"""
Google Sheet → swngfog 自動下單腳本

使用方式：
1. 安裝依賴：pip install -r requirements.txt
2. 確認 credentials.json 在同一目錄（Service Account 金鑰）
3. 在 config.py 設定 START_ROW
4. 執行：python main.py
"""

import re
import sys
import time
import signal
import atexit
import smtplib
import subprocess
import requests
import gspread
from gspread.exceptions import APIError
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os

from config import (
    SWNGFOG_API_KEY, SWNGFOG_API_URL,
    SHEET_ID, SHEET_TAB_NAME,
    START_ROW, SERVICE_MAP, MANUAL_SERVICES, SKIP_LINKS, ALLOWED_SERVICES, SERVICE_BATCH_SIZE,
    COL_ORDER_NO, COL_SERVICE, COL_LINK, COL_QTY, COL_AI_TAG, COL_STATUS,
    COL_FOG_IDS, COL_DELIVERY,
    BATCH_SIZE, AI_TAG_START_ROW,
    ALERT_EMAIL_TO, ALERT_EMAIL_FROM, ALERT_EMAIL_PASSWORD,
)

CREDENTIALS_FILE = os.path.join(os.path.dirname(__file__), "credentials.json")
PAUSE_FLAG_FILE  = os.path.join(os.path.dirname(__file__), ".paused")

# ─── Auto-resume 設定 ───────────────────────────────────────
# .paused 存在時,若查到 swngfog 餘額 >= 門檻,自動 resume(每天上限 N 次,避免 loop)
AUTO_RESUME_COUNT_FILE        = os.path.join(os.path.dirname(__file__), ".auto_resume_count")
AUTO_RESUME_BALANCE_THRESHOLD = 5.0   # USD,高於此才自動 resume
AUTO_RESUME_MAX_PER_DAY       = 3     # 一天最多 auto-resume 次數


# ─────────────────────────────────────────────
# SIGTERM-safe 進度 flush
# ─────────────────────────────────────────────
# GitHub Actions cancel 會先送 SIGTERM 給 10 秒 grace 期再 SIGKILL。
# 我們在 batch 迴圈中把「已完成 batch 數」存入 _progress_state,
# 真的被 cancel 時 handler/atexit 會把最新進度 flush 回 sheet,
# 避免「每 5 批才寫 sheet」的節流造成 resume 重複下單。
_progress_state = {}


def _flush_pending_progress():
    """把 in-memory 的 batch 進度緊急寫回 sheet(SIGTERM/atexit 呼叫)"""
    if not _progress_state:
        return
    try:
        _progress_state['ws'].update_cell(
            _progress_state['row'],
            _progress_state['col'],
            f"處理中:{_progress_state['done']}/{_progress_state['total']}",
        )
        print(f"[緊急 flush] 列{_progress_state['row']} 寫入 處理中:{_progress_state['done']}/{_progress_state['total']}")
    except Exception as e:
        print(f"[緊急 flush 失敗] {e}")
    _progress_state.clear()


def _sigterm_handler(signum, frame):
    print(f"[SIGTERM] 收到訊號 {signum},緊急 flush 進度後退出")
    _flush_pending_progress()
    sys.exit(0)


signal.signal(signal.SIGTERM, _sigterm_handler)
atexit.register(_flush_pending_progress)


# ─────────────────────────────────────────────
# Email 通知
# ─────────────────────────────────────────────

def send_alert_email(subject: str, body: str):
    """
    寄送警告 email 到 ALERT_EMAIL_TO。
    若 ALERT_EMAIL_PASSWORD 未設定則只印出警告，不寄信。
    """
    if not ALERT_EMAIL_PASSWORD:
        print(f"  [EMAIL 略過] 未設定 App Password，無法寄信。主旨：{subject}")
        return

    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"] = ALERT_EMAIL_FROM
        msg["To"] = ALERT_EMAIL_TO
        msg.attach(MIMEText(body, "plain", "utf-8"))

        with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=15) as smtp:
            smtp.login(ALERT_EMAIL_FROM, ALERT_EMAIL_PASSWORD)
            smtp.sendmail(ALERT_EMAIL_FROM, ALERT_EMAIL_TO, msg.as_string())

        print(f"  [EMAIL 已寄出] → {ALERT_EMAIL_TO}  主旨：{subject}")
    except Exception as e:
        print(f"  [EMAIL 失敗] 無法寄信：{e}")


# ─────────────────────────────────────────────
# 餘額不足：寄信記錄進度 + 停止 cron
# ─────────────────────────────────────────────

def query_swngfog_balance():
    """查 swngfog 餘額。成功回 float (USD),失敗回 None"""
    try:
        resp = requests.post(SWNGFOG_API_URL, data={
            "key": SWNGFOG_API_KEY, "action": "balance",
        }, timeout=15)
        resp.raise_for_status()
        j = resp.json()
        bal = j.get("balance")
        if bal is None:
            return None
        return float(bal)
    except Exception as e:
        print(f"  [balance 查詢失敗] {e}")
        return None


def _read_auto_resume_count():
    """讀 .auto_resume_count 回 (today_str, count_today)"""
    from datetime import date
    today = date.today().isoformat()
    try:
        with open(AUTO_RESUME_COUNT_FILE, "r", encoding="utf-8") as f:
            line = f.read().strip()
            d, c = line.split(":")
            if d == today:
                return today, int(c)
    except Exception:
        pass
    return today, 0


def _write_auto_resume_count(today, count):
    try:
        with open(AUTO_RESUME_COUNT_FILE, "w", encoding="utf-8") as f:
            f.write(f"{today}:{count}")
    except Exception as e:
        print(f"  [寫入 auto_resume_count 失敗] {e}")


def _try_auto_resume():
    """.paused 存在時嘗試自動 resume。
    回 True = 已刪 .paused + push,可繼續處理
    回 False = 仍 pause,呼叫端應 return 退出
    """
    today, count = _read_auto_resume_count()
    if count >= AUTO_RESUME_MAX_PER_DAY:
        print(f"[auto-resume] 今日已達上限 {count}/{AUTO_RESUME_MAX_PER_DAY} 次,不自動 resume")
        return False

    balance = query_swngfog_balance()
    if balance is None:
        print(f"[auto-resume] 無法查 swngfog 餘額,維持 pause")
        return False
    print(f"[auto-resume] 偵測餘額 ${balance:.2f} (門檻 ${AUTO_RESUME_BALANCE_THRESHOLD})")
    if balance < AUTO_RESUME_BALANCE_THRESHOLD:
        print(f"[auto-resume] 餘額不足門檻,維持 pause")
        return False

    new_count = count + 1
    _write_auto_resume_count(today, new_count)
    try:
        os.remove(PAUSE_FLAG_FILE)
        subprocess.run(["git", "config", "user.email", "bot@swngfog.autoorder"], check=False)
        subprocess.run(["git", "config", "user.name",  "swngfog-bot"],         check=False)
        subprocess.run(["git", "add", ".paused", AUTO_RESUME_COUNT_FILE], check=False)
        subprocess.run(["git", "commit", "-m",
            f"auto-resume: 餘額 ${balance:.2f} (今日第 {new_count}/{AUTO_RESUME_MAX_PER_DAY} 次)"
        ], check=False)
        r = subprocess.run(["git", "push"], capture_output=True, text=True)
        if r.returncode != 0:
            print(f"[auto-resume] push 失敗: {r.stderr[:200]}")
            return False
    except Exception as e:
        print(f"[auto-resume] 失敗: {e}")
        return False

    print(f"[auto-resume] ✓ 已解除 pause (餘額 ${balance:.2f}),繼續處理")
    send_alert_email(
        subject=f"[swngfog] 🔄 Auto-resume (餘額 ${balance:.2f})",
        body=(f"偵測到 swngfog 餘額 ${balance:.2f} ≥ 門檻 ${AUTO_RESUME_BALANCE_THRESHOLD},"
              f"自動解除 pause。\n今日第 {new_count}/{AUTO_RESUME_MAX_PER_DAY} 次 auto-resume。"),
    )
    return True


def _commit_and_push_pause_flag(timestamp, row_num, order_no, remaining_qty):
    """建立 .paused flag + commit + push,下一輪 workflow 啟動時會立即退出"""
    try:
        with open(PAUSE_FLAG_FILE, "w", encoding="utf-8") as f:
            f.write(
                f"Paused at {timestamp}\n"
                f"Row: {row_num}\n"
                f"Order: #{order_no}\n"
                f"Remaining qty: {remaining_qty}\n\n"
                f"充值後請刪除此檔:\n"
                f"  rm .paused && git commit -am 'resume' && git push\n"
            )
        # GitHub Actions 的 GITHUB_TOKEN 已經設定好 push 權限(workflow permissions: contents: write)
        subprocess.run(["git", "config", "user.email", "bot@swngfog.autoorder"], check=False)
        subprocess.run(["git", "config", "user.name",  "swngfog-bot"],         check=False)
        subprocess.run(["git", "add", ".paused"], check=False)
        subprocess.run(["git", "commit", "-m", f"auto-pause: 餘額不足 列{row_num} #{order_no}"], check=False)
        r = subprocess.run(["git", "push"], capture_output=True, text=True)
        if r.returncode == 0:
            print("  [⛔] 已 commit + push .paused flag,下輪 workflow 會立即退出")
        else:
            print(f"  [警告] push .paused 失敗: {r.stderr[:200]}")
    except Exception as e:
        print(f"  [警告] 建立 .paused flag 失敗: {e}")


def pause_due_to_balance(row_num, order_no, service_name, igid, qty, batch_done, batch_total, ws=None, batch_size=None):
    """偵測到餘額不足時:寄信告知進度 + 推送 .paused flag 暫停後續 workflow

    寄信去重:使用 H 欄(col 8) 當計數器,同一列最多寄 2 封 email。
    充值恢復後要人工刪除 .paused + git push,順便手動清空 H 欄。
    """
    from datetime import datetime

    done_qty = batch_done * (batch_size if batch_size else BATCH_SIZE)
    remaining_qty = qty - done_qty
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # ── 檢查 email 計數器(H 欄 = col 8)──
    MAX_EMAILS = 2
    email_col = 8  # H
    should_email = True
    new_count = 1
    if ws is not None:
        try:
            current = ws.cell(row_num, email_col).value
            try:
                sent_count = int(current) if current else 0
            except (ValueError, TypeError):
                sent_count = 0
            if sent_count >= MAX_EMAILS:
                should_email = False
                print(f"  [⛔ 餘額不足] 列{row_num} 已寄 {sent_count} 封 email(>={MAX_EMAILS}),不再寄")
            else:
                new_count = sent_count + 1
        except Exception as e:
            print(f"  [警告] 讀 H 欄 email 計數失敗:{e}")

    # 不論是否寄信,一律 push .paused flag(避免後續 workflow 繼續跑)
    _commit_and_push_pause_flag(timestamp, row_num, order_no, remaining_qty)

    if not should_email:
        return

    send_alert_email(
        subject=f"[swngfog] ⛔ 餘額不足,自動暫停({timestamp})",
        body=(
            f"swngfog 餘額不足,自動下單腳本已暫停所有活動。\n\n"
            f"=== 最後處理進度 ===\n"
            f"  時間     :{timestamp}\n"
            f"  列號     :第 {row_num} 列\n"
            f"  訂單號   :#{order_no}\n"
            f"  服務     :{service_name}\n"
            f"  IGID    :{igid}\n"
            f"  原始數量 :{qty}\n"
            f"  已下單   :{done_qty}({batch_done}/{batch_total} 批)\n"
            f"  未完成   :{remaining_qty} 個\n\n"
            f"=== 處理步驟 ===\n"
            f"  1. 到 swngfog 充值帳戶餘額\n"
            f"  2. 刪除 .paused flag 讓 workflow 恢復:\n"
            f"       cd /Users/yangkyun-kai/Desktop/Gsheet_to_fog\n"
            f"       rm .paused && git commit -am 'resume' && git push\n"
            f"  3. 下一輪 cron-job.org 觸發時會自動從斷點續跑\n"
        ),
    )

    # 寫回 H 欄 email 計數
    if ws is not None:
        try:
            ws.update_cell(row_num, email_col, str(new_count))
        except Exception as e:
            print(f"  [警告] 寫入 H 欄 email 計數失敗:{e}")

    print(f"  [⛔] 已寄送進度 email(第 {new_count}/{MAX_EMAILS} 封),腳本暫停")


# ─────────────────────────────────────────────
# Google Sheets 授權（Service Account）
# ─────────────────────────────────────────────

def get_gspread_client():
    """取得 Google Sheets 客戶端（使用 Service Account，無需瀏覽器）"""
    if not os.path.exists(CREDENTIALS_FILE):
        raise FileNotFoundError(
            f"找不到 {CREDENTIALS_FILE}！\n"
            "請確認 credentials.json（Service Account 金鑰）已放在同一目錄。"
        )
    return gspread.service_account(filename=CREDENTIALS_FILE)


# ─────────────────────────────────────────────
# 工具函式
# ─────────────────────────────────────────────

def extract_igid(raw: str) -> str:
    """
    從 C 欄原始值提取 IGID
    - IG 貼文/reel URL（/p/ 或 /reel/）→ 回傳乾淨 URL（保留大小寫）
    - IG 帳號 URL → 提取帳號名（小寫）
    - 純帳號 → 直接回傳小寫
    """
    raw = raw.strip()
    if "instagram.com" in raw:
        if "/p/" in raw or "/reel/" in raw:
            # 貼文/reel：去 query string，保留大小寫（shortcode 大小寫敏感）
            return raw.split("?")[0].rstrip("/")
        # 帳號 URL → 提取帳號名（小寫）
        match = re.search(r"instagram\.com/([^/?#]+)", raw)
        if match:
            return match.group(1).lower().rstrip("/")
    # 純帳號（粉絲類）→ 小寫
    return raw.lower()


def parse_processing_status(status: str):
    """
    解析「處理中:X/N」格式
    回傳 (completed_batches, total_batches)，格式不符則回傳 None
    """
    if not status.startswith("處理中:"):
        return None
    try:
        progress = status[len("處理中:"):]
        x, n = progress.split("/")
        return int(x), int(n)
    except Exception:
        return None


def place_order(service_id: int, link: str, quantity: int) -> dict:
    """呼叫 swngfog API 下單"""
    payload = {
        "key": SWNGFOG_API_KEY,
        "action": "add",
        "service": service_id,
        "link": link,
        "quantity": quantity,
    }
    resp = requests.post(SWNGFOG_API_URL, data=payload, timeout=30)
    resp.raise_for_status()
    try:
        result = resp.json()
    except Exception:
        raise ValueError(f"API 回傳非 JSON 內容：{resp.text[:200]}")
    # 明確失敗（success:false）
    if result.get("success") is False:
        code = result.get("code", "?")
        msg  = result.get("msg", "?")
        raise ValueError(f"API 伺服器錯誤 code:{code} msg:{msg}")
    # 舊式錯誤欄位
    if "error" in result:
        raise ValueError(f"API 錯誤：{result['error']}")
    # 未回傳訂單 ID
    if not result.get("order"):
        raise ValueError(f"API 未回傳訂單 ID，原始回應：{str(result)[:200]}")
    return result


# ─────────────────────────────────────────────
# 主流程
# ─────────────────────────────────────────────

def gsheet_retry(fn, max_retries=3, base_delay=10):
    """
    帶指數退避的 Google Sheets API 重試。
    處理 503 (Service Unavailable) 和 429 (Rate Limit) 等暫時性錯誤。
    """
    for attempt in range(max_retries):
        try:
            return fn()
        except APIError as e:
            status = e.response.status_code if hasattr(e, 'response') else 0
            if status in (429, 500, 502, 503) and attempt < max_retries - 1:
                delay = base_delay * (2 ** attempt)
                print(f"  [重試] Google API {status} 錯誤，{delay}秒後重試（{attempt+1}/{max_retries}）...")
                time.sleep(delay)
            else:
                raise


def process_orders(dry_run: bool = False):
    """
    讀取 Google Sheet，處理「等待處理」訂單。

    Args:
        dry_run: True 表示只印出，不實際呼叫 API（測試用）
    """
    # ── Pause flag 檢查 ───────────────────────────────────────
    # pause_due_to_balance 會 commit .paused + push,下輪 workflow 讀到就退出。
    # 但現在加 auto-resume: 若 .paused 存在且 swngfog 餘額 >= 門檻,自動刪掉 .paused
    # (每日最多 N 次,避免 balance 剛好勉強夠的 loop)。
    if os.path.exists(PAUSE_FLAG_FILE):
        print("[⛔ PAUSED] 偵測到 .paused flag")
        try:
            with open(PAUSE_FLAG_FILE, "r", encoding="utf-8") as f:
                print(f"--- .paused 內容 ---\n{f.read()}")
        except Exception:
            pass
        if _try_auto_resume():
            print("[auto-resume] 繼續執行處理流程...\n")
        else:
            print("  → 人工充值後執行: rm .paused && git commit -am 'resume' && git push")
            return

    print("連接 Google Sheets...")
    gc = get_gspread_client()
    sh = gsheet_retry(lambda: gc.open_by_key(SHEET_ID))
    try:
        ws = gsheet_retry(lambda: sh.worksheet(SHEET_TAB_NAME))
    except APIError as e:
        status = e.response.status_code if hasattr(e, 'response') else '?'
        msg = f"Google Sheets API 錯誤（HTTP {status}），重試後仍失敗"
        print(f"[錯誤] {msg}")
        send_alert_email(
            subject=f"[swngfog] ⚠️ Google Sheets API 錯誤（{status}）",
            body=f"{msg}\n\n錯誤詳情：{e}\n\n請稍後再試，Google 服務可能暫時不可用。",
        )
        return
    except Exception as e:
        available = [w.title for w in sh.worksheets()]
        msg = f"找不到分頁「{SHEET_TAB_NAME}」，目前可用分頁：{available}"
        print(f"[錯誤] {msg}")
        send_alert_email(
            subject=f"[swngfog] ⚠️ 找不到分頁：{SHEET_TAB_NAME}",
            body=f"{msg}\n\n請確認 Google Sheet 分頁名稱是否正確，或新增對應月份的分頁。",
        )
        return

    all_rows = gsheet_retry(lambda: ws.get_all_values())
    print(f"共讀取 {len(all_rows)} 列，從第 {START_ROW} 列開始處理\n")

    total_api_calls = 0
    total_orders = 0
    errors = []      # 收集所有錯誤，最後一起寄 email

    for row_idx in range(START_ROW - 1, len(all_rows)):
        row = all_rows[row_idx]
        row_num = row_idx + 1  # 人類可讀列號

        def cell(col):
            return row[col].strip() if col < len(row) else ""

        order_no    = cell(COL_ORDER_NO)
        service_name = cell(COL_SERVICE)
        link_raw    = cell(COL_LINK)
        qty_raw     = cell(COL_QTY)
        status      = cell(COL_STATUS)

        # 跳過：欄位B（服務）或欄位C（連結）任一為空 → 不執行
        if not service_name or not link_raw:
            if order_no:  # 有訂單號才印，避免空列噪音
                print(f"[跳過] 列{row_num} 訂單#{order_no}：B欄或C欄為空，跳過")
            continue

        # 跳過已完成
        if status.strip() == "完成":
            continue

        # 跳過手動標記「進行中」的列（人工暫停用途，不要碰）
        if status.strip() == "進行中":
            continue

        # 跳過已通知人工（避免每輪重複寄 email）
        # 等實際完成後人工把 I 欄改成「完成」即可
        if status.strip() == "已通知人工":
            continue

        # ── 讚類服務全部改寄 email,不自動下單 ───────────────────
        # 比對規則:服務名稱含「讚」字就改寄 email(涵蓋 台灣讚/台灣自動讚/真人讚/台港華人讚/頂級讚…)
        # 必須在白名單過濾前,避免被靜默跳過
        if "讚" in service_name:
            msg = f"[讚類改寄 email] 列{row_num} 訂單#{order_no}:{service_name} × {qty_raw}"
            print(msg)
            errors.append(msg)
            send_alert_email(
                subject=f"[swngfog] 需要安排(讚類):{service_name} × {qty_raw}(訂單#{order_no})",
                body=(
                    f"以下讚類訂單已改為人工安排,不會自動呼叫 swngfog:\n\n"
                    f"  列號   :{row_num}\n"
                    f"  訂單號 :{order_no}\n"
                    f"  服務   :{service_name}\n"
                    f"  IGID  :{link_raw}\n"
                    f"  數量   :{qty_raw}\n\n"
                    f"提示:腳本已將 I 欄標記為「已通知人工」,下次掃描不會再寄 email。\n"
                    f"      手動處理完後請把 I 欄改成「完成」。"
                ),
            )
            try:
                ws.update_cell(row_num, COL_STATUS + 1, "已通知人工")
                print(f"  [✓] 已標記列{row_num} I欄為「已通知人工」")
            except Exception as e:
                print(f"  [警告] 寫入「已通知人工」標記失敗:{e}")
            continue

        # 白名單過濾:只處理 ALLOWED_SERVICES 內的服務,其餘靜默跳過(不寄 email)
        if ALLOWED_SERVICES and service_name not in ALLOWED_SERVICES:
            continue

        # ── 跳過名單：C 欄（IGID/連結）命中 SKIP_LINKS → 直接標記完成，不送 swngfog ──
        skip_set_lower = {s.strip().lower() for s in SKIP_LINKS}
        if link_raw.strip().lower() in skip_set_lower:
            print(f"[跳過名單] 列{row_num} 訂單#{order_no}：IGID「{link_raw}」在 SKIP_LINKS，直接標記完成")
            try:
                ws.update_cell(row_num, COL_STATUS + 1, "完成")
                print(f"  [✓] 已標記列{row_num} I欄為「完成」")
                if row_num >= AI_TAG_START_ROW:
                    ws.update_cell(row_num, COL_AI_TAG + 1, "AI下單")
                    print(f"  [✓] 已標記列{row_num} G欄為「AI下單」")
            except Exception as e:
                print(f"  [警告] 寫入跳過名單完成標記失敗：{e}")
            continue

        # 解析斷點續跑狀態「處理中:X/N」
        resume_batch = 0
        parsed = parse_processing_status(status)
        if parsed:
            resume_batch, _ = parsed
            print(f"[續單] 列{row_num} 訂單#{order_no}：偵測到中斷進度，從第{resume_batch + 1}批繼續")
        elif status and status.strip() not in ("等待處理",):
            print(f"[提示] 列{row_num} 訂單#{order_no}：狀態為「{status}」，非預期值，仍繼續處理")

        # ── 人工處理服務 → 寄 email 通知，不自動下單 ───────────
        if service_name in MANUAL_SERVICES:
            msg = f"[人工處理] 列{row_num} 訂單#{order_no}：{service_name} × {qty_raw}，需手動安排"
            print(msg)
            errors.append(msg)
            send_alert_email(
                subject=f"[swngfog] 需要安排：{service_name} × {qty_raw}（訂單#{order_no}）",
                body=(
                    f"以下訂單需要人工安排，請手動到供應商下單：\n\n"
                    f"  列號   ：{row_num}\n"
                    f"  訂單號 ：{order_no}\n"
                    f"  服務   ：{service_name}\n"
                    f"  IGID  ：{link_raw}\n"
                    f"  數量   ：{qty_raw}\n\n"
                    f"此服務尚未設定自動下單，請手動處理。\n\n"
                    f"提示：腳本已將 I 欄標記為「已通知人工」，下次掃描不會再寄 email。\n"
                    f"      你手動處理完後，請把 I 欄改成「完成」。"
                ),
            )
            # 標記 I 欄為「已通知人工」，避免下輪重複寄 email
            try:
                ws.update_cell(row_num, COL_STATUS + 1, "已通知人工")
                print(f"  [✓] 已標記列{row_num} I欄為「已通知人工」(下次掃描不再寄 email)")
            except Exception as e:
                print(f"  [警告] 寫入「已通知人工」標記失敗：{e}")
            continue

        # ── 完全未知服務名稱 → 寄 email 警告 ──────────────────
        if service_name not in SERVICE_MAP:
            msg = f"[未知服務] 列{row_num} 訂單#{order_no}：服務「{service_name}」不在系統內，無法下單"
            print(msg)
            errors.append(msg)
            send_alert_email(
                subject=f"[swngfog] ⚠️ 未知服務名稱：{service_name}（訂單#{order_no}）",
                body=(
                    f"自動下單腳本遇到完全未知的服務名稱，請確認：\n\n"
                    f"  列號   ：{row_num}\n"
                    f"  訂單號 ：{order_no}\n"
                    f"  服務   ：{service_name}\n"
                    f"  IGID  ：{link_raw}\n"
                    f"  數量   ：{qty_raw}\n\n"
                    f"請在 config.py 的 SERVICE_MAP 或 MANUAL_SERVICES 新增此服務名稱。\n\n"
                    f"提示：腳本已將 I 欄標記為「已通知人工」，下次掃描不會再寄 email。\n"
                    f"      修正服務名稱或在 config.py 加入後，請把 I 欄改回空白或「等待處理」。"
                ),
            )
            # 標記 I 欄為「已通知人工」，避免下輪重複寄 email
            try:
                ws.update_cell(row_num, COL_STATUS + 1, "已通知人工")
                print(f"  [✓] 已標記列{row_num} I欄為「已通知人工」(下次掃描不再寄 email)")
            except Exception as e:
                print(f"  [警告] 寫入「已通知人工」標記失敗：{e}")
            continue

        # ── 數量驗證 ──────────────────────────────────────────
        try:
            qty = int(float(qty_raw))
        except (ValueError, TypeError):
            msg = f"[無效數量] 列{row_num} 訂單#{order_no}：數量「{qty_raw}」無法解析，跳過"
            print(msg)
            errors.append(msg)
            send_alert_email(
                subject=f"[swngfog] 無效數量：訂單#{order_no}",
                body=(
                    f"自動下單腳本遇到無法解析的數量：\n\n"
                    f"  列號   ：{row_num}\n"
                    f"  訂單號 ：{order_no}\n"
                    f"  服務   ：{service_name}\n"
                    f"  數量   ：{qty_raw}（無效）\n\n"
                    f"請手動確認 Google Sheet 中該列的數量欄位。"
                ),
            )
            continue

        if qty <= 0:
            print(f"[跳過] 列{row_num} 訂單#{order_no}：數量為 {qty}，跳過")
            continue

        # ── 拆單並下單 ────────────────────────────────────────
        # 規則：batch_size 同時作為最低單位（min quantity）。
        # 餘量 < batch_size 時併入最後一批，避免 API 拒收（例如 55 拆成 10,10,10,10,15）。
        service_id  = SERVICE_MAP[service_name]
        igid        = extract_igid(link_raw)
        batch_size  = SERVICE_BATCH_SIZE.get(service_name, BATCH_SIZE)

        # 數量小於最低單位無法下單
        if qty < batch_size:
            msg = f"[數量不足] 列{row_num} 訂單#{order_no}：數量 {qty} < 最低單位 {batch_size}，跳過"
            print(msg)
            errors.append(msg)
            send_alert_email(
                subject=f"[swngfog] 數量不足最低單位：訂單#{order_no}",
                body=(
                    f"自動下單腳本遇到數量小於最低單位：\n\n"
                    f"  列號     ：{row_num}\n"
                    f"  訂單號   ：{order_no}\n"
                    f"  服務     ：{service_name}\n"
                    f"  數量     ：{qty}\n"
                    f"  最低單位 ：{batch_size}\n\n"
                    f"請手動處理或補足數量後改回「等待處理」。"
                ),
            )
            try:
                ws.update_cell(row_num, COL_STATUS + 1, "已通知人工")
            except Exception:
                pass
            continue

        full_batches = qty // batch_size
        remainder   = qty % batch_size
        batches = [batch_size] * full_batches
        if remainder > 0:
            # remainder < batch_size 一定成立（否則 % 已 = 0），併入最後一批
            batches[-1] += remainder
        batch_count = len(batches)

        resume_info = f"（從第{resume_batch + 1}批繼續）" if resume_batch > 0 else ""
        breakdown = f"{full_batches}批×{batch_size}" + (f"（最後1批={batches[-1]}）" if remainder else "")
        print(f"[處理] 列{row_num} 訂單#{order_no} | {service_name}(ID:{service_id}) | {igid} | "
              f"數量:{qty} → {breakdown}" + resume_info)

        total_orders += 1
        order_failed = False
        api_calls_this_order = 0
        failed_batches = []   # 收集本訂單失敗批次，最後一封匯總 email
        fog_order_ids = []    # 本訂單所有 fog 訂單 ID,供 verify_delivery 查交付狀態

        # ── 開始前標記「處理中」，防止並行重複執行 / 支援斷點續跑 ──
        if not dry_run and resume_batch == 0:
            try:
                ws.update_cell(row_num, COL_STATUS + 1, f"處理中:0/{batch_count}")
            except Exception as e:
                print(f"  [警告] 無法寫入處理中狀態：{e}")

        for i in range(resume_batch, batch_count):
            q_this = batches[i]
            if dry_run:
                print(f"  [DRY RUN] service={service_id} link={igid} quantity={q_this}")
                api_calls_this_order += 1
            else:
                try:
                    result = place_order(service_id, igid, q_this)
                    oid = result.get('order', '?')
                    print(f"  [成功] 批次{i+1}/{batch_count}（數量:{q_this}） → fog訂單ID:{oid}")
                    if oid and oid != '?':
                        fog_order_ids.append(str(oid))
                    total_api_calls += 1
                    api_calls_this_order += 1
                    # ── SIGTERM-safe 進度追蹤 ──
                    # 每批都更新 in-memory 進度,workflow 被 cancel 時 handler 會 flush 回 sheet,
                    # 避免「每 5 批才寫」的節流造成 resume 重複下單。
                    _progress_state.clear()
                    _progress_state.update({
                        'ws': ws, 'row': row_num, 'col': COL_STATUS + 1,
                        'done': i + 1, 'total': batch_count,
                    })
                    # 節流:每 5 批才寫 sheet 一次(減少 API 呼叫),最後一批必寫
                    if (i + 1) % 5 == 0 or (i + 1) == batch_count:
                        try:
                            ws.update_cell(row_num, COL_STATUS + 1, f"處理中:{i+1}/{batch_count}")
                            _progress_state.clear()  # 已 flush,清掉 in-memory 狀態
                        except Exception as ue:
                            print(f"  [警告] 無法更新進度：{ue}")
                except Exception as e:
                    err_msg = str(e)
                    # ── 餘額不足 → 悲觀處理防重複下單 ──────────────────
                    # 關鍵:swngfog 可能回 success:false 但實際已 CREATE 訂單(常見 bug)。
                    # 若寫 i (未算成功),下輪 resume 會從 i+1 重試,導致重複下單。
                    # 改寫 i+1 (當作已下),犧牲最多 10 粉 coverage,換不會重複下單。
                    if "balance" in err_msg.lower():
                        try:
                            ws.update_cell(row_num, COL_STATUS + 1, f"處理中:{i+1}/{batch_count}")
                        except Exception:
                            pass
                        # 重要:清掉 _progress_state,否則 atexit/SIGTERM 會 flush 舊的
                        # done=i 狀態把悲觀寫入 i+1 蓋回 i → 重複下單風險
                        _progress_state.clear()
                        print(f"  [⛔ 餘額不足] 批次{i+1} 可能已下單,寫進度 {i+1}/{batch_count} 避免重複,停止所有活動")
                        pause_due_to_balance(row_num, order_no, service_name, igid, qty, i+1, batch_count, ws=ws, batch_size=batch_size)
                        return  # 立即中止整個流程
                    # ── 其他錯誤(非 balance) → 寫 i (未算成功),可 resume 重試 ──
                    try:
                        ws.update_cell(row_num, COL_STATUS + 1, f"處理中:{i}/{batch_count}")
                    except Exception:
                        pass
                    print(f"  [失敗] 批次{i+1}/{batch_count}（數量:{q_this}）：{err_msg}")
                    failed_batches.append(f"批次{i+1}/{batch_count}（數量:{q_this}）：{err_msg}")
                    order_failed = True
                time.sleep(0.2)

        # ── 部分/全部失敗 → 一封匯總 email ──────────────────────
        if order_failed and failed_batches:
            errors.append(f"訂單#{order_no} 失敗{len(failed_batches)}批")
            send_alert_email(
                subject=f"[swngfog] 下單部分失敗：訂單#{order_no}（成功{api_calls_this_order}批／失敗{len(failed_batches)}批）",
                body=(
                    f"自動下單部分批次失敗：\n\n"
                    f"  列號   ：{row_num}\n"
                    f"  訂單號 ：{order_no}\n"
                    f"  服務   ：{service_name}（ID:{service_id}）\n"
                    f"  IGID  ：{igid}\n"
                    f"  總批次 ：{batch_count}\n"
                    f"  本次成功：{api_calls_this_order}\n"
                    f"  失敗   ：{len(failed_batches)}\n\n"
                    f"=== 失敗明細（前10筆）===\n" +
                    "\n".join(f"  {j+1}. {m}" for j, m in enumerate(failed_batches[:10])) +
                    (f"\n  ...（共 {len(failed_batches)} 筆）" if len(failed_batches) > 10 else "") +
                    f"\n\nSheet I欄已保留「處理中:{resume_batch + api_calls_this_order}/{batch_count}」，下次執行將自動從斷點繼續。"
                ),
            )

        # ── 全部批次成功 → 回寫 I 欄「完成」，並視需要寫 G 欄「AI下單」─
        total_successful = resume_batch + api_calls_this_order
        if not order_failed and total_successful == batch_count:
            # 本訂單完成,清掉 in-memory 進度(避免 atexit flush 蓋到已經是「完成」的 I 欄)
            _progress_state.clear()
            if dry_run:
                print(f"  [DRY RUN] 將列{row_num} I欄更新為「完成」")
                if row_num >= AI_TAG_START_ROW:
                    print(f"  [DRY RUN] 將列{row_num} G欄更新為「AI下單」")
            else:
                try:
                    ws.update_cell(row_num, COL_STATUS + 1, "完成")  # gspread 欄位從1起算
                    print(f"  [✓] 已標記列{row_num} I欄為「完成」")
                except Exception as e:
                    print(f"  [警告] 無法更新狀態欄：{e}")

                # 重置 H 欄 email 計數器(此列之後若再餘額不足可重新寄信通知)
                try:
                    ws.update_cell(row_num, 8, "")
                except Exception as e:
                    print(f"  [警告] 無法重置 H 欄 email 計數:{e}")

                # 寫 J 欄: fog 訂單 ID 範圍(供 verify_delivery.py 查交付狀態)
                if fog_order_ids:
                    j_val = (f"{fog_order_ids[0]}-{fog_order_ids[-1]}"
                             if len(fog_order_ids) > 1 else fog_order_ids[0])
                    try:
                        ws.update_cell(row_num, COL_FOG_IDS + 1, j_val)
                    except Exception as e:
                        print(f"  [警告] 無法寫 J 欄 fog IDs:{e}")

                if row_num >= AI_TAG_START_ROW:
                    try:
                        ws.update_cell(row_num, COL_AI_TAG + 1, "AI下單")  # gspread 欄位從1起算
                        print(f"  [✓] 已標記列{row_num} G欄為「AI下單」")
                    except Exception as e:
                        print(f"  [警告] 無法更新 G 欄：{e}")
        elif not order_failed:
            print(f"  [⚠️ 未標記完成] 列{row_num} 預期{batch_count}批，實際完成{total_successful}批")

    # ── 執行摘要 ──────────────────────────────────────────────
    print(f"\n{'='*50}")
    mode_tag = "[DRY RUN] " if dry_run else ""
    print(f"{mode_tag}完成！處理訂單數：{total_orders}，API呼叫次數：{total_api_calls}")

    if errors:
        print(f"\n警告/錯誤 ({len(errors)} 筆)：")
        for e in errors:
            print(f"  {e}")
    else:
        print("無任何警告或錯誤 ✓")


if __name__ == "__main__":
    import traceback

    dry_run = "--dry-run" in sys.argv
    if dry_run:
        print("=== DRY RUN 模式：只印出，不實際呼叫 API ===\n")

    try:
        process_orders(dry_run=dry_run)
    except Exception as e:
        tb = traceback.format_exc()
        print(f"\n[💥 CRASH] 腳本發生未預期錯誤：{e}\n{tb}")
        send_alert_email(
            subject=f"[swngfog] 💥 腳本 Crash，請立即檢查",
            body=(
                f"自動下單腳本發生未預期錯誤，已中止執行。\n\n"
                f"=== 錯誤訊息 ===\n{e}\n\n"
                f"=== 完整 Traceback ===\n{tb}\n\n"
                f"請檢查 Google Sheet 中狀態為「處理中:X/N」的列，"
                f"確認下次執行時是否需要手動介入。"
            ),
        )
        raise
