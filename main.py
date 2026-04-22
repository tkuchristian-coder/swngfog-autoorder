"""
Google Sheet → swngfog 自動下單腳本

使用方式：
1. 安裝依賴：pip install -r requirements.txt
2. 確認 credentials.json 在同一目錄（Service Account 金鑰）
3. 在 config.py 設定 START_ROW
4. 執行：python main.py
"""

import re
import time
import smtplib
import requests
import gspread
from gspread.exceptions import APIError
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os

from config import (
    SWNGFOG_API_KEY, SWNGFOG_API_URL,
    SHEET_ID, SHEET_TAB_NAME,
    START_ROW, SERVICE_MAP, MANUAL_SERVICES, SKIP_LINKS,
    COL_ORDER_NO, COL_SERVICE, COL_LINK, COL_QTY, COL_AI_TAG, COL_STATUS,
    BATCH_SIZE, AI_TAG_START_ROW,
    ALERT_EMAIL_TO, ALERT_EMAIL_FROM, ALERT_EMAIL_PASSWORD,
)

CREDENTIALS_FILE = os.path.join(os.path.dirname(__file__), "credentials.json")


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

def pause_due_to_balance(row_num, order_no, service_name, igid, qty, batch_done, batch_total):
    """偵測到餘額不足時：寄信告知進度，並停用 cron job"""
    import subprocess
    from datetime import datetime

    done_qty = batch_done * BATCH_SIZE
    remaining_qty = qty - done_qty
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # 1. 寄 email
    send_alert_email(
        subject=f"[swngfog] ⛔ 餘額不足，自動暫停（{timestamp}）",
        body=(
            f"swngfog 餘額不足，自動下單腳本已暫停所有活動。\n\n"
            f"=== 最後處理進度 ===\n"
            f"  時間     ：{timestamp}\n"
            f"  列號     ：第 {row_num} 列\n"
            f"  訂單號   ：#{order_no}\n"
            f"  服務     ：{service_name}\n"
            f"  IGID    ：{igid}\n"
            f"  原始數量 ：{qty}\n"
            f"  已下單   ：{done_qty}（{batch_done}/{batch_total} 批）\n"
            f"  未完成   ：{remaining_qty} 個\n\n"
            f"=== 處理步驟 ===\n"
            f"  1. 請到 swngfog 充值帳戶餘額\n"
            f"  2. 充值後執行以下指令重新啟動自動排程：\n"
            f"     (crontab -l; echo '*/5 * * * * /Users/yangkyun-kai/Desktop/Gsheet_to_fog/run.sh') | crontab -\n"
            f"  3. 或直接執行補單：\n"
            f"     python3 /Users/yangkyun-kai/Desktop/Gsheet_to_fog/main.py\n"
        ),
    )

    # 2. 停用 cron（移除含 run.sh 的那行）
    try:
        result = subprocess.run(
            "crontab -l 2>/dev/null | grep -v 'Gsheet_to_fog/run.sh' | crontab -",
            shell=True, capture_output=True, text=True
        )
        print(f"  [⛔] Cron job 已停用")
    except Exception as e:
        print(f"  [警告] 停用 cron 失敗：{e}")

    print(f"  [⛔] 已寄送進度 email，腳本暫停")


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

        # 跳過已通知人工（避免每輪重複寄 email）
        # 等實際完成後人工把 I 欄改成「完成」即可
        if status.strip() == "已通知人工":
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
                    f"請在 config.py 的 SERVICE_MAP 或 MANUAL_SERVICES 新增此服務名稱。"
                ),
            )
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
        service_id  = SERVICE_MAP[service_name]
        igid        = extract_igid(link_raw)
        full_batches = qty // BATCH_SIZE
        remainder   = qty % BATCH_SIZE
        batch_count = full_batches + (1 if remainder else 0)

        resume_info = f"（從第{resume_batch + 1}批繼續）" if resume_batch > 0 else ""
        print(f"[處理] 列{row_num} 訂單#{order_no} | {service_name}(ID:{service_id}) | {igid} | "
              f"數量:{qty} → {full_batches}批×{BATCH_SIZE}" + (f" + 1批×{remainder}" if remainder else "") + resume_info)

        total_orders += 1
        order_failed = False
        api_calls_this_order = 0
        failed_batches = []   # 收集本訂單失敗批次，最後一封匯總 email

        # ── 開始前標記「處理中」，防止並行重複執行 / 支援斷點續跑 ──
        if not dry_run and resume_batch == 0:
            try:
                ws.update_cell(row_num, COL_STATUS + 1, f"處理中:0/{batch_count}")
            except Exception as e:
                print(f"  [警告] 無法寫入處理中狀態：{e}")

        for i in range(resume_batch, full_batches):
            if dry_run:
                print(f"  [DRY RUN] service={service_id} link={igid} quantity={BATCH_SIZE}")
                api_calls_this_order += 1
            else:
                try:
                    result = place_order(service_id, igid, BATCH_SIZE)
                    print(f"  [成功] 批次{i+1}/{batch_count} → fog訂單ID:{result.get('order','?')}")
                    total_api_calls += 1
                    api_calls_this_order += 1
                    try:
                        ws.update_cell(row_num, COL_STATUS + 1, f"處理中:{i+1}/{batch_count}")
                    except Exception as ue:
                        print(f"  [警告] 無法更新進度：{ue}")
                except Exception as e:
                    err_msg = str(e)
                    # ── 餘額不足 → 寄信記憶進度 + 停止 cron ──────────
                    if "balance" in err_msg.lower():
                        print(f"  [⛔ 餘額不足] 停止所有活動")
                        pause_due_to_balance(row_num, order_no, service_name, igid, qty, i, batch_count)
                        return  # 立即中止整個流程
                    print(f"  [失敗] 批次{i+1}/{batch_count}：{err_msg}")
                    failed_batches.append(f"批次{i+1}/{batch_count}：{err_msg}")
                    order_failed = True
                time.sleep(1.5)

        if remainder > 0 and resume_batch <= full_batches:
            if dry_run:
                print(f"  [DRY RUN] service={service_id} link={igid} quantity={remainder}")
                api_calls_this_order += 1
            else:
                try:
                    result = place_order(service_id, igid, remainder)
                    print(f"  [成功] 餘量批次({remainder}個) → fog訂單ID:{result.get('order','?')}")
                    total_api_calls += 1
                    api_calls_this_order += 1
                    try:
                        ws.update_cell(row_num, COL_STATUS + 1, f"處理中:{batch_count}/{batch_count}")
                    except Exception as ue:
                        print(f"  [警告] 無法更新進度：{ue}")
                except Exception as e:
                    err_msg = str(e)
                    if "balance" in err_msg.lower():
                        print(f"  [⛔ 餘額不足] 停止所有活動")
                        pause_due_to_balance(row_num, order_no, service_name, igid, qty, full_batches, batch_count)
                        return
                    print(f"  [失敗] 餘量批次（{remainder}個）：{err_msg}")
                    failed_batches.append(f"餘量批次({remainder}個)：{err_msg}")
                    order_failed = True
                time.sleep(1.5)

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
    import sys
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
