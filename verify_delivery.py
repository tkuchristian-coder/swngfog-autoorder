"""
verify_delivery.py
──────────────────
掃近期標「完成」的列,用 swngfog 公開 API (action=status) 查該 row 的
first/last fog 訂單 ID 實際交付狀態,寫入 K 欄「實際交付狀態」:
  - 已到帳:first+last 都 Completed
  - 派送中X%:部分 Completed
  - 處理中:都還沒開始
  - 異常:order 找不到或其他錯誤
"""
import os
import sys
import time
import requests
import gspread
from gspread.exceptions import APIError

sys.path.insert(0, os.path.dirname(__file__))
from config import (
    SWNGFOG_API_KEY, SWNGFOG_API_URL,
    SHEET_ID, SHEET_TAB_NAME, START_ROW,
    COL_STATUS, COL_FOG_IDS, COL_DELIVERY,
)

CREDENTIALS_FILE = os.path.join(os.path.dirname(__file__), "credentials.json")

# 只掃最後 N 列,避免掃全 sheet 浪費 API
SCAN_LAST_N = 500
# 已經寫過「已到帳」的就不重新驗證
FINAL_STATUSES = {"已到帳", "異常"}


def query_order_status(order_id):
    """查單筆 swngfog 訂單狀態。回 dict 或 None"""
    try:
        resp = requests.post(SWNGFOG_API_URL, data={
            "key": SWNGFOG_API_KEY, "action": "status", "order": order_id,
        }, timeout=15)
        resp.raise_for_status()
        j = resp.json()
        return j if isinstance(j, dict) else None
    except Exception as e:
        print(f"  [查 status 失敗] order={order_id}: {e}")
        return None


def classify(first_status, last_status):
    """根據 first+last order status 判斷整列交付狀態"""
    def norm(s):
        return (s or "").strip().lower()
    f, l = norm(first_status), norm(last_status)
    completed = {"completed", "complete"}
    pending = {"pending", "processing", "in progress", "in_progress"}
    partial_s = {"partial"}
    cancel_s = {"canceled", "cancelled"}
    if f in completed and l in completed:
        return "已到帳"
    if f in cancel_s or l in cancel_s:
        return "異常:含取消"
    if f in partial_s or l in partial_s:
        return "派送中(有 Partial)"
    if f in completed:
        return "派送中(頭段已完成)"
    if l in completed:
        return "派送中(尾段已完成)"
    if f in pending or l in pending:
        return "處理中"
    return f"異常:f={first_status[:20]},l={last_status[:20]}"


def main():
    print("verify_delivery: 連接 Google Sheets...")
    gc = gspread.service_account(filename=CREDENTIALS_FILE)
    ws = gc.open_by_key(SHEET_ID).worksheet(SHEET_TAB_NAME)
    all_rows = ws.get_all_values()
    total_rows = len(all_rows)
    scan_from = max(START_ROW - 1, total_rows - SCAN_LAST_N)
    print(f"共 {total_rows} 列,掃描 {scan_from+1}~{total_rows}")

    updated = 0
    skipped = 0
    processed = 0
    for row_idx in range(scan_from, total_rows):
        row = all_rows[row_idx]
        row_num = row_idx + 1

        def c(col):
            return row[col].strip() if col < len(row) else ""

        status = c(COL_STATUS)
        fog_ids = c(COL_FOG_IDS)
        delivery = c(COL_DELIVERY)

        # 只處理「完成」+ J 有 fog ID + K 不是終局狀態
        if status != "完成":
            skipped += 1; continue
        if not fog_ids:
            skipped += 1; continue
        if delivery in FINAL_STATUSES:
            skipped += 1; continue

        # 解析 J 欄:格式可能 "ID" 或 "FIRST-LAST"
        parts = fog_ids.split("-")
        first_id = parts[0].strip()
        last_id = parts[-1].strip() if len(parts) > 1 else first_id

        processed += 1
        print(f"\n列{row_num}: 查 fog ID first={first_id} last={last_id}")
        first_s = query_order_status(first_id)
        time.sleep(0.5)
        last_s = query_order_status(last_id) if last_id != first_id else first_s
        time.sleep(0.5)

        if first_s is None or last_s is None:
            new_val = "異常:查詢失敗"
        else:
            fs = first_s.get("status", "")
            ls = last_s.get("status", "")
            new_val = classify(fs, ls)
            print(f"  → first={fs!r} last={ls!r} → {new_val}")

        try:
            ws.update_cell(row_num, COL_DELIVERY + 1, new_val)
            updated += 1
        except APIError as e:
            print(f"  [寫入失敗] {e}")

    print(f"\n{'='*50}")
    print(f"完成:掃 {total_rows-scan_from} 列,處理 {processed} 列,"
          f"更新 {updated} 列,跳過 {skipped} 列")


if __name__ == "__main__":
    main()
