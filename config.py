import os
from datetime import datetime

# swngfog API 設定（支援 GitHub Actions Secrets 注入，本地執行 fallback 硬編碼）
SWNGFOG_API_KEY = os.environ.get("SWNGFOG_API_KEY", "")
SWNGFOG_API_URL = "https://www.swngfog.com/api/v1"

# Google Sheet 設定
SHEET_ID = os.environ.get("SHEET_ID", "1sbYoxrrMOPZsA2q6dPDqzU3OvoWykcsy-4ao9f9oyJk")
SHEET_TAB_NAME = os.environ.get("SHEET_TAB_NAME", "2026年3月")

# 從第幾列開始處理（含標題列後的第一行資料是第2列，index從1起算）
# 使用者執行前請修改此值
START_ROW = 778  # 行778以前不動

# ── 自動下單：服務名稱 → swngfog server ID ──────────────────
SERVICE_MAP = {
    # 粉絲類 (ID 3)
    "普通台灣粉":   3,
    "台港華人":     3,

    # 讚類 (ID 7)
    "台灣讚":           7,
    "快速台灣讚":       7,
    "真人讚":           7,
    "台港華人讚":       7,
    "陳輝自動讚/台港華人": 7,
    "台灣自動讚":       7,
    "真人自動讚":       7,
    "高品質台灣自動讚": 7,
    "普通台灣讚":       7,
}

# ── 人工處理：遇到這些服務改寄 email，不自動下單 ────────────
# Email 內容會告知服務名稱、IGID、數量，請手動安排
MANUAL_SERVICES = {
    "真人粉",
    "高品質(頂級)80-200",
    "高品質(頂級)300",
    "高品質(頂級)500",
    "高品質(頂級)600",
    "高品質(頂級)800",
    "互動粉",
    "台灣按讚粉中高互動",
    "智能按讚粉超高互動",
    "優質或高品質讚(200-500)",
    "頂級讚(追蹤500)",
    "快速普通台灣粉",
}

# 欄位 index（0-based）
COL_ORDER_NO = 0   # A欄：訂單號
COL_SERVICE  = 1   # B欄：服務類型
COL_LINK     = 2   # C欄：IG連結/IGID
COL_QTY      = 3   # D欄：數量
COL_AI_TAG   = 6   # G欄：AI下單標記
COL_STATUS   = 8   # I欄：狀態（"完成" 表示跳過）

# 從哪一行起才寫 AI下單 標記（G欄）
AI_TAG_START_ROW = 818

# 每批固定數量
BATCH_SIZE = 10

# Email 通知設定
ALERT_EMAIL_TO   = "snowmiecpay@gmail.com"
ALERT_EMAIL_FROM = "snowmiecpay@gmail.com"
# Gmail App Password（16碼）
# 申請：Google 帳號 → 安全性 → 兩步驟驗證 → 應用程式密碼
ALERT_EMAIL_PASSWORD = os.environ.get("GMAIL_PASSWORD", "")
