# ============================================================
#  config.py  —  路徑與參數設定
#  搬到公司電腦時，只需要修改這個檔案
# ============================================================

from pathlib import Path

# ── 來源 Excel 資料夾 ────────────────────────────────────────
# 開發環境（目前電腦）
BROKER_DIR   = Path(r"D:\Python\DailyReport\01.經紀")   # 融資餘額分佈
MARKET_DIR   = Path(r"D:\Python\DailyReport\02.市場")   # 風險管理摘要說明
WM_DIR       = Path(r"D:\Python\DailyReport\03.財管")   # 財管商品集中度

# 公司電腦上線後改成（範例）：
# BROKER_DIR = Path(r"\\172.21.131.92\風管部\01.經紀")
# MARKET_DIR = Path(r"\\172.21.131.92\風管部\02.市場")
# WM_DIR     = Path(r"\\172.21.131.92\風管部\03.財管")

# ── 檔名前綴（底線前的部分）────────────────────────────────
BROKER_PREFIX = "融資餘額分佈"
MARKET_PREFIX = "風險管理摘要說明"
WM_PREFIX     = "財管商品集中度管理報表"

# ── 輸出路徑 ─────────────────────────────────────────────────
BASE_DIR    = Path(r"D:\Python\DailyReport")
OUTPUT_DIR  = BASE_DIR / "output"     # HTML 每天存這
DB_PATH     = BASE_DIR / "風控日報.db"

# 公司電腦上線後改成（範例）：
# OUTPUT_DIR = Path(r"\\172.21.131.92\風管部\output")
# DB_PATH    = Path(r"\\172.21.131.92\風管部\風控日報.db")

# ── 其他設定 ─────────────────────────────────────────────────
REPORT_TITLE = "風險管理整合日報"
DEPT_NAME    = "風險管理部"

# ── 收件人名單 ────────────────────────────────────────────────
# 寄信時使用，可隨時新增/移除
EMAIL_TO = [
    "主管A@company.com",
    "主管B@company.com",
]
EMAIL_CC = [
    "同事C@company.com",
]

# 寄件人（預設用目前登入 Outlook 的帳號，通常不需要改）
EMAIL_FROM = ""   # 留空 = 用 Outlook 預設帳號

# 信件主旨格式（{date} 會自動帶入報告日期）
EMAIL_SUBJECT = "【風控整合日報】{date}"
