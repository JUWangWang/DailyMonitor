# 風險管理日報 自動產出系統

## 檔案結構
DailyReport/
├── config.py          ← 路徑設定
├── extract.py         ← 讀取 Excel、解析數據
├── db.py              ← SQLite 資料庫存取
├── render.py          ← 產出 HTML 報告
├── main.py            ← 主程式入口
├── requirements.txt   ← 套件清單
├── output/            ← HTML 每天自動產在這（自動建立）
└── 風控日報.db        ← 資料庫（自動建立）

## 安裝
pip install -r requirements.txt

## 使用方式

### 產今天的報告
python main.py

### 產指定日期的報告
python main.py 20260306

### 重跑所有歷史（Excel 都要在才能跑）
python main.py --rebuild-all

### 20260321 日誌 ###
第 1 步改 db.py
第 2 步改 dashboard.py
第 3 步改 render.py
第 4 步測試
┌──────────────────────────────┐
│         Excel 原始資料        │
│  市場 / 財管 / 經紀 報表檔案   │
└─────────────┬────────────────┘
              │
              ▼
┌──────────────────────────────┐
│          extract.py          │
│  讀 Excel、解析成 market/wm   │
│  /broker 結構化 dict         │
└─────────────┬────────────────┘
              │
              ▼
┌──────────────────────────────┐
│            db.py             │
│  daily_summary               │
│  market_pnl                  │
│  wm_concentration            │
│  broker_margin               │
│  custom_sections             │
└───────┬───────────────┬──────┘
        │               │
        │               ▼
        │      ┌──────────────────────┐
        │      │   dashboard.py       │
        │      │  報告區塊編輯器       │
        │      │  - 新增/編輯/停用     │
        │      │  - layout_mode       │
        │      │  - insert_after      │
        │      │  - 複製昨日區塊       │
        │      └─────────┬────────────┘
        │                │
        ▼                │
┌──────────────────────────────┐
│           main.py            │
│  產生每日日報資料流程         │
│  extract → db → render       │
└─────────────┬────────────────┘
              │
              ▼
┌──────────────────────────────┐
│          render.py           │
│  主報表頁面                  │
│  + custom_sections           │
│  - full_page                 │
│  - inline                    │
│  - summary / market / wm     │
│    / broker / appendix 插入  │
└─────────────┬────────────────┘
              │
              ▼
┌──────────────────────────────┐
│       HTML / PDF 輸出         │
│  正式報告 / 附錄 / 補充說明    │
└──────────────────────────────┘

