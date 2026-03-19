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

