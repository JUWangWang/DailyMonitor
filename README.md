# 風控整合日報 自動產出系統

## 檔案結構
```
DailyReport/
├── config.py          ← ⭐ 路徑設定（搬到公司只改這裡）
├── extract.py         ← 讀取 Excel、解析數據
├── db.py              ← SQLite 資料庫存取
├── render.py          ← 產出 HTML 報告
├── main.py            ← 主程式入口
├── requirements.txt   ← 套件清單
├── output/            ← HTML 每天自動產在這（自動建立）
└── 風控日報.db        ← 資料庫（自動建立）
```

## 安裝
```bash
pip install -r requirements.txt
```

## 使用方式

### 產今天的報告
```bash
python main.py
```

### 產指定日期的報告
```bash
python main.py 20260306
```

### 重跑所有歷史（Excel 都要在才能跑）
```bash
python main.py --rebuild-all
```

## 搬到公司電腦的步驟
1. 從 GitHub clone 下來
2. 打開 `config.py`，修改以下三個路徑：
   - `BROKER_DIR`  → 融資餘額分佈 Excel 的資料夾
   - `MARKET_DIR`  → 風險管理摘要說明 Excel 的資料夾
   - `WM_DIR`      → 財管商品集中度 Excel 的資料夾
   - `OUTPUT_DIR`  → HTML 輸出資料夾
   - `DB_PATH`     → 資料庫存放路徑
3. `pip install -r requirements.txt`
4. 執行 `python main.py` 測試

## 設定 Windows 工作排程器（每天自動跑）
1. 開啟「工作排程器」
2. 建立基本工作
3. 觸發程序：每天 08:30
4. 動作：啟動程式
   - 程式：`C:\Python311\python.exe`（你的 Python 路徑）
   - 引數：`main.py`
   - 起始位置：`D:\Python\DailyReport`
