"""
main.py  —  風控整合日報 主程式
=====================================================
用法：
  python main.py               # 產今天的報告
  python main.py 20260306      # 產指定日期的報告
  python main.py --rebuild-all # 重跑所有歷史（需要 Excel 都在）
=====================================================
"""

import sys
import argparse

# Windows CMD 編碼設定（只在真實 CMD 環境下才套用）
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except AttributeError:
        pass  # Jupyter 環境略過
from datetime import date, datetime
from pathlib import Path

import config
from extract import extract_all
from db     import init_db, save_report, load_custom_sections
from render import generate_html, save_html


def run(target_date: date, overwrite: bool = True):
    print(f"\n{'='*50}")
    print(f"  風控整合日報 產出中：{target_date.strftime('%Y/%m/%d')}")
    print(f"{'='*50}")

    # 1. 讀取 Excel
    print("\n[1/3] 讀取 Excel...")
    data = extract_all(target_date, config)

    # 2. 存入 DB
    print("\n[2/3] 儲存至資料庫...")
    init_db(config.DB_PATH)
    alert_level, alert_items = save_report(config.DB_PATH, data, overwrite=overwrite)
    data["alert_level"] = alert_level
    data["alert_items"] = alert_items

    # 3. 產出 HTML
    print("\n[3/3] 產出 HTML...")
    report_date_db = data["report_date"].replace("/", "-")
    custom_sections = load_custom_sections(config.DB_PATH, report_date_db)
    html = generate_html(data, custom_sections=custom_sections)
    out_path = save_html(html, config.OUTPUT_DIR, data["report_date"])
    print(f"  OK HTML 已存至：{out_path}")

    print(f"\nDONE 完成！燈號：{alert_level.upper()}")
    return out_path, data


def main():
    parser = argparse.ArgumentParser(description="風控整合日報產出工具")
    parser.add_argument("date", nargs="?", help="指定日期 YYYYMMDD，不填則用今天")
    parser.add_argument("--rebuild-all", action="store_true", help="重跑所有歷史日期")
    parser.add_argument("--email", action="store_true", help="產出後直接寄信")
    parser.add_argument("--preview-email", action="store_true", help="產出後開啟 Outlook 草稿（不直接寄出）")
    args = parser.parse_args()

    if args.rebuild_all:
        # 掃描三個資料夾，找出所有日期
        import re
        dates = set()
        for folder, prefix in [
            (config.BROKER_DIR, config.BROKER_PREFIX),
            (config.MARKET_DIR, config.MARKET_PREFIX),
            (config.WM_DIR,     config.WM_PREFIX),
        ]:
            for f in Path(folder).glob(f"{prefix}*.xls*"):
                m = re.search(r'(\d{8})', f.name)
                if m:
                    dates.add(m.group(1))
        dates = sorted(dates)
        print(f"找到 {len(dates)} 個日期：{dates}")
        for d_str in dates:
            try:
                d = datetime.strptime(d_str, "%Y%m%d").date()
                run(d, overwrite=True)
            except Exception as e:
                print(f"  ⚠ {d_str} 失敗：{e}")
    else:
        if args.date:
            target = datetime.strptime(args.date, "%Y%m%d").date()
        else:
            target = date.today()
        result = run(target)
        if result:
            out_path, data = result
            if args.email:
                print("\n[3.5] 寄送 email...")
                from email_sender import send_report
                send_report(target, out_path, data)
            elif args.preview_email:
                print("\n[3.5] 開啟 Outlook 草稿...")
                from email_sender import preview_report
                preview_report(target, out_path, data)


if __name__ == "__main__":
    main()