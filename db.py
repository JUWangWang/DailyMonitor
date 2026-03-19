# ============================================================
#  db.py  —  SQLite 存取
# ============================================================

import sqlite3
import json
from pathlib import Path
from datetime import date


def get_conn(db_path: Path):
    db_path.parent.mkdir(parents=True, exist_ok=True)
    return sqlite3.connect(str(db_path))


def init_db(db_path: Path):
    """建立資料表（第一次執行時）"""
    with get_conn(db_path) as conn:
        conn.executescript("""
        CREATE TABLE IF NOT EXISTS daily_summary (
            report_date  TEXT PRIMARY KEY,
            market_json  TEXT,   -- 自營業務完整資料
            wm_json      TEXT,   -- 財管商品完整資料
            broker_json  TEXT,   -- 經紀業務完整資料
            alert_level  TEXT,   -- 整體燈號：red / yellow / green
            alert_items  TEXT,   -- 今日重點（JSON array）
            created_at   TEXT DEFAULT (datetime('now','localtime'))
        );

        -- 自營損益（方便跨日比較）
        CREATE TABLE IF NOT EXISTS market_pnl (
            report_date TEXT,
            dept        TEXT,
            biz         TEXT,
            dtd         REAL,
            mtd         REAL,
            ytd         REAL,
            status      TEXT,
            PRIMARY KEY (report_date, dept, biz)
        );

        -- 財管集中度（方便趨勢圖）
        CREATE TABLE IF NOT EXISTS wm_concentration (
            report_date TEXT,
            category    TEXT,
            name        TEXT,
            pct         REAL,
            l1          REAL,
            l2          REAL,
            status      TEXT,
            PRIMARY KEY (report_date, category)
        );

        -- 超限/警示事件
        CREATE TABLE IF NOT EXISTS alert_events (
            report_date TEXT,
            source      TEXT,   -- market / wm / broker
            type        TEXT,   -- 超限 / 80%提醒 / L1 / L2
            dept        TEXT,
            name        TEXT,
            value       REAL,
            note        TEXT,
            PRIMARY KEY (report_date, source, name)
        );

        -- 融資維持率（方便趨勢圖）
        CREATE TABLE IF NOT EXISTS broker_margin (
            report_date   TEXT PRIMARY KEY,
            total_balance REAL,
            total_maint   REAL,
            abc_pct       REAL,
            grade_a_pct   REAL,
            grade_b_pct   REAL,
            grade_c_pct   REAL,
            grade_d_pct   REAL,
            grade_e_pct   REAL
        );

        CREATE TABLE IF NOT EXISTS custom_sections (
            report_date        TEXT,
            section_id         TEXT,
            title              TEXT,
            section_type       TEXT,   -- text / bullets / table
            content_json       TEXT,
            display_order      INTEGER DEFAULT 100,
            enabled            INTEGER DEFAULT 1,
            page_break_before  INTEGER DEFAULT 0,
            created_at         TEXT DEFAULT (datetime('now','localtime')),
            updated_at         TEXT DEFAULT (datetime('now','localtime')),
            PRIMARY KEY (report_date, section_id)
        );
        """)
    print(f"  OK DB 初始化完成：{db_path}")


def save_report(db_path: Path, data: dict, overwrite: bool = True):
    """把當天所有資料存進 DB"""
    report_date = data["report_date"].replace("/", "-")
    market  = data["market"]
    wm      = data["wm"]
    broker  = data["broker"]

    # 計算整體燈號
    has_red    = bool(market["loss_over"] or market["d3_over"] or
                      any(v["status"] in ("達L2",) for v in wm["conc"].values()))
    has_yellow = bool(market["loss_warn"] or market["d3_warn"] or
                      any(v["status"] in ("達L1","接近L1","80%提醒") for v in wm["conc"].values()))
    alert_level = "red" if has_red else ("yellow" if has_yellow else "green")

    # 今日重點條列
    alert_items = []
    for r in market["loss_over"]:
        alert_items.append({"type":"red","text":f"自營 {r['dept']}{r['biz']} 月損失超限（{r['m_pct']*100:.1f}%）"})
    for r in market["d3_over"]:
        alert_items.append({"type":"red","text":f"單檔損失超限 {r['code']} {r['name']}（{r['loss_rate']*100:.1f}%）"})
    for k, v in wm["conc"].items():
        if v["status"] in ("達L2","達L1"):
            alert_items.append({"type":"red","text":f"財管 {v['name']} {v['status']}（{v['pct']*100:.2f}%）"})
        elif v["status"] in ("接近L1","80%提醒"):
            alert_items.append({"type":"yellow","text":f"財管 {v['name']} {v['status']}（{v['pct']*100:.2f}%）"})
    for r in market["loss_warn"]:
        alert_items.append({"type":"yellow","text":f"自營 {r['dept']}{r['biz']} 月損失80%提醒"})

    with get_conn(db_path) as conn:
        # daily_summary
        if overwrite:
            conn.execute("DELETE FROM daily_summary WHERE report_date=?", (report_date,))
            conn.execute("DELETE FROM market_pnl WHERE report_date=?", (report_date,))
            conn.execute("DELETE FROM wm_concentration WHERE report_date=?", (report_date,))
            conn.execute("DELETE FROM alert_events WHERE report_date=?", (report_date,))
            conn.execute("DELETE FROM broker_margin WHERE report_date=?", (report_date,))

        conn.execute("""
            INSERT OR REPLACE INTO daily_summary
            (report_date, market_json, wm_json, broker_json, alert_level, alert_items)
            VALUES (?,?,?,?,?,?)
        """, (report_date, json.dumps(market, ensure_ascii=False),
              json.dumps(wm, ensure_ascii=False),
              json.dumps(broker, ensure_ascii=False),
              alert_level, json.dumps(alert_items, ensure_ascii=False)))

        # market_pnl
        all_pnl = (
            [(r["dept"], "投資銀行處", r["dtd"], r["mtd"], r["ytd"], r["status"]) for r in market["ib_rows"]] +
            [(r["dept"], "策略部位",   r["dtd"], r["mtd"], r["ytd"], r["status"]) for r in market["strategy_rows"]] +
            [(r["dept"], "交易部位",   r["dtd"], r["mtd"], r["ytd"], r["status"]) for r in market["trade_rows"]]
        )
        conn.executemany("""
            INSERT OR REPLACE INTO market_pnl
            (report_date, dept, biz, dtd, mtd, ytd, status) VALUES (?,?,?,?,?,?,?)
        """, [(report_date, d, b, dtd, mtd, ytd, s) for d, b, dtd, mtd, ytd, s in all_pnl])

        # wm_concentration
        conc_rows = [(k, v["name"], v["pct"], v["l1"], v["l2"], v["status"])
                     for k, v in wm["conc"].items()]
        conn.executemany("""
            INSERT OR REPLACE INTO wm_concentration
            (report_date, category, name, pct, l1, l2, status) VALUES (?,?,?,?,?,?,?)
        """, [(report_date, cat, nm, pct, l1, l2, st) for cat, nm, pct, l1, l2, st in conc_rows])

        # alert_events
        for item in alert_items:
            conn.execute("""
                INSERT OR REPLACE INTO alert_events
                (report_date, source, type, dept, name, value, note) VALUES (?,?,?,?,?,?,?)
            """, (report_date, "market", item["type"], "", item["text"], 0, ""))

        # broker_margin
        grades = {r["grade"]: r for r in broker["dist_rows"]}
        conn.execute("""
            INSERT OR REPLACE INTO broker_margin
            (report_date, total_balance, total_maint, abc_pct,
             grade_a_pct, grade_b_pct, grade_c_pct, grade_d_pct, grade_e_pct)
            VALUES (?,?,?,?,?,?,?,?,?)
        """, (report_date,
              broker["total_balance"], broker["total_maint"], broker["abc_pct"],
              (grades.get("A") or {}).get("pct", 0),
              (grades.get("B") or {}).get("pct", 0),
              (grades.get("C") or {}).get("pct", 0),
              (grades.get("D") or {}).get("pct", 0),
              (grades.get("E") or {}).get("pct", 0)))

    print(f"  OK DB 儲存完成：{report_date}（燈號：{alert_level}）")
    return alert_level, alert_items


def load_report(db_path: Path, report_date: str) -> dict | None:
    """從 DB 讀取某天的完整資料"""
    with get_conn(db_path) as conn:
        row = conn.execute(
            "SELECT market_json, wm_json, broker_json, alert_level, alert_items "
            "FROM daily_summary WHERE report_date=?", (report_date,)
        ).fetchone()
    if not row:
        return None
    return {
        "market":      json.loads(row[0]),
        "wm":          json.loads(row[1]),
        "broker":      json.loads(row[2]),
        "alert_level": row[3],
        "alert_items": json.loads(row[4]),
        "report_date": report_date.replace("-", "/"),
    }


def list_dates(db_path: Path) -> list[dict]:
    """列出所有已存的日期與燈號（供 Streamlit 用）"""
    with get_conn(db_path) as conn:
        rows = conn.execute(
            "SELECT report_date, alert_level, alert_items FROM daily_summary ORDER BY report_date DESC"
        ).fetchall()
    return [{"date": r[0], "level": r[1], "alerts": json.loads(r[2])} for r in rows]

-- 儲存區塊
def save_custom_section(db_path: Path, report_date: str, section: dict):
    with get_conn(db_path) as conn:
        conn.execute("""
            INSERT OR REPLACE INTO custom_sections
            (report_date, section_id, title, section_type, content_json,
             display_order, enabled, page_break_before, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, datetime('now','localtime'))
        """, (
            report_date,
            section["section_id"],
            section["title"],
            section["section_type"],
            json.dumps(section["content"], ensure_ascii=False),
            section.get("display_order", 100),
            1 if section.get("enabled", True) else 0,
            1 if section.get("page_break_before", False) else 0,
        ))

-- 讀取某日所有區塊
def load_custom_sections(db_path: Path, report_date: str) -> list[dict]:
    with get_conn(db_path) as conn:
        rows = conn.execute("""
            SELECT section_id, title, section_type, content_json,
                   display_order, enabled, page_break_before
            FROM custom_sections
            WHERE report_date=?
            ORDER BY display_order, section_id
        """, (report_date,)).fetchall()

    return [{
        "section_id": r[0],
        "title": r[1],
        "section_type": r[2],
        "content": json.loads(r[3]),
        "display_order": r[4],
        "enabled": bool(r[5]),
        "page_break_before": bool(r[6]),
    } for r in rows]

--刪除區塊
def delete_custom_section(db_path: Path, report_date: str, section_id: str):
    with get_conn(db_path) as conn:
        conn.execute("""
            DELETE FROM custom_sections
            WHERE report_date=? AND section_id=?
        """, (report_date, section_id))

--複製前一日區塊
def copy_custom_sections(db_path: Path, from_date: str, to_date: str):
    sections = load_custom_sections(db_path, from_date)
    with get_conn(db_path) as conn:
        conn.execute("DELETE FROM custom_sections WHERE report_date=?", (to_date,))
        for s in sections:
            conn.execute("""
                INSERT OR REPLACE INTO custom_sections
                (report_date, section_id, title, section_type, content_json,
                 display_order, enabled, page_break_before, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, datetime('now','localtime'))
            """, (
                to_date,
                s["section_id"],
                s["title"],
                s["section_type"],
                json.dumps(s["content"], ensure_ascii=False),
                s["display_order"],
                1 if s["enabled"] else 0,
                1 if s["page_break_before"] else 0,
            ))
