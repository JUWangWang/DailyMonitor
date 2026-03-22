# ============================================================
#  dashboard.py  —  風險管理日報 歷史查詢介面（Streamlit）
#  執行方式：py -m streamlit run dashboard.py
# ============================================================

import sqlite3
import json
import subprocess
import sys
from pathlib import Path
from datetime import datetime, timedelta, date

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px

import config

import uuid
import copy
from db import load_custom_sections, save_custom_section, delete_custom_section

# ── 頁面設定 ────────────────────────────────────────────────
st.set_page_config(
    page_title="風險管理日報 查詢系統",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS ─────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stSidebar"] { background: #f8f9fb; }
.metric-card {
    background: white; border: 1px solid #e3e8ef;
    border-radius: 8px; padding: 12px 16px; margin-bottom: 8px;
}
.up   { color: #1a9e6a; font-weight: 600; }
.dn   { color: #c62828; font-weight: 600; }
.tag-red    { background:#fef2f2; color:#c62828; border:1px solid #fccaca;
              border-radius:4px; padding:2px 8px; font-size:12px; font-weight:700; }
.tag-orange { background:#fff7ed; color:#ea580c; border:1px solid #fed7aa;
              border-radius:4px; padding:2px 8px; font-size:12px; font-weight:700; }
.tag-yellow { background:#fffbeb; color:#b45309; border:1px solid #fde68a;
              border-radius:4px; padding:2px 8px; font-size:12px; font-weight:700; }
.tag-green  { background:#edf7f3; color:#1a9e6a; border:1px solid #b2dfcf;
              border-radius:4px; padding:2px 8px; font-size:12px; font-weight:700; }
.tag-blue   { background:#e8f0fb; color:#1976d2; border:1px solid #c2d3f0;
              border-radius:4px; padding:2px 8px; font-size:12px; font-weight:700; }
div[data-testid="column"] { padding: 4px 6px; }
</style>
""", unsafe_allow_html=True)


# ── DB 工具 ──────────────────────────────────────────────────
@st.cache_resource
def get_db_path():
    return config.DB_PATH

def db_conn():
    return sqlite3.connect(str(get_db_path()))

@st.cache_data(ttl=60)
def load_date_list():
    try:
        with db_conn() as conn:
            rows = conn.execute(
                "SELECT report_date, alert_level, alert_items FROM daily_summary ORDER BY report_date DESC"
            ).fetchall()
        return [{"date": r[0], "level": r[1], "alerts": json.loads(r[2])} for r in rows]
    except Exception:
        return []

@st.cache_data(ttl=60)
def load_day(report_date: str):
    with db_conn() as conn:
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
        "report_date": report_date,
    }

@st.cache_data(ttl=60)
def load_pnl_trend(dept: str, biz: str, start: str, end: str):
    with db_conn() as conn:
        rows = conn.execute("""
            SELECT report_date, dtd, mtd, ytd, status
            FROM market_pnl
            WHERE dept=? AND biz=? AND report_date BETWEEN ? AND ?
            ORDER BY report_date
        """, (dept, biz, start, end)).fetchall()
    return pd.DataFrame(rows, columns=["date","dtd","mtd","ytd","status"])

@st.cache_data(ttl=60)
def load_conc_trend(category: str, start: str, end: str):
    with db_conn() as conn:
        rows = conn.execute("""
            SELECT report_date, name, pct, l1, l2, status
            FROM wm_concentration
            WHERE category=? AND report_date BETWEEN ? AND ?
            ORDER BY report_date
        """, (category, start, end)).fetchall()
    return pd.DataFrame(rows, columns=["date","name","pct","l1","l2","status"])

@st.cache_data(ttl=60)
def load_broker_trend(start: str, end: str):
    with db_conn() as conn:
        rows = conn.execute("""
            SELECT report_date, total_balance, total_maint, abc_pct,
                   grade_a_pct, grade_b_pct, grade_c_pct, grade_d_pct, grade_e_pct,
                   unlim_total_balance, unlim_total_maint, unlim_abc_pct,
                   unlim_a_pct, unlim_b_pct, unlim_c_pct, unlim_d_pct, unlim_e_pct
            FROM broker_margin
            WHERE report_date BETWEEN ? AND ?
            ORDER BY report_date
        """, (start, end)).fetchall()
    return pd.DataFrame(rows, columns=[
        "date","total_balance","total_maint","abc_pct","A","B","C","D","E",
        "unlim_total_balance","unlim_total_maint","unlim_abc_pct",
        "uA","uB","uC","uD","uE"
    ])

@st.cache_data(ttl=60)
def load_alert_events(start: str, end: str):
    with db_conn() as conn:
        rows = conn.execute("""
            SELECT report_date, source, type, name
            FROM alert_events
            WHERE report_date BETWEEN ? AND ?
            ORDER BY report_date DESC
        """, (start, end)).fetchall()
    return pd.DataFrame(rows, columns=["日期","來源","類型","說明"])

def save_alert_items(db_path, report_date: str, alert_items: list[dict]):
    conn = sqlite3.connect(str(db_path))
    try:
        conn.execute(
            "UPDATE daily_summary SET alert_items=? WHERE report_date=?",
            (json.dumps(alert_items, ensure_ascii=False), report_date)
        )
        conn.commit()
    finally:
        conn.close()
    st.cache_data.clear()

TEMPLATE_PATH = Path("section_templates.json")

def load_section_templates():
    if not TEMPLATE_PATH.exists():
        return []
    try:
        return json.loads(TEMPLATE_PATH.read_text(encoding="utf-8"))
    except Exception:
        return []

def save_section_templates(templates):
    TEMPLATE_PATH.write_text(
        json.dumps(templates, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

def instantiate_section_from_template(template: dict, display_order: int | None = None):
    s = copy.deepcopy(template)
    s.pop("template_id", None)
    s.pop("template_name", None)
    s["title"] = s.pop("default_title", s.get("title", "未命名區塊"))
    s["section_id"] = str(uuid.uuid4())[:8]
    if display_order is not None:
        s["display_order"] = display_order
    return s

# ── 格式工具 ──────────────────────────────────────────────────
def fmt_wan(v, unit="萬"):
    if not v:
        return "0"
    wan = float(v) / 10000
    if abs(wan) >= 10000:
        s = f"{wan/10000:.2f}億"
    else:
        s = f"{wan:,.0f}{unit}"
    return s

def fmt_pct(v, digits=1):
    if v is None:
        return "—"
    return f"{float(v)*100:.{digits}f}%"

def level_tag(level):
    if level == "red":
        return '<span class="tag-red">🔴 超限</span>'
    elif level == "orange":
        return '<span class="tag-orange">🟠 警示</span>'
    elif level == "yellow":
        return '<span class="tag-yellow">🟡 警示</span>'
    return '<span class="tag-green">✅ 正常</span>'

def badge(status):
    cls = {"超限":"tag-red","月損失超限":"tag-red","達L2":"tag-red",
           "達L1":"tag-orange",
           "80%提醒":"tag-yellow","接近L1":"tag-yellow",
           "正常":"tag-green"}.get(status, "tag-blue")
    return f'<span class="{cls}">{status or "—"}</span>'

def build_auto_alert_items(data: dict) -> list[dict]:
    m = data["market"]
    wm = data["wm"]

    items = []
    _all_pnl_rows = m.get("ib_rows", []) + m.get("strategy_rows", []) + m.get("trade_rows", [])
    seq = 1

    for r in _all_pnl_rows:
        mp = float(r.get("m_pct") or 0)
        yp = float(r.get("y_pct") or 0)
        dept = r.get("dept", "")

        if mp >= 1.0:
            items.append({
                "id": f"auto_market_m_{seq}",
                "source": "auto",
                "category": "market",
                "text": f"自營 {dept} 月損失超限（{mp*100:.1f}%）",
                "level": "r",
                "enabled": True,
                "sort_order": 10 + seq
            })
            seq += 1
        elif mp >= 0.8:
            items.append({
                "id": f"auto_market_m_{seq}",
                "source": "auto",
                "category": "market",
                "text": f"自營 {dept} 月損失80%提醒（{mp*100:.1f}%）",
                "level": "o",
                "enabled": True,
                "sort_order": 10 + seq
            })
            seq += 1

        if yp >= 1.0:
            items.append({
                "id": f"auto_market_y_{seq}",
                "source": "auto",
                "category": "market",
                "text": f"自營 {dept} 年損失超限（{yp*100:.1f}%）",
                "level": "r",
                "enabled": True,
                "sort_order": 20 + seq
            })
            seq += 1
        elif yp >= 0.8:
            items.append({
                "id": f"auto_market_y_{seq}",
                "source": "auto",
                "category": "market",
                "text": f"自營 {dept} 年損失80%提醒（{yp*100:.1f}%）",
                "level": "o",
                "enabled": True,
                "sort_order": 20 + seq
            })
            seq += 1

    for r in m.get("d3_over", []):
        items.append({
            "id": f'auto_d3_over_{r["code"]}',
            "source": "auto",
            "category": "market",
            "text": f'單檔損失超限 {r["code"]} {r["name"]}（{r["loss_rate"]*100:.1f}%）',
            "level": "r",
            "enabled": True,
            "sort_order": 200
        })

    for r in m.get("d3_warn", []):
        items.append({
            "id": f'auto_d3_warn_{r["code"]}',
            "source": "auto",
            "category": "market",
            "text": f'單檔損失80%提醒 {r["code"]} {r["name"]}（{r["loss_rate"]*100:.1f}%）',
            "level": "o",
            "enabled": True,
            "sort_order": 210
        })

    for v in wm.get("conc", {}).values():
        if v.get("status") in ("達L1", "達L2"):
            items.append({
                "id": f'auto_wm_{v.get("name","")}',
                "source": "auto",
                "category": "wm",
                "text": f'財管 {v.get("name","")} {v.get("status","")}（{(v.get("pct") or 0)*100:.2f}%）',
                "level": "o" if v.get("status") == "達L1" else "r",
                "enabled": True,
                "sort_order": 300
            })

    return items


def merge_alert_items(data: dict) -> list[dict]:
    saved_items = data.get("alert_items", []) or []
    auto_items = build_auto_alert_items(data)

    auto_override_map = {
        x.get("id"): x for x in saved_items
        if x.get("source") == "auto" and x.get("id")
    }

    merged = []
    for item in auto_items:
        override = auto_override_map.get(item["id"])
        if override:
            item["enabled"] = override.get("enabled", item["enabled"])
            item["level"] = override.get("level", item["level"])
            item["sort_order"] = override.get("sort_order", item["sort_order"])
        merged.append(item)

    for item in saved_items:
        if item.get("source") == "manual":
            merged.append(item)

    return sorted(merged, key=lambda x: x.get("sort_order", 9999))


def level_label_to_code(label: str) -> str:
    return {
        "紅燈": "r",
        "橙燈": "o",
        "黃燈": "y",
        "藍燈": "b",
        "綠燈": "g",
    }.get(label, "b")


def level_code_to_label(code: str) -> str:
    return {
        "r": "紅燈",
        "o": "橙燈",
        "y": "黃燈",
        "b": "藍燈",
        "g": "綠燈",
    }.get(code, "藍燈")

# ── 主畫面 ───────────────────────────────────────────────────
st.title("📊 風險管理日報 查詢系統")

dates = load_date_list()
if not dates:
    st.error("找不到資料庫，請先執行 `py main.py` 產生至少一筆報告")
    st.stop()

date_options = [d["date"] for d in dates]

# ── 進入畫面預設設定────────────────────────────────────────────
def set_active_query():
    st.session_state.active_group = "query"
 
def set_active_data():
    st.session_state.active_group = "data"
 
def set_active_report():
    st.session_state.active_group = "report"
 
def set_active_setting():
    st.session_state.active_group = "setting"

# ── Sidebar ──────────────────────────────────────────────────
with st.sidebar:
    st.header("功能選單")
 
    # 只保留一個主群組，避免 4 組 radio 同時都有預設值
    if "main_group" not in st.session_state:
        st.session_state.main_group = "查詢模式"
 
    main_group = st.radio(
        "主功能",
        ["查詢模式", "彙整資料", "報告產出與信件通知", "設定專區"],
        label_visibility="collapsed",
        key="main_group",
    )
 
    st.divider()
 
    if main_group == "查詢模式":
        mode = st.radio(
            "查詢",
            ["📅 單日報告", "⚖️ 雙日比較", "📈 趨勢圖", "🔔 超限事件清單"],
            label_visibility="collapsed",
            key="query_mode",
        )
 
    elif main_group == "彙整資料":
        mode = st.radio(
            "彙整",
            ["🔄 資料轉檔"],
            label_visibility="collapsed",
            key="data_mode",
        )
 
    elif main_group == "報告產出與信件通知":
        mode = st.radio(
            "報告",
            ["⚡ 今日重點說明編輯器", "🧩 報告區塊編輯器", "📄 產出報告", "✉️ 呈報信件"],
            label_visibility="collapsed",
            key="report_mode",
        )
 
    else:
        mode = st.radio(
            "設定",
            ["📁 資料來源路徑", "📁 產出報告路徑", "📧 信件設定"],
            label_visibility="collapsed",
            key="setting_mode",
        )
 
    st.divider()
    st.caption(f"資料庫共 {len(dates)} 筆報告")
    st.caption(f"最新：{dates[0]['date'] if dates else '—'}")
    st.caption(f"最早：{dates[-1]['date'] if dates else '—'}")

# ════════════════════════════════════════════════════════════
#  模式一：單日報告
# ════════════════════════════════════════════════════════════
if mode == "📅 單日報告":
    col1, col2 = st.columns([2, 5])
    with col1:
        sel_date = st.selectbox("選擇日期", date_options,
            format_func=lambda d: f"{d}  {'🔴' if next((x for x in dates if x['date']==d), {}).get('level')=='red' else '🟡' if next((x for x in dates if x['date']==d), {}).get('level')=='yellow' else '✅'}")

    data = load_day(sel_date)
    if not data:
        st.warning("找不到該日資料")
        st.stop()

    m  = data["market"]
    wm = data["wm"]
    b  = data["broker"]

    # 燈號（與 render.py 同源）
    _all_pnl = m.get("ib_rows",[]) + m.get("strategy_rows",[]) + m.get("trade_rows",[])
    _m_over  = sum(1 for r in _all_pnl if float(r.get("m_pct") or 0) >= 1.0)
    _y_over  = sum(1 for r in _all_pnl if float(r.get("y_pct") or 0) >= 1.0)
    _m_warn  = sum(1 for r in _all_pnl if 0.8 <= float(r.get("m_pct") or 0) < 1.0)
    _y_warn  = sum(1 for r in _all_pnl if 0.8 <= float(r.get("y_pct") or 0) < 1.0)
    _sig_market = "red"    if _m_over or _y_over or m["d3_over"] else \
                  "orange" if _m_warn or _y_warn or m["d3_warn"] else "green"
    _sig_wm     = "orange" if any(v.get("status") in ("達L1","達L2") for v in wm["conc"].values()) else "green"
    _broker_maint = float(b.get("total_maint", 0) or 0)
    _sig_broker = "orange" if _broker_maint > 0 and _broker_maint < 160 else "green"

    st.markdown(f"""
    <div style="display:flex;gap:12px;margin:8px 0 16px;align-items:center;">
      <div style="font-size:14px;font-weight:700;color:#4a6080;">資料日期：{sel_date}</div>
      <div>經紀業務：{level_tag(_sig_broker)}</div>
      <div>自營業務：{level_tag(_sig_market)}</div>
      <div>財管商品：{level_tag(_sig_wm)}</div>
    </div>
    """, unsafe_allow_html=True)

    merged_alert_items = merge_alert_items(data)

    with st.expander("⚡ 今日重點說明", expanded=True):
        enabled_items = [x for x in merged_alert_items if x.get("enabled", True)]
        if enabled_items:
            for item in enabled_items:
                tag_html = {
                    "r": '<span class="tag-red">紅燈</span>',
                    "o": '<span class="tag-orange">橙燈</span>',
                    "y": '<span class="tag-yellow">黃燈</span>',
                    "b": '<span class="tag-blue">藍燈</span>',
                    "g": '<span class="tag-green">綠燈</span>',
                }.get(item.get("level", "b"), '<span class="tag-blue">藍燈</span>')

                st.markdown(
                    f"{tag_html}　{item.get('text','')}",
                    unsafe_allow_html=True
                )
        else:
            st.write("✅ 今日各項指標正常")

    tab1, tab2, tab3, tab4 = st.tabs(["01 經紀業務", "02 自營損益", "03 單檔損失", "04 財管集中度"])

    # ── Tab1: 經紀業務 ────────────────────────────────────────
    with tab1:
        c1, c2, c3 = st.columns(3)
        c1.metric("整體融資維持率", f"{b['total_maint']:.1f}%")
        c2.metric("ABC合計", fmt_pct(b["abc_pct"]))
        c3.metric("融資餘額", fmt_wan(b["total_balance"]))

        c4, c5, c6 = st.columns(3)
        c4.metric("整體不限用途借款維持率", f"{b.get('unlim_total_maint',0):.1f}%" if b.get('unlim_total_maint') else "—")
        c5.metric("不限用途 ABC合計", fmt_pct(b.get("unlim_abc_pct")) if b.get("unlim_abc_pct") else "—")
        c6.metric("不限用途借款餘額", fmt_wan(b.get("unlim_total_balance", 0)) if b.get("unlim_total_balance") else "—")

        st.markdown("**融資 A~E 分佈**")
        if b["dist_rows"]:
            dist_df = pd.DataFrame([{
                "等級": r["grade"],
                "比重": fmt_pct(r["pct"]),
                "餘額": fmt_wan(r["balance"]),
                "維持率": f"{r['maint']:.1f}%",
            } for r in b["dist_rows"]])
            st.dataframe(dist_df, hide_index=True, use_container_width=True)

        col_l, col_r = st.columns(2)
        with col_l:
            st.markdown("**融資前5大個股**")
            if b["margin_top5"]:
                m5_df = pd.DataFrame([{
                    "代號": r["code"], "名稱": r["name"], "評等": r.get("grade","—"),
                    "融資(億)": f"{r['balance']/1e8:.2f}",
                    "維持率": f"{r['maint']:.1f}%",
                } for r in b["margin_top5"]])
                st.dataframe(m5_df, hide_index=True, use_container_width=True)

        with col_r:
            st.markdown("**融券前5大個股**")
            if b["short_top5"]:
                s5_df = pd.DataFrame([{
                    "代號": r["code"], "名稱": r["name"], "評等": r.get("grade","—"),
                    "擔保金(億)": f"{r['collat']/1e8:.2f}",
                    "維持率": f"{r['maint']:.1f}%",
                } for r in b["short_top5"]])
                st.dataframe(s5_df, hide_index=True, use_container_width=True)

        st.markdown("**不限用途借貸前5大客戶**")
        if b["unlim_top5"]:
            u5_df = pd.DataFrame([{
                "分公司": r.get("branch","—"),
                "客戶": r["name"],
                "借款(萬)": f"{r['amount']/10000:,.0f}",
                "維持率": f"{r['maint']:.1f}%",
            } for r in b["unlim_top5"]])
            st.dataframe(u5_df, hide_index=True, use_container_width=True)

    # ── Tab2: 自營損益 ────────────────────────────────────────
    with tab2:
        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("**損失超限 / 警示**")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("月損失超限", _m_over)
            c2.metric("年損失超限", _y_over)
            c3.metric("月損失80%提醒", _m_warn)
            c4.metric("年損失80%提醒", _y_warn)
            for r in _all_pnl:
                mp = float(r.get("m_pct") or 0)
                yp = float(r.get("y_pct") or 0)
                if mp >= 1.0:
                    st.error(f"🔴 {r['dept']} 月損失超限（{mp*100:.1f}%）")
                elif mp >= 0.8:
                    st.warning(f"🟠 {r['dept']} 月損失80%提醒（{mp*100:.1f}%）")
                if yp >= 1.0:
                    st.error(f"🔴 {r['dept']} 年損失超限（{yp*100:.1f}%）")
                elif yp >= 0.8:
                    st.warning(f"🟠 {r['dept']} 年損失80%提醒（{yp*100:.1f}%）")

        st.divider()
        col_ib, col_ft = st.columns(2)

        with col_ib:
            st.markdown("**🏦 投資銀行處**")
            ib_data = []
            for r in m["ib_rows"]:
                ib_data.append({
                    "部門/業務": r["dept"],
                    "DTD": fmt_wan(r["dtd"]),
                    "MTD": fmt_wan(r["mtd"]),
                    "YTD": fmt_wan(r["ytd"]),
                    "月損失%": fmt_pct(r.get("m_pct")),
                    "年損失%": fmt_pct(r.get("y_pct")),
                    "狀態": r["status"] or "—",
                })
            ib_data.append({
                "部門/業務": "投資銀行處 合計",
                "DTD": fmt_wan(m["ib_total"]["dtd"]),
                "MTD": fmt_wan(m["ib_total"]["mtd"]),
                "YTD": fmt_wan(m["ib_total"]["ytd"]),
                "月損失%": "—", "年損失%": "—", "狀態": "—",
            })
            st.dataframe(pd.DataFrame(ib_data), hide_index=True, use_container_width=True)

        with col_ft:
            st.markdown("**📊 金融交易處**")
            ft_data = []
            for r in m["strategy_rows"] + [m["strategy_total"]]:
                ft_data.append({
                    "部門/業務": r["dept"],
                    "DTD": fmt_wan(r["dtd"]),
                    "MTD": fmt_wan(r["mtd"]),
                    "YTD": fmt_wan(r["ytd"]),
                    "月損失%": fmt_pct(r.get("m_pct")),
                    "年損失%": fmt_pct(r.get("y_pct")),
                    "狀態": r["status"] or "—",
                })
            for r in m["trade_rows"] + [m["trade_total"], m["ft_total"]]:
                ft_data.append({
                    "部門/業務": r["dept"],
                    "DTD": fmt_wan(r["dtd"]),
                    "MTD": fmt_wan(r["mtd"]),
                    "YTD": fmt_wan(r["ytd"]),
                    "月損失%": fmt_pct(r.get("m_pct")),
                    "年損失%": fmt_pct(r.get("y_pct")),
                    "狀態": r["status"] or "—",
                })
            st.dataframe(pd.DataFrame(ft_data), hide_index=True, use_container_width=True)

    # ── Tab3: 單檔損失 ────────────────────────────────────────
    with tab3:
        c1, c2 = st.columns(2)
        c1.metric("單檔超限", len(m["d3_over"]))
        c2.metric("單檔80%提醒", len(m["d3_warn"]))
        if m["d3_top5"]:
            d3_df = pd.DataFrame([{
                "代號": r["code"], "名稱": r["name"],
                "未實現損益": fmt_wan(r["pnl"]),
                "損失率": fmt_pct(r["loss_rate"]),
                "狀態": r["status"] or "觀察",
            } for r in m["d3_top5"]])
            st.dataframe(d3_df, hide_index=True, use_container_width=True)

    # ── Tab4: 財管集中度 ──────────────────────────────────────
    with tab4:
        alloc = wm["alloc"]
        c1, c2, c3 = st.columns(3)
        c1.metric("海外債券", f"{alloc.get('bond',0)*100:.1f}%")
        c2.metric("基金商品", f"{alloc.get('fund',0)*100:.1f}%")
        c3.metric("結構型商品", f"{alloc.get('struct',0)*100:.1f}%")

        st.markdown("**集中度明細**")
        cat_labels = {
            "bond_inv": "債券｜投資等級",
            "bond_noninv": "債券｜非投資等級",
            "fund": "基金｜單一標的",
            "struct_target": "結構型｜連結標的",
            "struct_upper": "結構型上手｜BBB+(含)以上",
            "struct_lower": "結構型上手｜投資等級下緣",
        }
        conc_data = []
        for k, label in cat_labels.items():
            v = wm["conc"].get(k, {})
            conc_data.append({
                "類別": label,
                "最大集中標的": v.get("name","—"),
                "集中度": fmt_pct(v.get("pct")),
                "L1": fmt_pct(v.get("l1")),
                "L2": fmt_pct(v.get("l2")),
                "狀態": v.get("status","—"),
            })
        st.dataframe(pd.DataFrame(conc_data), hide_index=True, use_container_width=True)

        st.markdown("**高資產客戶**")
        ha = wm["ha"]
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("客戶人數", f"{int(ha.get('count',0))} 人")
        c2.metric("投資總額", f"{ha.get('total',0)/1e8:.1f} 億")
        c3.metric("BB-(含)以下債券", f"{int(ha.get('bb_count',0))} 人")
        c4.metric("境外非投信基金", f"{int(ha.get('offshore_count',0))} 人")


# ════════════════════════════════════════════════════════════
#  模式二：雙日比較
# ════════════════════════════════════════════════════════════
elif mode == "⚖️ 雙日比較":
    st.subheader("雙日比較")
    col1, col2 = st.columns(2)
    with col1:
        date_a = st.selectbox("日期 A（基準）", date_options, key="da")
    with col2:
        date_b = st.selectbox("日期 B（比較）", date_options, index=min(1, len(date_options)-1), key="db")

    if date_a == date_b:
        st.warning("請選擇不同的兩個日期")
        st.stop()

    da = load_day(date_a)
    db_ = load_day(date_b)
    if not da or not db_:
        st.error("其中一天找不到資料")
        st.stop()

    st.markdown(f"**{date_a}（A）vs {date_b}（B）**")

    # 融資維持率比較
    st.markdown("#### 融資維持率比較")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric(f"維持率 {date_a}", f"{da['broker']['total_maint']:.1f}%")
    c2.metric(f"維持率 {date_b}", f"{db_['broker']['total_maint']:.1f}%",
              delta=f"{db_['broker']['total_maint'] - da['broker']['total_maint']:+.1f}%")
    c3.metric(f"ABC合計 {date_a}", fmt_pct(da["broker"]["abc_pct"]))
    c4.metric(f"ABC合計 {date_b}", fmt_pct(db_["broker"]["abc_pct"]),
              delta=f"{(db_['broker']['abc_pct'] - da['broker']['abc_pct'])*100:+.2f}%")

    # ── 融資前5大個股比較 ──────────────────────────────────────
    st.markdown("#### 融資前5大個股比較")
    def _top5_cmp(list_a, list_b, bal_key, bal_unit, bal_fmt):
        codes_a = {r["code"]: r for r in (list_a or [])}
        codes_b = {r["code"]: r for r in (list_b or [])}
        all_codes = list(dict.fromkeys(
            [r["code"] for r in (list_a or [])] +
            [r["code"] for r in (list_b or [])]
        ))
        rows = []
        for code in all_codes:
            ra = codes_a.get(code)
            rb = codes_b.get(code)
            name = (ra or rb).get("name", "")
            if ra and rb:
                flag = "—"
            elif rb:
                flag = "🆕 新增"
            else:
                flag = "🗑 刪除"
            bal_a = bal_fmt(ra[bal_key]) if ra else "—"
            bal_b = bal_fmt(rb[bal_key]) if rb else "—"
            if ra and rb:
                delta_bal = bal_fmt(rb[bal_key] - ra[bal_key])
            else:
                delta_bal = "—"
            maint_a = f"{ra['maint']:.1f}%" if ra else "—"
            maint_b = f"{rb['maint']:.1f}%" if rb else "—"
            if ra and rb:
                delta_maint = f"{rb['maint'] - ra['maint']:+.1f}%"
            else:
                delta_maint = "—"
            rows.append({
                "代號": code, "名稱": name,
                f"{bal_unit}({date_a})": bal_a,
                f"{bal_unit}({date_b})": bal_b,
                f"{bal_unit}變動": delta_bal,
                f"維持率({date_a})": maint_a,
                f"維持率({date_b})": maint_b,
                "維持率變動": delta_maint,
                "異動": flag,
            })
        return rows

    margin_cmp = _top5_cmp(
        da["broker"].get("margin_top5"), db_["broker"].get("margin_top5"),
        "balance", "融資餘額(億)",
        lambda v: f"{v/1e8:.2f}"
    )
    if margin_cmp:
        st.dataframe(pd.DataFrame(margin_cmp), hide_index=True, use_container_width=True)
    else:
        st.info("兩日均無融資前5大資料")

    # ── 融券前5大個股比較 ──────────────────────────────────────
    st.markdown("#### 融券前5大個股比較")
    short_cmp = _top5_cmp(
        da["broker"].get("short_top5"), db_["broker"].get("short_top5"),
        "collat", "擔保金(億)",
        lambda v: f"{v/1e8:.2f}"
    )
    if short_cmp:
        st.dataframe(pd.DataFrame(short_cmp), hide_index=True, use_container_width=True)
    else:
        st.info("兩日均無融券前5大資料")

    # 自營損益比較
    st.markdown("#### 自營損益比較（萬元）")
    rows_a = {r["dept"]: r for r in da["market"]["ib_rows"] + da["market"]["trade_rows"]}
    rows_b = {r["dept"]: r for r in db_["market"]["ib_rows"] + db_["market"]["trade_rows"]}
    all_depts = list(dict.fromkeys(list(rows_a.keys()) + list(rows_b.keys())))

    cmp_data = []
    for dept in all_depts:
        ra = rows_a.get(dept, {})
        rb = rows_b.get(dept, {})
        mtd_a = float(ra.get("mtd") or 0) / 10000
        mtd_b = float(rb.get("mtd") or 0) / 10000
        diff = mtd_b - mtd_a
        cmp_data.append({
            "部門": dept,
            f"MTD {date_a}(萬)": f"{mtd_a:,.0f}",
            f"MTD {date_b}(萬)": f"{mtd_b:,.0f}",
            "變動(萬)": f"{diff:+,.0f}",
        })
    st.dataframe(pd.DataFrame(cmp_data), hide_index=True, use_container_width=True)

    # 財管集中度比較
    st.markdown("#### 財管集中度比較")
    cat_labels = {
        "bond_inv": "債券｜投資等級",
        "bond_noninv": "債券｜非投資等級",
        "fund": "基金｜單一標的",
        "struct_target": "結構型｜連結標的",
        "struct_upper": "結構型上手｜BBB+以上",
        "struct_lower": "結構型上手｜下緣",
    }
    conc_cmp = []
    for k, label in cat_labels.items():
        va = da["wm"]["conc"].get(k, {})
        vb = db_["wm"]["conc"].get(k, {})
        pct_a = float(va.get("pct") or 0) * 100
        pct_b = float(vb.get("pct") or 0) * 100
        conc_cmp.append({
            "類別": label,
            f"{date_a}": f"{pct_a:.2f}%",
            f"{date_b}": f"{pct_b:.2f}%",
            "變動": f"{pct_b - pct_a:+.2f}%",
            f"狀態({date_b})": vb.get("status","—"),
        })
    st.dataframe(pd.DataFrame(conc_cmp), hide_index=True, use_container_width=True)

    # 融資維持率比較
    st.markdown("#### 融資維持率比較")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric(f"維持率 {date_a}", f"{da['broker']['total_maint']:.1f}%")
    c2.metric(f"維持率 {date_b}", f"{db_['broker']['total_maint']:.1f}%",
              delta=f"{db_['broker']['total_maint'] - da['broker']['total_maint']:+.1f}%")
    c3.metric(f"ABC合計 {date_a}", fmt_pct(da["broker"]["abc_pct"]))
    c4.metric(f"ABC合計 {date_b}", fmt_pct(db_["broker"]["abc_pct"]),
              delta=f"{(db_['broker']['abc_pct'] - da['broker']['abc_pct'])*100:+.2f}%")

    # ── 融資前5大個股比較 ──────────────────────────────────────
    st.markdown("#### 融資前5大個股比較")
    def _top5_cmp(list_a, list_b, bal_key, bal_unit, bal_fmt):
        codes_a = {r["code"]: r for r in (list_a or [])}
        codes_b = {r["code"]: r for r in (list_b or [])}
        all_codes = list(dict.fromkeys(
            [r["code"] for r in (list_a or [])] +
            [r["code"] for r in (list_b or [])]
        ))
        rows = []
        for code in all_codes:
            ra = codes_a.get(code)
            rb = codes_b.get(code)
            name = (ra or rb).get("name", "")
            if ra and rb:
                flag = "—"
            elif rb:
                flag = "🆕 新增"
            else:
                flag = "🗑 刪除"
            bal_a = bal_fmt(ra[bal_key]) if ra else "—"
            bal_b = bal_fmt(rb[bal_key]) if rb else "—"
            if ra and rb:
                delta_bal = bal_fmt(rb[bal_key] - ra[bal_key])
            else:
                delta_bal = "—"
            maint_a = f"{ra['maint']:.1f}%" if ra else "—"
            maint_b = f"{rb['maint']:.1f}%" if rb else "—"
            if ra and rb:
                delta_maint = f"{rb['maint'] - ra['maint']:+.1f}%"
            else:
                delta_maint = "—"
            rows.append({
                "代號": code, "名稱": name,
                f"{bal_unit}({date_a})": bal_a,
                f"{bal_unit}({date_b})": bal_b,
                f"{bal_unit}變動": delta_bal,
                f"維持率({date_a})": maint_a,
                f"維持率({date_b})": maint_b,
                "維持率變動": delta_maint,
                "異動": flag,
            })
        return rows

    margin_cmp = _top5_cmp(
        da["broker"].get("margin_top5"), db_["broker"].get("margin_top5"),
        "balance", "融資餘額(億)",
        lambda v: f"{v/1e8:.2f}"
    )
    if margin_cmp:
        st.dataframe(pd.DataFrame(margin_cmp), hide_index=True, use_container_width=True)
    else:
        st.info("兩日均無融資前5大資料")

    # ── 融券前5大個股比較 ──────────────────────────────────────
    st.markdown("#### 融券前5大個股比較")
    short_cmp = _top5_cmp(
        da["broker"].get("short_top5"), db_["broker"].get("short_top5"),
        "collat", "擔保金(億)",
        lambda v: f"{v/1e8:.2f}"
    )
    if short_cmp:
        st.dataframe(pd.DataFrame(short_cmp), hide_index=True, use_container_width=True)
    else:
        st.info("兩日均無融券前5大資料")


# ════════════════════════════════════════════════════════════
#  模式三：趨勢圖
# ════════════════════════════════════════════════════════════
elif mode == "📈 趨勢圖":
    st.subheader("歷史趨勢圖")

    col1, col2, col3 = st.columns(3)
    with col1:
        chart_type = st.selectbox("指標類型", ["經紀業務", "自營損益", "財管集中度"])
    with col2:
        if len(date_options) >= 2:
            start_d = st.selectbox("起始日期", date_options, index=len(date_options)-1)
        else:
            start_d = date_options[-1]
    with col3:
        end_d = st.selectbox("結束日期", date_options, index=0)

    if start_d > end_d:
        start_d, end_d = end_d, start_d

    if chart_type == "經紀業務":
        broker_cat = st.selectbox("選擇類別", ["融資業務", "不限用途業務"])

        broker_charts = {
            "融資業務": [
                "全公司融資餘額趨勢",
                "全公司融資維持率趨勢",
                "融資A~E級比重趨勢",
            ],
            "不限用途業務": [
                "全公司不限用途借款餘額趨勢",
                "全公司不限用途借款維持率趨勢",
                "不限用途借款擔保品A~E級比重趨勢",
            ],
        }
        broker_chart_sel = st.selectbox("選擇圖表", broker_charts[broker_cat])

        df = load_broker_trend(start_d, end_d)
        if df.empty:
            st.info("該期間無資料")
        else:
            colors_grade = {"A":"#1a9e6a","B":"#1976d2","C":"#b45309","D":"#d97706","E":"#c62828"}

            if broker_chart_sel == "全公司融資餘額趨勢":
                df["bal_yi"] = df["total_balance"].astype(float) / 1e8
                fig = go.Figure()
                fig.add_trace(go.Scatter(x=df["date"], y=df["bal_yi"],
                    mode="lines+markers", name="融資餘額(億)",
                    line=dict(color="#1976d2", width=2), marker=dict(size=6)))
                fig.update_layout(title="全公司融資餘額趨勢", xaxis_title="日期",
                    yaxis_title="億元", height=400, margin=dict(t=40, b=20))
                st.plotly_chart(fig, use_container_width=True)

            elif broker_chart_sel == "全公司融資維持率趨勢":
                fig = go.Figure()
                fig.add_trace(go.Scatter(x=df["date"], y=df["total_maint"].astype(float),
                    mode="lines+markers", name="整體維持率(%)",
                    line=dict(color="#1976d2", width=2), marker=dict(size=6)))
                fig.add_hline(y=130, line_dash="dash", line_color="#c62828",
                    opacity=0.6, annotation_text="追繳線 130%")
                fig.update_layout(title="全公司融資維持率趨勢", xaxis_title="日期",
                    yaxis_title="%", height=400, margin=dict(t=40, b=20))
                st.plotly_chart(fig, use_container_width=True)

            elif broker_chart_sel == "融資A~E級比重趨勢":
                fig = go.Figure()
                for g in ["A","B","C","D","E"]:
                    fig.add_trace(go.Scatter(x=df["date"], y=df[g].astype(float)*100,
                        mode="lines", name=f"{g}級",
                        stackgroup="one", line=dict(color=colors_grade[g])))
                fig.update_layout(title="融資 A~E 等級比重趨勢", xaxis_title="日期",
                    yaxis_title="%", height=400, margin=dict(t=40, b=20))
                st.plotly_chart(fig, use_container_width=True)

            elif broker_chart_sel == "全公司不限用途借款餘額趨勢":
                df["ubal_yi"] = df["unlim_total_balance"].astype(float) / 1e8
                fig = go.Figure()
                fig.add_trace(go.Scatter(x=df["date"], y=df["ubal_yi"],
                    mode="lines+markers", name="不限用途借款餘額(億)",
                    line=dict(color="#7c4dff", width=2), marker=dict(size=6)))
                fig.update_layout(title="全公司不限用途借款餘額趨勢", xaxis_title="日期",
                    yaxis_title="億元", height=400, margin=dict(t=40, b=20))
                st.plotly_chart(fig, use_container_width=True)

            elif broker_chart_sel == "全公司不限用途借款維持率趨勢":
                fig = go.Figure()
                fig.add_trace(go.Scatter(x=df["date"], y=df["unlim_total_maint"].astype(float),
                    mode="lines+markers", name="不限用途整體維持率(%)",
                    line=dict(color="#7c4dff", width=2), marker=dict(size=6)))
                fig.add_hline(y=130, line_dash="dash", line_color="#c62828",
                    opacity=0.6, annotation_text="追繳線 130%")
                fig.update_layout(title="全公司不限用途借款維持率趨勢", xaxis_title="日期",
                    yaxis_title="%", height=400, margin=dict(t=40, b=20))
                st.plotly_chart(fig, use_container_width=True)

            elif broker_chart_sel == "不限用途借款擔保品A~E級比重趨勢":
                fig = go.Figure()
                for g, col in [("A","uA"),("B","uB"),("C","uC"),("D","uD"),("E","uE")]:
                    fig.add_trace(go.Scatter(x=df["date"], y=df[col].astype(float)*100,
                        mode="lines", name=f"{g}級",
                        stackgroup="one", line=dict(color=colors_grade[g])))
                fig.update_layout(title="不限用途借款擔保品 A~E 等級比重趨勢",
                    xaxis_title="日期", yaxis_title="%",
                    height=400, margin=dict(t=40, b=20))
                st.plotly_chart(fig, use_container_width=True)

    elif chart_type == "自營損益":
        # 直接從 DB 讀出所有存在的 dept + biz 組合，避免硬寫字串與實際資料不符
        with db_conn() as _conn:
            _pairs = _conn.execute(
                "SELECT DISTINCT dept, biz FROM market_pnl ORDER BY biz, dept"
            ).fetchall()

        if not _pairs:
            st.info("資料庫尚無自營損益資料，請先執行轉資料。")
        else:
            # 顯示名稱：「biz | dept」
            dept_biz_options = {f"{biz_} | {dept_}": (dept_, biz_)
                                for dept_, biz_ in _pairs}
            sel = st.selectbox("選擇部門", list(dept_biz_options.keys()))
            dept, biz = dept_biz_options[sel]
            metric = st.radio("損益類型", ["MTD", "YTD"], horizontal=True)

            df = load_pnl_trend(dept, biz, start_d, end_d)
            if df.empty:
                st.info("該期間無資料")
            else:
                col_map = {"MTD": "mtd", "YTD": "ytd"}
                y_col = col_map[metric]
                df[y_col] = df[y_col].astype(float) / 10000

                fig = go.Figure()
                fig.add_trace(go.Scatter(
                    x=df["date"], y=df[y_col],
                    mode="lines+markers",
                    name=f"{sel} {metric}",
                    line=dict(color="#1976d2", width=2),
                    marker=dict(size=6),
                ))
                fig.add_hline(y=0, line_dash="dash", line_color="gray", opacity=0.5)
                fig.update_layout(
                    title=f"{sel} — {metric} 損益趨勢（萬元）",
                    xaxis_title="日期", yaxis_title="萬元",
                    height=400, margin=dict(t=40, b=20),
                )
                st.plotly_chart(fig, use_container_width=True)

    elif chart_type == "財管集中度":
        cat_options = {
            "債券｜投資等級": "bond_inv",
            "債券｜非投資等級": "bond_noninv",
            "基金｜單一標的": "fund",
            "結構型｜連結標的": "struct_target",
            "結構型上手｜BBB+以上": "struct_upper",
        }
        sel = st.selectbox("選擇類別", list(cat_options.keys()))
        cat = cat_options[sel]
        df = load_conc_trend(cat, start_d, end_d)
        if df.empty:
            st.info("該期間無資料")
        else:
            df["pct"] = df["pct"].astype(float) * 100
            df["l1"]  = df["l1"].astype(float) * 100
            df["l2"]  = df["l2"].astype(float) * 100

            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=df["date"], y=df["pct"],
                mode="lines+markers", name="集中度",
                line=dict(color="#1976d2", width=2), marker=dict(size=6),
            ))
            if df["l1"].iloc[0] > 0:
                fig.add_trace(go.Scatter(
                    x=df["date"], y=df["l1"],
                    mode="lines", name="L1警示點",
                    line=dict(color="#f59e0b", dash="dash", width=1.5),
                ))
            if df["l2"].iloc[0] > 0:
                fig.add_trace(go.Scatter(
                    x=df["date"], y=df["l2"],
                    mode="lines", name="L2警示點",
                    line=dict(color="#c62828", dash="dash", width=1.5),
                ))
            fig.update_layout(
                title=f"{sel} — 集中度趨勢（%）",
                xaxis_title="日期", yaxis_title="%",
                height=400, margin=dict(t=40, b=20),
            )
            st.plotly_chart(fig, use_container_width=True)



# ════════════════════════════════════════════════════════════
#  模式四：超限事件清單
# ════════════════════════════════════════════════════════════
elif mode == "🔔 超限事件清單":
    st.subheader("超限 / 警示事件清單")

    col1, col2, col3 = st.columns(3)
    with col1:
        start_d = st.selectbox("起始日期", date_options, index=len(date_options)-1)
    with col2:
        end_d = st.selectbox("結束日期", date_options, index=0)
    with col3:
        filter_type = st.selectbox("篩選類型", ["全部", "red（超限）", "yellow（警示）"])

    if start_d > end_d:
        start_d, end_d = end_d, start_d

    df = load_alert_events(start_d, end_d)
    if df.empty:
        st.info("該期間無超限/警示事件")
    else:
        if filter_type == "red（超限）":
            df = df[df["類型"] == "red"]
        elif filter_type == "yellow（警示）":
            df = df[df["類型"] == "yellow"]

        st.markdown(f"共 **{len(df)}** 筆事件")

        # 超限事件計數長條圖
        if not df.empty:
            count_df = df.groupby(["日期","類型"]).size().reset_index(name="件數")
            fig = px.bar(count_df, x="日期", y="件數", color="類型",
                         color_discrete_map={"red":"#c62828","yellow":"#f59e0b"},
                         title="每日超限/警示件數", height=300)
            fig.update_layout(margin=dict(t=40,b=20))
            st.plotly_chart(fig, use_container_width=True)

        st.dataframe(df[["日期","類型","說明"]], hide_index=True, use_container_width=True)


# ════════════════════════════════════════════════════════════
#  模式五：🔄 轉資料
# ════════════════════════════════════════════════════════════
elif mode == "🔄 資料轉檔":
    st.subheader("🔄 資料轉檔")
    st.markdown("執行 Excel → SQLite 資料轉換，等同於 `py main.py YYYYMMDD`。")

    sel_date = st.date_input("選擇資料日期", value=date.today())
    date_str = sel_date.strftime("%Y%m%d")

    # 檢查來源檔案是否存在
    from extract import find_file, find_broker2_file
    file_status = {}
    for label, directory, prefix in [
        ("市場風險（自營）", config.MARKET_DIR, config.MARKET_PREFIX),
        ("財管商品集中度",   config.WM_DIR,     config.WM_PREFIX),
        ("經紀業務當日作業（主檔）", config.BROKER_DIR, config.BROKER_PREFIX),
    ]:
        try:
            f = find_file(Path(directory), prefix, sel_date)
            file_status[label] = ("✅", str(f.name), True, False)
        except FileNotFoundError:
            file_status[label] = ("❌", "找不到檔案", False, False)

    # 第二個 broker 檔（追繳處分）：選填，找不到只顯示警告不擋路
    b2 = find_broker2_file(Path(config.BROKER2_DIR), sel_date)
    if b2:
        file_status["追繳及處分金額彙總表（選填）"] = ("✅", str(b2.name), True, True)
    else:
        file_status["追繳及處分金額彙總表（選填）"] = ("⚠️", "找不到，追繳/處分欄位將顯示 —", True, True)

    st.markdown("**檔案狀態檢查：**")
    all_ok = True
    for label, (icon, fname, ok, optional) in file_status.items():
        st.markdown(f"{icon} **{label}**　`{fname}`")
        if not ok and not optional:
            all_ok = False

    st.divider()
    if not all_ok:
        st.error("⚠️ 有檔案缺少，請確認檔案已放入對應資料夾後再執行。")
    else:
        if st.button("▶ 執行轉資料", type="primary"):
            with st.spinner(f"執行中：main.py {date_str} ..."):
                try:
                    result = subprocess.run(
                        [sys.executable, "main.py", date_str],
                        capture_output=True, text=True, encoding="utf-8",
                        errors="replace", timeout=120,
                    )
                    if result.returncode == 0:
                        st.success("✅ 轉資料完成！")
                        st.code(result.stdout, language="text")
                        st.cache_data.clear()   # 清快取，讓日期列表更新
                    else:
                        st.error("❌ 執行失敗")
                        st.code(result.stderr or result.stdout, language="text")
                except subprocess.TimeoutExpired:
                    st.error("❌ 執行逾時（超過 120 秒）")
                except Exception as e:
                    st.error(f"❌ 執行錯誤：{e}")


# ════════════════════════════════════════════════════════════
#  模式六：📄 產出報告
# ════════════════════════════════════════════════════════════
elif mode == "📄 產出報告":
    st.subheader("📄 產出 HTML 報告")
    st.markdown("從資料庫讀取已轉好的資料，重新產出 HTML 報告。")

    if not date_options:
        st.warning("資料庫尚無資料，請先執行「🔄 資料轉檔」。")
    else:
        sel_date = st.selectbox("選擇日期", date_options)

        if st.button("▶ 產出 HTML", type="primary"):
            from db import load_report, load_custom_sections
            from render import generate_html, save_html

            data = load_report(config.DB_PATH, sel_date)
            custom_sections = load_custom_sections(config.DB_PATH, sel_date)
            if not data:
                st.error("找不到該日資料")
            else:
                with st.spinner("產出中..."):
                    try:
                        html = generate_html(data, custom_sections=custom_sections)
                        out_path = save_html(html, config.OUTPUT_DIR, data["report_date"])
                        st.success(f"✅ 已產出：`{out_path}`")
                        with open(out_path, "r", encoding="utf-8") as f:
                            html_content = f.read()
                        st.download_button(
                            label="⬇ 下載 HTML",
                            data=html_content,
                            file_name=out_path.name,
                            mime="text/html",
                        )
                    except Exception as e:
                        st.error(f"❌ 產出失敗：{e}")
# ════════════════════════════════════════════════════════════
# ⚡ 今日重點說明編輯器
# ════════════════════════════════════════════════════════════
elif mode == "⚡ 今日重點說明編輯器":
    st.subheader("⚡ 今日重點說明編輯器")

    sel_date = st.selectbox("選擇報告日期", date_options, key="alert_editor_date")
    data = load_day(sel_date)

    if not data:
        st.warning("找不到該日資料")
        st.stop()

    merged_items = merge_alert_items(data)

    st.markdown("### 系統 / 人工重點項目")
    edited_rows = []

    for idx, item in enumerate(merged_items):
        with st.container(border=True):
            c1, c2, c3 = st.columns([1, 1, 5])
            with c1:
                enabled = st.checkbox("納入", value=item.get("enabled", True), key=f"alert_enabled_{idx}")
            with c2:
                level_label = st.selectbox(
                    "燈號",
                    ["紅燈", "橙燈", "黃燈", "藍燈", "綠燈"],
                    index=["紅燈", "橙燈", "黃燈", "藍燈", "綠燈"].index(level_code_to_label(item.get("level", "b"))),
                    key=f"alert_level_{idx}"
                )
            with c3:
                if item.get("source") == "manual":
                    text = st.text_input("內容", value=item.get("text", ""), key=f"alert_text_{idx}")
                else:
                    text = item.get("text", "")
                    st.text_input("內容", value=text, disabled=True, key=f"alert_text_{idx}")

            sort_order = st.number_input("排序", min_value=1, value=int(item.get("sort_order", 9999)), step=10, key=f"alert_sort_{idx}")

            edited_rows.append({
                "id": item.get("id"),
                "source": item.get("source", "manual"),
                "category": item.get("category", "manual"),
                "text": text,
                "level": level_label_to_code(level_label),
                "enabled": enabled,
                "sort_order": int(sort_order),
            })

    st.divider()
    st.markdown("### 人工新增重點")
    manual_text = st.text_area("新增內容", key="new_manual_alert_text", placeholder="請輸入人工補充說明")
    manual_level = st.selectbox("燈號", ["紅燈", "橙燈", "黃燈", "藍燈", "綠燈"], key="new_manual_alert_level")
    manual_sort = st.number_input("排序", min_value=1, value=900, step=10, key="new_manual_alert_sort")

    if st.button("➕ 新增人工重點"):
        if manual_text.strip():
            new_item = {
                "id": f"manual_{uuid.uuid4().hex[:8]}",
                "source": "manual",
                "category": "manual",
                "text": manual_text.strip(),
                "level": level_label_to_code(manual_level),
                "enabled": True,
                "sort_order": int(manual_sort),
            }
            current = data.get("alert_items", []) or []
            current.append(new_item)
            save_alert_items(config.DB_PATH, sel_date, current)
            st.success("✅ 已新增人工重點，請重新整理頁面查看。")
        else:
            st.warning("請先輸入人工重點內容")

    if st.button("💾 儲存今日重點設定", type="primary"):
        save_alert_items(config.DB_PATH, sel_date, edited_rows)
        st.success("✅ 已儲存今日重點說明設定")

# ════════════════════════════════════════════════════════════
#  🧩 報告區塊編輯器
# ════════════════════════════════════════════════════════════
elif mode == "🧩 報告區塊編輯器":
    st.subheader("🧩 報告區塊編輯器")

    if not date_options:
        st.warning("資料庫尚無資料，請先執行「🔄 資料轉檔」。")
        st.stop()

    sel_date = st.selectbox("選擇報告日期", date_options, key="custom_section_date")
    report_date_db = sel_date  # 這裡 date_options 本來就是 YYYY-MM-DD

    sections = load_custom_sections(config.DB_PATH, report_date_db)
    
    st.markdown("### 選擇既有區塊 / 新增區塊")

    sec_map = {f"{s['display_order']}｜{s['title']}": s for s in sections}

    selected_section_label = st.selectbox(
        "選擇既有區塊",
        ["(新增區塊)"] + list(sec_map.keys()),
        key="edit_custom_section_select"
    )

    if selected_section_label != "(新增區塊)":
        editing_section = sec_map[selected_section_label]
    else:
        editing_section = None

    templates = load_section_templates()

    st.markdown("### 新增來源")
    create_mode = st.radio(
        "新增方式",
        ["空白新增", "從模板新增"],
        horizontal=True,
        key="section_create_mode"
    )

    selected_template = None
    if not editing_section and create_mode == "從模板新增":
        if templates:
            selected_template_name = st.selectbox(
                "選擇模板",
                [t["template_name"] for t in templates],
                key="section_template_select"
            )
            selected_template = next(
                (t for t in templates if t["template_name"] == selected_template_name),
                None
            )
        else:
            st.info("目前尚無模板，請先建立 section_templates.json 或先從現有區塊另存模板。")

    st.markdown("### 既有區塊")
    if sections:
        list_df = pd.DataFrame([
            {
                "順序": s["display_order"],
                "啟用": s["enabled"],
                "類型": s["section_type"],
                "標題": s["title"],
                "新頁": s["page_break_before"],
                "section_id": s["section_id"],
            }
            for s in sections
        ])
        st.dataframe(list_df, hide_index=True, use_container_width=True)
    else:
        st.info("目前尚無自訂區塊。")

    st.divider()
    st.markdown("### 新增 / 編輯區塊")

    if editing_section:
        title_default = editing_section["title"]
        section_type_default = editing_section["section_type"]
        display_order_default = editing_section["display_order"]
        enabled_default = editing_section["enabled"]
        layout_mode_default = editing_section.get("layout_mode", "full_page")
        page_break_default = editing_section.get("page_break_before", False)
        insert_after_default = editing_section.get("insert_after", "appendix")
        section_id = editing_section["section_id"]
        content_default = editing_section["content"]

    elif selected_template:
        new_section = instantiate_section_from_template(selected_template, display_order=100)
        title_default = new_section["title"]
        section_type_default = new_section["section_type"]
        display_order_default = new_section.get("display_order", 100)
        enabled_default = new_section.get("enabled", True)
        layout_mode_default = new_section.get("layout_mode", "inline")
        page_break_default = new_section.get("page_break_before", False)
        insert_after_default = new_section.get("insert_after", "appendix")
        section_id = new_section["section_id"]
        content_default = new_section.get("content", {})

    else:
        title_default = ""
        section_type_default = "text"
        display_order_default = 100
        enabled_default = True
        layout_mode_default = "full_page"
        page_break_default = False
        insert_after_default = "appendix"
        section_id = str(uuid.uuid4())[:8]
        content_default = {}

    title = st.text_input("區塊標題", value=title_default)

    section_type_options = ["text", "bullets", "table"]
    section_type = st.selectbox(
        "區塊類型",
        section_type_options,
        index=section_type_options.index(section_type_default)
    )

    display_order = st.number_input(
        "顯示順序",
        min_value=1,
        step=10,
        value=int(display_order_default)
    )

    enabled = st.checkbox("納入本次報告", value=enabled_default)

    layout_mode_options = ["full_page", "inline"]
    layout_mode = st.selectbox(
        "版面模式",
        layout_mode_options,
        index=layout_mode_options.index(layout_mode_default),
        format_func=lambda x: "獨立頁" if x == "full_page" else "接續顯示"
    )

    insert_after_options = ["summary", "market", "wm", "broker", "appendix"]
    insert_after = st.selectbox(
        "插入位置",
        insert_after_options,
        index=insert_after_options.index(insert_after_default),
        format_func=lambda x: {
            "summary": "總覽後",
            "market": "自營後",
            "wm": "財管後",
            "broker": "經紀後",
            "appendix": "附加區"
        }[x]
    )

    page_break_before = st.checkbox(
        "此區塊前強制換頁（僅接續顯示時有意義）",
        value=page_break_default
    )
    

    content = {}

    if section_type == "text":
        text_value = st.text_area(
            "內容",
            value=content_default.get("text", ""),
            height=220,
            placeholder="請輸入一般說明文字，可分段。"
        )
        content = {"text": text_value}

    elif section_type == "bullets":
        bullet_default = "\n".join(content_default.get("items", []))
        bullet_text = st.text_area(
            "條列內容（每行一點）",
            value=bullet_default,
            height=180,
            placeholder="黃金ETF Vega 使用率偏高\n白銀ETF 波動加劇\n建議提高盤中監控頻率"
        )
        items = [x.strip() for x in bullet_text.splitlines() if x.strip()]
        content = {"items": items}

    elif section_type == "table":
        default_cols = content_default.get("columns", ["項目", "數值", "狀態"])
        columns_text = st.text_input(
            "欄位名稱（以逗號分隔）",
            value=",".join(default_cols)
        )
        cols = [c.strip() for c in columns_text.split(",") if c.strip()]
        if not cols:
            cols = ["欄位1", "欄位2"]
        default_rows = content_default.get("rows", [])
        if default_rows:
            default_table_df = pd.DataFrame(default_rows, columns=cols)
        else:
            default_table_df = pd.DataFrame(columns=cols)

        table_df = st.data_editor(
            default_table_df,
            num_rows="dynamic",
            use_container_width=True,
            key="custom_table_editor"
        )
        content = {
            "columns": cols,
            "rows": table_df.fillna("").values.tolist()
        }

    col1, col2 = st.columns(2)
    with col1:
        if st.button("💾 儲存區塊", type="primary"):
            if not title.strip():
                st.error("請輸入區塊標題")
            else:
                section = {
                    "section_id": section_id,
                    "title": title.strip(),
                    "section_type": section_type,
                    "content": content,
                    "display_order": int(display_order),
                    "enabled": enabled,
                    "layout_mode": layout_mode,
                    "page_break_before": page_break_before,
                    "insert_after": insert_after,
                }
                save_custom_section(config.DB_PATH, report_date_db, section)
                if editing_section:
                    st.success("✅ 已更新區塊")
                else:
                    st.success("✅ 已新增區塊")

    with col2:
        if sections:
            del_target = st.selectbox(
                "選擇要刪除的區塊",
                options=sections,
                format_func=lambda s: f"{s['display_order']}｜{s['title']}",
                key="delete_custom_section_target"
            )
            if st.button("🗑 刪除選取區塊"):
                delete_custom_section(config.DB_PATH, report_date_db, del_target["section_id"])
                st.success("✅ 已刪除區塊，請重新整理或切換頁面後查看。")

    st.divider()
    st.markdown("### 模板功能")

    template_name = st.text_input(
        "另存模板名稱",
        value=f"{title.strip() or '新模板'}｜{section_type}",
        key="save_template_name"
    )

    if st.button("📌 另存為模板"):
        if not title.strip():
            st.error("請先輸入區塊標題後再存成模板")
        else:
            templates = load_section_templates()
            template_obj = {
                "template_id": f"tpl_{uuid.uuid4().hex[:8]}",
                "template_name": template_name.strip(),
                "section_type": section_type,
                "layout_mode": layout_mode,
                "insert_after": insert_after,
                "enabled": enabled,
                "page_break_before": page_break_before,
                "default_title": title.strip(),
                "content": content,
            }
            templates.append(template_obj)
            save_section_templates(templates)
            st.success("✅ 已另存為模板")

# ════════════════════════════════════════════════════════════
#  ✉️ 呈報信件
# ════════════════════════════════════════════════════════════
elif mode == "✉️ 呈報信件":
    st.subheader("✉️ 呈報信件")

    if not date_options:
        st.warning("資料庫尚無資料，請先執行「🔄 資料轉檔」。")
    else:
        sel_date = st.selectbox("選擇報告日期", date_options)

        # 從 config 讀預設值
        subject_tmpl = getattr(config, "EMAIL_SUBJECT", "【風險管理日報】{date}")
        to_list      = getattr(config, "EMAIL_TO", [])
        cc_list      = getattr(config, "EMAIL_CC", [])

        report_date_fmt = sel_date.replace("-", "/")
        subject_default = subject_tmpl.replace("{date}", report_date_fmt)

        st.markdown("**信件設定確認（可於本頁臨時調整）**")
        col1, col2 = st.columns(2)
        with col1:
            subject_input = st.text_input("主旨", value=subject_default)
            to_input      = st.text_area("收件人（每行一個）", value="\n".join(to_list), height=100)
        with col2:
            cc_input      = st.text_area("副本（每行一個）", value="\n".join(cc_list), height=100)
            extra_note    = st.text_area("當天補充說明（選填）", placeholder="例：本日因資料系統延遲，資料截至 12:00", height=100)

        st.divider()

        # 確認 HTML 報告是否已產出
        date_str_file = sel_date.replace("-", "")
        html_path = config.OUTPUT_DIR / f"風險管理日報_{date_str_file}.html"
        if html_path.exists():
            st.success(f"✅ 已找到報告檔案：`{html_path.name}`")
        else:
            st.warning(f"⚠️ 找不到報告檔案 `{html_path.name}`，請先至「📄 產出報告」產出後再寄信。")

        if st.button("📨 建立 Outlook 草稿", type="primary", disabled=not html_path.exists()):
            try:
                import win32com.client as win32
                outlook = win32.Dispatch("outlook.application")
                mail    = outlook.CreateItem(0)   # 0 = olMailItem

                mail.Subject = subject_input
                for addr in [a.strip() for a in to_input.splitlines() if a.strip()]:
                    mail.Recipients.Add(addr).Type = 1   # olTo
                for addr in [a.strip() for a in cc_input.splitlines() if a.strip()]:
                    r = mail.Recipients.Add(addr)
                    r.Type = 2   # olCC

                # HTML 內文
                with open(html_path, "r", encoding="utf-8") as f:
                    html_body = f.read()
                # 簡化內文：只取 body 部分，避免整份報告直接在信內顯示
                note_html = f"<p style='color:#555;font-size:13px;'>{extra_note}</p>" if extra_note else ""
                mail.HTMLBody = f"""
                <p>您好，</p>
                <p>敬請查閱附件 {report_date_fmt} 風險管理日報。</p>
                {note_html}
                <p>風險管理部</p>
                """

                # 附上 HTML 報告
                mail.Attachments.Add(str(html_path.resolve()))

                mail.Recipients.ResolveAll()
                mail.Save()   # 存成草稿，不自動寄出
                st.success("✅ Outlook 草稿已建立，請至草稿匣確認後手動寄出。")

            except ImportError:
                st.error("❌ 找不到 win32com，請先安裝：`pip install pywin32`")
            except Exception as e:
                st.error(f"❌ 建立草稿失敗：{e}")


# ════════════════════════════════════════════════════════════
#  📁 資料來源路徑
# ════════════════════════════════════════════════════════════
elif mode == "📁 資料來源路徑":
    st.subheader("📁 資料來源路徑設定")
    st.info("修改後請點「💾 儲存」，系統會寫入 config.py 並立即生效（下次執行轉檔時套用）。")

    broker_dir  = st.text_input("經紀業務資料夾（BROKER_DIR）— 主檔 xlsb",  value=str(config.BROKER_DIR))
    broker2_dir = st.text_input("追繳處分資料夾（BROKER2_DIR）— 彙總表 xls", value=str(getattr(config, "BROKER2_DIR", config.BROKER_DIR)))
    market_dir  = st.text_input("市場風險資料夾（MARKET_DIR）", value=str(config.MARKET_DIR))
    wm_dir      = st.text_input("財管商品資料夾（WM_DIR）",     value=str(config.WM_DIR))

    col1, col2, col3 = st.columns(3)
    for label, path_str in [("經紀主檔", broker_dir), ("追繳處分", broker2_dir), ("市場", market_dir), ("財管", wm_dir)]:
        p = Path(path_str)
        if p.exists():
            st.markdown(f"✅ `{label}` 路徑存在")
        else:
            st.markdown(f"❌ `{label}` 路徑不存在：`{path_str}`")

    if st.button("💾 儲存資料來源路徑"):
        try:
            cfg_path = Path("config.py")
            txt = cfg_path.read_text(encoding="utf-8")
            import re
            txt = re.sub(r'BROKER_DIR\s*=\s*Path\([^\)]+\)',  f'BROKER_DIR   = Path(r"{broker_dir}")',  txt)
            txt = re.sub(r'BROKER2_DIR\s*=\s*Path\([^\)]+\)', f'BROKER2_DIR  = Path(r"{broker2_dir}")', txt)
            txt = re.sub(r'MARKET_DIR\s*=\s*Path\([^\)]+\)',  f'MARKET_DIR   = Path(r"{market_dir}")',  txt)
            txt = re.sub(r'WM_DIR\s*=\s*Path\([^\)]+\)',      f'WM_DIR       = Path(r"{wm_dir}")',      txt)
            cfg_path.write_text(txt, encoding="utf-8")
            st.success("✅ 已儲存，請重新整理頁面套用新設定。")
        except Exception as e:
            st.error(f"❌ 儲存失敗：{e}")


# ════════════════════════════════════════════════════════════
#  📁 產出報告路徑
# ════════════════════════════════════════════════════════════
elif mode == "📁 產出報告路徑":
    st.subheader("📁 產出報告路徑設定")
    st.info("修改後請點「💾 儲存」，寫入 config.py 後立即生效。")

    output_dir = st.text_input("HTML 輸出資料夾（OUTPUT_DIR）", value=str(config.OUTPUT_DIR))
    db_path    = st.text_input("資料庫路徑（DB_PATH）",         value=str(config.DB_PATH))

    p_out = Path(output_dir)
    p_db  = Path(db_path).parent
    st.markdown(f"{'✅' if p_out.exists() else '⚠️ 資料夾不存在（執行時會自動建立）'} OUTPUT_DIR：`{output_dir}`")
    st.markdown(f"{'✅' if p_db.exists()  else '⚠️ 資料夾不存在'} DB_PATH 所在目錄：`{p_db}`")

    if st.button("💾 儲存產出路徑"):
        try:
            cfg_path = Path("config.py")
            txt = cfg_path.read_text(encoding="utf-8")
            import re
            txt = re.sub(r'OUTPUT_DIR\s*=\s*\S+.*', f'OUTPUT_DIR  = Path(r"{output_dir}")', txt)
            txt = re.sub(r'DB_PATH\s*=\s*\S+.*',    f'DB_PATH     = Path(r"{db_path}")',    txt)
            cfg_path.write_text(txt, encoding="utf-8")
            st.success("✅ 已儲存，請重新整理頁面套用新設定。")
        except Exception as e:
            st.error(f"❌ 儲存失敗：{e}")


# ════════════════════════════════════════════════════════════
#  📧 信件設定
# ════════════════════════════════════════════════════════════
elif mode == "📧 信件設定":
    st.subheader("📧 信件設定")
    st.info("此處設定會永久儲存至 config.py，呈報信件頁面每次預設帶入這裡的值。")

    subject_tmpl = st.text_input(
        "主旨格式（{date} 會自動替換為報告日期）",
        value=getattr(config, "EMAIL_SUBJECT", "【風險管理日報】{date}")
    )
    to_default = "\n".join(getattr(config, "EMAIL_TO", []))
    cc_default = "\n".join(getattr(config, "EMAIL_CC", []))

    col1, col2 = st.columns(2)
    with col1:
        to_input = st.text_area("收件人（每行一個 email）", value=to_default, height=150)
    with col2:
        cc_input = st.text_area("副本（每行一個 email）",   value=cc_default, height=150)

    body_tmpl = st.text_area(
        "信件內文範本（HTML，{date} 會替換為報告日期）",
        value=getattr(config, "EMAIL_BODY_TMPL",
              "<p>您好，</p><p>敬請查閱附件 {date} 風險管理日報。</p><p>風險管理部</p>"),
        height=150,
    )

    if st.button("💾 儲存信件設定"):
        try:
            to_list  = [e.strip() for e in to_input.splitlines() if e.strip()]
            cc_list  = [e.strip() for e in cc_input.splitlines() if e.strip()]
            cfg_path = Path("config.py")
            txt = cfg_path.read_text(encoding="utf-8")
            import re
            txt = re.sub(r'EMAIL_TO\s*=\s*\[.*?\]', f'EMAIL_TO = {to_list!r}', txt, flags=re.S)
            txt = re.sub(r'EMAIL_CC\s*=\s*\[.*?\]', f'EMAIL_CC = {cc_list!r}', txt, flags=re.S)
            txt = re.sub(r'EMAIL_SUBJECT\s*=\s*".*?"', f'EMAIL_SUBJECT = {subject_tmpl!r}', txt)
            # EMAIL_BODY_TMPL 可能不存在，用 append
            if "EMAIL_BODY_TMPL" in txt:
                txt = re.sub(r'EMAIL_BODY_TMPL\s*=\s*".*?"', f'EMAIL_BODY_TMPL = {body_tmpl!r}', txt, flags=re.S)
            else:
                txt += f'\nEMAIL_BODY_TMPL = {body_tmpl!r}\n'
            cfg_path.write_text(txt, encoding="utf-8")
            st.success("✅ 已儲存信件設定。")
        except Exception as e:
            st.error(f"❌ 儲存失敗：{e}")