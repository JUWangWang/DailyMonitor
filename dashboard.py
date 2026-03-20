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
            SELECT report_date, total_maint, abc_pct,
                   grade_a_pct, grade_b_pct, grade_c_pct, grade_d_pct, grade_e_pct
            FROM broker_margin
            WHERE report_date BETWEEN ? AND ?
            ORDER BY report_date
        """, (start, end)).fetchall()
    return pd.DataFrame(rows, columns=["date","total_maint","abc_pct","A","B","C","D","E"])

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


# ── 格式工具 ──────────────────────────────────────────────────
def fmt_wan(v, unit="萬"):
    if not v:
        return "0"
    wan = float(v) / 10000
    if abs(wan) >= 10000:
        s = f"{wan/10000:.2f}億"
    else:
        s = f"{wan:,.0f}{unit}"
    return ("+" if wan > 0 else "") + s

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
            ["📄 產出報告", "🧩 報告區塊編輯器", "✉️ 呈報信件"],
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
      <div>自營業務：{level_tag(_sig_market)}</div>
      <div>財管商品：{level_tag(_sig_wm)}</div>
      <div>經紀業務：{level_tag(_sig_broker)}</div>
    </div>
    """, unsafe_allow_html=True)

    # 今日重點（與 render.py 同源，從 m_pct/y_pct 即時產生）
    ai_lines = []
    for r in _all_pnl:
        mp = float(r.get("m_pct") or 0)
        yp = float(r.get("y_pct") or 0)
        dept = r.get("dept","")
        if mp >= 1.0:  ai_lines.append(("red",    f"🔴 自營 {dept} 月損失超限（{mp*100:.1f}%）"))
        elif mp >= 0.8: ai_lines.append(("orange", f"🟠 自營 {dept} 月損失80%提醒（{mp*100:.1f}%）"))
        if yp >= 1.0:  ai_lines.append(("red",    f"🔴 自營 {dept} 年損失超限（{yp*100:.1f}%）"))
        elif yp >= 0.8: ai_lines.append(("orange", f"🟠 自營 {dept} 年損失80%提醒（{yp*100:.1f}%）"))
    for r in m.get("d3_over",[]):
        ai_lines.append(("red",    f"🔴 單檔損失超限 {r['code']} {r['name']}（{r['loss_rate']*100:.1f}%）"))
    for r in m.get("d3_warn",[]):
        ai_lines.append(("orange", f"🟠 單檔損失80%提醒 {r['code']} {r['name']}（{r['loss_rate']*100:.1f}%）"))
    for v in wm.get("conc",{}).values():
        if v.get("status") in ("達L1","達L2"):
            ai_lines.append(("orange", f"🟠 財管 {v.get('name','')} {v.get('status','')}（{(v.get('pct') or 0)*100:.2f}%）"))

    if ai_lines:
        with st.expander("⚡ 今日重點說明", expanded=True):
            for _, text in ai_lines:
                st.write(text)
    else:
        with st.expander("⚡ 今日重點說明", expanded=False):
            st.write("✅ 今日各項指標正常")

    tab1, tab2, tab3, tab4 = st.tabs(["01 自營損益", "02 單檔損失", "03 財管集中度", "04~05 經紀業務"])

    # ── Tab1: 自營損益 ────────────────────────────────────────
    with tab1:
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

    # ── Tab2: 單檔損失 ────────────────────────────────────────
    with tab2:
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

    # ── Tab3: 財管集中度 ──────────────────────────────────────
    with tab3:
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

    # ── Tab4: 經紀業務 ────────────────────────────────────────
    with tab4:
        c1, c2, c3 = st.columns(3)
        c1.metric("整體維持率", f"{b['total_maint']:.1f}%")
        c2.metric("ABC合計", fmt_pct(b["abc_pct"]))
        c3.metric("融資餘額", fmt_wan(b["total_balance"]))

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
                    "代號": r["code"], "名稱": r["name"], "評等": r["grade"],
                    "融資(億)": f"{r['balance']/1e8:.2f}",
                    "集中度": fmt_pct(r["conc"]),
                } for r in b["margin_top5"]])
                st.dataframe(m5_df, hide_index=True, use_container_width=True)

        with col_r:
            st.markdown("**融券前5大個股**")
            if b["short_top5"]:
                s5_df = pd.DataFrame([{
                    "代號": r["code"], "名稱": r["name"],
                    "擔保金(億)": f"{r['collat']/1e8:.2f}",
                    "占比": fmt_pct(r["pct"]),
                    "維持率": f"{r['maint']:.1f}%",
                } for r in b["short_top5"]])
                st.dataframe(s5_df, hide_index=True, use_container_width=True)

        st.markdown("**不限用途借貸前5大客戶**")
        if b["unlim_top5"]:
            u5_df = pd.DataFrame([{
                "客戶": r["name"],
                "借款(萬)": f"{r['amount']/10000:,.0f}",
                "維持率": f"{r['maint']:.1f}%",
            } for r in b["unlim_top5"]])
            st.dataframe(u5_df, hide_index=True, use_container_width=True)


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
            f"MTD {date_a}(萬)": f"{mtd_a:+,.0f}",
            f"MTD {date_b}(萬)": f"{mtd_b:+,.0f}",
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
        pct_a = float(va.get("pct") or 0
