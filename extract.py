# ============================================================
#  extract.py  —  讀取三個 Excel，回傳結構化資料
# ============================================================

import re
from pathlib import Path
from datetime import date
import openpyxl
import pandas as pd


# ── 工具：找指定日期的檔案 ───────────────────────────────────
def find_file(directory: Path, prefix: str, target_date: date) -> Path:
    """
    在 directory 找符合 prefix_YYYYMMDD.xlsx / .xlsm 的檔案
    target_date: 要找哪一天，None 表示找最新的
    """
    date_str = target_date.strftime("%Y%m%d")
    patterns = [f"{prefix}_{date_str}.xlsx",
                f"{prefix}_{date_str}.xlsm",
                f"{prefix}-{date_str}.xlsx",
                f"{prefix}-{date_str}.xlsm"]
    for p in patterns:
        f = directory / p
        if f.exists():
            return f

    # 找不到指定日期 → 直接報錯，不使用舊資料
    raise FileNotFoundError(
        f"\n❌ 找不到 {target_date.strftime('%Y/%m/%d')} 的 [{prefix}] 檔案\n"
        f"   路徑：{directory}\n"
        f"   請確認檔案已放入資料夾，或改用指定日期：python main.py {target_date.strftime('%Y%m%d')}"
    )


# ── 工具：安全取值 ───────────────────────────────────────────
def _val(ws, row, col):
    v = ws.cell(row=row, column=col).value
    return v if v is not None else 0

def _str(ws, row, col):
    v = ws.cell(row=row, column=col).value
    return str(v).strip() if v is not None else ""

def _fmt_wan(v):
    """元 → 萬元字串，帶正負號"""
    if v is None or v == 0:
        return "0"
    wan = v / 10000
    if abs(wan) >= 10000:
        return f"{'+' if wan > 0 else ''}{wan/10000:.2f}億"
    return f"{'+' if wan > 0 else ''}{wan:,.0f}萬"

def _fmt_pct(v):
    if v is None:
        return "—"
    return f"{v*100:.1f}%"


# ════════════════════════════════════════════════════════════
#  01 + 02  市場風險（自營業務）
# ════════════════════════════════════════════════════════════
def extract_market(path: Path) -> dict:
    wb = openpyxl.load_workbook(path, data_only=True)

    # ── Sheet1：各部室損益 ────────────────────────────────
    ws1 = wb["Sheet1"]
    data_date = str(_val(ws1, 13, 7))[:10].replace("-", "/")

    # 投資銀行處（Row 16-18）
    ib_rows = []
    for r in [16, 17]:
        ib_rows.append({
            "dept":    _str(ws1, r, 3),
            "dtd":     _val(ws1, r, 5),
            "mtd":     _val(ws1, r, 6),
            "ytd":     _val(ws1, r, 7),
            "status":  _str(ws1, r, 8),
            "m_pct":   _val(ws1, r, 15),   # col O 月損益%
            "y_pct":   _val(ws1, r, 16),   # col P 年損益%
        })
    ib_total = {
        "dept": "投資銀行處 合計",
        "dtd":  _val(ws1, 18, 5),
        "mtd":  _val(ws1, 18, 6),
        "ytd":  _val(ws1, 18, 7),
        "status": "", "m_pct": None, "y_pct": None,
    }

    # 金融交易處 策略部位（Row 21-23）
    strategy_rows = []
    for r in [21, 22]:
        strategy_rows.append({
            "dept":   _str(ws1, r, 3),
            "dtd":    _val(ws1, r, 5),
            "mtd":    _val(ws1, r, 6),
            "ytd":    _val(ws1, r, 7),
            "status": _str(ws1, r, 8),
            "m_pct":  None, "y_pct": None,
        })
    strategy_total = {
        "dept": "策略部位小計",
        "dtd":  _val(ws1, 23, 5),
        "mtd":  _val(ws1, 23, 6),
        "ytd":  _val(ws1, 23, 7),
        "status": "", "m_pct": None, "y_pct": None,
    }

    # 金融交易處 交易部位（Row 24-29）
    trade_rows = []
    for r in [24, 25, 26, 27, 28, 29]:
        dept = _str(ws1, r, 3)
        if not dept:
            continue
        trade_rows.append({
            "dept":   dept,
            "dtd":    _val(ws1, r, 5),
            "mtd":    _val(ws1, r, 6),
            "ytd":    _val(ws1, r, 7),
            "status": _str(ws1, r, 8),
            "m_pct":  _val(ws1, r, 15),   # col O
            "y_pct":  _val(ws1, r, 16),   # col P
        })
    trade_total = {
        "dept": "交易部位小計",
        "dtd":  _val(ws1, 30, 5),
        "mtd":  _val(ws1, 30, 6),
        "ytd":  _val(ws1, 30, 7),
        "status": "", "m_pct": None, "y_pct": None,
    }
    ft_total = {
        "dept": "金融交易處 合計",
        "dtd":  _val(ws1, 31, 5),
        "mtd":  _val(ws1, 31, 6),
        "ytd":  _val(ws1, 31, 7),
        "status": "", "m_pct": None, "y_pct": None,
    }

    # ── 市場風險限額使用率（Row 28-42）────────────────────
    ws_mkt = wb["市場風險限額控管表"]
    limit_rows = []
    for r in range(28, 43):
        dept = _str(ws_mkt, r, 1)
        biz  = _str(ws_mkt, r, 2)
        if not dept and not biz:
            continue
        budget_pct = _val(ws_mkt, r, 3)
        m_pct      = _val(ws_mkt, r, 4)
        y_pct      = _val(ws_mkt, r, 5)
        # 判斷狀態
        def _status(v):
            if v is None or v == 0:
                return "正常"
            if v >= 1.0:
                return "超限"
            if v >= 0.8:
                return "80%提醒"
            return "正常"
        limit_rows.append({
            "dept":       dept,
            "biz":        biz,
            "budget_pct": budget_pct,
            "m_pct":      m_pct,
            "y_pct":      y_pct,
            "m_status":   _status(m_pct),
            "y_status":   _status(y_pct),
        })

    # ── D3合併：單檔損失 ──────────────────────────────────
    ws_d3 = wb["D3合併"]
    d3_rows = []
    for r in range(2, 200):
        code = _str(ws_d3, r, 4)
        if not code:
            break
        loss_rate = _val(ws_d3, r, 10)
        chk       = _str(ws_d3, r, 11)
        d3_rows.append({
            "date":      str(_val(ws_d3, r, 1))[:10],
            "type":      _str(ws_d3, r, 2),
            "code":      code,
            "name":      _str(ws_d3, r, 5),
            "market":    _str(ws_d3, r, 6),
            "amount":    _val(ws_d3, r, 8),
            "pnl":       _val(ws_d3, r, 9),
            "loss_rate": loss_rate,
            "status":    chk,
            # "status":    chk if chk else ("超限" if loss_rate <= -0.3 else
            #                               "80%提醒" if loss_rate <= -0.2 else ""),
            "note":      _str(ws_d3, r, 13),
        })

    d3_date = d3_rows[0]["date"] if d3_rows else data_date

    # 超限/警示計數（四類：月/年 × 超限/80%提醒）
    m_loss_over = [r for r in limit_rows if r["m_status"] == "超限"]    # 月損失超限
    m_loss_warn = [r for r in limit_rows if r["m_status"] == "80%提醒"] # 月損失80%提醒
    y_loss_over = [r for r in limit_rows if r["y_status"] == "超限"]    # 年損失超限
    y_loss_warn = [r for r in limit_rows if r["y_status"] == "80%提醒"] # 年損失80%提醒
    loss_over   = m_loss_over + y_loss_over   # 向下相容
    loss_warn   = m_loss_warn + y_loss_warn   # 向下相容
    d3_over     = [r for r in d3_rows if r["status"] == "超限"]
    d3_warn     = [r for r in d3_rows if r["status"] == "80%提醒"]

    return {
        "data_date":      data_date,
        "d3_date":        d3_date,
        # 投資銀行處
        "ib_rows":        ib_rows,
        "ib_total":       ib_total,
        # 金融交易處
        "strategy_rows":  strategy_rows,
        "strategy_total": strategy_total,
        "trade_rows":     trade_rows,
        "trade_total":    trade_total,
        "ft_total":       ft_total,
        # 限額使用率
        "limit_rows":     limit_rows,
        "loss_over":      loss_over,
        "loss_warn":      loss_warn,
        "m_loss_over":    m_loss_over,
        "m_loss_warn":    m_loss_warn,
        "y_loss_over":    y_loss_over,
        "y_loss_warn":    y_loss_warn,
        # 單檔損失
        "d3_rows":        d3_rows,
        "d3_over":        d3_over,
        "d3_warn":        d3_warn,
        "d3_top5":        d3_rows[:5],
    }


# ════════════════════════════════════════════════════════════
#  03  財管商品集中度
# ════════════════════════════════════════════════════════════
def extract_wm(path: Path) -> dict:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["總表"]

    data_date = ""
    # 從總表 Row 16 以後讀集中度
    conc = {
        "bond_inv":      {"pct": 0, "l1": 0.15, "l2": 0.20, "name": "", "rating": ""},
        "bond_noninv":   {"pct": 0, "l1": 0.05, "l2": 0.08, "name": "", "rating": ""},
        "fund":          {"pct": 0, "l1": 0.10, "l2": 0.15, "name": ""},
        "struct_target": {"pct": 0, "l1": 0.15, "l2": 0.20, "name": ""},
        "struct_target2":{"pct": 0, "l1": 0.15, "l2": 0.20, "name": ""},
        "struct_upper":  {"pct": 0, "l1": 0.20, "l2": 0.25, "name": "", "rating": ""},
        "struct_lower":  {"pct": 0, "l1": 0.10, "l2": 0.15, "name": ""},
    }

    # 直接用固定列讀取（依 user 確認的儲存格位置）
    # Row 19: 債券投資等級  G19=pct  H19=L1  I19=L2  J19=達警示點  L19=名稱
    # Row 20: 債券非投資等級
    # Row 23: 基金單一標的
    # Row 25: 結構型連結標的
    # Row 26: 結構型上手 BBB+以上
    # Row 27: 結構型上手 投資等級下緣
    def _row(r):
        return {
            "pct":    ws.cell(r, 7).value or 0,
            "l1":     ws.cell(r, 8).value or 0,
            "l2":     ws.cell(r, 9).value or 0,
            "alert":  str(ws.cell(r, 10).value or "").strip(),
            "name":   str(ws.cell(r, 12).value or "").strip(),
        }
    r19 = _row(19); r20 = _row(20); r23 = _row(23)
    r25 = _row(25); r26 = _row(26); r27 = _row(27)

    conc["bond_inv"].update({"pct": r19["pct"], "l1": r19["l1"] or 0.15, "l2": r19["l2"] or 0.20, "name": r19["name"]})
    conc["bond_noninv"].update({"pct": r20["pct"], "l1": r20["l1"] or 0.05, "l2": r20["l2"] or 0.08, "name": r20["name"]})
    conc["fund"].update({"pct": r23["pct"], "l1": r23["l1"] or 0.10, "l2": r23["l2"] or 0.15, "name": r23["name"]})
    conc["struct_target"].update({"pct": r25["pct"], "l1": r25["l1"] or 0.15, "l2": r25["l2"] or 0.20, "name": r25["name"]})
    conc["struct_upper"].update({"pct": r26["pct"], "l1": r26["l1"] or 0.20, "l2": r26["l2"] or 0.25, "name": r26["name"]})
    conc["struct_lower"].update({"pct": r27["pct"], "l1": r27["l1"] or 0.10, "l2": r27["l2"] or 0.15, "name": r27["name"]})

    # 整體配置比例
    alloc = {"bond": 0, "fund": 0, "struct": 0}
    for r in range(7, 12):
        cat = str(ws.cell(r, 13).value or "").strip()
        pct = ws.cell(r, 14).value or 0
        if "海外債" in cat:
            alloc["bond"] = pct
        elif "基金" in cat:
            alloc["fund"] = pct
        elif "結構型" in cat:
            alloc["struct"] = pct

    # 高資產客戶（固定儲存格：F5=人數, G5=投資總額(元), F8=BB-人數, G8=BB-金額, F11=境外人數, G11=境外金額）
    ws_ha = wb["高資產客戶"]
    ha = {
        "count":        ws_ha.cell(5, 6).value or 0,          # F5
        "total":        ws_ha.cell(5, 7).value or 0,          # G5 (元)
        "bb_count":     ws_ha.cell(8, 6).value or 0,          # F8
        "bb_amount":    ws_ha.cell(8, 7).value or 0,          # G8 (元)
        "offshore_count": ws_ha.cell(11, 6).value or 0,       # F11
        "offshore_amount": ws_ha.cell(11, 7).value or 0,      # G11 (元)
    }

    # 判斷集中度警示狀態
    def _conc_status(item):
        pct = item["pct"] or 0
        l1  = item["l1"] or 0
        l2  = item["l2"] or 0
        if l2 and pct >= l2:
            return "達L2"
        if l1 and pct >= l1:
            return "達L1"
        if l1 and pct >= l1 * 0.8:
            return "接近L1"
        return "正常"

    for k in conc:
        conc[k]["status"] = _conc_status(conc[k])

    return {
        "data_date": data_date,
        "alloc":     alloc,
        "conc":      conc,
        "ha":        ha,
    }


# ════════════════════════════════════════════════════════════
#  04 + 05  經紀業務
# ════════════════════════════════════════════════════════════
def extract_broker(path: Path) -> dict:
    wb = openpyxl.load_workbook(path, data_only=True)

    # ── 工具：模糊比對 Sheet 名稱 ─────────────────────────
    def _sheet(keyword):
        for name in wb.sheetnames:
            if keyword in name:
                return wb[name]
        return None

    # ── 融資分佈（A~E 等級）──────────────────────────────
    ws_dist = _sheet("融資分佈圖")
    dist_rows = []
    grade_map = {"A": None, "B": None, "C": None, "D": None, "E": None}
    total_balance = 0
    total_maint   = 0
    abc_pct = 0
    if ws_dist:
        for r in range(1, 20):
            raw = str(ws_dist.cell(r, 1).value or "").strip()
            grade = raw.replace("級", "")
            if grade in grade_map:
                pct     = float(ws_dist.cell(r, 2).value or 0)
                balance = float(ws_dist.cell(r, 3).value or 0)
                maint   = float(ws_dist.cell(r, 4).value or 0) * 100   # decimal → %
                grade_map[grade] = {"grade": grade, "pct": pct/100, "balance": balance*1000, "maint": maint}
            if "總計" in raw:
                total_balance = float(ws_dist.cell(r, 3).value or 0) * 1000
                total_maint   = float(ws_dist.cell(r, 4).value or 0) * 100
            # abc_pct: 在第5欄找 ABC合計 標籤，值在第6欄
            c5 = str(ws_dist.cell(r, 5).value or "").strip()
            if "ABC" in c5:
                abc_pct = float(ws_dist.cell(r, 6).value or 0)
    dist_rows = [v for v in grade_map.values() if v]

    # ── 款項借貸（解析文字格式）──────────────────────────
    ws_loan = _sheet("款項借貸")
    loans = {"half_year": 0, "t5": 0, "t30": 0}
    if ws_loan:
        # 找「只查有餘額:Y」區段後的第一個「餘額合計」= 實際未還金額
        in_active = False
        found_half = False
        for r in range(1, 200):
            txt = str(ws_loan.cell(r, 2).value or "").strip()
            if "只查有餘額:Y" in txt:
                in_active = True
            if in_active and not found_half and "餘額合計" in txt:
                nums = [int(n.replace(",","")) for n in re.findall(r"[\d,]+", txt)
                        if len(n.replace(",","")) >= 3]
                if nums:
                    loans["half_year"] = nums[0]   # 單位：元
                found_half = True
            # T+5：找類型含 "T5" 且未還借款 > 0
            if re.search(r"T5", txt) and "T5轉半" not in txt:
                nums = re.findall(r"[\d,]{5,}", txt)
                vals = [int(n.replace(",","")) for n in nums]
                if len(vals) >= 2 and vals[1] > 0:
                    loans["t5"] += vals[1]   # 未還借款金額（元）

    # ── 融資前5大個股 ─────────────────────────────────────
    ws_top = _sheet("18-27")
    margin_top5 = []
    if ws_top:
        header_found = False
        for r in range(1, 100):
            c1 = str(ws_top.cell(r, 1).value or "").strip()
            c2 = str(ws_top.cell(r, 2).value or "").strip()
            if "個股代號" in c1 or "代號" in c1:
                header_found = True
                continue
            if not header_found:
                continue
            code = c1
            name = c2
            if not code or not name:
                continue
            # 跳過純標題行
            if not re.match(r"^\d{4,}", code):
                continue
            grade = str(ws_top.cell(r, 3).value or "").strip()
            balance_amt = float(ws_top.cell(r, 7).value or 0) * 1e8   # 億 → 元
            conc  = float(ws_top.cell(r, 9).value or 0)   # 集中度
            maint  = float(ws_top.cell(r, 11).value or 0)   # 維持率
            margin_top5.append({
                "code": code, "name": name, "grade": grade,
                "balance": balance_amt, "conc": conc, "maint": maint,
            })
            if len(margin_top5) >= 5:
                break

    # ── 融券前5大 ─────────────────────────────────────────
    ws_short = _sheet("融券")
    short_top5 = []
    if ws_short:
        for r in range(3, 30):
            code  = str(ws_short.cell(r, 1).value or "").strip()
            name  = str(ws_short.cell(r, 2).value or "").strip()
            if not code or not name or not re.match(r"^\d{4}", code):
                continue
            collat = float(ws_short.cell(r, 4).value or 0)
            pct    = float(ws_short.cell(r, 5).value or 0) / 100
            maint  = float(ws_short.cell(r, 6).value or 0) * 100   # decimal → %
            short_top5.append({
                "code": code, "name": name,
                "collat": collat, "pct": pct, "maint": maint,
            })
            if len(short_top5) >= 5:
                break

    # ── 不限用途借貸前5大客戶（28-2 限用途）────────────────
    ws_unlim = _sheet("28-2")
    unlim_top5 = []
    if ws_unlim:
        for r in range(2, 50):
            acct  = str(ws_unlim.cell(r, 1).value or "").strip()
            name  = str(ws_unlim.cell(r, 2).value or "").strip()
            amt   = ws_unlim.cell(r, 4).value   # 借款金額(萬)
            maint = ws_unlim.cell(r, 5).value   # 整戶維持率%
            if not name or not acct or not re.match(r"^\d{4}", acct):
                continue
            unlim_top5.append({
                "name":   name.strip(),
                "amount": float(amt or 0) * 10000,   # 萬 → 元
                "maint":  float(maint or 0),
            })
            if len(unlim_top5) >= 5:
                break

    # ── 有價證券借貸（29-11，依類別編號區分）──────────────
    # 類別：0=全部 1=960T自營 2=外資 3=他券商
    # 金額單位：仟元
    ws_sec = _sheet("29-11")
    sec_lending = {"foreign": 0, "broker": 0, "prop": 0, "total": 0, "rate": 0}
    if ws_sec:
        cat_map = {}   # {類別編號: 借券金額(元)}
        rate_map = {}  # {類別編號: 加權費率}
        current_cat = None
        for r in range(1, 200):
            txt = str(ws_sec.cell(r, 1).value or "").strip()
            # 類別標頭：類別:N(...)
            m_cat = re.search(r"類別:(\d+)\(", txt)
            if m_cat:
                current_cat = int(m_cat.group(1))
            # 加權平均費率行，同時有借券金額
            if "加權平均費率" in txt and current_cat is not None:
                m_amt = re.search(r"借券金額:\s*([\d,]+)", txt)
                m_rate = re.search(r"加權平均費率:\s*([\d\.]+)", txt)
                if m_amt:
                    cat_map[current_cat] = int(m_amt.group(1).replace(",",""))   # 已是元
                if m_rate:
                    rate_map[current_cat] = float(m_rate.group(1))
        sec_lending["total"]   = cat_map.get(0, 0)
        sec_lending["prop"]    = cat_map.get(1, 0)   # 960T 自營
        sec_lending["foreign"] = cat_map.get(2, 0)   # 外資
        sec_lending["broker"]  = cat_map.get(3, 0)   # 他券商
        sec_lending["rate"]    = rate_map.get(0, 0)  # 整體加權費率

    return {
        "dist_rows":     dist_rows,
        "total_balance": total_balance,
        "total_maint":   total_maint,
        "abc_pct":       abc_pct,
        "loans":         loans,
        "margin_top5":   margin_top5,
        "short_top5":    short_top5,
        "unlim_top5":    unlim_top5,
        "sec_lending":   sec_lending,
    }


# ════════════════════════════════════════════════════════════
#  彙整：一次讀三個 Excel
# ════════════════════════════════════════════════════════════
def extract_all(target_date: date, config) -> dict:
    from config import BROKER_DIR, MARKET_DIR, WM_DIR
    from config import BROKER_PREFIX, MARKET_PREFIX, WM_PREFIX

    broker_file = find_file(BROKER_DIR, BROKER_PREFIX, target_date)
    market_file = find_file(MARKET_DIR, MARKET_PREFIX, target_date)
    wm_file     = find_file(WM_DIR,     WM_PREFIX,     target_date)

    print(f"  [經紀] {broker_file.name}")
    print(f"  [市場] {market_file.name}")
    print(f"  [財管] {wm_file.name}")

    market = extract_market(market_file)
    wm     = extract_wm(wm_file)
    broker = extract_broker(broker_file)

    return {"market": market, "wm": wm, "broker": broker,
            "report_date": target_date.strftime("%Y/%m/%d")}
