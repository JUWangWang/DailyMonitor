# ============================================================
#  alert_logic.py  —  今日重點 / 燈號共用邏輯
# ============================================================

def build_auto_alert_items(data: dict) -> list[dict]:
    m = data["market"]
    wm = data["wm"]

    items = []
    all_pnl_rows = m.get("ib_rows", []) + m.get("strategy_rows", []) + m.get("trade_rows", [])
    seq = 1

    for r in all_pnl_rows:
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
                "sort_order": 10 + seq,
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
                "sort_order": 10 + seq,
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
                "sort_order": 20 + seq,
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
                "sort_order": 20 + seq,
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
            "sort_order": 200,
        })

    for r in m.get("d3_warn", []):
        items.append({
            "id": f'auto_d3_warn_{r["code"]}',
            "source": "auto",
            "category": "market",
            "text": f'單檔損失80%提醒 {r["code"]} {r["name"]}（{r["loss_rate"]*100:.1f}%）',
            "level": "o",
            "enabled": True,
            "sort_order": 210,
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
                "sort_order": 300,
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


def calc_signal_levels(data: dict) -> dict:
    m = data["market"]
    wm = data["wm"]
    b = data["broker"]

    all_pnl_rows = m.get("ib_rows", []) + m.get("strategy_rows", []) + m.get("trade_rows", [])

    loss_over_cnt = sum(1 for r in all_pnl_rows if float(r.get("m_pct") or 0) >= 1.0)
    loss_warn_cnt = sum(1 for r in all_pnl_rows if 0.8 <= float(r.get("m_pct") or 0) < 1.0)
    y_over_cnt = sum(1 for r in all_pnl_rows if float(r.get("y_pct") or 0) >= 1.0)
    y_warn_cnt = sum(1 for r in all_pnl_rows if 0.8 <= float(r.get("y_pct") or 0) < 1.0)

    sig_market = (
        "red" if loss_over_cnt or y_over_cnt or m.get("d3_over")
        else "orange" if loss_warn_cnt or y_warn_cnt or m.get("d3_warn")
        else "green"
    )

    sig_wm = (
        "orange"
        if any(v.get("status") in ("達L1", "達L2") for v in wm.get("conc", {}).values())
        else "green"
    )

    broker_maint = float(b.get("total_maint", 0) or 0)
    unlim_maint = float(b.get("unlim_total_maint", 0) or 0)

    sig_broker = (
        "orange"
        if (
            (broker_maint > 0 and broker_maint < 160) or
            (unlim_maint > 0 and unlim_maint < 160)
        )
        else "green"
    )

    return {
        "market": sig_market,
        "wm": sig_wm,
        "broker": sig_broker,
        "loss_over_cnt": loss_over_cnt,
        "loss_warn_cnt": loss_warn_cnt,
        "y_over_cnt": y_over_cnt,
        "y_warn_cnt": y_warn_cnt,
    }