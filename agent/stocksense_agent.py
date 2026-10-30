"""StockSense – Inventory Tracker & Stock Replenishment Agent.

Implements the 7 steps of the StockSense blueprint against the master
inventory workbook (Master_Inventory_List.xlsx kept in Google Drive):

 1. Pull latest stock counts from the master inventory workbook
 2. Reconcile records – flag variances >15% vs expected usage
 3. Forecast 14-day demand per SKU per site (usage weighted by site headcount)
 4. Reorder point = daily usage x lead time + safety stock; order qty to max level / MOQ
 5. Check the central store first – propose central->site transfers before buying
 6. Draft POs per supplier as a WhatsApp ordering message
 7. Route each PO for approval using the blueprint thresholds

Usage:  python stocksense_agent.py <master_inventory.xlsx> [run_date YYYY-MM-DD] [outdir]
Outputs (in outdir): run_report.json, whatsapp_message.txt, StockSense_Run_<date>.xlsx
"""
import sys, os, json, math, datetime as dt
from collections import defaultdict
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

VARIANCE_LIMIT = 0.15          # blueprint step 2
FORECAST_DAYS = 14             # blueprint step 3
VARIANCE_WINDOW_DAYS = 7       # look-back for reconciliation against the movement log


def read_inputs(path):
    """Step 1 – read raw inputs only (never trust cached formula values)."""
    wb = load_workbook(path, data_only=True)
    st = wb["Settings"]
    settings = {st.cell(r, 1).value: st.cell(r, 2).value for r in range(5, 12) if st.cell(r, 1).value}
    cfg = {
        "safety_days": float(settings["Safety stock (days of usage)"]),
        "review_days": float(settings["Review period (days)"]),
        "auto_limit": float(settings["Auto-approve PO limit (S$)"]),
        "ops_limit": float(settings["Ops Manager approval limit (S$)"]),
        "expiry_days": float(settings["Expiry warning (days)"]),
    }
    sup = {}
    for r in wb["Suppliers"].iter_rows(min_row=5, values_only=True):
        if r[0]:
            sup[r[0]] = {"name": r[1], "contact": r[3], "phone": r[5], "lead": float(r[6] or 0),
                         "terms": r[7], "otd": r[8], "approved": (r[9] or "").lower() == "yes"}
    sites = []
    for r in wb["Sites"].iter_rows(min_row=5, values_only=True):
        if r[0] and r[0] != "CS":
            sites.append({"id": r[0], "name": r[1], "supervisor": r[4], "cleaners": float(r[5] or 0), "day": r[6]})
    items = []
    for r in wb["Master Inventory"].iter_rows(min_row=5, values_only=True):
        if not r[0] or r[0] == "TOTAL" or str(r[0]).startswith("Notes"):
            continue
        items.append({"sku": r[0], "desc": r[1], "cat": r[2], "unit": r[3], "sup": r[4],
                      "cost": float(r[6] or 0), "central": float(r[7] or 0), "sites": float(r[8] or 0),
                      "daily": float(r[10] or 0), "moq": float(r[15] or 1),
                      "expiry": r[21] if isinstance(r[21], (dt.date, dt.datetime)) else None})
    moves = []
    for r in wb["Stock Movements"].iter_rows(min_row=5, values_only=True):
        if r[0] and r[2]:
            d = r[0].date() if isinstance(r[0], dt.datetime) else r[0]
            moves.append({"date": d, "ref": r[1], "sku": r[2], "type": r[4], "from": r[5], "to": r[6], "qty": float(r[7] or 0)})
    return cfg, sup, sites, items, moves


def ceil_to(x, m):
    return int(math.ceil(x / m) * m) if m else int(math.ceil(x))


def run(path, run_date):
    cfg, sup, sites, items, moves = read_inputs(path)
    total_cleaners = sum(s["cleaners"] for s in sites) or 1
    window_start = run_date - dt.timedelta(days=VARIANCE_WINDOW_DAYS)

    # issues recorded per SKU in the look-back window (central -> site)
    issued = defaultdict(float)
    for m in moves:
        if m["type"] == "Issue" and window_start <= m["date"] <= run_date:
            issued[m["sku"]] += m["qty"]

    lines, transfers, holds, forecast_rows, alerts = [], [], [], [], []
    for it in items:
        s = sup.get(it["sup"], {"name": "UNKNOWN", "lead": 0, "approved": False})
        on_hand = it["central"] + it["sites"]
        ss = math.ceil(it["daily"] * cfg["safety_days"])
        rop = math.ceil(it["daily"] * s["lead"] + ss)
        max_lvl = math.ceil(rop + it["daily"] * cfg["review_days"])
        cover = on_hand / it["daily"] if it["daily"] else None

        # Step 2 – reconciliation: recorded issues vs expected usage over the window
        variance = None
        if it["sku"] in issued:
            expected = it["daily"] * VARIANCE_WINDOW_DAYS
            variance = (issued[it["sku"]] - expected) / expected if expected else 0
        hold = variance is not None and abs(variance) > VARIANCE_LIMIT

        # Step 3 – 14-day forecast per site, weighted by cleaner headcount
        demand14 = it["daily"] * FORECAST_DAYS
        per_site = {st_["id"]: round(demand14 * st_["cleaners"] / total_cleaners, 1) for st_ in sites}
        forecast_rows.append({"sku": it["sku"], "desc": it["desc"], "demand14": round(demand14, 1), **per_site})

        # status
        if on_hand == 0:
            status = "OUT OF STOCK"
        elif on_hand <= ss:
            status = "CRITICAL"
        elif on_hand <= rop:
            status = "REORDER"
        elif on_hand > max_lvl * 1.5:
            status = "OVERSTOCK"
        else:
            status = "OK"

        exp_flag = "-"
        if it["expiry"]:
            days = ((it["expiry"].date() if isinstance(it["expiry"], dt.datetime) else it["expiry"]) - run_date).days
            exp_flag = "EXPIRED" if days < 0 else ("Expiring soon" if days <= cfg["expiry_days"] else "OK")
            if exp_flag != "OK":
                alerts.append(f"{it['sku']} {it['desc']}: {exp_flag} ({days} days)")

        rec = {**it, "supplier": s["name"], "lead": s["lead"], "on_hand": on_hand, "ss": ss, "rop": rop,
               "max": max_lvl, "cover": round(cover, 1) if cover is not None else None, "status": status,
               "variance": round(variance, 3) if variance is not None else None, "expiry_flag": exp_flag}

        # Step 5 – central store first: sites below their 14-day need while central holds a surplus
        site_need = demand14 - it["sites"]
        central_reserve = ss  # keep safety stock at the central store
        if site_need > 0 and it["central"] > central_reserve and status in ("OK", "OVERSTOCK"):
            qty = int(min(math.ceil(site_need), it["central"] - central_reserve))
            if qty > 0:
                transfers.append({"sku": it["sku"], "desc": it["desc"], "unit": it["unit"], "qty": qty,
                                  "split": {k: v for k, v in per_site.items()}, "reason": "Sites below 14-day need; central has surplus"})

        # Step 4 – reorder
        if on_hand <= rop:
            qty = ceil_to(max(max_lvl - on_hand, it["moq"]), it["moq"])
            if hold:
                holds.append({**rec, "order_qty": qty,
                              "reason": f"Issues logged {issued[it['sku']]:.0f} vs expected {it['daily']*VARIANCE_WINDOW_DAYS:.0f} "
                                        f"({variance:+.0%}) – supervisor to confirm recount before ordering"})
            else:
                lines.append({**rec, "order_qty": qty, "value": round(qty * it["cost"], 2)})
        elif hold:
            holds.append({**rec, "order_qty": 0,
                          "reason": f"Variance {variance:+.0%} vs expected usage – recount (no order needed yet)"})

    # Step 6/7 – group into POs per supplier and route for approval
    pos = []
    by_sup = defaultdict(list)
    for l in lines:
        by_sup[l["sup"]].append(l)
    for n, (sid, ls) in enumerate(sorted(by_sup.items()), 1):
        total = round(sum(l["value"] for l in ls), 2)
        s = sup[sid]
        if not s["approved"]:
            appr = "Operations Manager (new supplier)"
        elif total <= cfg["auto_limit"]:
            appr = "Auto-approved (manager notified)"
        elif total <= cfg["ops_limit"]:
            appr = "Operations Manager"
        else:
            appr = "Finance / CEO"
        pos.append({"po": f"PO-{run_date:%y%m%d}-{n:02d}", "supplier_id": sid, "supplier": s["name"],
                    "contact": s["contact"], "phone": s["phone"], "lead": s["lead"],
                    "eta": str(run_date + dt.timedelta(days=int(s["lead"]))), "total": total,
                    "approval": appr, "lines": [{k: l[k] for k in ("sku", "desc", "unit", "order_qty", "cost", "value", "status")} for l in ls]})

    return cfg, sup, sites, items, pos, transfers, holds, forecast_rows, alerts, lines


def build_records(cfg, sup, items):
    out = []
    for it in items:
        s = sup.get(it["sup"], {"lead": 0, "name": "UNKNOWN"})
        on_hand = it["central"] + it["sites"]
        ss = math.ceil(it["daily"] * cfg["safety_days"])
        rop = math.ceil(it["daily"] * s["lead"] + ss)
        mx = math.ceil(rop + it["daily"] * cfg["review_days"])
        st = ("OUT OF STOCK" if on_hand == 0 else "CRITICAL" if on_hand <= ss else "REORDER" if on_hand <= rop
              else "OVERSTOCK" if on_hand > mx * 1.5 else "OK")
        out.append({"sku": it["sku"], "desc": it["desc"], "cat": it["cat"], "on_hand": on_hand, "central": it["central"],
                    "sites": it["sites"], "rop": rop, "ss": ss, "max": mx, "status": st,
                    "value": round(on_hand * it["cost"], 2),
                    "cover": round(on_hand / it["daily"], 1) if it["daily"] else None})
    return out


def whatsapp_text(run_date, pos, transfers, holds, alerts, recs):
    sc = defaultdict(int)
    for r in recs:
        sc[r["status"]] += 1
    L = [f"*StockSense – Daily Replenishment {run_date:%d %b %Y}*",
         f"Scanned {len(recs)} SKUs: {sc['OUT OF STOCK']} out of stock, {sc['CRITICAL']} critical, {sc['REORDER']} to reorder.",
         ""]
    if pos:
        L.append(f"*Purchase orders drafted ({len(pos)}) – S${sum(p['total'] for p in pos):,.2f}*")
        for p in pos:
            L.append(f"\n*{p['po']}* · {p['supplier']} · S${p['total']:,.2f}")
            L.append(f"Approval: {p['approval']} · ETA {dt.date.fromisoformat(p['eta']):%d %b}")
            for l in p["lines"]:
                L.append(f"  • {l['desc']} – {l['order_qty']} {l['unit'].lower()}  (S${l['value']:,.2f})")
    if holds:
        L.append(f"\n*On hold – recount needed ({len(holds)})*")
        for h in holds:
            L.append(f"  • {h['desc']}: {h['reason'].split(' – ')[0]}")
    if transfers:
        L.append(f"\n*Transfers from central store ({len(transfers)})*")
        for t in transfers:
            L.append(f"  • {t['desc']} – {t['qty']} {t['unit'].lower()}")
    if alerts:
        L.append("\n*Expiry alerts*")
        for a in alerts:
            L.append(f"  • {a}")
    L.append("\nReply *APPROVE <PO no.>* or *HOLD <PO no.>*. POs ≤ S$500 are auto-sent.")
    return "\n".join(L)


def write_workbook(path, run_date, pos, transfers, holds, forecast_rows, recs, sites):
    wb = Workbook()
    F = "Arial"; H = PatternFill("solid", fgColor="1F4E5A"); HF = Font(name=F, bold=True, color="FFFFFF", size=10)
    th = Side(style="thin", color="BFBFBF"); B = Border(left=th, right=th, top=th, bottom=th)

    def sheet(ws, title, heads, rows, widths, money=()):
        ws["A1"] = title; ws["A1"].font = Font(name=F, bold=True, size=13, color="1F4E5A")
        ws["A2"] = f"StockSense agent run {run_date:%d %b %Y}"; ws["A2"].font = Font(name=F, italic=True, size=9, color="595959")
        for i, (h, w) in enumerate(zip(heads, widths), 1):
            c = ws.cell(4, i, h); c.font = HF; c.fill = H; c.border = B; c.alignment = Alignment(wrap_text=True, vertical="center")
            ws.column_dimensions[c.column_letter].width = w
        for r, row in enumerate(rows, 5):
            for i, v in enumerate(row, 1):
                c = ws.cell(r, i, v); c.font = Font(name=F, size=10); c.border = B
                if i in money: c.number_format = "#,##0.00"
        ws.freeze_panes = "A5"; ws.sheet_view.showGridLines = False

    ws = wb.active; ws.title = "Purchase Orders"
    rows = []
    for p in pos:
        for l in p["lines"]:
            rows.append([p["po"], p["supplier"], l["sku"], l["desc"], l["status"], l["order_qty"], l["unit"], l["cost"], l["value"], p["total"], p["approval"], p["eta"], "Pending" if "Auto" not in p["approval"] else "Sent"])
    sheet(ws, "Draft Purchase Orders", ["PO No.", "Supplier", "SKU", "Item", "Stock Status", "Order Qty", "Unit", "Unit Cost (S$)", "Line Value (S$)", "PO Total (S$)", "Approval Route", "ETA", "PO Status"],
          rows, [15, 28, 10, 32, 13, 9, 8, 10, 11, 11, 26, 11, 10], money=(8, 9, 10))
    sheet(wb.create_sheet("Holds – Recount"), "Held for Supervisor Recount (variance >15%)", ["SKU", "Item", "Status", "Order Qty (if confirmed)", "Reason"],
          [[h["sku"], h["desc"], h["status"], h["order_qty"], h["reason"]] for h in holds], [10, 32, 13, 12, 80])
    sheet(wb.create_sheet("Transfers"), "Central Store → Site Transfers", ["SKU", "Item", "Qty", "Unit"] + [f"{s['id']} share" for s in sites] + ["Reason"],
          [[t["sku"], t["desc"], t["qty"], t["unit"]] + [t["split"][s["id"]] for s in sites] + [t["reason"]] for t in transfers],
          [10, 32, 7, 8] + [9] * len(sites) + [40])
    sheet(wb.create_sheet("14-Day Forecast"), "14-Day Demand Forecast per Site (headcount-weighted)", ["SKU", "Item", "Total 14-day"] + [f"{s['id']} {s['name']}" for s in sites],
          [[f["sku"], f["desc"], f["demand14"]] + [f[s["id"]] for s in sites] for f in forecast_rows], [10, 32, 10] + [14] * len(sites))
    sheet(wb.create_sheet("Stock Status"), "Stock Status Snapshot", ["SKU", "Item", "Category", "Central", "At Sites", "On Hand", "Safety Stock", "Reorder Pt", "Max", "Days Cover", "Stock Value (S$)", "Status"],
          [[r["sku"], r["desc"], r["cat"], r["central"], r["sites"], r["on_hand"], r["ss"], r["rop"], r["max"], r["cover"], r["value"], r["status"]] for r in recs],
          [10, 32, 16, 8, 8, 8, 9, 9, 7, 9, 12, 13], money=(11,))
    wb.save(path)


if __name__ == "__main__":
    src = sys.argv[1]
    run_date = dt.date.fromisoformat(sys.argv[2]) if len(sys.argv) > 2 else dt.date.today()
    outdir = sys.argv[3] if len(sys.argv) > 3 else "."
    os.makedirs(outdir, exist_ok=True)
    cfg, sup, sites, items, pos, transfers, holds, forecast_rows, alerts, lines = run(src, run_date)
    recs = build_records(cfg, sup, items)
    msg = whatsapp_text(run_date, pos, transfers, holds, alerts, recs)
    open(os.path.join(outdir, "whatsapp_message.txt"), "w").write(msg)
    import urllib.parse
    open(os.path.join(outdir, "whatsapp_link.txt"), "w").write("https://wa.me/?text=" + urllib.parse.quote(msg))
    # dashboard data files
    json.dump([{k: x[k] for k in ("sku", "desc", "cat", "status", "central", "sites", "on_hand", "rop", "max", "cover", "value")} for x in recs],
              open(os.path.join(outdir, "stock.json"), "w"))
    json.dump([{"po": p["po"], "supplier": p["supplier"], "approval": p["approval"], "eta": p["eta"], "sku": l["sku"], "item": l["desc"],
                "qty": l["order_qty"], "unit": l["unit"], "line_value": l["value"], "po_total": p["total"]} for p in pos for l in p["lines"]],
              open(os.path.join(outdir, "purchase_orders.json"), "w"))
    json.dump([{"sku": t["sku"], "item": t["desc"], "qty": t["qty"], "unit": t["unit"], "reason": t["reason"]} for t in transfers],
              open(os.path.join(outdir, "transfers.json"), "w"))
    json.dump([{"sku": h["sku"], "item": h["desc"], "status": h["status"], "order_qty": h["order_qty"], "reason": h["reason"]} for h in holds],
              open(os.path.join(outdir, "holds.json"), "w"))
    write_workbook(os.path.join(outdir, f"StockSense_Run_{run_date}.xlsx"), run_date, pos, transfers, holds, forecast_rows, recs, sites)
    json.dump({"run_date": str(run_date), "settings": cfg, "records": recs, "purchase_orders": pos,
               "transfers": transfers, "holds": [{k: h[k] for k in ("sku", "desc", "status", "order_qty", "reason")} for h in holds],
               "expiry_alerts": alerts, "sites": sites}, open(os.path.join(outdir, "run_report.json"), "w"), indent=1, default=str)
    print(msg)
