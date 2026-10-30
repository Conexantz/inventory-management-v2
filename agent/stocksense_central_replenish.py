"""StockSense – central store auto-replenishment (same rule as the StockSense Orders app).

Safety stock  = total daily usage x 7 days
Reorder point = daily usage x supplier lead time + safety stock
Order-up-to   = reorder point + 14 days of usage, rounded up to the supplier's MOQ
Position      = central stock + open supplier orders - branch items still waiting for approval
An item is ordered when position <= reorder point. Items are grouped into one purchase order per supplier.
POs up to S$500 with an approved supplier are auto-approved; larger ones wait for the General Manager.

Usage: python3 stocksense_central_replenish.py <export_dir> <out_writes.json> [trigger]
<export_dir> holds the app database exported with ArtifactData list out_dir:
  stock/*.json, orders/*.json, decisions/*.json, purchase/*.json (may be missing), config/suppliers.json
Writes a JSON list of ArtifactData batch entries (set purchase/<ref>) and prints a summary.
"""
import sys, os, json, glob, math, random, string, datetime as dt

SAFETY_DAYS, REVIEW_DAYS, AUTO_LIMIT = 7, 14, 500


def load(dirpath):
    out = {}
    for f in glob.glob(os.path.join(dirpath, "*.json")):
        out[os.path.basename(f)[:-5]] = json.load(open(f))
    return out


def main():
    src, dst = sys.argv[1], sys.argv[2]
    trigger = sys.argv[3] if len(sys.argv) > 3 else "daily check"
    stock = load(os.path.join(src, "stock"))
    orders = load(os.path.join(src, "orders"))
    decisions = load(os.path.join(src, "decisions"))
    purchase = load(os.path.join(src, "purchase"))
    sups = {s["id"]: s for s in json.load(open(os.path.join(src, "config", "suppliers.json")))["list"]}

    def line_dec(oid, i):
        d = decisions.get(oid)
        if not d:
            return None
        if "lines" in d:
            return d["lines"].get(str(i))
        return d if d.get("status") not in (None, "pending") else None

    def status_of(oid):
        d = decisions.get(oid)
        return d.get("status", "pending") if d else "pending"

    on_order, demand = {}, {}
    for p in purchase.values():
        if p.get("status") not in ("received", "cancelled"):
            for l in p["lines"]:
                on_order[l["sku"]] = on_order.get(l["sku"], 0) + float(l["qty"])
    for oid, o in orders.items():
        if status_of(oid) == "pending":
            for i, l in enumerate(o["lines"]):
                if not line_dec(oid, i):
                    demand[l["sku"]] = demand.get(l["sku"], 0) + float(l["qty"])

    by_sup = {}
    for sku, it in sorted(stock.items()):
        sp = sups.get(it.get("sup"))
        lead = float(sp["lead"]) if sp else 7
        d = float(it.get("daily") or 0)
        ss = math.ceil(d * SAFETY_DAYS)
        rop = math.ceil(d * lead + ss)
        mx = math.ceil(rop + d * REVIEW_DAYS)
        pos = float(it.get("central") or 0) + on_order.get(sku, 0) - demand.get(sku, 0)
        moq = max(1, int(it.get("moq") or 1))
        if d <= 0 or pos > rop:
            continue
        need = max(moq, math.ceil((mx - pos) / moq) * moq)
        by_sup.setdefault(it.get("sup") or "UNKNOWN", []).append({
            "sku": sku, "desc": it["desc"], "unit": it["unit"], "cost": float(it["cost"]), "qty": need,
            "value": round(need * float(it["cost"]), 2),
            "why": f"Position {pos:,.0f} at or below reorder point {rop:,} (safety {ss:,})"})

    now = dt.datetime.now(dt.timezone.utc)
    stamp = now.strftime("%Y-%m-%dT%H:%M:%S.") + f"{now.microsecond // 1000:03d}Z"
    writes, summary = [], []
    for sid, lines in by_sup.items():
        sp = sups.get(sid)
        total = round(sum(l["value"] for l in lines), 2)
        auto = bool(sp and sp.get("approved") and total <= AUTO_LIMIT)
        ref = "PO-" + now.strftime("%y%m%d") + "-" + sid.replace("SUP-", "S") + "-" + "".join(random.choices(string.ascii_uppercase + string.digits, k=3))
        eta = (now + dt.timedelta(days=float(sp["lead"]) if sp else 7)).strftime("%Y-%m-%d")
        writes.append({"op": "set", "collection": "purchase", "doc_id": ref, "data": {
            "ref": ref, "supplierId": sid, "supplier": sp["name"] if sp else sid, "lines": lines, "total": total, "eta": eta,
            "trigger": trigger, "status": "approved" if auto else "awaiting-approval",
            "approval": "auto (≤ S$500)" if auto else "", "createdAt": stamp, "createdBy": "auto"}})
        summary.append(f"{ref} {sp['name'] if sp else sid}: {len(lines)} items, S${total:,.2f}, {'auto-approved' if auto else 'needs GM approval'}")
    json.dump(writes, open(dst, "w"), indent=1)
    print(f"{len(writes)} supplier orders to raise")
    for s in summary:
        print(" -", s)


if __name__ == "__main__":
    main()
