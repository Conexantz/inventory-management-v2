# StockSense – inventory management for cleaning supplies

StockSense keeps a cleaning services company's central store and client sites stocked without stock-outs or over-ordering. It follows the workflow blueprint in `docs/`: track stock, flag what is running low, let branch managers order, route orders to the General Manager for approval, and replenish the central store from suppliers before it drops into safety stock.

All data in this repository is sample data built for the case study (one central store, six client sites, 36 items, six suppliers). It is not real company data.

## What's here

| Path | What it is |
|---|---|
| `app/index.html` | **StockSense Orders** – the interactive app for branch managers and the General Manager (see below). |
| `agent/stocksense_agent.py` | Daily replenishment agent: reads the master inventory workbook, reconciles usage, forecasts 14-day demand, calculates reorder points, proposes transfers, drafts POs and the WhatsApp message. |
| `agent/stocksense_central_replenish.py` | Central-store auto-replenishment check, the same rule the app uses. Used by the daily scheduled run. |
| `dashboard/index.html` | Page for the StockSense replenishment dashboard (data comes from the run outputs). |
| `data/Master_Inventory_List.xlsx` | Sample master inventory: items, suppliers, sites, stock movements and settings, with live formulas. |
| `data/StockSense_Run_2026-10-30.xlsx` | Example output of one agent run. |
| `docs/` | The workflow blueprint and an example WhatsApp order message. |

## StockSense Orders app

Tabs:

- **Low stock** – flags items where a branch holds less than 7 days of its usage and recommends a top-up to 14 days. Branch managers can order more than recommended, with a reason.
- **Place order** – catalogue with unit cost, central store stock and the branch's own balance, plus an order slip.
- **Approvals** (General Manager) – each line has an automatic check against the recommended restocking level and is flagged when it doesn't fit. The GM can tick individual items, a whole order, or all items, and approve or reject them together. Approval issues stock from the central store to the branch.
- **Order history** – status per order and item, and monthly spend per branch.
- **Supplier orders** (General Manager) – central store cover for every item and supplier purchase orders raised automatically.
- **Inventory & costing** – balances at the central store and each branch, valued at unit cost.

The app runs as a Claude artifact. It uses the artifact `db` capability for shared data and the `user` capability for roles, so `index.html` does not run on its own outside claude.ai.

### Central store replenishment rule

```
safety stock   = total daily usage x 7 days
reorder point  = daily usage x supplier lead time + safety stock
order up to    = reorder point + 14 days of usage, rounded up to the supplier's MOQ
position       = central stock + open supplier orders - branch items waiting for approval
```

When position is at or below the reorder point, a purchase order is raised, grouped by supplier. Orders up to S$500 with an approved supplier are approved automatically; larger orders wait for the General Manager. A delivery is added to the central store when it is marked received.

The check runs after the GM approves branch items, when the GM presses **Check now**, and every morning through the daily scheduled run.

## Running the agent locally

```bash
pip install openpyxl
python3 agent/stocksense_agent.py data/Master_Inventory_List.xlsx 2026-10-30 out
```

Outputs in `out/`: the run workbook, `whatsapp_message.txt`, `whatsapp_link.txt`, dashboard data files and `run_report.json`.
