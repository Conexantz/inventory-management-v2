# StockSense – inventory management for cleaning supplies

StockSense keeps a cleaning services company's central store and client sites stocked without stock-outs or over-ordering. It follows the workflow blueprint in `docs/`: track stock, flag what is running low, let branch managers order, route orders to the General Manager for approval, and replenish the central store from suppliers before it drops into safety stock.

All data in this repository is sample data built for the case study (one central store, six client sites, 36 items, six suppliers). It is not real company data.

## What's here

| Path | What it is |
|---|---|
| `index.html` | **StockSense Orders, standalone** – the version to deploy on Vercel or any static host. Data is saved in the viewer's browser. |
| `app/index.html` | **StockSense Orders** – the interactive app for branch managers and the General Manager, as it runs inside claude.ai (see below). |
| `vercel.json`, `.vercelignore` | Vercel settings: static site, no build step, only the root `index.html` is deployed. |
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

`app/index.html` runs as a Claude artifact. It uses the artifact `db` capability for shared data and the `user` capability for roles, so it does not run on its own outside claude.ai.

## Deploying on Vercel

The root `index.html` is the same app with a small built-in runtime in place of the claude.ai capabilities. It needs no build step.

1. In Vercel, **Add New → Project** and import this repository.
2. Leave **Framework Preset** as *Other*, **Root Directory** as `./`, and leave the build and output settings empty (they come from `vercel.json`).
3. Deploy.

In this standalone version:

- Orders, approvals, stock balances and supplier orders are saved in the viewer's browser (localStorage). Each person and each device has their own copy, so a branch manager's order won't reach the General Manager on another computer.
- The viewer is the General Manager, with a **Branch view / GM view** toggle at the top to act as a branch manager.
- **Reset demo data** at the top of the page restores the sample stock.

To share data between branch managers and the General Manager across devices, the storage functions at the top of `index.html` (the `StockSense standalone runtime` script) need a backend such as Supabase, with a sign-in to tell branch managers from the General Manager.

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
