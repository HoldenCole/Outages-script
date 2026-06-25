# Refinery Outage Analytics

Turns the Snowflake refinery-outage export into three commodities-desk-grade
deliverables from one data pipeline:

1. **Excel workbook** (`outage_workbook.xlsx`) — 12 visible sheets (incl. an
   **Assumptions & Methodology** source-of-truth sheet and dedicated **Mogas**,
   **Naphtha** and **Margin Context** sheets), native charts,
   a two-way sensitivity heatmap, a 2027 implied-total-by-scenario chart, and a **live**
   2027 **Scenario Analysis** model (Conservative/Average/Active fan) driven by
   data-validation dropdowns. *(Priority 1.)*
2. **Slide deck** (`outage_deck.pptx`) — 16:9 deck (two charts per content slide)
   that **opens with a month-over-month "what changed" section** — movers by
   PADD/unit, a trailing-13-month trend, and new-vs-resolved outages, all computed
   live from the data — then the 2027 outlook: the planned H1-vs-H1 like-for-like,
   the Conservative/Average/Active fan with an implied-total-by-scenario chart, a
   five-PADD small-multiples and a national turnaround table. *(Priority 2.)*
3. **HTML dashboard** (`outage_dashboard.html`) — a single self-contained file
   with a **dynamic focus-year selector** (KPIs show the picked year vs the prior
   year), a 2026-vs-2025 / H1'27-vs-H1'26 / 2027-forecast outlook strip, and a
   live 2027 scenario panel. *(Priority 3.)*

All three read from the **same** `engine.build_context()` bundle, so the numbers
always agree. Everything is re-runnable: point at a refreshed export and rebuild.

---

## Quick start (VS Code or any terminal)

```bash
# 1. create an isolated environment (recommended before sharing)
python -m venv .venv
source .venv/bin/activate            # macOS/Linux
# .venv\Scripts\Activate.ps1         # Windows PowerShell

# 2. install dependencies
pip install -r requirements.txt

# 3. build both deliverables (deck + dashboard)
python scripts/build_all.py
```

That reads `data/Refinery_Outages_Enhanced.xlsx` and writes the deck and
dashboard to `output/`. **The input data is not committed** — drop your outage
export at that path (or pass one as an argument). Paths resolve relative to the
repo root, so it runs from any directory. **No network is required** to build.

In **VS Code**, the committed `.vscode/` configs give you three one-click paths
(on first open it offers to install the recommended Python + Debugpy extensions):

* open any script and hit **▶ Run**;
* press **F5** → *Run and Debug → "Build all deliverables"* (or *"…pick a data
  file"* to be prompted for an export path);
* run the default build task — *Terminal → Run Build Task…* (`⌘⇧B` /
  `Ctrl+Shift+B`) — which also lists deck-only and dashboard-only tasks.

Build a single deliverable, or point at a refreshed export / different output dir:

```bash
python scripts/build_slides.py                         # -> output/outage_deck.pptx
python scripts/build_dashboard.py                      # -> output/outage_dashboard.html
python scripts/build_all.py path/to/export.xlsx --outdir dist/
```

### Run it on your own data

Drop your export in `data/`, pass its path, and rebuild. The engine
**auto-detects the schema** — no flag needed:

* **"Refinery Outages Enhanced" breakdown (primary)** — one row per outage event
  with columns `Source`, `Country/Region`, `Owner`, `Plant`, `Start Date`,
  `End Date`, `Unit Name`, `Outage Type`, `Offline Capacity (KBD)`. Each event is
  expanded to the calendar months it spans, and the free-text `Unit Name` is
  mapped to a unit class (CDU / FCC / hydrocracker / reformer / …). Only verified
  years (2023+) are kept and obvious placeholder spans are dropped.
* **Legacy Snowflake `Query1` export** — also recognised automatically
  (`CAP_OFFLINE_ADJUSTED_KBD`, `OUTAGE_YEAR`, `OUTAGE_MONTH`, plus the usual
  `OUTAGE_TYPE` / `PAD_DIST` / `UNIT_CATEGORY` / `REFINERY_OPERATOR` columns).
* `OUTAGE_TYPE` `PLANNED` / `UNPLANNED` (and `UNKNOWN`, folded into unplanned)
  drives the split; the 2027 unplanned scenario extrapolates off the 2023–2025
  baseline of actual unplanned offline.

---

## Repository layout

```
.
├── scripts/      pipeline code (run these)
│   ├── engine.py            data core — the only place raw data is touched
│   ├── charts.py            matplotlib renderers (shared by the deck)
│   ├── build_slides.py      PowerPoint deck (python-pptx)
│   ├── build_dashboard.py   self-contained HTML dashboard (Chart.js inlined)
│   ├── audit_exxon.py       reconcile the Exxon slate vs the corporate plan
│   └── build_all.py         orchestrator — loads once, builds deck + dashboard
├── data/         live input: Refinery_Outages_Enhanced.xlsx
├── output/       generated deliverables (.xlsx / .pptx / .html)
├── reference/    mogas & ethylene-margins workbooks, gasoline-weekly PDF, Yields.txt
├── docs/         CLAUDE_CODE_BUILD_SPEC.md, "What good output looks like"
├── legacy/       superseded openpyxl prototype (reference only)
├── README.md  ·  requirements.txt
```

---

## Refreshing with new data

Drop a new Snowflake export (same `Query1` schema) into `data/` and re-run
`python scripts/build_all.py` (or pass its path). Only `engine.py` touches raw
data; the builders consume its frames. The Excel **Scenario Analysis** sheet
(scenario cascade + sensitivity heatmap) is live formulas, so you can also just
edit the yellow input cells in Excel without rebuilding.

**Market data (gasoline crack).** The **Margin Context** sheet values outages in
dollars off a monthly gasoline crack vendored to `data/market_crack.csv`. Refresh
it with `python scripts/fetch_market_data.py` (pulls EIA NY-Harbor gasoline + WTI,
no API key), or overwrite the CSV with a Bloomberg pull (e.g. RBOB 321) using the
same columns. The crack is also a blue, editable input inside the workbook, so the
$ figures recompute live without a rebuild. If the CSV is absent the rest of the
pipeline still builds — the Margin Context sheet just shows a Bloomberg-fillable
template.

---

## Workbook sheets

**12 findable sheets** (+ a hidden `Data` backing sheet):

`Cover · Assumptions · Dashboard · Explorer · Trends · PADD · Units & Refineries ·
Mogas · Naphtha · Margin Context · Events & TAs · Scenario Analysis`

- **Cover** — contents (hyperlinks), read-before-use caveats (COVID/Uri outliers,
  2027-incomplete H1-focus, planned-only guardrail), colour key, and an
  **auto-generated "This Week's Reads"** block (top movers computed from the data).
- **Assumptions & Methodology** — single **source of truth**: every locked desk
  rule, the mogas yield map, baseline windows + scenario/forecast formula, the
  $-at-risk method, the **tested** (and rejected) crack↔outage relationship, and
  full data provenance — so a reviewer can audit the numbers without reading code.
- **Dashboard** — KPI tiles and headline charts.
- **Explorer** — *the interactive sheet.* Dropdowns for **PADD/unit**, **type**
  (All/Planned/Unplanned) and **measure** (Capacity vs Mogas-equivalent) drive a
  **SUMIFS** model: the monthly grid, the **month-by-month YoY%** ("Jan +7% vs
  last Jan"), and the chart all recompute live off the hidden `Data` sheet.
- **Trends** — annual table + a flattened (2020/21-capped) annual chart
  side-by-side, targeted comparisons incl. the **H1-vs-H1** like-for-like,
  Total/Planned/Unplanned monthly matrices with **in-cell sparklines**, and
  **month-by-month YoY%** grids (Planned shows 2027; Unplanned/Total n/a for 2027).
- **PADD** — per-PADD **two-axis** combo charts (2026 plan + unplanned-with-
  forecast-tail bars on the left axis; prior-year totals & 2027 plan lines on the
  right), laid out **two-per-row**, and the Total / Unplanned / Planned PADD×year
  matrices in a **2×2 grid** — so the sheet reads across, not just down.
- **Units & Refineries** — unit mix (share + YoY%), top refineries (autofiltered),
  operators and the event scatter.
- **Mogas** — gasoline-yield-weighted offline: yield map, annual + YoY%, and a
  stacked chart (the gasoline read).
- **Naphtha** — the octane complex (reforming + isomerization + aromatics/BTX):
  annual + YoY%, by-PADD and by-unit (octane/blending risk).
- **Margin Context** — values offline capacity in **$ of gross refining margin at
  risk** (`offline kbd × crack($/bbl) × days / 1000`). The gasoline crack is a
  **blue, Bloomberg-overwritable input** seeded from EIA (NY-Harbor gasoline − WTI);
  the $ figures are live formulas, with a two-axis offline-vs-crack overlay
  (margin-timing) and annual $-at-risk bars.
- **Events & TAs** — back-to-back FCC clusters (the **ExxonMobil Q1 turnaround**
  signal external trackers miss), the **ExxonMobil 2027 booked book** (by refinery
  / by unit + a refinery×month chart), and the 2026 & 2027 planned TA schedules.
- **Scenario Analysis** — live 2027 scenario (dropdown inputs → cascade → per-PADD
  split → **P25/P50/P90 bands**), the **Conservative/Average/Active fan** chart,
  the sensitivity heatmap, and a **scenario-summary** table + stacked column
  (booked planned + scenario unplanned = 2027 implied total).

Every sheet has a **Home** link and a data-vintage stamp. The build is
re-runnable and idempotent.

**Quality-of-life touches:** hover tooltips on metric headers, inputs and the
rationale/forecast cells; helper prompts when you click a dropdown; small-base
`n/m` and planned-only `n/a` guards so no nonsensical +1000% or -100% YoY's; a
live "Showing: …" echo of the active Explorer slice; one-click "Jump to"
cross-navigation on the Dashboard; a Quick-Start callout; and print-ready setup
(repeating title rows, page/date footers, fit-to-width).

### Deck (`outage_deck.pptx`)

Styled to a sell-side "weekly meeting" template: a navy title slide, white
content slides with a **red section header**, brand wordmark and page number,
and **dense layouts** — **two charts per content slide** (chart-pair + "key
takeaways" bullets), a five-PADD **small-multiples** regional overview, and a
**full-width national turnaround-schedule table**. Chart legends sit above the
plot where they would otherwise collide with the data. Set `BRAND_LOGO` in
`build_slides.py` to drop in your own logo (the layout reproduces the reference
look, not the trademarked mark).

### Dashboard (`outage_dashboard.html`)

A **dynamic focus-year selector** (KPIs show the picked year vs the prior year),
a **2026-vs-2025 / H1'27-vs-H1'26 / 2027-forecast** outlook strip, metric/PADD
filters, YoY bars, a range-band seasonality chart that **carries the 2026/27
forecast** (not zero), a monthly YoY% chart, an FCC-clusters table, a 2026
TA-schedule table (PADD selector) and the live 2027 scenario panel — all in one
self-contained file sharing the exact same palette and numbers as the workbook.

---

## Locked data decisions

These are desk rules baked into `engine.py` — don't re-derive them:

- **Primary metric:** `CAP_OFFLINE_ADJUSTED_KBD` (offline capacity, kbd, all
  units). Mogas-equivalent is a **secondary overlay only**.
- **Outage type** is binary `{PLANNED, UNPLANNED}`; `UNKNOWN → UNPLANNED`.
- **PADD** is parsed from `PAD_DIST` Roman numerals (100% resolved); a state map
  is the fallback.
- **2027 is planned-only.** Unplanned-2027 is a *modeled scenario*. Any
  Plan+Unplanned or Unplanned comparison vs 2027 shows **n/a**; only Planned is
  comparable.
- **2026 & 2027** are partial/special → rendered **grey italic** and footnoted.
  2027 planned is **incomplete past H1**, so H1-vs-H1 is the like-for-like read.
- **2020–2021** (COVID / Winter Storm Uri) are excluded from forecast baselines
  **and flattened on charts** (capped to the next-highest year) so they don't
  distort the trend; the real actuals stay in the tables/footnotes.
- **YoY% guards:** a tiny prior-year base shows `n/m` (not the +1000% blow-up);
  unplanned/total vs 2027 shows `n/a` (no unplanned-2027 actuals).
- **Event count** = distinct `OUTAGE_ID`s (rows are unit-month slices).

Scenario math (identical in workbook formulas, deck, and dashboard):

```
forecast[m] = baseline(window)[m] · (1 + growth) · multiplier
            + one_off          (added to the stress month only)
2027 unplanned = Σ forecast[m]
implied total  = 2027 unplanned + 2027 planned (booked)
```

---

## QA notes

- The deck and dashboard are **self-contained** — just open them; no external
  files or network are needed at view time.
- **`python scripts/audit_exxon.py`** reconciles the ExxonMobil slate against the
  vendored corporate turnaround plan (`data/exxon_ta_plan.csv`) and writes a
  per-unit report to `output/`.
- The build is idempotent — re-running reproduces the same outputs from the same
  input.
