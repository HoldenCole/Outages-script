# Refinery Outage Analytics

A trading-desk view of refinery outages built around **per-unit capacity
offline**: how much of each key unit (CDU, FCC, hydrocracker, reformer) is down,
by month, region and operator, with the 2027 outlook front and center. One data
pipeline feeds three deliverables that always agree:

1. **Slide deck** (`output/outage_deck.pptx`): an 8-slide, trader-focused deck.
   Total 2027 outages by unit, the biggest individual outages by PADD, the
   like-for-like H1 read per unit and month, where it tightens by PADD, the
   naphtha balance (CDU supply vs reformer demand), the verified ExxonMobil
   slate, and the 2027 unplanned scenario.
2. **Excel model** (`output/outage_model.xlsx`): every number the deck looks at,
   calculates or forecasts, in detail, with the exact deck charts embedded on
   each sheet (just copy and paste). The Naphtha and Forecast sheets are **live
   models**: change the shaded input cells and they recompute.
3. **HTML dashboard** (`output/outage_dashboard.html`): a single self-contained
   file with a focus-year selector, outlook strip and live 2027 scenario panel.

All three read the **same** `engine.build_context()` bundle, so the numbers
agree. Point at a refreshed export and rebuild.

---

## Quick start (VS Code or any terminal)

```bash
python -m venv .venv
source .venv/bin/activate            # macOS/Linux (.venv\Scripts\Activate.ps1 on Windows)
pip install -r requirements.txt
python scripts/build_all.py          # builds deck + Excel model + dashboard
```

That reads `data/Refinery_Outages_Enhanced.xlsx` and writes all three outputs to
`output/`. **The input data is not committed**: drop your outage export at that
path (or pass one as an argument). Paths resolve relative to the repo root, so it
runs from any directory, and no network is required to build.

Build a single deliverable, or point at a refreshed export / different output dir:

```bash
python scripts/build_slides.py                  # -> output/outage_deck.pptx
python scripts/build_workbook.py                # -> output/outage_model.xlsx
python scripts/build_dashboard.py               # -> output/outage_dashboard.html
python scripts/build_all.py path/to/export.xlsx --outdir dist/
```

In **VS Code**, the committed `.vscode/` configs give a one-click build (▶ Run,
or F5 "Build all deliverables", or the default build task).

### Run it on your own data

The engine **auto-detects the schema**:

* **"Refinery Outages Enhanced" breakdown (primary):** one row per outage event,
  columns `Source`, `Country/Region`, `Owner`, `Plant`, `Start Date`,
  `End Date`, `Unit Name`, `Outage Type`, `Offline Capacity (KBD)`. Each event is
  expanded to the calendar months it spans, and the free-text `Unit Name` is
  mapped to a unit class (CDU / FCC / hydrocracker / reformer). Only verified
  years (2023+) are kept; obvious placeholder spans are dropped.
* **Legacy Snowflake `Query1` export:** also recognised automatically
  (`CAP_OFFLINE_ADJUSTED_KBD`, `OUTAGE_YEAR`, `OUTAGE_MONTH`, plus the usual
  `OUTAGE_TYPE` / `PAD_DIST` / `UNIT_CATEGORY` / `REFINERY_OPERATOR` columns).

---

## How the numbers are built

These rules live in `engine.py`. They are the desk conventions behind every
chart and table.

* **Per unit, never summed.** CDU, FCC, hydrocracker and reformer are each read
  on their own. A 250-kbd CDU plus a 100-kbd FCC is never "350 offline".
* **Day-weighted concurrent offline.** For each month, every distinct physical
  unit is counted once at its days-down share of nameplate (nameplate x days-down
  / days-in-month). A unit offline only part of a month counts only for the days
  it is actually down, and it is never carried into a month it is back online.
* **CDU is atmospheric crude only.** Vacuum (VDU) is not folded in.
* **2027 completeness is asymmetric.** ExxonMobil gave a full-year plan (verified
  vs their corporate schedule); every other operator is H1-confirmed only, so
  non-Exxon H2 is shown as an indicative floor, not confirmed. H1 is the honest
  cross-year window.
* **Completeness-aware forecast baseline.** The 2027 unplanned scenario is the
  mean monthly shape of unplanned offline over a window (default 2023-2026), where
  each calendar month is averaged only over the years that actually reported it.
  The fresh 2026 H1 actuals sharpen the H1 baseline while H2 stays on the three
  complete years (2023-2025); annual range bands use complete years only.
* **Naphtha balance.** Crude makes naphtha (~35% of the barrel); reformers
  consume it (their charge is naphtha). A CDU outage removes naphtha supply, a
  reformer outage removes demand:
  `net = reformer_offline x intake (~1.0) - CDU_offline x naphtha_yield (~0.35)`.
  Net below zero is a naphtha deficit (short), above zero a surplus (long).
* **Outage type** is binary `{PLANNED, UNPLANNED}`; `UNKNOWN` folds into
  unplanned. 2027 is planned-only, so its unplanned figure is always a scenario.

The yield assumptions (gasoline/mogas yields, the naphtha yield and reformer
intake, the scenario multipliers) are tunable cells on the Excel **Assumptions**
sheet, and the Naphtha and Forecast sheets recompute live off them.

---

## Excel model sheets

`outage_model.xlsx`, eight sheets, each with its detail table and the matching
deck chart embedded:

| Sheet | What's in it |
|---|---|
| **Assumptions** | As-of date, yields (incl. naphtha), reformer intake, scenario multipliers, baseline window, methodology. Shaded input cells drive the live sheets. |
| **Per-Unit** | CDU / FCC / hydrocracker / reformer concurrent offline by month and year (day-weighted), plus busiest-month peaks. |
| **Biggest** | The biggest individual 2027 outages: refinery, unit, class, PADD, kbd, window, confirmed vs indicative. |
| **H1 by Unit** | H1 (Jan-Jun) planned offline per unit and month for 2025 / 2026 / 2027, plus the H1 averages. |
| **PADD by Unit** | 2027 CDU and FCC offline by PADD and month. |
| **Naphtha** | CDU supply vs reformer demand balance, **live** off the Assumptions yields, with the monthly net. |
| **ExxonMobil** | Per-unit 2027 turnarounds, verified against the corporate plan. |
| **Forecast** | Completeness-aware baseline and the Conservative / Average / Active scenario, **live** off the multipliers, with implied totals and history bands. |

---

## Repository layout

```
.
├── scripts/
│   ├── engine.py            data core: the only place raw data is touched
│   ├── charts.py            matplotlib renderers (shared by deck and workbook)
│   ├── build_slides.py      PowerPoint deck (python-pptx)
│   ├── build_workbook.py    Excel model (XlsxWriter), charts embedded
│   ├── build_dashboard.py   self-contained HTML dashboard (Chart.js inlined)
│   ├── build_all.py         orchestrator: load once, render once, build all three
│   └── audit_exxon.py       reconcile the Exxon slate vs the corporate plan
├── data/         live input: Refinery_Outages_Enhanced.xlsx (not committed)
├── output/       generated deliverables (.pptx / .xlsx / .html)
├── docs/         build spec and "what good output looks like"
├── README.md  ·  ROADMAP.md  ·  requirements.txt
```

---

## QA notes

* The deck, model and dashboard are **self-contained**: open them, no external
  files or network needed at view time.
* The build is idempotent: re-running reproduces the same outputs from the same
  input.
* `python scripts/audit_exxon.py` reconciles the ExxonMobil slate against the
  vendored corporate turnaround plan (`data/exxon_ta_plan.csv`) and writes a
  per-unit report to `output/`.
