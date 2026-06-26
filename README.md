# Refinery Outage Analytics

A trading-desk view of refinery outages built around **per-unit capacity
offline**: how much of each key unit (CDU, FCC, hydrocracker, reformer) is down,
by month, region and operator, with the forward outlook (the current year + 1)
front and center. It runs off the **Snowflake golden record** and is **dynamic**:
the Excel model is built around the same data, so a refreshed or expanded export
recomputes everything. The window is always **2023 .. current year + 1** (2023-2027
today) and rolls forward on its own. One data pipeline feeds three deliverables
that always agree:

1. **Slide decks** (`output/outage_deck.pptx`, `output/outage_deck_naphtha.pptx`):
   two chart-forward decks (presenter brings their own notes, so the slides are
   the charts). The **main deck** is gasoline/distillate-focused on the forward
   outlook year: total outages by unit, the biggest individual outages by PADD,
   the like-for-like H1 read per unit and month, outages by PADD by unit, the
   naphtha balance, recent unplanned context, and the unplanned scenario. The
   **naphtha / chem-feed deck** is the parallel octane & petrochemical-feedstock
   read, headlining the in-progress year ("rest of <yy>", H1 actual vs H2 booked)
   and tilted to **reformers**: rest-of-year outages by unit, the biggest outages,
   reformers (the octane read), the naphtha balance, the naphtha/octane/chem-feed
   complex (reforming + isomerization + aromatics/BTX), reformer & crude by PADD,
   and unplanned context. Both read the same engine context, so they agree.
2. **Excel model** (`output/outage_model.xlsx`): every number the deck looks at,
   calculates or forecasts, in detail, with the exact deck charts embedded on
   each sheet (just copy and paste). It is **live off the Snowflake**: the `Data`
   sheet holds the golden-record rows and every analysis sheet is `=SUMIFS` over
   it, with **Focus and PADD derived by Excel formula** (so pasted rows classify
   themselves) and spare formula rows below the data, so when you paste a refreshed
   or larger Snowflake into `Data` the whole model recomputes. The Naphtha and
   Forecast sheets additionally recompute off the shaded Assumptions input cells.
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

That reads `data/Golden_Record_Snowflake.xlsx` and writes all three outputs to
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

* **Snowflake golden record (primary):** the day-weighted monthly export
  (`CAP_OFFLINE_ADJUSTED_KBD`, `OUTAGE_YEAR`, `OUTAGE_MONTH`, `OUTAGE_TYPE`,
  `PAD_DIST`, `UNIT_CATEGORY`, `UNIT_NAME`, `REFINERY_OPERATOR`, `PLANT_NAME`, ...).
  `CAP_OFFLINE_ADJUSTED_KBD` is already day-weighted, so the desk metric is a
  straight `SUMIFS` over it — exactly what the live Excel model does. The window is
  clipped to 2023 .. current year + 1.
* **"Refinery Outages Enhanced" breakdown (also recognised):** one row per outage
  event (`Plant`, `Start Date`, `End Date`, `Unit Name`, `Outage Type`,
  `Offline Capacity (KBD)`, ...); each event is expanded to the months it spans and
  the free-text `Unit Name` is mapped to a unit class.

In both schemas, `UNIT_CATEGORY` maps to a focus class (CDU / FCC / hydrocracker /
reformer), and a **vacuum pipe still mislabelled `ATMOS DISTILLATION`** (name with
`VPS` or starting `VACUUM`) is demoted out of CDU — CDU is atmospheric crude only.
That same rule is mirrored in the Excel `Focus` formula, so the live model agrees.

---

## How the numbers are built

These rules live in `engine.py`. They are the desk conventions behind every
chart and table.

* **Per unit, never summed.** CDU, FCC, hydrocracker and reformer are each read
  on their own. A 250-kbd CDU plus a 100-kbd FCC is never "350 offline".
* **Day-weighted offline, summed.** Each Snowflake row's `CAP_OFFLINE_ADJUSTED_KBD`
  is already the day-weighted share of nameplate (nameplate x days-down /
  days-in-month), so a unit offline part of a month counts only for the days it is
  down. The monthly figure is a straight **sum** of those rows — exactly what the
  live `SUMIFS` model computes — so the deck and the workbook agree, and both keep
  working as the Snowflake grows.
* **CDU is atmospheric crude only.** Vacuum (VDU) is not folded in — including a
  vacuum pipe still the source mislabels `ATMOS DISTILLATION` (demoted by unit name).
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

`outage_model.xlsx`, fifteen sheets. The analysis computes off a single `Data`
source with **visible Excel formulas** (SUMIFS / AVERAGE / CORREL / SLOPE /
RSQ ...), so any number on a slide can be traced to a cell and you can see how it
is calculated. The shaded gold cells are editable inputs that the dependent
sheets recompute off. Every deck chart is embedded, and Historicals / Regression
add **native, live Excel charts**.

| Sheet | What's in it |
|---|---|
| **Index** | Maps each deck slide to its model sheet and how it is calculated. |
| **Assumptions** | As-of date, yields (incl. naphtha), reformer intake, scenario multipliers, baseline window, methodology. Drives the live sheets. |
| **Data** | The source = the Snowflake golden record (one row per year/month/plant/unit/type, day-weighted kbd + nameplate). `Focus` and `PADD` are Excel formulas; spare formula rows sit below the data. Paste a refresh here and everything recomputes. |
| **Historicals** | Monthly 2023-2027: total / planned / unplanned, unplanned %, per focus unit, per PADD, annual + YoY%, with a live line chart. |
| **Per-Unit** | CDU / FCC / hydrocracker / reformer concurrent offline by month and year, plus busiest-month peaks (=MAX). |
| **Biggest** | The biggest individual 2027 outages: refinery, unit, class, PADD, kbd, window, confirmed vs indicative. |
| **H1 by Unit** | H1 (Jan-Jun) planned offline per unit and month for 2025 / 2026 / 2027, plus the H1 averages (=AVERAGE). |
| **PADD by Unit** | 2027 CDU and FCC offline by PADD and month. |
| **Naphtha** | CDU supply vs reformer demand balance, **live** off the Assumptions yields. |
| **ExxonMobil** | Per-unit 2027 turnarounds, verified against the corporate plan. |
| **Forecast** | Completeness-aware baseline and the Conservative / Average / Active scenario, **live** off the multipliers, with implied totals and history bands. |
| **Sensitivity** | 2027 implied total across an unplanned-multiplier x one-off-shock grid (editable heatmap). |
| **Stress Test** | Named shocks (USGC hurricane, winter freeze, CDU trips, fall overlap) on the 2027 book, with tunable shock cells. |
| **Statistics** | Descriptive stats (mean / median / stdev / percentiles) and a Pearson correlation matrix over the historical series. |
| **Regression** | Least-squares best-fit (slope, intercept, R-squared, std err) on key relationships, each with a scatter + trendline chart. |

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
├── data/         live input: Golden_Record_Snowflake.xlsx (not committed)
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
