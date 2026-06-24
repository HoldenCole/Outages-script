# Refinery Outage Analytics

Turns the Snowflake refinery-outage export into three commodities-desk-grade
deliverables from one data pipeline:

1. **Excel workbook** (`outage_workbook.xlsx`) — 12 sheets, native charts,
   two-way sensitivity heatmap, tornado, and a **live** 2027 scenario model
   driven by data-validation dropdowns. *(Priority 1.)*
2. **Slide deck** (`outage_deck.pptx`) — 11-slide 16:9 deck mirroring the
   workbook's charts with data-driven takeaways. *(Priority 2.)*
3. **HTML dashboard** (`outage_dashboard.html`) — a single self-contained file
   with KPI tiles, metric/PADD filters, and a live 2027 scenario panel.
   *(Priority 3.)*

All three read from the **same** `engine.build_context()` bundle, so the numbers
always agree. Everything is re-runnable: point at a refreshed export and rebuild.

---

## Quick start

```bash
pip install -r requirements.txt
python scripts/build_all.py
```

That reads `data/rEFINERY oUTAGES.xlsx` and writes the three deliverables to
`output/`. Paths resolve relative to the repo root, so it works from any
directory. Point at a different export or output folder with
`python scripts/build_all.py path/to/export.xlsx --outdir somewhere/`.

Build a single deliverable instead:

```bash
python scripts/build_workbook.py            # -> output/outage_workbook.xlsx
python scripts/build_slides.py              # -> output/outage_deck.pptx
python scripts/build_dashboard.py           # -> output/outage_dashboard.html
```

Each takes an optional input path and `--out`; with no args it uses
`data/rEFINERY oUTAGES.xlsx`. Leading/trailing whitespace in the path and sheet
name is stripped automatically.

---

## Repository layout

```
.
├── scripts/      pipeline code (run these)
│   ├── engine.py            data core — the only place raw data is touched
│   ├── charts.py            matplotlib renderers (shared by the deck)
│   ├── build_workbook.py    Excel workbook (XlsxWriter) — 12-sheet deliverable
│   ├── build_slides.py      PowerPoint deck (python-pptx)
│   ├── build_dashboard.py   self-contained HTML dashboard (Chart.js inlined)
│   └── build_all.py         orchestrator — loads once, builds all three
├── data/         live input: rEFINERY oUTAGES.xlsx
├── output/       generated deliverables (.xlsx / .pptx / .html)
├── reference/    example workbook, gasoline-weekly PDF, Yields.txt
├── docs/         CLAUDE_CODE_BUILD_SPEC.md, "What good output looks like"
├── legacy/       superseded openpyxl prototype (reference only)
├── README.md  ·  requirements.txt
```

---

## Refreshing with new data

Drop a new Snowflake export (same `Query1` schema) into `data/` and re-run
`python scripts/build_all.py` (or pass its path). Only `engine.py` touches raw
data; the builders consume its frames. The Excel **Scenario** and
**Sensitivity** sheets are live formulas, so you can also just edit the yellow
input cells in Excel without rebuilding.

---

## Workbook sheets

Consolidated into **8 findable sheets** (+ a hidden `Data` backing sheet):

`Cover · Dashboard · Explorer · Trends · PADD · Units & Refineries ·
Events & TAs · Model`

- **Cover** — contents (hyperlinks), caveats, colour key, and an
  **auto-generated "This Week's Reads"** block (top movers computed from the data).
- **Dashboard** — KPI tiles and headline charts.
- **Explorer** — *the interactive sheet.* Dropdowns for **PADD/unit**, **type**
  (All/Planned/Unplanned) and **measure** (Capacity vs Mogas-equivalent) drive a
  **SUMIFS** model: the monthly grid, the **month-by-month YoY%** ("Jan +7% vs
  last Jan"), and the chart all recompute live off the hidden `Data` sheet.
- **Trends** — annual table + comparisons + Total/Planned/Unplanned monthly
  matrices with **in-cell sparklines** and heatmap colour-scales, plus
  **month-by-month YoY%** grids.
- **PADD** — per-PADD combo charts (2026 plan+unplanned vs prior-year totals &
  2027 plan) and PADD×year matrices.
- **Units & Refineries** — unit mix (share + YoY%), a **Naphtha / octane
  complex** block (reforming + isomerization + aromatics/BTX — the octane read),
  top refineries (autofiltered), operators, event scatter and the mogas overlay.
- **Events & TAs** — back-to-back FCC clusters (the **ExxonMobil Q1 turnaround**
  signal external trackers miss) and the 2026 planned turnaround schedule.
- **Model** — live 2027 scenario (dropdown inputs → cascade → per-PADD split →
  **P25/P50/P90 bands**) and the sensitivity heatmap + tornado.

Every sheet has a **Home** link and a data-vintage stamp. The build is
re-runnable and idempotent.

**Quality-of-life touches:** hover tooltips on metric headers and inputs;
helper prompts when you click a dropdown; ▲▼ direction arrows on the YoY%
grids; a live "Showing: …" echo of the active Explorer slice; one-click
"Jump to" cross-navigation on the Dashboard; a Quick-Start callout; and
print-ready setup (repeating title rows, page/date footers, fit-to-width).

### Deck (`outage_deck.pptx`)

Styled to a sell-side "weekly meeting" template: a navy title slide, white
content slides with a **red section header**, brand wordmark and page number,
and **dense layouts** — multi-chart grids, combo-charts + "key takeaways"
bullets, range-band seasonality charts, and **full-width turnaround-schedule
tables** per PADD. Set `BRAND_LOGO` in `build_slides.py` to drop in your own
logo (the layout reproduces the reference look, not the trademarked mark).

### Dashboard (`outage_dashboard.html`)

KPI tiles, metric/PADD filters, YoY bars, a range-band seasonality chart, a
monthly YoY% chart, an FCC-clusters table, a **2026 TA-schedule table** (PADD
selector) and the live 2027 scenario panel — all in one self-contained file
sharing the exact same palette and numbers as the workbook.

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
- **2020–2021** (COVID / Winter Storm Uri) are excluded from forecast baselines
  by default.
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

- Open `outage_workbook.xlsx` in **real Excel** (not LibreOffice) for final
  visual QA — confirm the combo charts show columns *and* lines, the heatmap
  shows a green→red gradient with the base case outlined, and changing a
  scenario dropdown recalculates the forecast, chart, heatmap and PADD split.
- The deck and dashboard are self-contained; just open them.
- The build is idempotent — re-running reproduces the same outputs from the same
  input.
