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
python build_all.py "rEFINERY oUTAGES.xlsx"
```

Outputs land next to the script: `outage_workbook.xlsx`, `outage_deck.pptx`,
`outage_dashboard.html`. Use `--outdir dist` to write them elsewhere.

Build a single deliverable instead:

```bash
python build_workbook.py  "rEFINERY oUTAGES.xlsx" --out outage_workbook.xlsx
python build_slides.py    "rEFINERY oUTAGES.xlsx" --out outage_deck.pptx
python build_dashboard.py "rEFINERY oUTAGES.xlsx" --out outage_dashboard.html
```

If no path is given, each script falls back to the `INPUT_PATH` constant at the
top of the file. Leading/trailing whitespace in the path and sheet name is
stripped automatically.

---

## Refreshing with new data

Drop a new Snowflake export (same `Query1` schema) in place and re-run
`build_all.py`. Only `engine.py` touches raw data; the builders consume its
frames. The Excel **Scenario** and **Sensitivity** sheets are live formulas, so
you can also just edit the yellow input cells in Excel without rebuilding.

---

## Files

| File | Role |
|---|---|
| `engine.py` | Data core: load/clean, all pivots, scenario/sensitivity math, `build_context()` bundle. **The only place raw data is touched.** |
| `build_workbook.py` | Excel workbook (XlsxWriter) — the 12-sheet deliverable. |
| `charts.py` | Matplotlib renderers (desk palette) shared by the deck. |
| `build_slides.py` | PowerPoint deck (python-pptx). |
| `build_dashboard.py` | Self-contained HTML dashboard (Chart.js vendored inline). |
| `build_all.py` | Orchestrator — loads once, builds all three. |
| `CLAUDE_CODE_BUILD_SPEC.md` | The original build specification. |
| `outage_workbook.py`, `outage_monthly.py`, `outage_analyzer.py` | Legacy openpyxl prototype, moved to `legacy/` for reference. Superseded by the scripts above. |

---

## Workbook sheets

`Cover · Dashboard · Summary · Monthly · PADD Charts · PADD Detail · Units ·
Refinery Detail · Scenario 2027 · Sensitivity · Mogas Overlay · Notes`

- **PADD Charts** — one native combo chart per PADD: 2026 plan+unplanned stacked
  columns with prior-year total lines and the 2027 plan line.
- **Scenario 2027** — yellow input cells + dropdowns (baseline window, growth,
  unplanned multiplier, one-off, stress month). The forecast cascade, the line
  chart, and the PADD allocation are all live formulas.
- **Sensitivity** — a 6×6 growth × multiplier heatmap (3-colour scale, base case
  outlined) wired to the scenario baseline, plus a tornado of the drivers.

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
