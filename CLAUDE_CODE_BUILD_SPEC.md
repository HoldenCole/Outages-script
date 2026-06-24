# Refinery Outage Workbook — Build Specification for Claude Code

**Audience:** Claude Code, building in the user's local repo with real Excel available for QA.
**Goal:** An institutional, sell-side/commodities-desk-grade Excel workbook generated from a Python script that ingests a Snowflake outage export and produces a polished, chart-rich, interactive `.xlsx`. It must be **re-runnable**: point at a new export, get the same workbook with updated numbers.

**Why this spec exists:** A prototype was built with `openpyxl`. Its charts and conditional-formatting heatmaps render as flat/weak output and do not meet the bar. This spec mandates `XlsxWriter` for charts, real native Excel chart objects, real heatmaps, real data tables, and a visual-QA loop in actual Excel. Build to the reference images the user provided (per-PADD stacked-bar + prior-year-line combo charts) and to standard hedge-fund / sell-side model conventions (two-way sensitivity heatmaps, tornado diagrams, scatter plots, KPI tiles, no gridlines, colored tabs).

---

## 0. Environment & dependencies

```bash
python -m venv .venv && source .venv/bin/activate      # or your env of choice
pip install pandas numpy xlsxwriter openpyxl
```

- **Primary chart/format engine: `XlsxWriter`** (write-only). It produces high-quality native Excel charts, supports `combine()` for column+line combos, secondary axes, per-series fills, markers, gradient fills, chart styles. Use it for the whole workbook build.
- Use `openpyxl` only if you need to post-process (e.g. inject a native data-table XML). Do **not** build the charts in openpyxl.
- **QA in real Excel** (Windows/Mac), not LibreOffice. LibreOffice flattens conditional-format color scales and some combo-chart styling, which is exactly what misled the prototype. Open the file, eyeball every chart and heatmap, confirm dropdowns recalc.

---

## 1. Input data contract

**File:** Snowflake export, e.g. `rEFINERY oUTAGES.xlsx`, sheet **`Query1`** (~46k rows, 2010–2038).
Path is configurable via a constant at top of script (`INPUT_PATH`). Strip leading/trailing whitespace from the path and from the sheet name — a past run failed on a leading space in the filename.

**Columns used (logical → source):**

| logical | source column | use |
|---|---|---|
| year | `OUTAGE_YEAR` | int |
| month | `OUTAGE_MONTH` | 1–12 |
| month_date | `OUTAGE_MONTH_DATE` | timestamp |
| start / end | `OUTAGE_START_DATE` / `OUTAGE_END_DATE` | dates |
| **cap_kbd** | `CAP_OFFLINE_ADJUSTED_KBD` | **PRIMARY metric**, capacity offline kbd |
| cap_raw | `OFFLINE_CAPACITY` | secondary |
| unit_cap | `UNIT_CAPACITY` | reference |
| duration | `TOTAL_OUTAGE_DAYS` | monthly-allocated days (caps at ~31), use as *intensity* not full event length |
| state | `REFINERY_STATE` | PADD fallback |
| city | `REFINERY_CITY` | detail |
| operator | `REFINERY_OPERATOR` | granular rows |
| plant | `PLANT_NAME` | granular rows |
| otype | `OUTAGE_TYPE` | PLANNED / UNPLANNED / UNKNOWN |
| otype2 | `OUTAGE_TYPE_LVL2` | detail |
| unit_name | `UNIT_NAME` | detail |
| unit_cat | `UNIT_CATEGORY` | 17 categories, for unit + mogas |
| pad_dist | `PAD_DIST` | **PADD source**, Roman numerals |
| cause | `OUTAGE_CAUSE` | detail |
| outage_id | `OUTAGE_ID` | event count |

### Cleaning rules (these are locked decisions — do not re-derive)

1. **Outage type:** map `UNKNOWN → UNPLANNED`. Final type is binary {PLANNED, UNPLANNED}. (Desk instruction: "when it's unknown, consider that unplanned.")
2. **PADD:** parse from `PAD_DIST` Roman numerals. Mapping: `PADD I→PADD 1, II→2, III→3, IV→4, V→5, Caribbean→PADD Caribbean`. If `PAD_DIST` ever null, fall back to a `REFINERY_STATE → PADD` map (table in Appendix A). In the current file PAD_DIST is 100% populated; the historical "all PADDs empty" bug was a Roman-numeral parsing failure, not missing data. **Do not** try to coerce digits out of the Roman string with `int()` — map explicitly.
3. **cap_kbd:** `pd.to_numeric(..., errors="coerce").fillna(0)`.
4. **Primary metric is capacity offline (kbd), all units and products.** Mogas is a *secondary overlay only* (Section 4.9). Do not make mogas the base.
5. **Partial years:** 2026 and 2027 have fewer rows than full years (data entry ongoing). Render 2026 & 2027 in **grey italic** everywhere and footnote them. Never present them as final.
6. **2027 is PLANNED-only.** There is no actual unplanned 2027 data. Unplanned 2027 is a *modeled scenario* (Section 4.8). Comparison guardrail: any "Plan+Unplanned" or "Unplanned" comparison **vs 2027 must show `n/a`**; only "Planned" comparisons against 2027 are valid.

Build a reusable `engine.py` (the prototype's `outage_engine.py` is a fine starting point — reuse its `load`, `annual_summary`, `padd_year_matrix`, `unit_year_matrix`, `seasonality`, `monthly_by_year`, `padd_month_year`, `operator_year`, `plant_detail` functions). The workbook builder imports from it. Keep data logic and presentation logic separate so a data refresh only touches inputs.

---

## 2. Global styling standard (apply to EVERY sheet)

- **Font:** Arial throughout. Titles 18pt bold navy; section bands 11pt bold white on color; body 10pt.
- **No gridlines anywhere:** `worksheet.hide_gridlines(2)` on every sheet.
- **Every tab colored** via `set_tab_color`. Suggested: Cover navy `#1F3864`, Dashboard red `#C00000`, analytic sheets blue `#2E5496`, PADD charts green `#548235`, Units gold `#BF9000`, Refinery detail purple `#7030A0`, Scenario red, Sensitivity orange `#C55A11`, Mogas green, Notes gray.
- **Palette (ExxonMobil-style):** navy `#1F3864`, blue `#2E5496`, red `#C00000`, gold `#BF9000`, green `#548235`, orange `#ED7D31`, light-blue fill `#D6E0F0`, light-gray fill `#F2F2F2`, yellow assumption `#FFF2CC`.
- **Number formats:** kbd `#,##0;(#,##0);"-"` (zeros show as dash, negatives in parens). Percent `0.0%`. Years as text `@` (so "2024" not "2,024"). Multiples `0.0"x"`.
- **Color-coding convention (financial-model standard):** blue font `#0000FF` = hardcoded inputs the user changes; black = formulas; green `#008000` = cross-sheet links; yellow fill = key assumption cells.
- **Banded section headers** (merged colored bar with white bold text) above each table block — not accent stripes, a full header bar.
- **Freeze panes** below the header row on data-heavy sheets. **Set print area**, landscape, fit-to-width.
- Alternating row shading (light-gray / white) on every multi-row table.
- **Zero formula errors.** After build, open in Excel and confirm no `#REF! #DIV/0! #VALUE! #N/A #NAME?`.

---

## 3. Reference charts to match (CRITICAL)

The user supplied reference images. Reproduce this exact chart archetype **as native Excel combo charts**, one per PADD:

> **"PADD _n_ Planned & Unplanned Offline (kbd)"** — a **combo chart**:
> - **Stacked columns:** current-year Plan (gold `#BF9000`) + current-year Unplanned estimate (light orange `#ED7D31`), stacked.
> - **Lines over the columns:** prior years as lines — 2025 (red), 2024 (blue), 2023 (gray) — each a smooth-free line, and where present a 2027-plan line (green).
> - Month categories Jan–Dec on X. kbd on Y. Legend on top. Title per PADD.

XlsxWriter recipe:
```python
col = wb.add_chart({'type': 'column', 'subtype': 'stacked'})
col.add_series({'name':..., 'categories':..., 'values':...,  # Plan
                'fill':{'color':'#BF9000'}})
col.add_series({'name':..., 'values':..., 'fill':{'color':'#ED7D31'}})  # Unplanned est
line = wb.add_chart({'type':'line'})
for yr,color in [('2025','#C00000'),('2024','#2E5496'),('2023','#808080'),('2027','#548235')]:
    line.add_series({'name':..., 'categories':..., 'values':..., 'line':{'color':color,'width':2.25}})
col.combine(line)
col.set_title({'name':'PADD 3 Planned & Unplanned Offline (kbd)'})
col.set_x_axis({'name':''}); col.set_y_axis({'name':'kbd'})
col.set_legend({'position':'top'})
col.set_size({'width': 720, 'height': 380})
ws.insert_chart('B20', col)
```
If a PADD's current-year bars are tiny vs prior-year lines (PADD 1 is genuinely small), that's data-accurate — keep one shared axis so magnitudes stay honest. Only use a secondary axis (`'y2_axis': True`) if the user later asks to emphasize shape over magnitude.

Provide the data each chart reads from in a **labeled data block directly above the chart on the same sheet** (series in rows, months in columns), so the chart is auditable and re-runs cleanly.

---

## 4. Sheet-by-sheet build

Order: **Cover · Dashboard · Summary · Monthly · PADD Charts · PADD Detail · Units · Refinery Detail · Scenario 2027 · Sensitivity · Mogas Overlay · Data Notes.**

### 4.1 Cover
Title block, scope, data vintage (`{n} rows | {min}-{max}`), primary-metric statement, the 2026/27 + 2027-planned-only caveat in red, and a clickable contents list (use `write_url` internal links `internal:'Sheet'!A1`).

### 4.2 Dashboard
- **KPI tile row** (4–5 tiles): latest full year total offline, unplanned, unplanned % of total, event count, top PADD. Big number (~24pt) on light-blue fill, label above on colored fill.
- **Stacked column** capacity-by-year (planned navy + unplanned red), 2018–2027.
- **Clustered column** unplanned-by-PADD across recent years.
- Optional **donut** of unplanned share by PADD (latest year).
- Charts native XlsxWriter, sized ~480×300, titled, legended.

### 4.3 Summary
- Annual table 2016–2027: Year, Planned, Unplanned, Total(formula), Events, Unpl%(formula), YoY ΔTotal(formula). Grey-italic 2026/27.
- **Targeted comparison blocks**: 2025v26, 2025v27, 2026v27, each with Plan+Unplanned / Planned / Unplanned rows. Enforce the **2027 guardrail** (n/a on non-planned vs-2027 blocks). Use real formulas for the Δ%.

### 4.4 Monthly
Three month×year matrices (Total, Planned, Unplanned) 2018–2027 with a Total column (`SUM`). Below: a **line chart** of unplanned-by-month with a series per recent year (the seasonality view).

### 4.5 PADD Charts  ← the reference deliverable
One combo chart **per PADD (1–5)** per Section 3, each with its own data block. This is the sheet the user most wants to see done right. Stack them vertically with clear PADD section bands.

### 4.6 PADD Detail
PADD×year matrices for Total / Unplanned / Planned, each with a "Total US" `SUM` row. Add a **clustered column** of unplanned-by-PADD.

### 4.7 Units
Unit-category × year matrix (all 17 categories), 2020–2027, sorted by total. **Bar chart** of the top categories. Add **data bars** (conditional format) on a "total" column for in-cell magnitude.

### 4.8 Refinery Detail (granular)
- **Top 15 refineries**: Refinery, PADD, Operator, Total, Planned, Unplanned, Events. **Data bars** on Total.
- **Top 10 operators × year** matrix.
- **Scatter plot**: event capacity (Y) vs monthly intensity/duration (X) for recent unplanned events; markers only, no line. Optionally color/series by PADD.

### 4.9 Scenario 2027 (driver-based, live)
Inputs panel (yellow fill, blue font, bordered), with **data-validation dropdowns**:
- Baseline window (dropdown: `2022-2025`, `2023-2025`, `2018-2019,2022-2025`, `All ex-2020/21`) — default `2022-2025`.
- Production growth % (input).
- Unplanned rate multiplier (input, the risk dial).
- One-off event adder (kbd) + Stress month (dropdown Jan–Dec).

Hidden/lookup **seasonality profile table** (avg unplanned kbd/month for each window). Forecast cascade as **live formulas**:
`Baseline (INDEX/MATCH on window) → × (1+growth) × multiplier → + one-off(IF stress month) → 2027 Unplanned forecast → + 2027 Planned booked → Implied total.`
Then a **line chart**: 2027 scenario vs 2027 planned vs 2024/2025 actual unplanned.
Also a **PADD allocation** mini-table splitting the scenario annual by historical PADD unplanned share (green cross-sheet-style links).

### 4.10 Sensitivity & Risk (the heatmaps)
- **Two-way sensitivity heatmap:** rows = production growth `{-10%,-5%,0,5%,10%,15%}`, cols = unplanned multiplier `{0.7,0.85,1.0,1.15,1.3,1.5}`, body = resulting 2027 unplanned kbd.
  - **Preferred:** build it as a **native Excel two-input Data Table** wired to the Scenario inputs so it recomputes through the full monthly model. XlsxWriter/openpyxl can't emit the `=TABLE()` array directly, so either (a) write the grid as explicit formulas `=$anchor*(1+g)*m` referencing a hard-coded baseline-annual anchor cell (robust, no macro — what the prototype did, keep as fallback), **or** (b) post-process the saved file to inject the `<f t="dataTable" .../>` array XML over the grid range, then verify recalculation in Excel. Implement (a) first so the sheet is never broken; attempt (b) as the upgrade and only keep it if Excel recalcs cleanly.
  - **Apply a real 3-color scale** (`conditional_format` `'type':'3_color_scale'`, green `#63BE7B` → yellow `#FFEB84` → red `#F8696B`). Outline the base-case cell (g=0, m=1.0) with a medium navy border.
  - Header row/col on navy with white bold; growth as `0.0%`, multiplier as `0.0"x"`.
- **Tornado diagram:** for each driver (unplanned multiplier ±30%, baseline-window swing ±15%, production growth ±10%, one-off event +300kbd), compute Low/Base/High and Swing, **sorted by descending swing**. Plot as a **stacked horizontal bar** centered on the base (classic tornado). Table beside the chart.
- Optional **spider/line** chart showing each driver swept across a range.

### 4.11 Mogas Overlay (secondary)
Clearly labeled secondary product view. Yield map table (unit category → bucket → factor): CDU 0.175, FCC 0.65, Ref 0.85, HDC 0.05, Coker 0.20; everything else 0 (see Appendix B for the full 17-category → bucket mapping). Mogas-equivalent = `cap_kbd × factor`. Show mogas annual planned/unplanned/total by year. Capacity is never discarded — mogas is additive.

### 4.12 Data Notes
Source, row count, primary metric, type/PADD logic, the 2027 + partial-year caveats, duration caveat, scenario method, sensitivity method, mogas method, refresh instructions, and the blue/black/green/yellow color key.

---

## 5. Interactivity requirements

- Scenario dropdowns via `data_validation` (`{'validate':'list','source':[...]}`).
- All scenario/sensitivity outputs are **formulas**, never Python-computed constants, so editing an input recalculates live.
- Cross-sheet links in green font.
- Consider a workbook-level **defined name** for the active baseline window so multiple sheets can reference one control.
- Freeze panes + autofilter on the big detail tables.

---

## 6. QA loop (do not skip — this is why v1 failed)

1. Build, then **open in real Excel** (not LibreOffice).
2. Verify every **chart**: combo charts show columns *and* lines; legends/titles present; colors correct; axes labeled; no empty/placeholder charts.
3. Verify the **heatmap** shows a true green→red gradient and the base case is outlined.
4. Verify the **tornado** is centered and sorted; **scatter** shows markers.
5. Change a **scenario dropdown / input** and confirm forecast, chart, sensitivity grid, and PADD allocation all recalc.
6. Confirm **no gridlines, all tabs colored, grey-italic partial years, dashes for zeros, no formula errors**.
7. Re-run the script against the file to confirm **idempotent, refreshable** output.

Deliver: `engine.py`, `build_workbook.py`, a short `README.md` (how to set `INPUT_PATH` and run), and the generated `outage_workbook.xlsx`.

---

## Appendix A — STATE → PADD fallback map
PADD 1: CT DE DC FL GA ME MD MA NH NJ NY NC PA RI SC VT VA WV
PADD 2: IL IN IA KS KY MI MN MO NE ND OH OK SD TN WI
PADD 3: AL AR LA MS NM TX
PADD 4: CO ID MT UT WY
PADD 5: AK AZ CA HI NV OR WA

## Appendix B — UNIT_CATEGORY → mogas bucket
CDU(0.175): ATMOS DISTILLATION, VACUUM DISTILLATION
FCC(0.65): FLUID CAT CRACKING
Ref(0.85): REFORMING
HDC(0.05): HYDROCRACKING, RESID_HYDROCRACKING
Coker(0.20): COKING, THERM CRACKING/VISBREAKING
Other(0.0): HYDROTREATING, ALKYLATION, ISOMERIZATION, ASPHALT, BTX, MTBE, AROMATICS, GAS PROCESSING, OTHER

## Appendix C — Known data facts (sanity checks)
- ~46,456 rows, years 2010–2038; PADD 100% resolved from PAD_DIST.
- Type split after UNKNOWN→UNPLANNED: ~33,221 unplanned / ~13,235 planned.
- 2027 = 621 rows, **all PLANNED** (15,923 kbd). Confirms planned-only.
- Unplanned seasonality peaks Feb (winter freeze) and Sep (turnaround season); summer trough.
- 2020–21 are COVID/Winter-Storm-Uri outliers (unplanned ~112k / ~52k kbd vs ~9–24k normal) — exclude from forecast baselines by default.
- PADD 3 ≈ 47% of unplanned; PADD 5 spiked in 2025 on California events.
- Top operators by capacity offline: Phillips 66, Marathon, Valero, ExxonMobil, Chevron, CITGO.
