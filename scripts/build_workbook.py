#!/usr/bin/env python3
"""
build_workbook.py
The Excel model. Everything the deck looks at, calculates or forecasts lives here
in detail, and the exact deck charts are embedded on each sheet so they are just
copy-and-paste. Built from the same engine.build_context() bundle as the deck and
dashboard, so the three always agree.

Design intent (this is the starting point of a model that will eventually live
fully in Excel): a single `Data` sheet holds the pullable source records, and the
analysis sheets compute off it with visible formulas (SUMIFS / AVERAGE), so any
number on a slide can be pointed to here and you can see how it is calculated. The
`Index` sheet maps each deck slide to its model sheet. The Naphtha and Forecast
sheets are live: change the shaded input cells on Assumptions and they recompute.

Sheets:
    Index         deck slide -> model sheet -> how it is calculated
    Assumptions   yields (incl. naphtha), reformer intake, scenario multipliers,
                  baseline window, as-of date, methodology   (tunable input cells)
    Data          one row per (year, month, plant, unit, type): the pullable source
    Per-Unit      CDU/FCC/HDC/Reformer concurrent offline by month & year (=SUMIFS)
    Biggest       the biggest individual 2027 outages, by PADD   (from Data)
    H1 by Unit    H1 (Jan-Jun) planned per unit & month, 2025/26/27 (=SUMIFS / AVG)
    PADD by Unit  CDU & FCC offline by PADD & month, 2027        (=SUMIFS)
    Naphtha       CDU supply vs reformer demand balance (=SUMIFS x Assumptions)
    ExxonMobil    per-unit 2027 turnarounds, verified            (from Data)
    Forecast      baseline + Conservative/Average/Active scenario (live formulas)

Usage:
    python build_workbook.py                       # uses the default input
    python build_workbook.py path/to/export.xlsx --out output/outage_model.xlsx
"""
import argparse
import sys
import tempfile
from pathlib import Path

import numpy as np
import pandas as pd
import xlsxwriter

import engine
import charts

_ROOT = Path(__file__).resolve().parent.parent
INPUT_PATH = str(_ROOT / "data" / "Golden_Record_Snowflake.xlsx")
OUT_PATH = str(_ROOT / "output" / "outage_model.xlsx")
ASOF = "June 26th, 2026"

MONTHS = engine.MONTHS
FOCUS = engine.FOCUS_ORDER
IMG = {"x_scale": 0.6, "y_scale": 0.6, "object_position": 1}

# Rolling window: 2023 .. current year + 1 (FY = the forward outlook year). Every
# sheet keys off these, so when the Snowflake (and the year) roll forward the model
# extends itself instead of dropping a year.
Y0 = engine.START_YEAR
FY = engine.FOCUS_YEAR
CY = engine.CURRENT_YEAR
YEARS = list(range(Y0, FY + 1))

# Data sheet column letters (keep in sync with _data_sheet header order)
D = {"year": "A", "mnum": "B", "month": "C", "plant": "D", "operator": "E",
     "padd": "F", "focus": "G", "unit_cat": "H", "unit_name": "I", "type": "J",
     "cap_kbd": "K", "cap_raw": "L"}


def _si(value_col, *conds):
    """Build a full-column SUMIFS over the Data sheet. conds = (col_letter, value)."""
    parts = [f"Data!${value_col}:${value_col}"]
    for col, val in conds:
        v = f'"{val}"' if isinstance(val, str) else val
        parts.append(f"Data!${col}:${col},{v}")
    return "=SUMIFS(" + ",".join(parts) + ")"


def _formats(wb):
    base = {"font_name": "Calibri", "font_size": 11}
    navy = "#1F3864"
    return {
        "title": wb.add_format({**base, "bold": True, "font_size": 16, "font_color": navy}),
        "sub": wb.add_format({**base, "italic": True, "font_color": "#808080"}),
        "h": wb.add_format({**base, "bold": True, "font_color": "white", "bg_color": navy,
                            "align": "center", "valign": "vcenter", "border": 1, "border_color": "white"}),
        "hl": wb.add_format({**base, "bold": True, "font_color": "white", "bg_color": navy,
                             "align": "left", "valign": "vcenter", "border": 1, "border_color": "white"}),
        "rowh": wb.add_format({**base, "bold": True, "bg_color": "#D6E0F0", "border": 1,
                               "border_color": "#BFBFBF"}),
        "num": wb.add_format({**base, "num_format": "#,##0", "border": 1, "border_color": "#E2E2E2"}),
        "f2": wb.add_format({**base, "num_format": "0.00", "border": 1, "border_color": "#E2E2E2"}),
        "pct": wb.add_format({**base, "num_format": "0%", "border": 1, "border_color": "#E2E2E2"}),
        "inph": wb.add_format({**base, "num_format": "#,##0", "bg_color": "#FFF2CC", "border": 1,
                               "border_color": "#BF9000", "align": "center"}),
        "txt": wb.add_format({**base, "border": 1, "border_color": "#E2E2E2"}),
        "txtw": wb.add_format({**base, "text_wrap": True, "valign": "top"}),
        "inp": wb.add_format({**base, "bold": True, "num_format": "0.00", "bg_color": "#FFF2CC",
                              "border": 1, "border_color": "#BF9000", "align": "center"}),
        "lbl": wb.add_format({**base, "bold": True}),
        "key": wb.add_format({**base, "italic": True, "font_color": "#BF9000"}),
        "secn": wb.add_format({**base, "bold": True, "font_size": 12, "font_color": navy,
                               "bottom": 2, "border_color": navy}),
    }


def _title(ws, fm, title, sub):
    ws.write(0, 0, title, fm["title"])
    ws.write(1, 0, sub, fm["sub"])
    ws.set_row(0, 22)


# --------------------------------------------------------------------------- sheets
def _index(wb, fm):
    ws = wb.add_worksheet("Index")
    ws.set_column(0, 0, 7); ws.set_column(1, 1, 40); ws.set_column(2, 2, 14)
    ws.set_column(3, 3, 78)
    _title(ws, fm, f"Refinery Outages Model {Y0}-{FY}: Index",
           f"Live off the Snowflake golden record ({Y0}-{FY} = 2023 .. current year + 1). Every deck slide "
           "maps to a model sheet; open the sheet to see the detail and the formula.")
    rows = [
        ("Slide", "Deck slide", "Model sheet", "What it shows and how it is calculated"),
        ("2", f"Total {FY} Outages by Unit", "Per-Unit",
         "Day-weighted concurrent offline per unit by month (=SUMIFS over Data), split confirmed vs indicative."),
        ("3", "What's Driving the Numbers (biggest outages)", "Biggest",
         f"The biggest individual {FY} outages by nameplate kbd, with PADD; one row per physical unit."),
        ("4", "H1 Planned per Unit & Month", "H1 by Unit",
         f"H1 (Jan-Jun) planned offline per unit and month, {FY-2}/{str(FY-1)[2:]}/{str(FY)[2:]} "
         "(=SUMIFS, type=Planned); H1 avg =AVERAGE."),
        ("5", "Outages by PADD by Unit", "PADD by Unit",
         f"{FY} CDU and FCC offline by PADD and month (=SUMIFS over Data)."),
        ("6", "Naphtha Balance", "Naphtha",
         "net = reformer offline x intake  -  CDU offline x naphtha yield. Live: =SUMIFS x Assumptions cells."),
        ("7", f"ExxonMobil {FY} by Unit", "ExxonMobil",
         f"Per-unit ExxonMobil {FY} turnarounds, verified against the corporate plan."),
        ("8", f"{FY} Unplanned Scenario", "Forecast",
         "baseline x multiplier (live off Assumptions); implied offline read by month - peak & average month (annual Σ kept as reference)."),
    ]
    r0 = 3
    for j, h in enumerate(rows[0]):
        ws.write(r0, j, h, fm["hl"] if j != 0 else fm["h"])
    for i, row in enumerate(rows[1:], 1):
        for j, v in enumerate(row):
            ws.write(r0 + i, j, v, fm["txt"] if j != 2 else fm["rowh"])
    ws.write(r0 + len(rows) + 1, 0,
             "Source of truth: the Data sheet (one row per year/month/plant/unit/type). The analysis "
             "sheets compute off it so the calculations are visible.", fm["key"])
    # analysis tabs beyond the deck
    ar = r0 + len(rows) + 3
    ws.write(ar, 0, "Analysis tabs (beyond the deck)", fm["secn"])
    extra = [
        ("What's Changed", "Rolling MoM (live) + week-over-week pull log + this month's new / back-online movers."),
        ("Historicals", f"Monthly {Y0}-{FY}: total/planned/unplanned, unplanned %, per unit, per PADD, YoY, chart."),
        ("Statistics", "Descriptive stats (mean/median/stdev/percentiles) and a correlation matrix."),
        ("Regression", "Least-squares best-fit (slope, R-squared, std err) + scatter trendlines."),
        ("Sensitivity", f"{FY} peak-month implied offline across multiplier x one-off shock (editable heatmap)."),
        ("Stress Test", f"Named shocks (hurricane, freeze, CDU trips) on the {FY} book, tunable."),
        ("PADD Connectivity", "Effective crude-outage impact = nominal x per-PADD pass-through (P3 Gulf buffered); tunable."),
        ("Data Quality", "Auto-flags: planned turnarounds recurring <5yr apart, and unit-months summing >100% of nameplate."),
    ]
    for i, (sh, what) in enumerate(extra):
        ws.write(ar + 1 + i, 2, sh, fm["rowh"])
        ws.write(ar + 1 + i, 3, what, fm["txt"])


def _assumptions(wb, fm, ctx):
    ws = wb.add_worksheet("Assumptions")
    ws.set_column(0, 0, 32); ws.set_column(1, 1, 14); ws.set_column(2, 4, 13); ws.set_column(5, 5, 58)
    _title(ws, fm, "Assumptions & Methodology",
           "Tunable cells are shaded gold. The Naphtha and Forecast sheets recompute live off them.")
    through = engine.latest_actual_month(ctx["df"])
    ws.write(3, 0, "As of", fm["lbl"]); ws.write(3, 1, ASOF, fm["txt"])
    ws.write(4, 0, "Actuals reported through", fm["lbl"])
    ws.write(4, 1, f"{MONTHS[through[1]-1]} {through[0]}" if through else "n/a", fm["txt"])
    ws.write(5, 0, "Data scope", fm["lbl"])
    ws.write(5, 1, f"{ctx['diag']['years'][0]}-{ctx['diag']['years'][1]} (2023+ verified)", fm["txt"])

    ws.write(7, 0, "Naphtha balance inputs", fm["secn"])
    ws.write(8, 0, "Naphtha yield (per bbl crude)", fm["lbl"])
    ws.write_number(8, 1, engine.NAPHTHA_YIELD, fm["inp"])
    ws.write(9, 0, "Reformer naphtha intake (per bbl capacity)", fm["lbl"])
    ws.write_number(9, 1, engine.REFORMER_NAPHTHA_INTAKE, fm["inp"])
    wb.define_name("naph_yield", "=Assumptions!$B$9")
    wb.define_name("ref_intake", "=Assumptions!$B$10")

    ws.write(11, 0, "Scenario multipliers (x baseline)", fm["secn"])
    for i, (nm, val) in enumerate([("Conservative", 0.8), ("Average", 1.0), ("Active", 1.3)]):
        ws.write(12 + i, 0, nm, fm["lbl"]); ws.write_number(12 + i, 1, val, fm["inp"])
    wb.define_name("mult_cons", "=Assumptions!$B$13")
    wb.define_name("mult_avg", "=Assumptions!$B$14")
    wb.define_name("mult_act", "=Assumptions!$B$15")
    ws.write(16, 0, "Baseline window", fm["lbl"]); ws.write(16, 1, ctx["scenario"]["window"], fm["txt"])

    ws.write(7, 3, "Gasoline (mogas) yields", fm["secn"])
    ws.write(8, 3, "Unit bucket", fm["hl"]); ws.write(8, 4, "yield", fm["h"])
    for i, (b, fac) in enumerate(engine.YIELD_FACTOR.items()):
        ws.write(9 + i, 3, b, fm["rowh"]); ws.write_number(9 + i, 4, fac, fm["f2"])

    notes = [
        "Per-unit, never summed: CDU, FCC, hydrocracker and reformer are read on their own.",
        "Day-weighted concurrent offline: a unit offline part of a month counts only for its days down "
        "(nameplate x days-down / days-in-month), each unit once per month.",
        "CDU = atmospheric crude only (vacuum / VDU is not folded in).",
        f"{FY} completeness: ExxonMobil gave a full-year plan (verified vs their schedule); every other "
        "operator is H1-confirmed only, so non-Exxon H2 is indicative.",
        "Forecast baseline is completeness-aware: each calendar month is averaged over the years that "
        "actually reported it, so 2026 H1 sharpens H1 while H2 stays on 2023-2025.",
        "Naphtha balance: net = reformer demand removed minus CDU supply removed (+ surplus / - deficit).",
    ]
    # PADD connectivity: crude-outage pass-through (tunable; PADD Connectivity sheet is live off these)
    ws.write(17, 0, "PADD connectivity: crude-outage pass-through", fm["secn"])
    ws.write(18, 0, "Share of a CDU outage that cascades downstream", fm["hl"])
    ws.write(18, 1, "pass-thru", fm["h"])
    for i, p in enumerate(["PADD 1", "PADD 2", "PADD 3", "PADD 4", "PADD 5"]):
        ws.write(19 + i, 0, p, fm["lbl"]); ws.write_number(19 + i, 1, engine.PADD_CONNECTIVITY[p], fm["inp"])
        wb.define_name(f"conn_p{i+1}", f"=Assumptions!$B${20+i}")
    ws.merge_range(24, 0, 24, 5, "Low = well-connected (P3 Gulf): the units downstream of the crude unit keep "
                   "running on piped-in intermediates. High (P2 / P4 / P5, parts of P1): islanded, so a crude "
                   "outage cascades and the downstream units are cut too.", fm["txtw"])
    ws.write(26, 0, "Methodology", fm["secn"])
    for i, n in enumerate(notes):
        ws.merge_range(27 + i, 0, 27 + i, 5, f"- {n}", fm["txtw"])


DATA_BUFFER = 6000           # spare formula rows so pasted Snowflake refreshes self-classify


def _data_sheet(wb, fm, base):
    """The model source = the Snowflake golden record (2023 .. current year + 1).
    Paste a refreshed Snowflake into the value columns; Focus and PADD are live
    Excel formulas (col H/M -> focus/padd), so new rows classify themselves and the
    analysis SUMIFS keep working as the file grows. cap_kbd = CAP_OFFLINE_ADJUSTED_KBD."""
    g = base.sort_values(["year", "month", "cap_kbd"], ascending=[True, True, False])
    recs = g.to_dict("records")
    ws = wb.add_worksheet("Data")
    ws.set_column(0, 1, 6); ws.set_column(2, 2, 6); ws.set_column(3, 3, 26); ws.set_column(4, 4, 24)
    ws.set_column(5, 6, 12); ws.set_column(7, 7, 20); ws.set_column(8, 8, 24); ws.set_column(9, 9, 10)
    ws.set_column(10, 12, 10)
    _title(ws, fm, "Data: Snowflake golden record (paste refreshes here; Focus/PADD auto-derive)",
           "Update = paste OUTAGE_YEAR/MONTH->A,B; PLANT/OPERATOR->D,E; UNIT_CATEGORY->H; UNIT_NAME->I; "
           "OUTAGE_TYPE->J; CAP_OFFLINE_ADJUSTED_KBD->K; OFFLINE_CAPACITY->L; PAD_DIST->M. Rest auto-fills.")
    heads = ["year", "monthnum", "month", "plant", "operator", "padd", "focus", "unit_cat",
             "unit_name", "type", "cap_kbd", "cap_raw", "pad_dist"]
    r0 = 3
    for j, h in enumerate(heads):
        ws.write(r0, j, h, fm["h"])
    first = r0 + 1                                       # 0-based first data row
    for i in range(len(recs) + DATA_BUFFER):
        rr = first + i
        xl = rr + 1                                      # 1-based Excel row
        live = i < len(recs)
        r = recs[i] if live else None
        # Focus (G) and PADD (F) are formulas on every row so pasted data self-classifies
        ws.write_formula(rr, 5, _padd_formula(xl), fm["txt"], str(r["padd"]) if live else "")
        ws.write_formula(rr, 6, _focus_formula(xl), fm["txt"], str(r["focus"]) if live else "")
        if not live:
            continue
        m = int(r["month"])
        ws.write_number(rr, 0, int(r["year"]), fm["txt"])
        ws.write_number(rr, 1, m, fm["txt"])
        ws.write(rr, 2, MONTHS[m - 1], fm["txt"])
        ws.write(rr, 3, str(r["plant"]), fm["txt"])
        ws.write(rr, 4, str(r["operator"]), fm["txt"])
        ws.write(rr, 7, str(r["unit_cat"]), fm["txt"])
        ws.write(rr, 8, str(r["unit_name"]), fm["txt"])
        ws.write(rr, 9, str(r["type"]), fm["txt"])
        ws.write_number(rr, 10, float(r["cap_kbd"]), fm["num"])
        ws.write_number(rr, 11, float(r["cap_raw"]) if pd.notna(r["cap_raw"]) else 0.0, fm["num"])
        ws.write(rr, 12, str(r["pad_dist"]), fm["txt"])
    ws.freeze_panes(first, 0)
    ws.autofilter(r0, 0, first + len(recs) - 1, len(heads) - 1)
    return len(recs)


def _matrix_formula(ws, fm, r0, c0, row_keys, col_keys, formula, cache, row_label):
    """Write a labelled matrix of formulas. formula(rk, ck)->str, cache(rk, ck)->float."""
    ws.write(r0, c0, row_label, fm["hl"])
    for j, ck in enumerate(col_keys):
        ws.write(r0, c0 + 1 + j, ck, fm["h"])
    for i, rk in enumerate(row_keys):
        ws.write(r0 + 1 + i, c0, rk, fm["rowh"])
        for j, ck in enumerate(col_keys):
            ws.write_formula(r0 + 1 + i, c0 + 1 + j, formula(rk, ck), fm["num"], float(cache(rk, ck)))
    return r0 + 1 + len(row_keys)


def _per_unit(wb, fm, ctx, assets):
    ws = wb.add_worksheet("Per-Unit")
    ws.set_column(0, 0, 8); ws.set_column(1, 12, 8)
    _title(ws, fm, "Per-Unit Concurrent Capacity Offline (kbd, day-weighted) = SUMIFS over Data",
           "Each focus unit on its own, by month and year. Never summed across units.")
    mnum = {mo: i + 1 for i, mo in enumerate(MONTHS)}
    years = YEARS
    blocks = {}
    r = 3
    for f in FOCUS:
        m = ctx["focus_monthly"][f]
        ws.write(r, 0, engine.FOCUS_LABEL[f], fm["secn"])
        top = r + 1
        r = _matrix_formula(
            ws, fm, top, 0, years, MONTHS,
            lambda y, mo, _f=f: _si("K", ("A", y), ("B", mnum[mo]), ("G", _f)),
            lambda y, mo, _m=m: (_m.loc[y, mo] if y in _m.index else 0.0), "Year") + 2
        blocks[f] = top                     # header row; data rows top+1..top+len(years)
    # busiest-month peak = MAX across that unit/year's 12 monthly cells above
    fp = ctx["focus_peak"]
    ws.write(r, 0, "Busiest-month concurrent offline (=MAX of the unit's months)", fm["secn"])
    hdr = r + 1
    ws.write(hdr, 0, "Year", fm["hl"])
    for j, f in enumerate(FOCUS):
        ws.write(hdr, 1 + j, engine.FOCUS_LABEL[f], fm["h"])
    for i, y in enumerate(years):
        ws.write(hdr + 1 + i, 0, y, fm["rowh"])
        for j, f in enumerate(FOCUS):
            drow = blocks[f] + 1 + i        # 0-based data row of (f, y)
            ws.write_formula(hdr + 1 + i, 1 + j, f"=MAX(B{drow+1}:M{drow+1})", fm["num"],
                             float(fp.loc[y, f] if y in fp.index else 0.0))
    ws.insert_image(3, 14, assets["splits_2027"], IMG)


def _biggest(wb, fm, ctx, assets):
    ws = wb.add_worksheet("Biggest")
    ws.set_column(0, 0, 30); ws.set_column(1, 1, 26); ws.set_column(2, 2, 10)
    ws.set_column(3, 3, 9); ws.set_column(4, 4, 9); ws.set_column(5, 5, 9)
    ws.set_column(6, 6, 9); ws.set_column(7, 7, 13)
    _title(ws, fm, f"Biggest {FY} Outages by Unit (nameplate kbd, from Data)",
           "Individual units, never added. PADD shows where the tonnage sits.")
    ev = engine.unit_events(ctx["df"], year=FY)
    ev = ev[ev["focus"].isin(FOCUS)].copy()
    ev["is_exxon"] = ev["operator"].astype(str).str.upper().str.contains("EXXON", na=False)
    ev["status"] = [("confirmed" if (x or min(ms) <= 6) else "indicative (H2)")
                    for x, ms in zip(ev["is_exxon"], ev["months"])]
    ev = ev.sort_values("kbd", ascending=False).head(30)
    heads = ["Refinery", "Unit", "Class", "PADD", "kbd", "Window", "Type", "Status"]
    r0 = 3
    for j, h in enumerate(heads):
        ws.write(r0, j, h, fm["h"] if j >= 2 else fm["hl"])
    for i, (_, row) in enumerate(ev.iterrows()):
        rr = r0 + 1 + i
        ws.write(rr, 0, str(row["plant"]).replace(" Refinery", ""), fm["txt"])
        ws.write(rr, 1, str(row["unit_name"]), fm["txt"])
        ws.write(rr, 2, str(row["focus"]), fm["txt"])
        ws.write(rr, 3, str(row["padd"]), fm["txt"])
        ws.write_number(rr, 4, float(row["kbd"]), fm["num"])
        ws.write(rr, 5, str(row["span"]), fm["txt"])
        ws.write(rr, 6, str(row["type"]).title(), fm["txt"])
        ws.write(rr, 7, str(row["status"]), fm["txt"])
    ws.insert_image(3, 9, assets["biggest_outages"], IMG)


def _h1(wb, fm, ctx, assets):
    ws = wb.add_worksheet("H1 by Unit")
    ws.set_column(0, 0, 10); ws.set_column(1, 7, 9)
    _title(ws, fm, "H1 (Jan-Jun) Planned Offline per Unit & Month = SUMIFS (type=Planned)",
           f"Like-for-like cross-year read; {FY} is confirmed through H1. H1 avg = AVERAGE of the row.")
    mnum = {mo: i + 1 for i, mo in enumerate(MONTHS)}
    h1m = MONTHS[:6]
    fp = ctx["focus_planned"]
    blocks = {}
    r = 3
    for f in FOCUS:
        m = fp[f]
        ws.write(r, 0, engine.FOCUS_LABEL[f], fm["secn"])
        top = r + 1
        r = _matrix_formula(
            ws, fm, top, 0, [FY - 2, FY - 1, FY], h1m,
            lambda y, mo, _f=f: _si("K", ("A", y), ("B", mnum[mo]), ("G", _f), ("J", "PLANNED")),
            lambda y, mo, _m=m: (_m.loc[y, mo] if y in _m.index else 0.0), "Year") + 2
        blocks[f] = top                     # header row of this block; data rows top+1..top+3
    # H1 averages = AVERAGE of each unit/year row in the blocks above
    h1 = ctx["h1_focus_planned"]
    ws.write(r, 0, "H1 average offline (=AVERAGE Jan:Jun)", fm["secn"])
    hdr = r + 1
    ws.write(hdr, 0, "Unit", fm["hl"])
    for j, y in enumerate([FY - 2, FY - 1, FY]):
        ws.write(hdr, 1 + j, y, fm["h"])
    for i, f in enumerate(FOCUS):
        ws.write(hdr + 1 + i, 0, engine.FOCUS_LABEL[f], fm["rowh"])
        for j, y in enumerate([FY - 2, FY - 1, FY]):
            drow = blocks[f] + 1 + j        # 0-based data row for (f, year)
            ws.write_formula(hdr + 1 + i, 1 + j, f"=AVERAGE(B{drow+1}:G{drow+1})", fm["num"],
                             float(h1.loc[f, y] if y in h1.columns else 0.0))
    ws.insert_image(3, 9, assets["h1_month_by_unit"], IMG)


def _padd(wb, fm, ctx, base, assets):
    ws = wb.add_worksheet("PADD by Unit")
    ws.set_column(0, 0, 7); ws.set_column(1, 1, 7); ws.set_column(2, 2, 6)
    ws.set_column(3, 3, 5); ws.set_column(4, 4, 9); ws.set_column(5, 7, 9)
    _title(ws, fm, "Outages by PADD by Unit: full history with MoM% and YoY% = SUMIFS over Data",
           f"CDU and FCC offline by PADD and month, {Y0}-{FY}: level, month-over-month % and year-over-year %.")
    units = [("CDU", "Crude (CDU)"), ("FCC", "FCC (cat cracker)")]
    padds = [f"PADD {k}" for k in range(1, 6)]
    years = YEARS
    series = {(f, p): _ym(base, focus=f, padd=p) for f, _ in units for p in padds}
    # tidy monthly fact table: one row per unit/PADD/month, with MoM% and YoY%
    r0 = 3
    for j, h in enumerate(["Unit", "PADD", "Year", "Mon", "Period", "kbd", "MoM %", "YoY %"]):
        ws.write(r0, j, h, fm["h"] if j >= 5 else fm["hl"])
    rr = r0 + 1
    for f, _lbl in units:
        for p in padds:
            vals = [_g(series[(f, p)], y, mi) for y in years for mi in range(1, 13)]
            if sum(vals) <= 0:
                continue                   # skip unit/PADD combos with no outages
            for k in range(60):
                y, mi = years[k // 12], (k % 12) + 1
                xl = rr + 1
                ws.write(rr, 0, f, fm["txt"]); ws.write(rr, 1, p.replace("PADD ", "P"), fm["txt"])
                ws.write_number(rr, 2, y, fm["txt"]); ws.write_number(rr, 3, mi, fm["txt"])
                ws.write(rr, 4, f"{MONTHS[mi-1]} {str(y)[2:]}", fm["rowh"])
                ws.write_formula(rr, 5, _si("K", ("A", y), ("B", mi), ("G", f), ("F", p)), fm["num"], vals[k])
                if k >= 1:
                    pv = vals[k - 1]
                    ws.write_formula(rr, 6, f'=IFERROR(F{xl}/F{xl-1}-1,"")', fm["pct"],
                                     (vals[k] / pv - 1) if pv else "")
                if k >= 12:
                    pv = vals[k - 12]
                    ws.write_formula(rr, 7, f'=IFERROR(F{xl}/F{xl-12}-1,"")', fm["pct"],
                                     (vals[k] / pv - 1) if pv else "")
                rr += 1
    ws.autofilter(r0, 0, rr - 1, 7)
    ws.freeze_panes(r0 + 1, 0)
    ws.insert_image(r0, 9, assets["cdu_padd_27"], IMG)
    ws.insert_image(r0 + 21, 9, assets["fcc_padd_27"], IMG)
    # annual pivot per unit: PADD x year level + YoY%
    ar = rr + 2
    for f, lbl in units:
        ws.write(ar, 0, f"{lbl}: annual by PADD (kbd) and YoY%", fm["secn"])
        h = ar + 1
        ws.write(h, 0, "PADD", fm["hl"])
        for j, y in enumerate(years):
            ws.write(h, 1 + j, y, fm["h"])
        for j, y in enumerate(years[1:], 1):
            ws.write(h, len(years) + j, f"{y} YoY%", fm["h"])
        for i, p in enumerate(padds):
            rr2 = h + 1 + i; xl = rr2 + 1
            ws.write(rr2, 0, p.replace("PADD ", "P"), fm["rowh"])
            ann = {y: float(series[(f, p)].get(y, pd.Series(dtype=float)).sum()) for y in years}
            for j, y in enumerate(years):
                ws.write_formula(rr2, 1 + j, _si("K", ("A", y), ("G", f), ("F", p)), fm["num"], ann[y])
            for j, y in enumerate(years[1:], 1):
                col = chr(ord("B") + j)      # this year's level column
                pcol = chr(ord("B") + j - 1)
                ws.write_formula(rr2, len(years) + j, f'=IFERROR({col}{xl}/{pcol}{xl}-1,"")', fm["pct"],
                                 (ann[y] / ann[years[j-1]] - 1) if ann[years[j-1]] else "")
        ar = h + len(padds) + 2


def _naphtha(wb, fm, ctx, assets):
    ws = wb.add_worksheet("Naphtha")
    ws.set_column(0, 0, 8); ws.set_column(1, 6, 15)
    _title(ws, fm, "Naphtha Balance: CDU Supply vs Reformer Demand (kbd) = live model",
           "CDU/reformer offline pulled from Data; supply, demand and net recompute off the Assumptions yields.")
    nb = ctx["naphtha_balance"]
    mnum = {mo: i + 1 for i, mo in enumerate(MONTHS)}
    heads = ["Month", "CDU offline", "Naphtha supply removed", "Reformer offline",
             "Naphtha demand removed", "Net balance"]
    r0 = 3
    ws.write(r0, 0, heads[0], fm["hl"])
    for j, h in enumerate(heads[1:], 1):
        ws.write(r0, j, h, fm["h"])
    for i, mo in enumerate(MONTHS):
        rr = r0 + 1 + i
        ws.write(rr, 0, mo, fm["rowh"])
        ws.write_formula(rr, 1, _si("K", ("A", FY),("B", mnum[mo]), ("G", "CDU")),
                         fm["num"], float(nb["cdu_offline"][i]))
        ws.write_formula(rr, 2, f"=B{rr+1}*naph_yield", fm["num"], float(nb["supply_removed"][i]))
        ws.write_formula(rr, 3, _si("K", ("A", FY),("B", mnum[mo]), ("G", "Reformer")),
                         fm["num"], float(nb["ref_offline"][i]))
        ws.write_formula(rr, 4, f"=D{rr+1}*ref_intake", fm["num"], float(nb["demand_removed"][i]))
        ws.write_formula(rr, 5, f"=E{rr+1}-C{rr+1}", fm["num"], float(nb["net"][i]))
    tot = r0 + 1 + 12
    ws.write(tot, 0, "Year", fm["rowh"])
    cached = [sum(nb["cdu_offline"]), sum(nb["supply_removed"]), sum(nb["ref_offline"]),
              sum(nb["demand_removed"]), nb["annual_net"]]
    for j, col, cv in zip((1, 2, 3, 4, 5), "BCDEF", cached):
        ws.write_formula(tot, j, f"=SUM({col}{r0+2}:{col}{r0+13})", fm["num"], float(cv))
    ws.write(tot + 2, 0, "Net < 0 = deficit (naphtha short, bullish reformate); net > 0 = surplus.", fm["key"])
    ws.write(tot + 3, 0, f"{FY} annual net = {nb['annual_net']:,.0f} kbd "
                         f"({nb['n_deficit']} deficit months, {nb['n_surplus']} surplus).", fm["sub"])
    ws.insert_image(3, 7, assets["naphtha_balance"], IMG)


def _exxon(wb, fm, ctx, assets):
    ws = wb.add_worksheet("ExxonMobil")
    ws.set_column(0, 0, 24); ws.set_column(1, 1, 26); ws.set_column(2, 3, 11)
    ws.set_column(4, 4, 9); ws.set_column(5, 6, 11); ws.set_column(7, 7, 24)
    _title(ws, fm, f"ExxonMobil {FY} Focus-Unit Turnarounds, per Unit (kbd)",
           "Cross-checked against ExxonMobil's corporate turnaround plan.")
    ev = ctx["exxon_verify"]["events"]
    ev = ev[ev["focus"].isin(FOCUS)].copy().sort_values("kbd", ascending=False)
    heads = ["Refinery", "Unit", "Class", "PADD", "kbd", "Window", "Type", "Vs corporate plan"]
    r0 = 3
    for j, h in enumerate(heads):
        ws.write(r0, j, h, fm["h"] if j >= 2 else fm["hl"])
    for i, (_, row) in enumerate(ev.iterrows()):
        rr = r0 + 1 + i
        ws.write(rr, 0, str(row["plant"]).replace(" Refinery", ""), fm["txt"])
        ws.write(rr, 1, str(row["unit_name"]), fm["txt"])
        ws.write(rr, 2, str(row["focus"]), fm["txt"])
        ws.write(rr, 3, str(row["padd"]), fm["txt"])
        ws.write_number(rr, 4, float(row["kbd"]), fm["num"])
        ws.write(rr, 5, str(row["span"]), fm["txt"])
        ws.write(rr, 6, str(row["type"]).title(), fm["txt"])
        ws.write(rr, 7, str(row.get("note", "")), fm["txt"])
    ws.insert_image(3, 9, assets["exxon_gantt"], IMG)


def _forecast(wb, fm, ctx, assets):
    ws = wb.add_worksheet("Forecast")
    ws.set_column(0, 0, 26); ws.set_column(1, 12, 8); ws.set_column(13, 15, 11)
    _title(ws, fm, f"{FY} Unplanned Scenario (kbd/month) = baseline x multiplier (live)",
           "Headline by month (peak/avg = real levels). Annual Sigma (sum of the 12 monthly "
           "concurrent figures) is kept for reference -- a magnitude, not a level; not for slides.")
    dff = ctx["df"]
    base = engine.baseline_profile(ctx["df"], ctx["scenario"]["window"])
    fan = ctx["scenario_fan"]
    planned_m = [float(dff[(dff["year"] == FY) & (dff["type"] == "PLANNED") & (dff["month"] == m)]["cap_kbd"].sum())
                 for m in range(1, 13)]
    r0 = 3
    # --- monthly scenario paths (unplanned), peak & avg per month (levels, not sums) ---
    ws.write(r0, 0, "Unplanned scenario (kbd/mo)", fm["hl"])
    for j, mo in enumerate(MONTHS):
        ws.write(r0, 1 + j, mo, fm["h"])
    ws.write(r0, 13, "Peak/mo", fm["h"]); ws.write(r0, 14, "Avg/mo", fm["h"])
    ws.write(r0, 15, "Annual Σ", fm["h"])
    ws.write(r0 + 1, 0, f"Baseline ({ctx['scenario']['window']}, completeness-aware)", fm["rowh"])
    for j, mo in enumerate(MONTHS):
        ws.write_number(r0 + 1, 1 + j, float(base[mo]), fm["num"])
    ws.write_formula(r0 + 1, 13, f"=MAX(B{r0+2}:M{r0+2})", fm["num"], float(base.max()))
    ws.write_formula(r0 + 1, 14, f"=AVERAGE(B{r0+2}:M{r0+2})", fm["num"], float(base.mean()))
    ws.write_formula(r0 + 1, 15, f"=SUM(B{r0+2}:M{r0+2})", fm["num"], float(base.sum()))
    for k, ref in enumerate(["mult_cons", "mult_avg", "mult_act"]):
        nm = ["Conservative", "Average", "Active"][k]
        rr = r0 + 2 + k
        ws.write(rr, 0, nm, fm["rowh"])
        for j in range(12):
            col = chr(ord("B") + j)
            ws.write_formula(rr, 1 + j, f"={col}{r0+2}*{ref}", fm["num"], float(fan[nm].iloc[j]))
        ws.write_formula(rr, 13, f"=MAX(B{rr+1}:M{rr+1})", fm["num"], float(fan[nm].max()))
        ws.write_formula(rr, 14, f"=AVERAGE(B{rr+1}:M{rr+1})", fm["num"], float(fan[nm].mean()))
        ws.write_formula(rr, 15, f"=SUM(B{rr+1}:M{rr+1})", fm["num"], float(fan[nm].sum()))
    # --- booked planned by month (so implied is a real monthly level) ---
    pr = r0 + 5                                            # planned row (Excel row pr+1)
    ws.write(pr, 0, f"Booked planned {FY} (by month)", fm["rowh"])
    for j in range(12):
        ws.write_formula(pr, 1 + j, _si("K", ("A", FY), ("B", j + 1), ("J", "PLANNED")),
                         fm["num"], planned_m[j])
    ws.write_formula(pr, 13, f"=MAX(B{pr+1}:M{pr+1})", fm["num"], max(planned_m))
    ws.write_formula(pr, 14, f"=AVERAGE(B{pr+1}:M{pr+1})", fm["num"], sum(planned_m) / 12)
    ws.write_formula(pr, 15, f"=SUM(B{pr+1}:M{pr+1})", fm["num"], sum(planned_m))
    # --- implied total offline BY MONTH = scenario unplanned + booked planned ---
    ir = r0 + 7
    ws.write(ir, 0, "Implied total offline by month = scenario unplanned + booked planned", fm["secn"])
    ws.write(ir + 1, 0, "Scenario", fm["hl"])
    for j, mo in enumerate(MONTHS):
        ws.write(ir + 1, 1 + j, mo, fm["h"])
    ws.write(ir + 1, 13, "Peak month", fm["h"]); ws.write(ir + 1, 14, "Avg month", fm["h"])
    ws.write(ir + 1, 15, "Annual Σ", fm["h"])
    for k, nm in enumerate(["Conservative", "Average", "Active"]):
        rr = ir + 2 + k
        srow = r0 + 3 + k                                  # this scenario's path row (Excel)
        impl = [float(fan[nm].iloc[j]) + planned_m[j] for j in range(12)]
        ws.write(rr, 0, nm, fm["rowh"])
        for j in range(12):
            col = chr(ord("B") + j)
            ws.write_formula(rr, 1 + j, f"={col}{srow}+{col}{pr+1}", fm["num"], impl[j])
        ws.write_formula(rr, 13, f"=MAX(B{rr+1}:M{rr+1})", fm["num"], max(impl))
        ws.write_formula(rr, 14, f"=AVERAGE(B{rr+1}:M{rr+1})", fm["num"], sum(impl) / 12)
        ws.write_formula(rr, 15, f"=SUM(B{rr+1}:M{rr+1})", fm["num"], sum(impl))
    ws.write(ir + 5, 0, "Peak month = worst single month offline (trade this). Avg month = average. "
             "Annual Σ = sum of the 12 monthly figures: a reference magnitude, not a level -- keep it off slides.",
             fm["sub"])
    sb = ctx["scenario_bands"]
    hb = ir + 7
    ws.write(hb, 0, "History (complete years, annual Σ unplanned)", fm["secn"])
    for k, (lab, key) in enumerate([("P25", "p25"), ("Median", "p50"), ("P90", "p90"), ("Mean", "mean")]):
        ws.write(hb + 1 + k, 0, lab, fm["lbl"]); ws.write_number(hb + 1 + k, 1, float(sb[key]), fm["num"])
    ws.insert_image(3, 17, assets["fan"], IMG)


def _base(df):
    """The Snowflake rows the Data sheet holds and every aggregate sums over.
    No dedup: summing CAP_OFFLINE_ADJUSTED_KBD is exactly what the live SUMIFS do,
    so the workbook keeps working (and agreeing with the deck) as the file grows."""
    d = df.copy()
    d["focus"] = d["focus"].fillna("")
    d["padd"] = d["padd"].fillna("")
    d["pad_dist"] = d["pad_dist"].astype(str).replace({"nan": "", "None": "", "NaN": ""})
    d = d.dropna(subset=["year", "month"]).copy()
    d["year"] = d["year"].astype(int); d["month"] = d["month"].astype(int)
    return d


def _focus_formula(r):
    """Excel formula mapping UNIT_CATEGORY (col H) -> focus class, so pasted rows
    classify themselves (keeps the model live). CDU excludes vacuum pipe stills the
    golden record mislabels as ATMOS DISTILLATION: a unit whose name (col I) has
    'VPS' or starts with 'VACUUM' is vacuum, not crude. Mirrors engine
    _focus_from_unitcat so the live model and the deck agree."""
    return (f'=IF(AND($H{r}="ATMOS DISTILLATION",ISNUMBER(SEARCH("VPS",$I{r}))=FALSE,'
            f'LEFT($I{r},6)<>"VACUUM"),"CDU",IF($H{r}="FLUID CAT CRACKING","FCC",'
            f'IF(OR($H{r}="HYDROCRACKING",$H{r}="RESID_HYDROCRACKING"),"Hydrocracker",'
            f'IF($H{r}="REFORMING","Reformer",""))))')


def _padd_formula(r):
    """Excel formula mapping PAD_DIST (col M) Roman -> canonical 'PADD n'."""
    return (f'=IF($M{r}="PADD I","PADD 1",IF($M{r}="PADD II","PADD 2",IF($M{r}="PADD III","PADD 3",'
            f'IF($M{r}="PADD IV","PADD 4",IF($M{r}="PADD V","PADD 5","")))))')


def _ym(base, **f):
    """year/month Series of summed cap_kbd for the given equality filters."""
    d = base
    for k, v in f.items():
        d = d[d[k] == v]
    return d.groupby(["year", "month"])["cap_kbd"].sum()


def _g(s, y, m):
    return float(s.get((y, m), 0.0))


def _historicals(wb, fm, base):
    ws = wb.add_worksheet("Historicals")
    ws.set_column(0, 1, 6); ws.set_column(2, 2, 9); ws.set_column(3, 16, 8)
    _title(ws, fm, f"Historicals: monthly offline {Y0}-{FY} (kbd) = SUMIFS over Data",
           "Total / planned / unplanned, unplanned %, per focus unit and per PADD. The series the stats use.")
    tot, pl, un = _ym(base), _ym(base, type="PLANNED"), _ym(base, type="UNPLANNED")
    cdu, fcc = _ym(base, focus="CDU"), _ym(base, focus="FCC")
    hdc, refm = _ym(base, focus="Hydrocracker"), _ym(base, focus="Reformer")
    padd = {k: _ym(base, padd=f"PADD {k}") for k in range(1, 6)}
    heads = ["Year", "Mon", "Period", "Total", "Planned", "Unplanned", "Unpl %",
             "CDU", "FCC", "HydroCk", "Reformer", "PADD1", "PADD2", "PADD3", "PADD4", "PADD5", "Naph net"]
    r0 = 3
    for j, h in enumerate(heads):
        ws.write(r0, j, h, fm["h"] if j >= 3 else fm["hl"])
    years = YEARS
    i = 0
    for y in years:
        for mi, mo in enumerate(MONTHS, 1):
            rr = r0 + 1 + i
            xl = rr + 1
            ws.write_number(rr, 0, y, fm["txt"]); ws.write_number(rr, 1, mi, fm["txt"])
            ws.write(rr, 2, f"{mo} {str(y)[2:]}", fm["rowh"])
            ws.write_formula(rr, 3, _si("K", ("A", y), ("B", mi)), fm["num"], _g(tot, y, mi))
            ws.write_formula(rr, 4, _si("K", ("A", y), ("B", mi), ("J", "PLANNED")), fm["num"], _g(pl, y, mi))
            ws.write_formula(rr, 5, _si("K", ("A", y), ("B", mi), ("J", "UNPLANNED")), fm["num"], _g(un, y, mi))
            ws.write_formula(rr, 6, f"=IFERROR(F{xl}/D{xl},0)", fm["f2"],
                             (_g(un, y, mi) / _g(tot, y, mi)) if _g(tot, y, mi) else 0.0)
            for c, (s, foc) in enumerate([(cdu, "CDU"), (fcc, "FCC"), (hdc, "Hydrocracker"), (refm, "Reformer")]):
                ws.write_formula(rr, 7 + c, _si("K", ("A", y), ("B", mi), ("G", foc)), fm["num"], _g(s, y, mi))
            for c, k in enumerate(range(1, 6)):
                ws.write_formula(rr, 11 + c, _si("K", ("A", y), ("B", mi), ("F", f"PADD {k}")),
                                 fm["num"], _g(padd[k], y, mi))
            ws.write_formula(rr, 16, f"=K{xl}*ref_intake-H{xl}*naph_yield", fm["num"],
                             engine.REFORMER_NAPHTHA_INTAKE * _g(refm, y, mi) - engine.NAPHTHA_YIELD * _g(cdu, y, mi))
            i += 1
    ws.freeze_panes(r0 + 1, 3)
    ws.autofilter(r0, 0, r0 + 60, len(heads) - 1)
    # annual summary with YoY%
    ta = tot.groupby(level=0).sum(); pa = pl.groupby(level=0).sum(); ua = un.groupby(level=0).sum()
    a0 = r0 + 62
    ws.write(a0, 0, "Annual (kbd) and YoY%", fm["secn"])
    ah = a0 + 1
    for j, h in enumerate(["Year", "Total", "Planned", "Unplanned", "Unpl %", "Total YoY%"]):
        ws.write(ah, j, h, fm["h"] if j else fm["hl"])
    for k, y in enumerate(years):
        rr = ah + 1 + k; xl = rr + 1
        tv, uv = float(ta.get(y, 0.0)), float(ua.get(y, 0.0))
        ws.write_number(rr, 0, y, fm["rowh"])
        ws.write_formula(rr, 1, _si("K", ("A", y)), fm["num"], tv)
        ws.write_formula(rr, 2, _si("K", ("A", y), ("J", "PLANNED")), fm["num"], float(pa.get(y, 0.0)))
        ws.write_formula(rr, 3, _si("K", ("A", y), ("J", "UNPLANNED")), fm["num"], uv)
        ws.write_formula(rr, 4, f"=IFERROR(D{xl}/B{xl},0)", fm["f2"], (uv / tv) if tv else 0.0)
        if k == 0:
            ws.write(rr, 5, "n/a", fm["txt"])
        else:
            prev = float(ta.get(years[k - 1], 0.0))
            ws.write_formula(rr, 5, f"=IFERROR(B{xl}/B{xl-1}-1,0)", fm["pct"],
                             (tv / prev - 1) if prev else 0.0)
    # live line chart of total/planned/unplanned by month
    ch = wb.add_chart({"type": "line"})
    for col, nm, color in [(3, "Total", "#1F3864"), (4, "Planned", "#BF9000"), (5, "Unplanned", "#C00000")]:
        ch.add_series({"name": nm, "categories": ["Historicals", r0 + 1, 2, r0 + 60, 2],
                       "values": ["Historicals", r0 + 1, col, r0 + 60, col],
                       "line": {"color": color, "width": 1.75}})
    ch.set_title({"name": "US Capacity Offline by Month (kbd)"})
    ch.set_size({"width": 760, "height": 320})
    ch.set_legend({"position": "bottom"})
    ws.insert_chart(r0, 18, ch)
    return r0  # header row (data rows r0+1 .. r0+60)


# stats window: 2023-2025 = 36 fully-reported months (rows 5..40 in Historicals, 1-based)
STAT_HEADS = [("Total", "D"), ("Planned", "E"), ("Unplanned", "F"), ("CDU", "H"),
              ("FCC", "I"), ("HydroCk", "J"), ("Reformer", "K"), ("Naph net", "Q")]
STAT_R1, STAT_R2 = 5, 40       # Historicals 1-based data rows for 2023-2025


def _stat_series(base):
    keys = [(y, m) for y in (2023, 2024, 2025) for m in range(1, 13)]
    s = {"Total": _ym(base), "Planned": _ym(base, type="PLANNED"), "Unplanned": _ym(base, type="UNPLANNED"),
         "CDU": _ym(base, focus="CDU"), "FCC": _ym(base, focus="FCC"),
         "HydroCk": _ym(base, focus="Hydrocracker"), "Reformer": _ym(base, focus="Reformer")}
    out = {nm: np.array([_g(s[nm], y, m) for (y, m) in keys]) for nm in s}
    out["Naph net"] = engine.REFORMER_NAPHTHA_INTAKE * out["Reformer"] - engine.NAPHTHA_YIELD * out["CDU"]
    return out


def _statistics(wb, fm, base):
    ws = wb.add_worksheet("Statistics")
    ws.set_column(0, 0, 16); ws.set_column(1, 9, 11)
    _title(ws, fm, "Statistics (2023-2025 monthly, n=36)",
           "Descriptive stats and a correlation matrix over the Historicals series. Live =formulas.")
    arr = _stat_series(base)
    names = [n for n, _ in STAT_HEADS]
    # descriptive
    ws.write(3, 0, "Descriptive (kbd)", fm["secn"])
    rows = [("Mean", "AVERAGE", lambda a: np.mean(a)),
            ("Median", "MEDIAN", lambda a: np.median(a)),
            ("Std dev", "STDEV", lambda a: np.std(a, ddof=1)),
            ("Min", "MIN", lambda a: np.min(a)),
            ("P25", None, lambda a: np.percentile(a, 25)),
            ("P75", None, lambda a: np.percentile(a, 75)),
            ("Max", "MAX", lambda a: np.max(a))]
    ws.write(4, 0, "Statistic", fm["hl"])
    for j, (nm, col) in enumerate(STAT_HEADS):
        ws.write(4, 1 + j, nm, fm["h"])
    for i, (lab, fn, pyf) in enumerate(rows):
        ws.write(5 + i, 0, lab, fm["rowh"])
        for j, (nm, col) in enumerate(STAT_HEADS):
            rng = f"Historicals!${col}${STAT_R1}:${col}${STAT_R2}"
            if fn:
                form = f"={fn}({rng})"
            else:
                pct = 0.25 if lab == "P25" else 0.75
                form = f"=PERCENTILE({rng},{pct})"
            ws.write_formula(5 + i, 1 + j, form, fm["num"], float(pyf(arr[nm])))
    # correlation matrix
    cr = 5 + len(rows) + 2
    ws.write(cr, 0, "Correlation matrix (Pearson)", fm["secn"])
    ws.write(cr + 1, 0, "", fm["hl"])
    for j, nm in enumerate(names):
        ws.write(cr + 1, 1 + j, nm, fm["h"])
    M = np.corrcoef([arr[n] for n in names])
    for i, ni in enumerate(names):
        ws.write(cr + 2 + i, 0, ni, fm["rowh"])
        for j, nj in enumerate(names):
            ci, cj = STAT_HEADS[i][1], STAT_HEADS[j][1]
            form = (f"=CORREL(Historicals!${ci}${STAT_R1}:${ci}${STAT_R2},"
                    f"Historicals!${cj}${STAT_R1}:${cj}${STAT_R2})")
            ws.write_formula(cr + 2 + i, 1 + j, form, fm["f2"], float(M[i, j]))
    ws.conditional_format(cr + 2, 1, cr + 1 + len(names), len(names),
                          {"type": "3_color_scale", "min_color": "#C00000",
                           "mid_color": "#FFFFFF", "max_color": "#2E5496",
                           "min_value": -1, "mid_value": 0, "max_value": 1,
                           "min_type": "num", "mid_type": "num", "max_type": "num"})
    ws.write(cr + 3 + len(names), 0,
             "+1 move together, -1 move opposite. Crude vs reformer turnaround clustering, "
             "planned vs unplanned substitution, etc.", fm["key"])


def _reg_block(wb, ws, fm, r, title, xlab, ylab, xname, yname, x, y):
    """Write one regression block (slope/intercept/R2/SE/corr/n) + a scatter+trendline."""
    xc, yc = STAT_HEADS_MAP[xname], STAT_HEADS_MAP[yname]
    rng = lambda c: f"Historicals!${c}${STAT_R1}:${c}${STAT_R2}"
    slope, intercept = np.polyfit(x, y, 1)
    r2 = float(np.corrcoef(x, y)[0, 1] ** 2)
    resid = y - (slope * x + intercept)
    steyx = float(np.sqrt(np.sum(resid ** 2) / (len(x) - 2)))
    ws.write(r, 0, title, fm["secn"])
    items = [("Slope", f"=SLOPE({rng(yc)},{rng(xc)})", slope),
             ("Intercept", f"=INTERCEPT({rng(yc)},{rng(xc)})", intercept),
             ("R-squared", f"=RSQ({rng(yc)},{rng(xc)})", r2),
             ("Correlation", f"=CORREL({rng(yc)},{rng(xc)})", float(np.corrcoef(x, y)[0, 1])),
             ("Std err (STEYX)", f"=STEYX({rng(yc)},{rng(xc)})", steyx),
             ("n", f"=COUNT({rng(xc)})", float(len(x)))]
    for i, (lab, form, val) in enumerate(items):
        ws.write(r + 1 + i, 0, lab, fm["rowh"])
        ws.write_formula(r + 1 + i, 1, form, fm["f2"], val)
    ch = wb.add_chart({"type": "scatter"})
    ch.add_series({"name": f"{yname} vs {xname}",
                   "categories": rng(xc), "values": rng(yc),
                   "marker": {"type": "circle", "size": 5, "fill": {"color": "#2E5496"}},
                   "trendline": {"type": "linear", "display_equation": True,
                                 "display_r_squared": True, "line": {"color": "#C00000", "width": 1.5}}})
    ch.set_title({"name": f"{ylab} vs {xlab}"})
    ch.set_x_axis({"name": xlab}); ch.set_y_axis({"name": ylab})
    ch.set_legend({"none": True}); ch.set_size({"width": 430, "height": 300})
    ws.insert_chart(r, 3, ch)
    return r + 9


STAT_HEADS_MAP = {n: c for n, c in STAT_HEADS}


def _regression(wb, fm, base):
    ws = wb.add_worksheet("Regression")
    ws.set_column(0, 0, 18); ws.set_column(1, 1, 12)
    _title(ws, fm, "Regression / best-fit (2023-2025 monthly, n=36)",
           "Least-squares fits with R-squared; the chart trendline is the best-fit line.")
    arr = _stat_series(base)
    r = 3
    r = _reg_block(wb, ws, fm, r, "1) Reformer offline vs CDU offline (turnaround clustering)",
                   "CDU offline (kbd)", "Reformer offline (kbd)", "CDU", "Reformer",
                   arr["CDU"], arr["Reformer"]) + 2
    r = _reg_block(wb, ws, fm, r, "2) Unplanned vs Planned offline (substitution?)",
                   "Planned offline (kbd)", "Unplanned offline (kbd)", "Planned", "Unplanned",
                   arr["Planned"], arr["Unplanned"]) + 2
    _reg_block(wb, ws, fm, r, "3) Naphtha net vs CDU offline (supply mechanic)",
               "CDU offline (kbd)", "Naphtha net (kbd)", "CDU", "Naph net",
               arr["CDU"], arr["Naph net"])


def _sensitivity(wb, fm, ctx):
    ws = wb.add_worksheet("Sensitivity")
    ws.set_column(0, 0, 22); ws.set_column(1, 7, 12)
    _title(ws, fm, f"Sensitivity: {FY} peak-month implied offline (kbd)",
           "Worst single month of total capacity offline. Rows = unplanned multiplier x baseline; "
           "columns = one-off shock added to that month. Live off the Forecast monthly cells.")
    mults = [0.7, 0.85, 1.0, 1.15, 1.3, 1.5]
    shocks = [0, 250, 500, 750, 1000]
    base_rng, plan_rng = "Forecast!$B$5:$M$5", "Forecast!$B$9:$M$9"   # monthly baseline & planned
    base_m = engine.baseline_profile(ctx["df"], ctx["scenario"]["window"]).values
    dff = ctx["df"]
    plan_m = [float(dff[(dff["year"] == FY) & (dff["type"] == "PLANNED") & (dff["month"] == m)]["cap_kbd"].sum())
              for m in range(1, 13)]
    r0 = 3                                   # header row (Excel r0+1); shock values are editable here
    hxl = r0 + 1
    ws.write(r0, 0, "mult v / shock >", fm["hl"])
    for j, sh in enumerate(shocks):
        ws.write_number(r0, 1 + j, sh, fm["inph"])
    for i, mu in enumerate(mults):
        rr = r0 + 1 + i; xl = rr + 1
        ws.write_number(rr, 0, mu, fm["inp"])
        peak = max(mu * base_m[m] + plan_m[m] for m in range(12))   # worst combined month
        for j in range(len(shocks)):
            col = chr(ord("B") + j)
            form = f"{{=MAX($A{xl}*{base_rng}+{plan_rng})+{col}${hxl}}}"
            ws.write_array_formula(rr, 1 + j, rr, 1 + j, form, fm["num"], peak + shocks[j])
    ws.conditional_format(r0 + 1, 1, r0 + len(mults), len(shocks),
                          {"type": "3_color_scale", "min_color": "#63BE7B",
                           "mid_color": "#FFEB84", "max_color": "#F8696B"})
    ws.write(r0 + len(mults) + 2, 0, "Editable: shock row and multiplier column are gold. Peak-month implied = "
             "worst month of (baseline x multiplier + booked planned) + one-off shock. Base = mult 1.0, shock 0.",
             fm["sub"])


def _stress(wb, fm, ctx):
    ws = wb.add_worksheet("Stress Test")
    ws.set_column(0, 0, 26); ws.set_column(1, 1, 46); ws.set_column(2, 5, 14)
    _title(ws, fm, f"Stress test: named shocks on the {FY} book (peak-month kbd)",
           "Edit the gold shock cells. Peak-month implied offline and naphtha impact recompute live.")
    avg_peak = "Forecast!$N$14"        # Average scenario implied, peak month (a level)
    dff = ctx["df"]
    base_m = engine.baseline_profile(ctx["df"], ctx["scenario"]["window"]).values
    plan_m = [float(dff[(dff["year"] == FY) & (dff["type"] == "PLANNED") & (dff["month"] == m)]["cap_kbd"].sum())
              for m in range(1, 13)]
    base_peak = max(base_m[m] + plan_m[m] for m in range(12))   # Average implied peak month
    ny = engine.NAPHTHA_YIELD
    rows = [
        ("USGC hurricane", "PADD 3 unplanned spike, ~1-2 months", 800, 0),
        ("Winter freeze (Uri-like)", "National unplanned spike in Feb", 0, 2000),
        ("Two large CDUs trip", "~500 kbd crude unplanned, 1 month", 500, 0),
        ("Heavy fall TA overlap", "Sep-Oct planned overlap adds load", 0, 600),
        ("Quiet year", "Light unplanned, fewer surprises", 0, -1500),
    ]
    r0 = 3
    for j, h in enumerate(["Scenario", "Description", "Crude shock (kbd)", "Other shock (kbd)",
                           "Peak-month implied", "Naphtha net impact"]):
        ws.write(r0, j, h, fm["h"] if j >= 2 else fm["hl"])
    for i, (nm, desc, crude, other) in enumerate(rows):
        rr = r0 + 1 + i; xl = rr + 1
        ws.write(rr, 0, nm, fm["rowh"]); ws.write(rr, 1, desc, fm["txt"])
        ws.write_number(rr, 2, crude, fm["inph"]); ws.write_number(rr, 3, other, fm["inph"])
        ws.write_formula(rr, 4, f"={avg_peak}+C{xl}+D{xl}", fm["num"], base_peak + crude + other)
        ws.write_formula(rr, 5, f"=-naph_yield*C{xl}", fm["num"], -ny * crude)
    ws.write(r0 + len(rows) + 2, 0, "Peak-month implied = Average scenario's worst-month offline + crude shock "
             "+ other shock. Naphtha impact = -naphtha_yield x crude shock (more crude down = shorter naphtha).",
             fm["sub"])


# Colored tabs grouped by purpose: navigation, source, per-unit analysis, balance, forecast, stats
def _whats_changed(wb, fm, ctx):
    """Rolling change tracker. Month-over-month is live off Data (the balance that
    prices contracts); the rolling pull-log compares weekly pulls (the source is
    monthly, so true week-over-week = pull-over-pull, accumulated each build)."""
    import datetime as _dt
    ws = wb.add_worksheet("What's Changed")
    ws.set_column(0, 0, 26); ws.set_column(1, 6, 12)
    asof = _dt.date.today()
    cy, cm = asof.year, asof.month
    py, pm = (cy, cm - 1) if cm > 1 else (cy - 1, 12)
    df = ctx["df"]
    _title(ws, fm, f"What's Changed: rolling MoM & WoW (as of {asof:%b %d, %Y})",
           "Month-over-month is live off Data -- the balance that prices. Week-over-week compares your "
           "weekly pulls (source is monthly, so the rolling log grows one row per build).")

    COL_XL = {"type": "J", "focus": "G", "padd": "F"}    # pandas col -> Data-sheet letter

    def msum(y, m, *crit):                               # crit = (pandas_col, value)
        d = df[(df["year"] == y) & (df["month"] == m)]
        for col, val in crit:
            d = d[d[col] == val]
        return float(d["cap_kbd"].sum())

    def si_cell(ycell, mcell, *crit):                    # crit = (pandas_col, value) -> SUMIFS
        parts = ["Data!$K:$K", f"Data!$A:$A,{ycell}", f"Data!$B:$B,{mcell}"]
        for col, val in crit:
            parts.append(f'Data!${COL_XL[col]}:${COL_XL[col]},"{val}"')
        return "=SUMIFS(" + ",".join(parts) + ")"

    # ---- 1) Month-over-month (live) ----
    ws.write(3, 0, "Month-over-month (live off Data -- this is the balance that prices)", fm["secn"])
    ws.write(4, 0, "As-of month (edit year / month)", fm["lbl"])
    ws.write_number(4, 1, cy, fm["inp"]); ws.write_number(4, 2, cm, fm["inp"])
    ws.write(5, 0, "Prior month", fm["lbl"])
    ws.write_formula(5, 1, "=IF(C5=1,B5-1,B5)", fm["txt"], py)
    ws.write_formula(5, 2, "=IF(C5=1,12,C5-1)", fm["txt"], pm)
    hdr = 7
    for j, h in enumerate(["Metric", "This month", "Prior month", "Δ", "Δ%"]):
        ws.write(hdr, j, h, fm["hl"] if j == 0 else fm["h"])
    metrics = [("Total offline", ()), ("Planned", (("type", "PLANNED"),)), ("Unplanned", (("type", "UNPLANNED"),)),
               ("CDU", (("focus", "CDU"),)), ("FCC", (("focus", "FCC"),)),
               ("Hydrocracker", (("focus", "Hydrocracker"),)), ("Reformer", (("focus", "Reformer"),))]
    for i, (label, crit) in enumerate(metrics):
        rr = hdr + 1 + i; xl = rr + 1
        ws.write(rr, 0, label, fm["rowh"])
        ws.write_formula(rr, 1, si_cell("$B$5", "$C$5", *crit), fm["num"], msum(cy, cm, *crit))
        ws.write_formula(rr, 2, si_cell("$B$6", "$C$6", *crit), fm["num"], msum(py, pm, *crit))
        ws.write_formula(rr, 3, f"=B{xl}-C{xl}", fm["num"], msum(cy, cm, *crit) - msum(py, pm, *crit))
        ws.write_formula(rr, 4, f'=IFERROR(D{xl}/C{xl},"")', fm["pct"],
                         (msum(cy, cm, *crit) / msum(py, pm, *crit) - 1) if msum(py, pm, *crit) else "")

    # ---- 2) Rolling trailing 6 months (live; window rolls each build) ----
    seq = []
    yy, mm = cy, cm
    for _ in range(6):
        seq.append((yy, mm)); mm, yy = (mm - 1, yy) if mm > 1 else (12, yy - 1)
    seq = list(reversed(seq))
    r2 = hdr + len(metrics) + 2
    ws.write(r2, 0, "Rolling: trailing 6 months (live)", fm["secn"])
    for j, h in enumerate(["Month", "Total", "Planned", "Unplanned"]):
        ws.write(r2 + 1, j, h, fm["hl"] if j == 0 else fm["h"])
    for i, (yy, mm) in enumerate(seq):
        rr = r2 + 2 + i
        ws.write(rr, 0, f"{MONTHS[mm-1]} {yy}", fm["rowh"])
        ws.write_formula(rr, 1, _si("K", ("A", yy), ("B", mm)), fm["num"], msum(yy, mm))
        ws.write_formula(rr, 2, _si("K", ("A", yy), ("B", mm), ("J", "PLANNED")), fm["num"], msum(yy, mm, ("type", "PLANNED")))
        ws.write_formula(rr, 3, _si("K", ("A", yy), ("B", mm), ("J", "UNPLANNED")), fm["num"], msum(yy, mm, ("type", "UNPLANNED")))

    # ---- 3) Week-over-week: pull log (accumulates one row per build) ----
    log = engine.update_snapshot_log(ctx["df"])
    curkey = f"{cy}-{cm:02d}"; nxtkey = f"{cy if cm < 12 else cy+1}-{(cm % 12)+1:02d}"
    r3 = r2 + 9
    ws.write(r3, 0, "Week-over-week: your weekly pulls (rolling; one row per build)", fm["secn"])
    for j, h in enumerate(["Pull date", f"{MONTHS[cm-1]} {cy}", "Δ vs prior pull", f"{nxtkey} next mo"]):
        ws.write(r3 + 1, j, h, fm["hl"] if j == 0 else fm["h"])
    tail = log.tail(6).reset_index(drop=True)
    prev = None
    for i, row in tail.iterrows():
        rr = r3 + 2 + i
        cur_val = float(row[curkey]) if curkey in tail.columns and pd.notna(row.get(curkey)) else 0.0
        nxt_val = float(row[nxtkey]) if nxtkey in tail.columns and pd.notna(row.get(nxtkey)) else 0.0
        ws.write(rr, 0, str(row["as_of"]), fm["txt"])
        ws.write_number(rr, 1, round(cur_val, 1), fm["num"])
        ws.write(rr, 2, "" if prev is None else round(cur_val - prev, 1), fm["num"] if prev is not None else fm["txt"])
        ws.write_number(rr, 3, round(nxt_val, 1), fm["num"])
        prev = cur_val
    note_r = r3 + 2 + len(tail)
    ws.merge_range(note_r, 0, note_r, 4, "Each build appends today's snapshot (current + next 6 months' offline) to "
                   "data/whatschanged_log.csv. Run it weekly after refreshing the Snowflake and the Δ shows what the "
                   "vendor added/pulled since last week. Δ on the current month is the contract-relevant move.", fm["txtw"])

    # ---- 4) This month's movers ----
    def ids(y, m):
        d = df[(df["year"] == y) & (df["month"] == m)].dropna(subset=["outage_id"])
        if d.empty:
            return pd.DataFrame(columns=["kbd", "plant", "unit", "padd"])
        return d.groupby("outage_id").agg(kbd=("cap_kbd", "sum"), plant=("plant", "first"),
                                          unit=("unit_cat", "first"), padd=("padd", "first"))
    ci, pi = ids(cy, cm), ids(py, pm)
    new = ci.loc[ci.index.difference(pi.index)].sort_values("kbd", ascending=False).head(6)
    gone = pi.loc[pi.index.difference(ci.index)].sort_values("kbd", ascending=False).head(6)
    r4 = note_r + 2
    ws.write(r4, 0, f"New this month ({MONTHS[cm-1]} {cy}) -- not down the prior month", fm["secn"])
    for j, h in enumerate(["Plant", "Unit", "PADD", "kbd"]):
        ws.write(r4 + 1, j, h, fm["hl"] if j < 2 else fm["h"])
    for i, (_, r) in enumerate(new.iterrows()):
        rr = r4 + 2 + i
        ws.write(rr, 0, str(r["plant"])[:24], fm["txt"]); ws.write(rr, 1, str(r["unit"])[:18], fm["txt"])
        ws.write(rr, 2, str(r["padd"]), fm["txt"]); ws.write_number(rr, 3, round(r["kbd"], 1), fm["num"])
    r5 = r4 + 3 + len(new)
    ws.write(r5, 0, f"Back online this month -- down the prior month, not now", fm["secn"])
    for j, h in enumerate(["Plant", "Unit", "PADD", "kbd"]):
        ws.write(r5 + 1, j, h, fm["hl"] if j < 2 else fm["h"])
    for i, (_, r) in enumerate(gone.iterrows()):
        rr = r5 + 2 + i
        ws.write(rr, 0, str(r["plant"])[:24], fm["txt"]); ws.write(rr, 1, str(r["unit"])[:18], fm["txt"])
        ws.write(rr, 2, str(r["padd"]), fm["txt"]); ws.write_number(rr, 3, round(r["kbd"], 1), fm["num"])


def _padd_connectivity(wb, fm, ctx):
    """Sensitize the crude-outage read to regional pipeline connectivity. A CDU
    outage in a well-connected PADD (P3 Gulf) lets the downstream units keep running
    on piped-in intermediates (low pass-through); an islanded PADD has to cut them
    too (high pass-through). Effective = nominal x pass-through, live off Assumptions."""
    ws = wb.add_worksheet("PADD Connectivity")
    ws.set_column(0, 0, 12); ws.set_column(1, 5, 15)
    _title(ws, fm, "PADD Connectivity: effective crude-outage impact (live)",
           "Effective = nominal CDU offline x pass-through (Assumptions). Low pass-through (P3 Gulf) = "
           "well-connected, downstream keeps running; high = islanded, the outage cascades.")
    fp = ctx["focus_padd"]
    PADDS = ["PADD 1", "PADD 2", "PADD 3", "PADD 4", "PADD 5"]

    def nom(year, padd):
        g = fp.get(year, {}).get("CDU")
        return float(g.loc[padd].sum()) if (g is not None and padd in g.index) else 0.0
    r0 = 3
    heads = ["PADD", "Pass-thru", f"{CY} CDU (Σ kbd)", f"{CY} effective", f"{FY} CDU (Σ kbd)", f"{FY} effective"]
    for j, h in enumerate(heads):
        ws.write(r0, j, h, fm["hl"] if j == 0 else fm["h"])
    first = r0 + 1
    for i, p in enumerate(PADDS):
        rr = first + i; xl = rr + 1; n = i + 1
        f = engine.PADD_CONNECTIVITY[p]
        ws.write(rr, 0, p, fm["rowh"])
        ws.write_formula(rr, 1, f"=conn_p{n}", fm["f2"], f)
        ws.write_formula(rr, 2, _si("K", ("A", CY), ("G", "CDU"), ("F", p)), fm["num"], nom(CY, p))
        ws.write_formula(rr, 3, f"=C{xl}*conn_p{n}", fm["num"], nom(CY, p) * f)
        ws.write_formula(rr, 4, _si("K", ("A", FY), ("G", "CDU"), ("F", p)), fm["num"], nom(FY, p))
        ws.write_formula(rr, 5, f"=E{xl}*conn_p{n}", fm["num"], nom(FY, p) * f)
    tot = first + 5; txl = tot + 1
    ws.write(tot, 0, "Total", fm["rowh"]); ws.write(tot, 1, "", fm["txt"])
    cyn = sum(nom(CY, p) for p in PADDS); cye = sum(nom(CY, p) * engine.PADD_CONNECTIVITY[p] for p in PADDS)
    fyn = sum(nom(FY, p) for p in PADDS); fye = sum(nom(FY, p) * engine.PADD_CONNECTIVITY[p] for p in PADDS)
    for col, cached in [(2, cyn), (3, cye), (4, fyn), (5, fye)]:
        c = chr(ord("A") + col)
        ws.write_formula(tot, col, f"=SUM({c}{first+1}:{c}{first+5})", fm["num"], cached)
    br = tot + 2
    ws.write(br, 0, "Buffered by connectivity (nominal - effective):", fm["lbl"])
    ws.write_formula(br, 3, f"=C{txl}-D{txl}", fm["num"], cyn - cye)
    ws.write_formula(br, 5, f"=E{txl}-F{txl}", fm["num"], fyn - fye)
    ws.write_formula(br + 1, 0, f'="Effective crude tightness is "&TEXT(1-F{txl}/E{txl},"0%")&" below nominal '
                     f'in {FY} -- the P3 connectivity buffer."', fm["key"],
                     f"Effective crude tightness is {1-fye/fyn:.0%} below nominal in {FY} -- the P3 connectivity buffer."
                     if fyn else "")
    ws.write(br + 3, 0, "Tune the pass-through on the Assumptions sheet (gold cells). Applied to CDU only: "
             "crude is the unit whose outage can or cannot be routed around.", fm["key"])


def _data_quality(wb, fm, ctx):
    """Re-runnable data-quality flags: planned turnarounds recurring inside the
    ~5-year cycle (planned->planned only), and unit-months that sum to >100% of
    nameplate (overlap / double-count). Review only -- nothing is dropped."""
    ws = wb.add_worksheet("Data Quality")
    ws.set_column(0, 0, 28); ws.set_column(1, 1, 26); ws.set_column(2, 8, 10)
    _title(ws, fm, "Data Quality: turnaround-cadence & double-count flags (review only)",
           "Auto-flagged from the source each build; nothing is dropped. Verify against the source. "
           f"Turnarounds run on ~a {engine.TURNAROUND_CYCLE_YEARS}-year cycle.")
    # --- 1) turnaround cadence ---
    r0 = 3
    ws.write(r0, 0, f"Planned turnaround again within <{engine.TURNAROUND_CYCLE_YEARS}yr "
             "(CDU & FCC; planned->planned only -- unplanned never resets the cycle)", fm["secn"])
    cf_all = engine.cadence_flags(ctx["df_full"], reach_from=CY)
    cf = cf_all[(cf_all["next_TA"] >= CY) & (cf_all["next_TA"] <= FY)].sort_values(
        "kbd_next", ascending=False).head(16)
    heads = ["Plant", "Unit", "Class", "Prev TA", "Next TA", "Gap yr", "kbd prev", "kbd next"]
    hr = r0 + 1
    for j, h in enumerate(heads):
        ws.write(hr, j, h, fm["hl"] if j < 2 else fm["h"])
    for i, (_, r) in enumerate(cf.iterrows()):
        rr = hr + 1 + i
        ws.write(rr, 0, str(r["plant"])[:28], fm["txt"]); ws.write(rr, 1, str(r["unit"])[:26], fm["txt"])
        ws.write(rr, 2, r["focus"], fm["txt"])
        ws.write_number(rr, 3, int(r["prev_TA"]), fm["txt"]); ws.write_number(rr, 4, int(r["next_TA"]), fm["txt"])
        ws.write_number(rr, 5, int(r["gap_yrs"]), fm["txt"])
        ws.write_number(rr, 6, round(r["kbd_prev"]), fm["num"]); ws.write_number(rr, 7, round(r["kbd_next"]), fm["num"])
    nrow = hr + 1 + len(cf)
    ws.merge_range(nrow, 0, nrow, 8, f"{len(cf_all)} CDU/FCC planned-pairs are <={engine.TURNAROUND_CYCLE_YEARS-1}yr "
                   f"apart and reach {CY}+ (top 16 by size shown). 'kbd' = day-weighted kbd-months that year. "
                   "A short gap can be real minor maintenance, not a full turnaround -- check magnitude.", fm["txtw"])
    # --- 2) double count ---
    dr = nrow + 2
    ws.write(dr, 0, "Unit-months summing to >100% of nameplate (overlap / double-count)", fm["secn"])
    dc = engine.double_count_flags(ctx["df"])
    dcw = dc[dc["year"].between(CY, FY)] if len(dc) else dc
    dcw = dcw.sort_values("excess_kbd", ascending=False).head(12) if len(dcw) else dcw
    heads2 = ["Year", "Mon", "Plant", "Unit", "Class", "Σ kbd", "Nameplate", "Ratio", "Type"]
    hr2 = dr + 1
    order = [0, 1, 2, 3, 4, 5, 6, 7, 8]
    for j, h in enumerate(heads2):
        ws.write(hr2, j, h, fm["hl"] if j in (2, 3) else fm["h"])
    for i, (_, r) in enumerate(dcw.iterrows()):
        rr = hr2 + 1 + i
        ws.write_number(rr, 0, int(r["year"]), fm["txt"]); ws.write_number(rr, 1, int(r["month"]), fm["txt"])
        ws.write(rr, 2, str(r["plant"])[:22], fm["txt"]); ws.write(rr, 3, str(r["unit_name"])[:22], fm["txt"])
        ws.write(rr, 4, str(r["focus"]), fm["txt"])
        ws.write_number(rr, 5, round(r["sum_kbd"], 1), fm["num"]); ws.write_number(rr, 6, round(r["nameplate"], 1), fm["num"])
        ws.write_number(rr, 7, round(r["ratio"], 2), fm["f2"]); ws.write(rr, 8, str(r["types"]), fm["txt"])
    drow = hr2 + 1 + len(dcw)
    total_excess = float(dc[dc["year"].between(CY, FY)]["excess_kbd"].sum()) if len(dc) else 0.0
    ws.merge_range(drow, 0, drow, 8, f"{len(dc[dc['year'].between(CY, FY)]) if len(dc) else 0} focus unit-months "
                   f"in {CY}-{FY} exceed nameplate (~{total_excess:,.0f} excess kbd-months) -- small vs the totals, "
                   "but each is an overlap or a duplicate record (or one event logged as both planned and unplanned).",
                   fm["txtw"])


TAB = {"Index": "#1F3864", "Assumptions": "#BF9000", "Data": "#808080",
       "Historicals": "#538DD5", "Per-Unit": "#2E5496", "Biggest": "#2E5496",
       "H1 by Unit": "#2E5496", "PADD by Unit": "#2E5496", "Naphtha": "#548235",
       "ExxonMobil": "#2E5496", "Forecast": "#C55A11", "Sensitivity": "#C55A11",
       "Stress Test": "#C55A11", "Statistics": "#7030A0", "Regression": "#7030A0",
       "PADD Connectivity": "#C55A11", "Data Quality": "#A6A6A6", "What's Changed": "#C00000"}


def build_workbook(ctx, assets, path):
    """Write the Excel model. `assets` is the render_all() chart dict."""
    wb = xlsxwriter.Workbook(path, {"nan_inf_to_errors": True})
    fm = _formats(wb)
    base = _base(ctx["df"])
    _index(wb, fm)
    _assumptions(wb, fm, ctx)
    _data_sheet(wb, fm, base)
    _whats_changed(wb, fm, ctx)
    _historicals(wb, fm, base)
    _per_unit(wb, fm, ctx, assets)
    _biggest(wb, fm, ctx, assets)
    _h1(wb, fm, ctx, assets)
    _padd(wb, fm, ctx, base, assets)
    _naphtha(wb, fm, ctx, assets)
    _exxon(wb, fm, ctx, assets)
    _forecast(wb, fm, ctx, assets)
    _sensitivity(wb, fm, ctx)
    _stress(wb, fm, ctx)
    _statistics(wb, fm, base)
    _regression(wb, fm, base)
    _padd_connectivity(wb, fm, ctx)
    _data_quality(wb, fm, ctx)
    # professional polish: colored tabs, hidden gridlines (except the data table)
    for ws in wb.worksheets():
        ws.set_tab_color(TAB.get(ws.name, "#808080"))
        if ws.name != "Data":
            ws.hide_gridlines(2)
        ws.set_zoom(110)
    wb.worksheets()[0].activate()
    wb.close()
    return path


def main():
    ap = argparse.ArgumentParser(description="Refinery outage Excel model")
    ap.add_argument("excel", nargs="?", default=INPUT_PATH)
    ap.add_argument("--out", default=OUT_PATH)
    args = ap.parse_args()
    print(f"Loading {args.excel.strip()} ...")
    ctx = engine.build_context(args.excel)
    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        print("Rendering charts ...")
        assets = charts.render_all(ctx, tmp)
        print(f"Building model -> {args.out}")
        build_workbook(ctx, assets, args.out)
    print("Done.")


if __name__ == "__main__":
    sys.exit(main())
