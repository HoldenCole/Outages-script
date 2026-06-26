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
INPUT_PATH = str(_ROOT / "data" / "Refinery_Outages_Enhanced.xlsx")
OUT_PATH = str(_ROOT / "output" / "outage_model.xlsx")
ASOF = "June 26th, 2026"

MONTHS = engine.MONTHS
FOCUS = engine.FOCUS_ORDER
IMG = {"x_scale": 0.6, "y_scale": 0.6, "object_position": 1}

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
    _title(ws, fm, "Refinery Outages 2027 Model: Index",
           "Every deck slide maps to a model sheet. Open the sheet to see the detail and the formula.")
    rows = [
        ("Slide", "Deck slide", "Model sheet", "What it shows and how it is calculated"),
        ("2", "Total 2027 Outages by Unit", "Per-Unit",
         "Day-weighted concurrent offline per unit by month (=SUMIFS over Data), split confirmed vs indicative."),
        ("3", "What's Driving the Numbers (biggest outages)", "Biggest",
         "The biggest individual 2027 outages by nameplate kbd, with PADD; one row per physical unit."),
        ("4", "H1 Planned per Unit & Month", "H1 by Unit",
         "H1 (Jan-Jun) planned offline per unit and month, 2025/26/27 (=SUMIFS, type=Planned); H1 avg =AVERAGE."),
        ("5", "Outages by PADD by Unit", "PADD by Unit",
         "2027 CDU and FCC offline by PADD and month (=SUMIFS over Data)."),
        ("6", "Naphtha Balance", "Naphtha",
         "net = reformer offline x intake  -  CDU offline x naphtha yield. Live: =SUMIFS x Assumptions cells."),
        ("7", "ExxonMobil 2027 by Unit", "ExxonMobil",
         "Per-unit ExxonMobil 2027 turnarounds, verified against the corporate plan."),
        ("8", "2027 Unplanned Scenario", "Forecast",
         "baseline x multiplier (live off Assumptions); implied total = scenario unplanned + booked planned."),
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
        "2027 completeness: ExxonMobil gave a full-year plan (verified vs their schedule); every other "
        "operator is H1-confirmed only, so non-Exxon H2 is indicative.",
        "Forecast baseline is completeness-aware: each calendar month is averaged over the years that "
        "actually reported it, so 2026 H1 sharpens H1 while H2 stays on 2023-2025.",
        "Naphtha balance: net = reformer demand removed minus CDU supply removed (+ surplus / - deficit).",
    ]
    ws.write(18, 0, "Methodology", fm["secn"])
    for i, n in enumerate(notes):
        ws.merge_range(19 + i, 0, 19 + i, 5, f"- {n}", fm["txtw"])


def _data_sheet(wb, fm, ctx):
    """One row per (year, month, plant, unit, type): the pullable source the
    analysis sheets SUMIFS over. Day-weighted cap (kbd) and nameplate (raw)."""
    df = ctx["df"].copy()
    df["focus"] = df["focus"].fillna("")
    df["padd"] = df["padd"].fillna("")
    g = (df.groupby(["year", "month", "plant", "unit_name", "type"], dropna=False)
         .agg(cap_kbd=("cap_kbd", "max"), cap_raw=("cap_raw", "max"),
              operator=("operator", "first"), padd=("padd", "first"),
              focus=("focus", "first"), unit_cat=("unit_cat", "first")).reset_index())
    g = g.dropna(subset=["year", "month"])
    g["year"] = g["year"].astype(int); g["month"] = g["month"].astype(int)
    g["month_name"] = g["month"].map(lambda m: MONTHS[m - 1])
    g = g.sort_values(["year", "month", "cap_kbd"], ascending=[True, True, False]).reset_index(drop=True)

    ws = wb.add_worksheet("Data")
    ws.set_column(0, 1, 6); ws.set_column(2, 2, 6); ws.set_column(3, 3, 28); ws.set_column(4, 4, 22)
    ws.set_column(5, 5, 9); ws.set_column(6, 7, 14); ws.set_column(8, 8, 26); ws.set_column(9, 9, 10)
    ws.set_column(10, 11, 9)
    _title(ws, fm, "Data: deduped per year / month / plant / unit / type (the model source)",
           "Concurrent offline, day-weighted kbd and nameplate. Every analysis sheet SUMIFS over this.")
    heads = ["year", "monthnum", "month", "plant", "operator", "padd", "focus", "unit_cat",
             "unit_name", "type", "cap_kbd", "cap_raw"]
    r0 = 3
    for j, h in enumerate(heads):
        ws.write(r0, j, h, fm["h"])
    order = ["year", "month", "month_name", "plant", "operator", "padd", "focus", "unit_cat",
             "unit_name", "type", "cap_kbd", "cap_raw"]
    for i, (_, row) in enumerate(g.iterrows()):
        rr = r0 + 1 + i
        for j, col in enumerate(order):
            v = row[col]
            if col in ("cap_kbd", "cap_raw"):
                ws.write_number(rr, j, float(v), fm["num"])
            elif col in ("year", "month"):
                ws.write_number(rr, j, int(v), fm["txt"])
            else:
                ws.write(rr, j, "" if (isinstance(v, float) and np.isnan(v)) else str(v), fm["txt"])
    ws.freeze_panes(r0 + 1, 0)
    ws.autofilter(r0, 0, r0 + len(g), len(heads) - 1)
    return len(g)


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
    years = [2023, 2024, 2025, 2026, 2027]
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
    _title(ws, fm, "Biggest 2027 Outages by Unit (nameplate kbd, from Data)",
           "Individual units, never added. PADD shows where the tonnage sits.")
    ev = engine.unit_events(ctx["df"], year=2027)
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
           "Like-for-like cross-year read; 2027 is confirmed through H1. H1 avg = AVERAGE of the row.")
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
            ws, fm, top, 0, [2025, 2026, 2027], h1m,
            lambda y, mo, _f=f: _si("K", ("A", y), ("B", mnum[mo]), ("G", _f), ("J", "PLANNED")),
            lambda y, mo, _m=m: (_m.loc[y, mo] if y in _m.index else 0.0), "Year") + 2
        blocks[f] = top                     # header row of this block; data rows top+1..top+3
    # H1 averages = AVERAGE of each unit/year row in the blocks above
    h1 = ctx["h1_focus_planned"]
    ws.write(r, 0, "H1 average offline (=AVERAGE Jan:Jun)", fm["secn"])
    hdr = r + 1
    ws.write(hdr, 0, "Unit", fm["hl"])
    for j, y in enumerate([2025, 2026, 2027]):
        ws.write(hdr, 1 + j, y, fm["h"])
    for i, f in enumerate(FOCUS):
        ws.write(hdr + 1 + i, 0, engine.FOCUS_LABEL[f], fm["rowh"])
        for j, y in enumerate([2025, 2026, 2027]):
            drow = blocks[f] + 1 + j        # 0-based data row for (f, year)
            ws.write_formula(hdr + 1 + i, 1 + j, f"=AVERAGE(B{drow+1}:G{drow+1})", fm["num"],
                             float(h1.loc[f, y] if y in h1.columns else 0.0))
    ws.insert_image(3, 9, assets["h1_month_by_unit"], IMG)


def _padd(wb, fm, ctx, assets):
    ws = wb.add_worksheet("PADD by Unit")
    ws.set_column(0, 0, 11); ws.set_column(1, 12, 8)
    _title(ws, fm, "2027 Offline by PADD & Month, per Unit (kbd) = SUMIFS over Data",
           "Where it tightens. CDU (crude) and FCC (cat cracker), each kept separate.")
    mnum = {mo: i + 1 for i, mo in enumerate(MONTHS)}
    r = 3
    for f in ("CDU", "FCC"):
        g = ctx["focus_padd"][2027][f]
        ws.write(r, 0, f"{engine.FOCUS_LABEL[f]} by PADD", fm["secn"])
        r = _matrix_formula(
            ws, fm, r + 1, 0, list(g.index), MONTHS,
            lambda p, mo, _f=f: _si("K", ("A", 2027), ("B", mnum[mo]), ("G", _f), ("F", p)),
            lambda p, mo, _g=g: _g.loc[p, mo], "PADD") + 2
    ws.insert_image(3, 14, assets["cdu_padd_27"], IMG)
    ws.insert_image(22, 14, assets["fcc_padd_27"], IMG)


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
        ws.write_formula(rr, 1, _si("K", ("A", 2027), ("B", mnum[mo]), ("G", "CDU")),
                         fm["num"], float(nb["cdu_offline"][i]))
        ws.write_formula(rr, 2, f"=B{rr+1}*naph_yield", fm["num"], float(nb["supply_removed"][i]))
        ws.write_formula(rr, 3, _si("K", ("A", 2027), ("B", mnum[mo]), ("G", "Reformer")),
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
    ws.write(tot + 3, 0, f"2027 annual net = {nb['annual_net']:,.0f} kbd "
                         f"({nb['n_deficit']} deficit months, {nb['n_surplus']} surplus).", fm["sub"])
    ws.insert_image(3, 7, assets["naphtha_balance"], IMG)


def _exxon(wb, fm, ctx, assets):
    ws = wb.add_worksheet("ExxonMobil")
    ws.set_column(0, 0, 24); ws.set_column(1, 1, 26); ws.set_column(2, 3, 11)
    ws.set_column(4, 4, 9); ws.set_column(5, 6, 11); ws.set_column(7, 7, 24)
    _title(ws, fm, "ExxonMobil 2027 Focus-Unit Turnarounds, per Unit (kbd)",
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
    ws.set_column(0, 0, 22); ws.set_column(1, 12, 9); ws.set_column(13, 13, 11)
    _title(ws, fm, "2027 Unplanned Scenario (kbd) = baseline x multiplier (live)",
           "Each path recomputes off the Assumptions multipliers. A risk range, not a forecast.")
    base = engine.baseline_profile(ctx["df"], ctx["scenario"]["window"])
    r0 = 3
    ws.write(r0, 0, "Series", fm["hl"])
    for j, mo in enumerate(MONTHS):
        ws.write(r0, 1 + j, mo, fm["h"])
    ws.write(r0, 13, "Annual", fm["h"])
    ws.write(r0 + 1, 0, f"Baseline ({ctx['scenario']['window']}, completeness-aware)", fm["rowh"])
    for j, mo in enumerate(MONTHS):
        ws.write_number(r0 + 1, 1 + j, float(base[mo]), fm["num"])
    ws.write_formula(r0 + 1, 13, f"=SUM(B{r0+2}:M{r0+2})", fm["num"], float(base.sum()))
    fan = ctx["scenario_fan"]
    for k, (nm, ref) in enumerate([("Conservative", "mult_cons"), ("Average", "mult_avg"),
                                   ("Active", "mult_act")]):
        rr = r0 + 2 + k
        ws.write(rr, 0, nm, fm["rowh"])
        for j in range(12):
            col = chr(ord("B") + j)
            ws.write_formula(rr, 1 + j, f"={col}{r0+2}*{ref}", fm["num"], float(fan[nm].iloc[j]))
        ws.write_formula(rr, 13, f"=SUM(B{rr+1}:M{rr+1})", fm["num"], float(fan[nm].sum()))
    pl = float(ctx["summary"].loc[2027, "Planned"]) if 2027 in ctx["summary"].index else 0.0
    br = r0 + 6
    ws.write(br, 0, "Booked planned 2027", fm["lbl"]); ws.write_number(br, 1, pl, fm["num"])
    ws.write(br + 1, 0, "Implied total (planned + unplanned)", fm["secn"])
    ws.write(br + 2, 0, "Scenario", fm["hl"])
    ws.write(br + 2, 1, "Unplanned", fm["h"]); ws.write(br + 2, 2, "+ Planned", fm["h"])
    ws.write(br + 2, 3, "Implied total", fm["h"])
    for k, nm in enumerate(["Conservative", "Average", "Active"]):
        rr = br + 3 + k
        srow = r0 + 3 + k                   # 1-based Excel row of this scenario's annual cell
        ws.write(rr, 0, nm, fm["rowh"])
        ws.write_formula(rr, 1, f"=N{srow}", fm["num"], float(fan[nm].sum()))
        ws.write_formula(rr, 2, f"=$B${br+1}", fm["num"], pl)
        ws.write_formula(rr, 3, f"=B{rr+1}+C{rr+1}", fm["num"], float(fan[nm].sum()) + pl)
    sb = ctx["scenario_bands"]
    ws.write(br + 7, 0, "History (complete years, annual unplanned)", fm["secn"])
    for k, (lab, key) in enumerate([("P25", "p25"), ("Median", "p50"), ("P90", "p90"), ("Mean", "mean")]):
        ws.write(br + 8 + k, 0, lab, fm["lbl"]); ws.write_number(br + 8 + k, 1, float(sb[key]), fm["num"])
    ws.insert_image(3, 15, assets["fan"], IMG)
    ws.insert_image(26, 15, assets["scenario_total"], IMG)


# Colored tabs grouped by purpose: navigation, source, per-unit analysis, balance, forecast
TAB = {"Index": "#1F3864", "Assumptions": "#BF9000", "Data": "#808080",
       "Per-Unit": "#2E5496", "Biggest": "#2E5496", "H1 by Unit": "#2E5496",
       "PADD by Unit": "#2E5496", "Naphtha": "#548235", "ExxonMobil": "#2E5496",
       "Forecast": "#C55A11"}


def build_workbook(ctx, assets, path):
    """Write the Excel model. `assets` is the render_all() chart dict."""
    wb = xlsxwriter.Workbook(path, {"nan_inf_to_errors": True})
    fm = _formats(wb)
    _index(wb, fm)
    _assumptions(wb, fm, ctx)
    _data_sheet(wb, fm, ctx)
    _per_unit(wb, fm, ctx, assets)
    _biggest(wb, fm, ctx, assets)
    _h1(wb, fm, ctx, assets)
    _padd(wb, fm, ctx, assets)
    _naphtha(wb, fm, ctx, assets)
    _exxon(wb, fm, ctx, assets)
    _forecast(wb, fm, ctx, assets)
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
