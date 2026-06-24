#!/usr/bin/env python3
"""
build_workbook.py
Institutional, trading-floor refinery-outage workbook (XlsxWriter).

Consolidated into eight findable sheets plus a hidden backing-data sheet:

  Cover | Dashboard | Explorer | Trends | PADD | Units & Refineries |
  Events & TAs | Model      (+ hidden Data)

Explorer is fully interactive: dropdowns (PADD / Unit / type) drive SUMIFS over
the hidden Data sheet, so the monthly grid, the month-by-month YoY% grid and the
charts all recompute live. Scenario & Sensitivity (the "Model" sheet) are live
formulas wired to data-validation inputs.

Usage:
    python build_workbook.py                       # uses INPUT_PATH
    python build_workbook.py path/to/export.xlsx --out outage_workbook.xlsx
"""
import argparse
import sys
from pathlib import Path

import xlsxwriter
from xlsxwriter.utility import (xl_rowcol_to_cell as A1, xl_range_abs, xl_range,
                                xl_col_to_name)

import engine

_ROOT = Path(__file__).resolve().parent.parent
INPUT_PATH = str(_ROOT / "data" / "rEFINERY oUTAGES.xlsx")
OUT_PATH = str(_ROOT / "output" / "outage_workbook.xlsx")

# --------------------------------------------------------------------------- palette
NAVY = "#1F3864"
BLUE = "#2E5496"
RED = "#C00000"
GOLD = "#BF9000"
GREEN = "#548235"
ORANGE = "#ED7D31"
PURPLE = "#7030A0"
LT_BLUE = "#D6E0F0"
LT_GRAY = "#F2F2F2"
YELLOW = "#FFF2CC"
GRAY = "#808080"
WHITE = "#FFFFFF"

KBD = '#,##0;(#,##0);"-"'
PCT = "0.0%"
PCTS = "+0.0%;-0.0%;0.0%"          # signed percent for YoY
MULT = '0.0"x"'

MONTHS = engine.MONTHS
PADDS = engine.PADD_ORDER
PARTIAL = set(engine.PARTIAL_YEARS)
DISP_YEARS = list(range(2016, 2028))            # 2016-2027 standard display window


def short_op(name, n=18):
    s = str(name).title()
    for w in ("Corporation", "Company", "Incorporated", "Petroleum", "Refining",
              " Llc", " Lp", " L.P.", "North America", "Products", "Energy", " Inc"):
        s = s.replace(w, "")
    return (" ".join(s.split()) or str(name).title())[:n]


class Build:
    def __init__(self, ctx, out_path):
        self.ctx = ctx
        self.wb = xlsxwriter.Workbook(out_path, {"nan_inf_to_errors": True})
        self.f = self._formats()
        self.data_refs = {}
        self.sc_refs = {}

    # ----------------------------------------------------------------- formats
    def _formats(self):
        wb = self.wb
        base = {"font_name": "Arial", "font_size": 10}
        mk = lambda **kw: wb.add_format({**base, **kw})
        bands = {"h_navy": NAVY, "h_blue": BLUE, "h_red": RED, "h_green": GREEN,
                 "h_gold": "#9C6500", "h_orange": "#C55A11", "h_purple": PURPLE}
        f = {
            "title":    mk(font_size=17, bold=True, font_color=NAVY),
            "subtitle": mk(font_size=10, italic=True, font_color="#595959"),
            "colhdr":   mk(bold=True, font_color=WHITE, bg_color=NAVY, align="center",
                           valign="vcenter", border=1, border_color=WHITE, font_size=9),
            "colhdr_l": mk(bold=True, font_color=WHITE, bg_color=NAVY, align="left",
                           valign="vcenter", border=1, border_color=WHITE, indent=1, font_size=9),
            "rowlab":   mk(align="left", valign="vcenter", indent=1),
            "rowlab_b": mk(bold=True, align="left", valign="vcenter", indent=1),
            "rowlab_p": mk(align="left", valign="vcenter", indent=1, italic=True, font_color=GRAY),
            "kbd":      mk(num_format=KBD, align="right"),
            "kbd_b":    mk(num_format=KBD, align="right", bold=True),
            "kbd_p":    mk(num_format=KBD, align="right", italic=True, font_color=GRAY),
            "kbd_sh":   mk(num_format=KBD, align="right", bg_color=LT_GRAY),
            "kbd_p_sh": mk(num_format=KBD, align="right", italic=True, font_color=GRAY, bg_color=LT_GRAY),
            "pct":      mk(num_format=PCT, align="right"),
            "pct_p":    mk(num_format=PCT, align="right", italic=True, font_color=GRAY),
            "pcts":     mk(num_format=PCTS, align="right"),
            "pct_g":    mk(num_format=PCT, align="right", font_color="#008000"),
            "mult":     mk(num_format=MULT, align="right"),
            "yr":       mk(num_format="@", bold=True, align="center", valign="vcenter",
                           font_color=WHITE, bg_color=NAVY, border=1, border_color=WHITE, font_size=9),
            "yr_p":     mk(num_format="@", bold=True, align="center", valign="vcenter",
                           font_color="#D9D9D9", bg_color=NAVY, border=1, border_color=WHITE, font_size=9),
            "yrn":      mk(num_format="0", bold=True, align="center", valign="vcenter",
                           font_color=WHITE, bg_color=NAVY, border=1, border_color=WHITE, font_size=9),
            "yrn_p":    mk(num_format="0", bold=True, align="center", valign="vcenter",
                           font_color="#D9D9D9", bg_color=NAVY, border=1, border_color=WHITE, font_size=9),
            "na":       mk(align="right", italic=True, font_color=GRAY),
            "note":     mk(font_size=9.5, font_color="#404040", text_wrap=True, valign="top"),
            "note_b":   mk(font_size=9.5, bold=True, font_color=NAVY, valign="top"),
            "red_note": mk(font_size=9.5, bold=True, font_color=RED, text_wrap=True, valign="top"),
            "link":     mk(font_color=BLUE, underline=1, align="left", indent=1, bold=True),
            "kpi_lab":  mk(font_size=9.5, bold=True, font_color=WHITE, bg_color=BLUE,
                           align="center", valign="vcenter", text_wrap=True),
            "kpi_num":  mk(font_size=18, bold=True, font_color=NAVY, bg_color=LT_BLUE,
                           align="center", valign="vcenter", num_format="#,##0"),
            "kpi_pct":  mk(font_size=18, bold=True, font_color=NAVY, bg_color=LT_BLUE,
                           align="center", valign="vcenter", num_format="0%"),
            "kpi_txt":  mk(font_size=14, bold=True, font_color=NAVY, bg_color=LT_BLUE,
                           align="center", valign="vcenter"),
            "kpi_sub":  mk(font_size=8, italic=True, font_color="#595959", bg_color=LT_BLUE,
                           align="center", valign="vcenter"),
            "in_lab":   mk(bold=True, align="left", valign="vcenter", indent=1, bg_color=YELLOW,
                           border=1, border_color="#BF8F00"),
            "in_val":   mk(font_color="#0000FF", bold=True, align="center", valign="vcenter",
                           bg_color=YELLOW, border=1, border_color="#BF8F00"),
            "in_pct":   mk(num_format=PCT, font_color="#0000FF", bold=True, align="center",
                           valign="vcenter", bg_color=YELLOW, border=1, border_color="#BF8F00"),
            "in_mult":  mk(num_format=MULT, font_color="#0000FF", bold=True, align="center",
                           valign="vcenter", bg_color=YELLOW, border=1, border_color="#BF8F00"),
            "in_kbd":   mk(num_format=KBD, font_color="#0000FF", bold=True, align="center",
                           valign="vcenter", bg_color=YELLOW, border=1, border_color="#BF8F00"),
            "calc":     mk(num_format=KBD, align="right"),
            "calc_b":   mk(num_format=KBD, align="right", bold=True, bg_color=LT_BLUE),
            "calc_grn": mk(num_format=KBD, align="right", font_color="#008000"),
            "hm_cell":  mk(num_format="#,##0", align="center", valign="vcenter", border=1,
                           border_color="#BFBFBF"),
            "hm_base":  mk(num_format="#,##0", align="center", valign="vcenter", border=2,
                           border_color=NAVY, bold=True),
            "data_hdr": mk(bold=True, font_color=WHITE, bg_color=NAVY, align="center", font_size=9),
            "data":     mk(font_size=9),
            "date":     mk(num_format="mm/dd/yy", align="center"),
            "date_sh":  mk(num_format="mm/dd/yy", align="center", bg_color=LT_GRAY),
        }
        for name, color in bands.items():
            f[name] = mk(font_size=10.5, bold=True, font_color=WHITE, bg_color=color,
                         align="left", valign="vcenter", indent=1)
        return f

    def colors(self):
        return {"navy": NAVY, "blue": BLUE, "red": RED, "gold": GOLD}

    # ----------------------------------------------------------------- helpers
    def _setup(self, ws, tab, landscape=True, gridlines=True):
        if gridlines:
            ws.hide_gridlines(2)
        ws.set_tab_color(tab)
        if landscape:
            ws.set_landscape()
            ws.set_paper(1)
            ws.fit_to_pages(1, 0)
        ws.set_margins(0.3, 0.3, 0.35, 0.35)

    def _band(self, ws, r, c0, c1, text, fmt="h_navy"):
        ws.merge_range(r, c0, r, c1, text, self.f[fmt])
        ws.set_row(r, 19)
        return r + 1

    def _chrome(self, ws, last_col=15):
        """Home link + data-vintage stamp on the top row of every sheet."""
        ws.write_url(0, 1, "internal:'Cover'!A1", self.f["link"], "Home")
        d = self.ctx["diag"]
        y0, y1 = d["years"]
        fmt = self.wb.add_format({"font_name": "Arial", "font_size": 8, "italic": True,
                                  "font_color": "#808080", "align": "right"})
        ws.merge_range(0, max(3, last_col - 5), 0, last_col,
                       f"Snowflake export  |  {d['rows']:,} rows  |  {y0}-{y1}  |  primary metric kbd",
                       fmt)

    def _yr(self, y, num=False):
        if num:
            return self.f["yrn_p"] if y in PARTIAL else self.f["yrn"]
        return self.f["yr_p"] if y in PARTIAL else self.f["yr"]

    def _kf(self, y, shade=False):
        if y in PARTIAL:
            return self.f["kbd_p_sh"] if shade else self.f["kbd_p"]
        return self.f["kbd_sh"] if shade else self.f["kbd"]

    def _pf(self, y):
        return self.f["pct_p"] if y in PARTIAL else self.f["pct"]

    def _month_header(self, ws, r, label="Year", total=True):
        ws.write(r, 1, label, self.f["colhdr_l"])
        for j, m in enumerate(MONTHS):
            ws.write(r, 2 + j, m, self.f["colhdr"])
        if total:
            ws.write(r, 14, "Total", self.f["colhdr"])
        return r + 1

    # ============================================================= block writers
    def _annual_block(self, ws, r):
        s = self.ctx["summary"]
        r = self._band(ws, r, 1, 8, "Annual Capacity Offline (kbd)")
        for j, h in enumerate(["Year", "Planned", "Unplanned", "Total", "Events", "Unpl %", "YoY Δ", "YoY %"]):
            ws.write(r, 1 + j, h, self.f["colhdr_l"] if j == 0 else self.f["colhdr"])
        r += 1
        first = r
        for i, y in enumerate(DISP_YEARS):
            kf = self.f["kbd_p"] if y in PARTIAL else self.f["kbd"]
            pf = self._pf(y)
            ws.write_number(r, 1, y, self._yr(y))
            ws.write_number(r, 2, float(s.loc[y, "Planned"]), kf)
            ws.write_number(r, 3, float(s.loc[y, "Unplanned"]), kf)
            ws.write_formula(r, 4, f"={A1(r,2)}+{A1(r,3)}", kf)
            ws.write_number(r, 5, int(s.loc[y, "Events"]), kf)
            ws.write_formula(r, 6, f"=IF({A1(r,4)}=0,0,{A1(r,3)}/{A1(r,4)})", pf)
            if i == 0:
                ws.write_blank(r, 7, None, kf); ws.write_blank(r, 8, None, pf)
            else:
                ws.write_formula(r, 7, f"={A1(r,4)}-{A1(r-1,4)}", kf)
                ws.write_formula(r, 8, f"=IF({A1(r-1,4)}=0,0,{A1(r,7)}/{A1(r-1,4)})",
                                 self.f["pcts"] if y not in PARTIAL else self.f["pct_p"])
            r += 1
        return r + 1

    def _compare_block(self, ws, r):
        cmp = self.ctx["compare"]
        r = self._band(ws, r, 1, 5, "Targeted Year Comparisons (Δ% vs base year; n/a vs 2027 unplanned)")
        for key, (ya, yb) in [("2025v2026", (2025, 2026)), ("2025v2027", (2025, 2027)),
                              ("2026v2027", (2026, 2027))]:
            ws.write(r, 1, f"{ya} → {yb}", self.f["rowlab_b"])
            for j, h in enumerate([str(ya), str(yb), "Δ kbd", "Δ %"]):
                ws.write(r, 2 + j, h, self.f["colhdr"])
            r += 1
            for metric in ["Plan + Unplanned", "Planned", "Unplanned"]:
                blk = cmp[key][metric]
                ws.write(r, 1, metric, self.f["rowlab"])
                if blk["a"] is None or blk["b"] is None:
                    if blk["a"] is not None:
                        ws.write_number(r, 2, blk["a"], self.f["kbd"])
                    else:
                        ws.write(r, 2, "n/a", self.f["na"])
                    for c in (3, 4, 5):
                        ws.write(r, c, "n/a", self.f["na"])
                else:
                    ws.write_number(r, 2, blk["a"], self.f["kbd"])
                    ws.write_number(r, 3, blk["b"], self.f["kbd"])
                    ws.write_formula(r, 4, f"={A1(r,3)}-{A1(r,2)}", self.f["kbd"])
                    ws.write_formula(r, 5, f"=IF({A1(r,2)}=0,0,{A1(r,4)}/{A1(r,2)})", self.f["pcts"])
                r += 1
        return r + 1

    def _monthly_levels(self, ws, r):
        """Three monthly matrices with heatmap colour-scale + in-cell sparklines;
        returns (next_r, {title: first_row})."""
        blocks = [("Total Offline", self.ctx["monthly_total"], "h_navy", "#BDD7EE"),
                  ("Planned", self.ctx["monthly_planned"], "h_blue", "#BDD7EE"),
                  ("Unplanned", self.ctx["monthly_unplanned"], "h_red", "#F8C9C4")]
        firsts = {}
        for title, mat, band, spcol in blocks:
            r = self._band(ws, r, 1, 15, f"{title} - Monthly (kbd)", band)
            r = self._month_header(ws, r)
            ws.write(r - 1, 15, "Trend", self.f["colhdr"])
            first = r
            for y in DISP_YEARS:
                kf = self.f["kbd_p"] if y in PARTIAL else self.f["kbd"]
                ws.write_number(r, 1, y, self._yr(y))
                for j, m in enumerate(MONTHS):
                    ws.write_number(r, 2 + j, float(mat.loc[y, m]) if y in mat.index else 0.0, kf)
                ws.write_formula(r, 14, f"=SUM({A1(r,2)}:{A1(r,13)})",
                                 self.f["kbd_p"] if y in PARTIAL else self.f["kbd_b"])
                ws.add_sparkline(r, 15, {"range": f"{ws.name}!{xl_range(r,2,r,13)}",
                                         "type": "line", "series_color": NAVY,
                                         "high_point": True, "high_color": RED})
                r += 1
            ws.conditional_format(first, 2, r - 1, 13,
                                  {"type": "2_color_scale", "min_type": "num", "min_value": 0,
                                   "min_color": WHITE, "max_color": spcol})
            firsts[title] = first
            r += 1
        ws.set_column("P:P", 8)
        return r, firsts

    def _monthly_yoy(self, ws, r, firsts):
        """Month-by-month YoY% matrices (live formulas) for Unplanned & Total."""
        for title, band in [("Unplanned", "h_red"), ("Total Offline", "h_navy")]:
            lvl = firsts[title]
            r = self._band(ws, r, 1, 13, f"{title} - YoY % Change by Month (this month vs same month, prior year)", band)
            r = self._month_header(ws, r, total=False)
            grid_first = r
            for i, y in enumerate(DISP_YEARS):
                pf = self.f["pct_p"] if y in PARTIAL else self.f["pcts"]
                ws.write_number(r, 1, y, self._yr(y))
                for j in range(12):
                    if i == 0:
                        ws.write(r, 2 + j, "-", self.f["na"])
                    else:
                        cur, prev = A1(lvl + i, 2 + j), A1(lvl + i - 1, 2 + j)
                        ws.write_formula(r, 2 + j, f"=IF({prev}=0,\"n/a\",({cur}-{prev})/{prev})", pf)
                r += 1
            ws.conditional_format(grid_first, 2, r - 1, 13,
                                  {"type": "3_color_scale", "min_color": "#63BE7B",
                                   "mid_color": WHITE, "max_color": "#F8696B"})
            r += 1
        return r

    # ===================================================================== DATA
    def data_sheet(self):
        ws = self.wb.add_worksheet("Data")
        cols = ["year", "month", "key", "type", "kbd", "mogas"]
        for dim, c0 in [("padd", 0), ("unit", 8)]:
            df = self.ctx["tidy_padd"] if dim == "padd" else self.ctx["tidy_unit"]
            for j, h in enumerate(cols):
                ws.write(0, c0 + j, h, self.f["data_hdr"])
            for i, (_, row) in enumerate(df.iterrows(), start=1):
                ws.write_number(i, c0 + 0, int(row["year"]), self.f["data"])
                ws.write_number(i, c0 + 1, int(row["month"]), self.f["data"])
                ws.write(i, c0 + 2, str(row["key"]), self.f["data"])
                ws.write(i, c0 + 3, str(row["type"]), self.f["data"])
                ws.write_number(i, c0 + 4, float(row["kbd"]), self.f["data"])
                ws.write_number(i, c0 + 5, float(row["mogas"]), self.f["data"])
            n = len(df)
            self.data_refs[dim] = {
                "yr": "Data!" + xl_range_abs(1, c0, n, c0),
                "mo": "Data!" + xl_range_abs(1, c0 + 1, n, c0 + 1),
                "key": "Data!" + xl_range_abs(1, c0 + 2, n, c0 + 2),
                "type": "Data!" + xl_range_abs(1, c0 + 3, n, c0 + 3),
                "kbd": "Data!" + xl_range_abs(1, c0 + 4, n, c0 + 4),
                "mogas": "Data!" + xl_range_abs(1, c0 + 5, n, c0 + 5),
                "keys": sorted(df["key"].unique(), key=lambda k: (k not in ("Total US", "All Units"), k)),
            }
        ws.hide()

    # =============================================================== EXPLORER
    def _explorer_panel(self, ws, r, dim, title, band, default_key):
        """One interactive panel: dropdowns -> SUMIFS monthly grid + YoY% + chart."""
        D = self.data_refs[dim]
        keys = D["keys"]
        SHEET = "Explorer"
        r = self._band(ws, r, 1, 14, title, band)
        # control row: Select | Type | Measure (all dropdowns)
        ws.write(r, 1, "Select", self.f["in_lab"])
        ws.merge_range(r, 2, r, 3, default_key, self.f["in_val"])
        ws.data_validation(r, 2, r, 2, {"validate": "list", "source": keys})
        ws.write(r, 4, "Type", self.f["in_lab"])
        ws.merge_range(r, 5, r, 6, "Unplanned", self.f["in_val"])
        ws.data_validation(r, 5, r, 5, {"validate": "list", "source": ["All", "Planned", "Unplanned"]})
        ws.write(r, 7, "Measure", self.f["in_lab"])
        ws.merge_range(r, 8, r, 10, "Capacity (kbd)", self.f["in_val"])
        ws.data_validation(r, 8, r, 8, {"validate": "list", "source": ["Capacity (kbd)", "Mogas-eq (kbd)"]})
        ws.write(r, 12, "Sel. 2025 total:", self.f["note_b"])
        key_cell = A1(r, 2, row_abs=True, col_abs=True)
        type_cell = A1(r, 5, row_abs=True, col_abs=True)
        meas_cell = A1(r, 8, row_abs=True, col_abs=True)
        kpi_row = r
        r += 1

        def sumifs(year_cell, j):
            args = (f"{D['yr']},{year_cell},{D['mo']},{j+1},"
                    f"{D['key']},{key_cell},{D['type']},{type_cell}")
            return (f'=IF({meas_cell}="Mogas-eq (kbd)",'
                    f"SUMIFS({D['mogas']},{args}),SUMIFS({D['kbd']},{args}))")

        # monthly level grid
        r = self._band(ws, r, 1, 14, "Monthly (kbd) - live", "h_navy")
        r = self._month_header(ws, r)
        lvl_first = r
        for y in DISP_YEARS:
            ws.write_number(r, 1, y, self._yr(y, num=True))
            for j in range(12):
                ws.write_formula(r, 2 + j, sumifs(A1(r, 1), j),
                                 self.f["kbd_p"] if y in PARTIAL else self.f["kbd"])
            ws.write_formula(r, 14, f"=SUM({A1(r,2)}:{A1(r,13)})",
                             self.f["kbd_p"] if y in PARTIAL else self.f["kbd_b"])
            r += 1
        lvl_last = r - 1
        ws.conditional_format(lvl_first, 2, lvl_last, 13,
                              {"type": "2_color_scale", "min_type": "num", "min_value": 0,
                               "min_color": WHITE, "max_color": "#9DC3E6"})
        r += 1
        # monthly YoY% grid
        r = self._band(ws, r, 1, 13, "YoY % Change by Month - live (this month vs same month, prior year)", "h_red")
        r = self._month_header(ws, r, total=False)
        yoy_first = r
        for i, y in enumerate(DISP_YEARS):
            ws.write_number(r, 1, y, self._yr(y, num=True))
            for j in range(12):
                if i == 0:
                    ws.write(r, 2 + j, "-", self.f["na"])
                else:
                    cur, prev = A1(lvl_first + i, 2 + j), A1(lvl_first + i - 1, 2 + j)
                    ws.write_formula(r, 2 + j, f"=IF({prev}=0,\"n/a\",({cur}-{prev})/{prev})",
                                     self.f["pct_p"] if y in PARTIAL else self.f["pcts"])
            r += 1
        ws.conditional_format(yoy_first, 2, r - 1, 13,
                              {"type": "3_color_scale", "min_color": "#63BE7B",
                               "mid_color": WHITE, "max_color": "#F8696B"})
        # KPI formula now that grid exists (selected 2025 total)
        row2025 = lvl_first + DISP_YEARS.index(2025)
        ws.write_formula(kpi_row, 13, f"={A1(row2025,14)}", self.f["calc_b"])

        # live line chart of recent years
        chart = self.wb.add_chart({"type": "line"})
        cmap = {2022: GRAY, 2023: GREEN, 2024: BLUE, 2025: RED, 2026: GOLD, 2027: PURPLE}
        for y in (2022, 2023, 2024, 2025, 2026):
            rr = lvl_first + DISP_YEARS.index(y)
            chart.add_series({"name": str(y),
                              "categories": [SHEET, lvl_first - 1, 2, lvl_first - 1, 13],
                              "values": [SHEET, rr, 2, rr, 13],
                              "line": {"color": cmap[y], "width": 2.0,
                                       "dash_type": "dash" if y == 2026 else "solid"}})
        chart.set_title({"name": "Selected slice - monthly (kbd)"})
        chart.set_y_axis({"name": "kbd"})
        chart.set_legend({"position": "bottom"})
        chart.set_size({"width": 820, "height": 300})
        chart.set_chartarea({"border": {"none": True}})
        ws.insert_chart(r + 1, 1, chart)
        return r + 17

    def explorer(self):
        ws = self.wb.add_worksheet("Explorer")
        self._setup(ws, "#0F6FC6")
        self._chrome(ws, 14)
        ws.set_column("A:A", 2)
        ws.set_column("B:B", 9)
        ws.set_column("C:N", 7.2)
        ws.set_column("O:O", 9)
        ws.set_column("J:J", 13)
        ws.merge_range("B2:O2", "Explorer - Interactive Monthly View", self.f["title"])
        ws.merge_range("B3:O3",
                       "Pick a PADD or unit and a type from the yellow dropdowns - the grids, the "
                       "month-by-month YoY% and the chart all recompute live.", self.f["subtitle"])
        r = 5
        r = self._explorer_panel(ws, r, "padd", "View by PADD", "h_green", "Total US")
        r += 1
        r = self._explorer_panel(ws, r, "unit", "View by Unit Category", "h_gold", "All Units")
        ws.freeze_panes(5, 2)

    # ===================================================================== COVER
    def cover(self):
        ws = self.wb.add_worksheet("Cover")
        self._setup(ws, NAVY, landscape=False)
        ws.set_column("A:A", 2)
        ws.set_column("B:B", 20)
        ws.set_column("C:H", 13)
        d = self.ctx["diag"]
        y0, y1 = d["years"]
        ws.merge_range("B2:H2", "REFINERY OUTAGE ANALYTICS", self.f["title"])
        ws.merge_range("B3:H3",
                       "Capacity offline (kbd) - planned & unplanned - interactive model & 2027 scenario",
                       self.f["subtitle"])
        rows = [
            ("Primary metric", "CAP_OFFLINE_ADJUSTED_KBD - capacity offline, thousand bbl/day, all units."),
            ("Data vintage", f"{d['rows']:,} rows | {y0}-{y1} | {d['events_distinct']:,} distinct outages | PADD 100% resolved."),
            ("Outage type", "Binary {PLANNED, UNPLANNED}; UNKNOWN folded into UNPLANNED (desk rule)."),
            ("Mogas overlay", "Mogas-equivalent = capacity x unit yield (secondary; on Units & Refineries)."),
        ]
        r = 4
        for lab, val in rows:
            ws.write(r, 1, lab, self.f["note_b"])
            ws.merge_range(r, 2, r, 7, val, self.f["note"]); r += 1
        r += 1
        r = self._band(ws, r, 1, 7, "This Week's Reads (auto-generated from the data)", "h_green")
        for line in self.ctx["top_movers"]:
            ws.merge_range(r, 1, r, 7, "•  " + line, self.f["note"]); ws.set_row(r, 22); r += 1
        r += 1
        r = self._band(ws, r, 1, 7, "Read-Before-Use Caveats", "h_red")
        for c in ["2026 & 2027 are PARTIAL / special - shown grey italic and never presented as final.",
                  "2027 is PLANNED-ONLY; unplanned-2027 is a MODELED scenario, not an actual.",
                  "Guardrail: any Plan+Unplanned / Unplanned comparison vs 2027 shows n/a; only Planned is valid.",
                  "2020-2021 (COVID / Winter Storm Uri) are excluded from forecast baselines by default."]:
            ws.merge_range(r, 1, r, 7, "•  " + c, self.f["red_note"]); ws.set_row(r, 24); r += 1
        r += 1
        r = self._band(ws, r, 1, 7, "Contents", "h_navy")
        toc = [
            ("Dashboard", "KPIs and headline charts"),
            ("Explorer", "Interactive: dropdown PADD/unit/type -> live monthly grid + YoY% + chart"),
            ("Trends", "Annual table, comparisons, monthly matrices + month-by-month YoY%"),
            ("PADD", "Per-PADD combo charts and PADD x year detail"),
            ("Units & Refineries", "Unit categories (share/YoY%), top refineries, operators, scatter, mogas"),
            ("Events & TAs", "Back-to-back FCC clusters and the 2026 turnaround schedule"),
            ("Model", "Live 2027 scenario (dropdowns) + sensitivity heatmap & tornado"),
        ]
        for name, desc in toc:
            ws.write_url(r, 1, f"internal:'{name}'!A1", self.f["link"], name)
            ws.merge_range(r, 2, r, 7, desc, self.f["note"]); r += 1
        r += 1
        r = self._band(ws, r, 1, 7, "Color Key", "h_navy")
        for color, name, desc in [("#0000FF", "Blue font", "Hard-coded input you change"),
                                  ("#000000", "Black font", "Formula / calculation"),
                                  ("#008000", "Green font", "Cross-sheet link"),
                                  (YELLOW, "Yellow fill", "Key assumption / dropdown")]:
            ws.write(r, 1, name, self.wb.add_format(
                {"font_name": "Arial", "bold": True,
                 "font_color": color if color != YELLOW else "#000000",
                 "bg_color": YELLOW if color == YELLOW else WHITE}))
            ws.merge_range(r, 2, r, 7, desc, self.f["note"]); r += 1

    # ================================================================= DASHBOARD
    def dashboard(self):
        ws = self.wb.add_worksheet("Dashboard")
        self._setup(ws, RED)
        self._chrome(ws, 16)
        ws.set_column("A:A", 2)
        ws.set_column("B:Q", 9.5)
        s = self.ctx["summary"]
        ly = max(y for y in s.index if y not in PARTIAL and s.loc[y, "Unplanned"] > 0)
        ws.merge_range("B2:Q2", "Outage Dashboard", self.f["title"])
        ws.merge_range("B3:Q3", f"Headline metrics - latest full year {ly}. Use Explorer to slice by PADD/unit.",
                       self.f["subtitle"])
        padd_un = self.ctx["padd_unplanned"]
        top_padd = padd_un[ly].idxmax()
        tiles = [("Total Offline", float(s.loc[ly, "Total"]), "kbd", f"FY{ly}"),
                 ("Unplanned", float(s.loc[ly, "Unplanned"]), "kbd", f"FY{ly}"),
                 ("Unplanned %", float(s.loc[ly, "Unpl%"]), "pct", "of total"),
                 ("Distinct Outages", int(s.loc[ly, "Events"]), "kbd", f"FY{ly}"),
                 ("Top PADD", top_padd, "txt", "by unplanned")]
        col = 1
        for lab, val, kind, sub in tiles:
            ws.merge_range(4, col, 4, col + 2, lab, self.f["kpi_lab"])
            fmt = {"pct": "kpi_pct", "txt": "kpi_txt"}.get(kind, "kpi_num")
            ws.merge_range(5, col, 6, col + 2, val, self.f[fmt])
            ws.merge_range(7, col, 7, col + 2, sub, self.f["kpi_sub"])
            col += 3
        ws.set_row(5, 22); ws.set_row(6, 10)

        # data block (compact, just below the charts area)
        d0 = 26
        ws.write(d0 - 1, 1, "Chart data (kbd)", self.f["note_b"])
        years = [y for y in s.index if 2016 <= y <= 2027]
        ws.write(d0, 1, "Year", self.f["colhdr_l"])
        ws.write(d0, 2, "Planned", self.f["colhdr"]); ws.write(d0, 3, "Unplanned", self.f["colhdr"])
        for i, y in enumerate(years):
            ws.write_number(d0 + 1 + i, 1, y, self.f["yr"])
            ws.write_number(d0 + 1 + i, 2, float(s.loc[y, "Planned"]), self.f["kbd"])
            ws.write_number(d0 + 1 + i, 3, float(s.loc[y, "Unplanned"]), self.f["kbd"])
        n = len(years)
        stack = self.wb.add_chart({"type": "column", "subtype": "stacked"})
        for ci, (nm, c) in enumerate([("Planned", NAVY), ("Unplanned", RED)]):
            stack.add_series({"name": ["Dashboard", d0, 2 + ci],
                              "categories": ["Dashboard", d0 + 1, 1, d0 + n, 1],
                              "values": ["Dashboard", d0 + 1, 2 + ci, d0 + n, 2 + ci],
                              "fill": {"color": c}, "gap": 50})
        stack.set_title({"name": "Capacity Offline by Year (Planned + Unplanned)"})
        stack.set_y_axis({"name": "kbd"}); stack.set_legend({"position": "bottom"})
        stack.set_size({"width": 520, "height": 290}); stack.set_chartarea({"border": {"none": True}})
        ws.insert_chart("B9", stack)

        rec = [y for y in [2022, 2023, 2024, 2025] if y in padd_un.columns]
        p0 = d0 + n + 2
        ws.write(p0, 1, "PADD", self.f["colhdr_l"])
        for j, y in enumerate(rec):
            ws.write_number(p0, 2 + j, y, self.f["yr"])
        for i, p in enumerate(PADDS):
            ws.write(p0 + 1 + i, 1, p, self.f["rowlab"])
            for j, y in enumerate(rec):
                ws.write_number(p0 + 1 + i, 2 + j, float(padd_un.loc[p, y]), self.f["kbd"])
        clu = self.wb.add_chart({"type": "column"})
        for j, y in enumerate(rec):
            clu.add_series({"name": ["Dashboard", p0, 2 + j],
                            "categories": ["Dashboard", p0 + 1, 1, p0 + 5, 1],
                            "values": ["Dashboard", p0 + 1, 2 + j, p0 + 5, 2 + j],
                            "fill": {"color": [BLUE, GOLD, GREEN, RED][j % 4]}})
        clu.set_title({"name": "Unplanned Offline by PADD"})
        clu.set_y_axis({"name": "kbd"}); clu.set_legend({"position": "bottom"})
        clu.set_size({"width": 520, "height": 290}); clu.set_chartarea({"border": {"none": True}})
        ws.insert_chart("J9", clu)

        sh0 = p0 + 8
        for i, p in enumerate(PADDS):
            ws.write(sh0 + i, 1, p, self.f["rowlab"])
            ws.write_number(sh0 + i, 2, float(padd_un.loc[p, ly]), self.f["kbd"])
        donut = self.wb.add_chart({"type": "doughnut"})
        donut.add_series({"name": f"Unplanned share {ly}",
                          "categories": ["Dashboard", sh0, 1, sh0 + 4, 1],
                          "values": ["Dashboard", sh0, 2, sh0 + 4, 2],
                          "points": [{"fill": {"color": c}} for c in [NAVY, BLUE, GOLD, GREEN, ORANGE]],
                          "data_labels": {"percentage": True, "font": {"color": WHITE, "bold": True, "size": 9}}})
        donut.set_title({"name": f"Unplanned Share by PADD ({ly})"})
        donut.set_legend({"position": "right"}); donut.set_hole_size(55)
        donut.set_size({"width": 430, "height": 290}); donut.set_chartarea({"border": {"none": True}})
        ws.insert_chart("B25", donut)
        # seasonality band-ish: unplanned by month recent years
        mu = self.ctx["monthly_unplanned"]
        m0 = sh0
        ws.write(m0, 8, "Month", self.f["colhdr_l"])
        for j, mm in enumerate(MONTHS):
            ws.write(m0, 9 + j, mm, self.f["colhdr"])
        for i, y in enumerate([2024, 2025, 2026]):
            ws.write_number(m0 + 1 + i, 8, y, self.f["yr"])
            for j, mm in enumerate(MONTHS):
                ws.write_number(m0 + 1 + i, 9 + j, float(mu.loc[y, mm]) if y in mu.index else 0.0, self.f["kbd"])
        ln = self.wb.add_chart({"type": "line"})
        for i, (y, c) in enumerate([(2024, BLUE), (2025, RED), (2026, GOLD)]):
            ln.add_series({"name": ["Dashboard", m0 + 1 + i, 8],
                           "categories": ["Dashboard", m0, 9, m0, 20],
                           "values": ["Dashboard", m0 + 1 + i, 9, m0 + 1 + i, 20],
                           "line": {"color": c, "width": 2.0, "dash_type": "dash" if y == 2026 else "solid"}})
        ln.set_title({"name": "Unplanned Seasonality (kbd by month)"})
        ln.set_legend({"position": "bottom"})
        ln.set_size({"width": 430, "height": 290}); ln.set_chartarea({"border": {"none": True}})
        ws.insert_chart("J25", ln)
        ws.freeze_panes(8, 0)

    # ===================================================================== TRENDS
    def trends(self):
        ws = self.wb.add_worksheet("Trends")
        self._setup(ws, BLUE)
        self._chrome(ws, 15)
        ws.set_column("A:A", 2)
        ws.set_column("B:B", 9)
        ws.set_column("C:N", 7.2)
        ws.set_column("O:O", 9)
        ws.merge_range("B2:O2", "Trends - Annual & Monthly", self.f["title"])
        ws.merge_range("B3:O3",
                       "Annual summary, targeted comparisons, monthly matrices and month-by-month YoY%.",
                       self.f["subtitle"])
        r = 5
        r = self._annual_block(ws, r)
        r = self._compare_block(ws, r)
        r, firsts = self._monthly_levels(ws, r)
        r = self._monthly_yoy(ws, r, firsts)
        # seasonality chart from the unplanned monthly block
        uf = firsts["Unplanned"]
        line = self.wb.add_chart({"type": "line"})
        cmap = {2022: GRAY, 2023: GREEN, 2024: BLUE, 2025: RED, 2026: GOLD}
        for y in (2022, 2023, 2024, 2025, 2026):
            rr = uf + DISP_YEARS.index(y)
            line.add_series({"name": str(y),
                             "categories": ["Trends", uf - 1, 2, uf - 1, 13],
                             "values": ["Trends", rr, 2, rr, 13],
                             "line": {"color": cmap[y], "width": 2.0,
                                      "dash_type": "dash" if y == 2026 else "solid"}})
        line.set_title({"name": "Unplanned Offline Seasonality (kbd by month)"})
        line.set_y_axis({"name": "kbd"}); line.set_legend({"position": "bottom"})
        line.set_size({"width": 820, "height": 300}); line.set_chartarea({"border": {"none": True}})
        ws.insert_chart(r + 1, 1, line)
        ws.freeze_panes(6, 2)

    # ===================================================================== PADD
    def padd(self):
        ws = self.wb.add_worksheet("PADD")
        self._setup(ws, GREEN)
        self._chrome(ws, 13)
        ws.set_column("A:A", 2)
        ws.set_column("B:B", 16)
        ws.set_column("C:N", 7.5)
        ws.merge_range("B2:N2", "PADD Detail & Charts", self.f["title"])
        ws.merge_range("B3:N3",
                       "Per-PADD combo charts (2026 plan+unplanned bars vs prior-year totals & 2027 plan) "
                       "and PADD x year matrices.", self.f["subtitle"])
        pm = self.ctx["padd_month"]
        CUR = 2026
        r = 5
        for p in PADDS:
            r = self._band(ws, r, 1, 13, f"{p} - {CUR} plan+unplanned vs 2023-25 totals & 2027 plan", "h_green")
            hdr = r
            ws.write(r, 1, "Series", self.f["colhdr_l"])
            for j, m in enumerate(MONTHS):
                ws.write(r, 2 + j, m, self.f["colhdr"])
            r += 1

            def row_of(mat, yr):
                return [float(mat.loc[yr, m]) if yr in mat.index else 0.0 for m in MONTHS]
            series = [(f"{CUR} Planned", row_of(pm[p]["planned"], CUR)),
                      (f"{CUR} Unplanned", row_of(pm[p]["unplanned"], CUR)),
                      ("2025 Total", row_of(pm[p]["total"], 2025)),
                      ("2024 Total", row_of(pm[p]["total"], 2024)),
                      ("2023 Total", row_of(pm[p]["total"], 2023)),
                      ("2027 Planned", row_of(pm[p]["planned"], 2027))]
            data_first = r
            for name, vals in series:
                ws.write(r, 1, name, self.f["rowlab"])
                for j, v in enumerate(vals):
                    ws.write_number(r, 2 + j, v, self.f["kbd"])
                r += 1
            col = self.wb.add_chart({"type": "column", "subtype": "stacked"})
            col.add_series({"name": ["PADD", data_first, 1], "categories": ["PADD", hdr, 2, hdr, 13],
                            "values": ["PADD", data_first, 2, data_first, 13], "fill": {"color": GOLD}})
            col.add_series({"name": ["PADD", data_first + 1, 1], "categories": ["PADD", hdr, 2, hdr, 13],
                            "values": ["PADD", data_first + 1, 2, data_first + 1, 13], "fill": {"color": ORANGE}})
            ln = self.wb.add_chart({"type": "line"})
            for k, (yr, c) in enumerate([(2025, RED), (2024, BLUE), (2023, GRAY), (2027, GREEN)]):
                rr = data_first + 2 + k
                ln.add_series({"name": ["PADD", rr, 1], "categories": ["PADD", hdr, 2, hdr, 13],
                               "values": ["PADD", rr, 2, rr, 13], "line": {"color": c, "width": 2.25}})
            col.combine(ln)
            col.set_title({"name": f"{p} Planned & Unplanned Offline (kbd)"})
            col.set_y_axis({"name": "kbd"}); col.set_legend({"position": "top"})
            col.set_size({"width": 840, "height": 320}); col.set_chartarea({"border": {"none": True}})
            ws.insert_chart(r, 1, col)
            r += 17

        # PADD x year matrices (Total / Unplanned / Planned) + YoY%
        for title, mat, band in [("Total Offline", self.ctx["padd_total"], "h_navy"),
                                 ("Unplanned", self.ctx["padd_unplanned"], "h_red"),
                                 ("Planned", self.ctx["padd_planned"], "h_blue")]:
            years = [y for y in mat.columns if 2018 <= y <= 2027]
            r = self._band(ws, r, 1, 1 + len(years), f"{title} - PADD x Year (kbd)", band)
            ws.write(r, 1, "PADD", self.f["colhdr_l"])
            for j, y in enumerate(years):
                ws.write_number(r, 2 + j, y, self._yr(y));
            r += 1
            blk = r
            for p in PADDS:
                ws.write(r, 1, p, self.f["rowlab"])
                for j, y in enumerate(years):
                    ws.write_number(r, 2 + j, float(mat.loc[p, y]),
                                    self.f["kbd_p"] if y in PARTIAL else self.f["kbd"])
                r += 1
            ws.write(r, 1, "Total US", self.f["rowlab_b"])
            for j, y in enumerate(years):
                ws.write_formula(r, 2 + j, f"=SUM({A1(blk,2+j)}:{A1(r-1,2+j)})",
                                 self.f["kbd_p"] if y in PARTIAL else self.f["kbd_b"])
            r += 2
        ws.freeze_panes(6, 2)

    # ====================================================== UNITS & REFINERIES
    def units_refineries(self):
        ws = self.wb.add_worksheet("Units & Refineries")
        self._setup(ws, GOLD)
        self._chrome(ws, 13)
        ws.set_column("A:A", 2)
        ws.set_column("B:B", 24)
        ws.set_column("C:C", 26)
        ws.set_column("D:D", 18)
        ws.set_column("E:N", 9)
        ws.merge_range("B2:N2", "Units & Refineries", self.f["title"])
        ws.merge_range("B3:N3",
                       "Unit-category mix (share & YoY%), top refineries and operators, event scatter, "
                       "and the mogas overlay.", self.f["subtitle"])
        # unit matrix with share + YoY columns
        mat = self.ctx["unit_total"]
        years = [y for y in mat.columns if 2020 <= y <= 2027]
        i25 = years.index(2025) if 2025 in years else len(years) - 1
        i24 = years.index(2024) if 2024 in years else i25 - 1
        tot_col, share_col, yoy_col = 2 + len(years), 3 + len(years), 4 + len(years)
        r = 5
        r = self._band(ws, r, 1, yoy_col, "Capacity Offline by Unit Category (kbd) + share & YoY%")
        ws.write(r, 1, "Unit Category", self.f["colhdr_l"])
        for j, y in enumerate(years):
            ws.write_number(r, 2 + j, y, self._yr(y))
        ws.write(r, tot_col, "Total", self.f["colhdr"])
        ws.write(r, share_col, "% '25", self.f["colhdr"])
        ws.write(r, yoy_col, "YoY%", self.f["colhdr"])
        r += 1
        first = r
        nU = len(mat.index)
        for i, u in enumerate(mat.index):
            sh = i % 2 == 1
            ws.write(r, 1, str(u).title(), self.f["rowlab"])
            for j, y in enumerate(years):
                ws.write_number(r, 2 + j, float(mat.loc[u, y]), self._kf(y, sh))
            ws.write_formula(r, tot_col, f"=SUM({A1(r,2)}:{A1(r,1+len(years))})", self.f["kbd_b"])
            c25 = A1(r, 2 + i25)
            colsum = f"SUM({A1(first,2+i25)}:{A1(first+nU-1,2+i25)})"
            ws.write_formula(r, share_col, f"=IF({colsum}=0,0,{c25}/{colsum})", self.f["pct"])
            ws.write_formula(r, yoy_col, f"=IF({A1(r,2+i24)}=0,\"n/a\",({c25}-{A1(r,2+i24)})/{A1(r,2+i24)})",
                             self.f["pcts"])
            r += 1
        last = r - 1
        ws.conditional_format(first, tot_col, last, tot_col, {"type": "data_bar", "bar_color": GOLD, "bar_solid": True})
        topn = min(10, nU)
        bar = self.wb.add_chart({"type": "bar"})
        bar.add_series({"name": "Total offline (kbd)",
                        "categories": ["Units & Refineries", first, 1, first + topn - 1, 1],
                        "values": ["Units & Refineries", first, tot_col, first + topn - 1, tot_col],
                        "fill": {"color": GOLD}, "data_labels": {"value": True, "num_format": "#,##0"}})
        bar.set_title({"name": f"Top {topn} Unit Categories"}); bar.set_legend({"none": True})
        bar.set_size({"width": 470, "height": 300}); bar.set_chartarea({"border": {"none": True}})
        bar.reverse_series_order = True
        ws.insert_chart(first, yoy_col + 2, bar)
        r = max(r + 1, first + topn + 1)

        # --- Naphtha / octane complex (reforming + isomerization + aromatics/BTX) ---
        nap = self.ctx["naphtha"]
        na = nap["annual"]
        nyears = [y for y in na.index if 2018 <= y <= 2027]
        r = self._band(ws, r, 1, 8, "Naphtha / Octane Complex Offline (kbd) - reforming + isomerization + aromatics/BTX", "h_purple")
        ws.merge_range(r, 1, r, 8,
                       "These naphtha-fed units set gasoline octane; offline here squeezes octane/blending "
                       "even if crude runs hold - a read CDU-only trackers miss.", self.f["subtitle"])
        r += 1
        for j, h in enumerate(["Year", "Planned", "Unplanned", "Total"]):
            ws.write(r, 1 + j, h, self.f["colhdr_l"] if j == 0 else self.f["colhdr"])
        # by-PADD mini header
        ws.write(r, 6, "PADD (2025)", self.f["colhdr_l"]); ws.write(r, 7, "kbd", self.f["colhdr"])
        r += 1
        nf = r
        for y in nyears:
            kf = self.f["kbd_p"] if y in PARTIAL else self.f["kbd"]
            ws.write_number(r, 1, y, self._yr(y))
            ws.write_number(r, 2, float(na.loc[y, "Planned"]), kf)
            ws.write_number(r, 3, float(na.loc[y, "Unplanned"]), kf)
            ws.write_formula(r, 4, f"={A1(r,2)}+{A1(r,3)}", self.f["kbd_p"] if y in PARTIAL else self.f["kbd_b"])
            r += 1
        # by-PADD list alongside
        bp = nap["by_padd"]
        for k, p in enumerate(PADDS):
            rr = nf + k
            if rr < r:
                ws.write(rr, 6, p, self.f["rowlab"])
                ws.write_number(rr, 7, float(bp.loc[p, 2025]) if (p in bp.index and 2025 in bp.columns) else 0.0, self.f["kbd"])
        nch = self.wb.add_chart({"type": "column", "subtype": "stacked"})
        for ci, (nm, c) in enumerate([("Planned", NAVY), ("Unplanned", PURPLE)]):
            nch.add_series({"name": nm, "categories": ["Units & Refineries", nf, 1, r - 1, 1],
                            "values": ["Units & Refineries", nf, 2 + ci, r - 1, 2 + ci], "fill": {"color": c}})
        nch.set_title({"name": "Naphtha/Octane Complex Offline (kbd)"}); nch.set_legend({"position": "bottom"})
        nch.set_size({"width": 470, "height": 270}); nch.set_chartarea({"border": {"none": True}})
        ws.insert_chart(nf, 9, nch)
        r = max(r + 1, nf + 13)

        # top refineries
        pl = self.ctx["plants"]
        r = self._band(ws, r, 1, 7, "Top 15 Refineries by Capacity Offline (kbd)")
        for j, h in enumerate(["Refinery", "Operator", "PADD", "Total", "Planned", "Unplanned", "Events"]):
            ws.write(r, 1 + j, h, self.f["colhdr_l"] if j in (0, 1) else self.f["colhdr"])
        r += 1
        rf = r
        for _, row in pl.iterrows():
            ws.write(r, 1, str(row["plant"]), self.f["rowlab"])
            ws.write(r, 2, short_op(row["operator"], 24), self.f["rowlab"])
            ws.write(r, 3, str(row["padd"]), self.f["rowlab"])
            ws.write_number(r, 4, float(row["total"]), self.f["kbd"])
            ws.write_number(r, 5, float(row["planned"]), self.f["kbd"])
            ws.write_number(r, 6, float(row["unplanned"]), self.f["kbd"])
            ws.write_number(r, 7, int(row["events"]), self.f["kbd"])
            r += 1
        ws.conditional_format(rf, 4, r - 1, 4, {"type": "data_bar", "bar_color": PURPLE, "bar_solid": True})
        ws.autofilter(rf - 1, 1, r - 1, 7)        # sortable/filterable for the floor
        r += 1

        # operators x year
        ops = self.ctx["operators"]
        oyears = [y for y in ops.columns if 2020 <= y <= 2027]
        r = self._band(ws, r, 1, 1 + len(oyears), "Top 10 Operators x Year (kbd)")
        ws.write(r, 1, "Operator", self.f["colhdr_l"])
        for j, y in enumerate(oyears):
            ws.write_number(r, 2 + j, y, self._yr(y))
        r += 1
        for i, op in enumerate(ops.index):
            sh = i % 2 == 1
            ws.write(r, 1, short_op(op, 22), self.f["rowlab"])
            for j, y in enumerate(oyears):
                ws.write_number(r, 2 + j, float(ops.loc[op, y]), self._kf(y, sh))
            r += 1
        r += 1

        # scatter + mogas side by side (scatter left as a small block, mogas table)
        sc = self.ctx["scatter"]
        r = self._band(ws, r, 1, 6, "Unplanned Events: Capacity vs Duration (2023-25) | Mogas Overlay")
        ws.write(r, 1, "Duration (days)", self.f["colhdr"])
        ws.write(r, 2, "Capacity (kbd)", self.f["colhdr"])
        # mogas yield map header
        ws.write(r, 4, "Bucket", self.f["colhdr_l"]); ws.write(r, 5, "Factor", self.f["colhdr"])
        ws.write(r, 6, "Unit categories", self.f["colhdr_l"])
        r += 1
        sfirst = r
        for _, row in sc.iterrows():
            ws.write_number(r, 1, float(row["duration"]), self.f["kbd"])
            ws.write_number(r, 2, float(row["cap_kbd"]), self.f["kbd"])
            r += 1
        slast = r - 1
        # mogas yield map rows (written alongside, starting at sfirst)
        ws.set_column("G:G", 34)
        for k, (bucket, factor, cats) in enumerate(self.ctx["mogas_yield_map"]):
            rr = sfirst + k
            ws.write(rr, 4, bucket, self.f["rowlab_b"])
            ws.write_number(rr, 5, factor, self.wb.add_format({"font_name": "Arial", "num_format": "0.000", "align": "center"}))
            ws.write(rr, 6, ", ".join(c.title() for c in cats) or "-", self.f["rowlab"])
        scat = self.wb.add_chart({"type": "scatter"})
        scat.add_series({"categories": ["Units & Refineries", sfirst, 1, slast, 1],
                         "values": ["Units & Refineries", sfirst, 2, slast, 2],
                         "marker": {"type": "circle", "size": 5, "fill": {"color": PURPLE}, "border": {"none": True}}})
        scat.set_title({"name": "Unplanned Events - Capacity vs Duration"})
        scat.set_x_axis({"name": "Duration (days)", "min": 0}); scat.set_y_axis({"name": "kbd", "min": 0})
        scat.set_legend({"none": True})
        scat.set_size({"width": 430, "height": 290}); scat.set_chartarea({"border": {"none": True}})
        ws.insert_chart(sfirst + len(self.ctx["mogas_yield_map"]) + 1, 4, scat)

        # mogas annual
        rr = sfirst + max(len(self.ctx["mogas_yield_map"]) + 14, slast - sfirst + 2)
        ma = self.ctx["mogas_annual"]
        myears = [y for y in ma.index if 2018 <= y <= 2027]
        rr = self._band(ws, rr, 1, 4, "Mogas-Equivalent Offline by Year (kbd)", "h_green")
        for j, h in enumerate(["Year", "Planned", "Unplanned", "Total"]):
            ws.write(rr, 1 + j, h, self.f["colhdr_l"] if j == 0 else self.f["colhdr"])
        rr += 1
        mf = rr
        for y in myears:
            kf = self.f["kbd_p"] if y in PARTIAL else self.f["kbd"]
            ws.write_number(rr, 1, y, self._yr(y))
            ws.write_number(rr, 2, float(ma.loc[y, "Planned"]), kf)
            ws.write_number(rr, 3, float(ma.loc[y, "Unplanned"]), kf)
            ws.write_formula(rr, 4, f"={A1(rr,2)}+{A1(rr,3)}", self.f["kbd_p"] if y in PARTIAL else self.f["kbd_b"])
            rr += 1
        mch = self.wb.add_chart({"type": "column", "subtype": "stacked"})
        for ci, (nm, c) in enumerate([("Planned", NAVY), ("Unplanned", GREEN)]):
            mch.add_series({"name": nm, "categories": ["Units & Refineries", mf, 1, rr - 1, 1],
                            "values": ["Units & Refineries", mf, 2 + ci, rr - 1, 2 + ci], "fill": {"color": c}})
        mch.set_title({"name": "Mogas-Equivalent Offline (kbd)"}); mch.set_legend({"position": "bottom"})
        mch.set_size({"width": 470, "height": 290}); mch.set_chartarea({"border": {"none": True}})
        ws.insert_chart(mf, 6, mch)
        ws.freeze_panes(6, 0)

    # ===================================================================== EVENTS
    def events_tas(self):
        ws = self.wb.add_worksheet("Events & TAs")
        self._setup(ws, "#843C0C")
        self._chrome(ws, 14)
        ws.set_column("A:A", 2)
        ws.set_column("B:B", 27)
        ws.set_column("C:C", 20)
        ws.set_column("D:N", 7)
        ws.set_column("O:O", 8)
        ws.merge_range("B2:O2", "Events & Turnarounds", self.f["title"])
        ws.merge_range("B3:O3",
                       "Back-to-back FCC clusters external trackers miss, and the 2026 planned "
                       "turnaround schedule by PADD.", self.f["subtitle"])
        grid = self.ctx["fcc_grid"]
        r = 5
        r = self._band(ws, r, 1, 14, "ExxonMobil FCC Offline (kbd) - Plant x Month, 2022-2026  (adjacent shaded cells = a run)", "h_orange")
        ws.write(r, 1, "Plant (Year)", self.f["colhdr_l"])
        for j, m in enumerate(MONTHS):
            ws.write(r, 2 + j, m, self.f["colhdr"])
        r += 1
        gfirst = r
        for (plant, year), row in sorted(grid.items(), key=lambda kv: (kv[0][0], kv[0][1])):
            ws.write(r, 1, f"{plant.replace(' Refinery','')} {year}", self.f["rowlab"])
            for j in range(12):
                ws.write_number(r, 2 + j, row[j], self.wb.add_format(
                    {"font_name": "Arial", "num_format": '#,##0;;""', "align": "center",
                     "border": 1, "border_color": "#E0E0E0"}))
            r += 1
        ws.conditional_format(gfirst, 2, r - 1, 13, {"type": "2_color_scale", "min_type": "num",
                              "min_value": 0, "min_color": WHITE, "max_color": ORANGE})
        r += 1
        # Exxon FCC table + all-operator table side considerations: stack tight
        r = self._band(ws, r, 1, 8, "ExxonMobil FCC Back-to-Back Runs (>=3 months, ex-2020)", "h_orange")
        for j, h in enumerate(["Refinery", "PADD", "Year", "Span", "Mo", "kbd", "Unpl%"]):
            ws.write(r, 1 + j, h, self.f["colhdr_l"] if j == 0 else self.f["colhdr"])
        r += 1
        ef = r
        for c in self.ctx["fcc_exxon"]:
            ws.write(r, 1, c["plant"], self.f["rowlab"]); ws.write(r, 2, c["padd"], self.f["rowlab"])
            ws.write_number(r, 3, c["year"], self._yr(c["year"])); ws.write(r, 4, c["span"], self.f["rowlab"])
            ws.write_number(r, 5, c["n"], self.f["kbd"]); ws.write_number(r, 6, c["kbd"], self.f["kbd"])
            ws.write_number(r, 7, c["unpl_share"], self.f["pct"])
            r += 1
        ws.conditional_format(ef, 6, r - 1, 6, {"type": "data_bar", "bar_color": ORANGE, "bar_solid": True})
        r += 1

        # TA schedule by PADD (compact, top events per PADD)
        dfmt = self.f["date"]; dfmt_sh = self.f["date_sh"]
        for p in ["PADD 3", "PADD 2", "PADD 5"]:
            ta = self.ctx["ta_schedule"][p]
            if not len(ta):
                continue
            r = self._band(ws, r, 1, 8, f"{p} - 2026 Planned TAs (top {min(12,len(ta))} by size)", "h_navy")
            for j, h in enumerate(["Operator", "Refinery", "Unit", "Offline", "% PADD", "Start", "End"]):
                ws.write(r, 1 + j, h, self.f["colhdr_l"] if j < 3 else self.f["colhdr"])
            r += 1
            for i, (_, row) in enumerate(ta.head(12).iterrows()):
                sh = i % 2 == 1
                lf = self.wb.add_format({"font_name": "Arial", "align": "left", "indent": 1, "bg_color": LT_GRAY}) if sh else self.f["rowlab"]
                ws.write(r, 1, short_op(row["operator"]), lf)
                ws.write(r, 2, str(row["plant"]), lf)
                ws.write(r, 3, str(row["unit_cat"]).title(), lf)
                ws.write_number(r, 4, float(row["kbd"]), self._kf(2025, sh))
                ws.write_number(r, 5, float(row["pct_padd"]), self.f["pct"])
                ws.write_datetime(r, 6, row["start"].to_pydatetime(), dfmt_sh if sh else dfmt)
                ws.write_datetime(r, 7, row["end"].to_pydatetime(), dfmt_sh if sh else dfmt)
                r += 1
            r += 1
        ws.freeze_panes(gfirst, 2)

    # ===================================================================== MODEL
    def model(self):
        ws = self.wb.add_worksheet("Model")
        self._setup(ws, "#C55A11")
        self._chrome(ws, 13)
        ws.set_column("A:A", 2)
        ws.set_column("B:B", 24)
        ws.set_column("C:N", 8)
        SHEET = "Model"
        ws.merge_range("B2:N2", "2027 Scenario & Sensitivity (Live Model)", self.f["title"])
        ws.merge_range("B3:N3",
                       "Edit the yellow inputs - the forecast, per-PADD split, heatmap and tornado all recompute.",
                       self.f["subtitle"])
        windows = list(engine.BASELINE_WINDOWS.keys())
        df = self.ctx["df"]
        r = 5
        r = self._band(ws, r, 1, 4, "Scenario Inputs", "h_orange")
        inrow = {}
        ws.write(r, 1, "Baseline window", self.f["in_lab"]); ws.merge_range(r, 2, r, 4, engine.DEFAULT_WINDOW, self.f["in_val"])
        ws.data_validation(r, 2, r, 2, {"validate": "list", "source": windows}); inrow["window"] = r; r += 1
        ws.write(r, 1, "Production growth %", self.f["in_lab"]); ws.merge_range(r, 2, r, 4, 0.0, self.f["in_pct"]); inrow["growth"] = r; r += 1
        ws.write(r, 1, "Unplanned multiplier", self.f["in_lab"]); ws.merge_range(r, 2, r, 4, 1.0, self.f["in_mult"]); inrow["mult"] = r; r += 1
        ws.write(r, 1, "One-off event (kbd)", self.f["in_lab"]); ws.merge_range(r, 2, r, 4, 0, self.f["in_kbd"]); inrow["oneoff"] = r; r += 1
        ws.write(r, 1, "Stress month", self.f["in_lab"]); ws.merge_range(r, 2, r, 4, "Sep", self.f["in_val"])
        ws.data_validation(r, 2, r, 2, {"validate": "list", "source": MONTHS}); inrow["stress"] = r; r += 1
        c_w = A1(inrow["window"], 2, row_abs=True, col_abs=True)
        c_g = A1(inrow["growth"], 2, row_abs=True, col_abs=True)
        c_m = A1(inrow["mult"], 2, row_abs=True, col_abs=True)
        c_o = A1(inrow["oneoff"], 2, row_abs=True, col_abs=True)
        c_s = A1(inrow["stress"], 2, row_abs=True, col_abs=True)
        r += 1

        r = self._band(ws, r, 1, 13, "Lookup - Avg Unplanned Offline (kbd/month) by Window", "h_navy")
        ws.write(r, 1, "Window", self.f["colhdr_l"])
        for j, m in enumerate(MONTHS):
            ws.write(r, 2 + j, m, self.f["colhdr"])
        r += 1
        pf = r
        for w in windows:
            prof = engine.baseline_profile(df, w)
            ws.write(r, 1, w, self.f["rowlab"])
            for j, m in enumerate(MONTHS):
                ws.write_number(r, 2 + j, float(prof[m]), self.f["kbd"])
            r += 1
        pl = r - 1
        prof_range = xl_range_abs(pf, 2, pl, 13)
        wlab_range = xl_range_abs(pf, 1, pl, 1)
        r += 1

        r = self._band(ws, r, 1, 13, "2027 Unplanned Forecast Cascade (live)", "h_orange")
        ws.write(r, 1, "", self.f["colhdr_l"])
        for j, m in enumerate(MONTHS):
            ws.write(r, 2 + j, m, self.f["colhdr"])
        hdr = r; r += 1
        ws.write(r, 1, "Baseline (window)", self.f["rowlab"]); brow = r
        for j in range(12):
            ws.write_formula(r, 2 + j, f"=INDEX({prof_range},MATCH({c_w},{wlab_range},0),{j+1})", self.f["calc"])
        r += 1
        ws.write(r, 1, "Forecast unplanned", self.f["rowlab_b"]); frow = r
        for j in range(12):
            mc, bc = A1(hdr, 2 + j), A1(brow, 2 + j)
            ws.write_formula(r, 2 + j, f"={bc}*(1+{c_g})*{c_m}+IF({c_s}={mc},{c_o},0)", self.f["calc_b"])
        r += 2

        r = self._band(ws, r, 1, 4, "Scenario Outputs", "h_orange")
        ws.write(r, 1, "Baseline annual (pre-shock)", self.f["rowlab"]); anchor = A1(r, 2, row_abs=True, col_abs=True)
        ws.write_formula(r, 2, f"=SUM({A1(brow,2)}:{A1(brow,13)})", self.f["calc"]); r += 1
        ws.write(r, 1, "2027 Unplanned forecast", self.f["rowlab_b"]); fc_annual = A1(r, 2, row_abs=True, col_abs=True)
        ws.write_formula(r, 2, f"=SUM({A1(frow,2)}:{A1(frow,13)})", self.f["calc_b"]); r += 1
        ws.write(r, 1, "2027 Planned (booked)", self.f["rowlab"])
        planned = float(self.ctx["summary"].loc[2027, "Planned"]) if 2027 in self.ctx["summary"].index else 0.0
        pcell = A1(r, 2, row_abs=True, col_abs=True); ws.write_number(r, 2, planned, self.f["calc_grn"]); r += 1
        ws.write(r, 1, "2027 Implied total", self.f["rowlab_b"])
        ws.write_formula(r, 2, f"={fc_annual}+{pcell}", self.f["calc_b"]); r += 1
        # P25/P50/P90 range band on unplanned (historical window distribution)
        sb = self.ctx["scenario_bands"]
        ws.write(r, 1, "Unplanned range (P25 / P50 / P90)", self.f["rowlab"])
        for k, key in enumerate(["p25", "p50", "p90"]):
            ws.write_number(r, 2 + k, float(sb[key]), self.f["calc"])
        r += 1
        ws.write(r, 1, "Implied total (P50 / P90)", self.f["rowlab"])
        ws.write_formula(r, 2, f"={A1(r-1,3)}+{pcell}", self.f["calc"])
        ws.write_formula(r, 3, f"={A1(r-1,4)}+{pcell}", self.f["calc"]); r += 2

        # per-PADD scenario (live)
        sp = self.ctx["scenario_padd"]
        r = self._band(ws, r, 1, 4, "2027 Scenario by PADD (each PADD's own seasonality, live)", "h_green")
        for j, h in enumerate(["PADD", "Baseline", "Scenario", "Share"]):
            ws.write(r, 1 + j, h, self.f["colhdr_l"] if j == 0 else self.f["colhdr"])
        r += 1
        spf = r
        for p in PADDS:
            ws.write(r, 1, p, self.f["rowlab"]); ws.write_number(r, 2, float(sp[p]["baseline_annual"]), self.f["calc"])
            ws.write_formula(r, 3, f"={A1(r,2)}*(1+{c_g})*{c_m}", self.f["calc_b"]); r += 1
        spl = r - 1
        ws.write(r, 1, "Total US", self.f["rowlab_b"])
        ws.write_formula(r, 2, f"=SUM({A1(spf,2)}:{A1(spl,2)})", self.f["calc_b"])
        ws.write_formula(r, 3, f"=SUM({A1(spf,3)}:{A1(spl,3)})", self.f["calc_b"])
        tot_scn = A1(r, 3, row_abs=True, col_abs=True)
        for i in range(len(PADDS)):
            ws.write_formula(spf + i, 4, f"=IF({tot_scn}=0,0,{A1(spf+i,3)}/{tot_scn})", self.f["pct_g"])
        r += 2

        # forecast vs actuals chart data
        r = self._band(ws, r, 1, 13, "Forecast vs Actuals (chart data)", "h_navy")
        ws.write(r, 1, "Series", self.f["colhdr_l"])
        for j, m in enumerate(MONTHS):
            ws.write(r, 2 + j, m, self.f["colhdr"])
        chdr = r; r += 1
        ws.write(r, 1, "2027 Scenario", self.f["rowlab"])
        for j in range(12):
            ws.write_formula(r, 2 + j, f"={A1(frow,2+j)}", self.f["calc"])
        scrow = r; r += 1
        mp = self.ctx["monthly_planned"]; mu = self.ctx["monthly_unplanned"]
        ws.write(r, 1, "2027 Planned", self.f["rowlab"])
        for j, m in enumerate(MONTHS):
            ws.write_number(r, 2 + j, float(mp.loc[2027, m]) if 2027 in mp.index else 0.0, self.f["calc"])
        plrow = r; r += 1
        for yr in (2025, 2024):
            ws.write(r, 1, f"{yr} Unplanned", self.f["rowlab"])
            for j, m in enumerate(MONTHS):
                ws.write_number(r, 2 + j, float(mu.loc[yr, m]) if yr in mu.index else 0.0, self.f["calc"])
            r += 1
        chart = self.wb.add_chart({"type": "line"})
        for rr, c, wd, dash in [(scrow, RED, 3.0, "solid"), (plrow, GOLD, 2.0, "dash"),
                                (plrow + 1, BLUE, 1.5, "solid"), (plrow + 2, GRAY, 1.5, "solid")]:
            chart.add_series({"name": [SHEET, rr, 1], "categories": [SHEET, chdr, 2, chdr, 13],
                              "values": [SHEET, rr, 2, rr, 13], "line": {"color": c, "width": wd, "dash_type": dash}})
        chart.set_title({"name": "2027 Scenario vs Planned vs Recent Actual Unplanned (kbd)"})
        chart.set_y_axis({"name": "kbd"}); chart.set_legend({"position": "bottom"})
        chart.set_size({"width": 640, "height": 300}); chart.set_chartarea({"border": {"none": True}})
        ws.insert_chart(chdr, 15, chart)
        r += 1

        # --- sensitivity heatmap ---
        growths = [-0.10, -0.05, 0.0, 0.05, 0.10, 0.15]
        mults = [0.7, 0.85, 1.0, 1.15, 1.3, 1.5]
        r = self._band(ws, r, 1, 1 + len(mults), "Sensitivity - 2027 Unplanned (kbd): Growth (rows) x Multiplier (cols)", "h_orange")
        ws.write(r, 1, "Growth \\ Mult", self.wb.add_format(
            {"font_name": "Arial", "bold": True, "font_color": WHITE, "bg_color": NAVY,
             "align": "center", "valign": "vcenter", "border": 1}))
        for j, m in enumerate(mults):
            ws.write_number(r, 2 + j, m, self.wb.add_format(
                {"font_name": "Arial", "bold": True, "font_color": WHITE, "bg_color": NAVY,
                 "align": "center", "valign": "vcenter", "border": 1, "num_format": MULT}))
        top = r; r += 1
        bf = r
        for gi, g in enumerate(growths):
            ws.write_number(r, 1, g, self.wb.add_format(
                {"font_name": "Arial", "bold": True, "font_color": WHITE, "bg_color": NAVY,
                 "align": "center", "valign": "vcenter", "border": 1, "num_format": PCT}))
            for mj, m in enumerate(mults):
                isbase = abs(g) < 1e-9 and abs(m - 1.0) < 1e-9
                fmt = self.f["hm_base"] if isbase else self.f["hm_cell"]
                gref = A1(top + 1 + gi, 1, col_abs=True); mref = A1(top, 2 + mj, row_abs=True)
                ws.write_formula(r, 2 + mj, f"={anchor}*(1+{gref})*{mref}", fmt)
            r += 1
        ws.conditional_format(bf, 2, r - 1, 1 + len(mults),
                              {"type": "3_color_scale", "min_color": "#63BE7B",
                               "mid_color": "#FFEB84", "max_color": "#F8696B"})
        r += 1

        # --- tornado ---
        torn = self.ctx["tornado"]
        r = self._band(ws, r, 1, 6, "Tornado - 2027 Unplanned Sensitivity by Driver (kbd)", "h_orange")
        for j, h in enumerate(["Driver", "Low", "Base", "High", "Δ Low", "Δ High"]):
            ws.write(r, 1 + j, h, self.f["colhdr_l"] if j == 0 else self.f["colhdr"])
        r += 1
        tf = r
        for row in torn:
            ws.write(r, 1, row["driver"], self.f["rowlab"])
            ws.write_number(r, 2, row["low"], self.f["kbd"]); ws.write_number(r, 3, row["base"], self.f["kbd"])
            ws.write_number(r, 4, row["high"], self.f["kbd"])
            ws.write_formula(r, 5, f"={A1(r,2)}-{A1(r,3)}", self.f["kbd"])
            ws.write_formula(r, 6, f"={A1(r,4)}-{A1(r,3)}", self.f["kbd"])
            r += 1
        tl = r - 1
        tch = self.wb.add_chart({"type": "bar", "subtype": "stacked"})
        tch.add_series({"name": "Downside", "categories": [SHEET, tf, 1, tl, 1],
                        "values": [SHEET, tf, 5, tl, 5], "fill": {"color": BLUE}})
        tch.add_series({"name": "Upside", "categories": [SHEET, tf, 1, tl, 1],
                        "values": [SHEET, tf, 6, tl, 6], "fill": {"color": ORANGE}})
        tch.set_title({"name": "Tornado - Swing vs base (kbd)"}); tch.set_legend({"position": "bottom"})
        tch.set_size({"width": 640, "height": 280}); tch.set_chartarea({"border": {"none": True}})
        ws.insert_chart(tf, 8, tch)

    # ===================================================================== run
    def run(self):
        self.data_sheet()        # hidden backing first (Explorer references it)
        self.cover()
        self.dashboard()
        self.explorer()
        self.trends()
        self.padd()
        self.units_refineries()
        self.events_tas()
        self.model()
        # order tabs: Cover first, Data last/hidden
        self.wb.worksheets_objs.sort(key=lambda w: (
            ["Cover", "Dashboard", "Explorer", "Trends", "PADD", "Units & Refineries",
             "Events & TAs", "Model", "Data"].index(w.name)))
        self.wb.close()


def main():
    ap = argparse.ArgumentParser(description="Trading-floor refinery-outage workbook")
    ap.add_argument("excel", nargs="?", default=INPUT_PATH, help="path to the outage .xlsx export")
    ap.add_argument("--out", default=OUT_PATH, help="output .xlsx path")
    args = ap.parse_args()
    print(f"Loading {args.excel.strip()} ...")
    ctx = engine.build_context(args.excel)
    print(f"  {ctx['diag']['rows']:,} rows | {ctx['diag']['events_distinct']:,} distinct outages")
    print(f"Building workbook -> {args.out}")
    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    Build(ctx, args.out).run()
    print("Done.")


if __name__ == "__main__":
    sys.exit(main())
