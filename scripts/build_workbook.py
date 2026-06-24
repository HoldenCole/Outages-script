#!/usr/bin/env python3
"""
build_workbook.py
Institutional refinery-outage Excel workbook (XlsxWriter).

Generates a polished, chart-rich, interactive .xlsx from the Snowflake outage
export. Re-runnable: point INPUT_PATH at a new export and re-run to get the same
workbook with refreshed numbers.

Sheets (in order):
  Cover - Dashboard - Summary - Monthly - PADD Charts - PADD Detail -
  Units - Refinery Detail - Scenario 2027 - Sensitivity - Mogas Overlay - Notes

Primary metric: CAP_OFFLINE_ADJUSTED_KBD (offline capacity, all units). The
Scenario and Sensitivity sheets are driven by LIVE Excel formulas wired to
data-validation dropdowns, so editing an input recalculates everything.

Usage:
    python build_workbook.py                       # uses INPUT_PATH below
    python build_workbook.py path/to/export.xlsx --out outage_workbook.xlsx
"""
import argparse
import sys

import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell as A1, xl_range_abs

from pathlib import Path

import engine

_ROOT = Path(__file__).resolve().parent.parent          # repo root (scripts/ -> ..)
INPUT_PATH = str(_ROOT / "data" / "rEFINERY oUTAGES.xlsx")
OUT_PATH = str(_ROOT / "output" / "outage_workbook.xlsx")

# --------------------------------------------------------------------------- palette
NAVY = "#1F3864"
BLUE = "#2E5496"
RED = "#C00000"
GOLD = "#BF9000"
GREEN = "#548235"
ORANGE = "#ED7D31"
LT_BLUE = "#D6E0F0"
LT_GRAY = "#F2F2F2"
YELLOW = "#FFF2CC"
GRAY = "#808080"
WHITE = "#FFFFFF"

PADD_LINE_YEARS = [(2025, RED), (2024, BLUE), (2023, GRAY)]   # prior-year lines
KBD = '#,##0;(#,##0);"-"'
KBD1 = '#,##0.0;(#,##0.0);"-"'
PCT = "0.0%"
MULT = '0.0"x"'
YR = "@"

MONTHS = engine.MONTHS
PADDS = engine.PADD_ORDER
PARTIAL = set(engine.PARTIAL_YEARS)


class Build:
    def __init__(self, ctx, out_path):
        self.ctx = ctx
        self.wb = xlsxwriter.Workbook(out_path, {"nan_inf_to_errors": True})
        self.f = self._formats()
        # filled in by the scenario sheet, consumed by sensitivity:
        self.sc_refs = {}

    # ----------------------------------------------------------------- formats
    def _formats(self):
        wb = self.wb
        base = {"font_name": "Arial", "font_size": 10}
        mk = lambda **kw: wb.add_format({**base, **kw})
        f = {
            "title":     mk(font_size=18, bold=True, font_color=NAVY),
            "subtitle":  mk(font_size=11, italic=True, font_color="#595959"),
            "h_navy":    mk(font_size=11, bold=True, font_color=WHITE, bg_color=NAVY,
                            align="left", valign="vcenter", indent=1),
            "h_blue":    mk(font_size=11, bold=True, font_color=WHITE, bg_color=BLUE,
                            align="left", valign="vcenter", indent=1),
            "h_red":     mk(font_size=11, bold=True, font_color=WHITE, bg_color=RED,
                            align="left", valign="vcenter", indent=1),
            "h_green":   mk(font_size=11, bold=True, font_color=WHITE, bg_color=GREEN,
                            align="left", valign="vcenter", indent=1),
            "h_orange":  mk(font_size=11, bold=True, font_color=WHITE, bg_color="#C55A11",
                            align="left", valign="vcenter", indent=1),
            "colhdr":    mk(bold=True, font_color=WHITE, bg_color=NAVY, align="center",
                            valign="vcenter", border=1, border_color=WHITE),
            "colhdr_l":  mk(bold=True, font_color=WHITE, bg_color=NAVY, align="left",
                            valign="vcenter", border=1, border_color=WHITE, indent=1),
            "rowlab":    mk(align="left", valign="vcenter", indent=1),
            "rowlab_b":  mk(bold=True, align="left", valign="vcenter", indent=1),
            "rowlab_p":  mk(align="left", valign="vcenter", indent=1, italic=True,
                            font_color=GRAY),         # partial-year label
            "kbd":       mk(num_format=KBD, align="right"),
            "kbd_b":     mk(num_format=KBD, align="right", bold=True),
            "kbd_p":     mk(num_format=KBD, align="right", italic=True, font_color=GRAY),
            "kbd_sh":    mk(num_format=KBD, align="right", bg_color=LT_GRAY),
            "kbd_p_sh":  mk(num_format=KBD, align="right", italic=True, font_color=GRAY,
                            bg_color=LT_GRAY),
            "pct":       mk(num_format=PCT, align="right"),
            "pct_p":     mk(num_format=PCT, align="right", italic=True, font_color=GRAY),
            "pct_g":     mk(num_format=PCT, align="right", font_color="#008000"),
            "mult":      mk(num_format=MULT, align="right"),
            "yr":        mk(num_format=YR, bold=True, align="center", valign="vcenter",
                            font_color=WHITE, bg_color=NAVY, border=1, border_color=WHITE),
            "yr_p":      mk(num_format=YR, bold=True, align="center", valign="vcenter",
                            font_color="#D9D9D9", bg_color=NAVY, border=1, border_color=WHITE),
            "na":        mk(align="right", italic=True, font_color=GRAY),
            "note":      mk(font_size=10, font_color="#404040", text_wrap=True, valign="top"),
            "note_b":    mk(font_size=10, bold=True, font_color=NAVY, valign="top"),
            "red_note":  mk(font_size=10, bold=True, font_color=RED, text_wrap=True, valign="top"),
            "link":      mk(font_color=BLUE, underline=1, align="left", indent=1),
            # KPI tiles
            "kpi_lab":   mk(font_size=10, bold=True, font_color=WHITE, bg_color=BLUE,
                            align="center", valign="vcenter", text_wrap=True),
            "kpi_num":   mk(font_size=22, bold=True, font_color=NAVY, bg_color=LT_BLUE,
                            align="center", valign="vcenter"),
            "kpi_sub":   mk(font_size=9, italic=True, font_color="#595959", bg_color=LT_BLUE,
                            align="center", valign="vcenter"),
            # scenario input cells (blue font = user input, yellow fill = assumption)
            "in_lab":    mk(bold=True, align="left", valign="vcenter", indent=1,
                            bg_color=YELLOW, border=1, border_color="#BF8F00"),
            "in_val":    mk(font_color="#0000FF", bold=True, align="center", valign="vcenter",
                            bg_color=YELLOW, border=1, border_color="#BF8F00"),
            "in_pct":    mk(num_format=PCT, font_color="#0000FF", bold=True, align="center",
                            valign="vcenter", bg_color=YELLOW, border=1, border_color="#BF8F00"),
            "in_mult":   mk(num_format=MULT, font_color="#0000FF", bold=True, align="center",
                            valign="vcenter", bg_color=YELLOW, border=1, border_color="#BF8F00"),
            "in_kbd":    mk(num_format=KBD, font_color="#0000FF", bold=True, align="center",
                            valign="vcenter", bg_color=YELLOW, border=1, border_color="#BF8F00"),
            "calc":      mk(num_format=KBD, align="right"),         # black = formula
            "calc_b":    mk(num_format=KBD, align="right", bold=True, bg_color=LT_BLUE),
            "calc_grn":  mk(num_format=KBD, align="right", font_color="#008000"),  # cross-sheet
            "hm_hdr":    mk(bold=True, font_color=WHITE, bg_color=NAVY, align="center",
                            valign="vcenter", border=1),
            "hm_cell":   mk(num_format="#,##0", align="center", valign="vcenter", border=1,
                            border_color="#BFBFBF"),
            "hm_base":   mk(num_format="#,##0", align="center", valign="vcenter",
                            border=2, border_color=NAVY, bold=True),
        }
        return f

    def colors(self):
        return {"navy": NAVY, "blue": BLUE, "red": RED, "gold": GOLD}

    # ----------------------------------------------------------------- helpers
    def _setup(self, ws, tab_color, landscape=True):
        ws.hide_gridlines(2)
        ws.set_tab_color(tab_color)
        if landscape:
            ws.set_landscape()
            ws.set_paper(1)
            ws.fit_to_pages(1, 0)
        ws.set_margins(0.3, 0.3, 0.4, 0.4)

    def _band(self, ws, row, c0, c1, text, fmt="h_navy"):
        ws.merge_range(row, c0, row, c1, text, self.f[fmt])
        ws.set_row(row, 20)

    def _yr_fmt(self, year):
        return self.f["yr_p"] if year in PARTIAL else self.f["yr"]

    def _lab_fmt(self, year):
        return self.f["rowlab_p"] if year in PARTIAL else self.f["rowlab_b"]

    def _kbd_fmt(self, year, shade=False):
        if year in PARTIAL:
            return self.f["kbd_p_sh"] if shade else self.f["kbd_p"]
        return self.f["kbd_sh"] if shade else self.f["kbd"]

    # ===================================================================== COVER
    def cover(self):
        ws = self.wb.add_worksheet("Cover")
        self._setup(ws, NAVY, landscape=False)
        ws.set_column("A:A", 2)
        ws.set_column("B:B", 32)
        ws.set_column("C:H", 13)
        d = self.ctx["diag"]
        y0, y1 = d["years"]
        ws.set_row(1, 8)
        ws.merge_range("B3:H3", "REFINERY OUTAGE ANALYTICS", self.f["title"])
        ws.merge_range("B4:H4",
                       "Capacity Offline (kbd) - Planned & Unplanned - 2027 Scenario & Sensitivity",
                       self.f["subtitle"])
        rows = [
            ("Primary metric", "Capacity offline, thousand barrels/day (kbd) - CAP_OFFLINE_ADJUSTED_KBD"),
            ("Scope", "All US refining units & products; mogas-equivalent shown as a secondary overlay"),
            ("Data vintage", f"{d['rows']:,} rows  |  {y0}-{y1}  |  PADD {d['padd_source']}"),
            ("Event basis", f"{d['events_distinct']:,} distinct outages (OUTAGE_ID); rows are unit-month slices"),
            ("Outage type", "Binary {PLANNED, UNPLANNED}; UNKNOWN folded into UNPLANNED (desk rule)"),
        ]
        r = 6
        for lab, val in rows:
            ws.write(r, 1, lab, self.f["note_b"])
            ws.merge_range(r, 2, r, 7, val, self.f["note"])
            r += 1
        r += 1
        self._band(ws, r, 1, 7, "Read-Before-Use Caveats", "h_red"); r += 1
        caveats = [
            "2026 & 2027 are PARTIAL / special years - shown grey italic everywhere and never presented as final.",
            "2027 is PLANNED-ONLY in the source. Unplanned-2027 is a MODELED scenario, not an actual.",
            "Comparison guardrail: any Plan+Unplanned or Unplanned comparison vs 2027 shows n/a; only Planned is valid vs 2027.",
            "2020-2021 are COVID / Winter-Storm-Uri outliers and are excluded from forecast baselines by default.",
        ]
        for c in caveats:
            ws.merge_range(r, 1, r, 7, "-  " + c, self.f["red_note"])
            ws.set_row(r, 26)
            r += 1
        r += 1
        self._band(ws, r, 1, 7, "Contents", "h_navy"); r += 1
        toc = [
            ("Dashboard", "At-a-glance KPIs and headline charts"),
            ("Summary", "Annual table 2016-2027 + targeted year comparisons"),
            ("Monthly", "Total / Planned / Unplanned month x year matrices + seasonality"),
            ("PADD Charts", "Per-PADD planned+unplanned combo charts vs prior years"),
            ("PADD Detail", "PADD x year matrices and unplanned-by-PADD"),
            ("Units", "Unit-category x year matrix + magnitude bars"),
            ("Refinery Detail", "Top refineries, operators and event scatter"),
            ("Scenario 2027", "Live driver-based 2027 unplanned forecast (dropdowns)"),
            ("Sensitivity", "Two-way heatmap and tornado of the scenario drivers"),
            ("Mogas Overlay", "Secondary mogas-equivalent view"),
            ("Notes", "Sources, methodology and color key"),
        ]
        for name, desc in toc:
            ws.write_url(r, 1, f"internal:'{name}'!A1", self.f["link"], name)
            ws.merge_range(r, 2, r, 7, desc, self.f["note"])
            r += 1

    # ================================================================= DASHBOARD
    def dashboard(self):
        ws = self.wb.add_worksheet("Dashboard")
        self._setup(ws, RED)
        ws.set_column("A:A", 2)
        ws.set_column("B:M", 11)
        s = self.ctx["summary"]
        # latest full year = most recent non-partial year with actual unplanned data
        # (2026/27 are partial; 2028+ are future planned-only with no unplanned).
        full = [y for y in s.index if y not in PARTIAL and s.loc[y, "Unplanned"] > 0]
        ly = max(full)                       # latest full year (2025)
        ws.merge_range("B2:M2", "Outage Dashboard", self.f["title"])
        ws.merge_range("B3:M3", f"Headline metrics - latest full year {ly}", self.f["subtitle"])

        # KPI tiles
        padd_un = self.ctx["padd_unplanned"]
        top_padd = padd_un[ly].idxmax() if ly in padd_un.columns else "-"
        tiles = [
            ("Total Offline", s.loc[ly, "Total"], "kbd", f"FY{ly}"),
            ("Unplanned", s.loc[ly, "Unplanned"], "kbd", f"FY{ly}"),
            ("Unplanned %", s.loc[ly, "Unpl%"], "pct", "of total offline"),
            ("Distinct Outages", s.loc[ly, "Events"], "int", f"FY{ly}"),
            ("Top PADD", top_padd, "txt", "by unplanned kbd"),
        ]
        col = 1
        for lab, val, kind, sub in tiles:
            ws.merge_range(4, col, 4, col + 1, lab, self.f["kpi_lab"])
            if kind == "pct":
                cell = self.wb.add_format({"font_name": "Arial", "font_size": 22, "bold": True,
                                           "font_color": NAVY, "bg_color": LT_BLUE,
                                           "align": "center", "valign": "vcenter", "num_format": "0%"})
                ws.merge_range(5, col, 6, col + 1, val, cell)
            elif kind in ("kbd", "int"):
                cell = self.wb.add_format({"font_name": "Arial", "font_size": 20, "bold": True,
                                           "font_color": NAVY, "bg_color": LT_BLUE,
                                           "align": "center", "valign": "vcenter", "num_format": "#,##0"})
                ws.merge_range(5, col, 6, col + 1, val, cell)
            else:
                ws.merge_range(5, col, 6, col + 1, val, self.wb.add_format(
                    {"font_name": "Arial", "font_size": 16, "bold": True, "font_color": NAVY,
                     "bg_color": LT_BLUE, "align": "center", "valign": "vcenter"}))
            ws.merge_range(7, col, 7, col + 1, sub, self.f["kpi_sub"])
            col += 2
        ws.set_row(5, 26); ws.set_row(6, 12)

        # --- data block for charts (hidden-ish, below the fold) ---
        d0 = 40
        ws.write(d0 - 1, 1, "Chart data (capacity offline, kbd)", self.f["note_b"])
        years = [y for y in s.index if 2018 <= y <= 2027]
        ws.write(d0, 1, "Year", self.f["colhdr_l"])
        ws.write(d0, 2, "Planned", self.f["colhdr"])
        ws.write(d0, 3, "Unplanned", self.f["colhdr"])
        for i, y in enumerate(years):
            ws.write(d0 + 1 + i, 1, y, self.f["yr"])
            ws.write_number(d0 + 1 + i, 2, float(s.loc[y, "Planned"]), self.f["kbd"])
            ws.write_number(d0 + 1 + i, 3, float(s.loc[y, "Unplanned"]), self.f["kbd"])
        n = len(years)

        # Stacked column: planned (navy) + unplanned (red) by year
        col_chart = self.wb.add_chart({"type": "column", "subtype": "stacked"})
        for ci, (name, color) in enumerate([("Planned", NAVY), ("Unplanned", RED)]):
            col_chart.add_series({
                "name": ["Dashboard", d0, 2 + ci],
                "categories": ["Dashboard", d0 + 1, 1, d0 + n, 1],
                "values": ["Dashboard", d0 + 1, 2 + ci, d0 + n, 2 + ci],
                "fill": {"color": color},
                "gap": 60,
            })
        col_chart.set_title({"name": "Capacity Offline by Year (Planned + Unplanned)"})
        col_chart.set_x_axis({"num_format": "0"})
        col_chart.set_y_axis({"name": "kbd", "major_gridlines": {"visible": True,
                              "line": {"color": "#E0E0E0"}}})
        col_chart.set_legend({"position": "bottom"})
        col_chart.set_size({"width": 560, "height": 300})
        col_chart.set_chartarea({"border": {"none": True}})
        ws.insert_chart("B9", col_chart)

        # Clustered column: unplanned by PADD across recent years
        rec = [y for y in [2022, 2023, 2024, 2025] if y in padd_un.columns]
        p0 = d0 + n + 3
        ws.write(p0 - 1, 1, "Unplanned by PADD (kbd)", self.f["note_b"])
        ws.write(p0, 1, "PADD", self.f["colhdr_l"])
        for j, y in enumerate(rec):
            ws.write(p0, 2 + j, y, self.f["yr"])
        for i, p in enumerate(PADDS):
            ws.write(p0 + 1 + i, 1, p, self.f["rowlab"])
            for j, y in enumerate(rec):
                ws.write_number(p0 + 1 + i, 2 + j, float(padd_un.loc[p, y]), self.f["kbd"])
        clu = self.wb.add_chart({"type": "column"})
        palette = [BLUE, GOLD, GREEN, RED]
        for j, y in enumerate(rec):
            clu.add_series({
                "name": ["Dashboard", p0, 2 + j],
                "categories": ["Dashboard", p0 + 1, 1, p0 + 5, 1],
                "values": ["Dashboard", p0 + 1, 2 + j, p0 + 5, 2 + j],
                "fill": {"color": palette[j % len(palette)]},
            })
        clu.set_title({"name": "Unplanned Offline by PADD"})
        clu.set_y_axis({"name": "kbd"})
        clu.set_legend({"position": "bottom"})
        clu.set_size({"width": 560, "height": 300})
        clu.set_chartarea({"border": {"none": True}})
        ws.insert_chart("H9", clu)

        # Donut: unplanned share by PADD latest full year
        sh0 = p0 + 8
        ws.write(sh0 - 1, 1, f"Unplanned share by PADD ({ly})", self.f["note_b"])
        for i, p in enumerate(PADDS):
            ws.write(sh0 + i, 1, p, self.f["rowlab"])
            ws.write_number(sh0 + i, 2, float(padd_un.loc[p, ly]), self.f["kbd"])
        donut = self.wb.add_chart({"type": "doughnut"})
        donut.add_series({
            "name": f"Unplanned share {ly}",
            "categories": ["Dashboard", sh0, 1, sh0 + 4, 1],
            "values": ["Dashboard", sh0, 2, sh0 + 4, 2],
            "points": [{"fill": {"color": c}} for c in [NAVY, BLUE, GOLD, GREEN, ORANGE]],
        })
        donut.set_title({"name": f"Unplanned Share by PADD ({ly})"})
        donut.set_legend({"position": "right"})
        donut.set_size({"width": 380, "height": 300})
        donut.set_hole_size(55)
        donut.set_chartarea({"border": {"none": True}})
        ws.insert_chart("B25", donut)

        ws.freeze_panes(8, 0)

    # =================================================================== SUMMARY
    def summary(self):
        ws = self.wb.add_worksheet("Summary")
        self._setup(ws, BLUE)
        ws.set_column("A:A", 2)
        ws.set_column("B:B", 10)
        ws.set_column("C:I", 12)
        s = self.ctx["summary"]
        ws.merge_range("B2:I2", "Annual Summary & Comparisons", self.f["title"])
        ws.merge_range("B3:I3", "Capacity offline (kbd). 2026/27 grey italic; 2027 planned-only.",
                       self.f["subtitle"])

        # --- annual table 2016-2027 with formula columns ---
        r = 5
        self._band(ws, r, 1, 8, "Annual Capacity Offline (kbd)"); r += 1
        heads = ["Year", "Planned", "Unplanned", "Total", "Events", "Unpl %", "YoY Δ", "YoY %"]
        for j, h in enumerate(heads):
            ws.write(r, 1 + j, h, self.f["colhdr_l"] if j == 0 else self.f["colhdr"])
        r += 1
        years = [y for y in s.index if 2016 <= y <= 2027]
        first = r
        for i, y in enumerate(years):
            partial = y in PARTIAL
            kf = self.f["kbd_p"] if partial else self.f["kbd"]
            pf = self.f["pct_p"] if partial else self.f["pct"]
            ws.write_number(r, 1, y, self._yr_fmt(y))
            ws.write_number(r, 2, float(s.loc[y, "Planned"]), kf)
            ws.write_number(r, 3, float(s.loc[y, "Unplanned"]), kf)
            ws.write_formula(r, 4, f"={A1(r,2)}+{A1(r,3)}", kf)         # Total
            ws.write_number(r, 5, int(s.loc[y, "Events"]),
                            self.f["kbd_p"] if partial else self.f["kbd"])
            ws.write_formula(r, 6, f"=IF({A1(r,4)}=0,0,{A1(r,3)}/{A1(r,4)})", pf)  # Unpl%
            if i == 0:
                ws.write_blank(r, 7, None, kf)
                ws.write_blank(r, 8, None, pf)
            else:
                ws.write_formula(r, 7, f"={A1(r,4)}-{A1(r-1,4)}", kf)   # YoY delta
                ws.write_formula(r, 8, f"=IF({A1(r-1,4)}=0,0,{A1(r,7)}/{A1(r-1,4)})", pf)
            r += 1
        ws.write(r, 1, "Note", self.f["rowlab"])
        ws.merge_range(r, 2, r, 8,
                       "2027 Unplanned = 0 by construction (planned-only year); see Scenario 2027.",
                       self.f["subtitle"])
        ws.freeze_panes(first, 0)

        # --- targeted comparison blocks (2027 guardrail) ---
        r += 3
        self._band(ws, r, 1, 8, "Targeted Comparisons (Δ% vs base year)"); r += 1
        cmp = self.ctx["compare"]
        for key, (ya, yb) in [("2025v2026", (2025, 2026)),
                              ("2025v2027", (2025, 2027)),
                              ("2026v2027", (2026, 2027))]:
            ws.write(r, 1, f"{ya} → {yb}", self.f["rowlab_b"])
            ws.write(r, 2, f"{ya}", self.f["colhdr"])
            ws.write(r, 3, f"{yb}", self.f["colhdr"])
            ws.write(r, 4, "Δ kbd", self.f["colhdr"])
            ws.write(r, 5, "Δ %", self.f["colhdr"])
            r += 1
            for metric in ["Plan + Unplanned", "Planned", "Unplanned"]:
                blk = cmp[key][metric]
                ws.write(r, 1, metric, self.f["rowlab"])
                if blk["a"] is None or blk["b"] is None:
                    ws.write_number(r, 2, blk["a"], self.f["kbd"]) if blk["a"] is not None \
                        else ws.write(r, 2, "n/a", self.f["na"])
                    ws.write(r, 3, "n/a", self.f["na"])
                    ws.write(r, 4, "n/a", self.f["na"])
                    ws.write(r, 5, "n/a", self.f["na"])
                else:
                    ws.write_number(r, 2, blk["a"], self.f["kbd"])
                    ws.write_number(r, 3, blk["b"], self.f["kbd"])
                    ws.write_formula(r, 4, f"={A1(r,3)}-{A1(r,2)}", self.f["kbd"])
                    ws.write_formula(r, 5, f"=IF({A1(r,2)}=0,0,{A1(r,4)}/{A1(r,2)})", self.f["pct"])
                r += 1
            r += 1
        ws.write(r, 1, "Guardrail: only Planned is comparable vs 2027 (no actual unplanned-2027 exists).",
                 self.f["subtitle"])

    # =================================================================== MONTHLY
    def monthly(self):
        ws = self.wb.add_worksheet("Monthly")
        self._setup(ws, BLUE)
        ws.set_column("A:A", 2)
        ws.set_column("B:B", 8)
        ws.set_column("C:N", 7.5)
        ws.set_column("O:O", 9)
        ws.merge_range("B2:O2", "Monthly Detail & Seasonality", self.f["title"])
        ws.merge_range("B3:O3", "Capacity offline (kbd) by month x year.", self.f["subtitle"])

        blocks = [("Total Offline", self.ctx["monthly_total"], "h_navy"),
                  ("Planned", self.ctx["monthly_planned"], "h_blue"),
                  ("Unplanned", self.ctx["monthly_unplanned"], "h_red")]
        r = 5
        unpl_first_row = unpl_last_row = None
        for title, mat, band in blocks:
            self._band(ws, r, 1, 14, f"{title} (kbd)", band); r += 1
            ws.write(r, 1, "Year", self.f["colhdr_l"])
            for j, m in enumerate(MONTHS):
                ws.write(r, 2 + j, m, self.f["colhdr"])
            ws.write(r, 14, "Total", self.f["colhdr"])
            r += 1
            years = [y for y in mat.index if 2018 <= y <= 2027]
            if title == "Unplanned":
                unpl_first_row = r
            for y in years:
                partial = y in PARTIAL
                kf = self.f["kbd_p"] if partial else self.f["kbd"]
                ws.write_number(r, 1, y, self._yr_fmt(y))
                for j, m in enumerate(MONTHS):
                    ws.write_number(r, 2 + j, float(mat.loc[y, m]), kf)
                ws.write_formula(r, 14, f"=SUM({A1(r,2)}:{A1(r,13)})",
                                 self.f["kbd_p"] if partial else self.f["kbd_b"])
                r += 1
            if title == "Unplanned":
                unpl_last_row = r - 1
                unpl_years = years
            r += 1

        # Seasonality line chart: unplanned by month, one series per recent year
        line = self.wb.add_chart({"type": "line"})
        recent = [y for y in unpl_years if y in (2022, 2023, 2024, 2025, 2026)]
        cmap = {2022: GRAY, 2023: GREEN, 2024: BLUE, 2025: RED, 2026: GOLD}
        for y in recent:
            rr = unpl_first_row + (unpl_years.index(y))
            line.add_series({
                "name": str(y),
                "categories": ["Monthly", unpl_first_row - 1, 2, unpl_first_row - 1, 13],
                "values": ["Monthly", rr, 2, rr, 13],
                "line": {"color": cmap.get(y, NAVY), "width": 2.0,
                         "dash_type": "dash" if y == 2026 else "solid"},
            })
        line.set_title({"name": "Unplanned Offline Seasonality (kbd by month)"})
        line.set_y_axis({"name": "kbd"})
        line.set_legend({"position": "bottom"})
        line.set_size({"width": 820, "height": 320})
        line.set_chartarea({"border": {"none": True}})
        ws.insert_chart(r + 1, 1, line)
        ws.freeze_panes(7, 2)

    # =============================================================== PADD CHARTS
    def padd_charts(self):
        ws = self.wb.add_worksheet("PADD Charts")
        self._setup(ws, GREEN)
        ws.set_column("A:A", 2)
        ws.set_column("B:B", 18)
        ws.set_column("C:N", 8)
        ws.merge_range("B2:N2", "PADD Planned & Unplanned Offline (kbd)", self.f["title"])
        ws.merge_range("B3:N3",
                       "Columns = 2026 plan + unplanned (stacked). Lines = prior-year total offline; "
                       "green = 2027 plan.", self.f["subtitle"])
        pm = self.ctx["padd_month"]
        r = 5
        CUR = 2026
        for p in PADDS:
            self._band(ws, r, 1, 13, f"{p}  -  {CUR} stack vs 2023-2025 totals & 2027 plan", "h_green")
            r += 1
            hdr = r
            ws.write(r, 1, "Series", self.f["colhdr_l"])
            for j, m in enumerate(MONTHS):
                ws.write(r, 2 + j, m, self.f["colhdr"])
            r += 1
            # data rows
            def row_of(mat, yr):
                return [float(mat.loc[yr, m]) if yr in mat.index else 0.0 for m in MONTHS]
            series = [
                (f"{CUR} Planned", row_of(pm[p]["planned"], CUR), self.f["rowlab"]),
                (f"{CUR} Unplanned", row_of(pm[p]["unplanned"], CUR), self.f["rowlab"]),
                ("2025 Total", row_of(pm[p]["total"], 2025), self.f["rowlab"]),
                ("2024 Total", row_of(pm[p]["total"], 2024), self.f["rowlab"]),
                ("2023 Total", row_of(pm[p]["total"], 2023), self.f["rowlab"]),
                ("2027 Planned", row_of(pm[p]["planned"], 2027), self.f["rowlab"]),
            ]
            data_first = r
            for name, vals, lf in series:
                ws.write(r, 1, name, lf)
                for j, v in enumerate(vals):
                    ws.write_number(r, 2 + j, v, self.f["kbd"])
                r += 1

            col = self.wb.add_chart({"type": "column", "subtype": "stacked"})
            col.add_series({"name": ["PADD Charts", data_first, 1],
                            "categories": ["PADD Charts", hdr, 2, hdr, 13],
                            "values": ["PADD Charts", data_first, 2, data_first, 13],
                            "fill": {"color": GOLD}})
            col.add_series({"name": ["PADD Charts", data_first + 1, 1],
                            "categories": ["PADD Charts", hdr, 2, hdr, 13],
                            "values": ["PADD Charts", data_first + 1, 2, data_first + 1, 13],
                            "fill": {"color": ORANGE}})
            line = self.wb.add_chart({"type": "line"})
            for k, (yr, color) in enumerate([(2025, RED), (2024, BLUE), (2023, GRAY), (2027, GREEN)]):
                rr = data_first + 2 + k
                line.add_series({"name": ["PADD Charts", rr, 1],
                                 "categories": ["PADD Charts", hdr, 2, hdr, 13],
                                 "values": ["PADD Charts", rr, 2, rr, 13],
                                 "line": {"color": color, "width": 2.25}})
            col.combine(line)
            col.set_title({"name": f"{p} Planned & Unplanned Offline (kbd)"})
            col.set_x_axis({"name": ""})
            col.set_y_axis({"name": "kbd"})
            col.set_legend({"position": "top"})
            col.set_size({"width": 760, "height": 360})
            col.set_chartarea({"border": {"none": True}})
            ws.insert_chart(data_first, 15, col)
            r = max(r, data_first + 19) + 1

    # =============================================================== PADD DETAIL
    def padd_detail(self):
        ws = self.wb.add_worksheet("PADD Detail")
        self._setup(ws, GREEN)
        ws.set_column("A:A", 2)
        ws.set_column("B:B", 12)
        ws.set_column("C:N", 9)
        ws.merge_range("B2:N2", "PADD Detail Matrices", self.f["title"])
        ws.merge_range("B3:N3", "Capacity offline (kbd) by PADD x year, with Total US.",
                       self.f["subtitle"])
        mats = [("Total Offline", self.ctx["padd_total"], "h_navy"),
                ("Unplanned", self.ctx["padd_unplanned"], "h_red"),
                ("Planned", self.ctx["padd_planned"], "h_blue")]
        r = 5
        un_block = None
        for title, mat, band in mats:
            years = [y for y in mat.columns if 2018 <= y <= 2027]
            self._band(ws, r, 1, 1 + len(years), f"{title} (kbd)", band); r += 1
            ws.write(r, 1, "PADD", self.f["colhdr_l"])
            for j, y in enumerate(years):
                ws.write_number(r, 2 + j, y, self._yr_fmt(y))
            r += 1
            blk_first = r
            for p in PADDS:
                ws.write(r, 1, p, self.f["rowlab"])
                for j, y in enumerate(years):
                    ws.write_number(r, 2 + j, float(mat.loc[p, y]),
                                    self.f["kbd_p"] if y in PARTIAL else self.f["kbd"])
                r += 1
            # Total US row = SUM
            ws.write(r, 1, "Total US", self.f["rowlab_b"])
            for j, y in enumerate(years):
                ws.write_formula(r, 2 + j,
                                 f"=SUM({A1(blk_first,2+j)}:{A1(r-1,2+j)})",
                                 self.f["kbd_p"] if y in PARTIAL else self.f["kbd_b"])
            if title == "Unplanned":
                un_block = (blk_first, years)
            r += 2

        # clustered column of unplanned by PADD (recent years)
        if un_block:
            bf, years = un_block
            rec = [y for y in years if y in (2022, 2023, 2024, 2025)]
            clu = self.wb.add_chart({"type": "column"})
            for p_i, p in enumerate(PADDS):
                clu.add_series({
                    "name": ["PADD Detail", bf + p_i, 1],
                    "categories": ["PADD Detail", bf - 1, 2, bf - 1, 1 + len(years)],
                    "values": ["PADD Detail", bf + p_i, 2, bf + p_i, 1 + len(years)],
                })
            clu.set_title({"name": "Unplanned Offline by PADD (kbd)"})
            clu.set_y_axis({"name": "kbd"})
            clu.set_legend({"position": "bottom"})
            clu.set_size({"width": 820, "height": 320})
            clu.set_chartarea({"border": {"none": True}})
            ws.insert_chart(r, 1, clu)
        ws.freeze_panes(7, 2)

    # ===================================================================== UNITS
    def units(self):
        ws = self.wb.add_worksheet("Units")
        self._setup(ws, GOLD)
        ws.set_column("A:A", 2)
        ws.set_column("B:B", 24)
        ws.set_column("C:K", 9)
        ws.set_column("L:L", 11)
        ws.merge_range("B2:L2", "Unit-Category Detail", self.f["title"])
        ws.merge_range("B3:L3", "Capacity offline (kbd) by unit category x year, sorted by total.",
                       self.f["subtitle"])
        mat = self.ctx["unit_total"]
        years = [y for y in mat.columns if 2020 <= y <= 2027]
        r = 5
        self._band(ws, r, 1, 2 + len(years), "Capacity Offline by Unit Category (kbd)"); r += 1
        ws.write(r, 1, "Unit Category", self.f["colhdr_l"])
        for j, y in enumerate(years):
            ws.write_number(r, 2 + j, y, self._yr_fmt(y))
        ws.write(r, 2 + len(years), "Total", self.f["colhdr"])
        r += 1
        first = r
        tot_col = 2 + len(years)
        for i, u in enumerate(mat.index):
            shade = i % 2 == 1
            ws.write(r, 1, str(u).title(), self.f["rowlab"])
            for j, y in enumerate(years):
                ws.write_number(r, 2 + j, float(mat.loc[u, y]),
                                self._kbd_fmt(y, shade))
            ws.write_formula(r, tot_col, f"=SUM({A1(r,2)}:{A1(r,1+len(years))})",
                             self.f["kbd_b"])
            r += 1
        last = r - 1
        # data bars on the Total column
        ws.conditional_format(first, tot_col, last, tot_col,
                              {"type": "data_bar", "bar_color": GOLD,
                               "bar_solid": True})
        ws.freeze_panes(first, 2)

        # bar chart of top categories by total
        topn = min(10, len(mat.index))
        bar = self.wb.add_chart({"type": "bar"})
        bar.add_series({
            "name": "Total offline (kbd)",
            "categories": ["Units", first, 1, first + topn - 1, 1],
            "values": ["Units", first, tot_col, first + topn - 1, tot_col],
            "fill": {"color": GOLD},
            "data_labels": {"value": True, "num_format": "#,##0"},
        })
        bar.set_title({"name": f"Top {topn} Unit Categories by Capacity Offline"})
        bar.set_x_axis({"name": "kbd"})
        bar.set_legend({"none": True})
        bar.set_size({"width": 560, "height": 360})
        bar.set_chartarea({"border": {"none": True}})
        bar.reverse_series_order = True
        ws.insert_chart(first, tot_col + 2, bar)

    # =========================================================== REFINERY DETAIL
    def refinery_detail(self):
        ws = self.wb.add_worksheet("Refinery Detail")
        self._setup(ws, "#7030A0")
        ws.set_column("A:A", 2)
        ws.set_column("B:B", 30)
        ws.set_column("C:C", 9)
        ws.set_column("D:D", 26)
        ws.set_column("E:H", 11)
        ws.merge_range("B2:H2", "Refinery & Operator Detail", self.f["title"])
        ws.merge_range("B3:H3", "Top sites and operators by capacity offline (kbd), all years.",
                       self.f["subtitle"])

        # Top 15 refineries
        pl = self.ctx["plants"]
        r = 5
        self._band(ws, r, 1, 7, "Top 15 Refineries by Capacity Offline (kbd)"); r += 1
        for j, h in enumerate(["Refinery", "PADD", "Operator", "Total", "Planned", "Unplanned", "Events"]):
            ws.write(r, 1 + j, h, self.f["colhdr_l"] if j in (0, 2) else self.f["colhdr"])
        r += 1
        first = r
        for i, (_, row) in enumerate(pl.iterrows()):
            ws.write(r, 1, str(row["plant"]), self.f["rowlab"])
            ws.write(r, 2, str(row["padd"]), self.f["rowlab"])
            ws.write(r, 3, str(row["operator"]).title(), self.f["rowlab"])
            ws.write_number(r, 4, float(row["total"]), self.f["kbd"])
            ws.write_number(r, 5, float(row["planned"]), self.f["kbd"])
            ws.write_number(r, 6, float(row["unplanned"]), self.f["kbd"])
            ws.write_number(r, 7, int(row["events"]), self.f["kbd"])
            r += 1
        ws.conditional_format(first, 4, r - 1, 4,
                              {"type": "data_bar", "bar_color": "#7030A0", "bar_solid": True})

        # Top 10 operators x year
        r += 2
        ops = self.ctx["operators"]
        oyears = [y for y in ops.columns if 2020 <= y <= 2027]
        self._band(ws, r, 1, 1 + len(oyears), "Top 10 Operators x Year (kbd)"); r += 1
        ws.write(r, 1, "Operator", self.f["colhdr_l"])
        for j, y in enumerate(oyears):
            ws.write_number(r, 2 + j, y, self._yr_fmt(y))
        r += 1
        of = r
        for i, op in enumerate(ops.index):
            shade = i % 2 == 1
            ws.write(r, 1, str(op).title(), self.f["rowlab"])
            for j, y in enumerate(oyears):
                ws.write_number(r, 2 + j, float(ops.loc[op, y]), self._kbd_fmt(y, shade))
            r += 1

        # Scatter: capacity (Y) vs duration intensity (X) for recent unplanned events
        sc = self.ctx["scatter"]
        r += 2
        self._band(ws, r, 1, 4, "Unplanned Events: Capacity vs Monthly Intensity (2023-2025)"); r += 1
        ws.write(r, 1, "Duration (days)", self.f["colhdr"])
        ws.write(r, 2, "Capacity (kbd)", self.f["colhdr"])
        r += 1
        sfirst = r
        for _, row in sc.iterrows():
            ws.write_number(r, 1, float(row["duration"]), self.f["kbd"])
            ws.write_number(r, 2, float(row["cap_kbd"]), self.f["kbd"])
            r += 1
        slast = r - 1
        scat = self.wb.add_chart({"type": "scatter"})
        scat.add_series({
            "categories": ["Refinery Detail", sfirst, 1, slast, 1],
            "values": ["Refinery Detail", sfirst, 2, slast, 2],
            "marker": {"type": "circle", "size": 5,
                       "fill": {"color": "#7030A0"}, "border": {"none": True}},
        })
        scat.set_title({"name": "Unplanned Events - Capacity (kbd) vs Duration (days)"})
        scat.set_x_axis({"name": "Monthly intensity / duration (days)",
                         "min": 0, "major_gridlines": {"visible": True, "line": {"color": "#EEEEEE"}}})
        scat.set_y_axis({"name": "Capacity offline (kbd)", "min": 0})
        scat.set_legend({"none": True})
        scat.set_size({"width": 560, "height": 360})
        scat.set_chartarea({"border": {"none": True}})
        ws.insert_chart(sfirst, 4, scat)
        ws.freeze_panes(first, 0)

    # ============================================================== SCENARIO 2027
    def scenario(self):
        ws = self.wb.add_worksheet("Scenario 2027")
        self._setup(ws, RED)
        ws.set_column("A:A", 2)
        ws.set_column("B:B", 26)
        ws.set_column("C:N", 9)
        SHEET = "'Scenario 2027'"
        ws.merge_range("B2:N2", "2027 Unplanned Scenario (Live Model)", self.f["title"])
        ws.merge_range("B3:N3",
                       "Edit the yellow inputs - the forecast, chart, PADD split and the "
                       "Sensitivity sheet all recompute.", self.f["subtitle"])

        windows = list(engine.BASELINE_WINDOWS.keys())

        # --- inputs panel (yellow, blue font, dropdowns) ---
        r = 5
        self._band(ws, r, 1, 4, "Scenario Inputs", "h_red"); r += 1
        in_rows = {}
        # baseline window dropdown
        ws.write(r, 1, "Baseline window", self.f["in_lab"])
        ws.merge_range(r, 2, r, 4, engine.DEFAULT_WINDOW, self.f["in_val"])
        ws.data_validation(r, 2, r, 2, {"validate": "list", "source": windows})
        in_rows["window"] = r; r += 1
        ws.write(r, 1, "Production growth %", self.f["in_lab"])
        ws.merge_range(r, 2, r, 4, 0.0, self.f["in_pct"])
        in_rows["growth"] = r; r += 1
        ws.write(r, 1, "Unplanned rate multiplier", self.f["in_lab"])
        ws.merge_range(r, 2, r, 4, 1.0, self.f["in_mult"])
        in_rows["mult"] = r; r += 1
        ws.write(r, 1, "One-off event (kbd)", self.f["in_lab"])
        ws.merge_range(r, 2, r, 4, 0, self.f["in_kbd"])
        in_rows["oneoff"] = r; r += 1
        ws.write(r, 1, "Stress month", self.f["in_lab"])
        ws.merge_range(r, 2, r, 4, "Sep", self.f["in_val"])
        ws.data_validation(r, 2, r, 2, {"validate": "list", "source": MONTHS})
        in_rows["stress"] = r; r += 1

        c_window = A1(in_rows["window"], 2, row_abs=True, col_abs=True)
        c_growth = A1(in_rows["growth"], 2, row_abs=True, col_abs=True)
        c_mult = A1(in_rows["mult"], 2, row_abs=True, col_abs=True)
        c_oneoff = A1(in_rows["oneoff"], 2, row_abs=True, col_abs=True)
        c_stress = A1(in_rows["stress"], 2, row_abs=True, col_abs=True)

        # --- lookup: seasonality profiles per window (auditable data block) ---
        r += 1
        self._band(ws, r, 1, 13, "Lookup - Avg Unplanned Offline (kbd/month) by Baseline Window", "h_navy")
        r += 1
        ws.write(r, 1, "Window", self.f["colhdr_l"])
        for j, m in enumerate(MONTHS):
            ws.write(r, 2 + j, m, self.f["colhdr"])
        r += 1
        prof_first = r
        df = self.ctx["df"]
        for w in windows:
            prof = engine.baseline_profile(df, w)
            ws.write(r, 1, w, self.f["rowlab"])
            for j, m in enumerate(MONTHS):
                ws.write_number(r, 2 + j, float(prof[m]), self.f["kbd"])
            r += 1
        prof_last = r - 1
        prof_range = xl_range_abs(prof_first, 2, prof_last, 13)         # 12-col body
        wlabel_range = xl_range_abs(prof_first, 1, prof_last, 1)

        # --- forecast cascade (live formulas) ---
        r += 1
        self._band(ws, r, 1, 13, "2027 Unplanned Forecast Cascade (live)", "h_red"); r += 1
        ws.write(r, 1, "", self.f["colhdr_l"])
        for j, m in enumerate(MONTHS):
            ws.write(r, 2 + j, m, self.f["colhdr"])
        hdr_row = r; r += 1
        # baseline row (INDEX/MATCH into the profile block)
        ws.write(r, 1, "Baseline (window)", self.f["rowlab"])
        base_row = r
        for j in range(12):
            f = (f"=INDEX({prof_range},MATCH({c_window},{wlabel_range},0),{j+1})")
            ws.write_formula(r, 2 + j, f, self.f["calc"])
        r += 1
        # scaled / forecast row
        ws.write(r, 1, "Forecast unplanned", self.f["rowlab_b"])
        fc_row = r
        for j in range(12):
            mcell = A1(hdr_row, 2 + j)                          # month header e.g. "Jan"
            bcell = A1(base_row, 2 + j)
            f = (f"={bcell}*(1+{c_growth})*{c_mult}"
                 f"+IF({c_stress}={mcell},{c_oneoff},0)")
            ws.write_formula(r, 2 + j, f, self.f["calc_b"])
        r += 2

        # outputs
        self._band(ws, r, 1, 4, "Scenario Outputs", "h_red"); r += 1
        ws.write(r, 1, "Baseline annual (pre-shock)", self.f["rowlab"])
        anchor_cell = A1(r, 2, row_abs=True, col_abs=True)
        ws.write_formula(r, 2, f"=SUM({A1(base_row,2)}:{A1(base_row,13)})", self.f["calc"])
        r += 1
        ws.write(r, 1, "2027 Unplanned forecast", self.f["rowlab_b"])
        fc_annual_cell = A1(r, 2, row_abs=True, col_abs=True)
        ws.write_formula(r, 2, f"=SUM({A1(fc_row,2)}:{A1(fc_row,13)})", self.f["calc_b"])
        r += 1
        ws.write(r, 1, "2027 Planned (booked)", self.f["rowlab"])
        planned_2027 = float(self.ctx["summary"].loc[2027, "Planned"]) \
            if 2027 in self.ctx["summary"].index else 0.0
        planned_cell = A1(r, 2, row_abs=True, col_abs=True)
        ws.write_number(r, 2, planned_2027, self.f["calc_grn"])
        r += 1
        ws.write(r, 1, "2027 Implied total offline", self.f["rowlab_b"])
        ws.write_formula(r, 2, f"={fc_annual_cell}+{planned_cell}", self.f["calc_b"])
        r += 2

        # comparison block for the line chart: forecast + 2027 planned + 2024/25 actual unplanned
        self._band(ws, r, 1, 13, "Forecast vs Actuals (chart data)", "h_navy"); r += 1
        ws.write(r, 1, "Series", self.f["colhdr_l"])
        for j, m in enumerate(MONTHS):
            ws.write(r, 2 + j, m, self.f["colhdr"])
        chdr = r; r += 1
        # 2027 scenario (link to forecast row)
        ws.write(r, 1, "2027 Scenario (unplanned)", self.f["rowlab"])
        for j in range(12):
            ws.write_formula(r, 2 + j, f"={A1(fc_row,2+j)}", self.f["calc"])
        sc_chart_row = r; r += 1
        # 2027 planned monthly (from data)
        mp = self.ctx["monthly_planned"]
        ws.write(r, 1, "2027 Planned", self.f["rowlab"])
        for j, m in enumerate(MONTHS):
            v = float(mp.loc[2027, m]) if 2027 in mp.index else 0.0
            ws.write_number(r, 2 + j, v, self.f["calc"])
        pl_chart_row = r; r += 1
        mu = self.ctx["monthly_unplanned"]
        for yr in (2025, 2024):
            ws.write(r, 1, f"{yr} Unplanned (actual)", self.f["rowlab"])
            for j, m in enumerate(MONTHS):
                v = float(mu.loc[yr, m]) if yr in mu.index else 0.0
                ws.write_number(r, 2 + j, v, self.f["calc"])
            r += 1
        act_last = r - 1

        chart = self.wb.add_chart({"type": "line"})
        rows_for_chart = [(sc_chart_row, RED, 3.0), (pl_chart_row, GOLD, 2.0),
                          (pl_chart_row + 1, BLUE, 1.5), (pl_chart_row + 2, GRAY, 1.5)]
        for rr, color, w in rows_for_chart:
            chart.add_series({
                "name": ["Scenario 2027", rr, 1],
                "categories": ["Scenario 2027", chdr, 2, chdr, 13],
                "values": ["Scenario 2027", rr, 2, rr, 13],
                "line": {"color": color, "width": w,
                         "dash_type": "solid" if rr == sc_chart_row else
                                      ("dash" if rr == pl_chart_row else "solid")},
            })
        chart.set_title({"name": "2027 Scenario vs Planned vs Recent Actual Unplanned (kbd)"})
        chart.set_y_axis({"name": "kbd"})
        chart.set_legend({"position": "bottom"})
        chart.set_size({"width": 820, "height": 320})
        chart.set_chartarea({"border": {"none": True}})
        ws.insert_chart(r + 1, 1, chart)

        # PADD allocation mini-table (scenario annual x historical share)
        share = self.ctx["padd_share"]
        ar = r + 18
        self._band(ws, ar, 1, 3, "2027 Scenario - PADD Allocation (by historical unplanned share)", "h_green")
        ar += 1
        ws.write(ar, 1, "PADD", self.f["colhdr_l"])
        ws.write(ar, 2, "Share", self.f["colhdr"])
        ws.write(ar, 3, "Scenario kbd", self.f["colhdr"])
        ar += 1
        for p in PADDS:
            ws.write(ar, 1, p, self.f["rowlab"])
            ws.write_number(ar, 2, float(share[p]), self.f["pct_g"])
            ws.write_formula(ar, 3, f"={fc_annual_cell}*{A1(ar,2)}", self.f["calc_grn"])
            ar += 1

        # stash references for the sensitivity sheet
        self.sc_refs = {
            "anchor": f"{SHEET}!{anchor_cell}",
            "fc_annual": f"{SHEET}!{fc_annual_cell}",
            "growth": f"{SHEET}!{c_growth}",
            "mult": f"{SHEET}!{c_mult}",
        }

    # ================================================================= SENSITIVITY
    def sensitivity(self):
        ws = self.wb.add_worksheet("Sensitivity")
        self._setup(ws, "#C55A11")
        ws.set_column("A:A", 2)
        ws.set_column("B:B", 16)
        ws.set_column("C:K", 11)
        ws.merge_range("B2:K2", "Sensitivity & Risk", self.f["title"])
        ws.merge_range("B3:K3",
                       "Two-way heatmap recomputes from the Scenario baseline window. "
                       "Tornado ranks the drivers.", self.f["subtitle"])
        anchor = self.sc_refs["anchor"]

        # --- two-way heatmap: rows = growth, cols = multiplier ---
        growths = [-0.10, -0.05, 0.0, 0.05, 0.10, 0.15]
        mults = [0.7, 0.85, 1.0, 1.15, 1.3, 1.5]
        r = 5
        self._band(ws, r, 1, 1 + len(mults), "2027 Unplanned (kbd): Growth (rows) x Multiplier (cols)",
                   "h_orange"); r += 1
        ws.write(r, 1, "Growth \\ Mult", self.f["hm_hdr"])
        for j, m in enumerate(mults):
            ws.write_number(r, 2 + j, m, self.wb.add_format(
                {"font_name": "Arial", "bold": True, "font_color": WHITE, "bg_color": NAVY,
                 "align": "center", "valign": "vcenter", "border": 1, "num_format": MULT}))
        top = r; r += 1
        body_first = r
        for gi, g in enumerate(growths):
            ws.write_number(r, 1, g, self.wb.add_format(
                {"font_name": "Arial", "bold": True, "font_color": WHITE, "bg_color": NAVY,
                 "align": "center", "valign": "vcenter", "border": 1, "num_format": PCT}))
            for mj, m in enumerate(mults):
                # base case (g=0, m=1.0) gets a navy outline
                is_base = abs(g) < 1e-9 and abs(m - 1.0) < 1e-9
                fmt = self.f["hm_base"] if is_base else self.f["hm_cell"]
                gref = A1(top + 1 + gi, 1, col_abs=True)        # growth label cell
                mref = A1(top, 2 + mj, row_abs=True)            # mult header cell
                ws.write_formula(r, 2 + mj, f"={anchor}*(1+{gref})*{mref}", fmt)
            r += 1
        body_last = r - 1
        ws.conditional_format(body_first, 2, body_last, 1 + len(mults),
                              {"type": "3_color_scale",
                               "min_color": "#63BE7B", "mid_color": "#FFEB84",
                               "max_color": "#F8696B"})
        ws.write(r, 1, "Base case (0% / 1.0x) outlined; colors scale green→red across the grid.",
                 self.f["subtitle"])

        # --- tornado ---
        r += 3
        torn = self.ctx["tornado"]
        self._band(ws, r, 1, 6, "Tornado - 2027 Unplanned Sensitivity by Driver (kbd)", "h_orange")
        r += 1
        for j, h in enumerate(["Driver", "Low", "Base", "High", "Δ Low", "Δ High"]):
            ws.write(r, 1 + j, h, self.f["colhdr_l"] if j == 0 else self.f["colhdr"])
        r += 1
        t_first = r
        # plot biggest swing at top -> write in reverse so chart (bottom-up) shows largest on top
        for row in torn:
            ws.write(r, 1, row["driver"], self.f["rowlab"])
            ws.write_number(r, 2, row["low"], self.f["kbd"])
            ws.write_number(r, 3, row["base"], self.f["kbd"])
            ws.write_number(r, 4, row["high"], self.f["kbd"])
            ws.write_formula(r, 5, f"={A1(r,2)}-{A1(r,3)}", self.f["kbd"])   # negative
            ws.write_formula(r, 6, f"={A1(r,4)}-{A1(r,3)}", self.f["kbd"])   # positive
            r += 1
        t_last = r - 1

        tch = self.wb.add_chart({"type": "bar", "subtype": "stacked"})
        tch.add_series({
            "name": "Downside",
            "categories": ["Sensitivity", t_first, 1, t_last, 1],
            "values": ["Sensitivity", t_first, 5, t_last, 5],
            "fill": {"color": BLUE},
        })
        tch.add_series({
            "name": "Upside",
            "categories": ["Sensitivity", t_first, 1, t_last, 1],
            "values": ["Sensitivity", t_first, 6, t_last, 6],
            "fill": {"color": ORANGE},
        })
        tch.set_title({"name": "Tornado - Swing in 2027 Unplanned (kbd vs base)"})
        tch.set_x_axis({"name": "kbd vs base case"})
        tch.set_legend({"position": "bottom"})
        tch.set_size({"width": 720, "height": 320})
        tch.set_chartarea({"border": {"none": True}})
        ws.insert_chart(t_first, 8, tch)

    # ============================================================== MOGAS OVERLAY
    def mogas(self):
        ws = self.wb.add_worksheet("Mogas Overlay")
        self._setup(ws, GREEN)
        ws.set_column("A:A", 2)
        ws.set_column("B:B", 22)
        ws.set_column("C:F", 12)
        ws.set_column("H:H", 22)
        ws.merge_range("B2:F2", "Mogas-Equivalent Overlay (Secondary)", self.f["title"])
        ws.merge_range("B3:F3",
                       "Gasoline-equivalent = capacity x unit yield factor. Secondary view; "
                       "capacity is never discarded.", self.f["subtitle"])

        # yield map
        r = 5
        self._band(ws, r, 1, 4, "Yield Map: Unit Bucket → Mogas Factor", "h_green"); r += 1
        for j, h in enumerate(["Bucket", "Factor", "Unit categories"]):
            ws.write(r, 1 + j, h, self.f["colhdr_l"] if j != 1 else self.f["colhdr"])
        ws.set_column("D:D", 40)
        r += 1
        for bucket, factor, cats in self.ctx["mogas_yield_map"]:
            ws.write(r, 1, bucket, self.f["rowlab_b"])
            ws.write_number(r, 2, factor, self.wb.add_format(
                {"font_name": "Arial", "num_format": "0.000", "align": "center"}))
            ws.write(r, 3, ", ".join(c.title() for c in cats) or "-", self.f["rowlab"])
            r += 1

        # mogas annual
        r += 2
        ma = self.ctx["mogas_annual"]
        years = [y for y in ma.index if 2016 <= y <= 2027]
        self._band(ws, r, 1, 4, "Mogas-Equivalent Offline by Year (kbd)", "h_green"); r += 1
        for j, h in enumerate(["Year", "Planned", "Unplanned", "Total"]):
            ws.write(r, 1 + j, h, self.f["colhdr_l"] if j == 0 else self.f["colhdr"])
        r += 1
        first = r
        for y in years:
            partial = y in PARTIAL
            kf = self.f["kbd_p"] if partial else self.f["kbd"]
            ws.write_number(r, 1, y, self._yr_fmt(y))
            ws.write_number(r, 2, float(ma.loc[y, "Planned"]), kf)
            ws.write_number(r, 3, float(ma.loc[y, "Unplanned"]), kf)
            ws.write_formula(r, 4, f"={A1(r,2)}+{A1(r,3)}",
                             self.f["kbd_p"] if partial else self.f["kbd_b"])
            r += 1
        # stacked column
        chart = self.wb.add_chart({"type": "column", "subtype": "stacked"})
        for ci, (nm, color) in enumerate([("Planned", NAVY), ("Unplanned", GREEN)]):
            chart.add_series({
                "name": nm,
                "categories": ["Mogas Overlay", first, 1, r - 1, 1],
                "values": ["Mogas Overlay", first, 2 + ci, r - 1, 2 + ci],
                "fill": {"color": color},
            })
        chart.set_title({"name": "Mogas-Equivalent Offline (kbd)"})
        chart.set_y_axis({"name": "kbd (mogas-eq)"})
        chart.set_legend({"position": "bottom"})
        chart.set_size({"width": 560, "height": 320})
        chart.set_chartarea({"border": {"none": True}})
        ws.insert_chart(first, 6, chart)

    # ===================================================================== NOTES
    def notes(self):
        ws = self.wb.add_worksheet("Notes")
        self._setup(ws, "#7F7F7F", landscape=False)
        ws.set_column("A:A", 2)
        ws.set_column("B:B", 26)
        ws.set_column("C:H", 14)
        d = self.ctx["diag"]
        y0, y1 = d["years"]
        ws.merge_range("B2:H2", "Data Notes & Methodology", self.f["title"])
        items = [
            ("Source", f"Snowflake export, sheet Query1. {d['rows']:,} rows, {y0}-{y1}."),
            ("Primary metric", "CAP_OFFLINE_ADJUSTED_KBD - offline capacity, thousand bbl/day, all units."),
            ("Event count", f"{d['events_distinct']:,} distinct OUTAGE_IDs; rows are unit-month slices."),
            ("Outage type", "Binary {PLANNED, UNPLANNED}; UNKNOWN folded into UNPLANNED (desk rule)."),
            ("PADD", "Parsed from PAD_DIST Roman numerals (100% resolved); STATE map is fallback."),
            ("Partial years", "2026 & 2027 are incomplete/special - grey italic, never final."),
            ("2027 guardrail", "2027 is planned-only; unplanned-2027 is a scenario. Non-planned "
                               "comparisons vs 2027 show n/a."),
            ("Duration caveat", "TOTAL_OUTAGE_DAYS is monthly-allocated intensity (caps ~31), not full "
                                "event length."),
            ("Outliers", "2020-2021 (COVID / Winter Storm Uri) excluded from forecast baselines by default."),
            ("Scenario method", "Baseline(window) x (1+growth) x multiplier + one-off(stress month). "
                                 "All cells are live formulas off the yellow inputs."),
            ("Sensitivity method", "Heatmap = anchor x (1+growth) x multiplier across a grid (recalcs "
                                    "from the scenario window). Tornado ranks driver swings."),
            ("Mogas method", "Mogas-eq = capacity x bucket yield (CDU .175, FCC .65, Ref .85, HDC .05, "
                             "Coker .20, else 0). Additive overlay; capacity never discarded."),
            ("Refresh", "Set INPUT_PATH (or pass a path arg) and re-run build_workbook.py. Idempotent."),
        ]
        r = 4
        for lab, val in items:
            ws.write(r, 1, lab, self.f["note_b"])
            ws.merge_range(r, 2, r, 7, val, self.f["note"])
            ws.set_row(r, 30)
            r += 1
        r += 1
        self._band(ws, r, 1, 7, "Color Key (financial-model convention)", "h_navy"); r += 1
        keys = [
            ("#0000FF", "Blue font", "Hard-coded input the user changes"),
            ("#000000", "Black font", "Formula / calculation"),
            ("#008000", "Green font", "Cross-sheet link"),
            (YELLOW, "Yellow fill", "Key assumption cell"),
        ]
        for color, name, desc in keys:
            ws.write(r, 1, name, self.wb.add_format(
                {"font_name": "Arial", "bold": True, "font_color": color if color != YELLOW else "#000000",
                 "bg_color": YELLOW if color == YELLOW else WHITE}))
            ws.merge_range(r, 2, r, 7, desc, self.f["note"])
            r += 1

    # ===================================================================== run
    def run(self):
        self.cover()
        self.dashboard()
        self.summary()
        self.monthly()
        self.padd_charts()
        self.padd_detail()
        self.units()
        self.refinery_detail()
        self.scenario()
        self.sensitivity()      # depends on scenario refs
        self.mogas()
        self.notes()
        self.wb.close()


def main():
    ap = argparse.ArgumentParser(description="Institutional refinery-outage Excel workbook")
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
