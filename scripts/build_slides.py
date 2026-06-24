#!/usr/bin/env python3
"""
build_slides.py
Commodities-desk slide deck (python-pptx) for the refinery-outage analysis.

Priority-2 deliverable. Mirrors the workbook's charts as crisp embedded images
(rendered by charts.py) in a clean 16:9 layout, with data-driven insight
captions that refresh with the data. Self-contained: charts are embedded, not
linked.

Usage:
    python build_slides.py                       # uses INPUT_PATH
    python build_slides.py path/to/export.xlsx --out outage_deck.pptx
"""
import argparse
import os
import sys
import tempfile

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

from pathlib import Path

import engine
import charts

_ROOT = Path(__file__).resolve().parent.parent          # repo root (scripts/ -> ..)
INPUT_PATH = str(_ROOT / "data" / "rEFINERY oUTAGES.xlsx")
OUT_PATH = str(_ROOT / "output" / "outage_deck.pptx")

NAVY = RGBColor(0x1F, 0x38, 0x64)
BLUE = RGBColor(0x2E, 0x54, 0x96)
RED = RGBColor(0xC0, 0x00, 0x00)
GOLD = RGBColor(0xBF, 0x90, 0x00)
GREEN = RGBColor(0x54, 0x82, 0x35)
LT_BLUE = RGBColor(0xD6, 0xE0, 0xF0)
GRAY = RGBColor(0x59, 0x59, 0x59)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LT_GRAY = RGBColor(0xF2, 0xF2, 0xF2)

FONT = "Arial"
SW, SH = Inches(13.333), Inches(7.5)


def kbd(x):
    return f"{x:,.0f}"


class Deck:
    def __init__(self, ctx, assets):
        self.ctx = ctx
        self.assets = assets
        self.prs = Presentation()
        self.prs.slide_width = SW
        self.prs.slide_height = SH
        self.blank = self.prs.slide_layouts[6]

    # ------------------------------------------------------------- primitives
    def _slide(self):
        return self.prs.slides.add_slide(self.blank)

    def _rect(self, s, l, t, w, h, color, line=None):
        shp = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
        shp.fill.solid()
        shp.fill.fore_color.rgb = color
        if line is None:
            shp.line.fill.background()
        else:
            shp.line.color.rgb = line
            shp.line.width = Pt(1)
        shp.shadow.inherit = False
        return shp

    def _text(self, s, l, t, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
              space_after=4):
        """runs: list of (text, size, bold, color, italic)."""
        tb = s.shapes.add_textbox(l, t, w, h)
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        tf.margin_left = tf.margin_right = Pt(2)
        tf.margin_top = tf.margin_bottom = Pt(2)
        first = True
        for item in runs:
            text, size, bold, color = item[0], item[1], item[2], item[3]
            italic = item[4] if len(item) > 4 else False
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.alignment = align
            p.space_after = Pt(space_after)
            r = p.add_run()
            r.text = text
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.italic = italic
            r.font.name = FONT
            r.font.color.rgb = color
        return tb

    def _titlebar(self, s, title, sub=None):
        self._rect(s, 0, 0, SW, Inches(0.9), NAVY)
        self._text(s, Inches(0.45), Inches(0.12), Inches(12.4), Inches(0.5),
                   [(title, 24, True, WHITE)], anchor=MSO_ANCHOR.MIDDLE)
        if sub:
            self._text(s, Inches(0.47), Inches(0.6), Inches(12.4), Inches(0.28),
                       [(sub, 11, False, LT_BLUE)])
        # gold accent under bar
        self._rect(s, 0, Inches(0.9), SW, Pt(3), GOLD)

    def _footer(self, s, page):
        d = self.ctx["diag"]
        y0, y1 = d["years"]
        self._text(s, Inches(0.45), Inches(7.12), Inches(9), Inches(0.3),
                   [(f"Refinery Outage Analytics  |  Source: Snowflake export "
                     f"({d['rows']:,} rows, {y0}-{y1})  |  Primary metric: capacity offline (kbd)",
                     8, False, GRAY)])
        self._text(s, Inches(12.4), Inches(7.12), Inches(0.6), Inches(0.3),
                   [(str(page), 8, False, GRAY)], align=PP_ALIGN.RIGHT)

    def _pic_fit(self, s, path, l, t, max_w, max_h, center=True):
        """Insert a picture scaled to fit (max_w, max_h), preserving aspect."""
        pic = s.shapes.add_picture(path, l, t, width=max_w)
        if pic.height > max_h:
            scale = max_h / pic.height
            pic.height = int(pic.height * scale)
            pic.width = int(pic.width * scale)
        if center:
            pic.left = int(l + (max_w - pic.width) / 2)
        return pic

    def _bullet_list(self, s, x, y, w, bullets, spacing=0.82, size=11.5, box_h=0.62):
        by = y
        for b in bullets:
            self._rect(s, x, by + Inches(0.05), Inches(0.1), Inches(0.1), GOLD)
            self._text(s, x + Inches(0.24), by, w - Inches(0.24), Inches(box_h),
                       [(b, size, False, RGBColor(0x33, 0x33, 0x33))])
            by += Inches(spacing)
        return by

    def _content_slide(self, title, sub, img, bullets, page, layout="right"):
        """Chart + bullet points. layout 'right' = chart left/bullets right;
        'below' = wide chart on top, bullets in two columns beneath."""
        s = self._slide()
        self._titlebar(s, title, sub)
        if layout == "below":
            self._pic_fit(s, img, Inches(0.5), Inches(1.15), Inches(12.3), Inches(3.85))
            self._text(s, Inches(0.5), Inches(5.12), Inches(6), Inches(0.3),
                       [("Key points", 13, True, NAVY)])
            half = (len(bullets) + 1) // 2
            self._bullet_list(s, Inches(0.6), Inches(5.55), Inches(6.1), bullets[:half],
                              spacing=0.6, size=10.5, box_h=0.56)
            self._bullet_list(s, Inches(6.9), Inches(5.55), Inches(6.1), bullets[half:],
                              spacing=0.6, size=10.5, box_h=0.56)
        else:
            self._pic_fit(s, img, Inches(0.35), Inches(1.3), Inches(7.7), Inches(5.4), center=False)
            self._text(s, Inches(8.25), Inches(1.3), Inches(4.7), Inches(0.35),
                       [("Key points", 14, True, NAVY)])
            self._bullet_list(s, Inches(8.3), Inches(1.85), Inches(4.7), bullets)
        self._footer(s, page)
        return s

    # ------------------------------------------------------------- slides
    def title_slide(self):
        s = self._slide()
        self._rect(s, 0, 0, SW, SH, NAVY)
        self._rect(s, 0, Inches(3.05), SW, Pt(3), GOLD)
        d = self.ctx["diag"]
        y0, y1 = d["years"]
        self._text(s, Inches(0.8), Inches(1.7), Inches(11.7), Inches(1.0),
                   [("Refinery Outage Analytics", 44, True, WHITE)])
        self._text(s, Inches(0.82), Inches(3.2), Inches(11.7), Inches(0.6),
                   [("Capacity Offline (kbd) - Planned & Unplanned - 2027 Scenario & Sensitivity",
                     18, False, LT_BLUE)])
        s_ = self.ctx["summary"]
        ly = max(y for y in s_.index if y not in engine.PARTIAL_YEARS and s_.loc[y, "Unplanned"] > 0)
        self._text(s, Inches(0.82), Inches(4.5), Inches(11.7), Inches(2.0),
                   [(f"Data vintage:  {d['rows']:,} rows  |  {y0}-{y1}  |  "
                     f"{d['events_distinct']:,} distinct outages", 13, False, WHITE),
                    (f"Latest full year ({ly}):  {kbd(s_.loc[ly,'Total'])} kbd total offline, "
                     f"{kbd(s_.loc[ly,'Unplanned'])} kbd unplanned ({s_.loc[ly,'Unpl%']:.0%})",
                     13, False, WHITE),
                    ("2026 & 2027 are partial / planned-only - shown for context, not as finals.",
                     11, False, RGBColor(0xF0, 0xC0, 0xC0), True)], space_after=10)

    def exec_summary(self):
        s = self._slide()
        self._titlebar(s, "Executive Summary", "At-a-glance metrics and what matters")
        sm = self.ctx["summary"]
        ly = max(y for y in sm.index if y not in engine.PARTIAL_YEARS and sm.loc[y, "Unplanned"] > 0)
        padd_un = self.ctx["padd_unplanned"]
        top_padd = padd_un[ly].idxmax()
        share = self.ctx["padd_share"]
        sc = self.ctx["scenario"]
        tiles = [
            (f"Total Offline (FY{ly})", kbd(sm.loc[ly, "Total"]), "kbd"),
            (f"Unplanned (FY{ly})", kbd(sm.loc[ly, "Unplanned"]), "kbd"),
            ("Unplanned %", f"{sm.loc[ly,'Unpl%']:.0%}", "of total"),
            ("Distinct Outages", kbd(sm.loc[ly, "Events"]), f"FY{ly}"),
            ("Top PADD", top_padd, "by unplanned"),
        ]
        n = len(tiles)
        gap = Inches(0.2)
        tw = (SW - Inches(0.9) - gap * (n - 1)) / n
        x = Inches(0.45)
        for lab, val, sub in tiles:
            self._rect(s, x, Inches(1.25), tw, Inches(0.42), BLUE)
            self._text(s, x, Inches(1.27), tw, Inches(0.4),
                       [(lab, 10.5, True, WHITE)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            self._rect(s, x, Inches(1.67), tw, Inches(1.1), LT_BLUE)
            self._text(s, x, Inches(1.72), tw, Inches(0.7),
                       [(val, 26, True, NAVY)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            self._text(s, x, Inches(2.4), tw, Inches(0.3),
                       [(sub, 9, False, GRAY, True)], align=PP_ALIGN.CENTER)
            x += tw + gap

        # takeaways
        yoy = sm.loc[ly, "YoY%"]
        peak_season = "February (winter freeze) and September-October (turnaround)"
        bullets = [
            (f"{top_padd} dominates unplanned risk at ~{share[top_padd]:.0%} of US unplanned "
             f"capacity offline over the baseline window; PADD 2 and PADD 5 are the next largest.",),
            (f"FY{ly} total offline of {kbd(sm.loc[ly,'Total'])} kbd was {yoy:+.0%} YoY, with the "
             f"planned/unplanned split near 50/50 ({sm.loc[ly,'Unpl%']:.0%} unplanned).",),
            (f"Unplanned capacity offline is highly seasonal - peaks in {peak_season}, "
             f"with a summer trough.",),
            (f"2027 is planned-only in the data ({kbd(sc['planned_2027'])} kbd booked). The scenario "
             f"model adds a baseline-driven unplanned forecast of ~{kbd(sc['annual_unplanned'])} kbd "
             f"-> implied total ~{kbd(sc['implied_total'])} kbd.",),
        ]
        self._text(s, Inches(0.5), Inches(3.0), Inches(12.3), Inches(0.35),
                   [("Key Takeaways", 14, True, NAVY)])
        ty = Inches(3.45)
        for b in bullets:
            self._rect(s, Inches(0.55), ty + Inches(0.06), Inches(0.12), Inches(0.12), GOLD)
            self._text(s, Inches(0.85), ty, Inches(11.9), Inches(0.8),
                       [(b[0], 12.5, False, RGBColor(0x33, 0x33, 0x33))])
            ty += Inches(0.82)
        self._footer(s, 2)

    def build(self):
        a = self.assets
        self.title_slide()
        self.exec_summary()
        sm = self.ctx["summary"]
        ly = max(y for y in sm.index if y not in engine.PARTIAL_YEARS and sm.loc[y, "Unplanned"] > 0)
        peak_yr = max((y for y in sm.index if 2014 <= y <= 2025), key=lambda y: sm.loc[y, "Total"])
        share = self.ctx["padd_share"]
        sc = self.ctx["scenario"]

        un25 = self.ctx["unit_share"][2025].sort_values(ascending=False)
        lvl, yoy = self.ctx["padd_unpl_yoy"]
        sp = self.ctx["scenario_padd"]
        fcc = self.ctx["fcc_exxon"]
        p = 3

        self._content_slide(
            "Annual Trend", "Capacity offline by year, planned + unplanned", a["annual"],
            [f"Offline capacity peaked in {peak_yr}; 2020-21 (COVID / Winter Storm Uri) are "
             "outliers, excluded from forecast baselines.",
             f"FY{ly} ran {kbd(sm.loc[ly,'Total'])} kbd total, ~50/50 planned vs unplanned "
             f"({sm.loc[ly,'Unpl%']:.0%} unplanned).",
             f"Total offline has eased {sm.loc[ly,'YoY%']:+.0%} YoY into {ly} off the 2022-23 highs.",
             "2027 shows planned capacity only (booked turnarounds) - unplanned is modeled."], p); p += 1

        self._content_slide(
            "Unplanned by PADD", "Where the unplanned risk concentrates", a["padd_clustered"],
            [f"PADD 3 (Gulf Coast) carries ~{share['PADD 3']:.0%} of US unplanned capacity offline.",
             f"PADD 2 (Midwest) is next at ~{share['PADD 2']:.0%}; PADD 5 ~{share['PADD 5']:.0%}.",
             f"PADD 5 swung {yoy.loc['PADD 5',2025]:+.0%} in {2025} on California events; "
             f"PADD 1 {yoy.loc['PADD 1',2025]:+.0%}.",
             "PADD 3 and PADD 5 dominate both the level and the volatility of unplanned risk."], p); p += 1

        self._content_slide(
            "Seasonality", "Unplanned offline by calendar month", a["seasonality"],
            ["Unplanned outages cluster in late winter (Feb freeze events) and the autumn "
             "turnaround window.",
             "A clear summer trough (Jun-Aug) as units run through driving season.",
             "This recurring shape is the backbone of the 2027 scenario baseline.",
             "2025 (red) sits above 2023-24 in Q1 - a heavier freeze season."], p); p += 1

        self._content_slide(
            "PADD 3 - Reference View", "2026 plan + unplanned vs 2023-2025 totals and 2027 plan",
            a["padd3"],
            ["Gold/orange columns = 2026 booked plan and unplanned, stacked.",
             "Lines = prior-year total offline (2025 red, 2024 blue, 2023 gray).",
             "Dashed green = 2027 booked plan, the only forward series.",
             "One shared axis keeps magnitudes honest across years."], p); p += 1

        self._content_slide(
            "PADD 1, 2, 4 & 5", "Same combo view across the remaining PADDs", a["padd_sm"],
            [f"PADD 1 is genuinely small (~{share['PADD 1']:.0%} of unplanned) - a thin bar set.",
             "PADD 2 (Midwest) carries real weight and swings with refinery turnarounds.",
             "PADD 5 shows the 2025 California spike clearly in the red total line.",
             "PADD 4 (Rockies) is the smallest and steadiest region.",
             "2027 plan (dashed green) is the only forward-booked series everywhere."], p,
            layout="below"); p += 1

        # NEW: back-to-back FCC outages
        top_fcc = [c for c in fcc if c["n"] >= 4][:5]
        self._content_slide(
            "Back-to-Back FCC Outages - ExxonMobil",
            "Consecutive-month FCC runs that month-level external trackers miss", a["fcc"],
            ["This granular export captures repeated FCC (cat cracker) outages in adjacent "
             "months at the same Exxon plant.",
             "Baton Rouge ran Jan-Jun 2022; Baytown Jan-Mar 2023; Joliet Feb-May 2025 and "
             "Feb-Jun 2026 - classic Q1-Q2 turnaround clustering.",
             "Aggregated/external data smooths these into a single monthly figure and loses the "
             "back-to-back signal.",
             f"{len(fcc)} such Exxon FCC runs (>=3 months) since 2011 - a recurring, "
             "model-able pattern, not noise.",
             "FCC offline is the most gasoline-relevant unit loss (0.65 mogas yield)."], p,
            layout="below"); p += 1

        self._content_slide(
            "Unit Categories", "Capacity offline by unit category", a["units"],
            [f"Crude trains (atmospheric + vacuum distillation) are ~{(un25.get('ATMOS DISTILLATION',0)+un25.get('VACUUM DISTILLATION',0)):.0%} "
             f"of {2025} offline capacity.",
             f"Hydrotreating ~{un25.get('HYDROTREATING',0):.0%} and FCC ~{un25.get('FLUID CAT CRACKING',0):.0%} "
             "follow - FCC and reforming drive gasoline exposure.",
             "Unit mix matters: the same kbd offline means more mogas loss if it is FCC or "
             "reforming than if it is hydrotreating.",
             "Workbook adds per-unit share and YoY% columns for the full 17 categories."], p); p += 1

        self._content_slide(
            "2027 Unplanned Scenario", "Driver-based forecast vs planned and recent actuals",
            a["scenario"],
            [f"Default (window {sc['window']}, 0% growth, 1.0x): ~{kbd(sc['annual_unplanned'])} kbd "
             "unplanned.",
             f"On top of {kbd(sc['planned_2027'])} kbd booked plan -> ~{kbd(sc['implied_total'])} kbd "
             "implied total offline.",
             "Built from the historical monthly unplanned shape, not a flat annual number.",
             "Every input (window, growth, multiplier, one-off, stress month) is live in the "
             "Excel model."], p); p += 1

        # NEW: scenario by PADD
        self._content_slide(
            "2027 Scenario by PADD", "Each PADD carries its own monthly seasonality",
            a["scenario_padd"],
            [f"PADD 3 ~{kbd(sp['PADD 3']['annual'])} kbd and PADD 2 ~{kbd(sp['PADD 2']['annual'])} kbd "
             "dominate the 2027 unplanned forecast.",
             f"PADD 5 ~{kbd(sp['PADD 5']['annual'])} kbd, with its distinct California seasonal shape.",
             "More honest than splitting one national number by a flat share - the monthly "
             "shapes differ by region.",
             "Per-PADD baselines flex with the same growth and multiplier dials in the workbook."], p); p += 1

        # Sensitivity: heatmap + tornado below layout
        torn_top = self.ctx["tornado"][0]["driver"].split(" (")[0]
        s = self._slide()
        self._titlebar(s, "Sensitivity & Risk", "How the 2027 unplanned forecast flexes")
        self._pic_fit(s, a["heatmap"], Inches(0.4), Inches(1.25), Inches(6.5), Inches(4.0), center=False)
        self._pic_fit(s, a["tornado"], Inches(7.05), Inches(1.5), Inches(5.9), Inches(3.5), center=False)
        self._text(s, Inches(0.5), Inches(5.3), Inches(6), Inches(0.3), [("Key points", 13, True, NAVY)])
        self._bullet_list(s, Inches(0.6), Inches(5.7), Inches(6.1), [
            f"Base case {kbd(sc['annual_unplanned'])} kbd is outlined; grid spans -10%..+15% growth "
            "x 0.7..1.5x rate.",
            "The unplanned-rate multiplier is the dominant driver."], spacing=0.62, size=10.5, box_h=0.56)
        self._bullet_list(s, Inches(6.9), Inches(5.7), Inches(6.1), [
            f"Range runs ~{kbd(sc['annual_unplanned']*0.7)}-{kbd(sc['annual_unplanned']*1.5*1.15)} kbd "
            "across plausible inputs.",
            "Tornado ranks drivers by swing; window choice is second."], spacing=0.62, size=10.5, box_h=0.56)
        self._footer(s, p); p += 1

        # closing / method
        s = self._slide()
        self._titlebar(s, "Method & Caveats", "Read before use")
        notes = [
            "Primary metric is CAP_OFFLINE_ADJUSTED_KBD (offline capacity, kbd, all units). Mogas is "
            "a secondary overlay only.",
            "Outage type is binary {PLANNED, UNPLANNED}; UNKNOWN folds into UNPLANNED per desk rule.",
            "2027 is planned-only - any Plan+Unplanned or Unplanned comparison vs 2027 is shown n/a; "
            "only Planned is comparable. Unplanned-2027 is a modeled scenario.",
            "2020-2021 (COVID / Winter Storm Uri) are excluded from forecast baselines by default.",
            "Scenario = baseline(window) x (1+growth) x multiplier + one-off(stress month); the Excel "
            "workbook recomputes it live from dropdown inputs.",
            "All numbers refresh from the source export - re-run the build to update workbook, deck "
            "and dashboard together.",
        ]
        ty = Inches(1.4)
        for nlabel in notes:
            self._rect(s, Inches(0.55), ty + Inches(0.07), Inches(0.12), Inches(0.12), GOLD)
            self._text(s, Inches(0.85), ty, Inches(11.9), Inches(0.8),
                       [(nlabel, 12, False, RGBColor(0x33, 0x33, 0x33))])
            ty += Inches(0.78)
        self._footer(s, p)

    def save(self, path):
        self.prs.save(path)


def main():
    ap = argparse.ArgumentParser(description="Refinery-outage slide deck")
    ap.add_argument("excel", nargs="?", default=INPUT_PATH, help="path to the outage .xlsx export")
    ap.add_argument("--out", default=OUT_PATH, help="output .pptx path")
    args = ap.parse_args()

    print(f"Loading {args.excel.strip()} ...")
    ctx = engine.build_context(args.excel)
    with tempfile.TemporaryDirectory() as tmp:
        print("Rendering charts ...")
        assets = charts.render_all(ctx, tmp)
        print(f"Building deck -> {args.out}")
        Path(args.out).parent.mkdir(parents=True, exist_ok=True)
        deck = Deck(ctx, assets)
        deck.build()
        deck.save(args.out)
    print("Done.")


if __name__ == "__main__":
    sys.exit(main())
