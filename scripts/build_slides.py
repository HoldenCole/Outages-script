#!/usr/bin/env python3
"""
build_slides.py
Sell-side "weekly meeting" deck (python-pptx) styled to the reference template:
navy title slide, white content slides with a red section header, dense
multi-chart grids, combo-charts + bullets, and full-width turnaround-schedule
tables. Charts are embedded (self-contained).

Brand mark: set BRAND_LOGO to a logo image path to drop in your own logo;
otherwise BRAND_TEXT is rendered in the corner. The reference is an ExxonMobil
deck - this reproduces the layout/'look', not the trademarked logo.

Usage:
    python build_slides.py                       # uses INPUT_PATH
    python build_slides.py path/to/export.xlsx --out outage_deck.pptx
"""
import argparse
import os
import sys
import tempfile
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

import engine
import charts

_ROOT = Path(__file__).resolve().parent.parent
INPUT_PATH = str(_ROOT / "data" / "rEFINERY oUTAGES.xlsx")
OUT_PATH = str(_ROOT / "output" / "outage_deck.pptx")

# Brand mark (swap in your own logo to match exactly)
BRAND_TEXT = "Products Trading"
BRAND_LOGO = None

NAVY = RGBColor(0x1F, 0x38, 0x64)
NAVY2 = RGBColor(0x2E, 0x54, 0x96)
RED = RGBColor(0xC0, 0x00, 0x00)
GOLD = RGBColor(0xBF, 0x90, 0x00)
GREEN = RGBColor(0x54, 0x82, 0x35)
LT_BLUE = RGBColor(0xD6, 0xE0, 0xF0)
GRAY = RGBColor(0x59, 0x59, 0x59)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LT_GRAY = RGBColor(0xF2, 0xF2, 0xF2)
INK = RGBColor(0x26, 0x2B, 0x33)

FONT = "Arial"
SW, SH = Inches(13.333), Inches(7.5)


def kbd(x):
    return f"{x:,.0f}"


def _d(ts):
    try:
        return f"{ts.month}/{ts.day}/{str(ts.year)[2:]}"
    except Exception:
        return ""


class Deck:
    def __init__(self, ctx, assets, asof="June 15th, 2026"):
        self.ctx = ctx
        self.a = assets
        self.asof = asof
        self.prs = Presentation()
        self.prs.slide_width = SW
        self.prs.slide_height = SH
        self.blank = self.prs.slide_layouts[6]
        self.page = 0

    # ----------------------------------------------------------------- primitives
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
            shp.line.width = Pt(0.75)
        shp.shadow.inherit = False
        return shp

    def _text(self, s, l, t, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, sa=3):
        tb = s.shapes.add_textbox(l, t, w, h)
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
            setattr(tf, m, Pt(1))
        first = True
        for item in runs:
            text, size, bold, color = item[0], item[1], item[2], item[3]
            italic = item[4] if len(item) > 4 else False
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.alignment = align
            p.space_after = Pt(sa)
            r = p.add_run()
            r.text = text
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.italic = italic
            r.font.name = FONT
            r.font.color.rgb = color
        return tb

    def _brand(self, s, l, t, h=Inches(0.32)):
        if BRAND_LOGO and os.path.exists(BRAND_LOGO):
            s.shapes.add_picture(BRAND_LOGO, l, t, height=h)
        else:
            self._text(s, l, t, Inches(3), h, [(BRAND_TEXT, 13, True, RED)],
                       anchor=MSO_ANCHOR.MIDDLE)

    def _pic_fit(self, s, path, l, t, max_w, max_h, center="x"):
        pic = s.shapes.add_picture(path, l, t, width=max_w)
        if pic.height > max_h:
            sc = max_h / pic.height
            pic.height = int(pic.height * sc)
            pic.width = int(pic.width * sc)
        if center == "x":
            pic.left = int(l + (max_w - pic.width) / 2)
        elif center == "both":
            pic.left = int(l + (max_w - pic.width) / 2)
            pic.top = int(t + (max_h - pic.height) / 2)
        return pic

    # ----------------------------------------------------------------- chrome
    def _section(self, title, sub=None):
        """White content slide with a red section header top-left."""
        s = self._slide()
        self.page += 1
        self._text(s, Inches(0.45), Inches(0.18), Inches(11.5), Inches(0.7),
                   [(title, 30, False, RED)], anchor=MSO_ANCHOR.MIDDLE)
        if sub:
            self._text(s, Inches(0.5), Inches(0.92), Inches(11.5), Inches(0.3),
                       [(sub, 12, False, GRAY)])
        self._brand(s, Inches(0.45), Inches(6.98))
        self._text(s, Inches(12.5), Inches(7.0), Inches(0.7), Inches(0.3),
                   [(str(self.page), 9, False, GRAY)], align=PP_ALIGN.RIGHT)
        return s

    def _footnote(self, s, text):
        self._text(s, Inches(0.5), Inches(6.74), Inches(11.5), Inches(0.24),
                   [(text, 8, False, GRAY, True)])

    def _bullets(self, s, x, y, w, bullets, size=12, spacing=0.52, head="Key takeaways"):
        if head:
            self._text(s, x, y, w, Inches(0.3), [(head + ":", size + 1.5, True, NAVY)])
            y = y + Inches(0.45)
        for b in bullets:
            sub = b.startswith("- ")
            txt = b[2:] if sub else b
            bx = x + (Inches(0.45) if sub else Inches(0.0))
            self._rect(s, bx, y + Inches(0.07), Inches(0.09), Inches(0.09),
                       GOLD if not sub else GRAY)
            self._text(s, bx + Inches(0.22), y, w - Inches(0.22) - (bx - x), Inches(0.7),
                       [(txt, size if not sub else size - 1.5, False, INK)])
            y += Inches(spacing if not sub else spacing - 0.06)
        return y

    # ----------------------------------------------------------------- slides
    def title_slide(self):
        s = self._slide()
        self._rect(s, 0, 0, SW, SH, NAVY)
        self._rect(s, 0, Inches(1.05), SW, Inches(0.02), NAVY2)
        self._rect(s, 0, Inches(3.05), SW, Pt(2.5), GOLD)
        self._text(s, Inches(0.7), Inches(0.45), Inches(8), Inches(0.4),
                   [(self.asof, 13, False, LT_BLUE)])
        # brand top-right
        if BRAND_LOGO and os.path.exists(BRAND_LOGO):
            s.shapes.add_picture(BRAND_LOGO, Inches(10.3), Inches(0.4), height=Inches(0.5))
        else:
            self._text(s, Inches(9.3), Inches(0.4), Inches(3.5), Inches(0.5),
                       [(BRAND_TEXT, 18, True, WHITE)], align=PP_ALIGN.RIGHT,
                       anchor=MSO_ANCHOR.MIDDLE)
        self._text(s, Inches(0.7), Inches(3.35), Inches(11.8), Inches(1.6),
                   [("Refinery Outage Analytics", 40, False, WHITE),
                    ("Weekly Gasoline / Mogas Meeting", 26, False, LT_BLUE)], sa=8)
        d = self.ctx["diag"]
        self._text(s, Inches(0.72), Inches(5.5), Inches(11), Inches(0.4),
                   [(f"Capacity offline (kbd) - planned & unplanned - 2027 scenario   |   "
                     f"{d['rows']:,} rows, {d['years'][0]}-{d['years'][1]}", 12, False, LT_BLUE)])
        self._text(s, Inches(11.4), Inches(7.0), Inches(1.7), Inches(0.3),
                   [("PROPRIETARY", 9, False, LT_BLUE)], align=PP_ALIGN.RIGHT)

    def exec_summary(self):
        s = self._section("Executive Summary", "At-a-glance metrics and what matters this week")
        sm = self.ctx["summary"]
        ly = max(y for y in sm.index if y not in engine.PARTIAL_YEARS and sm.loc[y, "Unplanned"] > 0)
        padd_un = self.ctx["padd_unplanned"]
        top_padd = padd_un[ly].idxmax()
        share = self.ctx["padd_share"]
        sc = self.ctx["scenario"]
        tiles = [(f"Total Offline FY{ly}", kbd(sm.loc[ly, "Total"]), "kbd"),
                 (f"Unplanned FY{ly}", kbd(sm.loc[ly, "Unplanned"]), "kbd"),
                 ("Unplanned %", f"{sm.loc[ly,'Unpl%']:.0%}", "of total"),
                 ("Distinct Outages", kbd(sm.loc[ly, "Events"]), f"FY{ly}"),
                 ("Top PADD", top_padd, "unplanned")]
        n = len(tiles)
        gap = Inches(0.18)
        tw = (SW - Inches(0.9) - gap * (n - 1)) / n
        x = Inches(0.45)
        for lab, val, sub in tiles:
            self._rect(s, x, Inches(1.35), tw, Inches(0.4), NAVY2)
            self._text(s, x, Inches(1.37), tw, Inches(0.36), [(lab, 10.5, True, WHITE)],
                       align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            self._rect(s, x, Inches(1.75), tw, Inches(1.0), LT_BLUE)
            self._text(s, x, Inches(1.8), tw, Inches(0.62), [(val, 25, True, NAVY)],
                       align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            self._text(s, x, Inches(2.42), tw, Inches(0.28), [(sub, 9, False, GRAY, True)],
                       align=PP_ALIGN.CENTER)
            x += tw + gap
        self._bullets(s, Inches(0.5), Inches(3.2), Inches(12.3), [
            f"{top_padd} carries ~{share[top_padd]:.0%} of US unplanned capacity offline; "
            f"PADD 2 ~{share['PADD 2']:.0%} and PADD 5 ~{share['PADD 5']:.0%} are the next watch regions.",
            f"FY{ly} total offline {kbd(sm.loc[ly,'Total'])} kbd, {sm.loc[ly,'YoY%']:+.0%} YoY, "
            f"a roughly even planned/unplanned split ({sm.loc[ly,'Unpl%']:.0%} unplanned).",
            "Unplanned offline is highly seasonal - Feb freeze and the autumn turnaround window, "
            "with a summer trough.",
            "Recurring back-to-back FCC turnarounds at the ExxonMobil plants (Baton Rouge, Baytown, "
            "Joliet) are visible in this data but washed out of month-level external trackers.",
            f"2027 is planned-only ({kbd(sc['planned_2027'])} kbd booked); the scenario adds "
            f"~{kbd(sc['annual_unplanned'])} kbd modeled unplanned -> ~{kbd(sc['implied_total'])} kbd implied total.",
        ], size=12.5, spacing=0.62, head=None)

    def grid4_slide(self, title, sub, imgs, takeaways, foot=None):
        s = self._section(title, sub)
        # 2x2 grid in the upper area, takeaways line at the bottom
        cells = [(Inches(0.45), Inches(1.28)), (Inches(6.85), Inches(1.28)),
                 (Inches(0.45), Inches(3.62)), (Inches(6.85), Inches(3.62))]
        for img, (cx, cy) in zip(imgs, cells):
            self._pic_fit(s, img, cx, cy, Inches(6.0), Inches(2.28), center="both")
        self._bullets(s, Inches(0.5), Inches(6.12), Inches(12.3), takeaways, size=11,
                      spacing=0.32, head=None)
        if foot:
            self._footnote(s, foot)
        return s

    def charts_bullets_slide(self, title, sub, imgs, bullets, foot=None):
        """1-2 charts stacked on the left, bullets on the right (TA-style)."""
        s = self._section(title, sub)
        if len(imgs) == 1:
            self._pic_fit(s, imgs[0], Inches(0.45), Inches(1.45), Inches(7.4), Inches(5.0),
                          center="both")
        else:
            self._pic_fit(s, imgs[0], Inches(0.45), Inches(1.35), Inches(7.4), Inches(2.55),
                          center="both")
            self._pic_fit(s, imgs[1], Inches(0.45), Inches(4.05), Inches(7.4), Inches(2.55),
                          center="both")
        self._bullets(s, Inches(8.15), Inches(1.5), Inches(4.85), bullets, size=12, spacing=0.62)
        if foot:
            self._footnote(s, foot)
        return s

    def ta_table_slide(self, title, df, note=None, maxrows=15):
        s = self._section(title, "Planned turnaround schedule - offline capacity, kbd")
        heads = ["Company", "Refinery", "PADD", "Offline\n(kbd)", "% of\nPADD",
                 "Start", "End", "Unit"]
        widths = [2.6, 2.7, 0.8, 0.85, 0.8, 0.95, 0.95, 2.0]
        rows = min(maxrows, len(df)) + 1
        x, y = Inches(0.45), Inches(1.35)
        tbl = s.shapes.add_table(rows, len(heads), x, y, Inches(12.4), Inches(0.3 * rows)).table
        tbl.first_row = False
        tbl.horz_banding = True
        for j, wd in enumerate(widths):
            tbl.columns[j].width = Inches(wd)
        for j, h in enumerate(heads):
            c = tbl.cell(0, j)
            c.fill.solid(); c.fill.fore_color.rgb = NAVY
            c.margin_left = c.margin_right = Pt(3); c.margin_top = c.margin_bottom = Pt(1)
            tf = c.text_frame; tf.word_wrap = True
            p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT if j < 2 else PP_ALIGN.CENTER
            r = p.add_run(); r.text = h
            r.font.size = Pt(9); r.font.bold = True; r.font.color.rgb = WHITE; r.font.name = FONT
        for i, (_, row) in enumerate(df.head(maxrows).iterrows(), start=1):
            vals = [str(row["operator"]).title()[:30], str(row["plant"])[:30],
                    str(row["padd"]).replace("PADD ", "P"), f"{row['kbd']:.1f}",
                    f"{row['pct_padd']:.1%}", _d(row["start"]), _d(row["end"]),
                    str(row["unit_cat"]).title()[:22]]
            for j, v in enumerate(vals):
                c = tbl.cell(i, j)
                c.fill.solid(); c.fill.fore_color.rgb = WHITE if i % 2 else LT_GRAY
                c.margin_left = c.margin_right = Pt(3); c.margin_top = c.margin_bottom = Pt(0)
                tf = c.text_frame; tf.word_wrap = False
                p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT if j < 2 else PP_ALIGN.CENTER
                r = p.add_run(); r.text = v
                r.font.size = Pt(8.5); r.font.color.rgb = INK; r.font.name = FONT
        for i in range(rows):
            tbl.rows[i].height = Inches(0.42 if i == 0 else 0.3)
        if note:
            self._footnote(s, note)
        return s

    def method_slide(self):
        s = self._section("Method & Caveats", "Read before use")
        notes = [
            "Primary metric is CAP_OFFLINE_ADJUSTED_KBD (offline capacity, kbd, all units). Mogas is "
            "a secondary overlay (capacity x unit yield).",
            "Outage type is binary {PLANNED, UNPLANNED}; UNKNOWN folds into UNPLANNED per desk rule.",
            "2027 is planned-only - Plan+Unplanned / Unplanned comparisons vs 2027 are n/a; only "
            "Planned is comparable. Unplanned-2027 is a modeled scenario.",
            "2020-2021 (COVID / Winter Storm Uri) are excluded from forecast baselines by default.",
            "Back-to-back = consecutive months of outage at one plant/unit; surfaces clustered FCC "
            "turnarounds external trackers miss.",
            "Scenario = baseline(window) x (1+growth) x multiplier + one-off(stress month); the Excel "
            "workbook recomputes it live from dropdown inputs.",
            "All numbers refresh from the source export - re-run the build to update workbook, deck "
            "and dashboard together.",
        ]
        self._bullets(s, Inches(0.5), Inches(1.5), Inches(12.3), notes, size=12.5, spacing=0.66,
                      head=None)

    def build(self):
        a = self.a
        ctx = self.ctx
        sm = ctx["summary"]
        ly = max(y for y in sm.index if y not in engine.PARTIAL_YEARS and sm.loc[y, "Unplanned"] > 0)
        share = ctx["padd_share"]
        sc = ctx["scenario"]
        sp = ctx["scenario_padd"]
        _, yoy = ctx["padd_unpl_yoy"]
        un25 = ctx["unit_share"][2025].sort_values(ascending=False)
        fcc = ctx["fcc_exxon"]
        peak_yr = max((y for y in sm.index if 2014 <= y <= 2025), key=lambda y: sm.loc[y, "Total"])

        self.title_slide()
        self.exec_summary()

        self.grid4_slide(
            "US Overview", "Capacity offline - levels, mix and momentum",
            [a["annual"], a["padd_clustered"], a["padd_donut"], a["yoy_month"]],
            [f"Offline peaked in {peak_yr} (2020-21 outliers excluded from baselines); "
             f"FY{ly} ran ~50/50 planned/unplanned.",
             f"PADD 3 ~{share['PADD 3']:.0%} of unplanned; PADD 5 the most volatile "
             f"({yoy.loc['PADD 5',2025]:+.0%} in 2025 on California events)."])

        self.charts_bullets_slide(
            "Seasonality & Monthly Change",
            "Unplanned offline - range band, recent years and YoY% by month",
            [a["season_band"], a["yoy_month"]],
            ["Unplanned offline clusters in Feb (freeze) and the Sep-Oct turnaround window, "
             "with a clear summer trough.",
             "The grey band is the 2022-25 monthly range; recent years ride near the top of "
             "the band in Q1.",
             "YoY% by month (lower chart) shows where each year diverged - e.g. a heavier Feb "
             "in some years, lighter spring in others.",
             "This monthly shape - not a flat annual number - is what drives the 2027 scenario."],
            foot="Range band = monthly min-max across 2022-2025. Partial years (2026/27) flagged.")

        self.charts_bullets_slide(
            "PADD 3 - Planned & Unplanned Offline",
            "2026 plan + unplanned (bars) vs 2023-25 totals and 2027 plan (lines)",
            [a["padd3"]],
            ["PADD 3 (Gulf Coast) is the single largest source of offline capacity and unplanned "
             "risk in the US.",
             "Gold/orange bars = 2026 booked plan and unplanned; lines = prior-year totals; dashed "
             "green = 2027 plan.",
             f"Carries ~{share['PADD 3']:.0%} of US unplanned offline over the baseline window.",
             "Heavy Q1-Q2 activity, consistent with Gulf Coast spring turnaround timing.",
             "Magnitudes share one axis with the other PADDs - honest scale."])

        self.charts_bullets_slide(
            "PADD 2 & PADD 5 - Planned & Unplanned Offline",
            "The next two watch regions", [a["padd2"], a["padd5"]],
            [f"PADD 2 (Midwest) ~{share['PADD 2']:.0%} of unplanned offline; swings with refinery "
             "turnaround cycles.",
             f"PADD 5 (West Coast) ~{share['PADD 5']:.0%}, but the most volatile - the 2025 "
             "California spike is clear in the red total line.",
             "- Joliet (PADD 2) shows the recurring Q1-Q2 FCC clustering covered next.",
             "- PADD 5 tightness leaves little slack when an unplanned event hits.",
             "2027 plan (dashed green) is the only forward-booked series."])

        # FCC back-to-back
        self.charts_bullets_slide(
            "Back-to-Back FCC Outages - ExxonMobil",
            "Consecutive-month FCC runs external trackers miss", [a["fcc"]],
            ["This granular export captures repeated FCC (cat cracker) outages in adjacent months "
             "at the same Exxon plant.",
             "Baton Rouge ran Jan-Jun 2022; Baytown Jan-Mar 2023; Joliet Feb-May 2025 and Feb-Jun "
             "2026 - classic Q1-Q2 turnaround clustering.",
             "Month-aggregated external data smooths these into one figure and loses the back-to-back "
             "signal.",
             f"{len(fcc)} such Exxon FCC runs (>=3 months) since 2011 - a recurring, model-able "
             "pattern.",
             "FCC offline is the most gasoline-relevant unit loss (0.65 mogas yield)."],
            foot="Back-to-back = consecutive calendar months of FCC outage at one plant within a year; 2020 excluded.")

        # TA schedule tables per PADD
        for padd, label in [("PADD 3", "PADD 3"), ("PADD 2", "PADD 2"), ("PADD 5", "PADD 5")]:
            df = ctx["ta_schedule"][padd]
            if len(df):
                self.ta_table_slide(
                    f"{label} 2026 Planned TAs", df,
                    note="One row per outage/unit. % of PADD = offline kbd / PADD total planned offline. "
                         "Source: Snowflake export.")

        self.grid4_slide(
            "Units & Mogas", "Where the offline capacity sits, and the gasoline read-through",
            [a["units"], a["mogas"], a["scenario"], a["scenario_padd"]],
            [f"Crude trains (~{(un25.get('ATMOS DISTILLATION',0)+un25.get('VACUUM DISTILLATION',0)):.0%}), "
             f"hydrotreating (~{un25.get('HYDROTREATING',0):.0%}) and FCC (~{un25.get('FLUID CAT CRACKING',0):.0%}) "
             "lead; FCC/reforming drive mogas exposure.",
             f"2027 scenario ~{kbd(sc['annual_unplanned'])} kbd unplanned (PADD 3 ~{kbd(sp['PADD 3']['annual'])}, "
             f"PADD 2 ~{kbd(sp['PADD 2']['annual'])}, PADD 5 ~{kbd(sp['PADD 5']['annual'])})."])

        # Sensitivity
        torn_top = ctx["tornado"][0]["driver"].split(" (")[0]
        s = self._section("Sensitivity & Risk", "How the 2027 unplanned forecast flexes")
        self._pic_fit(s, a["heatmap"], Inches(0.4), Inches(1.35), Inches(6.5), Inches(4.2), center="both")
        self._pic_fit(s, a["tornado"], Inches(7.05), Inches(1.55), Inches(5.9), Inches(3.6), center="both")
        self._bullets(s, Inches(0.5), Inches(5.75), Inches(12.3), [
            f"Base case {kbd(sc['annual_unplanned'])} kbd outlined; grid spans -10%..+15% growth x "
            f"0.7..1.5x rate -> ~{kbd(sc['annual_unplanned']*0.7)}-{kbd(sc['annual_unplanned']*1.5*1.15)} kbd.",
            f"The unplanned-rate multiplier is the dominant driver ({torn_top}); baseline-window "
            "choice is second."], size=11, spacing=0.32, head=None)

        self.method_slide()

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
