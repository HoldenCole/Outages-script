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
import math
import os
import sys
import tempfile
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

import pandas as pd

import engine
import charts

_ROOT = Path(__file__).resolve().parent.parent
INPUT_PATH = str(_ROOT / "data" / "Refinery_Outages_Data.xlsx")
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
            fs = size if not sub else size - 1.5
            boxw = w - Inches(0.22) - (bx - x)                  # text width (EMU)
            # estimate wrapped line count so multi-line bullets don't overlap
            cpl = max(18, int((boxw / 914400) * 72 / (fs * 0.53)))
            nlines = max(1, math.ceil(len(txt) / cpl))
            self._rect(s, bx, y + Inches(0.07), Inches(0.09), Inches(0.09),
                       GOLD if not sub else GRAY)
            self._text(s, bx + Inches(0.22), y, boxw, Inches(0.26 * nlines + 0.1),
                       [(txt, fs, False, INK)])
            base = spacing if not sub else spacing - 0.06
            y += Inches(max(base, nlines * (fs * 1.2 / 72) + 0.16))
        return y

    def _dotgrid(self, s, x0, y0, cols, rows, step, color, d=Pt(5)):
        for r in range(rows):
            for c in range(cols):
                shp = s.shapes.add_shape(MSO_SHAPE.OVAL, int(x0 + c * step), int(y0 + r * step), d, d)
                shp.fill.solid(); shp.fill.fore_color.rgb = color
                shp.line.fill.background(); shp.shadow.inherit = False

    # ----------------------------------------------------------------- slides
    def title_slide(self):
        s = self._slide()
        self._rect(s, 0, 0, SW, SH, NAVY)
        # textured right panel + halftone dots (evokes the reference cover)
        self._rect(s, Inches(8.3), 0, SW - Inches(8.3), SH, RGBColor(0x24, 0x40, 0x70))
        self._dotgrid(s, Inches(8.65), Inches(0.5), 10, 14, Inches(0.45), RGBColor(0x3C, 0x5A, 0x90))
        # date + thin rule (top-left)
        self._text(s, Inches(0.7), Inches(0.5), Inches(6), Inches(0.4), [(self.asof, 14, False, LT_BLUE)])
        self._rect(s, Inches(0.72), Inches(1.04), Inches(6.6), Pt(1.2), RGBColor(0x4A, 0x66, 0x99))
        # brand top-right
        if BRAND_LOGO and os.path.exists(BRAND_LOGO):
            s.shapes.add_picture(BRAND_LOGO, Inches(10.5), Inches(0.45), height=Inches(0.5))
        else:
            self._text(s, Inches(9.0), Inches(0.45), Inches(3.8), Inches(0.5),
                       [(BRAND_TEXT, 18, True, WHITE)], align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
        # title block (lower-left, two lines like the reference)
        self._text(s, Inches(0.7), Inches(3.0), Inches(7.6), Inches(0.4),
                   [("Products Trading", 20, False, LT_BLUE)])
        self._text(s, Inches(0.68), Inches(3.5), Inches(8.0), Inches(1.7),
                   [("Refinery Outage &", 38, True, WHITE),
                    ("2027 Turnaround Outlook", 38, True, WHITE)], sa=2)
        self._rect(s, Inches(0.72), Inches(5.45), Inches(2.8), Pt(2.5), GOLD)
        d = self.ctx["diag"]
        self._text(s, Inches(0.72), Inches(5.65), Inches(7.4), Inches(0.4),
                   [(f"Capacity offline (kbd) - planned & unplanned - 2027 scenario   |   "
                     f"{d['rows']:,} rows, {d['years'][0]}-{d['years'][1]}", 11.5, False, LT_BLUE)])
        self._text(s, Inches(11.3), Inches(7.0), Inches(1.85), Inches(0.3),
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

    def wide_chart_slide(self, title, sub, img, bullets, foot=None):
        """One wide chart (left ~2/3) + a narrow takeaways column (right)."""
        s = self._section(title, sub)
        self._pic_fit(s, img, Inches(0.4), Inches(1.3), Inches(8.7), Inches(5.3), center="both")
        self._bullets(s, Inches(9.35), Inches(1.5), Inches(3.75), bullets, size=11, spacing=0.46)
        if foot:
            self._footnote(s, foot)
        return s

    def padd_overview_slide(self):
        """Replaces five sparse per-PADD slides with one dense small-multiples view."""
        pp = self.ctx["padd_planned"]
        bul = ["2027 planned offline vs 2026, by region (kbd):"]
        for p in engine.PADD_ORDER:
            v26 = float(pp.loc[p, 2026]); v27 = float(pp.loc[p, 2027])
            pct = (v27 / v26 - 1) if v26 else None
            bul.append(f"- {p}: ~{kbd(v27)} vs ~{kbd(v26)}"
                       + (f"  ({pct:+.0%})." if pct is not None else "."))
        nat26 = float(pp[2026].sum()); nat27 = float(pp[2027].sum())
        bul.append(f"National ~{kbd(nat27)} kbd in 2027 vs ~{kbd(nat26)} in 2026 "
                   f"({nat27/nat26-1:+.0%}) - Gulf Coast (PADD 3) leads the build.")
        bul.append("Each panel: 2026 plan+unplanned (bars), recent-year totals (lines), 2027 plan (green dashed).")
        self.wide_chart_slide(
            "Planned 2027 by Region - All PADDs",
            "Per-PADD monthly offline: 2026 actuals, recent-year totals and the 2027 booked plan",
            self.a["padd_all"], bul,
            foot="2027 is planned-only and complete through H1; H2 books fill in. 2020-21 excluded.")

    def national_ta_slide(self):
        """Dense national 2027 turnaround schedule (largest by offline capacity)."""
        allta = pd.concat([self.ctx["ta_2027"][p] for p in engine.PADD_ORDER], ignore_index=True)
        allta = allta.sort_values("kbd", ascending=False)
        self.ta_table_slide(
            "Largest 2027 Turnarounds Nationally", allta,
            note="Top planned 2027 turnarounds by offline capacity (kbd), all PADDs. 2027 is "
                 "planned-only and complete through H1; H2 books continue to fill in.", maxrows=14)

    def build(self):
        a = self.a
        ctx = self.ctx
        sm = ctx["summary"]
        pp = ctx["padd_planned"]
        sc = ctx["scenario"]
        sb = ctx["scenario_bands"]
        fcc = ctx["fcc_exxon"]
        ex = ctx["exxon_2027"]
        pl27 = float(sm.loc[2027, "Planned"]); pl26 = float(sm.loc[2026, "Planned"]); pl25 = float(sm.loc[2025, "Planned"])

        def ppct(p):
            v26 = float(pp.loc[p, 2026]); v27 = float(pp.loc[p, 2027])
            return (v27 / v26 - 1) if v26 else 0.0

        self.title_slide()

        # 1) 2027 vs 2026 & 2025 planned  (lead with the H1 like-for-like)
        h1 = ctx["h1_planned"]; h1_26 = float(h1.get(2026, 0)); h1_27 = float(h1.get(2027, 0))
        self.charts_bullets_slide(
            "2027 vs 2026 & 2025 - Planned Offline",
            "US planned capacity offline (kbd); 2027 book complete through H1 only",
            [a["planned_xyear"], a["padd_pair"]],
            [f"H1 is the clean read: 2027 H1 planned ~{kbd(h1_27)} kbd vs 2026 H1 ~{kbd(h1_26)} "
             f"({h1_27/h1_26-1:+.0%}) - a real, like-for-like step up.",
             f"Full-year looks larger ({pl27/pl26-1:+.0%} vs 2026) only because 2026 H2 is light/incomplete; "
             "don't over-read it - 2027 H2 is still being scheduled.",
             f"Gulf Coast leads the build: PADD 3 planned {ppct('PADD 3'):+.0%} y/y; "
             f"PADD 2 is the only region lighter ({ppct('PADD 2'):+.0%}).",
             "Work concentrates in spring (Mar-May) and autumn (Sep-Oct) - the usual turnaround windows.",
             "Only planned is comparable vs 2027 (no actual unplanned-2027 exists); the unplanned outlook is next."],
            foot="2027 planned data is incomplete past H1 - compare H1-vs-H1; the full-year bar fills in as operators book H2.")

        # 2) 2027 potential unplanned - Conservative / Average / Active fan
        fan = ctx["scenario_fan"]
        cons = float(fan["Conservative"].sum()); avg = float(fan["Average"].sum()); act = float(fan["Active"].sum())
        self.charts_bullets_slide(
            "2027 Potential Unplanned - Scenario Fan",
            "Conservative / Average / Active paths off the 2022-25 seasonal baseline",
            [a["fan"], a["tornado"]],
            [f"Average case ~{kbd(avg)} kbd unplanned; Conservative ~{kbd(cons)} (calm year, x0.8), "
             f"Active ~{kbd(act)} (heavy year, x1.3).",
             f"On top of {kbd(sc['planned_2027'])} kbd booked planned, that's an implied total of "
             f"~{kbd(cons+sc['planned_2027'])} / ~{kbd(avg+sc['planned_2027'])} / ~{kbd(act+sc['planned_2027'])} kbd.",
             f"Active vs Conservative spans ~{kbd(act-cons)} kbd ({act/cons-1:+.0%}) - the unplanned risk range to hedge.",
             "Risk peaks in Feb (winter freeze) and Sep-Oct (turnaround overlap), with a summer trough.",
             "Fully tunable in the Excel Scenario Analysis sheet - window, growth, multiplier, one-off and stress month."])

        # 2b) Outages in dollars - gross margin at risk (valued at the gasoline crack)
        di = ctx.get("dollar_impact") or {}
        if di:
            d25 = di.get(2025) or di[max(di)]
            self.charts_bullets_slide(
                "Outages in Dollars - Gross Margin at Risk",
                "Offline capacity valued at the gasoline crack spread (EIA NYH gasoline - WTI)",
                [a["dollar"], a["dollar_month"]],
                [f"Valuing offline capacity at the gasoline crack turns outages into P&L: 2025 unplanned "
                 f"outages = ~${d25['unplanned']/1000:.1f}bn gross margin at risk (crack ~${d25['crack_avg']:.0f}/bbl).",
                 f"Planned turnarounds ~${d25['planned']/1000:.1f}bn, timed into softer-margin shoulder seasons.",
                 f"Crack has firmed to ~${di.get(2026,{}).get('crack_avg',0):.0f}/bbl in 2026 (from "
                 f"~${d25['crack_avg']:.0f}), so the same offline volume is worth materially more this year.",
                 "2020-21 spikes ($35bn+) reflect the COVID / Winter-Storm-Uri extremes, not normal years.",
                 "Crack is a Bloomberg-overwritable input in the Excel Margin Context sheet (swap EIA seed for RBOB 321)."],
                foot="$MM = offline kbd x crack ($/bbl) x days / 1000. A gross-margin proxy; unplanned = unexpected "
                     "supply loss, planned = margin deferred via scheduled work.")

        # 3) back-to-back-to-back FCC
        self.charts_bullets_slide(
            "Back-to-Back-to-Back FCC Outages - ExxonMobil",
            "Consecutive-month FCC (cat cracker) runs external trackers miss",
            [a["fcc"], a["fcc_year"]],
            ["This granular book captures repeated FCC outages in adjacent months at the same Exxon "
             "plant - genuine back-to-back-to-back runs.",
             "Baton Rouge ran Jan-Jun 2022 (6 months); Joliet Feb-Jun 2026 (5); Baytown Jan-Mar 2023; "
             "Joliet Feb-May 2025.",
             "Month-aggregated external trackers smooth these into one figure and lose the consecutive-run signal.",
             f"{len(fcc)} such Exxon FCC runs (>=3 consecutive months) since 2011 - a recurring, model-able pattern.",
             "FCC is the most gasoline-relevant unit loss (0.65 mogas yield) - octane/blending risk stacks."],
            foot="Runs outlined are >=3 consecutive calendar months of FCC outage at one plant; 2020 excluded.")

        # 4) Exxon 2027 breakdown
        by_ref = ex["by_ref"]; by_unit = ex["by_unit"]; tot = ex["total"] or 1.0
        jol = float(by_ref.get("Joliet Refinery", 0))
        crude = float(by_unit.get("ATMOS DISTILLATION", 0) + by_unit.get("VACUUM DISTILLATION", 0))
        others = [(str(r).replace(" Refinery", ""), float(v)) for r, v in by_ref.items()
                  if r != "Joliet Refinery"][:3]
        others_txt = ", ".join(f"{r} ~{kbd(v)}" for r, v in others)
        self.charts_bullets_slide(
            "ExxonMobil - 2027 Outages Breakdown",
            "2027 is planned-only; this is Exxon's booked turnaround slate",
            [a["exxon27"], a["exxon_unit"]],
            [f"ExxonMobil has ~{kbd(tot)} kbd of planned offline booked for 2027 across "
             f"{len(by_ref)} refineries.",
             f"Joliet (PADD 2) dominates at ~{kbd(jol)} kbd (~{jol/tot:.0%}) - a major spring "
             "(Apr-May) turnaround including the FCC.",
             f"{others_txt} kbd round out the slate.",
             f"By unit: crude trains ~{kbd(crude)} (atmos+vac), FCC ~{kbd(by_unit.get('FLUID CAT CRACKING',0))}, "
             f"hydrotreating ~{kbd(by_unit.get('HYDROTREATING',0))} kbd.",
             "Calendar: heaviest in April with a second wave in Sep-Oct - watch the Joliet spring event."])

        # 5) regional outlook: one dense small-multiples overview + national TA table
        #    (replaces five sparse per-PADD slides)
        self.padd_overview_slide()
        self.national_ta_slide()

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
