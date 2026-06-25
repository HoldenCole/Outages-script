#!/usr/bin/env python3
"""
build_slides.py
Sell-side "weekly meeting" deck (python-pptx), built around PER-UNIT capacity
offline. The desk reads refinery outages one unit at a time -- CDU first, then
FCC, hydrocracker, reformer -- never as a single summed "total offline" (adding a
crude unit to an FCC to a coker is meaningless). Every figure is concurrent
capacity offline by month, 2021 onward, kept separate by unit and PADD.

The ExxonMobil 2027 slate is cross-checked, per unit, against ExxonMobil's own
corporate turnaround plan (data/exxon_ta_plan.csv); records with no counterpart
in the plan are flagged. Margin/$ estimates are intentionally excluded.

Brand mark: set BRAND_LOGO to a logo image path to drop in your own logo;
otherwise BRAND_TEXT is rendered in the corner.

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

    # ----------------------------------------------------------------- chart layouts
    def charts_bullets_slide(self, title, sub, imgs, bullets, foot=None):
        """1-2 charts stacked on the left, bullets on the right."""
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

    def wide_chart_slide(self, title, sub, img, bullets, foot=None):
        """One wide chart (left ~2/3) + a narrow takeaways column (right)."""
        s = self._section(title, sub)
        self._pic_fit(s, img, Inches(0.4), Inches(1.3), Inches(8.9), Inches(5.3), center="both")
        self._bullets(s, Inches(9.5), Inches(1.5), Inches(3.6), bullets, size=11, spacing=0.46)
        if foot:
            self._footnote(s, foot)
        return s

    def unit_table_slide(self, title, df, note=None, maxrows=14):
        """Per-unit turnaround schedule: one row per physical unit (no summing)."""
        s = self._section(title, "Planned turnarounds by unit - nameplate capacity offline, kbd")
        heads = ["Refinery", "Operator", "PADD", "Unit", "Class", "Offline\n(kbd)", "Window"]
        widths = [2.5, 2.5, 0.8, 2.7, 1.3, 0.95, 1.65]
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
            win = f"{_d(row['start'])}-{_d(row['end'])}" if pd.notna(row["start"]) else str(row["span"])
            vals = [str(row["plant"]).replace(" Refinery", "")[:26],
                    str(row["operator"]).title()[:26],
                    str(row["padd"]).replace("PADD ", "P"),
                    str(row["unit_name"])[:30],
                    str(row["focus"]) if isinstance(row["focus"], str)
                    else str(row["unit_cat"]).title()[:14],
                    f"{row['kbd']:.0f}", win]
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

    # ----------------------------------------------------------------- slides
    def title_slide(self):
        s = self._slide()
        self._rect(s, 0, 0, SW, SH, NAVY)
        self._rect(s, Inches(8.3), 0, SW - Inches(8.3), SH, RGBColor(0x24, 0x40, 0x70))
        self._dotgrid(s, Inches(8.65), Inches(0.5), 10, 14, Inches(0.45), RGBColor(0x3C, 0x5A, 0x90))
        self._text(s, Inches(0.7), Inches(0.5), Inches(6), Inches(0.4), [(self.asof, 14, False, LT_BLUE)])
        self._rect(s, Inches(0.72), Inches(1.04), Inches(6.6), Pt(1.2), RGBColor(0x4A, 0x66, 0x99))
        if BRAND_LOGO and os.path.exists(BRAND_LOGO):
            s.shapes.add_picture(BRAND_LOGO, Inches(10.5), Inches(0.45), height=Inches(0.5))
        else:
            self._text(s, Inches(9.0), Inches(0.45), Inches(3.8), Inches(0.5),
                       [(BRAND_TEXT, 18, True, WHITE)], align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
        self._text(s, Inches(0.7), Inches(2.9), Inches(7.6), Inches(0.4),
                   [("Products Trading", 20, False, LT_BLUE)])
        self._text(s, Inches(0.68), Inches(3.4), Inches(8.0), Inches(1.7),
                   [("Refinery Capacity Offline", 36, True, WHITE),
                    ("by Unit - 2027 Turnaround Outlook", 36, True, WHITE)], sa=2)
        self._rect(s, Inches(0.72), Inches(5.35), Inches(2.8), Pt(2.5), GOLD)
        self._text(s, Inches(0.72), Inches(5.55), Inches(7.8), Inches(0.9),
                   [("Per-unit capacity offline (kbd): CDU - FCC - hydrocracker - reformer   |   "
                     "2021-2027", 12, False, LT_BLUE),
                    ("2027: ExxonMobil full-year (verified vs their corporate plan); all other "
                     "operators H1-confirmed only", 11, False, RGBColor(0x9D, 0xB6, 0xDB))], sa=4)
        self._text(s, Inches(11.3), Inches(7.0), Inches(1.85), Inches(0.3),
                   [("PROPRIETARY", 9, False, LT_BLUE)], align=PP_ALIGN.RIGHT)

    def principle_slide(self):
        """Frame the whole deck: read capacity offline per unit, not as one total."""
        ver = self.ctx["exxon_verify"]
        nconf = int((ver["events"]["verified"] == True).sum())   # noqa: E712
        nflag = len(ver["flagged"])
        self.charts_bullets_slide(
            "Read Capacity Offline Per Unit - Not as One Big Total",
            "Why a single summed 'offline' figure misleads, and how this deck reports instead",
            [self.a["joliet_decode"]],
            ["A single turnaround takes many units down at once. Summing them is meaningless: "
             "Joliet's one April-2027 event adds to ~715 kbd at a refinery that runs ~250 kbd of crude.",
             "The honest read is per unit: ~250 kbd of crude (CDU) offline, with the FCC, coker and "
             "hydrotreaters shown separately - never added together.",
             "So this deck reports the four units that matter, in priority order: "
             "CDU -> FCC -> hydrocracker -> reformer.",
             "Every figure is concurrent capacity offline by month - a unit down Apr-May is counted "
             "once at its nameplate, never stacked across months or across unit types.",
             f"ExxonMobil's 2027 records are cross-checked unit-by-unit against Exxon's own corporate "
             f"TA plan: {nconf} confirmed, {nflag} flagged as not-in-plan.",
             "Margin / $-at-risk estimates are intentionally excluded - this is physical capacity offline only."],
            foot="Concurrent offline = for each month, each distinct unit's nameplate capacity counted once "
                 "(deduped by plant+unit). History shown from 2021.")

    def overview_slide(self):
        fp = self.ctx["focus_peak"]

        def pk(unit, yr):
            return float(fp.loc[yr, unit]) if yr in fp.index else 0.0
        self.wide_chart_slide(
            "Capacity Offline by Unit & Month, 2021-2027",
            "Concurrent capacity offline (kbd) for each focus unit - the per-unit timeline",
            self.a["focus_heat"],
            [f"Crude (CDU+VDU) dominates: peak concurrent ~{kbd(pk('CDU',2025))} kbd offline in 2025, "
             f"vs FCC ~{kbd(pk('FCC',2025))}, hydrocracker ~{kbd(pk('Hydrocracker',2025))}, "
             f"reformer ~{kbd(pk('Reformer',2025))}.",
             "Each unit carries a clear spring (Mar-May) and autumn (Sep-Oct) turnaround window every year.",
             f"2021 stands out across every unit - Winter Storm Uri (Feb) - {kbd(pk('CDU',2021))} kbd of "
             "crude offline at the peak; it's the floor year of this view.",
             "FCC, hydrocracker and reformer are smaller in kbd but gasoline-octane-critical: their loss "
             "squeezes blending even when crude runs hold.",
             "2026-27 are partial / booked-only; for 2027, only ExxonMobil is full-year - all other "
             "operators are H1-confirmed (so the 2027 H2 cells are an incomplete, non-Exxon floor)."],
            foot="Read each panel left-to-right as a month-by-month timeline; cells are kbd of that unit "
                 "class concurrently offline. Never sum across panels. 2027 H2 (non-Exxon) is not yet confirmed.")

    def unit_deepdive_slide(self, focus, lines_img, padd_img, extra_bullets, foot):
        fp = self.ctx["focus_peak"]
        label = engine.FOCUS_LABEL[focus]

        def pk(yr):
            return float(fp.loc[yr, focus]) if yr in fp.index else 0.0
        head = [
            f"Peak concurrent {focus} offline: ~{kbd(pk(2025))} kbd (2025), ~{kbd(pk(2026))} (2026), "
            f"~{kbd(pk(2027))} booked (2027).",
        ]
        self.charts_bullets_slide(
            f"{label} - Capacity Offline by Month & PADD",
            "Concurrent capacity offline (kbd); booked-only beyond H1-2027",
            [lines_img, padd_img], head + extra_bullets, foot=foot)

    def cdu_slide(self):
        ev = self.ctx["unit_events_2027"]
        cdu = ev[ev["focus"].eq("CDU")].sort_values("kbd", ascending=False)
        tops = ", ".join(f"{str(r.plant).replace(' Refinery','')} ~{kbd(r.kbd)}"
                         for _, r in cdu.head(3).iterrows())
        self.unit_deepdive_slide(
            "CDU", self.a["cdu_lines"], self.a["cdu_padd_27"],
            ["Crude is the #1 unit to watch - when the CDU is down, the whole refinery's throughput is cut.",
             "By PADD, the 2027 crude book concentrates on the Gulf Coast (PADD 3) in autumn and "
             "PADD 2 in spring (the Joliet event).",
             f"Largest booked 2027 crude turnarounds: {tops} kbd.",
             "Spring and autumn are the windows; the summer trough is when crude runs are protected for "
             "driving-season gasoline."],
            foot="Crude = atmospheric (CDU) + vacuum (VDU) distillation. 2027 is planned-only; H2 books "
                 "continue to fill in.")

    def fcc_slide(self):
        ev = self.ctx["unit_events_2027"]
        fcc = ev[ev["focus"].eq("FCC")].sort_values("kbd", ascending=False)
        tops = ", ".join(f"{str(r.plant).replace(' Refinery','')} ~{kbd(r.kbd)}"
                         for _, r in fcc.head(3).iterrows())
        self.unit_deepdive_slide(
            "FCC", self.a["fcc_lines"], self.a["fcc_padd_27"],
            ["FCC (cat cracker) is the most gasoline-relevant unit loss (~0.65 mogas yield) - octane and "
             "blending risk stack when it's down.",
             f"Largest booked 2027 FCC turnarounds: {tops} kbd - concentrated in Q1 (Beaumont, Baytown "
             "Fuels South).",
             "FCC turnarounds recur in adjacent months at the same plants - a back-to-back pattern that "
             "month-aggregated trackers smooth away.",
             "Gulf Coast (PADD 3) again leads the FCC slate; watch the spring overlap with crude work."],
            foot="FCC offline shown concurrent by month; a plant's FCC counted once per month at nameplate.")

    def hdc_ref_slide(self):
        fp = self.ctx["focus_peak"]

        def pk(u, y):
            return float(fp.loc[y, u]) if y in fp.index else 0.0
        s = self._section("Hydrocracker & Reformer - Capacity Offline by Month",
                          "The octane/distillate complex: smaller in kbd, but gasoline-quality critical")
        self._pic_fit(s, self.a["hdc_lines"], Inches(0.4), Inches(1.35), Inches(6.4), Inches(4.7),
                      center="both")
        self._pic_fit(s, self.a["ref_lines"], Inches(6.85), Inches(1.35), Inches(6.4), Inches(4.7),
                      center="both")
        self._bullets(s, Inches(0.5), Inches(6.15), Inches(12.3), [
            f"Hydrocracker peak concurrent offline ~{kbd(pk('Hydrocracker',2025))} kbd (2025) -> "
            f"~{kbd(pk('Hydrocracker',2027))} booked 2027; reformer ~{kbd(pk('Reformer',2025))} -> "
            f"~{kbd(pk('Reformer',2027))}.",
            "Hydrocracker = HYDROCRACKING only (distinct from diesel hydrotreating); reformer makes the "
            "high-octane reformate that backstops gasoline blending.",
            "Both are lower-volume than CDU/FCC but their loss tightens octane and distillate quality - a "
            "read CDU-only trackers miss.",
        ], size=11, spacing=0.34, head=None)
        self._footnote(s, "Concurrent capacity offline (kbd) by month, 2021-2027; 2026-27 partial / booked-only.")
        return s

    def exxon_slide(self):
        ver = self.ctx["exxon_verify"]
        ev = ver["events"]
        foc = ev[ev["focus"].isin(engine.FOCUS_ORDER)]
        conf = foc[foc["verified"] == True]                       # noqa: E712
        flag = foc[foc["verified"] == False]                      # noqa: E712
        bullets = [
            "ExxonMobil is the ONLY operator with a full-year 2027 plan - so it's the one refiner whose "
            "H2 turnarounds we can confirm. Everyone else is H1-only.",
            "Shown per unit, never summed: each bar is one unit's nameplate offline over its window - no "
            "meaningless Exxon 'total'.",
            f"Confirmed against Exxon's corporate plan ({len(conf)} focus-unit events): Baytown & Beaumont "
            "FCC in Q1, Joliet crude+vacuum Apr-May, Baton Rouge PSLA-9 crude in autumn.",
        ]
        for _, r in flag.iterrows():
            bullets.append(f"- Flagged: {str(r.plant).replace(' Refinery','')} {str(r.unit_name)[:18]} "
                           f"(~{kbd(r.kbd)} kbd, {r.span}) - {r.note}.")
        bullets.append("The Joliet 'Crude' in Sep-Oct is the '700 kbd' culprit: a duplicate of the real "
                       "April crude TA. Exxon's only Sep-2027 Joliet event is FT Cogen (a utility).")
        self.wide_chart_slide(
            "ExxonMobil 2027 - Per Unit, Verified vs Corporate Plan",
            "Each focus-unit turnaround as its own bar; red-hatched = no match in Exxon's plan",
            self.a["exxon_gantt"], bullets,
            foot="Cross-checked against data/exxon_ta_plan.csv (vendored from the AMR Turnaround Schedule). "
                 "Match = same refinery + unit class overlapping the same months.")

    def basis_2027_slide(self):
        c2 = self.ctx["confirmed2027"]
        pk = lambda f: max(c2[f]["confirmed"])
        self.wide_chart_slide(
            "2027 - What's Confirmed vs Still Being Booked",
            "Capacity offline by unit & month: solid = confirmed, hatched = non-Exxon H2 (not yet booked)",
            self.a["splits_2027"],
            ["Only ExxonMobil gave a full-year 2027 plan - and we verified it against their own "
             "corporate turnaround schedule. For every other operator, only H1 (Jan-Jun) 2027 is booked.",
             "So in H2, the solid bars are ExxonMobil alone; the hatched H2 is other operators' work "
             "still being scheduled - a floor that fills in, not a confirmed number.",
             f"Confirmed peak concurrent offline (all in H1): CDU ~{kbd(pk('CDU'))} kbd, "
             f"FCC ~{kbd(pk('FCC'))}, hydrocracker ~{kbd(pk('Hydrocracker'))}, reformer ~{kbd(pk('Reformer'))}.",
             "The apparent autumn CDU spike is almost entirely unconfirmed non-Exxon H2 - don't read it "
             "as a booked surge.",
             "Bottom line: compare 2027 H1 like-for-like vs prior years; read H2 as Exxon-confirmed plus "
             "an open book."],
            foot="Confirmed = ExxonMobil (any month, plan-verified) + all operators' H1. The Joliet Sep-Oct "
                 "'Crude' duplicate and the 2027 Joliet FCC (the plan books it 2026/2030) are excluded.")

    def build(self):
        self.title_slide()
        self.principle_slide()
        self.basis_2027_slide()
        self.overview_slide()
        self.cdu_slide()
        self.fcc_slide()
        self.hdc_ref_slide()
        self.exxon_slide()
        # per-unit 2027 turnaround schedule (focus units, largest offline first)
        ev = self.ctx["unit_events_2027"]
        foc = ev[ev["focus"].isin(engine.FOCUS_ORDER)].sort_values("kbd", ascending=False)
        self.unit_table_slide(
            "Largest 2027 Turnarounds by Unit - CDU / FCC / HC / Reformer", foc,
            note="One row per physical unit (deduped to peak nameplate offline), focus units only, all PADDs. "
                 "ExxonMobil rows are full-year plan-verified; non-Exxon H2 rows are not yet confirmed. "
                 "Excludes the verified-bad Joliet Sep-Oct crude & 2027 FCC.")

    def save(self, path):
        self.prs.save(path)


def main():
    ap = argparse.ArgumentParser(description="Refinery per-unit outage slide deck")
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
