#!/usr/bin/env python3
"""
build_slides.py
Trading-desk deck (python-pptx): four things a trader needs, kept simple --
    1. Total 2027 outages by unit  (and what each unit tightens)
    2. Outages by PADD by unit     (where it tightens)
    3. ExxonMobil outages          (per unit, verified vs their corporate plan)
    4. 2027 unplanned scenario     (the risk on top of the booked plan)

Everything is per-unit capacity offline (never a summed "total"), 2027-forward.
2027 completeness is asymmetric: only ExxonMobil gave a full-year plan (verified
against their corporate schedule); every other operator is H1-confirmed only, so
the non-Exxon H2 is flagged as not-yet-booked throughout.

Brand mark: set BRAND_LOGO to a logo image path; otherwise BRAND_TEXT is used.

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

import engine
import charts

_ROOT = Path(__file__).resolve().parent.parent
INPUT_PATH = str(_ROOT / "data" / "Refinery_Outages_Enhanced.xlsx")
OUT_PATH = str(_ROOT / "output" / "outage_deck.pptx")

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
        s = self._slide()
        self.page += 1
        self._text(s, Inches(0.45), Inches(0.18), Inches(12.4), Inches(0.7),
                   [(title, 30, False, RED)], anchor=MSO_ANCHOR.MIDDLE)
        if sub:
            self._text(s, Inches(0.5), Inches(0.92), Inches(12.3), Inches(0.3),
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
            boxw = w - Inches(0.22) - (bx - x)
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
                   [("Refinery Outages", 38, True, WHITE),
                    ("2027 Outlook by Unit", 38, True, WHITE)], sa=2)
        self._rect(s, Inches(0.72), Inches(5.35), Inches(2.8), Pt(2.5), GOLD)
        self._text(s, Inches(0.72), Inches(5.55), Inches(7.9), Inches(0.9),
                   [("Capacity offline by unit, region & operator - and what it tightens", 12.5, False, LT_BLUE),
                    ("2027: ExxonMobil full-year (verified vs their plan); all other operators "
                     "H1-confirmed only", 11, False, RGBColor(0x9D, 0xB6, 0xDB))], sa=4)
        self._text(s, Inches(11.3), Inches(7.0), Inches(1.85), Inches(0.3),
                   [("PROPRIETARY", 9, False, LT_BLUE)], align=PP_ALIGN.RIGHT)

    def total_by_unit_slide(self):
        c2 = self.ctx["confirmed2027"]
        pk = lambda f: max(c2[f]["confirmed"])
        self.wide_chart_slide(
            "Total 2027 Outages by Unit - and What They Tighten",
            "Capacity offline by unit & month; solid = confirmed, hatched = non-Exxon H2 (not yet booked)",
            self.a["splits_2027"],
            ["Crude (CDU) down = the whole refinery's run is cut - every product from that site. "
             "FCC down = gasoline + octane. Reformer = octane. Hydrocracker = diesel / jet.",
             f"Confirmed 2027 peak: CDU ~{kbd(pk('CDU'))}, FCC ~{kbd(pk('FCC'))}, "
             f"hydrocracker ~{kbd(pk('Hydrocracker'))}, reformer ~{kbd(pk('Reformer'))} kbd - all in H1.",
             "Read units separately, never added: a 250-kbd CDU plus a 100-kbd FCC is not '350 offline'.",
             "Solid = confirmed (Exxon full-year + everyone's H1). Hatched = non-Exxon H2, still being "
             "booked - don't trade the autumn spike as real."],
            foot="Concurrent capacity offline, each unit counted once per month. Non-Exxon H2 2027 is a floor "
                 "that fills in. Verified-bad Exxon records excluded.")

    def padd_by_unit_slide(self):
        self.charts_bullets_slide(
            "Outages by PADD by Unit - Where It Tightens",
            "2027 crude (CDU, top) and cat-cracker (FCC, bottom) offline by region & month",
            [self.a["cdu_padd_27"], self.a["fcc_padd_27"]],
            ["PADD 3 (Gulf) is the swing region - most crude and FCC work lands there, tightening USGC "
             "supply and the export barrel.",
             "PADD 2 (Midwest) carries the spring crude (the Joliet event). PADD 5 (West) is islanded - a "
             "California outage doesn't get bailed out by other regions.",
             "Timing: spring (Mar-May) and autumn (Sep-Oct) are the windows; summer is protected for "
             "driving-season gasoline.",
             "Past the dotted line (H2) is non-Exxon-unconfirmed - those autumn bars are a floor and grow "
             "as operators book."],
            foot="Concurrent capacity offline by month, stacked by PADD. P1 NE, P2 Midwest, P3 Gulf, "
                 "P4 Rockies, P5 West.")

    def exxon_slide(self):
        ev = self.ctx["exxon_verify"]["events"]
        conf = ev[ev["focus"].isin(engine.FOCUS_ORDER) & (ev["verified"] == True)]   # noqa: E712
        has_plan = len(conf) > 0          # corporate-plan reconciliation available (optional)
        bullets = [
            "ExxonMobil is the only operator with a full-year 2027 plan - so it's the one refiner whose H2 "
            "we can confirm. Everyone else is H1-only.",
            "Per unit, never summed - the 'Exxon ~700 kbd' figure was 8 Joliet units in one April "
            "turnaround added together; it's really ~250 kbd of crude.",
            (f"Reconciled to Exxon's own corporate plan ({len(conf)} units): " if has_plan
             else "Exxon's verified 2027 slate: ")
            + "Baytown & Beaumont FCC in Q1, Joliet crude Apr-May, Baton Rouge crude in autumn.",
        ]
        self.wide_chart_slide(
            "ExxonMobil 2027 - Per Unit",
            "Each unit's turnaround as its own bar, nameplate capacity offline (kbd)",
            self.a["exxon_gantt"], bullets,
            foot=("Cross-checked against ExxonMobil's corporate turnaround plan; match = same refinery + "
                  "unit class, overlapping months." if has_plan
                  else "Per-unit nameplate offline from the verified outage book."))

    def scenario_slide(self):
        sc = self.ctx["scenario"]
        fan = self.ctx["scenario_fan"]
        cons = float(fan["Conservative"].sum()); avg = float(fan["Average"].sum()); act = float(fan["Active"].sum())
        pl = float(sc["planned_2027"])
        self.charts_bullets_slide(
            "2027 Unplanned Scenario - the Risk on Top of Planned",
            "Potential unplanned offline (kbd) modeled off the 2022-25 seasonal pattern",
            [self.a["fan"], self.a["scenario_total"]],
            ["2027 has no actual unplanned yet - this is the risk range to carry on top of the booked "
             "planned slate.",
             f"Average ~{kbd(avg)} kbd unplanned; Conservative ~{kbd(cons)} (calm year), Active ~{kbd(act)} "
             f"(heavy). Booked planned is ~{kbd(pl)} kbd.",
             f"Implied total ~{kbd(cons + pl)} / ~{kbd(avg + pl)} / ~{kbd(act + pl)} kbd "
             "(Conservative / Average / Active).",
             "Risk peaks in Feb (winter freeze) and Sep-Oct (turnaround overlap); summer is the trough.",
             "Trade it: Active = stress case for supply tightness; Conservative = the floor."],
            foot="Scenario = mean 2022-25 monthly unplanned shape x {0.8 / 1.0 / 1.3}. A risk range, not a "
                 "forecast; fully tunable in the Excel Scenario tab.")

    def build(self):
        self.title_slide()
        self.total_by_unit_slide()     # 4) total outages by unit
        self.padd_by_unit_slide()      # 1) outages by PADD by unit
        self.exxon_slide()             # 2) ExxonMobil outages
        self.scenario_slide()          # 3) 2027 unplanned scenario

    def save(self, path):
        self.prs.save(path)


def main():
    ap = argparse.ArgumentParser(description="Refinery outage trading-desk deck")
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
