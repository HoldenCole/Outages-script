#!/usr/bin/env python3
"""
build_slides_naphtha.py
Parallel trading-desk deck, tilted to NAPHTHA / CHEM FEED and REFORMERS, and
headlining the in-progress year ("rest of <yy>") rather than the forward outlook
year. Same shape as the main deck (build_slides.py) - per-unit capacity offline,
chart-forward, presenter brings their own notes - but the read is octane and
petrochemical feedstock, not gasoline/distillate yields:

    1. Rest-of-year outages by unit   (H1 actual vs H2 booked, all four units)
    2. What's driving it              (biggest individual outages of the year)
    3. Reformers: the octane read     (reformer offline by month & PADD)   <- focus
    4. Naphtha balance                (CDU supply vs reformer demand, this year)
    5. Naphtha / octane / chem-feed   (the reforming-isom-aromatics complex)
    6. Reformer & crude by PADD       (where the octane / naphtha goes down)
    7. Unplanned context              (recent actuals, grounds the rest of year)

Naphtha is the reformer's charge and the steam-cracker / petrochemical feed, so
a reformer or naphtha-complex outage tightens octane and chem feed even when
crude runs hold - a read CDU-only trackers miss. Reuses the main deck's Deck
primitives and the shared engine context + charts, so the two decks agree.

Usage:
    python build_slides_naphtha.py                  # uses INPUT_PATH
    python build_slides_naphtha.py export.xlsx --out outage_deck_naphtha.pptx
"""
import argparse
import sys
import tempfile
from pathlib import Path

from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

import engine
import charts
from build_slides import (Deck, kbd, BRAND_LOGO, BRAND_TEXT,
                           NAVY, RED, GOLD, GREEN, LT_BLUE, GRAY, WHITE, SW, SH)
import os

_ROOT = Path(__file__).resolve().parent.parent
INPUT_PATH = str(_ROOT / "data" / "Golden_Record_Snowflake.xlsx")
OUT_PATH = str(_ROOT / "output" / "outage_deck_naphtha.pptx")

CY = engine.CURRENT_YEAR            # the in-progress year this deck headlines (rest of <yy>)


class NaphthaDeck(Deck):
    """Reuses Deck's primitives/chrome; swaps in the naphtha/chem-feed slide set."""

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
                    (f"Rest of {str(CY)[2:]}: Naphtha, Reformers & Chem Feed", 30, True, WHITE)], sa=2)
        self._rect(s, Inches(0.72), Inches(5.35), Inches(2.8), Pt(2.5), GOLD)
        self._text(s, Inches(0.72), Inches(5.55), Inches(7.9), Inches(0.95),
                   [("The octane and petrochemical-feedstock read: reformers and the naphtha complex, "
                     "not just crude", 12.5, False, LT_BLUE),
                    (f"{CY} H1 is actual; H2 is the booked plan still to come", 11, False,
                     RGBColor(0x9D, 0xB6, 0xDB))], sa=4)
        self._text(s, Inches(11.3), Inches(7.0), Inches(1.85), Inches(0.3),
                   [("PROPRIETARY", 9, False, LT_BLUE)], align=PP_ALIGN.RIGHT)

    def rest_of_year_slide(self):
        fm = self.ctx["focus_monthly"]
        h2 = {f: sum(float(self.ctx["focus_planned"][f].loc[CY, m])
                     for m in engine.MONTHS[6:]) if CY in self.ctx["focus_planned"][f].index else 0.0
              for f in engine.FOCUS_ORDER}
        ref_h2 = kbd(h2.get("Reformer", 0.0))
        self.wide_chart_slide(
            f"Rest of {CY} Outages by Unit",
            "Capacity offline by unit & month; solid = H1 actual, hatched = H2 booked plan (still to come)",
            self.a["cy_splits"],
            foot="Day-weighted offline (a unit down part of a month counts only its days down), each unit "
                 f"once per month. H1 reported; H2 is the booked plan. Reformer H2 booked ~{ref_h2} kbd.")

    def drivers_slide(self):
        ev = engine.unit_events(self.ctx["df"], year=CY)
        ev = ev[ev["focus"].isin(engine.FOCUS_ORDER)].sort_values("kbd", ascending=False)
        refs = ev[ev["focus"].eq("Reformer")].head(3)
        rnames = ", ".join(f"{r['plant'].replace(' Refinery', '')} ({kbd(r['kbd'])})" for _, r in refs.iterrows()) \
            if len(refs) else "n/a"
        self.wide_chart_slide(
            f"What's Driving It: the Biggest {CY} Outages",
            f"Each bar is one unit's nameplate offline (kbd) in {CY}, colored by PADD (region)",
            self.a["biggest_cy"],
            foot=f"Per-unit nameplate offline, the 12 biggest focus-unit outages of {CY}. Color = PADD "
                 f"region. Biggest reformers: {rnames} kbd.")

    def reformer_slide(self):
        nb = self.ctx["naphtha_balance_cy"]
        ref_off = sum(nb["ref_offline"])
        reformate = sum(nb["reformate_lost"])
        self.charts_bullets_slide(
            "Reformers: the Octane Read",
            f"{CY} catalytic-reformer capacity offline by month & PADD, with reformate (octane) lost",
            [self.a["reformer_focus"]],
            foot=f"Reformers run naphtha -> reformate (the octane in gasoline). ~{kbd(ref_off)} kbd of "
                 f"reformer offline over {CY} = ~{kbd(reformate)} kbd of reformate (octane) not made. "
                 f"Dashed line is {CY-1} for context; past the dotted line is H2 (booked).")

    def naphtha_balance_slide(self):
        nb = self.ctx["naphtha_balance_cy"]
        net = nb["net"]
        order = sorted(range(12), key=lambda i: net[i])
        m1, m2 = engine.MONTHS[order[0]], engine.MONTHS[order[1]]
        ny = int(round(nb["naphtha_yield"] * 100))
        state = "deficit (short)" if nb["annual_net"] < 0 else "surplus (long)"
        self.wide_chart_slide(
            "Naphtha Balance: CDU Supply vs Reformer Demand",
            f"{CY} outages read as naphtha length. CDU makes naphtha; reformers consume it",
            self.a["naphtha_balance_cy"],
            foot=f"Net = reformer offline x {nb['reformer_intake']:.0f} (demand) minus CDU offline x "
                 f"{nb['naphtha_yield']:.2f} (supply, ~{ny}% of crude), day-weighted. {CY} runs a {state} "
                 f"(net {kbd(nb['annual_net'])} kbd), tightest {m1}/{m2}; + surplus / - deficit.")

    def chemfeed_slide(self):
        na = self.ctx["naphtha"]["annual"]
        cy_tot = float(na.loc[CY, "Total"]) if CY in na.index else 0.0
        py = CY - 1
        chg = (cy_tot / float(na.loc[py, "Total"]) - 1.0) if (py in na.index and na.loc[py, "Total"]) else 0.0
        self.wide_chart_slide(
            "Naphtha / Octane / Chem-Feed Complex",
            "Reforming, isomerization and aromatics/BTX offline by year - the octane & petrochem-feed read",
            self.a["naphtha_complex"],
            foot=f"The naphtha/octane complex (reformer charge + steam-cracker / petrochem feed). {CY} "
                 f"offline ~{kbd(cy_tot)} kbd ({chg:+.0%} vs {py}). Reforming dominates the complex; this "
                 "is the octane/chem-feed availability a CDU-only tracker misses.")

    def padd_slide(self):
        self.charts_bullets_slide(
            "Reformer & Crude Outages by PADD",
            f"{CY} reformer (octane, left) and crude (CDU, right) offline by region & month",
            [self.a["ref_padd_cy"], self.a["cdu_padd_cy"]],
            foot="Day-weighted concurrent capacity offline by month, stacked by PADD. P1 NE, P2 Midwest, "
                 "P3 Gulf, P4 Rockies, P5 West. Reformer octane and crude naphtha go down together in P3 (Gulf).")

    def unplanned_context_slide(self):
        self.wide_chart_slide(
            f"Unplanned Offline: {CY-2}-{CY} Context",
            f"What unplanned actually looked like recently, to ground the rest of {CY}",
            self.a["unplanned_context"],
            foot=f"Actual unplanned capacity offline by month, day-weighted. {CY} reported through June; "
                 "the Feb-freeze and autumn-turnaround windows are the recurring risk to carry into H2.")

    def build(self):
        self.title_slide()
        self.rest_of_year_slide()         # all four units, H1 actual vs H2 booked
        self.drivers_slide()              # biggest individual outages of the year
        self.reformer_slide()             # reformers: the octane read   (focus)
        self.naphtha_balance_slide()      # CDU supply vs reformer demand
        self.chemfeed_slide()             # naphtha/octane/chem-feed complex
        self.padd_slide()                 # reformer & crude by PADD
        self.unplanned_context_slide()    # recent unplanned context


def main():
    ap = argparse.ArgumentParser(description="Naphtha / chem-feed rest-of-year deck")
    ap.add_argument("excel", nargs="?", default=INPUT_PATH, help="path to the outage .xlsx export")
    ap.add_argument("--out", default=OUT_PATH, help="output .pptx path")
    args = ap.parse_args()

    print(f"Loading {args.excel.strip()} ...")
    ctx = engine.build_context(args.excel)
    with tempfile.TemporaryDirectory() as tmp:
        print("Rendering naphtha/chem-feed charts ...")
        assets = charts.render_naphtha_assets(ctx, tmp)
        print(f"Building naphtha deck -> {args.out}")
        Path(args.out).parent.mkdir(parents=True, exist_ok=True)
        deck = NaphthaDeck(ctx, assets)
        deck.build()
        deck.save(args.out)
    print("Done.")


if __name__ == "__main__":
    sys.exit(main())
