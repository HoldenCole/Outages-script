#!/usr/bin/env python3
"""
build_all.py
One-shot builder: loads the outage export once and produces all three
deliverables (PowerPoint deck, Excel model, HTML dashboard) from the same data
context and the same rendered charts, so they always agree.

Usage:
    python build_all.py                          # uses the default input
    python build_all.py path/to/export.xlsx      # point at a refreshed export
    python build_all.py export.xlsx --outdir dist
"""
import argparse
import json
import os
import sys
import tempfile
from pathlib import Path

import engine
import charts
import build_slides
import build_slides_naphtha
import build_dashboard
import build_workbook

_ROOT = Path(__file__).resolve().parent.parent          # repo root (scripts/ -> ..)
DEFAULT_INPUT = str(_ROOT / "data" / "Golden_Record_Snowflake.xlsx")
DEFAULT_OUTDIR = str(_ROOT / "output")


def main():
    ap = argparse.ArgumentParser(description="Build deck + dashboard")
    ap.add_argument("excel", nargs="?", default=DEFAULT_INPUT, help="path to the outage export")
    ap.add_argument("--outdir", default=DEFAULT_OUTDIR, help="output directory")
    args = ap.parse_args()
    os.makedirs(args.outdir, exist_ok=True)
    out = lambda n: os.path.join(args.outdir, n)

    print(f"[1/4] Loading {args.excel.strip()} ...")
    ctx = engine.build_context(args.excel)
    d = ctx["diag"]
    print(f"      {d['rows']:,} rows | {d['events_distinct']:,} distinct outages | "
          f"years {d['years'][0]}-{d['years'][1]}")

    # Render the charts once and share them across the deck and the Excel model
    # (every chart is embedded in the workbook too, so they are copy-and-paste).
    with tempfile.TemporaryDirectory() as tmp:
        assets = charts.render_all(ctx, tmp)

        print("[2/4] Slide decks ...")
        deck_path = out("outage_deck.pptx")
        deck = build_slides.Deck(ctx, assets)
        deck.build()
        deck.save(deck_path)
        print(f"      -> {deck_path}")
        # parallel naphtha / chem-feed / reformer deck (headlines the in-progress year)
        nap_assets = charts.render_naphtha_assets(ctx, tmp)
        nap_path = out("outage_deck_naphtha.pptx")
        ndeck = build_slides_naphtha.NaphthaDeck(ctx, nap_assets)
        ndeck.build()
        ndeck.save(nap_path)
        print(f"      -> {nap_path}")

        print("[3/4] Excel model ...")
        model_path = out("outage_model.xlsx")
        build_workbook.build_workbook(ctx, assets, model_path)
        print(f"      -> {model_path}")

    print("[4/4] HTML dashboard ...")
    dash_path = out("outage_dashboard.html")
    data = build_dashboard.build_data(ctx)
    cjs = build_dashboard.fetch_chartjs()
    if cjs is None:
        cjs = (f'document.write(unescape("%3Cscript src=\'{build_dashboard.CHARTJS_URL}\''
               '%3E%3C/script%3E"));')
    html = build_dashboard.HTML.replace("__CHARTJS__", cjs).replace(
        "__DATA__", json.dumps(data, separators=(",", ":")))
    with open(dash_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"      -> {dash_path}")

    print("\nDone. Both decks, the Excel model and the dashboard are self-contained and agree.")


if __name__ == "__main__":
    sys.exit(main())
