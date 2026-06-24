#!/usr/bin/env python3
"""
build_all.py
One-shot builder: loads the outage export once and produces all three
deliverables (Excel workbook, PowerPoint deck, HTML dashboard) from the same
data context, so they always agree.

Usage:
    python build_all.py                          # uses the default INPUT_PATH
    python build_all.py path/to/export.xlsx      # point at a refreshed export
    python build_all.py export.xlsx --outdir dist
"""
import argparse
import json
import os
import sys
import tempfile

import engine
import charts
import build_workbook
import build_slides
import build_dashboard

DEFAULT_INPUT = "rEFINERY oUTAGES.xlsx"


def main():
    ap = argparse.ArgumentParser(description="Build workbook + deck + dashboard")
    ap.add_argument("excel", nargs="?", default=DEFAULT_INPUT, help="path to the outage .xlsx export")
    ap.add_argument("--outdir", default=".", help="output directory")
    args = ap.parse_args()
    os.makedirs(args.outdir, exist_ok=True)
    out = lambda n: os.path.join(args.outdir, n)

    print(f"[1/4] Loading {args.excel.strip()} ...")
    ctx = engine.build_context(args.excel)
    d = ctx["diag"]
    print(f"      {d['rows']:,} rows | {d['events_distinct']:,} distinct outages | "
          f"years {d['years'][0]}-{d['years'][1]}")

    print("[2/4] Excel workbook ...")
    wb_path = out("outage_workbook.xlsx")
    build_workbook.Build(ctx, wb_path).run()
    print(f"      -> {wb_path}")

    print("[3/4] Slide deck ...")
    deck_path = out("outage_deck.pptx")
    with tempfile.TemporaryDirectory() as tmp:
        assets = charts.render_all(ctx, tmp)
        deck = build_slides.Deck(ctx, assets)
        deck.build()
        deck.save(deck_path)
    print(f"      -> {deck_path}")

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

    print("\nDone. Open the workbook in real Excel for final visual QA (charts, "
          "heatmap, dropdowns); the deck and dashboard are self-contained.")


if __name__ == "__main__":
    sys.exit(main())
