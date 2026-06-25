#!/usr/bin/env python3
"""
audit_exxon.py
Reconcile the IIR ExxonMobil outage breakdown against ExxonMobil's OWN corporate
turnaround plan (data/exxon_ta_plan.csv, vendored from the AMR schedule), unit by
unit, for the years the plan covers (2026-2027). The plan is treated as ground
truth; this flags what's wrong in the IIR breakdown:

  * MATCH                - IIR unit + dates agree with the plan
  * DATE OFF ~Nd         - right unit/window but the start drifts > ~18 days
  * WRONG WINDOW         - plan has this unit class at this site, but other months
                           (e.g. the Joliet 'Crude' dated Sep-Oct -> plan = Apr-May)
  * NOT IN PLAN          - no such unit-class turnaround at this site that year
                           (e.g. a Joliet FCC dated 2027 -> plan runs it in 2026)
  * MISSING from IIR (US)- a US turnaround the plan books but IIR omits
  * MISSING (Canada)     - Imperial/Exxon Canadian sites; IIR is US-only (PADD I-V)

Writes, per year: output/exxon_<year>_reconciliation.csv and a color-coded .png.

Usage:
    python scripts/audit_exxon.py                 # default data file, 2026 & 2027
    python scripts/audit_exxon.py path/to/export.xlsx
"""
import sys
from pathlib import Path

import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

import engine

_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_INPUT = str(_ROOT / "data" / "Refinery_Outages_Enhanced.xlsx")
OUTDIR = _ROOT / "output"
YEARS = (2026, 2027)

# IIR unit_category -> the same bucket scheme the vendored plan uses
IIR_BUCKET = {
    "ATMOS DISTILLATION": "CDU", "VACUUM DISTILLATION": "CDU",
    "FLUID CAT CRACKING": "FCC",
    "HYDROCRACKING": "Hydrocracker", "RESID_HYDROCRACKING": "Hydrocracker",
    "REFORMING": "Reformer",
    "COKING": "Coker", "THERM CRACKING, VISBREAKING": "Coker",
    "HYDROTREATING": "Hydrotreater",
}
CORE = ("CDU", "FCC", "Hydrocracker", "Reformer")
CLASS_ORDER = {"CDU": 0, "FCC": 1, "Hydrocracker": 2, "Reformer": 3,
               "Coker": 4, "Hydrotreater": 5, "Other": 6}
CANADA = {"NANTICOKE", "STRATHCONA", "SARNIA"}              # Imperial/Exxon, non-US

ROW_COLOR = {
    "MATCH": "#E2EFDA", "DATE": "#FFF2CC", "WRONG": "#FCE4D6", "NOT": "#F8CBAD",
    "MISSING from IIR (US)": "#DDEBF7", "MISSING (Canada": "#EEEEEE",
}


def _bucket(uc):
    return IIR_BUCKET.get(str(uc), "Other")


def _fmt(s, e):
    try:
        return f"{s.strftime('%-m/%-d')}-{e.strftime('%-m/%-d')}"
    except Exception:
        return "?"


def reconcile(df, plan, year):
    """Return a DataFrame: one row per IIR Exxon unit-event for `year` with a
    verdict vs the plan, followed by core turnarounds the plan books that IIR
    is missing."""
    ev = engine.unit_events(df, operator_contains="EXXON", year=year)
    ev["bk"] = ev["unit_cat"].map(_bucket)
    ev["sk"] = ev["plant"].map(engine._site_key)

    py = plan[plan["year"] == year].copy()
    py["sk"] = py["site"].map(engine._site_key)
    cover = {}                                   # (site, bucket) -> [(start, end, months, unit)]
    for _, r in py.iterrows():
        cover.setdefault((r["sk"], r["bucket"]), []).append(
            (r["start"], r["end"], engine._months_between(r["start"], r["end"]), r["unit"]))

    rows = []
    for _, r in ev.iterrows():
        cand = cover.get((r["sk"], r["bk"]), [])
        best, overlap = None, 0
        for c in cand:
            ov = len(c[2] & set(r["months"]))
            if ov > overlap:
                overlap, best = ov, c
        if best and overlap:
            dd = abs((r["start"] - best[0]).days) if pd.notna(r["start"]) and pd.notna(best[0]) else 0
            verdict = "MATCH" if dd <= 18 else f"DATE OFF ~{dd}d"
            planw = _fmt(best[0], best[1])
        elif cand:
            allm = sorted(set().union(*[c[2] for c in cand]))
            verdict, planw = "WRONG WINDOW", f"plan months {allm}"
        else:
            verdict = "NOT IN PLAN" if r["bk"] in CORE + ("Coker",) else "non-core unit"
            planw = "-"
        rows.append([r["plant"].replace(" Refinery", "").replace(" Complex", ""),
                     str(r["unit_name"])[:22], r["bk"], f"{round(r['kbd'])}",
                     _fmt(r["start"], r["end"]), planw, verdict])

    iir = pd.DataFrame(rows, columns=["Refinery", "Unit (IIR breakdown)", "Class", "kbd",
                                      "IIR window", "Plan window", "Verdict vs Exxon plan"])
    iir = iir.sort_values(["Refinery", "Class"],
                          key=lambda c: c.map(CLASS_ORDER) if c.name == "Class" else c)

    miss = []
    for (sk, bk), lst in cover.items():
        if bk in CORE and not ((ev["sk"] == sk) & (ev["bk"] == bk)).any():
            v = "MISSING (Canada/non-US)" if sk in CANADA else "MISSING from IIR (US)"
            miss.append([sk.title(), "+".join(sorted({x[3] for x in lst}))[:22], bk, "-", "-",
                         _fmt(min(x[0] for x in lst), max(x[1] for x in lst)), v])
    return pd.concat([iir, pd.DataFrame(miss, columns=iir.columns)], ignore_index=True)


def _row_color(verdict):
    for k, c in ROW_COLOR.items():
        if verdict.startswith(k):
            return c
    return "#FFFFFF"


def render(tbl, year, path):
    fig, ax = plt.subplots(figsize=(13, 0.34 * len(tbl) + 1.1))
    ax.axis("off")
    t = ax.table(cellText=tbl.values, colLabels=list(tbl.columns), cellLoc="left", loc="center")
    t.auto_set_font_size(False); t.set_fontsize(8.4); t.scale(1, 1.32)
    for j in range(len(tbl.columns)):
        c = t[0, j]; c.set_facecolor("#1F3864"); c.set_text_props(color="white", fontweight="bold")
    for i in range(len(tbl)):
        col = _row_color(tbl.iloc[i]["Verdict vs Exxon plan"])
        for j in range(len(tbl.columns)):
            cell = t[i + 1, j]; cell.set_facecolor(col)
            if j == 6:
                cell.set_text_props(fontweight="bold")
    for j, wd in enumerate([0.11, 0.22, 0.11, 0.05, 0.13, 0.18, 0.20]):
        for i in range(len(tbl) + 1):
            t[i, j].set_width(wd)
    fig.suptitle(f"ExxonMobil {year} outage breakdown (IIR)  vs  Exxon corporate turnaround plan",
                 fontsize=12.5, fontweight="bold", color="#1F3864", y=0.99)
    fig.text(0.5, 0.005, "green=matches plan  amber=date drift  orange=wrong window  "
             "red=not in plan  blue=US unit missing from IIR  gray=Canadian site (outside US data)",
             ha="center", fontsize=7.6, color="#555")
    fig.savefig(path, dpi=170, bbox_inches="tight"); plt.close(fig)


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_INPUT
    df = engine.load(src)
    plan = engine.load_exxon_plan()
    if plan.empty:
        print("No corporate plan at data/exxon_ta_plan.csv - nothing to reconcile.")
        return 1
    OUTDIR.mkdir(parents=True, exist_ok=True)
    for year in YEARS:
        tbl = reconcile(df, plan, year)
        csv_p = OUTDIR / f"exxon_{year}_reconciliation.csv"
        png_p = OUTDIR / f"exxon_{year}_reconciliation.png"
        tbl.to_csv(csv_p, index=False)
        render(tbl, year, png_p)
        vc = tbl["Verdict vs Exxon plan"].str.replace(r" ~\d+d", "", regex=True).value_counts()
        print(f"{year}: {len(tbl)} rows -> {csv_p.name}, {png_p.name}")
        for k, v in vc.items():
            print(f"      {v:2d}  {k}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
