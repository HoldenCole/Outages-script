#!/usr/bin/env python3
"""Monthly KBD aggregation engine for the refinery outage workbook.

Batch 1 of the Excel-workbook build. This module is intentionally standalone so
it does not disturb the working outage_analyzer.py / HTML pipeline. It turns the
raw outage export into the PADD x month x year matrix (in KBD) that the example
"Offline Mogas Production (KBD)" workbook is built from, split by OUTAGE_TYPE
(Planned / Unplanned / combined) for Total US and each PADD 1-5.

Goal for this batch: PROVE THE NUMBERS before writing any Excel. The --validate
mode prints the PADD-2 2025 and 2026 monthly rows under two candidate
conventions so we can confirm which one reproduces the known ground-truth row
from the example sheet:

    2025 PADD-2 Plan+Unplanned: 47 100 210 371 245 45 60 80 159 313 83 16

Conventions tested:
  A) "as_is": each row is already a single outage-month slice, so the month's
     value is simply the SUM of CAP_OFFLINE_ADJUSTED_KBD over rows in
     (region, type, year, month).
  B) "prorated": CAP_OFFLINE_ADJUSTED_KBD is a whole-outage figure that must be
     spread across the months it touches, weighted by days in each month.

Whichever matches becomes the locked convention for Batch 2 (Excel tables,
%-deltas, totals/averages, conditional formatting) and Batch 3 (bar+line charts).

Usage:
    python outage_monthly.py "rEFINERY oUTAGES.xlsx" --validate
    python outage_monthly.py "rEFINERY oUTAGES.xlsx"            # summary only
"""
import argparse
import re
import sys

import numpy as np
import pandas as pd


# Logical field -> actual Excel header (matched case/space/punct-insensitive).
# Mirrors outage_analyzer.py and adds the two columns this workbook needs:
# CAP_OFFLINE_ADJUSTED_KBD (the KBD value) and OUTAGE_TYPE (planned/unplanned).
COLUMN_MAP = {
    "kbd": "CAP_OFFLINE_ADJUSTED_KBD",
    "outage_type": "OUTAGE_TYPE",
    "year": "OUTAGE_YEAR",
    "month": "OUTAGE_MONTH",
    "month_date": "OUTAGE_MONTH_DATE",
    "start_date": "OUTAGE_START_DATE",
    "end_date": "OUTAGE_END_DATE",
    "pct_month": "PERCENTAGE_MONTH",
    "pct_month_cal": "PERCENTAGE_MONTH_CAL",
    "total_days": "TOTAL_OUTAGE_DAYS",
    "month_days": "TOTAL_MONTH_DAYS",
    "padd": "PAD_DIST",
    "unit_type": "UNIT_CATEGORY",
    "refinery": "PLANT_NAME",
    "offline_cap": "OFFLINE_CAPACITY",
}

# Fields we cannot proceed without for the monthly matrix.
REQUIRED = ["kbd", "year", "padd"]

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# Known ground-truth row from the example workbook, used by --validate.
GROUND_TRUTH_PADD2_2025 = [47, 100, 210, 371, 245, 45, 60, 80, 159, 313, 83, 16]


def _norm(s):
    return re.sub(r"[^a-z0-9]+", "", str(s).strip().lower())


def _classify_type(v):
    """Map a raw OUTAGE_TYPE value to one of planned / unplanned / unknown."""
    t = _norm(v)
    if t.startswith("plan"):
        return "planned"
    if t.startswith("unplan") or t.startswith("unscheduled") or t == "forced":
        return "unplanned"
    return "unknown"


def load(path):
    """Read the export and return a tidy frame with normalized logical columns."""
    try:
        raw = pd.read_excel(path)
    except FileNotFoundError:
        raise SystemExit(f"ERROR: could not read Excel file '{path}': file not found.")
    except Exception as e:
        raise SystemExit(f"ERROR: could not read Excel file '{path}': {e}")

    actual = {_norm(c): c for c in raw.columns}
    rename, found, missing = {}, {}, []
    for logical, header in COLUMN_MAP.items():
        key = _norm(header)
        if key in actual:
            rename[actual[key]] = logical
            found[logical] = actual[key]
        else:
            missing.append((logical, header))

    df = raw.rename(columns=rename)
    df = df[[c for c in df.columns if c in COLUMN_MAP]].copy()

    missing_required = [m for m in missing if m[0] in REQUIRED]
    if missing_required:
        lines = "\n".join(f" - logical '{lg}' expected header '{hd}'" for lg, hd in missing_required)
        have = "\n".join(f" - {c}" for c in raw.columns)
        raise SystemExit(
            "ERROR: required columns for the monthly matrix not found.\n"
            + lines
            + "\n\nHeaders present in the file:\n"
            + have
        )

    # Numerics.
    for c in ("kbd", "year", "month", "pct_month", "pct_month_cal", "total_days", "month_days", "offline_cap"):
        if c in df:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    for c in ("start_date", "end_date", "month_date"):
        if c in df:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    # Derive month from a date if the month column is absent/empty.
    if ("month" not in df or df["month"].isna().all()):
        if "month_date" in df and df["month_date"].notna().any():
            df["month"] = df["month_date"].dt.month
        elif "start_date" in df:
            df["month"] = df["start_date"].dt.month

    # PADD -> integer 1..5 (handles "PADD 3" or numeric).
    if "padd" in df:
        df["padd"] = df["padd"].astype(str).str.extract(r"(\d)").astype(float)

    # Outage type bucket.
    if "outage_type" in df:
        df["type_bucket"] = df["outage_type"].map(_classify_type)
    else:
        df["type_bucket"] = "unknown"

    return df, found, [m for m in missing if m[0] not in REQUIRED]


def _region_label(padd):
    return f"PADD {int(padd)}" if pd.notna(padd) else "PADD ?"


def monthly_matrix(df, value="kbd", convention="as_is"):
    """Return a tidy frame: region, type_bucket, year, month(1-12), value(KBD).

    convention:
      "as_is"    -> month value = sum of `value` within (region, type, year, month).
                    Assumes each row is already a single outage-month slice.
      "prorated" -> spread each row's `value` across the calendar months between
                    start_date and end_date, weighted by days in each month.
                    Falls back to as_is when dates are unavailable.
    """
    d = df.dropna(subset=["year"]).copy()
    d = d[d["month"].notna()]
    d["year"] = d["year"].astype(int)
    d["month"] = d["month"].astype(int)
    d = d[(d["month"] >= 1) & (d["month"] <= 12)]

    if convention == "prorated" and "start_date" in d and "end_date" in d:
        rows = []
        for _, r in d.iterrows():
            v = r.get(value)
            if pd.isna(v):
                continue
            s, e = r.get("start_date"), r.get("end_date")
            if pd.isna(s) or pd.isna(e) or e < s:
                rows.append((r["padd"], r["type_bucket"], int(r["year"]), int(r["month"]), float(v)))
                continue
            # Build the set of (year, month) the outage spans, weighted by days.
            span = pd.date_range(s.normalize(), e.normalize(), freq="D")
            if len(span) == 0:
                rows.append((r["padd"], r["type_bucket"], int(r["year"]), int(r["month"]), float(v)))
                continue
            counts = {}
            for day in span:
                counts[(day.year, day.month)] = counts.get((day.year, day.month), 0) + 1
            total = sum(counts.values())
            for (yy, mm), c in counts.items():
                rows.append((r["padd"], r["type_bucket"], yy, mm, float(v) * c / total))
        tidy = pd.DataFrame(rows, columns=["padd", "type_bucket", "year", "month", "value"])
    else:
        tidy = d.rename(columns={value: "value"})[["padd", "type_bucket", "year", "month", "value"]].copy()

    tidy["region"] = tidy["padd"].map(_region_label)

    # Per-PADD rollup.
    per_padd = (
        tidy.groupby(["region", "type_bucket", "year", "month"], dropna=False)["value"]
        .sum()
        .reset_index()
    )
    # Total US rollup (all PADDs combined).
    total_us = (
        tidy.groupby(["type_bucket", "year", "month"], dropna=False)["value"]
        .sum()
        .reset_index()
    )
    total_us["region"] = "Total US"
    out = pd.concat([per_padd, total_us[["region", "type_bucket", "year", "month", "value"]]], ignore_index=True)
    return out


def combined_over_types(matrix):
    """Collapse planned+unplanned+unknown into a single 'all' bucket per cell."""
    allt = (
        matrix.groupby(["region", "year", "month"])["value"].sum().reset_index()
    )
    allt["type_bucket"] = "all"
    return allt[["region", "type_bucket", "year", "month", "value"]]


def monthly_row(matrix, region, year, type_bucket):
    """Return the 12-month list (Jan..Dec) for one region/year/type, 0 where absent."""
    sub = matrix[(matrix["region"] == region) & (matrix["year"] == year) & (matrix["type_bucket"] == type_bucket)]
    by_month = sub.set_index("month")["value"].to_dict()
    return [round(by_month.get(m, 0.0), 1) for m in range(1, 13)]


def period_summaries(row):
    """Total (12-mo avg), 1H avg (Jan-Jun), 2H avg (Jul-Dec) over present months."""
    vals = [v for v in row if v is not None]
    h1 = [row[i] for i in range(0, 6) if row[i] is not None]
    h2 = [row[i] for i in range(6, 12) if row[i] is not None]
    avg = lambda xs: round(sum(xs) / len(xs), 1) if xs else 0.0
    return {"total_avg": avg(vals), "h1_avg": avg(h1), "h2_avg": avg(h2)}


def print_validation(df):
    print("=" * 78)
    print("VALIDATION - does the engine reproduce the example workbook's PADD-2 row?")
    print("=" * 78)

    print("\nOUTAGE_TYPE value counts (raw -> bucket):")
    if "outage_type" in df:
        vc = df.groupby([df["outage_type"].astype(str), "type_bucket"]).size().reset_index(name="rows")
        for _, r in vc.iterrows():
            print(f"  {r['outage_type']!r:40s} -> {r['type_bucket']:9s} {int(r['rows']):>8d}")
    else:
        print("  (no OUTAGE_TYPE column found)")

    print("\nRows per year:")
    yr = df.dropna(subset=["year"]).copy()
    yr["year"] = yr["year"].astype(int)
    for y, c in yr["year"].value_counts().sort_index().items():
        print(f"  {y}: {int(c):>8d}")

    for conv in ("as_is", "prorated"):
        m = monthly_matrix(df, value="kbd", convention=conv)
        allt = combined_over_types(m)
        full = pd.concat([m, allt], ignore_index=True)
        print("\n" + "-" * 78)
        print(f"Convention: {conv}")
        print("-" * 78)
        for yr_ in (2025, 2026):
            row = monthly_row(full, "PADD 2", yr_, "all")
            s = period_summaries(row)
            print(f"  PADD 2 {yr_} (Plan+Unplanned):")
            print("    " + "  ".join(f"{MONTHS[i]}:{row[i]:>6.0f}" for i in range(12)))
            print(f"    Total(avg)={s['total_avg']:.0f}  1H={s['h1_avg']:.0f}  2H={s['h2_avg']:.0f}")
        # Compare 2025 to ground truth.
        row25 = monthly_row(full, "PADD 2", 2025, "all")
        diffs = [round(row25[i] - GROUND_TRUTH_PADD2_2025[i], 1) for i in range(12)]
        close = all(abs(x) <= 2 for x in diffs)  # allow small rounding wobble
        print("    ground truth 2025:  " + "  ".join(f"{MONTHS[i]}:{GROUND_TRUTH_PADD2_2025[i]:>6d}" for i in range(12)))
        print("    delta vs truth:     " + "  ".join(f"{MONTHS[i]}:{diffs[i]:>6.0f}" for i in range(12)))
        print(f"    --> {'MATCH (within +/-2)' if close else 'NO MATCH'} for convention '{conv}'")

    print("\n" + "=" * 78)
    print("Pick whichever convention shows MATCH; that is what Batch 2 will use.")
    print("=" * 78)


def print_summary(df, found, missing_optional):
    print(f"Loaded {len(df):,} rows.")
    print("Matched columns:")
    for lg, hd in found.items():
        print(f"  - {lg}: {hd}")
    if missing_optional:
        print("Optional columns not found:")
        for lg, hd in missing_optional:
            print(f"  - {lg} (expected '{hd}')")
    if "type_bucket" in df:
        print("Type buckets:", df["type_bucket"].value_counts().to_dict())


def main():
    ap = argparse.ArgumentParser(description="Monthly KBD aggregation engine (Batch 1)")
    ap.add_argument("excel", help="path to the outage .xlsx export")
    ap.add_argument("--validate", action="store_true", help="print PADD-2 reproduction check under both conventions")
    args = ap.parse_args()

    df, found, missing_optional = load(args.excel)
    print_summary(df, found, missing_optional)
    if args.validate:
        print()
        print_validation(df)


if __name__ == "__main__":
    try:
        main()
    except SystemExit as e:
        if isinstance(e.code, str):
            print(e.code)
            sys.exit(1)
        raise
