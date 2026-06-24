"""
engine.py
Reusable aggregation core for the refinery-outage analysis suite.

Turns the raw Snowflake export (rEFINERY oUTAGES.xlsx / Query1) into clean,
analysis-ready frames, pivots and a single `build_context()` bundle. Every
downstream deliverable -- the Excel workbook, the PowerPoint deck and the HTML
dashboard -- is built on top of what this module returns, so a data refresh
only ever touches inputs here.

Locked decisions (do not re-derive -- see CLAUDE_CODE_BUILD_SPEC.md):
  * PRIMARY metric  = CAP_OFFLINE_ADJUSTED_KBD  (offline capacity, all units)
  * SECONDARY view  = mogas-equivalent (capacity x unit yield) -- overlay only
  * UNKNOWN outage type folds into UNPLANNED
  * PADD parsed from Roman-numeral PAD_DIST ; state map is a fallback
  * 2027 is planned-only -> unplanned 2027 is a *scenario*, never an actual
  * 2026 & 2027 are partial / special -> render grey-italic, footnote them
  * 2020-21 are COVID / Winter-Storm-Uri outliers -> excluded from forecast
    baselines by default
"""

import re
import numpy as np
import pandas as pd

# ----------------------------------------------------------------------------- config
RAW_SHEET = "Query1"

# logical column -> source column
COLMAP = {
    "month_date":  "OUTAGE_MONTH_DATE",
    "month":       "OUTAGE_MONTH",
    "year":        "OUTAGE_YEAR",
    "start":       "OUTAGE_START_DATE",
    "end":         "OUTAGE_END_DATE",
    "cap_kbd":     "CAP_OFFLINE_ADJUSTED_KBD",
    "cap_raw":     "OFFLINE_CAPACITY",
    "unit_cap":    "UNIT_CAPACITY",
    "duration":    "TOTAL_OUTAGE_DAYS",
    "country":     "REFINERY_COUNTRY",
    "state":       "REFINERY_STATE",
    "city":        "REFINERY_CITY",
    "operator":    "REFINERY_OPERATOR",
    "plant":       "PLANT_NAME",
    "otype":       "OUTAGE_TYPE",
    "otype2":      "OUTAGE_TYPE_LVL2",
    "unit_name":   "UNIT_NAME",
    "unit_cat":    "UNIT_CATEGORY",
    "pad_dist":    "PAD_DIST",
    "cause":       "OUTAGE_CAUSE",
    "outage_id":   "OUTAGE_ID",
}

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# Roman PADD label -> canonical "PADD n"
ROMAN_PADD = {
    "PADD I": "PADD 1", "PADD II": "PADD 2", "PADD III": "PADD 3",
    "PADD IV": "PADD 4", "PADD V": "PADD 5", "PADD CARIBBEAN": "PADD Caribbean",
}

# state -> PADD (fallback when PAD_DIST unusable)
STATE_PADD = {
    "CT": 1, "DE": 1, "DC": 1, "FL": 1, "GA": 1, "ME": 1, "MD": 1, "MA": 1,
    "NH": 1, "NJ": 1, "NY": 1, "NC": 1, "PA": 1, "RI": 1, "SC": 1, "VT": 1,
    "VA": 1, "WV": 1,
    "IL": 2, "IN": 2, "IA": 2, "KS": 2, "KY": 2, "MI": 2, "MN": 2, "MO": 2,
    "NE": 2, "ND": 2, "OH": 2, "OK": 2, "SD": 2, "TN": 2, "WI": 2,
    "AL": 3, "AR": 3, "LA": 3, "MS": 3, "NM": 3, "TX": 3,
    "CO": 4, "ID": 4, "MT": 4, "UT": 4, "WY": 4,
    "AK": 5, "AZ": 5, "CA": 5, "HI": 5, "NV": 5, "OR": 5, "WA": 5,
}

# Mogas yield buckets (from Yields.txt). Unit category -> bucket -> yield factor.
YIELD_FACTOR = {"CDU": 0.175, "FCC": 0.65, "Ref": 0.85,
                "HDC": 0.05, "Coker": 0.20, "Other": 0.0}
UNITCAT_TO_BUCKET = {
    "ATMOS DISTILLATION": "CDU",
    "VACUUM DISTILLATION": "CDU",
    "FLUID CAT CRACKING": "FCC",
    "REFORMING": "Ref",
    "HYDROCRACKING": "HDC",
    "RESID_HYDROCRACKING": "HDC",
    "COKING": "Coker",
    "THERM CRACKING, VISBREAKING": "Coker",
    "THERM CRACKING/VISBREAKING": "Coker",
    "HYDROTREATING": "Other", "ALKYLATION": "Other", "ISOMERIZATION": "Other",
    "ASPHALT": "Other", "BTX": "Other", "MTBE": "Other", "AROMATICS": "Other",
    "GAS PROCESSING": "Other", "OTHER": "Other",
}

PADD_ORDER = ["PADD 1", "PADD 2", "PADD 3", "PADD 4", "PADD 5"]

# Years that are partial / special-cased (rendered grey-italic, footnoted).
PARTIAL_YEARS = [2026, 2027]
PLANNED_ONLY_YEARS = [2027]            # no actual unplanned data exists
OUTLIER_YEARS = [2020, 2021]          # COVID / Winter Storm Uri -- excluded from baselines

# Forecast baseline windows offered in the scenario model.
BASELINE_WINDOWS = {
    "2022-2025": [2022, 2023, 2024, 2025],
    "2023-2025": [2023, 2024, 2025],
    "2018-19 & 22-25": [2018, 2019, 2022, 2023, 2024, 2025],
    "All ex-2020/21": [2014, 2015, 2016, 2017, 2018, 2019,
                       2022, 2023, 2024, 2025],
}
DEFAULT_WINDOW = "2022-2025"


# ----------------------------------------------------------------------------- load + clean
def _to_padd_from_roman(val):
    if pd.isna(val):
        return None
    return ROMAN_PADD.get(str(val).strip().upper())


def _to_padd_from_state(st):
    if pd.isna(st):
        return None
    n = STATE_PADD.get(str(st).strip().upper())
    return f"PADD {n}" if n else None


def load(path):
    """Load the raw export, return a cleaned long-form dataframe.

    The path and sheet name are stripped of stray whitespace (a past run failed
    on a leading space in the filename).
    """
    df = pd.read_excel(str(path).strip(), sheet_name=RAW_SHEET.strip())
    out = pd.DataFrame()
    for logical, src in COLMAP.items():
        out[logical] = df[src] if src in df.columns else np.nan

    out["year"] = pd.to_numeric(out["year"], errors="coerce").astype("Int64")
    out["month"] = pd.to_numeric(out["month"], errors="coerce").astype("Int64")
    out["cap_kbd"] = pd.to_numeric(out["cap_kbd"], errors="coerce").fillna(0.0)
    out["duration"] = pd.to_numeric(out["duration"], errors="coerce")

    # outage type: fold UNKNOWN -> UNPLANNED, normalize to {PLANNED, UNPLANNED}
    t = out["otype"].astype(str).str.strip().str.upper()
    out["type"] = np.where(t.eq("PLANNED"), "PLANNED", "UNPLANNED")

    # PADD: prefer Roman label, fall back to state map
    p_roman = out["pad_dist"].apply(_to_padd_from_roman)
    p_state = out["state"].apply(_to_padd_from_state)
    out["padd"] = p_roman.where(p_roman.notna(), p_state)
    out["padd_source"] = np.where(
        p_roman.notna(), "PAD_DIST",
        np.where(p_state.notna(), "STATE", "UNRESOLVED"))

    # mogas overlay
    out["bucket"] = out["unit_cat"].map(UNITCAT_TO_BUCKET).fillna("Other")
    out["mogas_kbd"] = out["cap_kbd"] * out["bucket"].map(YIELD_FACTOR).fillna(0.0)

    out["month_name"] = out["month"].map(
        lambda m: MONTHS[int(m) - 1] if pd.notna(m) and 1 <= int(m) <= 12 else None)
    return out


# ----------------------------------------------------------------------------- diagnostics
def diagnostics(df):
    return {
        "rows": len(df),
        "years": (int(df["year"].min()), int(df["year"].max())),
        "type_counts": df["type"].value_counts().to_dict(),
        "padd_source": df["padd_source"].value_counts().to_dict(),
        "padd_counts": df["padd"].value_counts(dropna=False).to_dict(),
        "unresolved_padd": int((df["padd"].isna()).sum()),
        "events_distinct": int(df["outage_id"].nunique()),
        "2027_types": df.loc[df.year.eq(2027), "type"].value_counts().to_dict(),
    }


# ----------------------------------------------------------------------------- pivots
def _month_matrix(sub, value="cap_kbd"):
    """year x month matrix (sum of value), months as Jan..Dec columns."""
    g = (sub.groupby(["year", "month"])[value].sum()
         .unstack("month").reindex(columns=range(1, 13)))
    g.columns = MONTHS
    g = g.reindex(sorted(int(y) for y in g.index if pd.notna(y)))
    return g.fillna(0.0)


def monthly_by_year(df, value="cap_kbd", type_filter=None, padd=None):
    sub = df
    if type_filter:
        sub = sub[sub["type"].eq(type_filter)]
    if padd:
        sub = sub[sub["padd"].eq(padd)]
    return _month_matrix(sub, value)


def annual_summary(df, value="cap_kbd"):
    """year x {Planned, Unplanned, Total, Events} table.

    Events = distinct OUTAGE_IDs (one physical outage can span many
    unit/month rows), which is the honest event count for a desk product.
    """
    g = df.groupby(["year", "type"])[value].sum().unstack("type").fillna(0.0)
    for c in ["PLANNED", "UNPLANNED"]:
        if c not in g:
            g[c] = 0.0
    g = g.rename(columns={"PLANNED": "Planned", "UNPLANNED": "Unplanned"})
    g["Total"] = g["Planned"] + g["Unplanned"]
    ev = df.groupby("year")["outage_id"].nunique().rename("Events")
    g = g.join(ev)
    g.index = [int(y) for y in g.index]
    return g.sort_index()


def padd_year_matrix(df, value="cap_kbd", type_filter=None):
    sub = df if not type_filter else df[df["type"].eq(type_filter)]
    sub = sub[sub["padd"].isin(PADD_ORDER)]
    g = sub.groupby(["padd", "year"])[value].sum().unstack("year").fillna(0.0)
    g = g.reindex(PADD_ORDER)
    g.columns = [int(c) for c in g.columns]
    return g


def unit_year_matrix(df, value="cap_kbd", type_filter=None):
    sub = df if not type_filter else df[df["type"].eq(type_filter)]
    g = sub.groupby(["unit_cat", "year"])[value].sum().unstack("year").fillna(0.0)
    g.columns = [int(c) for c in g.columns]
    g["__tot"] = g.sum(axis=1)
    g = g.sort_values("__tot", ascending=False).drop(columns="__tot")
    return g


def seasonality(df, years, type_filter="UNPLANNED", padd=None, value="cap_kbd"):
    """Average monthly profile (kbd/month) across the given years.

    This is the backbone of the 2027 scenario: the mean calendar-month shape of
    unplanned capacity offline over a chosen baseline window.
    """
    sub = df[df["type"].eq(type_filter) & df["year"].isin(years)]
    if padd:
        sub = sub[sub["padd"].eq(padd)]
    by_ym = (sub.groupby(["year", "month"])[value].sum()
             .unstack("month").reindex(columns=range(1, 13)).fillna(0.0))
    # reindex rows so windows with a missing year still divide by len(years)
    by_ym = by_ym.reindex(years).fillna(0.0)
    prof = by_ym.mean(axis=0)
    prof.index = MONTHS
    return prof


def padd_month_year(df, padd, value="cap_kbd", type_filter=None):
    """For one PADD: year x month matrix (the per-PADD chart backbone)."""
    sub = df[df["padd"].eq(padd)]
    if type_filter:
        sub = sub[sub["type"].eq(type_filter)]
    return _month_matrix(sub, value)


def operator_year(df, value="cap_kbd", type_filter=None, top=12):
    sub = df if not type_filter else df[df["type"].eq(type_filter)]
    g = sub.groupby(["operator", "year"])[value].sum().unstack("year").fillna(0.0)
    g.columns = [int(c) for c in g.columns]
    g["__t"] = g.sum(axis=1)
    g = g.sort_values("__t", ascending=False).drop(columns="__t").head(top)
    return g


def plant_detail(df, value="cap_kbd", top=15, years=None):
    """Refinery-level detail: PADD, operator, total / planned / unplanned / events."""
    sub = df if years is None else df[df["year"].isin(years)]
    grp = sub.groupby(["plant", "padd", "operator"])
    base = grp.agg(total=(value, "sum"), events=("outage_id", "nunique"))
    planned = (sub[sub["type"].eq("PLANNED")]
               .groupby(["plant", "padd", "operator"])[value].sum().rename("planned"))
    out = base.join(planned).reset_index()
    out["planned"] = out["planned"].fillna(0.0)
    out["unplanned"] = out["total"] - out["planned"]
    return out.sort_values("total", ascending=False).head(top)


def event_scatter(df, years, type_filter="UNPLANNED", top=400):
    """Event-level points: monthly intensity (duration days) vs capacity (kbd).

    One marker per outage-month slice for recent unplanned events, used by the
    Refinery-Detail scatter. Returns the largest `top` by capacity so the chart
    stays readable.
    """
    sub = df[df["type"].eq(type_filter) & df["year"].isin(years)].copy()
    sub = sub[(sub["cap_kbd"] > 0) & sub["duration"].notna()]
    cols = ["duration", "cap_kbd", "padd", "plant", "year"]
    sub = sub[cols].sort_values("cap_kbd", ascending=False).head(top)
    return sub.reset_index(drop=True)


# ----------------------------------------------------------------------------- mogas overlay
def mogas_annual(df):
    """year x {Planned, Unplanned, Total} mogas-equivalent kbd."""
    g = (df.groupby(["year", "type"])["mogas_kbd"].sum()
         .unstack("type").fillna(0.0))
    for c in ["PLANNED", "UNPLANNED"]:
        if c not in g:
            g[c] = 0.0
    g = g.rename(columns={"PLANNED": "Planned", "UNPLANNED": "Unplanned"})
    g["Total"] = g["Planned"] + g["Unplanned"]
    g.index = [int(y) for y in g.index]
    return g.sort_index()


def mogas_yield_map():
    """bucket -> (factor, [unit categories]) for the Mogas-Overlay yield table."""
    rows = []
    for bucket, factor in YIELD_FACTOR.items():
        cats = sorted(c for c, b in UNITCAT_TO_BUCKET.items() if b == bucket)
        rows.append((bucket, factor, cats))
    return rows


# ----------------------------------------------------------------------------- comparisons (2027 guardrail)
def yoy_delta(df, value="cap_kbd"):
    """Annual summary with YoY delta + unplanned-% columns added."""
    s = annual_summary(df, value).copy()
    s["Unpl%"] = np.where(s["Total"] > 0, s["Unplanned"] / s["Total"], np.nan)
    s["YoY"] = s["Total"].diff()
    s["YoY%"] = s["Total"].pct_change()
    return s


def compare_block(df, year_a, year_b, value="cap_kbd"):
    """Plan+Unplanned / Planned / Unplanned totals for two years + % delta.

    Enforces the 2027 guardrail: any non-planned metric that touches a
    planned-only year (2027) returns None (renders as n/a).
    """
    s = annual_summary(df, value)

    def grab(year, metric):
        if year not in s.index:
            return None
        if year in PLANNED_ONLY_YEARS and metric in ("Total", "Unplanned"):
            return None
        return float(s.loc[year, metric])

    out = {}
    for metric, label in [("Total", "Plan + Unplanned"),
                          ("Planned", "Planned"),
                          ("Unplanned", "Unplanned")]:
        a, b = grab(year_a, metric), grab(year_b, metric)
        delta = (b - a) if (a is not None and b is not None) else None
        pct = (delta / a) if (delta is not None and a not in (None, 0)) else None
        out[label] = {"a": a, "b": b, "delta": delta, "pct": pct}
    return out


# ----------------------------------------------------------------------------- scenario model
def baseline_profile(df, window_key, padd=None):
    """Monthly unplanned profile (kbd/month) for a named baseline window."""
    years = BASELINE_WINDOWS[window_key]
    return seasonality(df, years, type_filter="UNPLANNED", padd=padd)


def scenario_forecast(df, window_key=DEFAULT_WINDOW, growth=0.0, multiplier=1.0,
                      oneoff_kbd=0.0, stress_month="Sep"):
    """Driver-based 2027 unplanned forecast (the default Python view).

    Mirrors the live Excel cascade so the deck/dashboard show the same numbers
    the workbook recomputes from its inputs:

        baseline(window) x (1+growth) x multiplier
        + one-off (added to the stress month only)
        = 2027 unplanned forecast (monthly, summed to annual)
    """
    prof = baseline_profile(df, window_key)
    scaled = prof * (1.0 + growth) * multiplier
    monthly = scaled.copy()
    if stress_month in MONTHS:
        monthly[stress_month] = monthly[stress_month] + oneoff_kbd
    planned_2027 = float(annual_summary(df).loc[2027, "Planned"]) \
        if 2027 in annual_summary(df).index else 0.0
    return {
        "window": window_key,
        "growth": growth,
        "multiplier": multiplier,
        "oneoff_kbd": oneoff_kbd,
        "stress_month": stress_month,
        "monthly_unplanned": monthly,
        "annual_unplanned": float(monthly.sum()),
        "planned_2027": planned_2027,
        "implied_total": float(monthly.sum()) + planned_2027,
    }


def padd_unplanned_share(df, window_key=DEFAULT_WINDOW):
    """Historical PADD share of unplanned capacity over the baseline window."""
    years = BASELINE_WINDOWS[window_key]
    sub = df[df["type"].eq("UNPLANNED") & df["year"].isin(years)
             & df["padd"].isin(PADD_ORDER)]
    by_padd = sub.groupby("padd")["cap_kbd"].sum().reindex(PADD_ORDER).fillna(0.0)
    total = by_padd.sum()
    return (by_padd / total) if total else by_padd


def tornado(df, window_key=DEFAULT_WINDOW):
    """Low / Base / High swing for each scenario driver, sorted by swing.

    Used by the Sensitivity sheet's tornado diagram. Base case is
    growth=0, multiplier=1.0, no one-off.
    """
    base_annual = float(baseline_profile(df, window_key).sum())
    drivers = []

    # unplanned multiplier +/-30%
    drivers.append(("Unplanned rate multiplier (0.7x / 1.3x)",
                    base_annual * 0.7, base_annual, base_annual * 1.3))
    # production growth +/-10%
    drivers.append(("Production growth (-10% / +10%)",
                    base_annual * 0.9, base_annual, base_annual * 1.1))
    # baseline-window swing +/-15% (proxy for window choice)
    drivers.append(("Baseline window swing (+/-15%)",
                    base_annual * 0.85, base_annual, base_annual * 1.15))
    # one-off event +300 kbd (one-sided)
    drivers.append(("One-off event (+300 kbd)",
                    base_annual, base_annual, base_annual + 300.0))

    rows = []
    for name, low, base, high in drivers:
        rows.append({"driver": name, "low": low, "base": base, "high": high,
                     "swing": abs(high - low)})
    rows.sort(key=lambda r: r["swing"], reverse=True)
    return rows


# ----------------------------------------------------------------------------- context bundle
def build_context(path):
    """One-shot bundle of every frame the deliverables need.

    The slide deck and HTML dashboard consume this so they never re-aggregate
    raw data; the Excel workbook also uses it for its data blocks (its
    interactive scenario/sensitivity cells are live formulas, not these values).
    """
    df = load(path)
    diag = diagnostics(df)
    summary = yoy_delta(df)

    ctx = {
        "df": df,
        "diag": diag,
        "summary": summary,                         # annual w/ YoY + Unpl%
        "padd_total": padd_year_matrix(df),
        "padd_unplanned": padd_year_matrix(df, type_filter="UNPLANNED"),
        "padd_planned": padd_year_matrix(df, type_filter="PLANNED"),
        "unit_total": unit_year_matrix(df),
        "unit_unplanned": unit_year_matrix(df, type_filter="UNPLANNED"),
        "monthly_total": monthly_by_year(df),
        "monthly_planned": monthly_by_year(df, type_filter="PLANNED"),
        "monthly_unplanned": monthly_by_year(df, type_filter="UNPLANNED"),
        "operators": operator_year(df, top=10),
        "plants": plant_detail(df, top=15),
        "scatter": event_scatter(df, [2023, 2024, 2025]),
        "mogas_annual": mogas_annual(df),
        "mogas_yield_map": mogas_yield_map(),
        "scenario": scenario_forecast(df),
        "padd_share": padd_unplanned_share(df),
        "tornado": tornado(df),
        "padd_month": {p: {
            "total": padd_month_year(df, p),
            "planned": padd_month_year(df, p, type_filter="PLANNED"),
            "unplanned": padd_month_year(df, p, type_filter="UNPLANNED"),
        } for p in PADD_ORDER},
        "compare": {
            "2025v2026": compare_block(df, 2025, 2026),
            "2025v2027": compare_block(df, 2025, 2027),
            "2026v2027": compare_block(df, 2026, 2027),
        },
    }
    return ctx


if __name__ == "__main__":
    import json
    import sys
    df = load(sys.argv[1] if len(sys.argv) > 1 else "rEFINERY oUTAGES.xlsx")
    print(json.dumps(diagnostics(df), indent=2, default=str))
