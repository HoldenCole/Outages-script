"""
engine.py
Reusable aggregation core for the refinery-outage analysis suite.

Turns the raw Snowflake export (Refinery_Outages_Data.xlsx / Query1) into clean,
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


# ----------------------------------------------------------------------------- back-to-back clusters
def _span(months):
    return f"{MONTHS[months[0] - 1]}-{MONTHS[months[-1] - 1]}" if len(months) > 1 \
        else MONTHS[months[0] - 1]


def consecutive_runs(df, operator_contains=None, unit_cat=None, min_len=3,
                     years=None, exclude_years=None):
    """Detect back-to-back (consecutive calendar-month) outage runs at the same
    plant within a year.

    This is the granular signal that month-aggregated external trackers wash
    out: e.g. the recurring Q1 FCC turnaround clusters at the ExxonMobil plants
    (Baton Rouge, Baytown, Beaumont, Joliet). Returns a list of dicts sorted by
    run length then capacity:
        plant, operator, padd, year, months[list], span, n, kbd, unplanned_kbd,
        planned_kbd, unpl_share.
    """
    sub = df.copy()
    if unit_cat:
        sub = sub[sub["unit_cat"].eq(unit_cat)]
    if operator_contains:
        sub = sub[sub["operator"].astype(str).str.upper()
                  .str.contains(operator_contains.upper(), na=False)]
    if years is not None:
        sub = sub[sub["year"].isin(years)]
    if exclude_years:
        sub = sub[~sub["year"].isin(exclude_years)]

    out = []
    for (plant, year), g in sub.groupby(["plant", "year"]):
        months = sorted(int(x) for x in g["month"].dropna().unique())
        if not months:
            continue
        runs, run = [], [months[0]]
        for m in months[1:]:
            if m == run[-1] + 1:
                run.append(m)
            else:
                runs.append(run)
                run = [m]
        runs.append(run)
        for r in runs:
            if len(r) < min_len:
                continue
            gg = g[g["month"].isin(r)]
            kbd = float(gg["cap_kbd"].sum())
            unpl = float(gg[gg["type"].eq("UNPLANNED")]["cap_kbd"].sum())
            out.append({
                "plant": str(plant), "operator": str(g["operator"].iloc[0]),
                "padd": str(g["padd"].iloc[0]), "year": int(year),
                "months": r, "span": _span(r), "n": len(r),
                "kbd": kbd, "unplanned_kbd": unpl, "planned_kbd": kbd - unpl,
                "unpl_share": (unpl / kbd if kbd else 0.0),
            })
    out.sort(key=lambda d: (-d["n"], -d["kbd"]))
    return out


def fcc_exxon_clusters(df, min_len=3):
    """The headline finding: ExxonMobil FCC back-to-back runs, COVID year excluded."""
    return consecutive_runs(df, operator_contains="EXXON", unit_cat="FLUID CAT CRACKING",
                            min_len=min_len, exclude_years=OUTLIER_YEARS)


def exxon_fcc_month_grid(df, years):
    """plant x month FCC capacity (kbd) for ExxonMobil, for the given years.

    Backs the back-to-back FCC heat-strip: consecutive shaded cells are runs.
    Returns {(plant, year): [12 kbd]} for plants with any FCC outage in `years`.
    """
    m = (df["operator"].astype(str).str.upper().str.contains("EXXON", na=False)
         & df["unit_cat"].eq("FLUID CAT CRACKING") & df["year"].isin(years))
    sub = df[m]
    grid = {}
    for (plant, year), g in sub.groupby(["plant", "year"]):
        row = [0.0] * 12
        for _, r in g.iterrows():
            if pd.notna(r["month"]):
                row[int(r["month"]) - 1] += float(r["cap_kbd"])
        grid[(str(plant), int(year))] = row
    return grid


# ----------------------------------------------------------------------------- percentage views
def padd_yoy(df, type_filter="UNPLANNED"):
    """(levels matrix, YoY% matrix) for capacity offline by PADD x year."""
    m = padd_year_matrix(df, type_filter=type_filter)
    yoy = m.pct_change(axis=1)
    return m, yoy


def unit_share(df, type_filter=None):
    """Unit-category share of total capacity offline per year (prod-by-unit %)."""
    m = unit_year_matrix(df, type_filter=type_filter)
    col_tot = m.sum(axis=0).replace(0, np.nan)
    return m.div(col_tot, axis=1).fillna(0.0)


def scenario_by_padd(df, window_key=DEFAULT_WINDOW, growth=0.0, multiplier=1.0):
    """Per-PADD 2027 unplanned scenario: each PADD carries its own seasonality.

    More honest than splitting one national number by a flat share - PADD 3 and
    PADD 5 have different monthly shapes.
    """
    res = {}
    for p in PADD_ORDER:
        prof = seasonality(df, BASELINE_WINDOWS[window_key], "UNPLANNED", padd=p)
        scaled = prof * (1.0 + growth) * multiplier
        res[p] = {"baseline_annual": float(prof.sum()),
                  "monthly": scaled, "annual": float(scaled.sum())}
    return res


def monthly_yoy(df, type_filter="UNPLANNED"):
    """year x month matrix of YoY % change (each cell vs the same month a year
    earlier). The 'percent difference in each month by year' view."""
    m = monthly_by_year(df, type_filter=type_filter)
    yoy = m.pct_change(axis=0)            # down the years, per month column
    return m, yoy


def monthly_range_band(df, type_filter="UNPLANNED", years=None, padd=None):
    """Per-calendar-month min / max / avg across `years` (default the baseline
    window) - the shaded 'N-yr range' band behind the seasonality lines."""
    if years is None:
        years = BASELINE_WINDOWS[DEFAULT_WINDOW]
    sub = df[df["type"].eq(type_filter) & df["year"].isin(years)]
    if padd:
        sub = sub[sub["padd"].eq(padd)]
    by_ym = (sub.groupby(["year", "month"])["cap_kbd"].sum()
             .unstack("month").reindex(index=years, columns=range(1, 13)).fillna(0.0))
    out = pd.DataFrame({"min": by_ym.min(axis=0), "max": by_ym.max(axis=0),
                        "avg": by_ym.mean(axis=0)})
    out.index = MONTHS
    return out


def turnaround_schedule(df, year, padd=None, type_filter="PLANNED", top=40):
    """Event-level outage schedule (the 'Fall TAs' tables): one row per
    outage/unit with operator, refinery, PADD, offline kbd, % of PADD, start &
    end dates, type and unit. Sorted by start date."""
    sub = df[(df["year"] == year) & df["type"].eq(type_filter)]
    if padd:
        sub = sub[sub["padd"].eq(padd)]
    if sub.empty:
        return pd.DataFrame(columns=["operator", "plant", "padd", "unit_cat",
                                     "kbd", "pct_padd", "start", "end", "type"])
    g = sub.groupby(["outage_id", "operator", "plant", "padd", "unit_cat"]).agg(
        kbd=("cap_kbd", "mean"), start=("start", "min"), end=("end", "max")).reset_index()
    padd_tot = sub.groupby("padd")["cap_kbd"].mean().to_dict()  # avg offline rate per PADD
    denom = sub.groupby("padd")["cap_kbd"].sum().to_dict()
    g["pct_padd"] = g.apply(
        lambda r: (r["kbd"] / denom[r["padd"]]) if denom.get(r["padd"]) else 0.0, axis=1)
    g["type"] = type_filter.title()
    g = g.sort_values(["kbd"], ascending=False).head(top).sort_values("start")
    return g[["operator", "plant", "padd", "unit_cat", "kbd", "pct_padd", "start", "end", "type"]]


def tidy_monthly(df, dim="padd", years=range(2014, 2028)):
    """Long backing table for the interactive Explorer's SUMIFS.

    Columns: year, month(1-12), key, type, kbd. `dim` is 'padd' or 'unit_cat'.
    Includes pre-aggregated rollups ('Total US'/'All Units' and type 'All') so a
    dropdown value maps to exactly one summable slice.
    """
    years = set(int(y) for y in years)
    if dim == "padd":
        opts = ["Total US"] + PADD_ORDER
        rollup = "Total US"
    else:
        present = [u for u in df["unit_cat"].dropna().unique()]
        # order units by total offline, cap to keep the table lean
        order = (df.groupby("unit_cat")["cap_kbd"].sum().sort_values(ascending=False).index.tolist())
        opts = ["All Units"] + [u for u in order if u in present]
        rollup = "All Units"
    rows = []
    for o in opts:
        sub_o = df if o == rollup else df[df[dim].eq(o)]
        for t in ["All", "PLANNED", "UNPLANNED"]:
            sub = sub_o if t == "All" else sub_o[sub_o["type"].eq(t)]
            g = sub.groupby(["year", "month"])[["cap_kbd", "mogas_kbd"]].sum()
            tlab = "All" if t == "All" else t.title()
            for (y, m), rr in g.iterrows():
                cap, mog = float(rr["cap_kbd"]), float(rr["mogas_kbd"])
                if pd.notna(y) and pd.notna(m) and int(y) in years and 1 <= int(m) <= 12 and (cap or mog):
                    rows.append((int(y), int(m), str(o), tlab, cap, mog))
    return pd.DataFrame(rows, columns=["year", "month", "key", "type", "kbd", "mogas"])


# ----------------------------------------------------------------------------- naphtha / octane
NAPHTHA_UNITS = ["REFORMING", "ISOMERIZATION", "AROMATICS", "BTX"]


def naphtha_analysis(df):
    """The naphtha / octane complex: catalytic reforming (naphtha -> high-octane
    reformate), isomerization (light naphtha -> isomerate) and aromatics/BTX.
    When these are offline, gasoline octane and blending are squeezed even if
    crude runs hold - a read external CDU-only trackers miss.
    """
    sub = df[df["unit_cat"].isin(NAPHTHA_UNITS)]
    annual = sub.groupby(["year", "type"])["cap_kbd"].sum().unstack("type").fillna(0.0)
    for c in ["PLANNED", "UNPLANNED"]:
        if c not in annual:
            annual[c] = 0.0
    annual = annual.rename(columns={"PLANNED": "Planned", "UNPLANNED": "Unplanned"})
    annual["Total"] = annual["Planned"] + annual["Unplanned"]
    annual.index = [int(y) for y in annual.index]
    monthly_unpl = _month_matrix(sub[sub["type"].eq("UNPLANNED")])
    by_padd = (sub[sub["padd"].isin(PADD_ORDER)].groupby(["padd", "year"])["cap_kbd"].sum()
               .unstack("year").fillna(0.0))
    by_padd.columns = [int(c) for c in by_padd.columns]
    by_unit = sub.groupby("unit_cat")["cap_kbd"].sum().sort_values(ascending=False)
    # share of reforming offline that is naphtha-octane vs total offline
    return {"annual": annual.sort_index(), "monthly_unpl": monthly_unpl,
            "by_padd": by_padd, "by_unit": by_unit, "units": NAPHTHA_UNITS}


def top_movers(df):
    """Auto-insight one-liners for the dashboard commentary block."""
    s = yoy_delta(df)
    ly = max(y for y in s.index if y not in PARTIAL_YEARS and s.loc[y, "Unplanned"] > 0)
    lines = []
    _, pyoy = padd_yoy(df, "UNPLANNED")
    if ly in pyoy.columns:
        mv = pyoy[ly].dropna()
        if len(mv):
            up, dn = mv.idxmax(), mv.idxmin()
            lines.append(f"{up} unplanned {pyoy.loc[up, ly]:+.0%} YoY in {ly} (biggest riser); "
                         f"{dn} {pyoy.loc[dn, ly]:+.0%} (biggest faller).")
    mu = monthly_by_year(df, type_filter="UNPLANNED")
    if ly in mu.index:
        pk = mu.loc[ly].idxmax()
        lines.append(f"{ly} unplanned peaked in {pk} ({mu.loc[ly, pk]:,.0f} kbd) - watch the "
                     "Feb-freeze / autumn-turnaround windows.")
    nap = naphtha_analysis(df)["annual"]
    if ly in nap.index and (ly - 1) in nap.index and nap.loc[ly - 1, "Total"]:
        ch = nap.loc[ly, "Total"] / nap.loc[ly - 1, "Total"] - 1
        lines.append(f"Naphtha/octane-complex offline {nap.loc[ly, 'Total']:,.0f} kbd in {ly} "
                     f"({ch:+.0%} YoY) - direct octane/blending read.")
    fc = scenario_forecast(df)
    lines.append(f"2027 implied total ~{fc['implied_total']:,.0f} kbd "
                 f"({fc['planned_2027']:,.0f} planned + ~{fc['annual_unplanned']:,.0f} modeled unplanned).")
    return lines


def scenario_bands(df, window_key=DEFAULT_WINDOW):
    """P25 / P50 / P90 of historical annual unplanned over the window -> a range
    around the point forecast."""
    years = BASELINE_WINDOWS[window_key]
    annuals = [df[df["type"].eq("UNPLANNED") & df["year"].eq(y)]["cap_kbd"].sum() for y in years]
    annuals = [a for a in annuals if a]
    if not annuals:
        return {"p25": 0.0, "p50": 0.0, "p90": 0.0, "mean": 0.0}
    return {"p25": float(np.percentile(annuals, 25)), "p50": float(np.percentile(annuals, 50)),
            "p90": float(np.percentile(annuals, 90)), "mean": float(np.mean(annuals))}


def exxon_2027_breakdown(df, operator_contains="EXXON", year=2027):
    """Breakdown of a single operator's planned outages for one year (default
    ExxonMobil 2027): offline kbd by refinery, by unit, and a refinery x month
    matrix for a stacked chart. 2027 is planned-only, so this is the booked book."""
    sub = df[df["operator"].astype(str).str.upper().str.contains(operator_contains.upper(), na=False)
             & df["year"].eq(year) & df["type"].eq("PLANNED")]
    by_ref = sub.groupby("plant")["cap_kbd"].sum().sort_values(ascending=False)
    by_unit = sub.groupby("unit_cat")["cap_kbd"].sum().sort_values(ascending=False)
    refs = by_ref.index.tolist()
    month_ref = {str(r): [float(sub[(sub["plant"].eq(r)) & (sub["month"].eq(m))]["cap_kbd"].sum())
                          for m in range(1, 13)] for r in refs}
    return {"by_ref": by_ref, "by_unit": by_unit, "month_ref": month_ref,
            "refs": [str(r) for r in refs], "total": float(sub["cap_kbd"].sum())}


# ----------------------------------------------------------------------------- outliers, %, forecasts
OUTLIER_NOTE = {2020: "COVID-19 demand collapse", 2021: "Winter Storm Uri freeze"}


def flatten_outliers(series, years=tuple(OUTLIER_YEARS)):
    """Cap outlier-year values to the max normal year so charts stay readable;
    return (display_series, footnote) with the real numbers in the footnote."""
    s = series.copy()
    normal = [float(v) for y, v in s.items() if int(y) not in years]
    cap = max(normal) if normal else float(s.max())
    notes = []
    for y in years:
        if y in s.index and float(s.loc[y]) > cap:
            notes.append(f"{y} ({OUTLIER_NOTE.get(y, 'outlier')}): actual "
                         f"{float(s.loc[y]):,.0f} kbd, shown flattened to {cap:,.0f}")
            s.loc[y] = cap
    return s, "   ".join(notes)


def display_annual(df, value="cap_kbd"):
    """Annual summary with 2020/21 unplanned flattened for chart display."""
    s = annual_summary(df, value).copy()
    flU, note = flatten_outliers(s["Unplanned"])
    s["UnplDisp"] = flU
    s["TotDisp"] = s["Planned"] + s["UnplDisp"]
    return s, note


def pct_change(cur, prev, min_base=0.0, cap=5.0):
    """YoY %, suppressed (None) when the base is tiny or the result is absurd -
    avoids the +1000%/+2000% blow-ups from near-zero prior-year values."""
    try:
        cur, prev = float(cur), float(prev)
    except (TypeError, ValueError):
        return None
    if prev == 0 or abs(prev) < min_base:
        return None
    p = (cur - prev) / prev
    return None if abs(p) > cap else p


def h1_planned(df):
    """H1 (Jan-Jun) planned offline by year - the honest like-for-like view
    while 2027 H2 data is still incomplete."""
    mp = monthly_by_year(df, type_filter="PLANNED")
    h1 = MONTHS[:6]
    return {int(y): float(sum(mp.loc[y, m] for m in h1)) for y in mp.index}


def scenario_fan(df, window_key=DEFAULT_WINDOW, lo=0.8, hi=1.3):
    """Conservative / Average / Active monthly unplanned paths for 2027."""
    prof = baseline_profile(df, window_key)
    return {"Conservative": prof * lo, "Average": prof * 1.0, "Active": prof * hi}


def completed_unplanned(df, years=(2024, 2025, 2026, 2027)):
    """Monthly unplanned per year for charts, with the forecast filled in where
    the actuals are clearly incomplete (data still trickling in for the current
    partial year) instead of left at zero. A planned-only year (2027) is all
    forecast; a complete year keeps all 12 actuals. Returns
    {year: {'vals':[12], 'fc_from': first_forecast_month_idx}}."""
    mu = monthly_by_year(df, type_filter="UNPLANNED")
    fc = scenario_forecast(df)["monthly_unplanned"]
    out = {}
    for y in years:
        if y not in mu.index or y in PLANNED_ONLY_YEARS:
            out[y] = {"vals": [float(fc[m]) for m in MONTHS], "fc_from": 0}
            continue
        row = [float(mu.loc[y, m]) for m in MONTHS]
        cut = 11   # default: every month is a real actual (complete year)
        pos = sorted(v for v in row if v > 0)
        if y in PARTIAL_YEARS and pos:
            # walk back from December; the incomplete tail is the trailing run of
            # months sitting well below this year's own typical month (40% of the
            # median positive month). The last month at/above that line is real.
            thresh = 0.4 * pos[len(pos) // 2]
            cut = -1
            for i in range(11, -1, -1):
                if row[i] >= thresh:
                    cut = i
                    break
        for i in range(cut + 1, 12):
            row[i] = float(fc[MONTHS[i]])
        out[y] = {"vals": row, "fc_from": cut + 1}
    return out


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
        "scenario_padd": scenario_by_padd(df),
        "padd_share": padd_unplanned_share(df),
        "tornado": tornado(df),
        "clusters": consecutive_runs(df, min_len=4, exclude_years=OUTLIER_YEARS),
        "fcc_exxon": fcc_exxon_clusters(df),
        "fcc_grid": exxon_fcc_month_grid(df, [2022, 2023, 2024, 2025, 2026]),
        "padd_unpl_yoy": padd_yoy(df, "UNPLANNED"),
        "unit_share": unit_share(df),
        "monthly_yoy": monthly_yoy(df, "UNPLANNED"),
        "monthly_yoy_total": monthly_yoy(df, None),
        "range_band": monthly_range_band(df, "UNPLANNED"),
        "ta_schedule": {p: turnaround_schedule(df, 2026, padd=p) for p in PADD_ORDER},
        "ta_all": turnaround_schedule(df, 2026, top=60),
        "tidy_padd": tidy_monthly(df, "padd"),
        "tidy_unit": tidy_monthly(df, "unit_cat"),
        "naphtha": naphtha_analysis(df),
        "top_movers": top_movers(df),
        "scenario_bands": scenario_bands(df),
        "exxon_2027": exxon_2027_breakdown(df),
        "ta_2027": {p: turnaround_schedule(df, 2027, padd=p, top=8) for p in PADD_ORDER},
        "display_annual": display_annual(df),
        "h1_planned": h1_planned(df),
        "scenario_fan": scenario_fan(df),
        "completed_unplanned": completed_unplanned(df),
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
    from pathlib import Path
    default = str(Path(__file__).resolve().parent.parent / "data" / "Refinery_Outages_Data.xlsx")
    df = load(sys.argv[1] if len(sys.argv) > 1 else default)
    print(json.dumps(diagnostics(df), indent=2, default=str))
