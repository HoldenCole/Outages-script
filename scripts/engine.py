"""
engine.py
Reusable aggregation core for the refinery-outage analysis suite.

Turns an outage export -- the cleaned "Refinery Outages Enhanced" breakdown
(one row per event; the primary and only required input) or, optionally, a legacy
Snowflake 'Query1' export -- into clean, analysis-ready frames, pivots and a
single `build_context()` bundle. The schema is auto-detected. Every
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
from datetime import date
from pathlib import Path

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

# Naphtha balance assumptions: crude MAKES naphtha (the naphtha cut is ~35% of
# crude); a reformer's charge is essentially naphtha, so it CONSUMES ~1.0x its
# capacity. A CDU outage removes naphtha supply; a reformer outage removes naphtha
# demand. (Reformate, the gasoline/octane the reformer makes, is YIELD_FACTOR['Ref'].)
NAPHTHA_YIELD = 0.35            # naphtha produced per barrel of crude run (CDU)
REFORMER_NAPHTHA_INTAKE = 1.0   # naphtha consumed per barrel of reformer capacity
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

# The model + slides always cover 2023 .. current year + 1 -- a rolling window that
# advances itself each January (no manual bump), so a refreshed Snowflake just works.
# 2027 is a floor: the current curated outlook year, kept even if the clock is behind.
START_YEAR = 2023
END_YEAR = max(2027, date.today().year + 1)
FOCUS_YEAR = END_YEAR                  # the forward outlook year the deck + model headline
CURRENT_YEAR = FOCUS_YEAR - 1          # the in-progress year (H1 actual, H2 still coming) -- the
                                       # "rest of <yy>" naphtha/chem-feed deck headlines this one

# Vendored external market context for the "what it means for the market" read.
# Hand-updated from the EIA Weekly Petroleum Status Report; edit these when a fresh
# WPSR lands (keep AS_OF and the source in sync). Inventory figures are the % vs the
# five-year average for the report week. The summer-grade transition dates are the
# EPA RVP timeline (refiners switch to summer blend in Mar-Apr; retail June 1).
MARKET_CONTEXT = {
    "as_of": "week ending Jun 19, 2026 (EIA WPSR, released Jun 24, 2026)",
    "gasoline_vs_5yr_pct": -5,         # US motor gasoline stocks vs 5-yr avg
    "distillate_vs_5yr_pct": -10,      # distillate stocks vs 5-yr avg
    "crude_vs_5yr_pct": -7,            # commercial crude stocks vs 5-yr avg
    "crude_stocks_mmbbl": 412.1,
    "gasoline_demand_mbpd": 8.8,       # 4-wk avg product supplied
    "gasoline_demand_yoy_pct": -3.0,
    "summer_grade_window": (3, 6),     # refiners switch to summer blend Mar; retail deadline Jun 1
    "source": "https://www.eia.gov/petroleum/supply/weekly/",
}

# PADD connectivity -- the share of a CDU (crude) outage that CASCADES into lost
# downstream product. Where a refinery is well pipeline-connected (PADD 3 / Gulf),
# the units that run off the crude unit can keep going on piped-in intermediates,
# so a crude outage buffers (low pass-through). Islanded regions (PADD 2 / 4 / 5
# and parts of PADD 1) must cut the downstream units too (high pass-through).
# Tunable on the Assumptions sheet; these are the defaults (edit for your view).
PADD_CONNECTIVITY = {
    "PADD 1": 0.85, "PADD 2": 0.90, "PADD 3": 0.40, "PADD 4": 0.90, "PADD 5": 0.95,
}


# ----------------------------------------------------------------------------- data-quality audit
TURNAROUND_CYCLE_YEARS = 5             # focus-unit turnarounds run on ~a 5-year cycle


def cadence_flags(df_full, classes=("CDU", "FCC"), max_gap=4, reach_from=None):
    """Flag physical units that take a PLANNED outage again within < TURNAROUND_CYCLE
    years -- i.e. planned work in years <= max_gap apart, which is short vs the
    ~5-year turnaround cycle. UNPLANNED -> PLANNED is legitimate and never flagged
    (we only look at planned-to-planned). Uses the FULL history so prior turnarounds
    are visible. Returns a DataFrame sorted by the later year, biggest first.
    Reformers/hydrocrackers cycle faster, so the default classes are CDU + FCC."""
    reach_from = (FOCUS_YEAR - 1) if reach_from is None else reach_from
    pl = df_full[(df_full["type"] == "PLANNED") & (df_full["focus"].isin(classes))]
    rows = []
    for (plant, unit, focus), s in pl.groupby(["plant", "unit_name", "focus"]):
        by_year = s.groupby("year")["cap_kbd"].sum()
        yrs = sorted(int(y) for y in by_year.index)
        for a, b in zip(yrs, yrs[1:]):
            if b - a <= max_gap and b >= reach_from:
                rows.append({"plant": str(plant), "unit": str(unit), "focus": focus,
                             "prev_TA": a, "next_TA": b, "gap_yrs": b - a,
                             "kbd_prev": float(by_year.loc[a]), "kbd_next": float(by_year.loc[b])})
    cols = ["plant", "unit", "focus", "prev_TA", "next_TA", "gap_yrs", "kbd_prev", "kbd_next"]
    out = pd.DataFrame(rows, columns=cols)
    return out.sort_values(["next_TA", "kbd_next"], ascending=[False, False]).reset_index(drop=True)


def double_count_flags(df, tol=1.02):
    """Flag unit-months whose summed day-weighted offline exceeds the unit nameplate
    (> 100% offline is physically impossible -> overlapping/duplicate records). One
    row per (year, month, plant, unit), with the summed kbd, nameplate and ratio."""
    g = (df.groupby(["year", "month", "plant", "unit_name", "focus"])
         .agg(sum_kbd=("cap_kbd", "sum"), nameplate=("cap_raw", "max"), n_rows=("cap_kbd", "size"),
              types=("type", lambda s: "+".join(sorted(set(s))))).reset_index())
    g = g[(g["nameplate"] > 0) & (g["sum_kbd"] > g["nameplate"] * tol)].copy()
    g["ratio"] = g["sum_kbd"] / g["nameplate"]
    g["excess_kbd"] = g["sum_kbd"] - g["nameplate"]
    return g.sort_values(["year", "excess_kbd"], ascending=[False, False]).reset_index(drop=True)

# Years that are partial / special-cased (rendered grey-italic, footnoted).
PARTIAL_YEARS = [FOCUS_YEAR - 1, FOCUS_YEAR]   # current (partial actuals) + outlook (planned only)
PLANNED_ONLY_YEARS = [FOCUS_YEAR]      # no actual unplanned data exists for the outlook year
OUTLIER_YEARS = [2020, 2021]          # COVID / Winter Storm Uri -- excluded from baselines

# Forecast baseline windows offered in the scenario model.
BASELINE_WINDOWS = {
    "2023-2026": [2023, 2024, 2025, 2026],
    "2022-2025": [2022, 2023, 2024, 2025],
    "2023-2025": [2023, 2024, 2025],
    "2018-19 & 22-25": [2018, 2019, 2022, 2023, 2024, 2025],
    "All ex-2020/21": [2014, 2015, 2016, 2017, 2018, 2019,
                       2022, 2023, 2024, 2025],
}
DEFAULT_WINDOW = "2023-2026"   # complete 2023-25 plus the latest reported 2026 months (completeness-aware)

# Focus units, in the priority order the desk reads them (CDU first, then FCC,
# hydrocracker, reformer). These four are reported per-unit; everything else is
# context. Hydrocracker = HYDROCRACKING only (NOT hydrotreating, a different unit).
FOCUS_ORDER = ["CDU", "FCC", "Hydrocracker", "Reformer"]
FOCUS_LABEL = {
    "CDU": "Crude (CDU)",
    "FCC": "FCC (cat cracker)",
    "Hydrocracker": "Hydrocracker",
    "Reformer": "Reformer",
}
UNITCAT_TO_FOCUS = {
    "ATMOS DISTILLATION": "CDU",          # atmospheric crude ONLY -- vacuum (VDU) is not folded in
    "FLUID CAT CRACKING": "FCC",
    "HYDROCRACKING": "Hydrocracker", "RESID_HYDROCRACKING": "Hydrocracker",
    "REFORMING": "Reformer",
}


def _focus_from_unitcat(unit_cat, unit_name):
    """Map UNIT_CATEGORY -> focus class, then DEMOTE the rows the golden record
    mislabels: a vacuum pipe still (VPS) tagged ATMOS DISTILLATION is not CDU
    (CDU is atmospheric crude only -- vacuum is never folded in). A unit is treated
    as vacuum when its name contains 'VPS' or starts with 'VACUUM'. This rule is
    mirrored exactly by the Excel _focus_formula (build_workbook.py) so the live,
    paste-to-refresh model agrees with the engine and keeps working as data grows."""
    focus = unit_cat.map(UNITCAT_TO_FOCUS)
    nm = unit_name.astype(str).str.upper().str.strip()
    is_vac = nm.str.contains("VPS", regex=False, na=False) | nm.str.startswith("VACUUM")
    return focus.mask(unit_cat.eq("ATMOS DISTILLATION") & is_vac)


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


def _pick_sheet(xls):
    """Choose the worksheet that looks like the outage export: prefer RAW_SHEET,
    else the sheet whose header best matches the expected columns."""
    names = xls.sheet_names
    if RAW_SHEET in names:
        return RAW_SHEET
    want = set(COLMAP.values())
    best, best_score = names[0], -1
    for nm in names:
        try:
            cols = set(pd.read_excel(xls, sheet_name=nm, nrows=0).columns)
        except Exception:
            continue
        score = len(want & cols)
        if score > best_score:
            best, best_score = nm, score
    return best


def _read_table(path):
    """Read the raw export into a dataframe. Accepts .xlsx/.xls (auto-picks the
    sheet) or .csv, with clear errors so others can run this on their own data."""
    path = str(path).strip()
    if not Path(path).exists():
        raise FileNotFoundError(
            f"Input file not found: {path}\n"
            "Pass the path to your outage export, e.g.:\n"
            "    python scripts/build_all.py data/MyExport.xlsx")
    if path.lower().endswith((".csv", ".txt")):
        return pd.read_csv(path)
    xls = pd.ExcelFile(path)
    return pd.read_excel(xls, sheet_name=_pick_sheet(xls))


# Source columns the analysis genuinely needs; everything else degrades gracefully.
ESSENTIAL_SRC = ["CAP_OFFLINE_ADJUSTED_KBD", "OUTAGE_YEAR", "OUTAGE_MONTH"]


# --------------------------------------------------------------- "Enhanced" breakdown schema
def _detect_schema(cols):
    """Snowflake export (OUTAGE_YEAR/CAP_OFFLINE_ADJUSTED_KBD) vs the cleaned
    'Refinery Outages Enhanced' breakdown (Unit Name / Start Date / Offline
    Capacity (KBD), one row per outage event)."""
    cols = {str(c).strip() for c in cols}
    if "OUTAGE_YEAR" in cols or "CAP_OFFLINE_ADJUSTED_KBD" in cols:
        return "snowflake"
    if {"Unit Name", "Outage Type", "Start Date"} <= cols or any("Offline Capacity" in c for c in cols):
        return "enhanced"
    return "snowflake"


def _unit_cat_from_name(name):
    """Map an Enhanced free-text Unit Name to a canonical UNIT_CATEGORY so the same
    focus/bucket maps work. Order matters: hydro-words are checked before the
    plain crude/FCC keywords (e.g. 'Selective Hydrogenation (FCCU)' is a
    hydrotreater; 'Reformer Feed Hydrotreater' is a hydrotreater, not a reformer)."""
    s = str(name).upper()
    H = lambda *k: any(x in s for x in k)
    if H("VACUUM", "VPS", "VDU"):
        return "VACUUM DISTILLATION"
    if H("HYDROCRACK", "ULTRACRACK", "ISOCRACK", "ISOMAX", "HEAVY OIL CRACK") or "(HCU)" in s or s.startswith("HCU"):
        return "HYDROCRACKING"
    if (H("HYDROTREAT", "HYDROFIN", "HYDROGENATION", "DESULF", "UNIFINER", "GULFINER", "GOFINER",
          "UNIBON", "HDS", "HDT", "DHT", " HT ", "HTU", "CHD", "ULSD", "NHT", "CGH", "CFH", "GFH",
          "CAT FEED", "GAS OIL HYDRO", "DIESEL HT") or s.startswith("HT ")):
        return "HYDROTREATING"
    if H("FCC", "CAT CRACK", "CATALYTIC CRACK", "FLUID CAT", "CCU"):
        return "FLUID CAT CRACKING"
    if H("COKER", "COKING", "FLEXICOK"):
        return "COKING"
    if H("REFORMER", "PLATFORMER", "ULTRAFORMER", "POWERFORMER", "REFORMING", "REFINING UNIT") \
            or "CRU " in s or "CRU(" in s:
        return "REFORMING"
    if H("CRUDE", "PIPESTILL", "PIPE STILL", "PSLA", "CTU", "ACU", "AVU", "TOPPING", "DISTILLING",
         "ATMOS", "COMBO", "CDU", "DU-", "CONDENSATE SPLIT"):
        return "ATMOS DISTILLATION"
    if H("ALKYLATION", "ALKY"):
        return "ALKYLATION"
    if H("ISOMER", "PENEX", "PENTANE"):
        return "ISOMERIZATION"
    if H("MTBE"):
        return "MTBE"
    if H("NAPHTHA"):
        return "HYDROTREATING"
    if H("DISTILLATION", "LIGHT ENDS"):
        return "ATMOS DISTILLATION"
    return "OTHER"


ENHANCED_MAX_SPAN_DAYS = 500     # drop obvious placeholder/error spans (e.g. an outage 'ending' 2031)
ENHANCED_MIN_YEAR = 2023         # only 2023+ is verified data -- 2021/2022 are dropped entirely

# Confirmed capacity overrides: cap a unit's offline where the site always keeps
# part of a multi-train unit online (match plant + exact unit name).
ENHANCED_CAP = [
    # Garyville runs two crude trains under 'Crude 210' and always leaves one
    # online, so no more than ~one train (138 kbd) is ever really offline.
    {"plant": "GARYVILLE", "unit": "CRUDE 210", "max_kbd": 138.0},
]


def _capped(plant, unit, cap):
    p, u = str(plant).upper(), str(unit).strip().upper()
    for c in ENHANCED_CAP:
        if c["plant"] in p and u == c["unit"]:
            return min(cap, c["max_kbd"])
    return cap


def _load_enhanced(raw):
    """Load the cleaned 'Refinery Outages Enhanced' breakdown (one row per outage
    event, with start/end dates and nameplate capacity) into the same long-form
    frame the rest of the engine expects -- expanding each event to the calendar
    months it spans (cap_kbd = nameplate x days-down / days-in-month, cap_raw =
    full nameplate). Source is kept so callers can filter by data provider."""
    raw = raw.rename(columns=lambda c: str(c).strip()).drop_duplicates()
    capcol = next((c for c in raw.columns if "Offline Capacity" in c), None)
    rows = []
    for i, r in raw.iterrows():
        s = pd.to_datetime(r.get("Start Date"), errors="coerce")
        e = pd.to_datetime(r.get("End Date"), errors="coerce")
        if pd.isna(s):
            continue
        if pd.isna(e) or e < s:
            e = s
        if (e - s).days > ENHANCED_MAX_SPAN_DAYS and e.year >= 2029:
            continue                       # run-off placeholder / data error (e.g. Anacortes -> 2031)
        cap = pd.to_numeric(r.get(capcol), errors="coerce")
        cap = float(cap) if pd.notna(cap) else 0.0
        otype = str(r.get("Outage Type", "")).strip().upper()
        typ = "PLANNED" if otype == "PLANNED" else "UNPLANNED"   # UNKNOWN folds into UNPLANNED
        uname = str(r.get("Unit Name", "")).strip()
        cap = _capped(r.get("Plant"), uname, cap)   # multi-train cap (see ENHANCED_CAP)
        region = r.get("Country/Region")
        oid = f"ENH{i}"
        for per in pd.period_range(s.to_period("M"), e.to_period("M"), freq="M"):
            mstart, mend = max(s, per.start_time), min(e, per.end_time)
            days_down = (mend.normalize() - mstart.normalize()).days + 1
            frac = min(1.0, max(0.0, days_down / per.days_in_month))
            rows.append({
                "year": int(per.year), "month": int(per.month),
                "month_date": per.start_time, "start": s, "end": e,
                "cap_raw": cap, "cap_kbd": cap * frac, "unit_cap": np.nan,
                "duration": int((e - s).days), "operator": str(r.get("Owner", "")).strip(),
                "plant": str(r.get("Plant", "")).strip(), "otype": typ, "otype2": np.nan,
                "unit_name": uname, "unit_cat": _unit_cat_from_name(uname),
                "pad_dist": region, "padd": _to_padd_from_roman(region),
                "outage_id": oid, "source": str(r.get("Source", "")).strip(),
                "country": np.nan, "state": np.nan, "city": np.nan, "cause": np.nan,
            })
    out = pd.DataFrame(rows)
    if out.empty:
        raise ValueError("The Enhanced export produced no usable rows (check Start Date / Unit Name).")
    out = out[out["year"] >= ENHANCED_MIN_YEAR].copy()     # 2021/2022 not verified -> drop entirely
    out["year"] = out["year"].astype("Int64")
    out["month"] = out["month"].astype("Int64")
    out["type"] = out["otype"]
    out["padd_source"] = np.where(out["padd"].notna(), "PAD_DIST", "UNRESOLVED")
    out["bucket"] = out["unit_cat"].map(UNITCAT_TO_BUCKET).fillna("Other")
    out["focus"] = _focus_from_unitcat(out["unit_cat"], out["unit_name"])
    out["is_exxon"] = out["operator"].astype(str).str.upper().str.contains("EXXON", na=False)
    out["mogas_kbd"] = out["cap_kbd"] * out["bucket"].map(YIELD_FACTOR).fillna(0.0)
    out["month_name"] = out["month"].map(
        lambda m: MONTHS[int(m) - 1] if pd.notna(m) and 1 <= int(m) <= 12 else None)
    out["schema"] = "enhanced"
    return out


def load(path):
    """Load the raw export, return a cleaned long-form dataframe.

    Robust to other people's files: the sheet is auto-selected (preferring
    'Query1'), .csv is accepted, and a clear error is raised if the essential
    columns are absent rather than silently producing an empty analysis.
    """
    df = _read_table(path)
    if _detect_schema(df.columns) == "enhanced":
        return _load_enhanced(df)
    missing = [c for c in ESSENTIAL_SRC if c not in df.columns]
    if missing:
        shown = ", ".join(map(str, list(df.columns)[:25])) + (" ..." if len(df.columns) > 25 else "")
        raise ValueError(
            "The input is missing required column(s): " + ", ".join(missing) + ".\n"
            f"Found columns: {shown}\n"
            "This tool expects a refinery-outage export with at least "
            "CAP_OFFLINE_ADJUSTED_KBD, OUTAGE_YEAR and OUTAGE_MONTH "
            "(see README -> 'Run it on your own data' for the full schema).")
    out = pd.DataFrame()
    for logical, src in COLMAP.items():
        out[logical] = df[src] if src in df.columns else np.nan

    out["year"] = pd.to_numeric(out["year"], errors="coerce").astype("Int64")
    out["month"] = pd.to_numeric(out["month"], errors="coerce").astype("Int64")
    out["cap_kbd"] = pd.to_numeric(out["cap_kbd"], errors="coerce").fillna(0.0)
    out["cap_raw"] = pd.to_numeric(out["cap_raw"], errors="coerce").fillna(0.0)  # full unit nameplate offline
    out["unit_cap"] = pd.to_numeric(out["unit_cap"], errors="coerce")
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
    out["focus"] = _focus_from_unitcat(out["unit_cat"], out["unit_name"])   # NaN for non-focus / vacuum
    out["is_exxon"] = (out["operator"].astype(str).str.upper()
                       .str.contains("EXXON", na=False))
    out["mogas_kbd"] = out["cap_kbd"] * out["bucket"].map(YIELD_FACTOR).fillna(0.0)

    out["month_name"] = out["month"].map(
        lambda m: MONTHS[int(m) - 1] if pd.notna(m) and 1 <= int(m) <= 12 else None)
    out["schema"] = "snowflake"
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
        "2027_types": df.loc[df.year.eq(FOCUS_YEAR), "type"].value_counts().to_dict(),
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

    Completeness-aware: the most recent year in the window is usually partial
    (reported only through some month), so the months it has not reported yet are
    left out of those months' average instead of being counted as zero. A fresh
    partial year (e.g. the current year's H1) therefore sharpens the early-month
    baseline without dragging the unreported later months toward zero. Backbone of
    the 2027 scenario: the mean calendar-month shape of unplanned offline over the
    chosen window.
    """
    sub = df[df["type"].eq(type_filter) & df["year"].isin(years)]
    if padd:
        sub = sub[sub["padd"].eq(padd)]
    by_ym = (sub.groupby(["year", "month"])[value].sum()
             .unstack("month").reindex(columns=range(1, 13)))
    # average over the window years actually present in the data (so a dataset
    # that only reaches back to 2024 isn't diluted by dividing over empty years)
    present = [y for y in years if y in by_ym.index]
    if not present:
        prof = pd.Series(0.0, index=range(1, 13))
        prof.index = MONTHS
        return prof
    by_ym = by_ym.reindex(present).fillna(0.0)
    cy, cm = latest_actual_month(df) or (max(present), 12)
    for y in present:                       # blank a partial current year's unreported tail
        if y >= cy and cm < 12:
            by_ym.loc[y, [m for m in range(cm + 1, 13)]] = np.nan
    prof = by_ym.mean(axis=0, skipna=True).reindex(range(1, 13)).fillna(0.0)
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
    planned_2027 = float(annual_summary(df).loc[FOCUS_YEAR, "Planned"]) \
        if FOCUS_YEAR in annual_summary(df).index else 0.0
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
    years = _complete_years(df, years)        # a partial current year would floor the band at 0
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
    around the point forecast. Fully-reported years only (a partial current year
    would understate the annual total)."""
    years = _complete_years(df, BASELINE_WINDOWS[window_key])
    annuals = [df[df["type"].eq("UNPLANNED") & df["year"].eq(y)]["cap_kbd"].sum() for y in years]
    annuals = [a for a in annuals if a]
    if not annuals:
        return {"p25": 0.0, "p50": 0.0, "p90": 0.0, "mean": 0.0}
    return {"p25": float(np.percentile(annuals, 25)), "p50": float(np.percentile(annuals, 50)),
            "p90": float(np.percentile(annuals, 90)), "mean": float(np.mean(annuals))}


def exxon_2027_breakdown(df, operator_contains="EXXON", year=FOCUS_YEAR):
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


def completed_unplanned(df, years=(FOCUS_YEAR - 3, FOCUS_YEAR - 2, FOCUS_YEAR - 1, FOCUS_YEAR)):
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


# ----------------------------------------------------------------------------- market context ($ at risk)
MARKET_CRACK_PATH = Path(__file__).resolve().parent.parent / "data" / "market_crack.csv"
DAYS_IN_MONTH = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]


def load_crack(path=None):
    """Monthly gasoline crack ($/bbl) from data/market_crack.csv ->
    {(year, month): crack}. Returns {} if the file is absent so the rest of the
    pipeline degrades gracefully (the Margin Context sheet just stays a blank,
    Bloomberg-fillable template). Refresh it with scripts/fetch_market_data.py."""
    p = Path(path) if path else MARKET_CRACK_PATH
    if not p.exists():
        return {}
    df = pd.read_csv(p, comment="#")
    return {(int(r.year), int(r.month)): float(r.crack) for r in df.itertuples()}


def crack_matrix(crack, years):
    """year -> [12] monthly crack, filling any missing month with that year's own
    mean (or the all-history mean) so charts/formulas never see gaps."""
    allv = list(crack.values())
    gmean = float(np.mean(allv)) if allv else 0.0
    out = {}
    for y in years:
        yvals = [crack.get((y, m)) for m in range(1, 13)]
        present = [v for v in yvals if v is not None]
        ymean = float(np.mean(present)) if present else gmean
        out[y] = [v if v is not None else ymean for v in yvals]
    return out


def outage_dollar_impact(df, crack, years=range(2018, 2028)):
    """Gross refining-margin **at risk** from offline capacity, valued at the
    gasoline crack: $MM[m] = offline_kbd[m] * days[m] * crack[m] / 1000.

    A desk shorthand - it values offline capacity at the gasoline gross margin;
    unplanned = unexpected supply loss (the bullish/at-risk number), planned =
    margin deferred via scheduled work. Returns
    {year: {planned, unplanned, total, crack_avg}} in $MM, plus 'monthly_unpl'
    per year for charting."""
    if not crack:
        return {}
    mp = monthly_by_year(df, type_filter="PLANNED")
    mu = monthly_by_year(df, type_filter="UNPLANNED")
    yrs = [int(y) for y in years if any((y, m) in crack for m in range(1, 13))]
    cm = crack_matrix(crack, yrs)
    out = {}
    for y in yrs:
        cr = cm[y]

        def dollars(mat):
            if y not in mat.index:
                return [0.0] * 12
            return [float(mat.loc[y, MONTHS[m]] * DAYS_IN_MONTH[m] * cr[m] / 1000.0)
                    for m in range(12)]
        pl, un = dollars(mp), dollars(mu)
        out[y] = {"planned": sum(pl), "unplanned": sum(un), "total": sum(pl) + sum(un),
                  "crack_avg": float(np.mean(cr)), "monthly_unpl": un}
    return out


def crack_outage_relationship(df, crack, y0=2018, y1=2025):
    """Test whether outages actually track the gasoline crack, over full years
    [y0, y1]. The honest result on this book: planned r~0 (turnarounds are
    operational/seasonal, NOT margin-timed) and unplanned only weakly negative
    (driven by the 2020-21 demand/freeze extremes). Conclusion the desk should
    take: the crack is a *valuation lens* (what outages are worth), not a
    predictor of when they happen. Returns {} when no crack data is present."""
    if not crack:
        return {}
    mp = monthly_by_year(df, type_filter="PLANNED")
    mu = monthly_by_year(df, type_filter="UNPLANNED")
    cr, pl, un = [], [], []
    for y in range(y0, y1 + 1):
        for mi, m in enumerate(MONTHS):
            c = crack.get((y, mi + 1))
            if c is None:
                continue
            cr.append(c)
            pl.append(float(mp.loc[y, m]) if y in mp.index else 0.0)
            un.append(float(mu.loc[y, m]) if y in mu.index else 0.0)
    if len(cr) < 12:
        return {}
    cr, pl, un = np.array(cr), np.array(pl), np.array(un)

    def corr(a, b):
        return float(np.corrcoef(a, b)[0, 1]) if a.std() > 0 and b.std() > 0 else 0.0
    return {"n": len(cr), "planned_r": corr(cr, pl), "unplanned_r": corr(cr, un),
            "total_r": corr(cr, pl + un)}


# ----------------------------------------------------------------------------- period-over-period
def latest_actual_month(df, min_frac=0.4):
    """Most recent *substantially-reported* (year, month).

    Anchoring naively on the very last month is wrong when a refresh's tail is
    still filling in (unplanned reporting trails off, or the month is a
    planned-only future book): that would read as a fake month-on-month cliff.
    So we take the last month whose UNPLANNED offline is at least `min_frac` of
    the median monthly level - which lands on the last genuinely-reported month
    and self-adjusts to whatever month a real refresh ends on. Returns
    (year, month) or None.
    """
    base = df.dropna(subset=["year", "month"])
    un = base[base["type"] == "UNPLANNED"].groupby(["year", "month"])["cap_kbd"].sum()
    if un.empty or un[un > 0].empty:
        un = base.groupby(["year", "month"])["cap_kbd"].sum()      # fall back to total
    pos = un[un > 0]
    if pos.empty:
        return None
    good = un[un >= min_frac * float(pos.median())]
    if good.empty:
        return None
    y, m = max(good.index)
    return int(y), int(m)


def _complete_years(df, years):
    """Subset of `years` that are fully reported through December. Drops a partial
    current year (whose unreported tail would distort annual or range statistics);
    leaves complete historical years untouched."""
    cy, cm = latest_actual_month(df) or ((max(years) if len(years) else 0), 12)
    return [y for y in years if (y < cy) or (y == cy and cm >= 12)]


def period_change(df):
    """Month-over-month change: the latest actual month vs the prior calendar
    month. Returns headline deltas (total / planned / unplanned), the biggest
    movers by PADD / unit / operator, and the outages that newly appeared or
    dropped off. None when there aren't two comparable months.

    The deck's "what changed" section is built entirely from this, so it adapts
    to whatever month a refreshed export lands on - no hard-coded dates.
    """
    cur = latest_actual_month(df)
    if cur is None:
        return None
    cy, cm = cur
    py, pm = (cy, cm - 1) if cm > 1 else (cy - 1, 12)

    def sub(y, m):
        return df[(df["year"] == y) & (df["month"] == m)]
    cur_df, prv_df = sub(cy, cm), sub(py, pm)
    if cur_df["cap_kbd"].sum() == 0 and prv_df["cap_kbd"].sum() == 0:
        return None

    def tot(d, t=None):
        return float((d if t is None else d[d["type"] == t])["cap_kbd"].sum())

    def movers(col):
        c = cur_df.groupby(col)["cap_kbd"].sum()
        p = prv_df.groupby(col)["cap_kbd"].sum()
        idx = c.index.union(p.index)
        delta = (c.reindex(idx, fill_value=0.0) - p.reindex(idx, fill_value=0.0))
        return delta.sort_values(ascending=False)            # +ve = month-on-month increase

    def ids(d):
        d = d.dropna(subset=["outage_id"])
        if d.empty:
            return pd.DataFrame(columns=["kbd", "plant", "unit", "padd", "type"])
        return d.groupby("outage_id").agg(
            kbd=("cap_kbd", "sum"), plant=("plant", "first"),
            unit=("unit_cat", "first"), padd=("padd", "first"), type=("type", "first"))
    ci, pi = ids(cur_df), ids(prv_df)
    new = ci.loc[ci.index.difference(pi.index)].sort_values("kbd", ascending=False)
    gone = pi.loc[pi.index.difference(ci.index)].sort_values("kbd", ascending=False)

    # trailing 13-month context up to (and including) the anchor month
    seq, yy, mm = [], cy, cm
    for _ in range(13):
        seq.append((yy, mm))
        mm, yy = (mm - 1, yy) if mm > 1 else (12, yy - 1)
    trail = []
    for (yy, mm) in reversed(seq):
        s = sub(yy, mm)
        trail.append({"label": f"{MONTHS[mm - 1]} {str(yy)[2:]}",
                      "total": float(s["cap_kbd"].sum()),
                      "unplanned": float(s[s["type"] == "UNPLANNED"]["cap_kbd"].sum())})

    return {
        "cur": (cy, cm), "prev": (py, pm), "trail": trail,
        "cur_label": f"{MONTHS[cm - 1]} {cy}", "prev_label": f"{MONTHS[pm - 1]} {py}",
        "total": (tot(cur_df), tot(prv_df)),
        "unplanned": (tot(cur_df, "UNPLANNED"), tot(prv_df, "UNPLANNED")),
        "planned": (tot(cur_df, "PLANNED"), tot(prv_df, "PLANNED")),
        "events": (int(cur_df["outage_id"].nunique()), int(prv_df["outage_id"].nunique())),
        "by_padd": movers("padd"), "by_unit": movers("unit_cat"), "by_operator": movers("operator"),
        "by_focus": movers("focus"),
        "new": new, "gone": gone,
    }


SNAPSHOT_LOG = Path(__file__).resolve().parent.parent / "data" / "whatschanged_log.csv"


def update_snapshot_log(df, path=None, asof=None, n_fwd=6):
    """The 'What's Changed' week-over-week engine. The source is monthly (and its
    start/end dates are unreliable wide envelopes), so true week-over-week is done
    by comparing weekly PULLS, not intra-data weeks: each build appends a dated row
    of the total day-weighted offline for the current month and the next n_fwd
    calendar months. Re-running the same day overwrites that day's row; running
    weekly accumulates the rolling history. Returns the log (oldest->newest)."""
    path = Path(SNAPSHOT_LOG if path is None else path)
    asof = date.today() if asof is None else asof
    y, m = asof.year, asof.month
    row = {"as_of": asof.isoformat()}
    for k in range(n_fwd):
        yy, mm = y, m + k
        while mm > 12:
            mm -= 12; yy += 1
        row[f"{yy}-{mm:02d}"] = round(float(df[(df["year"] == yy) & (df["month"] == mm)]["cap_kbd"].sum()), 1)
    if path.exists():
        log = pd.read_csv(path)
        log = log[log["as_of"] != row["as_of"]]
    else:
        log = pd.DataFrame()
    log = pd.concat([log, pd.DataFrame([row])], ignore_index=True)
    path.parent.mkdir(parents=True, exist_ok=True)
    log.to_csv(path, index=False)
    return log.sort_values("as_of").reset_index(drop=True)


# ----------------------------------------------------------------------------- per-unit capacity offline
def unit_offline_monthly(df, focus=None, type_filter=None, padd=None,
                         y0=START_YEAR, y1=END_YEAR, value="cap_kbd"):
    """Concurrent capacity offline (kbd) by month for one focus unit class.

    The honest 'per-unit' read: for each (year, month) we take each *distinct
    physical unit's* offline capacity once (deduped on plant+unit), so a unit is
    never double-counted across months and we never add a CDU to an FCC. Returns
    a year x month matrix (Jan..Dec).

    Day-weighted (`cap_kbd`) by default: a unit offline only part of a month is
    credited only for the days it is actually down (nameplate x days-down /
    days-in-month) -- which equals the average daily concurrent offline that
    month, and never carries a unit's capacity into a month it is back online.
    Pass value="cap_raw" for the full-nameplate 'peak in-month' read instead."""
    sub = df[(df["year"] >= y0) & (df["year"] <= y1)].dropna(subset=["month"]).copy()
    if focus:
        sub = sub[sub["focus"].eq(focus)]
    if type_filter:
        sub = sub[sub["type"].eq(type_filter)]
    if padd:
        sub = sub[sub["padd"].eq(padd)]
    if sub.empty:
        return pd.DataFrame(0.0, index=list(range(y0, y1 + 1)), columns=MONTHS)
    # sum the day-weighted offline (CAP_OFFLINE_ADJUSTED_KBD) by month -- this is
    # exactly what a live Excel SUMIFS does, so the deck and the model agree and
    # both keep working as the Snowflake grows.
    g = (sub.groupby(["year", "month"])[value].sum()
           .unstack("month").reindex(columns=range(1, 13)))
    g.columns = MONTHS
    return g.reindex(range(y0, y1 + 1)).fillna(0.0)


def focus_unit_monthly(df, type_filter=None, **kw):
    """{focus_class: year x month concurrent-offline matrix} for the four focus units."""
    return {f: unit_offline_monthly(df, focus=f, type_filter=type_filter, **kw)
            for f in FOCUS_ORDER}


def focus_unit_padd_month(df, focus, year, value="cap_kbd"):
    """PADD x month concurrent offline (day-weighted kbd) for one focus unit in
    one year -- the 'timeline by month and PADD' view (each unit kept separate)."""
    sub = df[(df["year"] == year) & df["focus"].eq(focus)
             & df["padd"].isin(PADD_ORDER)].dropna(subset=["month"])
    if sub.empty:
        return pd.DataFrame(0.0, index=PADD_ORDER, columns=MONTHS)
    g = (sub.groupby(["padd", "month"])[value].sum()
           .unstack("month").reindex(index=PADD_ORDER, columns=range(1, 13)))
    g.columns = MONTHS
    return g.fillna(0.0)


def focus_annual_peak(df, type_filter=None, y0=START_YEAR, y1=END_YEAR):
    """year x focus-class table of the busiest month's concurrent offline (kbd,
    day-weighted): 'at its worst month this year, how much of each unit class was
    offline on an average day'. Summing a unit's months would double-count; the
    single worst month is the honest scalar to compare across years."""
    out = {}
    for f in FOCUS_ORDER:
        m = unit_offline_monthly(df, focus=f, type_filter=type_filter, y0=y0, y1=y1)
        out[f] = m.max(axis=1)
    return pd.DataFrame(out).reindex(range(y0, y1 + 1))


def h1_focus_planned(df, years=(FOCUS_YEAR - 2, FOCUS_YEAR - 1, FOCUS_YEAR)):
    """H1 (Jan-Jun) PLANNED offline per focus unit, day-weighted average kbd, by
    year. The like-for-like window: 2027 is booked only through H1, so comparing
    the Jan-Jun planned slate year-on-year avoids 2027's still-incomplete H2.
    Returns a DataFrame (rows = focus units, columns = years) of the average kbd
    offline across Jan-Jun -- each unit counted once per month, never summed
    across unit classes."""
    fp = focus_unit_monthly(df, type_filter="PLANNED")
    h1 = MONTHS[:6]
    data = {}
    for f in FOCUS_ORDER:
        m = fp[f]
        data[f] = {int(y): (float(np.mean([m.loc[y, mo] for mo in h1])) if y in m.index else 0.0)
                   for y in years}
    return pd.DataFrame(data).T.reindex(FOCUS_ORDER)


def naphtha_balance(df, year=FOCUS_YEAR, padd=None):
    """CDU vs reformer outages read as a naphtha supply/demand balance.

    Crude distillation MAKES naphtha (~35% of crude); reformers CONSUME it (their
    charge is naphtha). So a CDU outage removes naphtha SUPPLY and a reformer
    outage removes naphtha DEMAND. The net per month says whether outages leave
    naphtha long (surplus) or short (deficit):

        supply_removed = CDU offline  x NAPHTHA_YIELD            (~0.35)
        demand_removed = Reformer offline x REFORMER_NAPHTHA_INTAKE (~1.0)
        net            = demand_removed - supply_removed   (+ surplus / - deficit)

    Day-weighted concurrent offline, each unit once per month. Also returns the
    reformate (gasoline/octane) the offline reformers would have made, valued at
    YIELD_FACTOR['Ref'], for the octane read. Returns a dict of 12-month lists
    plus summary scalars."""
    cdu = unit_offline_monthly(df, focus="CDU", padd=padd)
    ref = unit_offline_monthly(df, focus="Reformer", padd=padd)
    cdu_m = [float(cdu.loc[year, m]) if year in cdu.index else 0.0 for m in MONTHS]
    ref_m = [float(ref.loc[year, m]) if year in ref.index else 0.0 for m in MONTHS]
    supply = [c * NAPHTHA_YIELD for c in cdu_m]
    demand = [r * REFORMER_NAPHTHA_INTAKE for r in ref_m]
    net = [d - s for d, s in zip(demand, supply)]
    reformate = [r * YIELD_FACTOR["Ref"] for r in ref_m]      # gasoline/octane lost when reformer is down
    return {
        "year": year, "months": MONTHS,
        "cdu_offline": cdu_m, "ref_offline": ref_m,
        "supply_removed": supply, "demand_removed": demand, "net": net,
        "reformate_lost": reformate,
        "annual_net": float(sum(net)),
        "n_deficit": int(sum(1 for v in net if v < -1e-6)),
        "n_surplus": int(sum(1 for v in net if v > 1e-6)),
        "naphtha_yield": NAPHTHA_YIELD, "reformer_intake": REFORMER_NAPHTHA_INTAKE,
        "padd": padd,
    }


H1_MONTHS = [1, 2, 3, 4, 5, 6]


def focus_2027_split(df, focus, value="cap_kbd", year=FOCUS_YEAR):
    """The 2027 data-completeness split for one focus unit (day-weighted kbd).

    We have ExxonMobil's full-year 2027 plan (verified against their corporate
    schedule), but for every other operator only H1-2027 (Jan-Jun) is actually
    booked -- H2 is still being scheduled. So each month's concurrent offline is
    split into:
        confirmed  = Exxon (any month)  +  non-Exxon in H1
        indicative = non-Exxon in H2 (incomplete, a floor that fills in)
    Returns {'confirmed': [12], 'indicative': [12]} (kbd)."""
    sub = df[(df["year"] == year) & df["focus"].eq(focus)].dropna(subset=["month"]).copy()
    conf = sub[sub["is_exxon"] | sub["month"].isin(H1_MONTHS)]
    indic = sub[(~sub["is_exxon"]) & (~sub["month"].isin(H1_MONTHS))]

    def monthly(d):
        if d.empty:
            return [0.0] * 12
        s = d.groupby("month")[value].sum().reindex(range(1, 13)).fillna(0.0)
        return [float(s[m]) for m in range(1, 13)]
    return {"confirmed": monthly(conf), "indicative": monthly(indic)}


def unit_events(df, focus=None, operator_contains=None, year=None, padd=None,
                type_filter=None, top=None, min_kbd=0.0):
    """Per-unit event timeline: one row per physical outage (deduped to its peak
    nameplate, full date span) -- the 'per unit, not summed' detail behind the
    monthly view. Columns: plant, operator, padd, focus, unit_cat, unit_name,
    kbd, start, end, months, span, type, year. Sorted by offline kbd."""
    cols = ["outage_id", "plant", "operator", "padd", "focus", "unit_cat", "unit_name",
            "kbd", "start", "end", "months", "span", "type", "year"]
    sub = df.copy()
    if focus:
        sub = sub[sub["focus"].eq(focus)]
    if operator_contains:
        sub = sub[sub["operator"].astype(str).str.upper()
                  .str.contains(operator_contains.upper(), na=False)]
    if year is not None:
        sub = sub[sub["year"].eq(year)]
    if padd:
        sub = sub[sub["padd"].eq(padd)]
    if type_filter:
        sub = sub[sub["type"].eq(type_filter)]
    if min_kbd:
        sub = sub[sub["cap_raw"] >= min_kbd]
    if sub.empty:
        return pd.DataFrame(columns=cols)
    rows = []
    for (oid, plant, unit), g in sub.groupby(["outage_id", "plant", "unit_name"], dropna=False):
        months = sorted(int(m) for m in g["month"].dropna().unique())
        rows.append({
            "outage_id": str(oid),
            "plant": str(plant), "operator": str(g["operator"].iloc[0]),
            "padd": str(g["padd"].iloc[0]),
            "focus": (g["focus"].dropna().iloc[0] if g["focus"].notna().any() else None),
            "unit_cat": str(g["unit_cat"].iloc[0]), "unit_name": str(unit),
            "kbd": float(g["cap_raw"].max()),
            "start": g["start"].min(), "end": g["end"].max(),
            "months": months, "span": _span(months) if months else "",
            "type": str(g["type"].iloc[0]),
            "year": int(g["year"].dropna().iloc[0]) if g["year"].notna().any() else None,
        })
    out = pd.DataFrame(rows, columns=cols).sort_values("kbd", ascending=False)
    return out.head(top) if top else out.reset_index(drop=True)


# ----------------------------------------------------------------------- Exxon plan verification
EXXON_PLAN_PATH = Path(__file__).resolve().parent.parent / "data" / "exxon_ta_plan.csv"


def load_exxon_plan(path=None):
    """ExxonMobil's own corporate turnaround plan (vendored from the AMR schedule
    into data/exxon_ta_plan.csv) -> DataFrame[site, unit, platform, subunit,
    bucket, year, start, end, event, region]. Used to verify the IIR Exxon
    records per-unit. Empty frame if the file is absent (verification just
    degrades to 'unverified')."""
    p = Path(path) if path else EXXON_PLAN_PATH
    if not p.exists():
        return pd.DataFrame()
    plan = pd.read_csv(p, comment="#")
    for c in ("start", "end"):
        plan[c] = pd.to_datetime(plan[c], errors="coerce")
    plan["year"] = pd.to_numeric(plan["year"], errors="coerce").astype("Int64")
    return plan


_REFINERY_KEYS = ["BAYTOWN", "BATON ROUGE", "BEAUMONT", "JOLIET", "SARNIA",
                  "STRATHCONA", "NANTICOKE", "BILLINGS", "CHALMETTE", "TORRANCE", "FAWLEY"]


def _site_key(name):
    """Map an IIR plant name or Exxon-plan site name to a common refinery key."""
    s = str(name).upper()
    for k in _REFINERY_KEYS:
        if k in s:
            return k
    return s.replace(" REFINERY", "").replace(" COMPLEX", "").replace(" CHEMICAL PLANT", "").strip()


def _months_between(s, e):
    if pd.isna(s) or pd.isna(e):
        return set()
    cur, last = pd.Timestamp(s.year, s.month, 1), pd.Timestamp(e.year, e.month, 1)
    out = set()
    while cur <= last:
        out.add(cur.month)
        cur = cur + pd.offsets.MonthBegin(1)
    return out


def verify_exxon(df, year=FOCUS_YEAR, plan=None):
    """Cross-check IIR ExxonMobil records for `year` against Exxon's own corporate
    plan, per unit. An IIR event is 'confirmed' when the plan has an outage at the
    same refinery + same focus class overlapping the same months; otherwise it is
    'unverified' -- a likely phantom / mis-dated duplicate (e.g. the IIR Joliet
    'Crude' that appears in Sep-Oct 2027 with no counterpart in the plan, whose
    only Sep-2027 Joliet event is FT Cogen).

    Returns {events (with verified/note cols), flagged, confirmed_kbd,
    flagged_kbd, plan_year}."""
    if plan is None:
        plan = load_exxon_plan()
    ev = unit_events(df, operator_contains="EXXON", year=year)
    if ev.empty:
        return {"events": ev, "flagged": ev, "confirmed_kbd": 0.0, "flagged_kbd": 0.0,
                "plan_year": plan}
    if plan.empty:
        ev = ev.assign(verified=None, note="no corporate plan vendored")
        return {"events": ev, "flagged": ev.iloc[0:0], "confirmed_kbd": float(ev["kbd"].sum()),
                "flagged_kbd": 0.0, "plan_year": plan}
    py = plan[plan["year"] == year].copy()
    py["skey"] = py["site"].map(_site_key)
    cover = {}                                   # (skey, bucket) -> set(plan months)
    for _, r in py.iterrows():
        cover.setdefault((r["skey"], r["bucket"]), set()).update(
            _months_between(r["start"], r["end"]))
    verified, notes = [], []
    for _, e in ev.iterrows():
        sk, fb = _site_key(e["plant"]), e["focus"]
        if not isinstance(fb, str):              # NaN/None -> non-focus unit
            verified.append(None)
            notes.append("non-focus unit - not cross-checked")
            continue
        pm = cover.get((sk, fb), set())
        if pm & set(e["months"]):
            verified.append(True)
            notes.append("matches corporate plan")
        elif pm:
            verified.append(False)
            notes.append(f"plan has {fb} at {sk.title()} but in month(s) "
                         f"{sorted(pm)} not {e['months']} - check dating")
        else:
            verified.append(False)
            notes.append(f"no {fb} outage at {sk.title()} in the {year} corporate "
                         "plan - likely phantom / duplicate")
    ev = ev.assign(verified=verified, note=notes)
    flagged = ev[ev["verified"] == False]
    confirmed = ev[ev["verified"] != False]
    return {"events": ev, "flagged": flagged,
            "confirmed_kbd": float(confirmed["kbd"].sum()),
            "flagged_kbd": float(flagged["kbd"].sum()), "plan_year": py}


# ----------------------------------------------------------------------------- context bundle
def build_context(path):
    """One-shot bundle of every frame the deliverables need.

    The slide deck and HTML dashboard consume this so they never re-aggregate
    raw data; the Excel workbook also uses it for its data blocks (its
    interactive scenario/sensitivity cells are live formulas, not these values).
    """
    df_full = load(path)
    # the model + slides only ever concern 2023 .. current year + 1 (END_YEAR).
    # The golden-record Snowflake spans 2010-2038, so clip it to the window here.
    # Keep the full history for the Data-Quality tab (turnaround-cadence needs it).
    df = df_full[df_full["year"].between(START_YEAR, END_YEAR)].copy()
    _enhanced = ("schema" in df.columns and len(df) and df["schema"].iloc[0] == "enhanced")
    if _enhanced:
        # The Enhanced file is the user's already-reconciled book -> trust it as-is
        # (don't re-flag/drop records the curated source chose to keep).
        _excluded = []
        exxon_ver = verify_exxon(df, FOCUS_YEAR)
        _ev = exxon_ver["events"]
        if "verified" in _ev.columns:
            _ev = _ev.assign(verified=_ev["verified"].map(lambda v: True if v is False else v))
        exxon_ver = {"events": _ev, "flagged": _ev.iloc[0:0],
                     "confirmed_kbd": exxon_ver.get("confirmed_kbd", 0.0),
                     "flagged_kbd": 0.0, "plan_year": exxon_ver.get("plan_year")}
    else:
        # Legacy Snowflake export: EXCLUDE the ExxonMobil records that fail the
        # corporate-plan check (Joliet Sep-Oct 'Crude' duplicate + the 2027 FCC the
        # plan books in 2026/2030) from every deliverable.
        _ver0 = verify_exxon(df, FOCUS_YEAR)
        _excluded = (_ver0["flagged"][["plant", "unit_name", "kbd", "span", "note"]].to_dict("records")
                     if not _ver0["flagged"].empty else [])
        _bad = ({str(x) for x in _ver0["flagged"]["outage_id"].dropna()}
                if not _ver0["flagged"].empty else set())
        if _bad:
            df = df[~df["outage_id"].astype(str).isin(_bad)].copy()
        exxon_ver = verify_exxon(df, FOCUS_YEAR)

    diag = diagnostics(df)
    summary = yoy_delta(df)

    ctx = {
        "df": df,
        "df_full": df_full,                 # unclipped history, for the Data-Quality cadence audit
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
        "ta_2027": {p: turnaround_schedule(df, FOCUS_YEAR, padd=p, top=8) for p in PADD_ORDER},
        "display_annual": display_annual(df),
        "h1_planned": h1_planned(df),
        "scenario_fan": scenario_fan(df),
        "completed_unplanned": completed_unplanned(df),
        "crack": load_crack(),
        "dollar_impact": outage_dollar_impact(df, load_crack()),
        "crack_corr": crack_outage_relationship(df, load_crack()),
        "period_change": period_change(df),
        # --- per-unit capacity offline (2021+), the focus of the deck ---
        #     df is already plan-cleaned above; 2027 = Exxon full-year + non-Exxon
        #     H1 confirmed, non-Exxon H2 indicative (see confirmed2027).
        "focus_monthly": focus_unit_monthly(df),                       # all outages
        "focus_planned": focus_unit_monthly(df, type_filter="PLANNED"),
        "focus_peak": focus_annual_peak(df),
        "h1_focus_planned": h1_focus_planned(df),       # H1 planned per unit, 2025/26/27
        "naphtha_balance": naphtha_balance(df, FOCUS_YEAR),   # CDU supply vs reformer demand (outlook yr)
        "naphtha_balance_cy": naphtha_balance(df, CURRENT_YEAR),   # same, for the in-progress year
        "naphtha_balance_p3": naphtha_balance(df, FOCUS_YEAR, padd="PADD 3"),       # HVN balance, PADD 3 (Gulf) only
        "naphtha_balance_cy_p3": naphtha_balance(df, CURRENT_YEAR, padd="PADD 3"),
        "focus_padd": {y: {f: focus_unit_padd_month(df, f, y) for f in FOCUS_ORDER}
                       for y in (FOCUS_YEAR - 1, FOCUS_YEAR)},
        "confirmed2027": {f: focus_2027_split(df, f) for f in FOCUS_ORDER},
        "unit_events_2027": unit_events(df, year=FOCUS_YEAR, type_filter="PLANNED"),
        "exxon_verify": exxon_ver,
        "deck_excluded": _excluded,
        "padd_month": {p: {
            "total": padd_month_year(df, p),
            "planned": padd_month_year(df, p, type_filter="PLANNED"),
            "unplanned": padd_month_year(df, p, type_filter="UNPLANNED"),
        } for p in PADD_ORDER},
        "compare": {
            "2025v2026": compare_block(df, 2025, 2026),
            "2025v2027": compare_block(df, 2025, FOCUS_YEAR),
            "2026v2027": compare_block(df, FOCUS_YEAR - 1, FOCUS_YEAR),
        },
    }
    return ctx


if __name__ == "__main__":
    import json
    import sys
    from pathlib import Path
    default = str(Path(__file__).resolve().parent.parent / "data" / "Golden_Record_Snowflake.xlsx")
    df = load(sys.argv[1] if len(sys.argv) > 1 else default)
    print(json.dumps(diagnostics(df), indent=2, default=str))
