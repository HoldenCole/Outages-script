"""
outage_engine.py
Reusable aggregation core for refinery outage analysis.

Turns the raw Snowflake export (rEFINERY oUTAGES.xlsx / Query1) into clean,
analysis-ready frames and pivots. Everything downstream (workbook, slides,
dashboard) is built on top of the frames this module returns.

Design decisions (locked with desk):
  * PRIMARY metric  = CAP_OFFLINE_ADJUSTED_KBD  (total offline capacity, all units/products)
  * SECONDARY view  = mogas-equivalent (capacity x unit yield) -- an optional overlay
  * UNKNOWN outage type folds into UNPLANNED
  * PADD parsed from Roman-numeral PAD_DIST ; state map is a fallback
  * 2027 is planned-only in this dataset -> unplanned 2027 is a *scenario*, never an actual
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

MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

# Roman PADD label -> canonical "PADD n"
ROMAN_PADD = {
    "PADD I": "PADD 1", "PADD II": "PADD 2", "PADD III": "PADD 3",
    "PADD IV": "PADD 4", "PADD V": "PADD 5", "PADD CARIBBEAN": "PADD Caribbean",
}

# state -> PADD (fallback when PAD_DIST unusable)
STATE_PADD = {
    # PADD 1
    "CT":1,"DE":1,"DC":1,"FL":1,"GA":1,"ME":1,"MD":1,"MA":1,"NH":1,"NJ":1,
    "NY":1,"NC":1,"PA":1,"RI":1,"SC":1,"VT":1,"VA":1,"WV":1,
    # PADD 2
    "IL":2,"IN":2,"IA":2,"KS":2,"KY":2,"MI":2,"MN":2,"MO":2,"NE":2,"ND":2,
    "OH":2,"OK":2,"SD":2,"TN":2,"WI":2,
    # PADD 3
    "AL":3,"AR":3,"LA":3,"MS":3,"NM":3,"TX":3,
    # PADD 4
    "CO":4,"ID":4,"MT":4,"UT":4,"WY":4,
    # PADD 5
    "AK":5,"AZ":5,"CA":5,"HI":5,"NV":5,"OR":5,"WA":5,
}

# Mogas yield buckets (from Yields.txt). Unit category -> bucket.
# Buckets carry a mogas yield factor; everything else -> "Other" (0 mogas).
YIELD_FACTOR = {"CDU":0.175, "FCC":0.65, "Ref":0.85, "HDC":0.05, "Coker":0.20, "Other":0.0}
UNITCAT_TO_BUCKET = {
    "ATMOS DISTILLATION":"CDU",
    "VACUUM DISTILLATION":"CDU",       # topping/vacuum tied to crude train
    "FLUID CAT CRACKING":"FCC",
    "REFORMING":"Ref",
    "HYDROCRACKING":"HDC",
    "RESID_HYDROCRACKING":"HDC",
    "COKING":"Coker",
    "THERM CRACKING, VISBREAKING":"Coker",
    # non-mogas / negligible-gasoline units -> Other (capacity still tracked)
    "HYDROTREATING":"Other","ALKYLATION":"Other","ISOMERIZATION":"Other",
    "ASPHALT":"Other","BTX":"Other","MTBE":"Other","AROMATICS":"Other",
    "GAS PROCESSING":"Other","OTHER":"Other",
}

PADD_ORDER = ["PADD 1","PADD 2","PADD 3","PADD 4","PADD 5"]


# ----------------------------------------------------------------------------- load + clean
def _to_padd_from_roman(val):
    if pd.isna(val):
        return None
    key = str(val).strip().upper()
    return ROMAN_PADD.get(key)

def _to_padd_from_state(st):
    if pd.isna(st):
        return None
    n = STATE_PADD.get(str(st).strip().upper())
    return f"PADD {n}" if n else None

def load(path):
    """Load raw export, return cleaned long-form dataframe."""
    df = pd.read_excel(path, sheet_name=RAW_SHEET)
    out = pd.DataFrame()
    for logical, src in COLMAP.items():
        out[logical] = df[src] if src in df.columns else np.nan

    out["year"]  = pd.to_numeric(out["year"], errors="coerce").astype("Int64")
    out["month"] = pd.to_numeric(out["month"], errors="coerce").astype("Int64")
    out["cap_kbd"] = pd.to_numeric(out["cap_kbd"], errors="coerce").fillna(0.0)

    # outage type: fold UNKNOWN -> UNPLANNED, normalize to {PLANNED, UNPLANNED}
    t = out["otype"].astype(str).str.strip().str.upper()
    out["type"] = np.where(t.eq("PLANNED"), "PLANNED", "UNPLANNED")

    # PADD: prefer Roman label, fallback to state
    p_roman = out["pad_dist"].apply(_to_padd_from_roman)
    p_state = out["state"].apply(_to_padd_from_state)
    out["padd"] = p_roman.where(p_roman.notna(), p_state)
    out["padd_source"] = np.where(p_roman.notna(), "PAD_DIST",
                          np.where(p_state.notna(), "STATE", "UNRESOLVED"))

    # mogas overlay
    out["bucket"] = out["unit_cat"].map(UNITCAT_TO_BUCKET).fillna("Other")
    out["mogas_kbd"] = out["cap_kbd"] * out["bucket"].map(YIELD_FACTOR).fillna(0.0)

    out["month_name"] = out["month"].map(lambda m: MONTHS[int(m)-1] if pd.notna(m) and 1<=int(m)<=12 else None)
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
        "2027_types": df.loc[df.year.eq(2027),"type"].value_counts().to_dict(),
    }


# ----------------------------------------------------------------------------- pivots
def _month_matrix(sub, value="cap_kbd"):
    """year x month matrix (sum of value), months as Jan..Dec columns."""
    g = (sub.groupby(["year","month"])[value].sum()
            .unstack("month").reindex(columns=range(1,13)))
    g.columns = MONTHS
    g = g.reindex(sorted([int(y) for y in g.index if pd.notna(y)]))
    return g.fillna(0.0)

def monthly_by_year(df, value="cap_kbd", type_filter=None, padd=None):
    sub = df
    if type_filter: sub = sub[sub["type"].eq(type_filter)]
    if padd:        sub = sub[sub["padd"].eq(padd)]
    return _month_matrix(sub, value)

def annual_summary(df, value="cap_kbd"):
    """year x {Planned, Unplanned, Total, EventCount} table."""
    g = df.groupby(["year","type"])[value].sum().unstack("type").fillna(0.0)
    for c in ["PLANNED","UNPLANNED"]:
        if c not in g: g[c] = 0.0
    g = g.rename(columns={"PLANNED":"Planned","UNPLANNED":"Unplanned"})
    g["Total"] = g["Planned"] + g["Unplanned"]
    ev = df.groupby("year")["outage_id"].count().rename("Events")
    g = g.join(ev)
    g.index = [int(y) for y in g.index]
    return g.sort_index()

def padd_year_matrix(df, value="cap_kbd", type_filter=None):
    sub = df if not type_filter else df[df["type"].eq(type_filter)]
    sub = sub[sub["padd"].isin(PADD_ORDER)]
    g = sub.groupby(["padd","year"])[value].sum().unstack("year").fillna(0.0)
    g = g.reindex(PADD_ORDER)
    g.columns = [int(c) for c in g.columns]
    return g

def unit_year_matrix(df, value="cap_kbd", type_filter=None):
    sub = df if not type_filter else df[df["type"].eq(type_filter)]
    g = sub.groupby(["unit_cat","year"])[value].sum().unstack("year").fillna(0.0)
    g.columns = [int(c) for c in g.columns]
    g["__tot"] = g.sum(axis=1)
    g = g.sort_values("__tot", ascending=False).drop(columns="__tot")
    return g

def seasonality(df, years, type_filter="UNPLANNED", padd=None, value="cap_kbd"):
    """Average monthly profile across given years (the forecast backbone)."""
    sub = df[df["type"].eq(type_filter) & df["year"].isin(years)]
    if padd: sub = sub[sub["padd"].eq(padd)]
    by_ym = sub.groupby(["year","month"])[value].sum().unstack("month").reindex(columns=range(1,13)).fillna(0.0)
    prof = by_ym.mean(axis=0)            # avg kbd per calendar month
    prof.index = MONTHS
    return prof


if __name__ == "__main__":
    import json, sys
    df = load(sys.argv[1] if len(sys.argv)>1 else "/mnt/user-data/uploads/rEFINERY_oUTAGES.xlsx")
    print(json.dumps(diagnostics(df), indent=2, default=str))


# ----------------------------------------------------------------------------- v2 extensions
def padd_month_year(df, padd, value="cap_kbd", type_filter=None):
    """For one PADD: year x month matrix (the per-PADD chart backbone)."""
    sub = df[df["padd"].eq(padd)]
    if type_filter: sub = sub[sub["type"].eq(type_filter)]
    return _month_matrix(sub, value)

def operator_year(df, value="cap_kbd", type_filter=None, top=12):
    sub = df if not type_filter else df[df["type"].eq(type_filter)]
    g = sub.groupby(["operator","year"])[value].sum().unstack("year").fillna(0.0)
    g.columns=[int(c) for c in g.columns]
    g["__t"]=g.sum(axis=1); g=g.sort_values("__t",ascending=False).drop(columns="__t").head(top)
    return g

def plant_detail(df, value="cap_kbd", top=15):
    """Refinery-level detail with PADD, operator, planned/unplanned split (recent year focus)."""
    g = df.groupby(["plant","padd","operator"]).agg(
        total=(value,"sum"),
        planned=(value, lambda s: s[df.loc[s.index,"type"].eq("PLANNED")].sum()),
        events=("outage_id","count")).reset_index()
    g["unplanned"]=g["total"]-g["planned"]
    return g.sort_values("total",ascending=False).head(top)

def event_scatter(df, years, type_filter="UNPLANNED"):
    """Event-level: duration (monthly days) vs capacity, for scatter."""
    sub=df[df["type"].eq(type_filter) & df["year"].isin(years)].copy()
    sub["dur"]=pd.to_numeric(sub["__dummy"], errors="coerce") if "__dummy" in sub else np.nan
    return sub

def load_with_duration(path):
    df=load(path)
    raw=pd.read_excel(path, sheet_name=RAW_SHEET, usecols=["TOTAL_OUTAGE_DAYS"])
    df["dur"]=pd.to_numeric(raw["TOTAL_OUTAGE_DAYS"], errors="coerce")
    return df
