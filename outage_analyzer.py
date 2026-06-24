#!/usr/bin/env python3
import argparse
import os
import re
import sys

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio
from plotly.offline import get_plotlyjs


# ===========================================================================
# CONFIG — the only part you normally edit
# ===========================================================================

# Map each logical field to YOUR actual Excel header. Matching is case-insensitive
# and ignores spaces, underscores, punctuation, and hyphens.
COLUMN_MAP = {
    "month_date": "OUTAGE_MONTH_DATE",
    "month": "OUTAGE_MONTH",
    "year": "OUTAGE_YEAR",
    "start_date": "OUTAGE_START_DATE",
    "source": "OUTAGE_SOURCE",
    "end_date": "OUTAGE_END_DATE",
    "total_days": "TOTAL_OUTAGE_DAYS",
    "month_days": "TOTAL_MONTH_DAYS",
    "pct_month_cal": "PERCENTAGE_MONTH_CAL",
    "pct_month": "PERCENTAGE_MONTH",

    "offline_cap": "OFFLINE_CAPACITY",  # BBL/d offline for this outage
    "unit_capacity": "UNIT_CAPACITY",  # nameplate of the unit (BBL/d)
    "unit_type": "UNIT_CATEGORY",  # CDU / FCC / Coker / Reformer ...
    "unit_id": "UNIT_ID",  # real unit identifier from the export
    "refinery": "PLANT_NAME",  # refinery name
    "state": "REFINERY_STATE",  # state
    "padd": "PAD_DIST",  # 1..5 (numeric or "PADD 3")
}

# Bridge the export's UNIT_CATEGORY vocabulary to the Yields file's unit names.
# Matching elsewhere is done via _norm() (case/space/punct-insensitive). Only the
# categories we have yields for are mapped; every other category is intentionally
# left unmapped so its mogas stays blank (NaN), never zero.
UNIT_TYPE_ALIASES = {
    "ATMOS DISTILLATION": "CDU",
    "FLUID CAT CRACKING": "FCC",
    "REFORMING": "Ref",
    "HYDROCRACKING": "HDC",
    "COKING": "Coker",
    # Intentionally NOT mapped (left blank, no known mogas yield):
    #   RESID_HYDROCRACKING, ALKYLATION, AROMATICS, BTX, ASPHALT,
    #   GAS PROCESSING, HYDROTREATING, ISOMERIZATION, MTBE, OTHER,
    #   THERM CRACKING/VISBREAKING, VACUUM DISTILLATION
}

# Pipeline stops with a clear message if any of these don't map.
REQUIRED_FIELDS = ["year", "start_date", "end_date", "source", "offline_cap", "unit_type", "unit_id", "refinery", "padd"]

# --- the knobs ---
BIG_OUTAGE_BBL = 100_000  # big-outage flag fires at/above this BBL/d
YOY_HOT_THRESHOLD = 0.25  # month "hot" if >=25% above prior year
BACK_TO_BACK_MIN_MONTHS = 2  # consecutive months to flag a cluster
SPOF_REFINERY_MAX_UNITS = 1  # <= this many units of a type = SPOF
PADD_PRIORITY = [2, 1, 3, 5, 4]  # your ranking order, most important first
MANUAL_SOURCE_VALUES = ["manual"]  # source values treated as manual entry
YIELDS_FILE = "Yields"  # tab-separated PADD / Unit / Mogas Yield lookup
OUTPUT_HTML = "refinery_outage_dashboard.html"


# ===========================================================================
# LOADER + VALIDATOR
# ===========================================================================

def _norm(s):
    return re.sub(r"[^a-z0-9]+", "", str(s).strip().lower())


def _safe_int(v):
    """int() that tolerates NaN/None, returning an em dash for display."""
    return f"{int(v)}" if pd.notna(v) else "\u2014"


def load_yields(path=YIELDS_FILE):
    """Load the tab-separated mogas yield lookup keyed by (padd, unit_type).

    The file's structure is fixed (PADD / Unit / Mogas Yield); only the values
    change. Keys are normalized with _norm() so the join to the outage data is
    case- and spacing-insensitive. PADD labels like "PADD 1" are parsed to 1.
    """
    try:
        y = pd.read_csv(path, sep="\t")
    except FileNotFoundError:
        raise SystemExit(
            f"ERROR: mogas yields file '{path}' not found.\n"
            "It must sit next to the script with columns: PADD, Unit, Mogas Yield."
        )
    except Exception as e:
        raise SystemExit(f"ERROR: could not read mogas yields file '{path}': {e}")

    cols = {_norm(c): c for c in y.columns}
    need = {"padd": "PADD", "unit": "Unit", "mogasyield": "Mogas Yield"}
    missing = [label for key, label in need.items() if key not in cols]
    if missing:
        raise SystemExit(
            f"ERROR: mogas yields file '{path}' is missing column(s): {', '.join(missing)}.\n"
            "Expected columns: PADD, Unit, Mogas Yield."
        )

    y = y.rename(columns={cols["padd"]: "padd", cols["unit"]: "unit", cols["mogasyield"]: "yield"})
    y["padd"] = y["padd"].astype(str).str.extract(r"(\d)").astype(float)
    y["unit_key"] = y["unit"].apply(_norm)
    y["yield"] = pd.to_numeric(y["yield"], errors="coerce")
    lookup = {(p, u): yld for p, u, yld in zip(y["padd"], y["unit_key"], y["yield"])}
    return lookup


def attach_mogas(df, yields):
    """Add mogas_offline = offline_cap * yield, keyed by (padd, unit_type).

    The export's UNIT_CATEGORY values are first mapped through UNIT_TYPE_ALIASES
    to the Yields file's unit names. Unit types with no matching yield row get
    NaN (never silently zeroed); the caller reports the ORIGINAL category names
    that went unmatched. Yields are broken down by PADD and unit type.
    """
    aliased = df["unit_type"].map(lambda v: UNIT_TYPE_ALIASES.get(str(v).strip().upper(), v))
    unit_key = aliased.apply(_norm)
    yld = [yields.get((p, u), np.nan) for p, u in zip(df["padd"], unit_key)]
    df["mogas_yield"] = yld
    df["mogas_offline"] = df["offline_cap"] * df["mogas_yield"]
    unmatched = sorted(
        {str(ut) for ut, m in zip(df["unit_type"], df["mogas_yield"]) if pd.isna(m) and pd.notna(ut)}
    )
    return df, unmatched


def load(path):
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

    missing_required = [m for m in missing if m[0] in REQUIRED_FIELDS]
    if missing_required:
        lines = "\n".join(f" - logical '{lg}' expected header '{hd}'" for lg, hd in missing_required)
        matched = "\n".join(f" - logical '{lg}' matched '{hd}'" for lg, hd in found.items()) or " - none"
        raise SystemExit(
            "ERROR: required columns not found.\n"
            + lines
            + "\n\nMatched headers:\n"
            + matched
            + "\n\nFix COLUMN_MAP at the top of this file to match your headers."
        )

    for c in ("start_date", "end_date", "month_date"):
        if c in df:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    for c in ("year", "month", "total_days", "month_days", "offline_cap", "unit_capacity", "pct_month_cal"):
        if c in df:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    # 'month' is optional in the export but several views group/pivot on it.
    # Derive it from start_date when it is absent or entirely empty so the
    # year-over-year analysis never crashes with a KeyError.
    if ("month" not in df or df["month"].isna().all()) and "start_date" in df:
        df["month"] = df["start_date"].dt.month

    if "padd" in df:
        df["padd"] = df["padd"].astype(str).str.extract(r"(\d)").astype(float)

    manual = {m.lower() for m in MANUAL_SOURCE_VALUES}
    df["is_manual"] = df["source"].astype(str).str.strip().str.lower().isin(manual)
    df["source_clean"] = df["source"].astype(str).str.strip()

    if "total_days" not in df or df["total_days"].isna().all():
        df["total_days"] = (df["end_date"] - df["start_date"]).dt.days
    df["bbl_days"] = df["offline_cap"] * df["total_days"]

    if "unit_capacity" in df:
        df["pct_unit_offline"] = np.where(df["unit_capacity"] > 0, df["offline_cap"] / df["unit_capacity"], np.nan)

    report = {
        "n_rows": len(df),
        "found": found,
        "missing_optional": [m for m in missing if m[0] not in REQUIRED_FIELDS],
    }
    issues = pd.DataFrame(index=df.index)
    issues["end_before_start"] = df["end_date"] < df["start_date"]
    issues["negative_days"] = df["total_days"] < 0
    if "unit_capacity" in df:
        issues["offline_exceeds_nameplate"] = df["offline_cap"] > df["unit_capacity"]
    issues["missing_offline_cap"] = df["offline_cap"].isna()
    issues["missing_padd"] = df["padd"].isna()
    issues["missing_unit_id"] = df["unit_id"].isna()

    df["has_quality_issue"] = issues.any(axis=1)
    report["issue_counts"] = {c: int(issues[c].sum()) for c in issues.columns}
    report["n_issue_rows"] = int(df["has_quality_issue"].sum())
    report["n_manual"] = int(df["is_manual"].sum())
    report["n_feed"] = int((~df["is_manual"]).sum())
    report["source_breakdown"] = df["source_clean"].value_counts().to_dict()

    # Mogas: offline_cap * yield, yield chosen by (PADD, unit type).
    yields = load_yields()
    df, unmatched = attach_mogas(df, yields)
    report["mogas_total"] = float(df["mogas_offline"].sum(skipna=True))
    report["mogas_unmatched_units"] = unmatched
    return df, report


# ===========================================================================
# ANALYSIS + SIGNALS
# ===========================================================================

def padd_rank_key(padd):
    try:
        return PADD_PRIORITY.index(int(padd))
    except (ValueError, TypeError):
        return len(PADD_PRIORITY)


def build_redundancy(df):
    # Real unit-ID based counts: one unit_id = one unit.
    # NOTE: rows with a NaN padd are dropped from the PADD-level count (they are
    # already surfaced via the 'missing_padd' quality flag).
    ref_units = (
        df.groupby(["refinery", "unit_type"])["unit_id"].nunique().rename("units_in_refinery").reset_index()
    )
    # Distinct units per refinery, then sum across refineries within each PADD.
    # Using reset_index(name=...) avoids relying on groupby.apply returning a
    # Series (which broke .rename(...) on some pandas versions).
    per_ref = (
        df.dropna(subset=["padd"])
        .groupby(["padd", "unit_type", "refinery"])["unit_id"]
        .nunique()
        .reset_index(name="u")
    )
    padd_units = (
        per_ref.groupby(["padd", "unit_type"])["u"]
        .sum()
        .reset_index(name="units_in_padd")
    )
    return ref_units, padd_units


def attach_redundancy(df, ref_units, padd_units):
    df = df.merge(ref_units, on=["refinery", "unit_type"], how="left")
    df = df.merge(padd_units, on=["padd", "unit_type"], how="left")
    df["spof_refinery"] = df["units_in_refinery"] <= SPOF_REFINERY_MAX_UNITS
    df["thin_in_padd"] = df["units_in_padd"] <= 2
    return df


def monthly_series(df):
    return (
        df.assign(ym=df["start_date"].dt.to_period("M").dt.to_timestamp())
        .groupby("ym")
        .agg(bbl_days=("bbl_days", "sum"), events=("offline_cap", "size"), bbls_offline=("offline_cap", "sum"), mogas_offline=("mogas_offline", "sum"))
        .reset_index()
    )


def yoy_by_month(df):
    # year is required; month is derived in load() when missing. Drop any rows
    # still lacking month/year so the pivot in render() never sees NaN labels.
    d = df.dropna(subset=["year", "month"])
    g = d.groupby(["year", "month"]).agg(bbl_days=("bbl_days", "sum"), events=("offline_cap", "size")).reset_index()
    g["prev"] = g.groupby("month")["bbl_days"].shift(1)
    g["yoy_change"] = np.where(g["prev"].notna() & (g["prev"] != 0), (g["bbl_days"] - g["prev"]) / g["prev"], np.nan)
    g["hot"] = g["yoy_change"] >= YOY_HOT_THRESHOLD
    return g


def by_padd(df):
    g = df.groupby("padd").agg(events=("offline_cap", "size"), bbls_offline=("offline_cap", "sum"), bbl_days=("bbl_days", "sum"), mogas_offline=("mogas_offline", "sum")).reset_index()
    g["padd_priority"] = g["padd"].apply(padd_rank_key)
    return g.sort_values("padd_priority")


def by_unit(df):
    return (
        df.groupby("unit_type")
        .agg(events=("offline_cap", "size"), bbls_offline=("offline_cap", "sum"), bbl_days=("bbl_days", "sum"), mogas_offline=("mogas_offline", "sum"))
        .reset_index()
        .sort_values("bbl_days", ascending=False)
    )


def rank_outages(df, top=25):
    d = df.copy()
    d["padd_priority"] = d["padd"].apply(padd_rank_key)
    d["severity"] = d["offline_cap"] * (1.0 + 0.5 * d["spof_refinery"].astype(int) + 0.25 * d["thin_in_padd"].astype(int))
    d = d.sort_values(["severity", "padd_priority", "bbl_days"], ascending=[False, True, False])
    cols = [
        "refinery",
        "state",
        "padd",
        "unit_type",
        "unit_id",
        "offline_cap",
        "mogas_offline",
        "unit_capacity",
        "pct_unit_offline",
        "total_days",
        "bbl_days",
        "start_date",
        "end_date",
        "source_clean",
        "is_manual",
        "spof_refinery",
        "thin_in_padd",
        "severity",
    ]
    return d[[c for c in cols if c in d.columns]].head(top).reset_index(drop=True)


def signal_big_outages(df):
    big = df[df["offline_cap"] >= BIG_OUTAGE_BBL].copy()
    big["padd_priority"] = big["padd"].apply(padd_rank_key)
    return big.sort_values(["offline_cap", "padd_priority"], ascending=[False, True])


def signal_back_to_back(df):
    d = (
        df.assign(ym=df["start_date"].dt.to_period("M"))
        .groupby(["padd", "unit_type", "ym"])
        .agg(events=("offline_cap", "size"), bbls=("offline_cap", "sum"))
        .reset_index()
        .sort_values(["padd", "unit_type", "ym"])
    )
    clusters = []
    for (padd, unit), g in d.groupby(["padd", "unit_type"]):
        months = g["ym"].tolist()
        if not months:
            continue
        run = [months[0]]
        for prev, cur in zip(months, months[1:]):
            if cur == prev + 1:
                run.append(cur)
            else:
                if len(run) >= BACK_TO_BACK_MIN_MONTHS:
                    clusters.append((padd, unit, run[0], run[-1], len(run)))
                run = [cur]
        if len(run) >= BACK_TO_BACK_MIN_MONTHS:
            clusters.append((padd, unit, run[0], run[-1], len(run)))
    if not clusters:
        return pd.DataFrame(columns=["padd", "unit_type", "from", "to", "months"])
    return pd.DataFrame(clusters, columns=["padd", "unit_type", "from", "to", "months"]).sort_values("months", ascending=False)


def signal_spof(df):
    return df[df["spof_refinery"]].sort_values("offline_cap", ascending=False)


def build_all(df):
    ref_units, padd_units = build_redundancy(df)
    df = attach_redundancy(df, ref_units, padd_units)
    return df, {
        "monthly": monthly_series(df),
        "yoy": yoy_by_month(df),
        "by_padd": by_padd(df),
        "by_unit": by_unit(df),
        "ranked": rank_outages(df),
        "big": signal_big_outages(df),
        "back_to_back": signal_back_to_back(df),
        "spof": signal_spof(df),
    }


# ===========================================================================
# DASHBOARD RENDERER
# ===========================================================================

INK, PAPER, HAIR = "#14181d", "#0d1014", "#262d36"
TEXT, MUTE, ACCENT, HOT, COOL, MANUAL = "#e6e9ee", "#8a95a3", "#e8a13a", "#d6453d", "#4a8fb8", "#caa24a"

PLOT_LAYOUT = dict(
    paper_bgcolor=INK,
    plot_bgcolor=INK,
    font=dict(color=TEXT, family="ui-monospace, Menlo, monospace", size=12),
    margin=dict(l=56, r=20, t=36, b=40),
    xaxis=dict(gridcolor=HAIR, zerolinecolor=HAIR),
    yaxis=dict(gridcolor=HAIR, zerolinecolor=HAIR),
    legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(size=11)),
    colorway=[ACCENT, COOL, HOT, "#7bbf6a", "#b07bd1", MUTE],
)


def _fig_html(fig, div_id):
    fig.update_layout(**PLOT_LAYOUT)
    return pio.to_html(fig, include_plotlyjs=False, full_html=False, div_id=div_id, config={"displayModeBar": False})


def render(df, res, report, path=OUTPUT_HTML):
    n = lambda v: f"{v:,.0f}" if pd.notna(v) else "\u2014"
    pct = lambda v: f"{v*100:,.0f}%" if pd.notna(v) else "\u2014"
    d = lambda v: pd.to_datetime(v).strftime("%Y-%m-%d") if pd.notna(v) else "\u2014"

    big, b2b, spof, ranked = res["big"], res["back_to_back"], res["spof"], res["ranked"]
    takeaways = []
    if len(big):
        t = big.iloc[0]
        takeaways.append(f"Largest single outage: <b>{t['refinery']}</b> {t['unit_type']} (PADD {_safe_int(t['padd'])}) — <b>{n(t['offline_cap'])} BBL/d</b> offline for {n(t['total_days'])} days.")
        takeaways.append(f"<b>{len(big)}</b> outages at/above the {n(BIG_OUTAGE_BBL)} BBL/d threshold.")
    if len(b2b):
        c = b2b.iloc[0]
        takeaways.append(f"Longest back-to-back cluster: <b>{c['months']} months</b> of {c['unit_type']} in PADD {_safe_int(c['padd'])} ({c['from']} → {c['to']}).")
    hot = res["yoy"][res["yoy"]["hot"]]
    if len(hot):
        takeaways.append(f"<b>{len(hot)}</b> month(s) running above prior year by threshold.")
    if len(spof):
        takeaways.append(f"<b>{len(spof)}</b> outage(s) on single-point-of-failure units.")
    takeaways.append(f"Total mogas offline (capacity × yield): <b>{n(report['mogas_total'])} BBL/d</b>.")
    takeaways.append(f"Provenance: <b>{report['n_manual']}</b> manual, <b>{report['n_feed']}</b> feed rows. {report['n_issue_rows']} flagged for review.")

    fig_m = go.Figure()
    fig_m.add_bar(x=res["monthly"]["ym"], y=res["monthly"]["bbls_offline"], name="BBL/d offline", marker_color=ACCENT)
    fig_m.add_bar(x=res["monthly"]["ym"], y=res["monthly"]["mogas_offline"], name="mogas BBL/d", marker_color=COOL)
    fig_m.add_scatter(x=res["monthly"]["ym"], y=res["monthly"]["events"], name="events", yaxis="y2", mode="lines", line=dict(color=HOT, width=1.5))
    fig_m.update_layout(title="Monthly outage volume", barmode="overlay", yaxis2=dict(overlaying="y", side="right", showgrid=False, title="events"))

    bp = res["by_padd"]
    fig_p = go.Figure()
    fig_p.add_bar(x=[f"PADD {_safe_int(p)}" for p in bp["padd"]], y=bp["bbls_offline"], name="BBL/d offline", marker_color=ACCENT, text=[f"{v:,.0f}" for v in bp["bbls_offline"]], textposition="outside")
    fig_p.add_bar(x=[f"PADD {_safe_int(p)}" for p in bp["padd"]], y=bp["mogas_offline"], name="mogas BBL/d", marker_color=COOL, text=[f"{v:,.0f}" for v in bp["mogas_offline"]], textposition="outside")
    fig_p.update_layout(title="BBL/d offline by PADD — total vs. mogas (priority order)", barmode="group")

    bu = res["by_unit"]
    fig_u = go.Figure()
    fig_u.add_bar(x=bu["bbl_days"], y=bu["unit_type"], name="BBL-days", orientation="h", marker_color=COOL)
    fig_u.update_layout(title="BBL-days offline by unit type", yaxis=dict(autorange="reversed"))

    piv = res["yoy"].pivot(index="month", columns="year", values="bbl_days")
    fig_y = go.Figure()
    for yr in piv.columns:
        fig_y.add_scatter(x=piv.index, y=piv[yr], mode="lines", name=str(int(yr)), line=dict(width=1.2))
    fig_y.update_layout(title="Year-over-year BBL-days by month", xaxis=dict(tickmode="array", tickvals=list(range(1, 13))))

    charts = {"monthly": _fig_html(fig_m, "c_monthly"), "padd": _fig_html(fig_p, "c_padd"), "unit": _fig_html(fig_u, "c_unit"), "yoy": _fig_html(fig_y, "c_yoy")}

    def table(rows, cols, headers, fmt=None, hl_manual=True):
        fmt = fmt or {}
        out = ['<table class="grid"><thead><tr>'] + [f"<th>{h}</th>" for h in headers] + ["</tr></thead><tbody>"]
        for _, r in rows.iterrows():
            cls = []
            if hl_manual and r.get("is_manual"):
                cls.append("manual")
            if r.get("spof_refinery"):
                cls.append("spof")
            out.append(f'<tr class="{" ".join(cls)}">')
            for c in cols:
                v = r.get(c, "")
                if c in fmt:
                    try:
                        v = fmt[c](v)
                    except Exception:
                        pass
                badge = ' <span class="badge">●</span>' if c == "spof_refinery" and r.get(c) else ""
                out.append(f"<td>{v}{badge}</td>")
            out.append("</tr>")
        out.append("</tbody></table>")
        return "".join(out)

    ranked_tbl = table(
        ranked,
        ["refinery", "padd", "unit_type", "unit_id", "offline_cap", "mogas_offline", "pct_unit_offline", "total_days", "bbl_days", "start_date", "source_clean", "spof_refinery"],
        ["Refinery", "PADD", "Unit", "Unit ID", "BBL/d off", "Mogas BBL/d", "% unit", "Days", "BBL-days", "Start", "Source", "SPOF"],
        {"padd": _safe_int, "offline_cap": n, "mogas_offline": n, "pct_unit_offline": pct, "bbl_days": n, "start_date": d, "spof_refinery": lambda v: ""},
    )
    if len(b2b):
        b2b_tbl = table(
            b2b.assign(padd=b2b["padd"].apply(_safe_int), from_=b2b["from"].astype(str), to_=b2b["to"].astype(str)),
            ["padd", "unit_type", "from_", "to_", "months"],
            ["PADD", "Unit", "From", "To", "Months"],
            hl_manual=False,
        )
    else:
        b2b_tbl = "<table class='grid'><tbody><tr><td>No clusters found.</td></tr></tbody></table>"

    src_rows = "".join(f"<tr><td>{k}</td><td class='num'>{v}</td></tr>" for k, v in sorted(report["source_breakdown"].items(), key=lambda x: -x[1]))
    padd_order = ", ".join(str(p) for p in PADD_PRIORITY)
    plotly_js = get_plotlyjs()

    html = f"""<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Refinery Outage Dashboard</title>
<script>{plotly_js}</script>
<style>
:root{{--ink:{INK};--paper:{PAPER};--hair:{HAIR};--text:{TEXT};--mute:{MUTE};
--accent:{ACCENT};--hot:{HOT};--cool:{COOL};--manual:{MANUAL};}}
*{{box-sizing:border-box}}
body{{margin:0;background:var(--paper);color:var(--text);
font-family:ui-sans-serif,-apple-system,"Segoe UI",Roboto,sans-serif;font-size:14px;line-height:1.5}}
.num,td.num{{font-family:ui-monospace,Menlo,monospace;text-align:right}}
header{{padding:22px 28px;border-bottom:1px solid var(--hair);display:flex;align-items:baseline;gap:18px;flex-wrap:wrap}}
header h1{{font-size:18px;margin:0;letter-spacing:.3px}}
header .meta{{color:var(--mute);font-size:12px;font-family:ui-monospace,monospace}}
.wrap{{max-width:1280px;margin:0 auto;padding:24px 28px 64px}}
.takeaways{{background:var(--ink);border:1px solid var(--hair);border-left:3px solid var(--accent);
border-radius:8px;padding:18px 22px;margin-bottom:22px}}
.takeaways h2{{font-size:12px;text-transform:uppercase;letter-spacing:1.5px;color:var(--mute);margin:0 0 12px}}
.takeaways ul{{margin:0;padding-left:18px}} .takeaways li{{margin:6px 0}}
.controls{{display:flex;gap:14px;flex-wrap:wrap;margin:0 0 20px;align-items:end}}
.controls label{{display:flex;flex-direction:column;gap:4px;font-size:11px;text-transform:uppercase;letter-spacing:.8px;color:var(--mute)}}
.controls select{{background:var(--ink);color:var(--text);border:1px solid var(--hair);border-radius:6px;padding:7px 10px;font-size:13px;min-width:130px}}
.grid-charts{{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:24px}}
@media(max-width:900px){{.grid-charts{{grid-template-columns:1fr}}}}
.panel{{background:var(--ink);border:1px solid var(--hair);border-radius:8px;padding:8px}}
section h2.sh{{font-size:12px;text-transform:uppercase;letter-spacing:1.5px;color:var(--mute);
margin:28px 0 12px;border-bottom:1px solid var(--hair);padding-bottom:8px}}
table.grid{{width:100%;border-collapse:collapse;font-size:12.5px}}
table.grid th{{text-align:left;color:var(--mute);font-weight:600;font-size:11px;text-transform:uppercase;
letter-spacing:.6px;padding:8px 10px;border-bottom:1px solid var(--hair);position:sticky;top:0;background:var(--ink)}}
table.grid td{{padding:7px 10px;border-bottom:1px solid var(--hair);font-family:ui-monospace,Menlo,monospace}}
table.grid tr.manual td:first-child{{box-shadow:inset 3px 0 0 var(--manual)}}
table.grid tr.spof{{background:rgba(214,69,61,.08)}}
.badge{{color:var(--hot);font-size:10px}}
.legend{{display:flex;gap:18px;flex-wrap:wrap;font-size:12px;color:var(--mute);margin-top:10px}}
.legend span b{{color:var(--text)}}
.sw{{display:inline-block;width:10px;height:10px;border-radius:2px;margin-right:6px;vertical-align:middle}}
.tbl-scroll{{max-height:520px;overflow:auto;border:1px solid var(--hair);border-radius:8px}}
.src table{{border-collapse:collapse;font-size:12.5px}} .src td{{padding:4px 10px;font-family:ui-monospace,monospace}}
</style></head><body>
<header><h1>Refinery Outage Dashboard</h1>
<span class="meta">NA products &amp; feedstocks · PADD priority {padd_order} · big-outage ≥ {n(BIG_OUTAGE_BBL)} BBL/d</span></header>
<div class="wrap">
<div class="takeaways"><h2>Key takeaways</h2><ul>{"".join(f"<li>{t}</li>" for t in takeaways)}</ul></div>
<div class="controls">
<label>PADD<select id="fPadd"><option value="">All</option></select></label>
<label>Unit type<select id="fUnit"><option value="">All</option></select></label>
<label>Year<select id="fYear"><option value="">All</option></select></label>
<label>Source<select id="fSource"><option value="">All</option>
<option value="manual">Manual only</option><option value="feed">Feed only</option></select></label>
</div>
<div class="grid-charts">
<div class="panel">{charts['monthly']}</div><div class="panel">{charts['yoy']}</div>
<div class="panel">{charts['padd']}</div><div class="panel">{charts['unit']}</div>
</div>
<section><h2 class="sh">Ranked outages — severity (size × redundancy), then PADD priority</h2>
<div class="tbl-scroll">{ranked_tbl}</div>
<div class="legend">
<span><span class="sw" style="background:var(--manual)"></span>left bar = <b>manual</b> source</span>
<span><span class="sw" style="background:rgba(214,69,61,.4)"></span>red row = <b>single point of failure</b></span>
<span><span class="badge">●</span> = SPOF flag</span>
<span>Mogas BBL/d = offline capacity × (PADD, unit) yield</span></div></section>
<section><h2 class="sh">Back-to-back clusters — same unit type, same PADD, consecutive months</h2>
<div class="tbl-scroll">{b2b_tbl}</div></section>
<section><h2 class="sh">Data sources &amp; provenance</h2>
<div class="src"><table><tbody>{src_rows}</tbody></table></div>
<div class="legend"><span>Manual: <b>{report['n_manual']}</b></span>
<span>Feed: <b>{report['n_feed']}</b></span>
<span>Quality-flagged: <b>{report['n_issue_rows']}</b></span>
<span>Total mogas offline: <b>{n(report['mogas_total'])} BBL/d</b></span></div></section>
</div>
<script>
const rankRows = document.querySelectorAll('section:nth-of-type(1) .grid tbody tr');
const colIndex = {{padd: 1, unit: 2, year: 9}};
function opts(sel,vals){{const s=document.getElementById(sel);
[...new Set(vals)].filter(v=>v!=='').sort().forEach(v=>{{
const o=document.createElement('option');o.value=v;o.textContent=v;s.appendChild(o);}});}}
opts('fPadd',[...rankRows].map(r=>r.children[colIndex.padd].textContent.trim()));
opts('fUnit',[...rankRows].map(r=>r.children[colIndex.unit].textContent.trim()));
opts('fYear',[...rankRows].map(r=>r.children[colIndex.year].textContent.trim().slice(0,4)));
function apply(){{const p=fPadd.value,u=fUnit.value,y=fYear.value,s=fSource.value;
rankRows.forEach(r=>{{const m=r.classList.contains('manual');let show=true;
if(p&&r.children[colIndex.padd].textContent.trim()!==p)show=false;
if(u&&r.children[colIndex.unit].textContent.trim()!==u)show=false;
if(y&&!r.children[colIndex.year].textContent.trim().startsWith(y))show=false;
if(s==='manual'&&!m)show=false;if(s==='feed'&&m)show=false;
r.style.display=show?'':'none';}});}}
['fPadd','fUnit','fYear','fSource'].forEach(id=>document.getElementById(id).addEventListener('change',apply));
</script></body></html>"""

    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    return path


def export_ranked_xlsx(res, path):
    """Write the full ranked-outages frame to an .xlsx next to the HTML."""
    res["ranked"].to_excel(path, index=False, engine="openpyxl")
    return path


# ===========================================================================
# ENTRY POINT
# ===========================================================================

def main():
    ap = argparse.ArgumentParser(description="Refinery outage analyzer")
    ap.add_argument("excel", help="path to the outage .xlsx (Snowflake export)")
    ap.add_argument("--out", default=OUTPUT_HTML, help="output HTML path")
    ap.add_argument("--xlsx", default=None, help="output path for the ranked-outages Excel (defaults to the HTML stem + .xlsx)")
    args = ap.parse_args()

    print(f"Reading {args.excel} ...")
    df, report = load(args.excel)
    print(f" {report['n_rows']} rows | manual {report['n_manual']} feed {report['n_feed']} | quality-flagged {report['n_issue_rows']}")
    print(f" total mogas offline: {report['mogas_total']:,.0f} BBL/d")
    for ut in report["mogas_unmatched_units"]:
        print(f" note: no mogas yield for unit_type '{ut}' (mogas left blank for those rows).")
    for lg, hd in report["missing_optional"]:
        print(f" note: optional column '{lg}' (expected '{hd}') not found.")
    if report["found"]:
        print(" matched columns:")
        for lg, hd in report["found"].items():
            print(f"  - {lg}: {hd}")

    print("Analyzing ...")
    df, res = build_all(df)
    print(f" big outages: {len(res['big'])} | back-to-back: {len(res['back_to_back'])} | SPOF: {len(res['spof'])}")

    out = render(df, res, report, args.out)
    xlsx_path = args.xlsx or (os.path.splitext(out)[0] + ".xlsx")
    export_ranked_xlsx(res, xlsx_path)
    print(f"\nDone -> {out}")
    print(f"Ranked outages -> {xlsx_path}")
    print("Open the HTML in any browser or email it. No server needed.")


if __name__ == "__main__":
    try:
        main()
    except SystemExit as e:
        if isinstance(e.code, str):
            print(e.code)
            sys.exit(1)
        raise
