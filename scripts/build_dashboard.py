#!/usr/bin/env python3
"""
build_dashboard.py
Self-contained interactive HTML dashboard for the refinery-outage analysis.

Priority-3 deliverable. Emits a single .html file with the data embedded as JSON
and Chart.js vendored inline, so it opens with no server and (if Chart.js was
fetched at build time) no network. KPI tiles, metric/PADD filters, YoY bars,
seasonality lines, and a live 2027 scenario panel that recomputes exactly like
the workbook's formulas.

Usage:
    python build_dashboard.py                      # uses INPUT_PATH
    python build_dashboard.py path/to/export.xlsx --out outage_dashboard.html
"""
import argparse
import json
import os
import ssl
import sys
import urllib.request

from pathlib import Path

import engine

_ROOT = Path(__file__).resolve().parent.parent          # repo root (scripts/ -> ..)
INPUT_PATH = str(_ROOT / "data" / "Refinery_Outages_Data.xlsx")
OUT_PATH = str(_ROOT / "output" / "outage_dashboard.html")
CHARTJS_URL = "https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"


def fetch_chartjs():
    """Return Chart.js source for inlining, or None to fall back to the CDN."""
    try:
        ctx = ssl.create_default_context()
        ca = os.environ.get("REQUESTS_CA_BUNDLE") or "/root/.ccr/ca-bundle.crt"
        if os.path.exists(ca):
            ctx.load_verify_locations(ca)
        with urllib.request.urlopen(CHARTJS_URL, context=ctx, timeout=30) as r:
            return r.read().decode("utf-8", "replace")
    except Exception as e:
        print(f"  (could not vendor Chart.js: {e}; will reference CDN)")
        return None


def build_data(ctx):
    s = ctx["summary"]
    df = ctx["df"]
    years = [int(y) for y in s.index if 2016 <= int(y) <= 2027]
    ly = max(y for y in s.index if y not in engine.PARTIAL_YEARS and s.loc[y, "Unplanned"] > 0)

    summary = {}
    for y in years:
        summary[y] = {
            "planned": float(s.loc[y, "Planned"]),
            "unplanned": float(s.loc[y, "Unplanned"]),
            "total": float(s.loc[y, "Total"]),
            "events": int(s.loc[y, "Events"]),
            "unpl": (None if y in engine.PLANNED_ONLY_YEARS else float(s.loc[y, "Unpl%"])),
            "yoy_pct": (None if (y == years[0] or s.loc[y, "YoY%"] != s.loc[y, "YoY%"])
                        else float(s.loc[y, "YoY%"])),
        }

    padd = {}
    for key, mat in [("total", ctx["padd_total"]), ("unplanned", ctx["padd_unplanned"]),
                     ("planned", ctx["padd_planned"])]:
        for p in engine.PADD_ORDER:
            padd.setdefault(p, {})[key] = {int(y): float(mat.loc[p, y])
                                           for y in mat.columns if 2016 <= int(y) <= 2027}

    um = ctx["unit_total"]
    units = []
    for u in um.index[:12]:
        units.append({"name": str(u).title(),
                      "total": float(um.loc[u].sum()),
                      "by_year": {int(y): float(um.loc[u, y]) for y in um.columns
                                  if 2016 <= int(y) <= 2027}})

    monthly = {}
    for key, mat in [("total", ctx["monthly_total"]), ("planned", ctx["monthly_planned"]),
                     ("unplanned", ctx["monthly_unplanned"])]:
        monthly[key] = {int(y): [float(mat.loc[y, m]) for m in engine.MONTHS]
                        for y in mat.index if 2018 <= int(y) <= 2027}

    # per-PADD monthly (for the PADD-filtered seasonality view)
    padd_monthly = {}
    for p in engine.PADD_ORDER:
        padd_monthly[p] = {}
        for key in ("total", "planned", "unplanned"):
            mat = ctx["padd_month"][p][key]
            padd_monthly[p][key] = {int(y): [float(mat.loc[y, m]) for m in engine.MONTHS]
                                    for y in mat.index if 2018 <= int(y) <= 2027}

    profiles = {w: [float(v) for v in engine.baseline_profile(df, w).values]
                for w in engine.BASELINE_WINDOWS}
    mp = ctx["monthly_planned"]
    mu = ctx["monthly_unplanned"]
    planned_2027 = [float(mp.loc[2027, m]) if 2027 in mp.index else 0.0 for m in engine.MONTHS]
    actuals = {str(yr): [float(mu.loc[yr, m]) if yr in mu.index else 0.0 for m in engine.MONTHS]
               for yr in (2024, 2025)}

    # forward-looking series: H1 like-for-like, the scenario fan, and the
    # forecast-filled unplanned paths (so 2026/27 carry the forecast, not zero)
    h1 = {int(y): float(v) for y, v in ctx["h1_planned"].items() if 2024 <= int(y) <= 2027}
    fan = ctx["scenario_fan"]
    fan_out = {"conservative": float(fan["Conservative"].sum()),
               "average": float(fan["Average"].sum()),
               "active": float(fan["Active"].sum()),
               "monthly": {k.lower(): [float(fan[k][m]) for m in engine.MONTHS]
                           for k in ("Conservative", "Average", "Active")}}
    cu = ctx["completed_unplanned"]
    completed = {str(y): {"vals": [float(v) for v in cu[y]["vals"]], "fc_from": int(cu[y]["fc_from"])}
                 for y in (2024, 2025, 2026, 2027) if y in cu}

    # market context: $ of refining margin at risk + the gasoline crack overlay
    di = ctx["dollar_impact"]
    dollar_impact = {str(y): {k: float(v) for k, v in d.items() if k != "monthly_unpl"}
                     for y, d in di.items()}
    cmyears = [y for y in (2022, 2023, 2024, 2025, 2026) if di]
    crack_monthly = ({str(y): [float(v) for v in cmv]
                      for y, cmv in engine.crack_matrix(ctx["crack"], cmyears).items()}
                     if ctx["crack"] else {})

    fcc = [{"plant": c["plant"].replace(" Refinery", ""), "padd": c["padd"],
            "year": c["year"], "span": c["span"], "n": c["n"],
            "kbd": round(c["kbd"]), "unpl": c["unpl_share"]}
           for c in ctx["fcc_exxon"][:14]]
    unit_pct = {str(u).title(): float(ctx["unit_share"].loc[u, 2025])
                for u in ctx["unit_share"].index[:8]}

    rb = ctx["range_band"]
    range_band = {k: [float(rb.loc[m, k]) for m in engine.MONTHS] for k in ("min", "max", "avg")}
    _, ym = ctx["monthly_yoy"]
    yoy_month = {str(y): [None if ym.loc[y, m] != ym.loc[y, m] else float(ym.loc[y, m])
                          for m in engine.MONTHS]
                 for y in (2024, 2025) if y in ym.index}

    def _shop(s):
        s = str(s).title()
        for w in ("Corporation", "Company", "Incorporated", "Petroleum", "Refining",
                  " Llc", " Lp", "North America", "Products", "Energy", " Inc"):
            s = s.replace(w, "")
        return " ".join(s.split())[:22]

    def _dt(x):
        try:
            return f"{x.month}/{x.day}/{str(x.year)[2:]}"
        except Exception:
            return ""
    ta = {}
    for p in engine.PADD_ORDER:
        tdf = ctx["ta_schedule"][p]
        ta[p] = [{"op": _shop(rw["operator"]), "plant": str(rw["plant"])[:34],
                  "unit": str(rw["unit_cat"]).title(), "kbd": round(float(rw["kbd"]), 1),
                  "pct": float(rw["pct_padd"]), "start": _dt(rw["start"]), "end": _dt(rw["end"])}
                 for _, rw in tdf.head(18).iterrows()]

    d = ctx["diag"]
    return {
        "meta": {"rows": d["rows"], "years": list(d["years"]),
                 "events": d["events_distinct"], "latest_full_year": int(ly)},
        "months": engine.MONTHS,
        "year_list": years,
        "padd_order": engine.PADD_ORDER,
        "summary": summary,
        "padd": padd,
        "units": units,
        "unit_pct": unit_pct,
        "fcc": fcc,
        "range_band": range_band,
        "yoy_month": yoy_month,
        "ta": ta,
        "monthly": monthly,
        "padd_monthly": padd_monthly,
        "h1_planned": h1,
        "fan": fan_out,
        "completed_unplanned": completed,
        "dollar_impact": dollar_impact,
        "crack_monthly": crack_monthly,
        "scenario": {
            "windows": list(engine.BASELINE_WINDOWS.keys()),
            "default_window": engine.DEFAULT_WINDOW,
            "profiles": profiles,
            "planned_2027": planned_2027,
            "actuals": actuals,
            "padd_share": {p: float(ctx["padd_share"][p]) for p in engine.PADD_ORDER},
        },
    }


HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Refinery Outage Analytics</title>
<style>
:root{
  --navy:#1F3864; --blue:#2E5496; --red:#C00000; --gold:#BF9000; --green:#548235;
  --orange:#ED7D31; --ltblue:#D6E0F0; --ltgray:#F2F2F2; --gray:#595959; --line:#E3E8F0;
}
*{box-sizing:border-box}
body{margin:0;font-family:Arial,"Liberation Sans","Helvetica Neue",sans-serif;color:#23272e;background:#EEF1F6}
header{background:var(--navy);color:#fff;padding:16px 28px;border-bottom:3px solid var(--gold)}
header h1{margin:0;font-size:22px;letter-spacing:.3px}
header .sub{color:var(--ltblue);font-size:12.5px;margin-top:3px}
.wrap{max-width:1320px;margin:0 auto;padding:18px}
.kpis{display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-bottom:16px}
.kpi{background:#fff;border:1px solid var(--line);border-radius:8px;overflow:hidden;box-shadow:0 1px 2px rgba(0,0,0,.04)}
.kpi .lab{background:var(--blue);color:#fff;font-size:11.5px;font-weight:bold;text-align:center;padding:6px 4px}
.kpi .val{background:var(--ltblue);color:var(--navy);font-size:26px;font-weight:bold;text-align:center;padding:12px 4px}
.kpi .sub{color:var(--gray);font-size:10.5px;font-style:italic;text-align:center;padding:4px;background:var(--ltblue)}
.kpi .sub.up{color:#1d6f33;font-style:normal;font-weight:bold}
.kpi .sub.down{color:#b3261e;font-style:normal;font-weight:bold}
.outlook{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:12px;margin-bottom:16px}
.ocard{background:#fff;border:1px solid var(--line);border-left:4px solid var(--gold);border-radius:8px;padding:11px 15px;box-shadow:0 1px 2px rgba(0,0,0,.04)}
.ocard .ot{font-size:11.5px;font-weight:bold;color:var(--navy);text-transform:uppercase;letter-spacing:.3px}
.ocard .ov{font-size:21px;font-weight:bold;color:var(--navy);margin-top:3px}
.ocard .od{font-size:12px;margin-top:2px}
.ocard .od.up{color:#1d6f33;font-weight:bold}.ocard .od.down{color:#b3261e;font-weight:bold}
.ocard .on{font-size:10.5px;color:var(--gray);font-style:italic;margin-top:4px}
@media(max-width:980px){.outlook{grid-template-columns:1fr}}
.controls{background:#fff;border:1px solid var(--line);border-radius:8px;padding:12px 16px;margin-bottom:16px;display:flex;gap:22px;flex-wrap:wrap;align-items:center}
.controls label{font-size:12px;font-weight:bold;color:var(--navy);margin-right:6px}
.controls select{font-size:13px;padding:5px 8px;border:1px solid #c8d0dc;border-radius:5px;background:#fff}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:16px}
.card{background:#fff;border:1px solid var(--line);border-radius:8px;padding:14px 16px 18px;box-shadow:0 1px 2px rgba(0,0,0,.04)}
.card h3{margin:0 0 2px;font-size:15px;color:var(--navy)}
.card .note{font-size:11.5px;color:var(--gray);margin:0 0 10px}
.card.full{grid-column:1 / -1}
.chartbox{position:relative;height:330px}
.chartbox.tall{height:360px}
.scenario{background:#fff;border:1px solid var(--line);border-radius:8px;padding:0;margin:16px 0;overflow:hidden}
.scenario .hd{background:var(--red);color:#fff;font-weight:bold;padding:10px 16px;font-size:15px}
.scenario .body{display:grid;grid-template-columns:300px 1fr;gap:0}
.inputs{background:#FFF8E6;border-right:1px solid #EAD9A6;padding:16px}
.inputs .row{margin-bottom:14px}
.inputs .row label{display:block;font-size:12px;font-weight:bold;color:var(--navy);margin-bottom:5px}
.inputs select,.inputs input[type=number]{width:100%;padding:6px 8px;border:1px solid #d8c98f;border-radius:5px;font-size:13px;background:#fff;color:#0000FF;font-weight:bold}
.inputs input[type=range]{width:100%}
.inputs .rngval{float:right;color:#0000FF;font-weight:bold}
.outs{padding:16px}
.outrow{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:12px}
.outbox{background:var(--ltgray);border-radius:6px;padding:10px;text-align:center}
.outbox .l{font-size:11px;color:var(--gray);font-weight:bold}
.outbox .v{font-size:22px;color:var(--navy);font-weight:bold}
table.alloc{width:100%;border-collapse:collapse;font-size:12.5px;margin-top:6px}
table.alloc th{background:var(--navy);color:#fff;padding:5px 8px;text-align:right}
table.alloc th:first-child{text-align:left}
table.alloc td{padding:4px 8px;border-bottom:1px solid var(--line);text-align:right}
table.alloc td:first-child{text-align:left}
.foot{color:var(--gray);font-size:11px;text-align:center;padding:18px 0 26px}
.tag{display:inline-block;background:#FBEAEA;color:var(--red);font-size:10.5px;font-weight:bold;padding:2px 7px;border-radius:4px;margin-left:8px}
@media(max-width:980px){.grid{grid-template-columns:1fr}.kpis{grid-template-columns:repeat(2,1fr)}.scenario .body{grid-template-columns:1fr}}
</style>
</head>
<body>
<header>
  <h1>Refinery Outage Analytics <span class="tag" id="vintageTag"></span></h1>
  <div class="sub" id="subline"></div>
</header>
<div class="wrap">
  <div class="controls" style="margin-bottom:12px">
    <div><label>Focus year (KPIs vs prior year)</label>
      <select id="focusYear"></select></div>
    <div style="color:var(--gray);font-size:11.5px;font-style:italic">KPI tiles below show the focus year and its change vs the previous year - pick any year.</div>
  </div>
  <div class="kpis" id="kpis"></div>

  <div class="outlook" id="outlook"></div>

  <div class="controls">
    <div><label>Metric</label>
      <select id="metric"><option value="total">Total offline</option>
      <option value="unplanned" selected>Unplanned</option><option value="planned">Planned</option></select></div>
    <div><label>PADD (seasonality)</label>
      <select id="paddFilter"><option value="ALL">All US</option></select></div>
    <div style="color:var(--gray);font-size:11.5px;font-style:italic">Charts below respond to these filters; the 2027 panel is its own live model.</div>
  </div>

  <div class="grid">
    <div class="card"><h3>Capacity Offline by Year</h3><p class="note">Planned + Unplanned, kbd. 2026/27 partial.</p><div class="chartbox"><canvas id="cAnnual"></canvas></div></div>
    <div class="card"><h3 id="paddTitle">Unplanned by PADD</h3><p class="note">Selected metric across recent years, kbd.</p><div class="chartbox"><canvas id="cPadd"></canvas></div></div>
    <div class="card"><h3 id="seasTitle">Seasonality</h3><p class="note">Selected metric by month, one line per year.</p><div class="chartbox"><canvas id="cSeason"></canvas></div></div>
    <div class="card"><h3>Top Unit Categories</h3><p class="note">Capacity offline, all years, kbd.</p><div class="chartbox"><canvas id="cUnits"></canvas></div></div>
    <div class="card"><h3>Unplanned Seasonality &amp; Range Band</h3><p class="note">Grey band = 2022-25 monthly min-max; 2026 dashed = actual + forecast tail; 2027 dotted = Average scenario (not zero).</p><div class="chartbox"><canvas id="cBand"></canvas></div></div>
    <div class="card"><h3>Unplanned &mdash; YoY % Change by Month</h3><p class="note">Percent difference in each month vs the prior year.</p><div class="chartbox"><canvas id="cYoyMonth"></canvas></div></div>
    <div class="card full"><h3>$ Gross Margin at Risk &amp; Gasoline Crack</h3><p class="note">Offline capacity valued at the gasoline crack: unplanned/planned $MM per year (left axis, bars) vs avg crack $/bbl (right axis, line). 2020-21 inflated by COVID/Uri spikes; crack = EIA NYH&minus;WTI, Bloomberg-overwritable.</p><div class="chartbox"><canvas id="cDollar"></canvas></div></div>
    <div class="card full"><h3>Year-over-Year Change in Total Offline</h3><p class="note">YoY % change, total capacity offline.</p><div class="chartbox"><canvas id="cYoy"></canvas></div></div>
    <div class="card full"><h3>2026 Planned Turnaround Schedule
        <select id="taPadd" style="font-size:12px;padding:3px 6px;margin-left:8px;border:1px solid #c8d0dc;border-radius:5px"></select></h3>
      <p class="note">Event-level planned outages with offline capacity (kbd), % of PADD, and dates. Top by size.</p>
      <table class="alloc" id="taTbl"><thead><tr><th>Operator</th><th>Refinery</th><th>Unit</th><th>Offline (kbd)</th><th>% PADD</th><th>Start</th><th>End</th></tr></thead><tbody></tbody></table>
    </div>
    <div class="card full"><h3>Back-to-Back FCC Outages &mdash; ExxonMobil <span class="tag" id="fccTag"></span></h3>
      <p class="note">Consecutive-month FCC (cat cracker) runs at the same plant &mdash; the clustered signal month-level external trackers miss (2020 excluded).</p>
      <table class="alloc" id="fccTbl"><thead><tr><th>Refinery</th><th>PADD</th><th>Year</th><th>Span</th><th>Months</th><th>kbd</th><th>Unpl %</th></tr></thead><tbody></tbody></table>
    </div>
  </div>

  <div class="scenario">
    <div class="hd">2027 Unplanned Scenario &mdash; live model</div>
    <div class="body">
      <div class="inputs">
        <div class="row"><label>Baseline window</label><select id="scWindow"></select></div>
        <div class="row"><label>Production growth <span class="rngval" id="scGrowthV">0%</span></label>
          <input type="range" id="scGrowth" min="-10" max="15" step="1" value="0"></div>
        <div class="row"><label>Unplanned rate multiplier <span class="rngval" id="scMultV">1.00x</span></label>
          <input type="range" id="scMult" min="0.5" max="2" step="0.05" value="1"></div>
        <div class="row"><label>One-off event (kbd)</label><input type="number" id="scOneoff" value="0" step="50"></div>
        <div class="row"><label>Stress month</label><select id="scStress"></select></div>
      </div>
      <div class="outs">
        <div class="outrow">
          <div class="outbox"><div class="l">2027 Unplanned forecast</div><div class="v" id="oUnpl">-</div></div>
          <div class="outbox"><div class="l">2027 Planned (booked)</div><div class="v" id="oPlan">-</div></div>
          <div class="outbox"><div class="l">Implied total offline</div><div class="v" id="oTot">-</div></div>
        </div>
        <div style="display:grid;grid-template-columns:1.5fr 1fr;gap:14px">
          <div class="chartbox" style="height:260px"><canvas id="cScenario"></canvas></div>
          <div><table class="alloc" id="allocTbl"><thead><tr><th>PADD</th><th>Share</th><th>Scenario kbd</th></tr></thead><tbody></tbody></table></div>
        </div>
      </div>
    </div>
  </div>

  <div class="foot" id="foot"></div>
</div>

<script>__CHARTJS__</script>
<script>
const DATA = __DATA__;
const C = {navy:'#1F3864',blue:'#2E5496',red:'#C00000',gold:'#BF9000',green:'#548235',orange:'#ED7D31',gray:'#808080'};
const YRCOL = {2018:'#9DB0CE',2019:'#7E97C3',2020:'#cccccc',2021:'#bbbbbb',2022:'#808080',2023:'#548235',2024:'#2E5496',2025:'#C00000',2026:'#BF9000',2027:'#7030A0'};
const fmt = x => x==null ? 'n/a' : Math.round(x).toLocaleString();
const pct = x => x==null ? 'n/a' : (x*100).toFixed(0)+'%';
Chart.defaults.font.family = 'Arial, Liberation Sans, sans-serif';
Chart.defaults.font.size = 11;
Chart.defaults.plugins.legend.labels.boxWidth = 12;

// ---- header + KPIs ----
const M = DATA.meta;
document.getElementById('vintageTag').textContent = M.years[0]+'-'+M.years[1]+' | '+M.rows.toLocaleString()+' rows';
document.getElementById('subline').textContent =
  'Capacity offline (kbd), all units. '+M.events.toLocaleString()+' distinct outages. Latest full year '+M.latest_full_year+'. 2026/27 partial; 2027 planned-only.';
// ---- dynamic focus-year KPIs (selected year vs the prior year) ----
const ly = M.latest_full_year;
const PARTIAL = {2026:1, 2027:1};
const fySel = document.getElementById('focusYear');
// default focus = 2026 (forward-looking) when present, else latest full year
DATA.year_list.filter(y=>y>=2018).forEach(y=>{
  const o=document.createElement('option');o.value=y;o.textContent=y+(PARTIAL[y]?'  (partial)':'');
  if(y===(DATA.summary[2026]?2026:ly))o.selected=true;fySel.appendChild(o);});

function deltaSub(cur, prev){
  if(prev==null||prev===0||cur==null) return ['','',''];
  const d=cur/prev-1, cls=d>=0?'up':'down', arrowless=(d>=0?'+':'')+(d*100).toFixed(0)+'%';
  return [arrowless+' vs prior yr', cls, ''];
}
function renderKPIs(){
  const fy=+fySel.value, py=fy-1, S=DATA.summary[fy], P=DATA.summary[py];
  let topPadd='-',topVal=-1;
  DATA.padd_order.forEach(p=>{const v=(DATA.padd[p].unplanned[fy]||0); if(v>topVal){topVal=v;topPadd=p;}});
  // for the planned-only future year, unplanned isn't an actual -> label the forecast
  const unplVal = (fy===2027) ? ('~'+fmt(DATA.fan.average)) : fmt(S.unplanned);
  const unplSub = (fy===2027) ? ['Average scenario','',''] : deltaSub(S.unplanned, P&&P.unplanned);
  const cards = [
    ['Total Offline ('+fy+')', (fy===2027?'~':'')+fmt(S.total), deltaSub(S.total, P&&P.total)],
    ['Unplanned ('+fy+')', unplVal, unplSub],
    ['Planned ('+fy+')', fmt(S.planned), deltaSub(S.planned, P&&P.planned)],
    ['Distinct Outages', fmt(S.events), deltaSub(S.events, P&&P.events)],
    ['Top PADD (unpl)', topPadd, ['vs '+py+' prior yr','','']],
  ];
  document.getElementById('kpis').innerHTML = cards.map(k=>{
    const [txt,cls]=k[2];
    return `<div class="kpi"><div class="lab">${k[0]}</div><div class="val">${k[1]}</div>`+
           `<div class="sub ${cls}">${txt||'kbd'}</div></div>`;}).join('');
}
fySel.addEventListener('change',renderKPIs); renderKPIs();

// ---- outlook strip: 2026 vs 2025, H1'27 vs H1'26 planned, 2027 forecast ----
(function(){
  const s26=DATA.summary[2026], s25=DATA.summary[2025];
  const h1=DATA.h1_planned, fan=DATA.fan, plan27=DATA.summary[2027]?DATA.summary[2027].planned:0;
  function card(t,v,d,dcls,note){
    return `<div class="ocard"><div class="ot">${t}</div><div class="ov">${v}</div>`+
           (d?`<div class="od ${dcls}">${d}</div>`:'')+`<div class="on">${note}</div></div>`;}
  const cards=[];
  if(s26&&s25){
    const dt=s26.total/s25.total-1, du=s26.unplanned/s25.unplanned-1;
    cards.push(card('2026 vs 2025 - total offline', fmt(s26.total)+' kbd',
      (dt>=0?'+':'')+(dt*100).toFixed(0)+'% vs 2025', dt>=0?'up':'down',
      'Unplanned '+(du>=0?'+':'')+(du*100).toFixed(0)+'% ('+fmt(s26.unplanned)+' kbd). 2026 partial.'));
  }
  if(h1[2026]&&h1[2027]){
    const dh=h1[2027]/h1[2026]-1;
    cards.push(card('H1 2027 vs H1 2026 - planned', fmt(h1[2027])+' kbd',
      (dh>=0?'+':'')+(dh*100).toFixed(0)+'% vs H1 2026', dh>=0?'up':'down',
      'Like-for-like (Jan-Jun); H1 2026 = '+fmt(h1[2026])+' kbd. The clean 2027 read.'));
  }
  cards.push(card('2027 unplanned - forecast', '~'+fmt(fan.average)+' kbd',
    'range ~'+fmt(fan.conservative)+' - '+fmt(fan.active), 'up',
    'Conservative/Average/Active. Implied total ~'+fmt(fan.average+plan27)+' kbd with booked planned.'));
  const d25=(DATA.dollar_impact||{})['2025'];
  if(d25){
    cards.push(card('2025 unplanned - $ at risk', '~$'+(d25.unplanned/1000).toFixed(1)+'bn',
      'avg crack $'+Math.round(d25.crack_avg)+'/bbl', 'down',
      'Gross gasoline-refining margin on unplanned offline capacity. Planned ~$'+(d25.planned/1000).toFixed(1)+'bn.'));
  }
  document.getElementById('outlook').innerHTML=cards.join('');
})();

// ---- $ gross margin at risk + crack overlay ----
(function(){
  const DI=DATA.dollar_impact||{};
  const yrs=Object.keys(DI).filter(y=>+y>=2022).sort();
  if(!yrs.length){return;}
  new Chart(document.getElementById('cDollar'),{
    data:{labels:yrs.map(y=>+y>=2026?y+'*':y),datasets:[
      {type:'bar',label:'Unplanned $ at risk ($MM)',backgroundColor:C.navy,yAxisID:'y',data:yrs.map(y=>DI[y].unplanned)},
      {type:'bar',label:'Planned $ ($MM)',backgroundColor:C.gold,yAxisID:'y',data:yrs.map(y=>DI[y].planned)},
      {type:'line',label:'Avg gasoline crack ($/bbl)',borderColor:C.red,backgroundColor:C.red,borderWidth:2.6,pointRadius:3,tension:.25,yAxisID:'y1',data:yrs.map(y=>DI[y].crack_avg)}]},
    options:{responsive:true,maintainAspectRatio:false,scales:{
      x:{grid:{display:false}},
      y:{position:'left',title:{display:true,text:'$MM / yr'},ticks:{callback:v=>v.toLocaleString()}},
      y1:{position:'right',grid:{drawOnChartArea:false},title:{display:true,text:'crack $/bbl'}}},
      plugins:{legend:{position:'bottom'}}}});
})();

// ---- controls ----
const yearsAll = DATA.year_list;
const recent = yearsAll.filter(y=>y>=2022 && y<=2025);
const pf = document.getElementById('paddFilter');
DATA.padd_order.forEach(p=>{const o=document.createElement('option');o.value=p;o.textContent=p;pf.appendChild(o);});

// ---- annual stacked ----
const cAnnual = new Chart(document.getElementById('cAnnual'),{type:'bar',
  data:{labels:yearsAll.map(y=>DATA.summary[y]&& (y>=2026)?y+'*':''+y),
    datasets:[
      {label:'Planned',backgroundColor:C.navy,data:yearsAll.map(y=>DATA.summary[y].planned)},
      {label:'Unplanned',backgroundColor:C.red,data:yearsAll.map(y=>DATA.summary[y].unplanned)}]},
  options:{responsive:true,maintainAspectRatio:false,scales:{x:{stacked:true,grid:{display:false}},
    y:{stacked:true,ticks:{callback:v=>v.toLocaleString()}}},plugins:{legend:{position:'bottom'}}}});

// ---- PADD chart (responds to metric) ----
let cPadd;
function renderPadd(){
  const metric = document.getElementById('metric').value;
  document.getElementById('paddTitle').textContent =
    ({total:'Total',unplanned:'Unplanned',planned:'Planned'}[metric])+' by PADD';
  const cols=[C.blue,C.gold,C.green,C.red];
  const ds = recent.map((y,i)=>({label:''+y,backgroundColor:cols[i%4],
    data:DATA.padd_order.map(p=>DATA.padd[p][metric][y]||0)}));
  const cfg={type:'bar',data:{labels:DATA.padd_order.map(p=>p.replace('PADD ','P')),datasets:ds},
    options:{responsive:true,maintainAspectRatio:false,scales:{x:{grid:{display:false}},
      y:{ticks:{callback:v=>v.toLocaleString()}}},plugins:{legend:{position:'bottom'}}}};
  if(cPadd){cPadd.data=cfg.data;cPadd.update();}else{cPadd=new Chart(document.getElementById('cPadd'),cfg);}
}

// ---- seasonality (responds to metric + PADD) ----
let cSeason;
function renderSeason(){
  const metric=document.getElementById('metric').value, padd=document.getElementById('paddFilter').value;
  const src = padd==='ALL'?DATA.monthly[metric]:DATA.padd_monthly[padd][metric];
  document.getElementById('seasTitle').textContent='Seasonality - '+
    ({total:'Total',unplanned:'Unplanned',planned:'Planned'}[metric])+' ('+(padd==='ALL'?'All US':padd)+')';
  const yrs=[2022,2023,2024,2025,2026].filter(y=>src[y]);
  const ds=yrs.map(y=>({label:y>=2026?y+'*':''+y,borderColor:YRCOL[y]||C.navy,backgroundColor:YRCOL[y]||C.navy,
    borderWidth:y===2025?3:2,borderDash:y===2026?[6,4]:[],pointRadius:2,tension:.25,data:src[y],fill:false}));
  // for All-US unplanned, carry 2026's forecast tail and add the 2027 Average scenario (not zero)
  if(metric==='unplanned' && padd==='ALL'){
    const d26=ds.find(d=>d.label==='2026*'); if(d26)d26.data=DATA.completed_unplanned['2026'].vals;
    ds.push({label:'2027 (Avg scenario)',borderColor:'#7030A0',backgroundColor:'#7030A0',
      borderWidth:2.2,borderDash:[2,3],pointRadius:0,tension:.25,data:DATA.fan.monthly.average,fill:false});
  }
  const cfg={type:'line',data:{labels:DATA.months,datasets:ds},
    options:{responsive:true,maintainAspectRatio:false,scales:{x:{grid:{display:false}},
      y:{ticks:{callback:v=>v.toLocaleString()}}},plugins:{legend:{position:'bottom'}}}};
  if(cSeason){cSeason.data=cfg.data;cSeason.update();}else{cSeason=new Chart(document.getElementById('cSeason'),cfg);}
}

// ---- units ----
new Chart(document.getElementById('cUnits'),{type:'bar',
  data:{labels:DATA.units.map(u=>u.name),datasets:[{label:'kbd',backgroundColor:C.gold,data:DATA.units.map(u=>u.total)}]},
  options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,
    scales:{x:{ticks:{callback:v=>v.toLocaleString()}},y:{grid:{display:false}}},
    plugins:{legend:{display:false}}}});

// ---- YoY ----
new Chart(document.getElementById('cYoy'),{type:'bar',
  data:{labels:yearsAll.map(y=>y>=2026?y+'*':''+y),
    datasets:[{label:'YoY % total',data:yearsAll.map(y=>DATA.summary[y].yoy_pct==null?null:DATA.summary[y].yoy_pct*100),
      backgroundColor:yearsAll.map(y=>{const v=DATA.summary[y].yoy_pct;return v==null?'#ccc':(v>=0?C.green:C.red);})}]},
  options:{responsive:true,maintainAspectRatio:false,scales:{x:{grid:{display:false}},
    y:{ticks:{callback:v=>v+'%'}}},plugins:{legend:{display:false}}}});

// ---- seasonality range band ----
const RB = DATA.range_band;
new Chart(document.getElementById('cBand'),{type:'line',
  data:{labels:DATA.months,datasets:[
    {label:'Range max',data:RB.max,borderColor:'rgba(0,0,0,0)',backgroundColor:'rgba(150,150,150,0.25)',pointRadius:0,fill:'+1',tension:.3},
    {label:'Range min',data:RB.min,borderColor:'rgba(0,0,0,0)',backgroundColor:'rgba(150,150,150,0.25)',pointRadius:0,fill:false,tension:.3},
    {label:'5-yr avg',data:RB.avg,borderColor:C.gray,borderDash:[3,3],borderWidth:1.4,pointRadius:0,tension:.3},
    {label:'2024',data:DATA.monthly.unplanned[2024],borderColor:C.blue,borderWidth:1.8,pointRadius:2,tension:.25},
    {label:'2025',data:DATA.monthly.unplanned[2025],borderColor:C.red,borderWidth:2.6,pointRadius:2,tension:.25},
    {label:'2026 (+forecast tail)',data:DATA.completed_unplanned['2026'].vals,borderColor:C.gold,borderWidth:2.2,borderDash:[6,4],pointRadius:2,tension:.25},
    {label:'2027 (Avg scenario)',data:DATA.fan.monthly.average,borderColor:'#7030A0',borderWidth:2.2,borderDash:[2,3],pointRadius:0,tension:.25},
  ]},
  options:{responsive:true,maintainAspectRatio:false,scales:{x:{grid:{display:false}},
    y:{ticks:{callback:v=>v.toLocaleString()}}},
    plugins:{legend:{position:'bottom',labels:{filter:i=>!i.text.startsWith('Range')}}}}});

// ---- monthly YoY% ----
new Chart(document.getElementById('cYoyMonth'),{type:'bar',
  data:{labels:DATA.months,datasets:Object.keys(DATA.yoy_month).map((y,i)=>({
    label:y,backgroundColor:i?C.red:C.blue,
    data:DATA.yoy_month[y].map(v=>v==null?null:v*100)}))},
  options:{responsive:true,maintainAspectRatio:false,scales:{x:{grid:{display:false}},
    y:{ticks:{callback:v=>v+'%'}}},plugins:{legend:{position:'bottom'}}}});

// ---- TA schedule ----
const taSel=document.getElementById('taPadd');
DATA.padd_order.forEach((p,i)=>{const o=document.createElement('option');o.value=p;o.textContent=p+' ('+(DATA.ta[p]||[]).length+')';if(p==='PADD 3')o.selected=true;taSel.appendChild(o);});
function renderTA(){
  const rows=DATA.ta[taSel.value]||[];
  document.querySelector('#taTbl tbody').innerHTML=rows.map(r=>
    `<tr><td>${r.op}</td><td>${r.plant}</td><td>${r.unit}</td><td>${r.kbd.toFixed(1)}</td>`+
    `<td>${(r.pct*100).toFixed(1)}%</td><td>${r.start}</td><td>${r.end}</td></tr>`).join('')
    || '<tr><td colspan=7 style="text-align:center;color:#999">No planned TAs</td></tr>';
}
taSel.addEventListener('change',renderTA); renderTA();

// ---- back-to-back FCC clusters (ExxonMobil) ----
document.getElementById('fccTag').textContent = DATA.fcc.length + ' runs';
document.querySelector('#fccTbl tbody').innerHTML = DATA.fcc.map(c=>
  `<tr><td>${c.plant}</td><td>${c.padd}</td><td>${c.year}</td><td>${c.span}</td>`+
  `<td>${c.n}</td><td>${fmt(c.kbd)}</td><td>${pct(c.unpl)}</td></tr>`).join('');

// ---- scenario (live) ----
const S = DATA.scenario;
const swin=document.getElementById('scWindow'), sstr=document.getElementById('scStress');
S.windows.forEach(w=>{const o=document.createElement('option');o.value=w;o.textContent=w;if(w===S.default_window)o.selected=true;swin.appendChild(o);});
DATA.months.forEach((m,i)=>{const o=document.createElement('option');o.value=i;o.textContent=m;if(m==='Sep')o.selected=true;sstr.appendChild(o);});
let cScenario;
function renderScenario(){
  const w=swin.value, g=(+document.getElementById('scGrowth').value)/100,
    mult=+document.getElementById('scMult').value, oneoff=+document.getElementById('scOneoff').value||0,
    sm=+sstr.value;
  document.getElementById('scGrowthV').textContent=(g*100).toFixed(0)+'%';
  document.getElementById('scMultV').textContent=mult.toFixed(2)+'x';
  const base=S.profiles[w];
  const fc=base.map((b,i)=>b*(1+g)*mult+(i===sm?oneoff:0));
  const unpl=fc.reduce((a,b)=>a+b,0), plan=S.planned_2027.reduce((a,b)=>a+b,0);
  document.getElementById('oUnpl').textContent=fmt(unpl);
  document.getElementById('oPlan').textContent=fmt(plan);
  document.getElementById('oTot').textContent=fmt(unpl+plan);
  const ds=[
    {label:'2027 Scenario',borderColor:C.red,backgroundColor:C.red,borderWidth:3,pointRadius:2,tension:.2,data:fc,fill:false},
    {label:'2027 Planned',borderColor:C.gold,backgroundColor:C.gold,borderWidth:2,borderDash:[6,4],pointRadius:2,tension:.2,data:S.planned_2027,fill:false},
    {label:'2025 Unplanned',borderColor:C.blue,backgroundColor:C.blue,borderWidth:1.5,pointRadius:0,tension:.2,data:S.actuals['2025'],fill:false},
    {label:'2024 Unplanned',borderColor:C.gray,backgroundColor:C.gray,borderWidth:1.5,pointRadius:0,tension:.2,data:S.actuals['2024'],fill:false}];
  const cfg={type:'line',data:{labels:DATA.months,datasets:ds},
    options:{responsive:true,maintainAspectRatio:false,scales:{x:{grid:{display:false}},
      y:{ticks:{callback:v=>v.toLocaleString()}}},plugins:{legend:{position:'bottom'}}}};
  if(cScenario){cScenario.data=cfg.data;cScenario.update();}else{cScenario=new Chart(document.getElementById('cScenario'),cfg);}
  // PADD allocation
  const tb=document.querySelector('#allocTbl tbody');
  tb.innerHTML=DATA.padd_order.map(p=>{const sh=S.padd_share[p];
    return `<tr><td>${p}</td><td>${pct(sh)}</td><td>${fmt(unpl*sh)}</td></tr>`;}).join('');
}

['metric'].forEach(id=>document.getElementById(id).addEventListener('change',()=>{renderPadd();renderSeason();}));
document.getElementById('paddFilter').addEventListener('change',renderSeason);
['scWindow','scGrowth','scMult','scOneoff','scStress'].forEach(id=>{
  document.getElementById(id).addEventListener('input',renderScenario);
  document.getElementById(id).addEventListener('change',renderScenario);});

renderPadd();renderSeason();renderScenario();
document.getElementById('foot').textContent =
  'Generated from the Snowflake outage export. Primary metric: CAP_OFFLINE_ADJUSTED_KBD. '+
  'UNKNOWN folds into UNPLANNED. 2027 is planned-only; unplanned-2027 is a scenario. 2020-21 excluded from baselines.';
</script>
</body>
</html>"""


def main():
    ap = argparse.ArgumentParser(description="Refinery-outage HTML dashboard")
    ap.add_argument("excel", nargs="?", default=INPUT_PATH, help="path to the outage .xlsx export")
    ap.add_argument("--out", default=OUT_PATH, help="output .html path")
    args = ap.parse_args()

    print(f"Loading {args.excel.strip()} ...")
    ctx = engine.build_context(args.excel)
    data = build_data(ctx)
    print("Vendoring Chart.js ...")
    cjs = fetch_chartjs()
    if cjs is None:
        # fall back: load from CDN at view time
        cjs = f'document.write(unescape("%3Cscript src=\'{CHARTJS_URL}\'%3E%3C/script%3E"));'

    html = HTML.replace("__CHARTJS__", cjs).replace(
        "__DATA__", json.dumps(data, separators=(",", ":")))
    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        f.write(html)
    size = os.path.getsize(args.out) / 1024
    print(f"Building dashboard -> {args.out}  ({size:.0f} KB)")
    print("Done.")


if __name__ == "__main__":
    sys.exit(main())
