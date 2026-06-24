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
INPUT_PATH = str(_ROOT / "data" / "rEFINERY oUTAGES.xlsx")
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
        "monthly": monthly,
        "padd_monthly": padd_monthly,
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
  <div class="kpis" id="kpis"></div>

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
    <div class="card full"><h3>Year-over-Year Change in Total Offline</h3><p class="note">YoY % change, total capacity offline.</p><div class="chartbox"><canvas id="cYoy"></canvas></div></div>
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
const ly = M.latest_full_year, sly = DATA.summary[ly];
let topPadd='-',topVal=-1;
DATA.padd_order.forEach(p=>{const v=DATA.padd[p].unplanned[ly]||0; if(v>topVal){topVal=v;topPadd=p;}});
const kpis = [
  ['Total Offline ('+ly+')', fmt(sly.total), 'kbd'],
  ['Unplanned ('+ly+')', fmt(sly.unplanned), 'kbd'],
  ['Unplanned %', pct(sly.unpl), 'of total'],
  ['Distinct Outages', fmt(sly.events), 'FY'+ly],
  ['Top PADD', topPadd, 'by unplanned'],
];
document.getElementById('kpis').innerHTML = kpis.map(k=>
  `<div class="kpi"><div class="lab">${k[0]}</div><div class="val">${k[1]}</div><div class="sub">${k[2]}</div></div>`).join('');

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
