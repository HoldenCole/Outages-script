"""
charts.py
Matplotlib renderers for the slide deck (and any static-image needs).

Each function takes the engine.build_context() bundle and an output path, writes
a high-resolution PNG styled to the desk palette, and returns the path. Keeping
all static-chart styling here means the deck stays consistent and a refresh only
re-renders.
"""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
import numpy as np

import engine

# palette (mirrors the workbook)
NAVY = "#1F3864"
BLUE = "#2E5496"
RED = "#C00000"
GOLD = "#BF9000"
GREEN = "#548235"
ORANGE = "#ED7D31"
GRAY = "#808080"
LT = "#D6E0F0"

MONTHS = engine.MONTHS
PADDS = engine.PADD_ORDER

plt.rcParams.update({
    "font.family": ["Arial", "Liberation Sans", "DejaVu Sans"],
    "font.size": 11,
    "axes.edgecolor": "#BFBFBF",
    "axes.linewidth": 0.8,
    "axes.titlesize": 14,
    "axes.titleweight": "bold",
    "axes.titlecolor": NAVY,
    "figure.facecolor": "white",
    "savefig.facecolor": "white",
    "axes.grid": True,
    "grid.color": "#E8E8E8",
    "grid.linewidth": 0.7,
})

_thousands = FuncFormatter(lambda x, _: f"{x:,.0f}")


def _clean(ax, ygrid=True):
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(axis="y" if ygrid else "x", zorder=0)
    ax.grid(axis="x" if ygrid else "y", visible=False)
    ax.tick_params(length=0)


def _save(fig, path):
    fig.tight_layout()
    fig.savefig(path, dpi=200, bbox_inches="tight")
    plt.close(fig)
    return path


# --------------------------------------------------------------------------- charts
def annual_stack(ctx, path):
    s = ctx["summary"]
    years = [y for y in s.index if 2016 <= y <= 2027]
    planned = [s.loc[y, "Planned"] for y in years]
    unplanned = [s.loc[y, "Unplanned"] for y in years]
    fig, ax = plt.subplots(figsize=(10, 4.6))
    x = np.arange(len(years))
    ax.bar(x, planned, color=NAVY, label="Planned", zorder=3)
    ax.bar(x, unplanned, bottom=planned, color=RED, label="Unplanned", zorder=3)
    ax.set_xticks(x)
    ax.set_xticklabels([f"{y}*" if y in engine.PARTIAL_YEARS else str(y) for y in years])
    ax.yaxis.set_major_formatter(_thousands)
    ax.set_ylabel("kbd")
    ax.set_title("Capacity Offline by Year - Planned + Unplanned (kbd)")
    ax.legend(loc="upper right", frameon=False, ncol=2)
    _clean(ax)
    return _save(fig, path)


def padd_clustered(ctx, path):
    m = ctx["padd_unplanned"]
    years = [y for y in [2022, 2023, 2024, 2025] if y in m.columns]
    fig, ax = plt.subplots(figsize=(6.4, 4.3))
    x = np.arange(len(PADDS))
    w = 0.8 / len(years)
    cols = [BLUE, GOLD, GREEN, RED]
    for i, y in enumerate(years):
        ax.bar(x + i * w, [m.loc[p, y] for p in PADDS], w, color=cols[i % 4],
               label=str(y), zorder=3)
    ax.set_xticks(x + 0.4 - w / 2)
    ax.set_xticklabels([p.replace("PADD ", "P") for p in PADDS])
    ax.yaxis.set_major_formatter(_thousands)
    ax.set_ylabel("kbd")
    ax.set_title("Unplanned Offline by PADD (kbd)")
    ax.legend(frameon=False, ncol=2)
    _clean(ax)
    return _save(fig, path)


def padd_donut(ctx, path):
    m = ctx["padd_unplanned"]
    ly = max(y for y in m.columns if y not in engine.PARTIAL_YEARS
             and y <= 2025 and m[y].sum() > 0)
    vals = [m.loc[p, ly] for p in PADDS]
    fig, ax = plt.subplots(figsize=(5.2, 4.3))
    cols = [NAVY, BLUE, GOLD, GREEN, ORANGE]
    w, _, auto = ax.pie(vals, colors=cols, startangle=90,
                        wedgeprops=dict(width=0.42, edgecolor="white"),
                        autopct=lambda p: f"{p:.0f}%" if p > 4 else "",
                        pctdistance=0.78, textprops=dict(color="white", fontsize=10,
                                                         fontweight="bold"))
    ax.legend(PADDS, loc="center", frameon=False, fontsize=9)
    ax.set_title(f"Unplanned Share by PADD ({ly})")
    return _save(fig, path)


def seasonality(ctx, path):
    mu = ctx["monthly_unplanned"]
    fig, ax = plt.subplots(figsize=(10, 4.6))
    cmap = {2022: GRAY, 2023: GREEN, 2024: BLUE, 2025: RED, 2026: GOLD}
    for y in [2022, 2023, 2024, 2025, 2026]:
        if y in mu.index:
            ax.plot(MONTHS, [mu.loc[y, m] for m in MONTHS],
                    color=cmap[y], lw=2.4 if y == 2025 else 1.8,
                    ls="--" if y == 2026 else "-",
                    marker="o", ms=3.5, label=f"{y}*" if y in engine.PARTIAL_YEARS else str(y))
    ax.yaxis.set_major_formatter(_thousands)
    ax.set_ylabel("kbd")
    ax.set_title("Unplanned Offline Seasonality (kbd by month)")
    ax.legend(frameon=False, ncol=5, loc="upper right")
    _clean(ax)
    return _save(fig, path)


def padd_combo(ctx, padd, path, figsize=(7.2, 4.2), compact=False):
    """The reference archetype: 2026 plan+unplanned stacked columns, prior-year
    total lines, 2027 plan line."""
    pm = ctx["padd_month"][padd]
    CUR = 2026

    def row(mat, yr):
        return [mat.loc[yr, m] if yr in mat.index else 0.0 for m in MONTHS]

    fig, ax = plt.subplots(figsize=figsize)
    x = np.arange(12)
    plan = row(pm["planned"], CUR)
    unpl = row(pm["unplanned"], CUR)
    ax.bar(x, plan, color=GOLD, label="2026 Planned", zorder=3)
    ax.bar(x, unpl, bottom=plan, color=ORANGE, label="2026 Unplanned", zorder=3)
    for yr, color in [(2025, RED), (2024, BLUE), (2023, GRAY)]:
        ax.plot(x, row(pm["total"], yr), color=color, lw=2.2, marker="o", ms=3,
                label=f"{yr} Total")
    ax.plot(x, row(pm["planned"], 2027), color=GREEN, lw=2.2, ls="--", marker="s", ms=3,
            label="2027 Planned")
    ax.set_xticks(x)
    ax.set_xticklabels([m[0] for m in MONTHS] if compact else MONTHS,
                       fontsize=9 if not compact else 8)
    ax.yaxis.set_major_formatter(_thousands)
    ax.set_ylabel("kbd")
    ax.set_title(f"{padd} Planned & Unplanned Offline (kbd)",
                 fontsize=12 if not compact else 11)
    if not compact:
        ax.legend(frameon=False, ncol=3, fontsize=9, loc="upper center",
                  bbox_to_anchor=(0.5, -0.12))
    _clean(ax)
    return _save(fig, path)


def padd_small_multiples(ctx, padds, path):
    """2x2 small multiples of combo charts for the given PADDs (shared legend)."""
    fig, axes = plt.subplots(2, 2, figsize=(10.5, 6.4))
    CUR = 2026
    for ax, padd in zip(axes.flat, padds):
        pm = ctx["padd_month"][padd]

        def row(mat, yr):
            return [mat.loc[yr, m] if yr in mat.index else 0.0 for m in MONTHS]
        x = np.arange(12)
        plan, unpl = row(pm["planned"], CUR), row(pm["unplanned"], CUR)
        ax.bar(x, plan, color=GOLD, zorder=3)
        ax.bar(x, unpl, bottom=plan, color=ORANGE, zorder=3)
        for yr, color in [(2025, RED), (2024, BLUE), (2023, GRAY)]:
            ax.plot(x, row(pm["total"], yr), color=color, lw=1.6)
        ax.plot(x, row(pm["planned"], 2027), color=GREEN, lw=1.6, ls="--")
        ax.set_xticks(x)
        ax.set_xticklabels([m[0] for m in MONTHS], fontsize=7)
        ax.yaxis.set_major_formatter(_thousands)
        ax.set_title(padd, fontsize=11)
        _clean(ax)
    handles = [
        plt.Rectangle((0, 0), 1, 1, color=GOLD), plt.Rectangle((0, 0), 1, 1, color=ORANGE),
        plt.Line2D([0], [0], color=RED, lw=2), plt.Line2D([0], [0], color=BLUE, lw=2),
        plt.Line2D([0], [0], color=GRAY, lw=2), plt.Line2D([0], [0], color=GREEN, lw=2, ls="--")]
    labels = ["2026 Planned", "2026 Unplanned", "2025 Total", "2024 Total",
              "2023 Total", "2027 Planned"]
    fig.legend(handles, labels, loc="lower center", ncol=6, frameon=False, fontsize=9)
    fig.suptitle("PADD Planned & Unplanned Offline (kbd)", fontsize=14,
                 fontweight="bold", color=NAVY)
    fig.tight_layout(rect=[0, 0.05, 1, 0.96])
    fig.savefig(path, dpi=200, bbox_inches="tight")
    plt.close(fig)
    return path


def units_bar(ctx, path, topn=10):
    mat = ctx["unit_total"]
    tot = mat.sum(axis=1).sort_values(ascending=True).tail(topn)
    fig, ax = plt.subplots(figsize=(9.5, 5.2))
    ax.barh([str(u).title() for u in tot.index], tot.values, color=GOLD, zorder=3)
    ax.xaxis.set_major_formatter(_thousands)
    ax.set_xlabel("kbd (all years)")
    ax.set_title(f"Top {topn} Unit Categories by Capacity Offline")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(axis="x", zorder=0)
    ax.grid(axis="y", visible=False)
    ax.tick_params(length=0)
    for i, v in enumerate(tot.values):
        ax.text(v, i, f"  {v:,.0f}", va="center", fontsize=9, color=NAVY)
    return _save(fig, path)


def scenario_lines(ctx, path):
    sc = ctx["scenario"]
    mu = ctx["monthly_unplanned"]
    mp = ctx["monthly_planned"]
    fig, ax = plt.subplots(figsize=(10, 4.6))
    ax.plot(MONTHS, sc["monthly_unplanned"].values, color=RED, lw=3, marker="o", ms=4,
            label="2027 Scenario (unplanned)")
    if 2027 in mp.index:
        ax.plot(MONTHS, [mp.loc[2027, m] for m in MONTHS], color=GOLD, lw=2, ls="--",
                marker="s", ms=3, label="2027 Planned (booked)")
    for yr, c in [(2025, BLUE), (2024, GRAY)]:
        if yr in mu.index:
            ax.plot(MONTHS, [mu.loc[yr, m] for m in MONTHS], color=c, lw=1.6,
                    label=f"{yr} Unplanned (actual)")
    ax.yaxis.set_major_formatter(_thousands)
    ax.set_ylabel("kbd")
    ax.set_title("2027 Scenario vs Planned vs Recent Actual Unplanned (kbd)")
    ax.legend(frameon=False, ncol=2, loc="upper right")
    _clean(ax)
    return _save(fig, path)


def sensitivity_heatmap(ctx, path):
    anchor = float(ctx["scenario"]["monthly_unplanned"].sum()
                   / ((1 + ctx["scenario"]["growth"]) * ctx["scenario"]["multiplier"]))
    # recompute clean anchor = baseline annual (pre-shock) for default window
    anchor = float(engine.baseline_profile(ctx["df"], ctx["scenario"]["window"]).sum())
    growths = [-0.10, -0.05, 0.0, 0.05, 0.10, 0.15]
    mults = [0.7, 0.85, 1.0, 1.15, 1.3, 1.5]
    grid = np.array([[anchor * (1 + g) * m for m in mults] for g in growths])
    fig, ax = plt.subplots(figsize=(7.6, 5.0))
    im = ax.imshow(grid, cmap="RdYlGn_r", aspect="auto")
    ax.set_xticks(range(len(mults)), [f"{m:.2f}x" for m in mults])
    ax.set_yticks(range(len(growths)), [f"{g:+.0%}" for g in growths])
    ax.set_xlabel("Unplanned rate multiplier")
    ax.set_ylabel("Production growth")
    ax.set_title("2027 Unplanned Sensitivity (kbd)")
    for i in range(len(growths)):
        for j in range(len(mults)):
            ax.text(j, i, f"{grid[i, j]:,.0f}", ha="center", va="center",
                    fontsize=8, color="black")
    # outline base case (g=0, m=1.0)
    bi, bj = growths.index(0.0), mults.index(1.0)
    ax.add_patch(plt.Rectangle((bj - 0.5, bi - 0.5), 1, 1, fill=False,
                               edgecolor=NAVY, lw=2.5))
    ax.tick_params(length=0)
    return _save(fig, path)


def tornado(ctx, path):
    rows = sorted(ctx["tornado"], key=lambda r: r["swing"])  # smallest first -> bottom
    base = rows[0]["base"]
    fig, ax = plt.subplots(figsize=(8.4, 4.2))
    y = np.arange(len(rows))
    for i, r in enumerate(rows):
        ax.barh(i, r["low"] - base, left=base, color=BLUE, zorder=3)
        ax.barh(i, r["high"] - base, left=base, color=ORANGE, zorder=3)
    ax.axvline(base, color=NAVY, lw=1.5)
    ax.set_yticks(y, [r["driver"].split(" (")[0] for r in rows], fontsize=9)
    ax.xaxis.set_major_formatter(_thousands)
    ax.set_xlabel("2027 Unplanned (kbd)")
    ax.set_title("Tornado - Driver Sensitivity (centered on base case)")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(axis="x", zorder=0)
    ax.grid(axis="y", visible=False)
    ax.tick_params(length=0)
    handles = [plt.Rectangle((0, 0), 1, 1, color=BLUE), plt.Rectangle((0, 0), 1, 1, color=ORANGE)]
    ax.legend(handles, ["Downside", "Upside"], frameon=False, ncol=2, loc="lower right")
    return _save(fig, path)


def render_all(ctx, outdir):
    """Render every deck chart into outdir; return a dict name -> path."""
    import os
    os.makedirs(outdir, exist_ok=True)
    p = lambda n: os.path.join(outdir, n)
    out = {
        "annual": annual_stack(ctx, p("annual.png")),
        "padd_clustered": padd_clustered(ctx, p("padd_clustered.png")),
        "padd_donut": padd_donut(ctx, p("padd_donut.png")),
        "seasonality": seasonality(ctx, p("seasonality.png")),
        "padd3": padd_combo(ctx, "PADD 3", p("padd3.png")),
        "padd_sm": padd_small_multiples(ctx, ["PADD 1", "PADD 2", "PADD 4", "PADD 5"],
                                        p("padd_sm.png")),
        "units": units_bar(ctx, p("units.png")),
        "scenario": scenario_lines(ctx, p("scenario.png")),
        "heatmap": sensitivity_heatmap(ctx, p("heatmap.png")),
        "tornado": tornado(ctx, p("tornado.png")),
    }
    return out


if __name__ == "__main__":
    import sys
    from pathlib import Path
    default = str(Path(__file__).resolve().parent.parent / "data" / "rEFINERY oUTAGES.xlsx")
    ctx = engine.build_context(sys.argv[1] if len(sys.argv) > 1 else default)
    paths = render_all(ctx, "scratch_charts")
    for k, v in paths.items():
        print(f"{k}: {v}")
