"""
charts.py
Matplotlib renderers for the slide deck (and any static-image needs).

Each function takes the engine.build_context() bundle and an output path, writes
a high-resolution PNG styled to the desk palette, and returns the path. Keeping
all static-chart styling here means the deck stays consistent and a refresh only
re-renders.
"""
import math
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
    ax2 = ax.twinx()                       # prior-year total lines on a 2nd axis
    x = np.arange(12)
    plan = row(pm["planned"], CUR)
    unpl = row(pm["unplanned"], CUR)
    ax.bar(x, plan, color=GOLD, label="2026 Planned", zorder=3)
    ax.bar(x, unpl, bottom=plan, color=ORANGE, label="2026 Unplanned", zorder=3)
    for yr, color in [(2025, RED), (2024, BLUE), (2023, GRAY)]:
        ax2.plot(x, row(pm["total"], yr), color=color, lw=2.2, marker="o", ms=3,
                 label=f"{yr} Total")
    ax2.plot(x, row(pm["planned"], 2027), color=GREEN, lw=2.2, ls="--", marker="s", ms=3,
             label="2027 Planned")
    ax.set_xticks(x)
    ax.set_xticklabels([m[0] for m in MONTHS] if compact else MONTHS,
                       fontsize=9 if not compact else 8)
    ax.yaxis.set_major_formatter(_thousands)
    ax2.yaxis.set_major_formatter(_thousands)
    ax.set_ylabel("2026 plan/unplanned (kbd)")
    ax2.set_ylabel("prior-yr total / 2027 plan (kbd)", fontsize=9)
    ax.set_ylim(bottom=0); ax2.set_ylim(bottom=0)
    ax2.spines["top"].set_visible(False)
    ax.set_title(f"{padd} Planned & Unplanned Offline (kbd)",
                 fontsize=12 if not compact else 11)
    if not compact:
        h1, l1 = ax.get_legend_handles_labels()
        h2, l2 = ax2.get_legend_handles_labels()
        ax.legend(h1 + h2, l1 + l2, frameon=False, ncol=3, fontsize=9, loc="upper center",
                  bbox_to_anchor=(0.5, -0.12))
    ax.spines["top"].set_visible(False)
    ax.grid(axis="y", zorder=0)
    return _save(fig, path)


def padd_small_multiples(ctx, padds, path):
    """Small multiples of combo charts for the given PADDs (shared legend). Grid
    auto-sizes: 4 PADDs -> 2x2, 5-6 -> 2x3 (spare cell blanked)."""
    n = len(padds)
    ncol = 3 if n > 4 else 2
    nrow = int(math.ceil(n / ncol))
    fig, axes = plt.subplots(nrow, ncol, figsize=(3.6 * ncol, 3.15 * nrow))
    axes = np.atleast_1d(axes).reshape(-1)
    CUR = 2026
    for ax, padd in zip(axes, padds):
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
    for ax in axes[n:]:                       # blank any spare grid cells
        ax.axis("off")
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


def scenario_total_bars(ctx, path):
    """2027 implied total offline by scenario: booked planned (constant) plus the
    Conservative / Average / Active unplanned path. Replaces the old tornado -
    shows the actual outcome range, not abstract driver swings."""
    fan = ctx["scenario_fan"]
    pl = float(ctx["summary"].loc[2027, "Planned"]) if 2027 in ctx["summary"].index else 0.0
    names = ["Conservative", "Average", "Active"]
    mult = {"Conservative": 0.8, "Average": 1.0, "Active": 1.3}
    unpl = [float(fan[n].sum()) for n in names]
    fig, ax = plt.subplots(figsize=(9.6, 3.5))
    x = np.arange(3)
    ax.bar(x, [pl] * 3, color=NAVY, label="Booked planned", zorder=3)
    ax.bar(x, unpl, bottom=[pl] * 3, color=GOLD, label="Scenario unplanned", zorder=3)
    for i, u in enumerate(unpl):
        ax.text(i, pl + u, f"{pl + u:,.0f}", ha="center", va="bottom", fontsize=9.5,
                fontweight="bold", color=NAVY)
        ax.text(i, pl + u / 2, f"+{u:,.0f}", ha="center", va="center", fontsize=8, color="white")
        ax.text(i, pl / 2, f"{pl:,.0f}", ha="center", va="center", fontsize=8, color="white")
    ax.set_xticks(x, [f"{n}\n(x{mult[n]})" for n in names])
    ax.set_xlabel("2027 unplanned scenario")
    ax.yaxis.set_major_formatter(_thousands)
    ax.set_ylabel("kbd offline")
    ax.set_ylim(top=(pl + max(unpl)) * 1.16)
    ax.set_title("2027 Implied Total Offline by Scenario (kbd)")
    ax.legend(frameon=False, ncol=2, loc="upper left", fontsize=8.5)
    _clean(ax)
    return _save(fig, path)


def _runs_in(vec, min_len=3):
    """Index ranges of consecutive non-zero months in a 12-vector."""
    runs, start = [], None
    for i, v in enumerate(vec):
        if v > 0 and start is None:
            start = i
        elif v <= 0 and start is not None:
            if i - start >= min_len:
                runs.append((start, i - 1))
            start = None
    if start is not None and 12 - start >= min_len:
        runs.append((start, 11))
    return runs


def fcc_cluster_strip(ctx, path):
    """Heat-strip of ExxonMobil FCC capacity offline by plant x month, with
    back-to-back runs (>=3 consecutive months) outlined. The clustered signal
    external trackers miss."""
    grid = ctx["fcc_grid"]
    items = [(f"{pl.replace(' Refinery','')}  {yr}", vec)
             for (pl, yr), vec in sorted(grid.items(), key=lambda kv: (kv[0][0], kv[0][1]))
             if _runs_in(vec, 3)]
    if not items:
        items = [(f"{pl.replace(' Refinery','')}  {yr}", vec)
                 for (pl, yr), vec in sorted(grid.items())]
    labels = [a for a, _ in items]
    mat = np.array([b for _, b in items])
    fig, ax = plt.subplots(figsize=(10.5, max(3.2, 0.55 * len(items) + 1.4)))
    im = ax.imshow(mat, cmap="Oranges", aspect="auto", vmin=0)
    ax.set_xticks(range(12), MONTHS)
    ax.set_yticks(range(len(labels)), labels, fontsize=9)
    for i, (_, vec) in enumerate(items):
        for j in range(12):
            if vec[j] > 0:
                ax.text(j, i, f"{vec[j]:.0f}", ha="center", va="center", fontsize=7,
                        color="#23272e")
        for (s, e) in _runs_in(vec, 3):                       # outline the run
            ax.add_patch(plt.Rectangle((s - 0.5, i - 0.5), e - s + 1, 1, fill=False,
                                       edgecolor=RED, lw=2.4))
    ax.set_title("ExxonMobil FCC Offline by Month - back-to-back runs outlined (kbd)")
    ax.grid(False)
    ax.tick_params(length=0)
    return _save(fig, path)


def scenario_by_padd(ctx, path):
    """Stacked columns: 2027 scenario unplanned by month, decomposed by PADD."""
    sp = ctx["scenario_padd"]
    fig, ax = plt.subplots(figsize=(10, 4.6))
    x = np.arange(12)
    bottom = np.zeros(12)
    cols = {"PADD 1": "#9DB0CE", "PADD 2": BLUE, "PADD 3": NAVY, "PADD 4": GREEN, "PADD 5": GOLD}
    for p in PADDS:
        vals = sp[p]["monthly"].values
        ax.bar(x, vals, bottom=bottom, color=cols[p], label=p, zorder=3)
        bottom += vals
    ax.set_xticks(x, MONTHS)
    ax.yaxis.set_major_formatter(_thousands)
    ax.set_ylabel("kbd")
    ax.set_title("2027 Unplanned Scenario by PADD (each PADD's own seasonality)")
    ax.legend(frameon=False, ncol=5, loc="upper right")
    _clean(ax)
    return _save(fig, path)


def seasonality_band(ctx, path):
    """Range-band seasonality chart (matches the reference demand charts): a
    shaded historical min-max band with recent-year lines on top."""
    rb = ctx["range_band"]
    mu = ctx["monthly_unplanned"]
    fig, ax = plt.subplots(figsize=(10, 4.7))
    x = np.arange(12)
    ax.fill_between(x, rb["min"].values, rb["max"].values, color="#D9D9D9", alpha=0.85,
                    zorder=1, label="Range (22-25)")
    ax.plot(x, rb["avg"].values, color="#7F7F7F", lw=1.4, ls=":", zorder=2, label="5-yr avg")
    for y, c, w in [(2024, BLUE, 1.8), (2025, RED, 2.6), (2026, GOLD, 2.0)]:
        if y in mu.index:
            ax.plot(x, [mu.loc[y, m] for m in MONTHS], color=c, lw=w,
                    ls="--" if y == 2026 else "-", marker="o", ms=3,
                    label=f"{y}*" if y in engine.PARTIAL_YEARS else str(y), zorder=3)
    ax.set_xticks(x, MONTHS)
    ax.yaxis.set_major_formatter(_thousands)
    ax.set_ylabel("kbd")
    ax.set_title("Unplanned Offline by Month - range band & recent years (kbd)")
    ax.legend(frameon=False, ncol=5, loc="upper right", fontsize=9)
    _clean(ax)
    return _save(fig, path)


def monthly_yoy_bars(ctx, path):
    """Grouped bars: YoY % change in unplanned offline for each month (the
    'percent difference in each month by year' view)."""
    _, yoy = ctx["monthly_yoy"]
    fig, ax = plt.subplots(figsize=(10, 4.3))
    x = np.arange(12)
    yrs = [y for y in (2024, 2025) if y in yoy.index]
    w = 0.8 / max(1, len(yrs))
    cols = {2024: BLUE, 2025: RED}
    for i, y in enumerate(yrs):
        vals = [yoy.loc[y, m] * 100 if yoy.loc[y, m] == yoy.loc[y, m] else 0 for m in MONTHS]
        ax.bar(x + i * w, vals, w, color=cols[y], label=str(y), zorder=3)
    ax.axhline(0, color="#404040", lw=1)
    ax.set_xticks(x + 0.4 - w / 2, MONTHS)
    ax.set_ylabel("YoY % change")
    ax.set_title("Unplanned Offline - YoY % Change by Month")
    ax.legend(frameon=False, ncol=2, loc="upper right")
    _clean(ax)
    return _save(fig, path)


def mogas_annual_chart(ctx, path):
    ma = ctx["mogas_annual"]
    years = [y for y in ma.index if 2018 <= y <= 2027]
    fig, ax = plt.subplots(figsize=(6.4, 4.3))
    x = np.arange(len(years))
    ax.bar(x, [ma.loc[y, "Planned"] for y in years], color=NAVY, label="Planned", zorder=3)
    ax.bar(x, [ma.loc[y, "Unplanned"] for y in years],
           bottom=[ma.loc[y, "Planned"] for y in years], color=GREEN, label="Unplanned", zorder=3)
    ax.set_xticks(x, [f"{y}*" if y in engine.PARTIAL_YEARS else str(y) for y in years],
                  rotation=45, fontsize=8)
    ax.yaxis.set_major_formatter(_thousands)
    ax.set_ylabel("kbd (mogas-eq)")
    ax.set_title("Mogas-Equivalent Offline by Year")
    ax.legend(frameon=False, ncol=2)
    _clean(ax)
    return _save(fig, path)


def planned_cross_year(ctx, path):
    """US planned offline by month: 2027 (bars) vs 2026 & 2025 (lines), with an
    H1 | H2 divider because the 2027 book is only complete (booked) through H1."""
    mp = ctx["monthly_planned"]
    fig, ax = plt.subplots(figsize=(10, 4.7))
    x = np.arange(12)
    v27 = [mp.loc[2027, m] if 2027 in mp.index else 0.0 for m in MONTHS]
    ax.bar(x, v27, color=GOLD, label="2027 Planned", zorder=3, width=0.62)
    for yr, c in [(2026, BLUE), (2025, RED)]:
        if yr in mp.index:
            ax.plot(x, [mp.loc[yr, m] for m in MONTHS], color=c, lw=2.4, marker="o", ms=3.5,
                    label=f"{yr} Planned" + ("*" if yr in engine.PARTIAL_YEARS else ""))
    # mark the H1 / H2 boundary: 2027 H2 is still filling in -> read H1 for like-for-like
    ax.axvline(5.5, color=GRAY, ls=":", lw=1.4, zorder=2)
    ax.set_xticks(x, MONTHS)
    ax.yaxis.set_major_formatter(_thousands)
    ax.set_ylabel("kbd")
    ax.set_title("US Planned Offline by Month - 2027 vs 2026 & 2025 (kbd)", pad=28)
    # legend ABOVE the axes (between title and plot) so it never sits on the data
    ax.legend(frameon=False, ncol=3, loc="lower left", bbox_to_anchor=(0, 1.0),
              borderaxespad=0.3)
    _clean(ax)
    # headroom so the H1/H2 region labels float above the bars/lines, clear of legend
    top = ax.get_ylim()[1] * 1.15
    ax.set_ylim(top=top)
    ax.text(2.5, top * 0.97, "H1 (complete)", ha="center", va="top", fontsize=9, color=GRAY)
    ax.text(8.5, top * 0.97, "H2 2027 incomplete*", ha="center", va="top", fontsize=9,
            color=RED, style="italic")
    fig.text(0.5, 0.005, "*2027 planned book is complete through H1; H2 still being scheduled - "
             "compare H1-vs-H1 for a like-for-like read.", ha="center", fontsize=8, color=GRAY)
    return _save(fig, path)


def exxon_2027_chart(ctx, path):
    """ExxonMobil 2027 planned offline by month, stacked by refinery."""
    ex = ctx["exxon_2027"]
    fig, ax = plt.subplots(figsize=(9.6, 4.6))
    x = np.arange(12)
    bottom = np.zeros(12)
    cols = [NAVY, BLUE, GOLD, GREEN, ORANGE, GRAY]
    for i, ref in enumerate(ex["refs"][:6]):
        vals = np.array(ex["month_ref"][ref])
        ax.bar(x, vals, bottom=bottom, color=cols[i % len(cols)],
               label=ref.replace(" Refinery", ""), zorder=3)
        bottom += vals
    ax.set_xticks(x, MONTHS)
    ax.yaxis.set_major_formatter(_thousands)
    ax.set_ylabel("kbd")
    ax.set_title("ExxonMobil 2027 Planned Offline by Month & Refinery (kbd)", pad=24)
    ax.legend(frameon=False, ncol=4, loc="lower left", bbox_to_anchor=(0, 1.0),
              borderaxespad=0.3, fontsize=9)
    _clean(ax)
    return _save(fig, path)


def padd_planned_27v26(ctx, padd, path):
    """One PADD: planned offline by month, 2027 (bars) vs 2026 (line)."""
    pm = ctx["padd_month"][padd]["planned"]
    fig, ax = plt.subplots(figsize=(7.0, 4.0))
    x = np.arange(12)
    v27 = [pm.loc[2027, m] if 2027 in pm.index else 0.0 for m in MONTHS]
    v26 = [pm.loc[2026, m] if 2026 in pm.index else 0.0 for m in MONTHS]
    ax.bar(x, v27, color=GOLD, label="2027 Planned", zorder=3, width=0.62)
    ax.plot(x, v26, color=BLUE, lw=2.4, marker="o", ms=3.5, label="2026 Planned")
    ax.set_xticks(x, [m[0] for m in MONTHS])
    ax.yaxis.set_major_formatter(_thousands)
    ax.set_ylabel("kbd")
    ax.set_title(f"{padd} Planned Offline - 2027 vs 2026 (kbd)", pad=22)
    ax.legend(frameon=False, ncol=2, loc="lower left", bbox_to_anchor=(0, 1.0),
              borderaxespad=0.3, fontsize=9)
    _clean(ax)
    return _save(fig, path)


def scenario_fan_chart(ctx, path):
    """2027 unplanned forecast as Conservative / Average / Active paths."""
    fan = ctx["scenario_fan"]
    mu = ctx["monthly_unplanned"]
    fig, ax = plt.subplots(figsize=(10, 4.7))
    x = np.arange(12)
    ax.fill_between(x, fan["Conservative"].values, fan["Active"].values,
                    color="#D9E1F2", alpha=0.7, zorder=1, label="Conservative-Active range")
    styles = {"Conservative": (GREEN, "--", 3.2), "Average": (NAVY, "-", 4.0), "Active": (RED, "--", 3.2)}
    for name, prof in fan.items():
        c, ls, lw = styles[name]
        ax.plot(x, prof.values, color=c, lw=lw, ls=ls, marker="o", ms=6,
                label=f"{name} (~{prof.sum():,.0f} kbd)", zorder=3)
    if 2025 in mu.index:
        ax.plot(x, [mu.loc[2025, m] for m in MONTHS], color=GRAY, lw=2.0, ls=":", zorder=2,
                label="2025 actual")
    ax.set_xticks(x, MONTHS)
    ax.set_xlabel("Month (2027)")
    ax.yaxis.set_major_formatter(_thousands)
    ax.set_ylabel("kbd offline")
    ax.set_title("2027 Unplanned Forecast - Conservative / Average / Active (kbd)")
    ax.set_ylim(top=ax.get_ylim()[1] * 1.04)
    ax.legend(frameon=False, ncol=2, loc="upper right", fontsize=8.5)
    _clean(ax)
    return _save(fig, path)


def dollar_at_risk(ctx, path):
    """Annual $ of gross refining margin at risk (offline valued at the gasoline
    crack), unplanned vs planned bars + the average crack on a 2nd axis."""
    di = ctx.get("dollar_impact") or {}
    fig, ax = plt.subplots(figsize=(9.8, 4.7))
    if not di:
        ax.text(0.5, 0.5, "No crack data - run scripts/fetch_market_data.py",
                ha="center", va="center", fontsize=12, color=GRAY)
        ax.axis("off")
        return _save(fig, path)
    years = [y for y in sorted(di) if 2018 <= y <= 2026]
    x = np.arange(len(years))
    ax2 = ax.twinx()
    ax.bar(x - 0.2, [di[y]["unplanned"] for y in years], width=0.4, color=NAVY,
           label="Unplanned $ at risk", zorder=3)
    ax.bar(x + 0.2, [di[y]["planned"] for y in years], width=0.4, color=GOLD,
           label="Planned $", zorder=3)
    ax2.plot(x, [di[y]["crack_avg"] for y in years], color=RED, lw=2.6, marker="o", ms=4,
             label="Avg crack ($/bbl)")
    ax.set_xticks(x, [str(y) for y in years])
    ax.yaxis.set_major_formatter(_thousands)
    ax.set_ylabel("$MM / yr"); ax2.set_ylabel("crack ($/bbl)")
    ax.set_title("Gross Refining Margin at Risk from Outages ($MM) vs Gasoline Crack")
    h1, l1 = ax.get_legend_handles_labels()
    h2, l2 = ax2.get_legend_handles_labels()
    ax.legend(h1 + h2, l1 + l2, frameon=False, ncol=3, loc="upper right", fontsize=9)
    _clean(ax)
    return _save(fig, path)


def padd_planned_pair(ctx, path):
    """Where the 2027 build lands: planned offline by PADD, 2026 vs 2027, with
    the y/y% labelled over each region (pairs with the monthly cross-year view)."""
    pp = ctx["padd_planned"]
    padds = list(pp.index)
    fig, ax = plt.subplots(figsize=(9.6, 3.5))
    x = np.arange(len(padds))
    v26 = [float(pp.loc[p, 2026]) for p in padds]
    v27 = [float(pp.loc[p, 2027]) for p in padds]
    ax.bar(x - 0.2, v26, 0.4, color=BLUE, label="2026 Planned", zorder=3)
    ax.bar(x + 0.2, v27, 0.4, color=GOLD, label="2027 Planned", zorder=3)
    for i, (a, b) in enumerate(zip(v26, v27)):
        if a > 0:
            ax.text(i, max(a, b), f"  {b / a - 1:+.0%}", ha="center", va="bottom",
                    fontsize=8.5, color=NAVY, fontweight="bold")
    ax.set_xticks(x, padds)
    ax.yaxis.set_major_formatter(_thousands)
    ax.set_ylabel("kbd")
    ax.set_ylim(top=max(v26 + v27) * 1.18)
    ax.set_title("Planned Offline by PADD - 2026 vs 2027 (kbd, with y/y%)")
    ax.legend(frameon=False, ncol=2, loc="upper right", fontsize=9)
    _clean(ax)
    return _save(fig, path)


def dollar_monthly_profile(ctx, path):
    """When in the year the unplanned $-at-risk lands: monthly $MM, 2025 vs 2026
    (pairs with the annual $-at-risk bars)."""
    di = ctx.get("dollar_impact") or {}
    fig, ax = plt.subplots(figsize=(9.6, 3.5))
    if not di:
        ax.axis("off")
        return _save(fig, path)
    x = np.arange(12)
    for yr, c, off in [(2025, NAVY, -0.2), (2026, GOLD, 0.2)]:
        if yr in di:
            ax.bar(x + off, di[yr]["monthly_unpl"], 0.4, color=c,
                   label=f"{yr} unplanned $", zorder=3)
    ax.set_xticks(x, [m[0] for m in MONTHS])
    ax.yaxis.set_major_formatter(_thousands)
    ax.set_ylabel("$MM")
    ax.set_title("Monthly Unplanned Margin at Risk - 2025 vs 2026 ($MM)")
    ax.legend(frameon=False, ncol=2, loc="upper right", fontsize=9)
    _clean(ax)
    return _save(fig, path)


def exxon_by_unit(ctx, path):
    """ExxonMobil 2027 planned offline split by unit category (pairs with the
    by-month / by-refinery stacked view)."""
    bu = ctx["exxon_2027"]["by_unit"]
    items = sorted(bu.items(), key=lambda kv: kv[1])[-7:]
    fig, ax = plt.subplots(figsize=(9.6, 3.5))
    labels = [str(k).title() for k, _ in items]
    vals = [v for _, v in items]
    ax.barh(labels, vals, color=NAVY, zorder=3)
    ax.xaxis.set_major_formatter(_thousands)
    ax.set_xlabel("kbd")
    ax.set_xlim(right=max(vals) * 1.16)
    ax.set_title("ExxonMobil 2027 Planned Offline by Unit (kbd)")
    for i, v in enumerate(vals):
        ax.text(v, i, f"  {v:,.0f}", va="center", fontsize=8.5, color=NAVY)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(axis="x", zorder=0)
    ax.grid(axis="y", visible=False)
    ax.tick_params(length=0)
    return _save(fig, path)


def fcc_by_year(ctx, path):
    """US-wide FCC (cat cracker) offline by year - the recurring, multi-operator
    market pattern behind the Exxon-specific back-to-back runs (2020-21 capped)."""
    row = ctx["unit_total"].loc["FLUID CAT CRACKING"]
    years = [y for y in range(2014, 2027) if y in row.index]
    vals = [float(row[y]) for y in years]
    ref = max([v for y, v in zip(years, vals) if y not in engine.OUTLIER_YEARS] or vals)
    disp = [min(v, ref) if y in engine.OUTLIER_YEARS else v for y, v in zip(years, vals)]
    fig, ax = plt.subplots(figsize=(9.6, 3.5))
    x = np.arange(len(years))
    ax.bar(x, disp, color=GOLD, zorder=3)
    for i, y in enumerate(years):
        if y in engine.OUTLIER_YEARS:
            ax.text(i, disp[i], "^", ha="center", va="bottom", fontsize=11, color=RED)
    ax.set_xticks(x, [f"{y}*" if y in engine.PARTIAL_YEARS else str(y) for y in years], fontsize=8)
    ax.yaxis.set_major_formatter(_thousands)
    ax.set_ylabel("kbd")
    ax.set_title("US FCC (Cat Cracker) Offline by Year - All Operators (kbd)")
    _clean(ax)
    fig.text(0.5, 0.004, "2020-21 (COVID / Winter Storm Uri) capped for scale (^).",
             ha="center", fontsize=7.5, color=GRAY)
    return _save(fig, path)


def _dcolors(vals):
    return [GREEN if v >= 0 else RED for v in vals]


def mom_movers_chart(ctx, path):
    """Month-over-month change in offline capacity: diverging bars by PADD (left)
    and by the biggest-moving unit categories (right)."""
    pc = ctx.get("period_change")
    fig, (axp, axu) = plt.subplots(1, 2, figsize=(10.2, 3.7))
    if not pc:
        for ax in (axp, axu):
            ax.axis("off")
        axp.text(0.5, 0.5, "Need >= 2 reported months", ha="center", va="center", color=GRAY)
        return _save(fig, path)
    pad = pc["by_padd"]
    axp.barh([str(i).replace("PADD ", "P") for i in pad.index], pad.values,
             color=_dcolors(pad.values), zorder=3)
    axp.axvline(0, color=NAVY, lw=1)
    axp.set_title("By PADD (kbd)", fontsize=10)
    un = pc["by_unit"]
    top = un.reindex(un.abs().sort_values(ascending=False).index).head(8).iloc[::-1]
    axu.barh([str(i).title()[:16] for i in top.index], top.values,
             color=_dcolors(top.values), zorder=3)
    axu.axvline(0, color=NAVY, lw=1)
    axu.set_title("By unit - biggest movers (kbd)", fontsize=10)
    for ax in (axp, axu):
        ax.xaxis.set_major_formatter(_thousands)
        ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
        ax.grid(axis="x", zorder=0); ax.tick_params(length=0)
    fig.suptitle(f"Month-over-Month Change - {pc['cur_label']} vs {pc['prev_label']} (kbd)",
                 fontsize=13, fontweight="bold", color=NAVY)
    fig.tight_layout(rect=[0, 0, 1, 0.93])
    fig.savefig(path, dpi=200, bbox_inches="tight"); plt.close(fig)
    return path


def mom_trend_chart(ctx, path):
    """Trailing 13 months of total (bars) & unplanned (line) offline, latest month
    highlighted - the context the single MoM step sits in."""
    pc = ctx.get("period_change")
    fig, ax = plt.subplots(figsize=(10.2, 3.3))
    if not pc:
        ax.axis("off"); return _save(fig, path)
    trail = pc["trail"]
    x = np.arange(len(trail))
    tot = [t["total"] for t in trail]
    ax.bar(x, tot, color="#D6E0F0", label="Total offline", zorder=2)
    ax.bar([x[-1]], [tot[-1]], color=GOLD, zorder=3, label=f"Latest ({pc['cur_label']})")
    ax.plot(x, [t["unplanned"] for t in trail], color=RED, lw=2.2, marker="o", ms=3,
            label="Unplanned", zorder=4)
    ax.set_xticks(x, [t["label"] for t in trail], fontsize=7.5)
    ax.yaxis.set_major_formatter(_thousands); ax.set_ylabel("kbd")
    ax.set_title("Trailing 13 Months - Total & Unplanned Offline (kbd)")
    ax.legend(frameon=False, ncol=3, loc="upper right", fontsize=8.5)
    _clean(ax)
    return _save(fig, path)


def mom_newgone_chart(ctx, path):
    """The outages that newly appeared vs dropped off this month, top by kbd."""
    pc = ctx.get("period_change")
    fig, (a0, a1) = plt.subplots(1, 2, figsize=(10.2, 3.6))
    if not pc:
        for ax in (a0, a1):
            ax.axis("off")
        return _save(fig, path)

    def panel(ax, frame, color, title):
        d = frame.head(7).iloc[::-1]
        if len(d) == 0:
            ax.axis("off"); ax.text(0.5, 0.5, "none", ha="center", va="center", color=GRAY)
            ax.set_title(title, fontsize=10); return
        labels = [f"{str(r.plant).replace(' Refinery', '')[:20]} - {str(r.unit).title()[:10]}"
                  for _, r in d.iterrows()]
        ax.barh(range(len(d)), d["kbd"].values, color=color, zorder=3)
        ax.set_yticks(range(len(d)), labels, fontsize=7)
        for i, v in enumerate(d["kbd"].values):
            ax.text(v, i, f" {v:,.0f}", va="center", fontsize=7, color=NAVY)
        ax.set_xlim(right=max(d["kbd"].values) * 1.18)
        ax.xaxis.set_major_formatter(_thousands); ax.set_title(title, fontsize=10)
        ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
        ax.grid(axis="x", zorder=0); ax.tick_params(length=0)
    panel(a0, pc["new"], GREEN, f"New this month ({len(pc['new'])})")
    panel(a1, pc["gone"], GRAY, f"Resolved / came back ({len(pc['gone'])})")
    fig.suptitle(f"Outages Added vs Resolved - {pc['cur_label']} (kbd offline)",
                 fontsize=13, fontweight="bold", color=NAVY)
    fig.tight_layout(rect=[0, 0, 1, 0.92])
    fig.savefig(path, dpi=200, bbox_inches="tight"); plt.close(fig)
    return path


# --------------------------------------------------------------------------- per-unit (focus) charts
FOCUS_COLOR = {"CDU": NAVY, "FCC": RED, "Hydrocracker": GREEN, "Reformer": GOLD}


def focus_heat(ctx, path):
    """2x2 heatmaps: concurrent capacity offline (kbd) by month x year (2021-2027)
    for each focus unit (CDU/FCC/hydrocracker/reformer). Reads as a per-unit
    timeline and never sums across units - the core 'by month & unit' view."""
    fm = ctx["focus_monthly"]
    fig, axes = plt.subplots(2, 2, figsize=(11.6, 6.6))
    for ax, f in zip(axes.reshape(-1), engine.FOCUS_ORDER):
        m = fm[f]
        years = [int(y) for y in m.index]
        mat = m.values
        ax.imshow(mat, cmap="OrRd", aspect="auto", vmin=0)
        ax.set_xticks(range(12), [mo[0] for mo in MONTHS], fontsize=8)
        ax.set_yticks(range(len(years)), years, fontsize=8)
        vmax = mat.max() or 1.0
        for i in range(len(years)):
            for j in range(12):
                v = mat[i, j]
                if v >= 0.06 * vmax:
                    ax.text(j, i, f"{v:,.0f}", ha="center", va="center", fontsize=6,
                            color="white" if v > 0.55 * vmax else "#23272e")
        ax.set_title(engine.FOCUS_LABEL[f], fontsize=11)
        ax.grid(False)
        ax.tick_params(length=0)
    fig.suptitle("Capacity Offline by Unit & Month - concurrent kbd, 2021-2027",
                 fontsize=14, fontweight="bold", color=NAVY)
    fig.tight_layout(rect=[0, 0, 1, 0.95])
    fig.savefig(path, dpi=200, bbox_inches="tight")
    plt.close(fig)
    return path


def unit_year_lines(ctx, focus, path, figsize=(7.4, 4.5)):
    """One focus unit: concurrent capacity offline (kbd) by month, a line per year
    2021-2027 (the per-unit seasonal timeline)."""
    m = ctx["focus_monthly"][focus]
    fig, ax = plt.subplots(figsize=figsize)
    cmap = {2021: "#C9C9C9", 2022: "#9DB0CE", 2023: GREEN, 2024: BLUE,
            2025: RED, 2026: GOLD, 2027: NAVY}
    for y in [int(y) for y in m.index]:
        lw = 2.8 if y == 2027 else (2.3 if y == 2025 else 1.5)
        ls = "--" if y in (2026, 2027) else "-"
        ax.plot(MONTHS, [m.loc[y, mo] for mo in MONTHS], color=cmap.get(y, GRAY),
                lw=lw, ls=ls, marker="o", ms=3,
                label=f"{y}*" if y in engine.PARTIAL_YEARS else str(y))
    ax.yaxis.set_major_formatter(_thousands)
    ax.set_ylabel("kbd offline")
    ax.set_title(f"{engine.FOCUS_LABEL[focus]} - Capacity Offline by Month (kbd)", pad=26)
    ax.legend(frameon=False, ncol=7, fontsize=8, loc="lower left",
              bbox_to_anchor=(0, 1.0), borderaxespad=0.3, columnspacing=1.0)
    _clean(ax)
    return _save(fig, path)


def focus_padd_bars(ctx, focus, year, path, figsize=(7.4, 4.1)):
    """One focus unit, one year: concurrent offline by month, stacked by PADD -
    the 'timeline by month and PADD' read, kept per unit."""
    g = ctx["focus_padd"][year][focus]
    fig, ax = plt.subplots(figsize=figsize)
    x = np.arange(12)
    bottom = np.zeros(12)
    cols = {"PADD 1": "#9DB0CE", "PADD 2": BLUE, "PADD 3": NAVY, "PADD 4": GREEN, "PADD 5": GOLD}
    for pd_ in PADDS:
        vals = g.loc[pd_].values if pd_ in g.index else np.zeros(12)
        ax.bar(x, vals, bottom=bottom, color=cols[pd_], label=pd_.replace("PADD ", "P"), zorder=3)
        bottom += vals
    ax.set_xticks(x, [mo[0] for mo in MONTHS])
    ax.set_xlabel(f"Month ({year})")
    ax.yaxis.set_major_formatter(_thousands)
    ax.set_ylabel("kbd offline")
    ax.set_title(f"{focus} Offline by Month & PADD - {year} (kbd)")
    ax.legend(frameon=False, ncol=5, fontsize=8, loc="upper right")
    _clean(ax)
    if year == 2027:                       # mark the H1 (confirmed) | H2 (non-Exxon unconfirmed) line
        ax.set_ylim(top=ax.get_ylim()[1] * 1.15)
        top = ax.get_ylim()[1]
        ax.axvline(5.5, color=GRAY, ls=":", lw=1.3, zorder=2)
        ax.text(2.6, top * 0.99, "H1 confirmed", ha="center", va="top", fontsize=8, color=GRAY)
        ax.text(8.8, top * 0.99, "H2: non-Exxon unconfirmed", ha="center", va="top",
                fontsize=7.5, color=RED, style="italic")
    return _save(fig, path)


def splits_2027(ctx, path):
    """The 2027 completeness story, per focus unit: monthly concurrent offline
    split into CONFIRMED (Exxon full-year plan + every other operator's H1) and
    NON-EXXON H2 (still being booked, not confirmed). 2x2, one panel per unit."""
    fig, axes = plt.subplots(2, 2, figsize=(11.6, 6.6))
    x = np.arange(12)
    for idx, (ax, f) in enumerate(zip(axes.reshape(-1), engine.FOCUS_ORDER)):
        sp = ctx["confirmed2027"][f]
        conf, ind = np.array(sp["confirmed"]), np.array(sp["indicative"])
        ax.bar(x, conf, color=FOCUS_COLOR.get(f, NAVY), zorder=3)
        ax.bar(x, ind, bottom=conf, color="#E2E2E2", hatch="////", edgecolor="#A6A6A6",
               lw=0.4, zorder=3)
        ax.axvline(5.5, color=GRAY, ls=":", lw=1.2, zorder=2)
        ax.set_xticks(x, [mo[0] for mo in MONTHS], fontsize=7.5)
        ax.yaxis.set_major_formatter(_thousands)
        if idx % 2 == 0:                       # left column -> y units
            ax.set_ylabel("kbd offline", fontsize=9)
        if idx >= 2:                           # bottom row -> x units
            ax.set_xlabel("Month (2027)", fontsize=9)
        ax.set_title(engine.FOCUS_LABEL[f], fontsize=10.5)
        _clean(ax)
    handles = [plt.Rectangle((0, 0), 1, 1, color=NAVY),
               plt.Rectangle((0, 0), 1, 1, facecolor="#E2E2E2", hatch="////", edgecolor="#A6A6A6")]
    fig.legend(handles, ["Confirmed  (Exxon full-year + all others H1)",
                         "Non-Exxon H2  (still being booked - not confirmed)"],
               loc="lower center", ncol=2, frameon=False, fontsize=9.5)
    fig.suptitle("2027 Capacity Offline by Unit - Confirmed vs Not-Yet-Confirmed (kbd/month)",
                 fontsize=13, fontweight="bold", color=NAVY)
    fig.tight_layout(rect=[0, 0.05, 1, 0.95])
    fig.savefig(path, dpi=200, bbox_inches="tight")
    plt.close(fig)
    return path


def exxon_gantt(ctx, path):
    """Per-unit timeline of ExxonMobil's 2027 focus-unit turnarounds, each bar a
    single unit (plant-unit, nameplate kbd) spanning its outage months, colored by
    unit class. Confirmed against Exxon's corporate plan = solid; flagged (no plan
    match) = red hatched. Replaces the meaningless summed-Exxon total."""
    ev = ctx["exxon_verify"]["events"]
    foc = ev[ev["focus"].isin(engine.FOCUS_ORDER)].copy()
    foc["m0"] = foc["months"].apply(lambda m: min(m) if m else 13)
    foc["m1"] = foc["months"].apply(lambda m: max(m) if m else 0)
    foc = foc.sort_values(["m0", "kbd"], ascending=[True, False]).reset_index(drop=True)
    fig, ax = plt.subplots(figsize=(10.6, max(3.4, 0.52 * len(foc) + 1.1)))
    for i, (_, r) in enumerate(foc.iterrows()):
        s, w = r["m0"] - 1, r["m1"] - r["m0"] + 1
        flagged = (r["verified"] is False) or (r["verified"] == False)  # noqa: E712
        ax.barh(i, w, left=s, height=0.62, color=FOCUS_COLOR.get(r["focus"], GRAY),
                edgecolor=RED if flagged else "white", lw=2.4 if flagged else 0.5,
                hatch="///" if flagged else None, zorder=3)
        lbl = f"{r['plant'].replace(' Refinery', '')} {str(r['unit_name'])[:17]} ({r['kbd']:.0f})"
        ax.text(s + w + 0.15, i, lbl + ("   (!) not in plan" if flagged else ""),
                va="center", fontsize=8, color=RED if flagged else "#23272e")
    ax.set_yticks(range(len(foc)), [r["focus"] for _, r in foc.iterrows()], fontsize=7.5)
    ax.set_xticks(range(12), MONTHS)
    ax.set_xlabel("Month (2027)")
    ax.set_xlim(0, 17.5)
    ax.set_ylim(-0.6, len(foc) - 0.4)
    ax.invert_yaxis()
    handles = [plt.Rectangle((0, 0), 1, 1, color=FOCUS_COLOR[f]) for f in engine.FOCUS_ORDER]
    ax.legend(handles, engine.FOCUS_ORDER, frameon=False, ncol=4, fontsize=8,
              loc="lower right")
    ax.set_title("ExxonMobil 2027 - Focus-Unit Turnarounds, per unit (kbd), verified vs corporate plan",
                 fontsize=11.5)
    ax.grid(axis="x", zorder=0)
    ax.tick_params(length=0)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    return _save(fig, path)


def joliet_decode(ctx, path):
    """Why one 'Exxon ~700 kbd' figure is wrong. The units in Joliet's SINGLE
    Apr-2027 turnaround, shown per unit: adding them up = ~715 kbd at a ~250 kbd
    refinery. The honest read is 250 kbd of crude (CDU) offline."""
    ev = engine.unit_events(ctx["df"], operator_contains="EXXON", year=2027)
    jol = ev[ev["plant"].str.contains("Joliet", na=False)
             & ev["months"].apply(lambda m: bool(set(m) & {4, 5}))].copy()
    jol = jol.sort_values("kbd", ascending=True)
    labels = [f"{str(u).title()[:24]}" for u in jol["unit_name"]]
    vals = jol["kbd"].tolist()
    # highlight ONLY the atmospheric crude unit (the CDU) - the 250 kbd headline
    is_cdu = (jol["unit_cat"] == "ATMOS DISTILLATION").tolist()
    colors = [NAVY if c else "#C0C0C0" for c in is_cdu]
    total = sum(vals)
    crude = float(jol[jol["unit_cat"] == "ATMOS DISTILLATION"]["kbd"].sum())
    fig, ax = plt.subplots(figsize=(9.6, 4.8))
    y = np.arange(len(jol))
    ax.barh(y, vals, color=colors, zorder=3)
    ax.set_yticks(y, labels, fontsize=9)
    for i, v in enumerate(vals):
        ax.text(v, i, f"  {v:,.0f}", va="center", fontsize=8.5, color=NAVY)
    ax.xaxis.set_major_formatter(_thousands)
    ax.set_xlabel("nameplate offline (kbd)")
    ax.set_xlim(right=max(vals) * 1.28)
    ax.set_title("One Joliet turnaround, April 2027 - why you can't add units together")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(axis="x", zorder=0)
    ax.grid(axis="y", visible=False)
    ax.tick_params(length=0)
    ax.text(0.97, 0.06,
            f"Sum of all {len(jol)} units = {total:,.0f} kbd\n"
            f"...but Joliet is a ~250 kbd refinery.\n"
            f"Honest read: ~{crude:,.0f} kbd of crude (the CDU) offline;\n"
            f"the rest are separate downstream units.",
            transform=ax.transAxes, ha="right", va="bottom", fontsize=9.5, color=RED,
            bbox=dict(boxstyle="round,pad=0.5", fc="#FFF3F3", ec=RED, lw=1.2))
    return _save(fig, path)


PADD_COLOR = {"PADD 1": "#9DB0CE", "PADD 2": BLUE, "PADD 3": NAVY, "PADD 4": GREEN, "PADD 5": GOLD}


def biggest_outages(ctx, path, year=2027, topn=12):
    """'What's driving the numbers': the single biggest focus-unit outages in
    `year`, one bar per physical unit (nameplate kbd), colored by PADD so you see
    WHERE the big tonnage sits. Units are never added together - each bar stands
    alone. Non-Exxon H2 outages (still being booked) are hatched and tagged: an
    indicative floor, not confirmed."""
    ev = engine.unit_events(ctx["df"], year=year)
    ev = ev[ev["focus"].isin(engine.FOCUS_ORDER)].copy()
    fig, ax = plt.subplots(figsize=(10.8, max(3.6, 0.5 * min(topn, len(ev)) + 1.3)))
    if ev.empty:
        ax.axis("off")
        return _save(fig, path)
    ev["is_exxon"] = ev["operator"].astype(str).str.upper().str.contains("EXXON", na=False)
    ev["indic"] = (~ev["is_exxon"]) & ev["months"].apply(lambda m: bool(m) and min(m) >= 7)
    ev = ev.sort_values("kbd", ascending=False).head(topn).sort_values("kbd")   # biggest -> top
    ev = ev.reset_index(drop=True)
    xmax = float(ev["kbd"].max())
    for i, r in ev.iterrows():
        c = PADD_COLOR.get(r["padd"], GRAY)
        ind = bool(r["indic"])
        ax.barh(i, r["kbd"], height=0.66, color=c, zorder=3,
                alpha=0.5 if ind else 1.0, hatch="////" if ind else None,
                edgecolor="#7F7F7F" if ind else c, lw=0.6)
        note = f"  {r['kbd']:,.0f}   {r['focus']} · {r['span']}" + ("  H2*" if ind else "")
        ax.text(r["kbd"] + xmax * 0.015, i, note, va="center",
                fontsize=8.2, color=RED if ind else "#23272e")
    labels = [f"{str(r['plant']).replace(' Refinery', '')[:24]}: {str(r['unit_name'])[:16]}"
              for _, r in ev.iterrows()]
    ax.set_yticks(range(len(ev)), labels, fontsize=8)
    ax.set_ylim(-0.7, len(ev) - 0.3)
    ax.xaxis.set_major_formatter(_thousands)
    ax.set_xlabel("nameplate capacity offline (kbd) - per unit, never summed")
    ax.set_xlim(0, xmax * 1.5)
    present = [p for p in PADDS if p in set(ev["padd"])]
    handles = [plt.Rectangle((0, 0), 1, 1, color=PADD_COLOR[p]) for p in present]
    ax.legend(handles, present, frameon=False, ncol=len(present), fontsize=8.5,
              loc="lower right", title="Region", title_fontsize=8.5)
    ax.set_title(f"Biggest {year} Outages by Unit - colored by PADD (where the tonnage sits)",
                 fontsize=12)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(axis="x", zorder=0)
    ax.grid(axis="y", visible=False)
    ax.tick_params(length=0)
    if bool(ev["indic"].any()):
        fig.text(0.5, 0.004, "* non-Exxon H2 2027 - still being booked (indicative floor, not confirmed).",
                 ha="center", fontsize=7.5, color=RED, style="italic")
    return _save(fig, path)


# bright, high-contrast year palette for the cross-year comparison bars
YEAR_COLOR = {2025: "#5B9BD5", 2026: "#FFC000", 2027: "#ED7D31"}


def h1_monthly_by_unit(ctx, path):
    """H1 (Jan-Jun) planned offline per focus unit, BY MONTH - 2x2 small multiples,
    one panel per unit, grouped bars per month (one bright bar per year,
    2025/26/27). Day-weighted kbd, planned only; per-panel y-axis so each unit's
    monthly shape is readable. Like-for-like (2027 confirmed through H1)."""
    fp = ctx["focus_planned"]
    H1 = MONTHS[:6]
    years = [2025, 2026, 2027]
    fig, axes = plt.subplots(2, 2, figsize=(11.4, 6.4))
    x = np.arange(6)
    w = 0.82 / len(years)
    for idx, (ax, f) in enumerate(zip(axes.reshape(-1), engine.FOCUS_ORDER)):
        m = fp[f]
        for i, y in enumerate(years):
            vals = [float(m.loc[y, mo]) if y in m.index else 0.0 for mo in H1]
            ax.bar(x + i * w, vals, w, color=YEAR_COLOR[y], label=f"H1 {y}",
                   edgecolor="white", linewidth=0.4, zorder=3)
        ax.set_xticks(x + w, H1, fontsize=8.5)
        ax.yaxis.set_major_formatter(_thousands)
        ax.set_ylim(bottom=0)
        if idx % 2 == 0:
            ax.set_ylabel("kbd offline", fontsize=9)
        if idx >= 2:
            ax.set_xlabel("Month", fontsize=9)
        ax.set_title(engine.FOCUS_LABEL[f], fontsize=10.5)
        _clean(ax)
    handles = [plt.Rectangle((0, 0), 1, 1, color=YEAR_COLOR[y]) for y in years]
    fig.legend(handles, [f"H1 {y}" for y in years], loc="lower center", ncol=3,
               frameon=False, fontsize=10.5)
    fig.suptitle("H1 Planned Offline by Unit & Month - 2025 / 2026 / 2027 (kbd, day-weighted)",
                 fontsize=13, fontweight="bold", color=NAVY)
    fig.tight_layout(rect=[0, 0.05, 1, 0.95])
    fig.savefig(path, dpi=200, bbox_inches="tight")
    plt.close(fig)
    return path


def render_all(ctx, outdir):
    """Render every deck chart into outdir; return a dict name -> path.

    The deck is built around per-unit capacity offline (2021+): the focus is the
    four key units (CDU, FCC, hydrocracker, reformer) read per unit and by month
    and PADD, plus the verified ExxonMobil per-unit view. Margin/$ and aggregate
    'total offline' charts are intentionally not rendered here."""
    import os
    os.makedirs(outdir, exist_ok=True)
    p = lambda n: os.path.join(outdir, n)
    out = {
        # 1) total outages by unit (2027, confirmed vs not-yet-booked)
        "splits_2027": splits_2027(ctx, p("splits_2027.png")),
        # 1b) what's driving the numbers - biggest individual outages, by PADD
        "biggest_outages": biggest_outages(ctx, p("biggest_outages.png")),
        # 1c) H1 like-for-like: planned offline per unit & month, 2025 vs 2026 vs 2027
        "h1_month_by_unit": h1_monthly_by_unit(ctx, p("h1_month_by_unit.png")),
        # 2) outages by PADD by unit
        "cdu_padd_27": focus_padd_bars(ctx, "CDU", 2027, p("cdu_padd_27.png")),
        "fcc_padd_27": focus_padd_bars(ctx, "FCC", 2027, p("fcc_padd_27.png")),
        # 3) ExxonMobil outages (per unit, verified)
        "exxon_gantt": exxon_gantt(ctx, p("exxon_gantt.png")),
        # 4) 2027 unplanned scenario analysis
        "fan": scenario_fan_chart(ctx, p("fan.png")),
        "scenario_total": scenario_total_bars(ctx, p("scenario_total.png")),
    }
    return out


def _render_all_legacy(ctx, outdir):
    """Previous (aggregate/margin) chart set - kept for reference, not used by the
    per-unit deck."""
    import os
    os.makedirs(outdir, exist_ok=True)
    p = lambda n: os.path.join(outdir, n)
    out = {
        "annual": annual_stack(ctx, p("annual.png")),
        "padd_clustered": padd_clustered(ctx, p("padd_clustered.png")),
        "seasonality": seasonality(ctx, p("seasonality.png")),
        "padd3": padd_combo(ctx, "PADD 3", p("padd3.png")),
        "padd_sm": padd_small_multiples(ctx, ["PADD 1", "PADD 2", "PADD 4", "PADD 5"],
                                        p("padd_sm.png")),
        "units": units_bar(ctx, p("units.png")),
        "fcc": fcc_cluster_strip(ctx, p("fcc.png")),
        "scenario": scenario_lines(ctx, p("scenario.png")),
        "scenario_padd": scenario_by_padd(ctx, p("scenario_padd.png")),
        "heatmap": sensitivity_heatmap(ctx, p("heatmap.png")),
        "scenario_total": scenario_total_bars(ctx, p("scenario_total.png")),
        "season_band": seasonality_band(ctx, p("season_band.png")),
        "yoy_month": monthly_yoy_bars(ctx, p("yoy_month.png")),
        "mogas": mogas_annual_chart(ctx, p("mogas.png")),
        "padd1": padd_combo(ctx, "PADD 1", p("padd1.png")),
        "padd2": padd_combo(ctx, "PADD 2", p("padd2.png")),
        "padd5": padd_combo(ctx, "PADD 5", p("padd5.png")),
        "planned_xyear": planned_cross_year(ctx, p("planned_xyear.png")),
        "exxon27": exxon_2027_chart(ctx, p("exxon27.png")),
        "fan": scenario_fan_chart(ctx, p("fan.png")),
        "dollar": dollar_at_risk(ctx, p("dollar.png")),
        "padd_pair": padd_planned_pair(ctx, p("padd_pair.png")),
        "dollar_month": dollar_monthly_profile(ctx, p("dollar_month.png")),
        "exxon_unit": exxon_by_unit(ctx, p("exxon_unit.png")),
        "padd_all": padd_small_multiples(ctx, list(engine.PADD_ORDER), p("padd_all.png")),
        "fcc_year": fcc_by_year(ctx, p("fcc_year.png")),
        "mom_movers": mom_movers_chart(ctx, p("mom_movers.png")),
        "mom_trend": mom_trend_chart(ctx, p("mom_trend.png")),
        "mom_newgone": mom_newgone_chart(ctx, p("mom_newgone.png")),
        "padd_pl": {pd: padd_planned_27v26(ctx, pd, p(f"padd_pl_{pd[-1]}.png")) for pd in PADDS},
    }
    return out


if __name__ == "__main__":
    import sys
    from pathlib import Path
    default = str(Path(__file__).resolve().parent.parent / "data" / "Refinery_Outages_Enhanced.xlsx")
    ctx = engine.build_context(sys.argv[1] if len(sys.argv) > 1 else default)
    paths = render_all(ctx, "scratch_charts")
    for k, v in paths.items():
        print(f"{k}: {v}")
