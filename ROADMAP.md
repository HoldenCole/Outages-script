# Roadmap / Future Work

Captured ideas that are **not** built yet. Listed so they aren't lost between
sessions.

---

## 1. Refreshable data feed (requested — not yet built)

**Goal:** stop regenerating the workbook for every data update. Instead, point
the model at an external Excel file that holds the latest Snowflake export and
just hit **Data → Refresh All** to pull new numbers through the whole model.

**Why the current build is already set up for this.** Everything interactive in
the workbook reads from **one hidden `Data` sheet** (the tidy
`year · month · key · type · kbd · mogas` table with `Total US` / `All Units`
rollups). The Explorer's SUMIFS, the live grids, the YoY% and the charts are all
downstream of that single table. So "refreshable" means: *make the `Data` sheet
update from an external file instead of from a Python re-run.* The presentation
layer (formulas, charts, formatting, tooltips, dropdowns) never has to change.

### Recommended approach — Power Query connection to a tidy feed file

1. `build_workbook.py` gains a `--template` (or `--feed`) mode that emits **two**
   files:
   - `outage_model.xlsx` — the model **once**, with the `Data` sheet left as a
     Power Query "connection-only / loaded" table (not hardcoded values).
   - `outage_feed.xlsx` — just the tidy `Data` table (what `engine.tidy_monthly`
     already produces), regenerated whenever the Snowflake export changes.
2. One-time setup in Excel: **Get Data → From Workbook → `outage_feed.xlsx`**,
   load it to the `Data` table. From then on, dropping a fresh `outage_feed.xlsx`
   in place + **Refresh All** updates the entire model.
3. Refresh cadence: a tiny `python build_feed.py <new_export.xlsx>` re-emits only
   `outage_feed.xlsx` (fast — no charts/formatting), keeping all of `engine.py`'s
   cleaning rules (UNKNOWN→UNPLANNED, PADD parse, mogas yields) intact.

**Trade-off vs. doing the cleaning in Power Query (M):** porting `engine.py`'s
logic to M would make the workbook *fully* self-refreshing with no Python at all,
but it duplicates the (locked) cleaning rules in a second language. Keeping the
cleaning in Python + a thin feed file is lower-risk and keeps one source of truth.

### Things to verify when we build it
- SUMIFS ranges over the `Data` table should become **structured-table
  references** (e.g. `Data[kbd]`) so they auto-expand when the feed grows/shrinks
  — today they're fixed `$E$2:$E$2701` ranges sized to the current row count.
- The scenario lookup (baseline-window profiles) and the 2027-planned constants
  are currently written as values; decide whether those also come from the feed.
- Confirm Refresh preserves the dropdown selections and the scenario inputs.

---

## 2. Other parked ideas
- Named scenario presets (Base / Bull / Bear) that set the Model inputs in one click.
- "Compare two slices" mode in the Explorer (e.g. PADD 3 vs PADD 5 side by side).
- Naphtha: split light vs heavy naphtha / reformate octane barrels if the feed
  ever carries stream-level detail.
- Push the deck/dashboard off the same feed so all three refresh together.
