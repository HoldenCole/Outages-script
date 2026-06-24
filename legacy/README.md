# Legacy prototype

These three scripts are the original `openpyxl` prototype, kept for reference
only. They are **superseded** by the current pipeline in the repository root
(`engine.py`, `build_workbook.py`, `build_slides.py`, `build_dashboard.py`,
`build_all.py`).

| File | Was |
|---|---|
| `outage_workbook.py` | First-pass workbook builder (openpyxl). Imports `outage_monthly`. |
| `outage_monthly.py` | Monthly KBD aggregation engine used by the prototype. |
| `outage_analyzer.py` | Standalone HTML/analysis script. |

They still run from this folder (`python legacy/outage_workbook.py ...`), but the
charts/heatmaps render flat in openpyxl — which is why the production build moved
to `XlsxWriter`. Use the root pipeline instead.
