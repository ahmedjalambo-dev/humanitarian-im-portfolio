# Power BI dashboard

`ITCN_PowerBI_Data.xlsx` is the model produced by `python/clean_and_reconcile.py`, with one table per sheet:

| Table | Grain | Key |
|---|---|---|
| Sites | one row per site (with coordinates, households, RHU, CfW, avg progress) | site_id |
| Progress | site x component (target, achieved, % complete, status) | site_id |
| Weekly_Trend | week x component (cumulative achieved) | — |
| Reconciliation | site x component (field vs contractor, gap status) | site_id |
| CfW_Workdays | site x date x gender (workdays) | site_id |
| DQ_Log | one row per data-quality issue | — |

Relationships: `Sites[site_id]` 1 → * `Progress`, `Reconciliation`, `CfW_Workdays`.

Report pages: **Overview** (KPI cards, progress by component, map), **Sites** (table + slicers by governorate and site type), **Data quality** (open issues by dimension and dataset, reconciliation status).

Files: `ITCN_Dashboard.pbix`, plus screenshots in `docs/screenshots/`.
