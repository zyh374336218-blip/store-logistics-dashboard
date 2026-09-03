# Store Logistics Performance Dashboard (local)

## How to use (hosting mode A)

1. Update data: replace `datasource/daily.xlsx`
2. Generate dashboard data:

```bash
python scripts/build_logistics_dashboard.py
```

3. Open `dashboard/index.html` in a browser (double-click or drag into Chrome/Edge)

No server required for local use. Online mirror: GitHub Pages after push.

## Rules summary

| Item | Rule |
|------|------|
| Source | `datasource/daily.xlsx` (includes `qty`, `Weight Variance`) |
| Outlier | `Weight Variance > 0.2` (taken from source as-is) |
| PID | Same PID merged across SA/AE |
| Filters | Market / Store volume / PID / PH NO / Date drive KPI and Views A–E |
| QA | Empty/`#N/A` or duplicate PH NO → build fails |
| Decimals | Display with 2 places |

## Files

- `datasource/daily.xlsx` — source data
- `scripts/build_logistics_dashboard.py` — clean + QA
- `dashboard/data/logistics.js` — frontend data (generated)
- `dashboard/index.html` — dashboard page
