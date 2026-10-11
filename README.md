# Call Centre Performance Dashboard: Excel & Power BI

An interactive performance dashboard for a five-person call centre, built twice: once in **Excel (Power Pivot + DAX)** and once in **Power BI**. A manager picks a representative and the whole page re-reads for that person, while team benchmarks stay visible for comparison.

**Data:** 1,000 calls in 2023 · 15 customers · 5 representatives · $96,623 revenue *(sample dataset)*

📄 **Full write-up:** [PROJECT_REPORT.md](PROJECT_REPORT.md)

---

## What's in it

| | Excel version | Power BI version |
|---|---|---|
| **Data prep** | Excel Tables with formula columns | Power Query: parameterised, refreshable steps |
| **Model** | Power Pivot: 2 related tables | Star schema: 1 fact + 4 dimensions (incl. Date table) |
| **Measures** | 5 DAX measures | ~30 DAX measures (benchmarks, time intelligence, ranking) |
| **Interactivity** | 1 slicer driving 9 PivotTables and 6 charts | Rep slicer, cross-highlighting, drill-through page |
| **Security** | n/a | Row-level security (a rep sees only their own calls) |

### Design decisions

- **Benchmarks don't move.** When a rep is selected, team totals stay fixed so the selection has context. In Excel this is done by leaving 4 PivotTables unconnected to the slicer; in Power BI by `REMOVEFILTERS(Reps)` inside the benchmark measures.
- **Categories never disappear.** Every chart axis comes from a fixed list (Excel helper sheet / Power BI dimension tables), so a rep with no 0-star calls still shows a 0-star column.
- **Highlight, don't hide.** The selected rep's bar is coloured while the other four stay grey: an `NA()` overlay series in Excel, a disconnected table + `TREATAS` + conditional-format measure in Power BI.

## Findings

| Finding | Confidence |
|---|---|
| **Revenue per call** is where reps differ: $89.52 (R04) to $100.83 (R03), a 13% spread | Medium: no cost or call-type data |
| **Satisfaction does not separate reps**: every average sits between 3.86 and 3.92 | High that it's noise |
| Call duration vs satisfaction correlation is **+0.009**: no relationship | High (null result) |
| Volume swings **3×** across the year: March 155 calls vs August 50 | Low: one year, no seasonal baseline |

## Data quality issues caught

1. **A rating chart silently dropping 1 of 1,000 calls.** One call rounds to a 0-star rating; a 1–5 axis hid it. Caught by reconciling every breakdown back to the total.
2. **Banker's rounding in Power Query.** The default *Round* step rounds 4.5 down to 4, unlike Excel and DAX. With 100 ratings ending in .5, five-star calls came out at 248 instead of 307, with no error. Fixed with `RoundingMode.AwayFromZero`.

## Repository layout

```text
PROJECT_REPORT.md                        Full write-up: data, model, design, findings, validation
Call Center Dashboard.xlsx               Excel dashboard (Power Pivot model)
Power BI/
  Call Center Dashboard.pbip             Open this in Power BI Desktop
  Call Center Dashboard.SemanticModel/   Data model (TMDL: tables, measures, relationships, RLS)
  Call Center Dashboard.Report/          Report pages and visuals (PBIR)
  data/                                  Raw source CSVs (original messy headers)
  call-center-theme.json                 Colour theme
```

## Opening the Power BI version

1. Open `Power BI/Call Center Dashboard.pbip` in Power BI Desktop.
2. **Transform data → Edit parameters** → set `DataFolder` to the full path of the `Power BI/data/` folder on your machine (keep the trailing `\`).
3. Click **Refresh**.

## Limitations

This is a generated sample dataset and doesn't behave like a real call centre (average call ~90 minutes, every call ends in a purchase). The structure and method are the deliverable; the numbers would need real data before driving decisions.
