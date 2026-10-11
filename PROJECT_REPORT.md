# Call Centre Performance Dashboard: Project Report

**Tools:** Microsoft Excel (Power Pivot, DAX, PivotTables, Slicers) · Power BI Desktop (Power Query, DAX, PBIP/TMDL) · Python (pandas) for validation
**Data:** 1,000 calls · 15 customers · 5 representatives · calendar year 2023 · $96,623 revenue

---

## Contents

1. [The business question](#1-the-business-question)
2. [The data](#2-the-data)
3. [Data preparation](#3-data-preparation)
4. [The data model](#4-the-data-model)
5. [Measures](#5-measures)
6. [The dashboard](#6-the-dashboard)
7. [Design decisions](#7-design-decisions)
8. [Findings](#8-findings)
9. [Data quality and validation](#9-data-quality-and-validation)
10. [Excel vs Power BI](#10-excel-vs-power-bi)
11. [Limitations and next steps](#11-limitations-and-next-steps)
12. [Appendix: reference numbers](#appendix-reference-numbers)

---

## 1. The business question

A call-centre manager running a team of five needs one screen that answers four questions:

| Question | Answered by |
|---|---|
| How busy are we, and when? | Calls by month, calls by day of week |
| Who is carrying the load? | Calls and revenue by representative |
| Are customers happy? | Average rating, 5-star share, rating distribution |
| Where does the revenue come from? | Customer × representative revenue matrix |

The intended user runs weekly one-on-ones with each rep, so the core interaction is **pick a representative and see their numbers next to the team's**.

---

## 2. The data

Two source tables:

**Calls** (fact table, one row per call, 1,000 rows)

| Column | Type | Notes |
|---|---|---|
| Call ID | Text | Unique key |
| Customer ID | Text | Links to Customers |
| Call Duration | Whole number | Minutes, range 2–176 |
| Rep ID | Text | R01–R05 |
| Call Date | Date | 1 Jan – 31 Dec 2023, 330 distinct days |
| Purchase Amount | Whole number | $20 minimum; every call has a purchase |
| Satisfaction Rating | Decimal | 0.0–5.0, one decimal place |

**Customers** (dimension, one row per customer, 15 rows)

| Column | Type | Notes |
|---|---|---|
| Customer ID | Text | Unique key |
| Gender | Text | Source header was misspelled `Gedner` |
| Age | Whole number | 22–43 |
| City | Text | Cincinnati, Cleveland, Columbus |

---

## 3. Data preparation

### Derived columns

| Column | Logic | Why |
|---|---|---|
| **Rating Rounded** | Round rating to whole stars, half away from zero | Groups ratings into 0–5 stars |
| **Duration Bucket** | 5 tiers: under 10 min, 10–30, 30–60, 1–2 h, over 2 h | Slicing by call length |
| **Day of Week** | Weekday name from the date | Weekly pattern |
| **Month** | Month name from the date | Monthly trend |
| **FY** | Year + 1 if month ≥ July | The business runs a July–June financial year |

### How each tool does it

- **Excel:** formula columns inside Excel Tables (`IFS`, `TEXT`, `ROUND`, `YEAR`/`MONTH`), so new rows classify themselves.
- **Power BI:** Power Query steps over the raw CSVs: rename headers (including the `Gedner` → `Gender` fix), set types, add the columns. The folder path is a parameter, so the same steps replay on a new extract with one refresh.

---

## 4. The data model

### Excel (Power Pivot)

```text
Customers (15)  1 ──────── *  Calls (1,000)
              on Customer ID
```

One relationship, so a single PivotTable can show calls (from `Calls`) by city (from `Customers`) without copying city onto every call row.

### Power BI (star schema)

```text
              Date (365)       Customers (15)
                  │ 1               │ 1
                  ▼ *               ▼ *
              ┌──────────────────────────┐
              │      Calls (1,000)       │
              └──────────────────────────┘
                  ▲ *               ▲ *
                  │ 1               │ 1
              Reps (5)        Rating Scale (6)

              Rep Axis (5)  ← deliberately unrelated (see §7)
```

| Table | Built with | Purpose |
|---|---|---|
| Calls | Power Query | Fact table |
| Customers | Power Query | Customer attributes |
| Reps | Power Query (reference of Calls, distinct) | Rep dimension; target of the slicer and security role |
| Date | DAX `CALENDAR` | Continuous calendar for time intelligence; month/weekday sort orders; fiscal year |
| Rating Scale | DAX `GENERATESERIES(0, 5, 1)` | Guarantees all six star levels exist on an axis |
| Rep Axis | DAX `DISTINCT(Reps[Rep ID])` | Disconnected axis for highlighting |

All relationships are one-to-many and single-direction. Foreign keys on `Calls` are hidden so visuals must use the dimension columns. Auto date/time is off.

---

## 5. Measures

### Core (identical in both tools)

```dax
Call Count     = COUNTROWS ( Calls )
Total Amount   = SUM ( Calls[Purchase Amount] )
Total Duration = SUM ( Calls[Call Duration] )
Avg Rating     = AVERAGE ( Calls[Satisfaction Rating] )
5 Star Calls   = CALCULATE ( [Call Count], Calls[Rating Rounded] = 5 )
```

### Added in Power BI (~30 measures in display folders)

| Group | Examples | Technique |
|---|---|---|
| Ratios | Revenue per Call, 5 Star %, Low Rating % (0–2 stars) | `DIVIDE` for safe division |
| Benchmarks | Calls (All Reps), Share of Calls, Rev per Call vs Team | `CALCULATE` + `REMOVEFILTERS ( Reps )` |
| Time intelligence | Calls MoM %, Revenue YTD, Revenue FYTD | `DATEADD`, `TOTALYTD ( …, "6-30" )` |
| Ranking | Rev per Call Rank | `RANKX ( ALL ( Reps ), … )` guarded by `HASONEVALUE` |
| Labels | KPI sub-lines, dynamic subtitle | `ISFILTERED`, `CONCATENATEX`, `FORMAT` |
| Highlighting | Axis Calls, Axis Bar Colour | `TREATAS` from a disconnected table; colour returned as hex text |

---

## 6. The dashboard

### Excel: one page

```text
┌────────────────────────────────────────────────────────────────┐
│ Call Center Performance                  [ R01 R02 R03 R04 R05 ]│
├──────────┬──────────┬──────────┬──────────┬────────────────────┤
│  CALLS   │ REVENUE  │TALK TIME │AVG RATING│   5-STAR CALLS     │
├────────────────────────────────────────────────────────────────┤
│ Calls by month (area + line)                                   │
├─────────────────────┬─────────────────────┬────────────────────┤
│ Calls by rep        │ Revenue by rep      │ Calls by weekday   │
├─────────────────────┴──────┬──────────────┴────────────────────┤
│ Customer mix by city (100%)│ Satisfaction rating distribution  │
├────────────────────────────┴───────────────────────────────────┤
│ Revenue by customer × rep (heatmap, selected column bold)      │
└────────────────────────────────────────────────────────────────┘
```

Layers kept separate: `Data` → `Pivots` (9 PivotTables) → `Calc` (chart feeds) → `Dashboard` → `Notes`. Nothing is calculated on the dashboard sheet.

### Power BI: two pages

**Overview:** same layout, plus a revenue-per-call KPI and a ranked leaderboard. KPI cards show the selection as the main value and the team benchmark underneath (for example, R04: `$89.52`, *Team $96.62 (−7.10)*).

**Rep Detail (drill-through):** right-click a rep in the leaderboard → *Drill through*. Shows that rep's KPIs, rank, monthly trend, rating distribution, call-length mix and customers handled.

**Row-level security:** role `Rep R02` filters the `Reps` dimension; *View as role* shows 218 calls. Under the role the team benchmark also drops to 218: `REMOVEFILTERS` cannot override a security filter, which stops a rep from deriving colleagues' figures.

---

## 7. Design decisions

| Decision | Why | Excel mechanism | Power BI mechanism |
|---|---|---|---|
| **Team totals stay fixed when a rep is selected** | A rep's 218 calls means nothing without the 1,000 next to it | 4 of 9 PivotTables left unconnected to the slicer | `REMOVEFILTERS ( Reps )` inside benchmark measures |
| **No chart ever changes shape** | A vanishing category looks like a bug and hides that the category exists | Hard-coded category lists + `GETPIVOTDATA` wrapped in `IFERROR(…, 0)` | Dimension tables + *Show items with no data* |
| **Highlight the selection, keep the rest** | Comparison needs all five bars | Second series of `NA()` values at 100% overlap | Disconnected `Rep Axis` + `TREATAS` + colour measure |
| **Detect "has the user picked something?"** | Drives highlights and captions | `COUNTIF` over a slicer-filtered PivotTable | `ISFILTERED ( Reps[Rep ID] )` |
| **Months in calendar order** | Default sort is alphabetical | Months typed in order on the helper sheet | *Sort by column* on `Date[Month Number]` |
| **Unit in the chart title, not the labels** | Chart labels localise `$` to `₹` on an India-locale machine | `#,##0.0,` label format + "$ thousands" title | Currency format pinned to en-US |

---

## 8. Findings

Each finding is stated with how far the data supports it.

| # | Finding | So what | Confidence |
|---|---|---|---|
| 1 | **Revenue per call ranges from $89.52 (R04) to $100.83 (R03)**, a 13% spread | The clearest controllable difference between reps | Medium: no cost or call-type data to explain it |
| 2 | **R04 is weakest on three signals:** fewest calls (186), longest average call (95.7 min), lowest revenue per call | Three independent signals agree, so it's unlikely to be noise | Medium-high, but needs context (call mix, tenure) before acting |
| 3 | **R02 handles the most calls (218) at the second-lowest revenue per call ($94.41)** | If reps are measured on call count, that target may be producing this | Medium |
| 4 | **Satisfaction does not separate reps:** averages span 3.86–3.92 | Don't build a training programme around rating | High that this is noise |
| 5 | **Call length doesn't affect satisfaction:** correlation +0.009; average rating 3.93 for calls under 10 min vs 3.88 for calls over 2 h | Shortening calls won't lift satisfaction | High (a null result) |
| 6 | **Demand swings 3× across the year:** March 155 calls vs August 50 (monthly average 83) | Staff to the curve rather than flat | Low-medium: one year, no seasonal baseline |
| 7 | **73.5% of calls rate 4–5 stars; 6.7% rate 1–2** | Satisfaction is healthy; track the low-rating tail rather than the average | High |
| 8 | **R02 is the most polarised rep:** highest 5-star share (33.9%) *and* highest 0–2 star share (10.1% vs team 6.8%) | The two extremes cancel out in the average | Low: 22 low-rated calls |

**Headline recommendation:** manage on revenue per call, not satisfaction. Investigate what R03 and R05 do differently, and get call-type and tenure context on R04 before drawing conclusions.

---

## 9. Data quality and validation

### Reconciliation rule

Every breakdown must sum back to its headline: rep calls = 1,000, rep revenue = $96,623, rating distribution = 1,000, matrix grand total = $96,623. Both builds were checked against these totals at every stage, and the Power BI measures were cross-checked against an independent pandas calculation.

### Issues found

| Issue | How it showed up | Impact | Fix |
|---|---|---|---|
| **Missing 0-star category** | Rating chart built on a 1–5 axis summed to 999 | One call (Call_7340, R03) silently excluded | Axis covers 0–5 in both tools |
| **Banker's rounding in Power Query** | 5-star calls came out at 248 instead of 307, with no error | 100 ratings end in .5; the default *Round* step sends 4.5 → 4, moving 100 calls between categories | `Number.Round ( …, 0, RoundingMode.AwayFromZero )`, matching Excel and DAX |
| **Misspelled source header** | `Gedner` in the customer extract | Cosmetic, but breaks anything referencing `Gender` | Renamed in a Power Query step, so it's fixed on every refresh |

### Characteristics that suggest a generated dataset

- Average call length is **89.9 minutes**, and 76% of calls exceed an hour.
- **Every call has a purchase** (minimum $20).
- Customer call counts are flat: **50–82 calls each**, no long tail.

These don't block the analysis, but the numbers shouldn't be read as a real call centre's behaviour.

---

## 10. Excel vs Power BI

| | Excel | Power BI |
|---|---|---|
| **Best for** | One manager, monthly file, no licence or login needed | Many viewers, scheduled refresh, per-user security |
| **Cleaning** | Manual header fixes + formula columns | Recorded Power Query steps |
| **Benchmarks** | Configured in *Report Connections* (invisible to the next maintainer) | Written into the measure (self-documenting) |
| **Security** | Anyone with the file sees every row | Row-level security |
| **DAX** | Same engine; the 5 core measures carry over unchanged | Same, plus time intelligence and ranking |

Building it twice separated the **design** (fixed benchmarks, stable categories, highlight-not-hide) from the **mechanism**, which differs by tool.

---

## 11. Limitations and next steps

**Limitations**
- One year of data, so no year-on-year comparison or seasonal baseline.
- No call type, resolution, wait time, cost or rep tenure data, so performance gaps can be measured but not explained.
- Sample dataset (see §9).
- The Power BI data path is a local folder parameter and must be updated on another machine.

**Next steps**
1. Add a target line or prior-period comparison so every KPI has a reference point.
2. Add call type and resolution data to explain the revenue-per-call gap.
3. Dynamic RLS: an email column on `Reps` matched to `USERPRINCIPALNAME()`.
4. Publish to the Power BI Service with scheduled refresh through a gateway.

---

## Appendix: reference numbers

### Headline

| Metric | Value |
|---|---|
| Calls | 1,000 |
| Revenue | $96,623 |
| Talk time | 89,850 min (1,498 h) |
| Average rating | 3.89 |
| 5-star calls | 307 (30.7%) |
| Revenue per call | $96.62 |

### By representative

| Rep | Calls | Revenue | Rev / call | Avg min | Avg rating | 5★ % | 0–2★ % |
|---|---|---|---|---|---|---|---|
| R01 | 189 | $18,415 | $97.43 | 89.1 | 3.92 | 31.7% | 5.3% |
| R02 | 218 | $20,581 | $94.41 | 88.5 | 3.87 | 33.9% | 10.1% |
| R03 | 207 | $20,872 | $100.83 | 85.5 | 3.86 | 29.0% | 6.3% |
| R04 | 186 | $16,651 | $89.52 | 95.7 | 3.90 | 33.3% | 7.0% |
| R05 | 200 | $20,104 | $100.52 | 91.0 | 3.88 | 25.5% | 5.0% |

### Distributions

| Cut | Values |
|---|---|
| Months (Jan–Dec) | 79, 66, 155, 136, 83, 67, 64, 50, 76, 114, 57, 53 |
| Weekdays (Sun–Sat) | 146, 133, 138, 153, 128, 141, 161 |
| Ratings (0–5 ★) | 1, 8, 59, 197, 428, 307 |
| Duration buckets | <10 min 9 · 10–30 min 51 · 30–60 min 179 · 1–2 h 524 · >2 h 237 |
| Cities (calls) | Cincinnati 276 · Cleveland 389 · Columbus 335 |
| Fiscal year (Jul–Jun) | FY2023 586 calls · FY2024 414 calls |
| Revenue YTD / FYTD at Dec | $96,623 / $40,548 |
