# Call Center Performance Dashboard — Power BI Build Guide

**The same project, rebuilt in Power BI Desktop, from the raw CSV files.**

Same rules as the Excel guide: *do this → then this → check this number*.
If your number matches the ✅ checkpoint, move on. If not, stop and fix it there.

> Companion files:
> - `../Call Center Dashboard - Build Guide.md`: the Excel version. Build that first; this guide assumes you know *why* the design works.
> - `data/calls.csv`, `data/customers.csv`: the raw source extract, with the **original messy headers** (including the `Gedner` typo).
> - `call-center-theme.json`: the colour palette as a Power BI theme.

---

## Contents

| Part | What you build | Time |
|---|---|---|
| [0](#part-0--the-brief) | Brief, setup, what's different from Excel | 10 min |
| [1](#part-1--power-query) | Power Query: load, clean, derive columns | 30 min |
| [2](#part-2--the-star-schema) | Date table, rating table, relationships | 20 min |
| [3](#part-3--dax-measures) | ~30 DAX measures in 6 groups | 40 min |
| [4](#part-4--the-overview-page) | Overview page: cards, charts, matrix | 60 min |
| [5](#part-5--the-highlight-trick) | Selected-rep highlighting | 20 min |
| [6](#part-6--rep-detail-drill-through-page) | Drill-through page | 30 min |
| [7](#part-7--row-level-security) | Row-level security | 10 min |
| [8](#part-8--test-it) | Acceptance tests | 15 min |
| [9](#part-9--troubleshooting) | When it breaks | n/a |
| [10](#part-10--portfolio--publishing) | Screenshots, PDF, GitHub | 15 min |
| [A](#appendix--every-number-in-this-project) | Every correct number | n/a |

---

# Part 0 — The Brief

## What you are building

A two-page Power BI report on the same 1,000 calls:

1. **Overview**: the Excel dashboard, rebuilt. One rep slicer drives the page, and the team benchmarks stay visible.
2. **Rep Detail**: right-click any rep → *Drill through* → a page just for that person.

Plus **row-level security**, so a rep who opens the report only ever sees their own calls.

## What's different from the Excel version (read this; interviewers ask about it)

| Problem | Excel solution | Power BI solution |
|---|---|---|
| Clean the source data | Rename headers by hand, formula columns | **Power Query**: recorded steps that replay on every refresh |
| Relate calls to customers | One relationship in the Data Model | **Star schema**: 1 fact + 4 dimension tables |
| Keep company totals fixed when a rep is picked | 4 pivots deliberately **not connected** to the slicer | **DAX**: `REMOVEFILTERS(Reps)` inside the benchmark measures |
| Stop categories disappearing (the 0-star bar) | Hard-coded lists on `Calc` + `GETPIVOTDATA` | **Dimension tables** + *Show items with no data* |
| Month order (Jan, Feb… not Apr, Aug…) | Type the months by hand | **Sort by column** on a Date table |
| Highlight the selected rep's bar | `NA()` helper series at 100% overlap | **Disconnected table** + `TREATAS` + conditional-format colour measure |
| Per-rep deep dive | n/a | **Drill-through** page |
| Who can see what | Anyone with the file sees everything | **Row-level security** |
| Financial year (Jul–Jun) | Formula column | `TOTALYTD(..., "6-30")` time intelligence |

The **thinking** is identical. Only the mechanism changes. That's the point you make in an interview:
*"I solved the same design problems twice, with each tool's native mechanism."*

## The model you are building

```text
        ┌───────────┐        ┌────────────┐
        │   Date    │        │ Customers  │
        │ (365 rows)│        │ (15 rows)  │
        └─────┬─────┘        └─────┬──────┘
              │ 1                  │ 1
              │                    │
              ▼ *                  ▼ *
        ┌──────────────────────────────────┐
        │              Calls               │   ← fact table, 1,000 rows
        └──────────────────────────────────┘
              ▲ *                  ▲ *
              │                    │
              │ 1                  │ 1
        ┌─────┴─────┐        ┌─────┴──────┐
        │   Reps    │        │Rating Scale│
        │ (5 rows)  │        │  (6 rows)  │
        └───────────┘        └────────────┘

        ┌───────────┐
        │ Rep Axis  │   ← deliberately NOT related (Part 5)
        └───────────┘
```

## The numbers you are aiming for

Same as the Excel build, plus four new ones Power BI makes easy:

| Metric | Value |
|---|---|
| Calls | **1,000** |
| Revenue | **$96,623** |
| Talk time | **1,498 hrs** |
| Average rating | **3.89** |
| 5-star calls | **307** (30.7%) |
| **Revenue per call** | **$96.62** ← new |
| **0–2 star share** | **6.8%** (68 calls) ← new |
| **FYTD revenue at Dec-2023** | **$40,548** ← new (FY runs Jul–Jun) |
| **Rev/call rank #1** | **R03** ($100.83) ← new |

## Before you start

1. Install **Power BI Desktop** (free, Microsoft Store, Windows only).
2. Open it → **File → Options and settings → Options**:
   - **Current File → Data Load** → ❌ untick **Auto date/time**
   - **Global → Regional Settings** → leave locale as is, but note it; Part 9 covers the locale traps
3. **File → Save As** → `Call Center Dashboard.pbix` in this `Power BI` folder.

> **Why turn off Auto date/time?** Power BI otherwise builds a hidden date table for *every* date column. You're building one proper Date table in Part 2, and the hidden ones bloat the file and confuse time intelligence. Turning it off is standard practice.

---

# Part 1 — Power Query

**The idea:** in Excel you fixed the headers by hand once. In Power Query every fix is a **recorded step**. Drop in next month's CSV, hit Refresh, and all the cleaning replays itself.

## 1.1 Create a folder parameter

So the file isn't hard-wired to one path:

1. **Home → Transform data** (opens Power Query Editor)
2. **Home → Manage Parameters → New Parameter**
   - Name: `DataFolder`
   - Type: **Text**
   - Current Value: `C:\Users\Aravind Mahendra\Desktop\Excel\Power BI\data\` ← **keep the trailing `\`**
3. OK

## 1.2 Load and clean `Calls`

1. **Home → New Source → Text/CSV** → pick `data\calls.csv` → **OK**
2. Rename the query (right pane, *Query Settings → Name*) to **`Calls`**
3. Rename the headers (double-click each header):

| Source header | Rename to |
|---|---|
| Call number | `Call ID` |
| Customer ID | *(keep)* |
| Duration | `Call Duration` |
| Representative | `Rep ID` |
| Date of Call | `Call Date` |
| Purchase Amount | *(keep)* |
| Satisfaction Rating | *(keep)* |

4. Set data types (click the icon left of each header):

| Column | Type |
|---|---|
| Call ID, Customer ID, Rep ID | **Text** |
| Call Duration, Purchase Amount | **Whole Number** |
| Call Date | **Date** ← not Date/Time |
| Satisfaction Rating | **Decimal Number** |

### 1.2a Rating Rounded ⚠️ the rounding trap

**Do not use** *Add Column → Rounding → Round*. It looks right and is wrong.

Instead: **Add Column → Custom Column**
- Name: `Rating Rounded`
- Formula:

```m
Number.Round([Satisfaction Rating], 0, RoundingMode.AwayFromZero)
```

Then set its type to **Whole Number**.

**Why:** Power Query's `Number.Round` defaults to **banker's rounding** (round half to *even*), so 4.5 → **4**, not 5. Excel's `ROUND` and DAX's `ROUND` both round half *away from zero* (4.5 → 5). This dataset has **100 ratings ending in .5**, so the default silently moves 100 calls:

| Rating | Correct (AwayFromZero) | Wrong (default Round button) |
|---|---|---|
| 0 | 1 | 1 |
| 1 | 8 | 8 |
| 2 | 59 | **66** |
| 3 | 197 | **190** |
| 4 | 428 | **487** |
| 5 | **307** | **248** |

Every total still says 1,000. Nothing errors. Your 5-star KPI is just wrong by 59. **This is a great interview story; see the interview prep file.**

### 1.2b Duration Bucket

**Add Column → Custom Column** → Name `Duration Bucket`:

```m
if [Call Duration] <= 10 then "Under 10 mins"
else if [Call Duration] <= 30 then "10 to 30 mins"
else if [Call Duration] <= 60 then "30 to 60 mins"
else if [Call Duration] <= 120 then "1 to 2 hours"
else "More than 2 hours"
```

Type: **Text**.

Add one more custom column, **`Bucket Order`**, for sorting (type **Whole Number**):

```m
if [Call Duration] <= 10 then 1
else if [Call Duration] <= 30 then 2
else if [Call Duration] <= 60 then 3
else if [Call Duration] <= 120 then 4
else 5
```

### 1.2c Point the source at the parameter

In the *Applied Steps* list, click the gear ⚙ on **Source** → switch the file path box from *Basic* to *Advanced* → replace the path with: parameter `DataFolder` + text `calls.csv`. OK.

> **Shortcut:** instead of 1.2–1.2c, open **Home → Advanced Editor**, delete everything, and paste this. It does every step above:
>
> ```m
> let
>     Source = Csv.Document(File.Contents(DataFolder & "calls.csv"),
>         [Delimiter = ",", Columns = 7, Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
>     Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
>     Renamed = Table.RenameColumns(Promoted, {
>         {"Call number", "Call ID"}, {"Duration", "Call Duration"},
>         {"Representative", "Rep ID"}, {"Date of Call", "Call Date"}}),
>     Typed = Table.TransformColumnTypes(Renamed, {
>         {"Call ID", type text}, {"Customer ID", type text},
>         {"Call Duration", Int64.Type}, {"Rep ID", type text},
>         {"Call Date", type date}, {"Purchase Amount", Int64.Type},
>         {"Satisfaction Rating", type number}}),
>     AddRating = Table.AddColumn(Typed, "Rating Rounded",
>         each Number.Round([Satisfaction Rating], 0, RoundingMode.AwayFromZero), Int64.Type),
>     AddBucket = Table.AddColumn(AddRating, "Duration Bucket", each
>         if [Call Duration] <= 10 then "Under 10 mins"
>         else if [Call Duration] <= 30 then "10 to 30 mins"
>         else if [Call Duration] <= 60 then "30 to 60 mins"
>         else if [Call Duration] <= 120 then "1 to 2 hours"
>         else "More than 2 hours", type text),
>     AddOrder = Table.AddColumn(AddBucket, "Bucket Order", each
>         if [Call Duration] <= 10 then 1
>         else if [Call Duration] <= 30 then 2
>         else if [Call Duration] <= 60 then 3
>         else if [Call Duration] <= 120 then 4
>         else 5, Int64.Type)
> in
>     AddOrder
> ```
>
> Still *read* the manual steps: you'll be asked what each one does.

## 1.3 Load and clean `Customers`

1. **New Source → Text/CSV** → `data\customers.csv` → rename query to **`Customers`**
2. Rename **`Gedner` → `Gender`** (the typo from the source)
3. Types: Customer ID / Gender / City → **Text**, Age → **Whole Number**
4. Point Source at `DataFolder & "customers.csv"` as in 1.2c

Advanced Editor version:

```m
let
    Source = Csv.Document(File.Contents(DataFolder & "customers.csv"),
        [Delimiter = ",", Columns = 4, Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Renamed = Table.RenameColumns(Promoted, {{"Gedner", "Gender"}}),
    Typed = Table.TransformColumnTypes(Renamed, {
        {"Customer ID", type text}, {"Gender", type text},
        {"Age", Int64.Type}, {"City", type text}})
in
    Typed
```

## 1.4 Build the `Reps` dimension from the fact table

There's no rep file, so derive one:

1. Right-click the **`Calls`** query → **Reference** → rename the new query **`Reps`**
2. Right-click the `Rep ID` header → **Remove Other Columns**
3. **Home → Remove Rows → Remove Duplicates**
4. Sort `Rep ID` ascending

> **Reference, not Duplicate.** *Reference* starts from `Calls`' finished output, so if `Calls` changes, `Reps` follows. *Duplicate* copies the steps and the two drift apart.

## 1.5 Profile the data before you trust it

**View** tab → tick **Column quality**, **Column distribution**, **Column profile**.
Then at the bottom-left status bar click *"Column profiling based on top 1000 rows"* → switch to **entire data set**.

✅ **Checkpoint:**

| Query | Rows | Check |
|---|---|---|
| `Calls` | 1,000 | Every column: **Valid 100%**, Error 0%, Empty 0% |
| `Calls[Call ID]` | n/a | 1,000 **distinct**, 1,000 **unique** (it's a real key) |
| `Calls[Call Duration]` | n/a | min **2**, max **176** |
| `Customers` | 15 | `Customer ID` 15 distinct, 15 unique |
| `Reps` | 5 | R01 … R05 |

> The profile is also where you'd spot the dataset's oddities: **every call has a purchase** (Purchase Amount min > 0) and the average call is ~90 minutes. Note them; don't "fix" them.

## 1.6 Load

**Home → Close & Apply.**

✅ **Checkpoint:** the *Data* pane (right) shows `Calls`, `Customers`, `Reps`. The parameter doesn't load as a table; that's correct.

---

# Part 2 — The Star Schema

Go to **Model view** (third icon on the left rail).

## 2.1 Delete any auto-detected relationships

Power BI may have guessed some. Click each line → **Delete**. You'll create them deliberately in 2.5 so you know every one.

## 2.2 The Date table

**Modeling → New table**:

```dax
Date =
ADDCOLUMNS (
    CALENDAR (
        DATE ( YEAR ( MIN ( Calls[Call Date] ) ), 1, 1 ),
        DATE ( YEAR ( MAX ( Calls[Call Date] ) ), 12, 31 )
    ),
    "Year", YEAR ( [Date] ),
    "Month Number", MONTH ( [Date] ),
    "Month", FORMAT ( [Date], "mmm" ),
    "Quarter", "Q" & QUARTER ( [Date] ),
    "Day Number", WEEKDAY ( [Date], 1 ),
    "Day of Week", FORMAT ( [Date], "dddd" ),
    "FY", "FY" & ( YEAR ( [Date] ) + IF ( MONTH ( [Date] ) >= 7, 1, 0 ) )
)
```

Then:

1. **Table tools → Mark as date table** → column **`Date`** → OK
2. Click the `Month` column → **Column tools → Sort by column → `Month Number`**
3. Click `Day of Week` → **Sort by column → `Day Number`**
4. Format the `Date` column as `dd-mmm-yyyy`

✅ **Checkpoint:** 365 rows (1-Jan-2023 → 31-Dec-2023).

**Why 365 when calls only happen on 330 distinct days?** Time-intelligence functions (`DATEADD`, `TOTALYTD`) need an **unbroken** calendar. A Date table built from the call dates would have 35 holes and give wrong answers. Full years, no gaps, that's the rule.

**Why a Date table at all, when Calls already has a date?** Because it's where Month, Quarter, FY and weekday live **once**, with correct sort orders, and it guarantees all 12 months exist on an axis even when a filter leaves some empty.

## 2.3 The Rating Scale table

**Modeling → New table**:

```dax
Rating Scale =
SELECTCOLUMNS (
    GENERATESERIES ( 0, 5, 1 ),
    "Rating", [Value],
    "Rating Label", [Value] & " ★"
)
```

Sort `Rating Label` by `Rating`.

✅ **Checkpoint:** 6 rows: 0 ★ through 5 ★. **Six, not five.** Same trap as Excel Part 1.5: one call (Call_7340, rep **R03**) rates 0.

## 2.4 The Rep Axis table (disconnected, on purpose)

**Modeling → New table**:

```dax
Rep Axis = DISTINCT ( Reps[Rep ID] )
```

**Do not relate this to anything.** Part 5 explains why. If Power BI auto-creates a line to it, delete the line.

## 2.5 Create the relationships

Drag from the **dimension** column onto the **fact** column (or **Modeling → Manage relationships → New**):

| From (one side) | To (many side) | Cardinality | Cross-filter |
|---|---|---|---|
| `Customers[Customer ID]` | `Calls[Customer ID]` | One-to-many | Single |
| `Reps[Rep ID]` | `Calls[Rep ID]` | One-to-many | Single |
| `Date[Date]` | `Calls[Call Date]` | One-to-many | Single |
| `Rating Scale[Rating]` | `Calls[Rating Rounded]` | One-to-many | Single |

**Single** direction, every time. Filters flow *from* dimensions *into* the fact table. Bi-directional filtering is a common beginner fix that later creates ambiguous paths and wrong numbers.

## 2.6 Hide the foreign keys

In `Calls`, right-click → **Hide in report view**: `Customer ID`, `Rep ID`, `Call Date`, `Rating Rounded`.

Now anyone building a visual *must* use `Reps[Rep ID]`, `Date[Month]` and so on, the dimension columns that sort and filter correctly. This is the difference between a model and a pile of tables.

## 2.7 A home for the measures

**Home → Enter data** → name the table **`_Measures`**, leave the one column empty → **Load**. You'll put every measure here. (Once it holds a measure, delete its `Column1` and the table moves to the top of the Data pane with a calculator icon.)

✅ **Part 2 checkpoint:** Model view shows 4 dimensions each with a `1 → *` line into `Calls`, plus `Rep Axis` and `_Measures` floating alone.

---

# Part 3 — DAX Measures

Select `_Measures` → **Modeling → New measure** → type → Enter. Then set the format in **Measure tools**.

> Format strings: for currency, pick **Currency → $ English (United States)** from the Measure tools dropdown rather than typing `$`. On an India-locale machine a typed `$` can localise to `₹`, the same trap as the Excel chart labels.

## 3.1 Base measures (same five as Excel)

```dax
Call Count = COUNTROWS ( Calls )

Total Amount = SUM ( Calls[Purchase Amount] )

Total Duration (min) = SUM ( Calls[Call Duration] )

Avg Rating = AVERAGE ( Calls[Satisfaction Rating] )

5 Star Calls = CALCULATE ( [Call Count], Calls[Rating Rounded] = 5 )
```

| Measure | Format |
|---|---|
| Call Count | Whole number, thousands separator |
| Total Amount | Currency $, 0 decimals |
| Total Duration (min) | Whole number, thousands separator |
| Avg Rating | Decimal, 2 places |
| 5 Star Calls | Whole number |

## 3.2 Ratios, the metrics Excel didn't have

```dax
Talk Time (hrs) = DIVIDE ( [Total Duration (min)], 60 )

Revenue per Call = DIVIDE ( [Total Amount], [Call Count] )

5 Star % = DIVIDE ( [5 Star Calls], [Call Count] )

Low Rating % =
DIVIDE (
    CALCULATE ( [Call Count], Calls[Rating Rounded] <= 2 ),
    [Call Count]
)
```

Formats: Talk Time `#,0` · Revenue per Call Currency 2 dp · both % measures Percentage 1 dp.

**Always `DIVIDE`, never `/`.** `DIVIDE` returns blank on a zero denominator instead of an error, so an empty filter shows nothing instead of breaking the visual.

## 3.3 Team benchmarks, replacing Excel's "unconnected pivots"

In Excel you kept company totals fixed by *not connecting* four pivots to the slicer. In Power BI you do it **inside the measure**: remove the rep filter, whatever the slicer says.

```dax
Calls (All Reps) = CALCULATE ( [Call Count], REMOVEFILTERS ( Reps ) )

Revenue (All Reps) = CALCULATE ( [Total Amount], REMOVEFILTERS ( Reps ) )

Rev per Call (All Reps) = CALCULATE ( [Revenue per Call], REMOVEFILTERS ( Reps ) )

Avg Rating (All Reps) = CALCULATE ( [Avg Rating], REMOVEFILTERS ( Reps ) )

5 Star % (All Reps) = CALCULATE ( [5 Star %], REMOVEFILTERS ( Reps ) )

Share of Calls = DIVIDE ( [Call Count], [Calls (All Reps)] )

Share of Revenue = DIVIDE ( [Total Amount], [Revenue (All Reps)] )

Rev per Call vs Team = [Revenue per Call] - [Rev per Call (All Reps)]
```

**Why this is better than Excel's approach:** it's explicit. Six months later nobody has to open *Report Connections* to discover which pivots were left unticked and why. The rule is written in the measure.

**`REMOVEFILTERS(Reps)` only removes the rep filter.** A month or city filter still applies, so "team" always means *team, in the same period*. That's the benchmark you want.

## 3.4 Card sub-lines and the dynamic subtitle

These are text measures that show the benchmark under each KPI:

```dax
Calls Sub =
IF (
    ISFILTERED ( Reps[Rep ID] ),
    FORMAT ( [Share of Calls], "0.0%" ) & " of " & FORMAT ( [Calls (All Reps)], "#,0" ) & " team calls",
    "All 5 reps"
)

Revenue Sub =
IF (
    ISFILTERED ( Reps[Rep ID] ),
    FORMAT ( [Share of Revenue], "0.0%" ) & " of team revenue",
    "All 5 reps"
)

Rev per Call Sub =
IF (
    ISFILTERED ( Reps[Rep ID] ),
    "Team $" & FORMAT ( [Rev per Call (All Reps)], "0.00" )
        & "  (" & FORMAT ( [Rev per Call vs Team], "+0.00;-0.00" ) & ")",
    "All 5 reps"
)

Rating Sub =
IF (
    ISFILTERED ( Reps[Rep ID] ),
    "Team " & FORMAT ( [Avg Rating (All Reps)], "0.00" ),
    "All 5 reps"
)

5 Star Sub =
IF (
    ISFILTERED ( Reps[Rep ID] ),
    "Team " & FORMAT ( [5 Star % (All Reps)], "0.0%" ),
    "All 5 reps"
)

Report Subtitle =
IF (
    ISFILTERED ( Reps[Rep ID] ),
    "2023  ·  " & CONCATENATEX ( VALUES ( Reps[Rep ID] ), Reps[Rep ID], ", ", Reps[Rep ID] )
        & "  ·  " & FORMAT ( [Call Count], "#,0" )
        & " of " & FORMAT ( [Calls (All Reps)], "#,0" ) & " calls",
    "2023  ·  1,000 calls  ·  15 customers  ·  5 representatives"
)
```

> The `$` in `"Team $"` is plain text joined with `&`, so it never localises. Only `$` *inside a format string* is at risk.

`ISFILTERED` answers the same question as Excel's `Calc!B2 < 5`: *has the user picked something?* `CONCATENATEX` is DAX's `TEXTJOIN`.

## 3.5 Time intelligence

```dax
Calls Prev Month = CALCULATE ( [Call Count], DATEADD ( 'Date'[Date], -1, MONTH ) )

Calls MoM % =
DIVIDE ( [Call Count] - [Calls Prev Month], [Calls Prev Month] )

Revenue YTD = TOTALYTD ( [Total Amount], 'Date'[Date] )

Revenue FYTD = TOTALYTD ( [Total Amount], 'Date'[Date], "6-30" )
```

Formats: MoM % Percentage 1 dp · both YTD Currency $ 0 dp.

`"6-30"` = the financial year **ends 30 June**. Same Jul–Jun year as the Excel FY column.

## 3.6 Ranking

```dax
Rev per Call Rank =
IF (
    HASONEVALUE ( Reps[Rep ID] ),
    RANKX ( ALL ( Reps ), [Revenue per Call] )
)
```

`HASONEVALUE` stops it returning a meaningless "1" on a total row.

✅ **Part 3 checkpoint.** Build a scratch **Table** visual with `Reps[Rep ID]` and the measures below, compare, then delete the visual:

| Rep | Call Count | Total Amount | Revenue per Call | Avg Rating | 5 Star % | Low Rating % | Share of Calls | Rev per Call vs Team | Rank |
|---|---|---|---|---|---|---|---|---|---|
| R01 | 189 | $18,415 | $97.43 | 3.92 | 31.7% | 5.3% | 18.9% | +0.81 | 3 |
| R02 | 218 | $20,581 | $94.41 | 3.87 | 33.9% | 10.1% | 21.8% | −2.21 | 4 |
| R03 | 207 | $20,872 | **$100.83** | 3.86 | 29.0% | 6.3% | 20.7% | **+4.21** | **1** |
| R04 | 186 | $16,651 | **$89.52** | 3.90 | 33.3% | 7.0% | 18.6% | **−7.10** | **5** |
| R05 | 200 | $20,104 | $100.52 | 3.88 | 25.5% | 5.0% | 20.0% | +3.90 | 2 |
| **Total** | **1,000** | **$96,623** | **$96.62** | **3.89** | **30.7%** | **6.8%** | **100%** | n/a | blank |

And a scratch table with `Date[Month]`:

| Month | Calls | Calls MoM % | Revenue YTD | Revenue FYTD |
|---|---|---|---|---|
| Jan | 79 | *(blank)* | $7,790 | $7,790 |
| Mar | 155 | **+134.8%** | $28,475 | $28,475 |
| Jun | 67 | −19.3% | $56,075 | **$56,075** |
| Jul | 64 | −4.5% | $62,082 | **$6,007** ← FY resets |
| Dec | 53 | −7.0% | $96,623 | **$40,548** |

If **FYTD at Jul equals YTD** ($62,082), the `"6-30"` didn't parse on your locale. See Part 9.

If **5 Star Calls = 248**, you used the Round button in Power Query. Go back to 1.2a.

---

# Part 4 — The Overview Page

## 4.1 Page setup

1. Rename *Page 1* → **`Overview`**
2. **View → Themes → Browse for themes** → pick `call-center-theme.json`
3. **Format page** (paint roller with nothing selected):
   - **Canvas settings → Type: Custom → 1280 × 1600** (a tall scrolling page, like the Excel dashboard)
   - **Canvas background** → `#F0F0ED`, transparency **0%**
4. **View → Page view → Fit to width**

## 4.2 The layout

Approximate pixel positions (Format → General → Properties → Size and position):

```text
y=16    Title (text box)                      Rep slicer (tiles)
y=56    Subtitle (card: Report Subtitle)
y=100   ┌Calls┐┌Revenue┐┌Rev/Call┐┌Talk time┐┌Avg rating┐┌5-star %┐   6 cards, h=110
y=226   Calls & revenue by month (line + column)            full width, h=260
y=502   Calls by rep   │ Revenue by rep   │ Calls by day of week          h=280
y=798   Customer mix by city │ Rating distribution │ Rep leaderboard      h=280
y=1094  Revenue by customer × rep (matrix + heatmap)                      h=480
```

Width 1240, left edge at x=20. Use **View → Gridlines** and **Snap to grid** to line things up. **Format → Align** does the rest.

## 4.3 Title and slicer

1. **Insert → Text box** → `Call Center Performance`, Segoe UI Semibold 24, `#2D3134`
2. **Slicer** visual → field **`Reps[Rep ID]`**
   - Format → Slicer settings → **Style: Tile**
   - Selection → **Multi-select with Ctrl** on, **Show "Select all"** off
   - Slicer header text: `REPRESENTATIVE`

## 4.4 The KPI cards

For each card: **Card** visual → the main measure in *Fields*.

| Card | Main value | Sub-line measure | Accent |
|---|---|---|---|
| CALLS | `Call Count` | `Calls Sub` | brick `#C04433` |
| REVENUE | `Total Amount` | `Revenue Sub` | orange `#E67E22` |
| REVENUE / CALL | `Revenue per Call` | `Rev per Call Sub` | teal `#2E7D6F` |
| TALK TIME (HRS) | `Talk Time (hrs)` | `Calls Sub` *(share of calls)* | plum `#7A4C6C` |
| AVG RATING | `Avg Rating` | `Rating Sub` | gold `#D69E2E` |
| 5-STAR % | `5 Star %` | `5 Star Sub` | teal `#2E7D6F` |

**The sub-line.** Depending on your Power BI version:

- **Card with *Reference labels*** (current versions): Format → **Reference labels → Add** → put the sub-line measure in it → turn its title off.
- **Older card without reference labels**: add a second, smaller Card holding the sub-line measure, category label off, font 9 pt grey, and stack it under the main card. **Group** the two (select both → right-click → Group).

Card styling: white background, 1 px `#E0E0DC` border, rounded corners 6 px, callout value 20 pt semibold, label 9 pt uppercase grey. Add the accent colour as the **callout value colour** or as a 4 px **accent bar** (Format → Visual → *Accent bar*, where available).

For the subtitle: a Card with `Report Subtitle`, category label **off**, 10 pt grey, transparent background.

✅ **Checkpoint:** cards read `1,000` · `$96,623` · `$96.62` · `1,498` · `3.89` · `30.7%` with `All 5 reps` underneath. Click **R02** on the slicer:

| Card | Main value | Sub-line |
|---|---|---|
| Calls | 218 | `21.8% of 1,000 team calls` |
| Revenue | $20,581 | `21.3% of team revenue` |
| Revenue / call | $94.41 | `Team $96.62  (-2.21)` |
| Talk time | 322 | `21.8% of 1,000 team calls` |
| Avg rating | 3.87 | `Team 3.89` |
| 5-star % | 33.9% | `Team 30.7%` |

**Note the difference from Excel.** In the Excel build the *big* number stayed at company level and the sub-line showed the rep. Here the big number is the **selection** and the sub-line is the **benchmark**. Both are valid. Power BI's is the more common convention, since the thing you clicked is the thing that's biggest.

## 4.5 Calls & revenue by month

1. **Line and clustered column chart**
2. X-axis: `Date[Month]` · Column y-axis: `Total Amount` · Line y-axis: `Call Count`
3. Columns `#E0B8B0` (soft brick), line `#C04433` 2.5 px, markers on (circle, 6)
4. Data labels on the **line only**
5. Turn off gridlines; turn off both value axes, since data labels carry the values
6. Title: `Calls and revenue by month` (left-aligned)
7. **X-axis → Type: Categorical**

✅ **Checkpoint:** months run **Jan → Dec** (not Apr, Aug, Dec…). If alphabetical, you missed *Sort by column* in 2.2. Call labels: 79, 66, **155**, 136, 83, 67, 64, **50**, 76, 114, 57, 53.

## 4.6 Calls by day of week

1. **Clustered bar chart** → Y-axis `Date[Day of Week]`, X-axis `Call Count`
2. Brick bars, data labels on, value axis off
3. Sort: **More options (…) → Sort axis → Day of Week → Sort ascending** (so Sunday is on top)

✅ **Checkpoint:** Sun 146 · Mon 133 · Tue 138 · Wed 153 · Thu 128 · Fri 141 · Sat 161.

## 4.7 Customer mix by city

1. **100% stacked bar chart** → Y-axis `Customers[City]`, X-axis `Call Count`, Legend `Customers[Gender]`
2. Female → brick, Male → teal (Format → Bars → Colors)
3. Data labels **inside centre**, white, bold
4. Legend: top-right

✅ **Checkpoint:** all three bars the **same length**. Cincinnati 144/132 · Cleveland 326/63 · Columbus 129/206.

This visual *should* respond to the rep slicer in Power BI (unlike Excel, where it stayed fixed). Both are defensible. If you want the Excel behaviour, select the slicer → **Format → Edit interactions** → click the **⊘ None** icon above this chart.

## 4.8 Satisfaction rating distribution ⚠️

1. **Clustered column chart** → X-axis **`Rating Scale[Rating Label]`** (from the dimension, *not* from `Calls`), Y-axis `Call Count`
2. Gold `#D69E2E`, data labels on, value axis off
3. ⭐ Right-click `Rating Label` in the X-axis well → **Show items with no data**

✅ **Checkpoint:** six columns: 1, 8, 59, 197, 428, 307.
Now click **R02**: you should **still see six slots**, with `0 ★` empty (R02 has no zero-star call): 0, 3, 19, 35, 87, 74 = 218.
If the `0 ★` slot vanishes, you used `Calls[Rating Rounded]` on the axis or skipped *Show items with no data*.

**This is the Excel homework problem, solved the Power BI way:** a dimension table guarantees the category exists, so the chart never changes shape.

## 4.9 The matrix: revenue by customer × rep

Build it now with ordinary fields; Part 5 upgrades it.

1. **Matrix** visual
2. Rows: `Customers[City]`, then `Customers[Customer ID]`
3. Columns: `Rep Axis[Rep ID]` *(the disconnected table; Part 5 explains)*
4. Values: leave empty for now
5. Format → **Layout: Tabular**, turn **+/-** icons off, **expand all** (the ⤓ fork icon in the visual header)

## 4.10 Rep leaderboard

A small ranked table. It's also the launch point for the drill-through in Part 6.

1. **Table** visual → `Reps[Rep ID]`, `Revenue per Call`, `Rev per Call vs Team`, `Rev per Call Rank`, `Call Count`
2. Sort by `Rev per Call Rank` ascending
3. `Rev per Call vs Team` → **Cell elements → Font color → fx → Rules**: `< 0` → brick `#C04433`, `>= 0` → teal `#2E7D6F`
4. Select the **slicer** → **Format → Edit interactions** → set this table to **⊘ None**, so all five reps always stay ranked
5. Title: `Leaderboard · revenue per call`

✅ **Checkpoint:** R03 $100.83 (1) · R05 $100.52 (2) · R01 $97.43 (3) · R02 $94.41 (4) · R04 $89.52 (5).

---

# Part 5 — The Highlight Trick

**The problem.** You want "Calls by rep" to show **all five bars** (context) while the slicer **highlights** the selected rep. But the slicer filters `Reps`, so any chart on `Reps[Rep ID]` shrinks to one bar.

**Excel's answer** was an unconnected pivot plus an `NA()` series.
**Power BI's answer** is a **disconnected axis table** plus two measures:

1. A **value** measure that ignores the slicer and uses the *axis* rep instead
2. A **colour** measure that asks "is the axis rep one of the selected reps?"

## 5.1 The measures

```dax
Axis Calls =
CALCULATE (
    [Call Count],
    REMOVEFILTERS ( Reps ),
    TREATAS ( VALUES ( 'Rep Axis'[Rep ID] ), Reps[Rep ID] )
)

Axis Revenue =
CALCULATE (
    [Total Amount],
    REMOVEFILTERS ( Reps ),
    TREATAS ( VALUES ( 'Rep Axis'[Rep ID] ), Reps[Rep ID] )
)

Axis Bar Colour =
VAR AxisRep = SELECTEDVALUE ( 'Rep Axis'[Rep ID] )
RETURN
    IF (
        ISFILTERED ( Reps[Rep ID] ) && AxisRep IN VALUES ( Reps[Rep ID] ),
        "#C04433",
        "#C6C6C4"
    )

Axis Revenue Colour =
VAR AxisRep = SELECTEDVALUE ( 'Rep Axis'[Rep ID] )
RETURN
    IF (
        ISFILTERED ( Reps[Rep ID] ) && AxisRep IN VALUES ( Reps[Rep ID] ),
        "#E67E22",
        "#C6C6C4"
    )

Axis Font Colour =
VAR AxisRep = SELECTEDVALUE ( 'Rep Axis'[Rep ID] )
RETURN
    IF (
        ISFILTERED ( Reps[Rep ID] ) && AxisRep IN VALUES ( Reps[Rep ID] ),
        "#C04433",
        "#2D3134"
    )
```

Format `Axis Revenue` as Currency $ 0 dp.

**Reading `Axis Calls` in plain English:**
*"Throw away whatever rep the slicer picked (`REMOVEFILTERS`), and instead filter to the rep on this bar's axis (`TREATAS` maps the disconnected table's value onto the real `Reps` column)."*

**Reading `Axis Bar Colour`:**
*"If the user has picked something, and this bar's rep is among the picked ones, brick; otherwise grey."*
That's Excel's `AND($B$2<5, COUNTIF(...)>0)`, line for line, and it handles **multi-select** too.

## 5.2 Calls by rep

1. **Clustered bar chart** → Y-axis `Rep Axis[Rep ID]`, X-axis `Axis Calls`
2. Format → **Bars → Colors → fx** →
   - Format style: **Field value**
   - What field should we base this on: **`Axis Bar Colour`**
3. Data labels on, value axis off, sort by `Rep ID` ascending
4. Title: `Calls by rep`

## 5.3 Revenue by rep

Same, with `Axis Revenue` and colour field **`Axis Revenue Colour`**. Title: `Revenue by rep`.

## 5.4 Finish the matrix

1. Values: **`Axis Revenue`**
2. Format → **Cell elements → Background color → fx** → Format style **Gradient** → Lowest value `#FFF9F6`, Highest value `#F0B0A0`
3. **Cell elements → Font color → fx** → **Field value** → `Axis Font Colour`
4. **Column headers** → bold
5. **Row subtotals** and **Column subtotals** on (Format → Subtotals)

✅ **Part 5 checkpoint:**

- No selection → both rep charts show 5 **grey** bars: calls 189 / 218 / 207 / 186 / 200; revenue $18,415 / $20,581 / $20,872 / $16,651 / $20,104
- Click **R02** → R02 bar turns brick (calls) and orange (revenue); **the other four bars stay, in grey, with their values**
- Ctrl-click **R01 + R03** → two bars highlighted
- Matrix: first row C0003 = $891 · $1,332 · $1,282 · $739 · $560 → $4,804. Grand total **$96,623**. Selected rep's column in brick text.

> **Why the matrix and rep charts use `Rep Axis` columns, not `Reps`:** the matrix must show all five rep columns so the selected one can stand out, exactly the reason `RepAmounts` was unconnected in Excel.

## 5.5 Stop the rep charts filtering the page

Clicking a bar in *Calls by rep* now selects a `Rep Axis` value, which filters nothing (it's disconnected) but still dims other visuals, which is confusing. Select the *Calls by rep* chart → **Format → Edit interactions** → set every other visual to **⊘ None**. Repeat for *Revenue by rep* and the matrix.

The **slicer** is the one control. Tell the user that with a text hint: a small text box under the slicer, `Pick a rep · Ctrl-click for several`.

---

# Part 6 — Rep Detail Drill-Through Page

**What it's for:** the manager's one-on-one. Right-click a rep anywhere → *Drill through → Rep Detail* → one page about that person.

## 6.1 Set up the page

1. New page → rename **`Rep Detail`**
2. In the Visualizations pane, **Drill through** section → drag **`Reps[Rep ID]`** into *Add drill-through fields here*
3. Set **Keep all filters → Off**, so the page shows the rep you drilled on, not whatever the Overview slicer happened to say
4. Power BI adds a **back arrow** button top-left automatically. Keep it.
4. Canvas 1280 × 720, background `#F0F0ED`

## 6.2 The visuals

| Visual | Fields | Notes |
|---|---|---|
| Card: title | `Rep Profile` measure (below) | 20 pt |
| 4 cards | `Call Count`, `Revenue per Call`, `Avg Rating`, `Low Rating %` | sub-lines as in 4.4 |
| Card: rank | `Rank Text` measure (below) | |
| Line chart | `Date[Month]` × `Call Count` | monthly trend for this rep |
| Column chart | `Rating Scale[Rating Label]` × `Call Count` | *Show items with no data* on |
| Bar chart | `Calls[Duration Bucket]` × `Call Count` | sort by `Bucket Order`, see below |
| Table | `Customers[Customer ID]`, `Customers[City]`, `Call Count`, `Total Amount`, `Revenue per Call` | sort by Total Amount desc |

Two extra measures:

```dax
Rep Profile = "Representative " & SELECTEDVALUE ( Reps[Rep ID], "(several)" )

Rank Text =
IF (
    HASONEVALUE ( Reps[Rep ID] ),
    "Rank " & [Rev per Call Rank] & " of 5 by revenue per call"
)
```

**Sorting the duration buckets:** in Data view, click `Calls[Duration Bucket]` → **Sort by column → `Bucket Order`**. (This is why you made that column in 1.2b.)

## 6.3 Try it

Go to **Overview**, clear the slicer, right-click the **R04** row in the **leaderboard** (4.10) → **Drill through → Rep Detail**.

> **Why the leaderboard and not the rep bar charts?** Drill-through only appears on visuals that use the drill-through field, `Reps[Rep ID]`. The bar charts use the disconnected `Rep Axis[Rep ID]` (Part 5), so they won't offer it. That's expected.

✅ **Checkpoint (R04):**

| Item | Value |
|---|---|
| Title | `Representative R04` |
| Calls | 186 |
| Revenue per call | $89.52 |
| Avg rating | 3.90 |
| Low rating % | 7.0% |
| Rank | `Rank 5 of 5 by revenue per call` |
| Monthly trend | 10, 11, 30, 23, 22, 11, 17, 12, 13, 15, 8, 14 |

Click the back arrow to return.

---

# Part 7 — Row-Level Security

**The business case:** the manager sees everyone; a rep opening the same report sees only their own calls. One file, not five.

## 7.1 Create a role

1. **Modeling → Manage roles → New**
2. Role name: `Rep R02`
3. Table: **`Reps`** → filter: `[Rep ID] = "R02"` (switch to the DAX editor if the UI shows dropdowns)
4. Save

The filter sits on the `Reps` **dimension** and flows through the relationship into `Calls`. You never secure the fact table directly.

## 7.2 Test it

**Modeling → View as → tick `Rep R02`** → OK.

✅ **Checkpoint:**

| Item | Value |
|---|---|
| Calls card | **218** |
| `Calls (All Reps)` in the sub-line | **218**, *not* 1,000 |
| Rep charts | only R02 has a bar |

**That second row is the important one.** `REMOVEFILTERS(Reps)` cannot remove a security filter, so under RLS the "team" benchmark shrinks to the rep's own numbers. **That is security working correctly:** a rep must not be able to reverse-engineer colleagues' figures.

If the business *wants* reps to see the team average, you'd publish a separate pre-aggregated team table outside the secured path. That's a design conversation, not a DAX trick. Good interview material.

**Stop viewing as the role** before continuing (the yellow banner → *Stop viewing*).

## 7.3 (Extension) Dynamic RLS

In production you'd have one role for everyone: add an `Email` column to `Reps` and set the filter to `[Email] = USERPRINCIPALNAME()`. Each rep then sees their own data automatically after the file is published to the Power BI Service. You can't fully test this without a Power BI work account, but you should be able to explain it.

---

# Part 8 — Test It

## Test 1 — Nothing selected

| Check | Expected |
|---|---|
| Cards | 1,000 · $96,623 · $96.62 · 1,498 · 3.89 · 30.7% |
| Sub-lines | `All 5 reps` |
| Subtitle | `2023 · 1,000 calls · 15 customers · 5 representatives` |
| Rep bars | 5 grey |
| Rating chart | 6 columns summing to 1,000 |
| Months | Jan → Dec in order |

## Test 2 — Click R02

| Check | Expected |
|---|---|
| Cards | 218 · $20,581 · $94.41 · 322 · 3.87 · 33.9% |
| Sub-lines | `21.8% of 1,000 team calls`, `Team 3.89`, etc. |
| Subtitle | `2023 · R02 · 218 of 1,000 calls` |
| Rep charts | R02 coloured, 4 grey bars still present |
| Rating chart | still 6 slots; 0 ★ empty |
| Monthly line | 19, 10, 40, 40, 13, 12, 9, 13, 18, 21, 14, 9 |

## Test 3 — Ctrl-click R01 + R03

| Check | Expected |
|---|---|
| Calls | 396 (`39.6% of 1,000 team calls`) |
| Revenue | $39,287 |
| Revenue / call | $99.21 |
| Rep charts / matrix | **two** highlighted |
| Subtitle | `2023 · R01, R03 · 396 of 1,000 calls` |

## Test 4 — Refresh

1. Open `data\calls.csv` in Notepad (**not** Excel, since Excel rewrites the date format on save), change the first row's Purchase Amount from `128` to `1128`, save
2. **Home → Refresh**
3. Revenue → **$97,623**
4. Change it back, save, Refresh → **$96,623**

If revenue doesn't move, the query isn't reading the file you edited. Check the `DataFolder` parameter.

## Test 5 — Reconcile

Every breakdown must sum to the headline. Hover the totals:

- Rep bars: 189 + 218 + 207 + 186 + 200 = **1,000**
- Rating columns: 1 + 8 + 59 + 197 + 428 + 307 = **1,000**
- Matrix grand total = **$96,623**
- Revenue YTD at Dec = **$96,623**

---

# Part 9 — Troubleshooting

### 5-star calls shows 248, not 307
You used *Add Column → Rounding → Round* (banker's rounding). Replace with `Number.Round([Satisfaction Rating], 0, RoundingMode.AwayFromZero)`. See 1.2a.

### Months are in alphabetical order
`Date[Month]` → **Sort by column → Month Number**. Also check the axis uses `Date[Month]`, not a month column on `Calls`.

### "Sort by column" gives an error about multiple values
Each `Month` must map to exactly one `Month Number`. It will, if the Date table is built as in 2.2. You'll hit this if you try sorting a column from the wrong table.

### Relationship line is dotted
It's inactive, usually because a second path exists between the same tables. Delete the extra relationship; this model needs exactly the four in 2.5.

### "Mark as date table" is greyed out or errors
The `Date` column must be type **Date** with no blanks and no duplicates. Rebuild with the `CALENDAR` expression in 2.2.

### FYTD equals YTD (financial year didn't reset in July)
The `"6-30"` year-end string didn't parse in your locale. Try `"30-6"`, or replace with a locale-proof version:

```dax
Revenue FYTD =
CALCULATE (
    [Total Amount],
    FILTER (
        ALL ( 'Date' ),
        'Date'[FY] = MAX ( 'Date'[FY] )
            && 'Date'[Date] <= MAX ( 'Date'[Date] )
    )
)
```

### Revenue shows ₹ instead of $
You typed a format string with `$`. Use **Measure tools → Format → Currency → $ English (United States)** instead.

### The 0 ★ column disappears when a rep is selected
The axis must come from `Rating Scale[Rating Label]`, and *Show items with no data* must be on. See 4.8.

### All five rep bars highlight at once
`ISFILTERED(Reps[Rep ID])` is missing from the colour measure, or the slicer is on `Rep Axis` instead of `Reps`.

### Rep bars shrink to one bar when a rep is selected
The chart's axis is `Reps[Rep ID]`. Switch it to `Rep Axis[Rep ID]` and the value to `Axis Calls`.

### Nothing highlights
Check the slicer field is **`Reps[Rep ID]`**, and that `Rep Axis` has **no** relationship (Model view: it must float alone).

### Colour fx option doesn't show "Field value"
The colour measure must return **text**. Check that it returns `"#C04433"` in quotes, not a number.

### Refresh fails: "Could not find file"
The `DataFolder` parameter is wrong or missing its trailing `\`. **Transform data → Edit parameters**.

### Dates load as text or errors
Your CSV was re-saved by Excel in a local date format. Restore the original `yyyy-mm-dd` file, or in Power Query use **Change Type → Using Locale → Date → English (United States)**.

---

# Part 10 — Portfolio & Publishing

A `.pbix` on your laptop proves nothing to a recruiter. Make it visible:

1. **Screenshots**: Overview (no selection), Overview (R02 selected), Rep Detail (R04), Model view. Use **File → Export → Export to PDF** for a clean multi-page PDF too.
2. **GitHub**: one repo with the `.pbix`, the `data` folder, the theme, the screenshots, and a README with the four headline findings (see the interview prep file, §8).
3. **Power BI Service (optional)**: *Publish* needs a work/school Microsoft account. Personal Gmail/Outlook accounts can't sign in. If you have one, publish to *My workspace* and set up **scheduled refresh** (needs a gateway for local files; you can explain it even if you don't set it up).
4. **Don't** use *Publish to web* with anything you wouldn't put on a billboard. It's fully public. Fine for this sample data; never for real company data. Say that if asked.

---

# Appendix — Every Number in This Project

Everything in the Excel guide's appendix still holds. These are the Power BI additions.

## Per representative

| Rep | Calls | Revenue | Rev/call | vs team | Rank | Avg rating | 5★ % | 0–2★ % | Share of calls | Share of revenue |
|---|---|---|---|---|---|---|---|---|---|---|
| R01 | 189 | $18,415 | $97.43 | +0.81 | 3 | 3.92 | 31.7% | 5.3% | 18.9% | 19.1% |
| R02 | 218 | $20,581 | $94.41 | −2.21 | 4 | 3.87 | 33.9% | **10.1%** | 21.8% | 21.3% |
| R03 | 207 | $20,872 | **$100.83** | +4.21 | **1** | 3.86 | 29.0% | 6.3% | 20.7% | 21.6% |
| R04 | 186 | $16,651 | **$89.52** | −7.10 | **5** | 3.90 | 33.3% | 7.0% | 18.6% | 17.2% |
| R05 | 200 | $20,104 | $100.52 | +3.90 | 2 | 3.88 | 25.5% | 5.0% | 20.0% | 20.8% |
| **Team** | **1,000** | **$96,623** | **$96.62** | n/a | n/a | **3.89** | **30.7%** | **6.8%** | 100% | 100% |

## By month

| Month | Calls | MoM % | Revenue | Revenue YTD | Revenue FYTD |
|---|---|---|---|---|---|
| Jan | 79 | n/a | $7,790 | $7,790 | $7,790 |
| Feb | 66 | −16.5% | $6,122 | $13,912 | $13,912 |
| Mar | 155 | +134.8% | $14,563 | $28,475 | $28,475 |
| Apr | 136 | −12.3% | $13,536 | $42,011 | $42,011 |
| May | 83 | −39.0% | $8,171 | $50,182 | $50,182 |
| Jun | 67 | −19.3% | $5,893 | $56,075 | $56,075 |
| Jul | 64 | −4.5% | $6,007 | $62,082 | $6,007 |
| Aug | 50 | −21.9% | $4,950 | $67,032 | $10,957 |
| Sep | 76 | +52.0% | $6,689 | $73,721 | $17,646 |
| Oct | 114 | +50.0% | $10,875 | $84,596 | $28,521 |
| Nov | 57 | −50.0% | $6,775 | $91,371 | $35,296 |
| Dec | 53 | −7.0% | $5,252 | $96,623 | $40,548 |

## By quarter

| Quarter | Calls | Revenue |
|---|---|---|
| Q1 | 300 | $28,475 |
| Q2 | 286 | $27,600 |
| Q3 | 190 | $17,646 |
| Q4 | 224 | $22,902 |

## By city

| City | Calls | Revenue | Avg rating |
|---|---|---|---|
| Cincinnati | 276 | $26,061 | 4.03 |
| Cleveland | 389 | $37,849 | 3.88 |
| Columbus | 335 | $32,713 | 3.77 |

## By gender

| Gender | Calls | Revenue |
|---|---|---|
| Female | 599 | $57,788 |
| Male | 401 | $38,835 |

## By customer (sorted by revenue)

| Customer | Calls | Revenue |
|---|---|---|
| C0005 | 79 | $7,747 |
| C0004 | 82 | $7,689 |
| C0013 | 69 | $7,230 |
| C0007 | 67 | $7,216 |
| C0012 | 70 | $6,819 |
| C0001 | 65 | $6,785 |
| C0011 | 74 | $6,749 |
| C0009 | 69 | $6,601 |
| C0015 | 59 | $6,518 |
| C0010 | 72 | $6,242 |
| C0006 | 62 | $6,176 |
| C0002 | 63 | $5,634 |
| C0014 | 60 | $5,404 |
| C0008 | 59 | $5,009 |
| C0003 | 50 | $4,804 |

## Rating distribution for R02 (the "missing 0" test)

| 0 ★ | 1 ★ | 2 ★ | 3 ★ | 4 ★ | 5 ★ | Total |
|---|---|---|---|---|---|---|
| *(none)* | 3 | 19 | 35 | 87 | 74 | 218 |

## Model facts

| Item | Value |
|---|---|
| Date table rows | 365 |
| Distinct call dates | 330 (35 days with no calls) |
| Ratings ending in .5 | 100 (why banker's rounding matters) |
| The zero-star call | Call_7340 · C0006 · R03 · 27-Jan-2023 · $54 |
| Customer age range | 22–43, mean 32.6 |

---

## The five ideas, Power BI edition

1. **Clean once, replay forever.** Power Query steps are a recipe, not an edit.
2. **Star schema.** One fact table, dimensions around it, single-direction filters, foreign keys hidden.
3. **Benchmarks live in the measure.** `REMOVEFILTERS` makes "what shouldn't react" explicit, in place of Excel's unticked Report Connections.
4. **Dimensions guarantee categories.** A Date table and a Rating table mean no bar ever disappears.
5. **Disconnected table + `TREATAS` = highlighting.** The same idea as Excel's `NA()` series: show everything, colour the selection.
