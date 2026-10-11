# Call Center Performance Dashboard — Build Guide

**A step-by-step manual to build the whole project from a blank workbook.**

This is not a concepts file. Every step is: *do this → then this → check this number*.
If your number matches the ✅ checkpoint, move on. If it doesn't, stop and fix it there — errors compound.

> Companion file: `Excel Complete Notes.md` explains *why* each function works.
> This file only tells you *what to do*.
> Finished example: `Call Center Dashboard.xlsx` in this folder.
> Power BI version: `Power BI/Power BI Dashboard - Build Guide.md` rebuilds this same project in Power BI Desktop.

---

## Contents

| Part | What you build | Time |
|---|---|---|
| [0](#part-0--the-brief) | The brief + setup | 5 min |
| [1](#part-1--the-data-sheet) | Data sheet, 2 tables, 5 formula columns | 20 min |
| [2](#part-2--the-data-model) | Data Model + relationship | 10 min |
| [3](#part-3--the-five-dax-measures) | 5 DAX measures | 15 min |
| [4](#part-4--the-nine-pivot-tables) | 9 pivot tables | 30 min |
| [5](#part-5--the-slicer-and-its-wiring) | Slicer + wiring | 10 min |
| [6](#part-6--the-calc-sheet) | Calc sheet (chart feeds) | 30 min |
| [7](#part-7--dashboard-canvas--kpi-cards) | Canvas + 5 KPI cards | 30 min |
| [8](#part-8--the-six-charts) | 6 charts | 45 min |
| [9](#part-9--the-customer--rep-matrix) | Matrix + conditional formatting | 25 min |
| [10](#part-10--polish) | Polish | 15 min |
| [11](#part-11--test-it) | Acceptance tests | 10 min |
| [12](#part-12--troubleshooting) | When it breaks | — |
| [13](#part-13--extend-it) | Extensions | — |
| [A](#appendix--every-number-in-this-project) | Every correct number | — |

---

# Part 0 — The Brief

## What you are building

An interactive one-page dashboard. The user clicks a representative and **everything** updates: the KPI sub-lines, the trend, the weekday pattern, the rating spread, the highlighted bar, and the highlighted column in the revenue table.

## The finished layout

```text
┌──────────────────────────────────────────────────────────────────────┐
│  Call Center Performance                  ┌── SLICER: R01..R05 ──┐   │
│  2023 · R02 · 218 of 1,000 calls          └──────────────────────┘   │
├──────────┬──────────┬──────────┬──────────┬──────────────────────────┤
│ CALLS    │ REVENUE  │ TALK TIME│ AVG RATING│ 5-STAR CALLS            │  ← 5 KPI cards
│ 1,000    │ $96,623  │ 1,498 hrs│ 3.89      │ 307                     │
│ R02·218  │ R02·$20K │ R02·322h │ R02·3.87  │ R02·74                  │
├──────────────────────────────────────────────────────────────────────┤
│  Calls by month          (area + line, full width)                   │
├────────────────────┬────────────────────┬────────────────────────────┤
│ Calls by rep       │ Revenue by rep     │ Calls by day of week       │
│ (grey + highlight) │ (grey + highlight) │                            │
├────────────────────────────────┬─────────────────────────────────────┤
│ Customer mix by city (100%)    │ Satisfaction rating distribution    │
├──────────────────────────────────────────────────────────────────────┤
│  Revenue by customer and representative  (heatmap + column highlight)│
└──────────────────────────────────────────────────────────────────────┘
```

## Sheets you will create, in this order

| # | Sheet | Purpose | Pretty? |
|---|---|---|---|
| 1 | `Data` | The two source tables | No |
| 2 | `Pivots` | All 9 pivot tables | No |
| 3 | `Calc` | Helper formulas that feed charts | No |
| 4 | `Dashboard` | The only sheet anyone looks at | **Yes** |
| 5 | `Notes` | How it works | No |

**The one design rule for this whole project:**

> Never calculate on the Dashboard sheet. Never decorate on the Pivots sheet.

## Before you start

- Excel 2016 or later (365 ideal). You need **Data Model** support.
- Get the source file: `07.sample-data-excel-portfolio-project.xlsx`.
- If it opens in **Protected View**, click **Enable Editing**, then **Save As** to a real folder. You cannot build a data model from a Protected View file.

## The numbers you are aiming for

Write these on a sticky note. Every checkpoint refers back to them.

| Metric | Value |
|---|---|
| Calls | **1,000** |
| Revenue | **$96,623** |
| Talk time | **89,850 min = 1,498 hrs** |
| Average rating | **3.8854 → 3.89** |
| 5-star calls | **307** |
| Customers | **15** |
| Representatives | **5** |
| Date range | 1 Jan 2023 → 31 Dec 2023 |

---

# Part 1 — The Data Sheet

## 1.1 Set up the workbook

1. New blank workbook → **Save As** `Call Center Dashboard.xlsx` in your project folder.
2. Rename `Sheet1` to **`Data`**.

## 1.2 Bring in the calls data

1. Open the source file.
2. Copy the calls block (`Call number` header through the last row) — 1,000 rows + header.
3. Paste into your `Data` sheet at **`A1`**.
4. Copy the customers block and paste at **`N1`** (leave column M empty as a gap).

## 1.3 Rename the headers

The source headers are inconsistent and contain a typo. Fix them now — every later step depends on these exact names.

| Column | Source header | **Rename to** |
|---|---|---|
| A | Call number | `Call ID` |
| B | Customer ID | `Customer ID` |
| C | Duration | `Call Duration` |
| D | Representative | `Rep ID` |
| E | Date of Call | `Call Date` |
| F | Purchase Amount | `Purchase Amount` |
| G | Satisfaction Rating | `Satisfaction Rating` |
| H | FY | `FY` |
| I | Day of week | `Day of Week` |
| J | Duration Bucket | `Duration Bucket` |
| K | Rating rounded | `Rating Rounded` |
| L | *(new)* | `Month` |

Customers block:

| Column | Source header | **Rename to** |
|---|---|---|
| N | Customer ID | `Customer ID` |
| O | **Gedner** ← typo in source | `Gender` |
| P | Age | `Age` |
| Q | City | `City` |

## 1.4 Make them Excel Tables

1. Click any cell in the calls data → **`Ctrl + T`** → tick *My table has headers* → OK.
2. **Table Design → Table Name** → type **`Calls`**
3. Click any cell in the customer data → **`Ctrl + T`** → OK.
4. Name it **`Customers`**

> Naming matters. Every DAX measure and every pivot field refers to `Calls` and `Customers` by name. `Table1` will make the rest of this guide unreadable.

✅ **Checkpoint:** click inside `Calls`, look at the Name Box. It says `Calls`. Row count = 1,000 data rows (rows 2–1001).

## 1.5 Rebuild the five derived columns as formulas

Delete the pasted values in H, I, J, K and retype them as formulas. You want them live, so new data classifies itself.

Click cell **H2** and enter each of these. In a Table, the formula auto-fills all 1,000 rows.

### H — FY (financial year, July → June)

```excel
=YEAR([@[Call Date]])+IF(MONTH([@[Call Date]])>=7,1,0)
```

This company's financial year **ends in June**. So July 2023 belongs to FY2024.

✅ **Checkpoint:** FY 2023 = **586** rows, FY 2024 = **414** rows.

### I — Day of Week

```excel
=TEXT([@[Call Date]],"dddd")
```

`"dddd"` = full name (Monday). `"ddd"` = short (Mon).

### J — Duration Bucket

**Five** tiers, not three. Get these boundaries right or your bucket counts will be wrong:

```excel
=IFS([@[Call Duration]]<=10,"Under 10 mins",
     [@[Call Duration]]<=30,"10 to 30 mins",
     [@[Call Duration]]<=60,"30 to 60 mins",
     [@[Call Duration]]<=120,"1 to 2 hours",
     TRUE,"More than 2 hours")
```

✅ **Checkpoint:** Under 10 = **9** · 10–30 = **51** · 30–60 = **179** · 1–2 hrs = **524** · 2+ hrs = **237**. Total 1,000.

### K — Rating Rounded

```excel
=ROUND([@[Satisfaction Rating]],0)
```

✅ **Checkpoint — read this one carefully:**

| Rating | Calls |
|---|---|
| **0** | **1** ← yes, one call rates 0 |
| 1 | 8 |
| 2 | 59 |
| 3 | 197 |
| 4 | 428 |
| 5 | 307 |

That single 0-star call is a trap. Later, if you build the rating chart on a 1–5 axis, your chart will silently show 999 calls. **Your rating axis must be 0–5.**

### L — Month

```excel
=TEXT([@[Call Date]],"mmm")
```

✅ **Checkpoint:** Jan 79 · Feb 66 · Mar 155 · Apr 136 · May 83 · Jun 67 · Jul 64 · Aug 50 · Sep 76 · Oct 114 · Nov 57 · Dec 53 = 1,000.

## 1.6 Format and verify

1. Column E → `Ctrl+1` → Date → `14-Mar-12` style
2. Column F → `#,##0`
3. Column G → `0.0`
4. **View → Freeze Panes → Freeze Top Row**

✅ **Part 1 final check** — put these in blank cells, confirm, then delete them:

```excel
=COUNTA(Calls[Call ID])              →  1000
=SUM(Calls[Purchase Amount])         →  96623
=SUM(Calls[Call Duration])           →  89850
=AVERAGE(Calls[Satisfaction Rating]) →  3.8854
=COUNTIFS(Calls[Rating Rounded],5)   →  307
=COUNTA(Customers[Customer ID])      →  15
```

**Do not continue until all six match.**

---

# Part 2 — The Data Model

**Why:** `City` and `Gender` live in `Customers`; the calls live in `Calls`. To chart "calls by city" one pivot must read both tables. A relationship makes that possible.

## 2.1 Load the first table to the model

1. Click any cell in `Calls`
2. **Insert → PivotTable**
3. Source shows `Calls`
4. Choose **New Worksheet**
5. ⭐ **Tick "Add this data to the Data Model"** ← the whole part depends on this box
6. OK

Rename the new sheet to **`Pivots`**.

## 2.2 Load the second table

1. Go to `Data` sheet, click any cell in `Customers`
2. **Insert → PivotTable** → **Existing Worksheet** → pick a far-away cell like `Pivots!Z100`
3. ⭐ Tick **Add this data to the Data Model** → OK
4. You can delete this throwaway pivot afterwards — the table stays in the model.

## 2.3 Create the relationship

1. **Data → Relationships → New**
   *(or PivotTable Analyze → Relationships)*
2. Fill it in **exactly** like this:

| Box | Value |
|---|---|
| Table | **`Calls`** ← the many side |
| Column (Foreign) | `Customer ID` |
| Related Table | **`Customers`** ← the one side |
| Related Column (Primary) | `Customer ID` |

3. OK → Close

**Get the direction right.** `Customers` is the "one" side because each customer appears once. `Calls` is the "many" side. Reversed, Excel refuses to save it.

✅ **Checkpoint:** In the PivotTable Fields pane you now see an **All / Active** toggle, and under **All** both `Calls` and `Customers` are listed.

---

# Part 3 — The Five DAX Measures

**Why not just drag fields into Values?** A measure is written once, respects every slicer automatically, carries its own number format, and can be reused in all 9 pivots. Dragging fields means reformatting the same thing nine times.

## 3.1 How to add one

Click a pivot → in the Fields pane **right-click the `Calls` table → Add Measure…**

*(If you don't see it: File → Options → Add-ins → Manage: COM Add-ins → Go → tick **Microsoft Power Pivot for Excel**. Then use Power Pivot → Measures → New Measure.)*

## 3.2 Add these five

Fill in the dialog: **Measure Name**, **Formula**, then click **Number Format** before OK.

### 1. Call Count

```dax
=COUNTROWS(Calls)
```
Format: **Number**, 0 decimals, ✅ use thousands separator

### 2. Total Amount

```dax
=SUM(Calls[Purchase Amount])
```
Format: **Currency**, 0 decimals

### 3. Total Duration

```dax
=SUM(Calls[Call Duration])
```
Format: **Number**, 0 decimals, thousands separator

### 4. Avg Rating

```dax
=AVERAGE(Calls[Satisfaction Rating])
```
Format: **Decimal Number**, 2 decimals

### 5. 5 Star Calls

```dax
=CALCULATE([Call Count], Calls[Rating Rounded] = 5)
```
Format: **Number**, 0 decimals

**What `CALCULATE` does, in one line:** *"Work out Call Count, but add the filter Rating Rounded = 5."*

It still obeys the pivot row and the slicer, then layers its own filter on top. That's why selecting R02 gives R02's 74 five-star calls, not everyone's 307.

✅ **Checkpoint:** Drop all five into a scratch pivot's Values with nothing in Rows:

| Measure | Value |
|---|---|
| Call Count | 1,000 |
| Total Amount | $96,623 |
| Total Duration | 89,850 |
| Avg Rating | 3.89 |
| 5 Star Calls | 307 |

---

# Part 4 — The Nine Pivot Tables

All nine live on the `Pivots` sheet. Build them at the exact anchor cells below — Part 6's formulas point at these addresses.

## 4.1 The build recipe

For each row of the table:

1. **Insert → PivotTable → From Data Model**
   *(older Excel: Insert → PivotTable → tick "Use this workbook's Data Model")*
2. **Existing Worksheet** → type the anchor cell → OK
3. Drag the listed fields into Rows / Columns / Values
4. ⭐ **PivotTable Analyze → PivotTable Name** → type the name from the table

**Naming them is not optional.** Part 6 has 60+ formulas that reference these by position; when something breaks, names are how you find it.

## 4.2 The nine

| # | Name | Anchor | Rows | Columns | Values | Slicer? |
|---|---|---|---|---|---|---|
| 1 | `SummaryPivot` | `A3` | — | — | all 5 measures | ❌ **No** |
| 2 | `RepPivot` | `A8` | — | — | all 5 measures | ✅ Yes |
| 3 | `MonthTrend` | `A13` | `Calls[Month]` | — | Call Count | ✅ Yes |
| 4 | `WeekdayTrend` | `E13` | `Calls[Day of Week]` | — | Call Count | ✅ Yes |
| 5 | `RepsPivot` | `I13` | `Calls[Rep ID]` | — | Call Count, Total Amount | ❌ **No** |
| 6 | `SelectedRep` | `N13` | `Calls[Rep ID]` | — | *(none)* | ✅ Yes |
| 7 | `GenderCity` | `A32` | `Customers[City]` | `Customers[Gender]` | Call Count | ❌ No |
| 8 | `RepSatisfaction` | `G32` | `Calls[Rating Rounded]` | — | Call Count | ✅ Yes |
| 9 | `RepAmounts` | `A42` | `Customers[City]`, `Customers[Customer ID]` | `Calls[Rep ID]` | Total Amount | ❌ **No** |

## 4.3 Why four pivots are deliberately NOT connected

This is the single most important design decision in the project. Beginners connect everything and the dashboard becomes useless.

| Pivot | Not connected because… |
|---|---|
| `SummaryPivot` | The big KPI number must always show **1,000**, so you can compare the selection against the whole |
| `RepsPivot` | You need **all five bars** visible to highlight one of them |
| `GenderCity` | Customer demographics are a company fact, not a rep fact |
| `RepAmounts` | You need **all five columns** so the selected one can be highlighted |

> Rule of thumb: if the visual's job is to **show context**, don't connect it. If its job is to **respond**, connect it.

## 4.4 Notes on specific pivots

**`SelectedRep` (#6) has no Values on purpose.** It's a list of the currently-selected reps and nothing else. Part 6 reads it to work out what the user clicked. Turn off its Grand Total: **Design → Grand Totals → Off for Rows and Columns**.

**`RepAmounts` (#9)** — set **Design → Report Layout → Show in Tabular Form** so City and Customer ID sit in separate columns. Set **Design → Grand Totals → On for Rows Only**.

✅ **Checkpoint:**

- `A4` (SummaryPivot Call Count) = **1000**
- `RepsPivot` reads R01 189 · R02 218 · R03 207 · R04 186 · R05 200
- `RepSatisfaction` shows **6 rows: 0,1,2,3,4,5** — if you only see 5, revisit step 1.5
- `RepAmounts` grand total = **96,623**

---

# Part 5 — The Slicer and Its Wiring

## 5.1 Create it

1. Click **`RepPivot`** (the one at `A8`)
2. **PivotTable Analyze → Insert Slicer**
3. Tick **`Rep ID`** → OK

## 5.2 Wire it to exactly five pivots

Right-click the slicer → **Report Connections** → tick **only** these:

- ✅ `RepPivot`
- ✅ `MonthTrend`
- ✅ `WeekdayTrend`
- ✅ `SelectedRep`
- ✅ `RepSatisfaction`
- ❌ `SummaryPivot`
- ❌ `RepsPivot`
- ❌ `GenderCity`
- ❌ `RepAmounts`

## 5.3 Shape it

With the slicer selected, **Slicer** tab:

| Setting | Value |
|---|---|
| Columns | **5** (makes it a horizontal button strip) |
| Height | ~1.5 cm |
| Width | ~11 cm |
| Caption | `REPRESENTATIVE` |
| Style | pick one; right-click → **Duplicate** to recolour it later |

Leave the slicer on the Pivots sheet for now. You'll cut and paste it onto the Dashboard in Part 7.

✅ **Checkpoint:** Click **R02** on the slicer.

| Pivot | Should show |
|---|---|
| `RepPivot` (A9) | **218** |
| `SummaryPivot` (A4) | **1000** ← must NOT change |
| `RepsPivot` | all five reps still listed |

If `SummaryPivot` changed to 218, un-tick it in Report Connections.

Now click the slicer's **Clear Filter** ✗ before moving on.

---

# Part 6 — The Calc Sheet

**Why this sheet exists.** If you chart a pivot directly and the current rep has no 0-star calls, that bar **disappears** and the chart changes shape. It looks broken. (This is exactly the problem the tutorial leaves as homework.)

**The fix:** hard-code every category on a helper sheet and pull the numbers in with `GETPIVOTDATA`, wrapped in `IFERROR` so a missing category returns **0** instead of vanishing. The chart then never moves.

Insert a new sheet, name it **`Calc`**.

## 6.1 The trick for writing these formulas

Do **not** type the cube syntax by hand. Let Excel write it:

1. In a Calc cell, type `=`
2. Go to the Pivots sheet and **click the value you want**
3. Excel writes a full `GETPIVOTDATA(...)` for you
4. Press Enter, then edit it: replace the hard-coded item (e.g. `"...&[Jan]"`) with a cell reference (`"...&["&$A6&"]"`)
5. Wrap the whole thing in `IFERROR(...,0)`

*(If Excel writes a plain reference like `=Pivots!B14` instead, turn GETPIVOTDATA on: PivotTable Analyze → Options ▾ → tick **Generate GetPivotData**.)*

> Note: because these pivots come from the Data Model, the item syntax is cube-style — `"[Calls].[Month].&[Jan]"` — not the plain `"Month","Jan"` you'd get from a normal pivot. Let Excel generate it.

## 6.2 Selection state (the brain of the dashboard)

| Cell | Content |
|---|---|
| `A1` | `Selected reps` |
| `A2` | `Rep count` |
| `A3` | `Caption` |
| `A4` | `Subtitle` |

Scrub the pivot's Grand Total row out of the rep list — put in **`Z1:Z5`**:

```excel
=IF(LEFT(Pivots!N14,1)="R",Pivots!N14,"")
```
…then `Z2` → `N15`, `Z3` → `N16`, `Z4` → `N17`, `Z5` → `N18`.

Then:

**`B2` — how many reps are selected**
```excel
=COUNTIF(Pivots!$N$14:$N$19,"R0*")
```

**`B1` — their names**
```excel
=IF($B$2>=5,"All representatives",TEXTJOIN(", ",TRUE,$Z$1:$Z$5))
```

**`B3` — short caption**
```excel
=IF($B$2>=5,"Showing all 1,000 calls","Filtered to "&$B$2&" of 5 reps")
```

**`B4` — dashboard subtitle**
```excel
=IF($B$2>=5,
   "2023  ·  1,000 calls  ·  15 customers  ·  5 representatives",
   "2023  ·  "&$B$1&"  ·  "&TEXT(Pivots!$A$9,"#,##0")&" of 1,000 calls")
```

**How the selection detection works:** a pivot only lists the items that pass its filter. Nothing selected → all 5 reps appear → `B2` = 5. Select R02 → only R02 appears → `B2` = 1. So `B2 < 5` means "the user has filtered something", which is what drives every highlight in the dashboard.

✅ **Checkpoint:** `B1` = `All representatives`, `B2` = `5`. Click R02 on the slicer → `B1` = `R02`, `B2` = `1`. Clear the slicer.

## 6.3 KPI feeds

| Cell | Content |
|---|---|
| `V1` `W1` `X1` | `KPI` · `Overall` · `Selected` |

| Row | `V` (label) | `W` (overall) | `X` (selected) |
|---|---|---|---|
| 2 | `Calls` | `=Pivots!$A$4` | see below |
| 3 | `Amount` | `=Pivots!$B$4` | |
| 4 | `Talk time` | `=Pivots!$C$4/60` | |
| 5 | `Avg rating` | `=Pivots!$D$4` | |
| 6 | `5-star calls` | `=Pivots!$E$4` | |

Number-format column W: `#,##0` · `$#,##0` · `#,##0" hrs"` · `0.00` · `#,##0`

Column X (row 9 of the Pivots sheet = the slicer-driven `RepPivot`):

```excel
X2  =IF($B$2>=5,"all 5 reps",$B$1&"  ·  "&TEXT(Pivots!$A$9,"#,##0"))
X3  =IF($B$2>=5,"all 5 reps",$B$1&"  ·  "&TEXT(Pivots!$B$9,"$#,##0"))
X4  =IF($B$2>=5,"all 5 reps",$B$1&"  ·  "&TEXT(Pivots!$C$9/60,"#,##0")&" hrs")
X5  =IF($B$2>=5,"all 5 reps",$B$1&"  ·  "&TEXT(Pivots!$D$9,"0.00"))
X6  =IF($B$2>=5,"all 5 reps",$B$1&"  ·  "&TEXT(Pivots!$E$9,"#,##0"))
```

`TEXT()` is needed because joining a number to text with `&` strips its formatting. Format column X as **General** — it holds text now.

✅ **Checkpoint:** `W2:W6` = `1,000` · `$96,623` · `1,498 hrs` · `3.89` · `307`

## 6.4 Monthly trend — `A5:C17`

| Cell | Content |
|---|---|
| `A5` `B5` `C5` | `Month` · `Calls` · `Calls (area)` |
| `A6:A17` | Type `Jan` … `Dec` **manually**, in calendar order |

`B6`, fill down to `B17`:

```excel
=IFERROR(GETPIVOTDATA("[Measures].[Call Count]",Pivots!$A$13,
   "[Calls].[Month]","[Calls].[Month].&["&$A6&"]"),0)
```

`C6` (a duplicate — it becomes the shaded area in Part 8):

```excel
=$B6
```

> Typing the months by hand is the whole point. The pivot sorts them alphabetically (Apr, Aug, Dec…). Your hard-coded list fixes the order **and** guarantees 12 bars always exist.

✅ **Checkpoint:** `B6:B17` = 79, 66, 155, 136, 83, 67, 64, 50, 76, 114, 57, 53

## 6.5 Weekday — `E5:F12`

`E5` `F5` = `Day` · `Calls`
`E6:E12` = `Sunday`, `Monday`, `Tuesday`, `Wednesday`, `Thursday`, `Friday`, `Saturday`

```excel
F6  =IFERROR(GETPIVOTDATA("[Measures].[Call Count]",Pivots!$E$13,
       "[Calls].[Day of Week]","[Calls].[Day of Week].&["&$E6&"]"),0)
```

✅ **Checkpoint:** 146, 133, 138, 153, 128, 141, 161

## 6.6 Rep comparison + the highlight columns — `H5:L10`

Headers `H5:L5`: `Rep` · `Calls` · `Amount` · `Sel calls` · `Sel amount`
`H6:H10` = `R01` … `R05`

```excel
I6  =IFERROR(GETPIVOTDATA("[Measures].[Call Count]",Pivots!$I$13,
       "[Calls].[Rep ID]","[Calls].[Rep ID].&["&$H6&"]"),0)

J6  =IFERROR(GETPIVOTDATA("[Measures].[Total Amount]",Pivots!$I$13,
       "[Calls].[Rep ID]","[Calls].[Rep ID].&["&$H6&"]"),0)

K6  =IF(AND($B$2<5,COUNTIF(Pivots!$N$14:$N$19,$H6)>0),$I6,NA())

L6  =IF(AND($B$2<5,COUNTIF(Pivots!$N$14:$N$19,$H6)>0),$J6,NA())
```

Fill down to row 10.

**How columns K and L work — this is the highlight trick:**

- `$B$2<5` → only highlight when the user has actually filtered something
- `COUNTIF(...)>0` → is *this* rep among the selected ones?
- If yes → return the value. If no → return **`NA()`**

`NA()` is the key. Excel charts **draw nothing** for `#N/A`. In Part 8 you plot K and L as a second series at 100% overlap, so only the selected rep's bar gets a colour on top.

> Why `AND($B$2<5, ...)`? Without it, "nothing selected" shows all 5 reps in `SelectedRep`, so all 5 would highlight — meaningless. Bonus: this version also handles **multi-select** correctly, which the video's version does not.

✅ **Checkpoint:** `I6:I10` = 189, 218, 207, 186, 200 · `J6:J10` = 18415, 20581, 20872, 16651, 20104 · `K6:K10` all show `#N/A` (correct — nothing is selected). Click R02: `K7` = 218, the rest `#N/A`.

## 6.7 Gender by city — `N5:P8`

`N5` `O5` `P5` = `City` · `Female` · `Male`
`N6:N8` = `Cincinnati`, `Cleveland`, `Columbus`

```excel
O6  =IFERROR(GETPIVOTDATA("[Measures].[Call Count]",Pivots!$A$32,
       "[Customers].[City]","[Customers].[City].&["&$N6&"]",
       "[Customers].[Gender]","[Customers].[Gender].&[Female]"),0)
```

`P6` — same but `.&[Male]`.

✅ **Checkpoint:**

| City | Female | Male |
|---|---|---|
| Cincinnati | 144 | 132 |
| Cleveland | 326 | 63 |
| Columbus | 129 | 206 |

## 6.8 Rating distribution — `R5:S11` ⚠️

`R5` `S5` = `Rating` · `Calls`
`R6:R11` = `0 star`, `1 star`, `2 star`, `3 star`, `4 star`, `5 star`

**Six rows — 0 through 5.** This is the trap from Part 1.5.

```excel
S6  =IFERROR(GETPIVOTDATA("[Measures].[Call Count]",Pivots!$G$32,
       "[Calls].[Rating Rounded]","[Calls].[Rating Rounded].&[0]"),0)
```

…and rows 7–11 with `&[1]` … `&[5]`.

✅ **Checkpoint:** 1, 8, 59, 197, 428, 307 — **sums to 1,000**. If yours sums to 999 you built a 1–5 axis.

---

# Part 7 — Dashboard Canvas + KPI Cards

Insert a sheet, name it **`Dashboard`**, and drag it to the front.

## 7.1 Prepare the canvas

1. `Ctrl+A` → **Home → Fill Colour** → a light warm grey (`#F0F0ED`)
2. **View** → untick **Gridlines**, untick **Headings**
3. Set column widths for the Part 9 table now, before placing anything:

| Column | Width | Holds |
|---|---|---|
| A | 2 | left margin |
| B | 15 | City |
| C | 13 | Customer |
| D–H | 16 each | R01 … R05 |
| I | 15 | Total |
| J | 5 | right margin |

4. Set row heights: rows 62–76 to **22** (these hold the 15 customers later).

## 7.2 Title block

1. **Insert → Text Box** near `A2`, type `Call Center Performance` — 24 pt, semibold, dark grey
2. Second text box beneath it. Select it, click the **formula bar**, type:

```excel
=Calc!$B$4
```

That links the subtitle to the live selection. A shape linked to a cell shows that cell's **formatted** value and updates on every click.

> A shape can only ever display **one cell**. That's why each KPI card below is built from several shapes.

## 7.3 Move the slicer across

Go to `Pivots`, click the slicer, `Ctrl+X`, go to `Dashboard`, `Ctrl+V`. Position it top-right.
(A warning about moving a slicer is normal — accept it.)

**Hold `Alt` while dragging** to snap shapes to cell edges. This is how you get things aligned without fiddling.

## 7.4 Build one KPI card

Build the first one properly, then copy it four times.

**Four shapes per card:**

| # | Shape | Content | Format |
|---|---|---|---|
| 1 | Rounded Rectangle | *(empty)* | White fill, thin grey outline, no shadow |
| 2 | Small rectangle | *(empty)* | Accent colour, ~4 pt wide × 30 pt tall — the colour strip |
| 3 | Text box | `CALLS` | 8.5 pt semibold, grey, uppercase |
| 4 | Text box | `=Calc!$W$2` | **19 pt**, semibold, near-black |
| 5 | Text box | `=Calc!$X$2` | 9 pt, grey — the sub-line |

For 4 and 5: select the shape, click in the **formula bar**, type `=`, then click the Calc cell.

Card colours:

| Card | Accent | Overall | Selected |
|---|---|---|---|
| CALLS | brick `#C04433` | `=Calc!$W$2` | `=Calc!$X$2` |
| REVENUE | orange `#E67E22` | `$W$3` | `$X$3` |
| TALK TIME | plum `#7A4C6C` | `$W$4` | `$X$4` |
| AVG RATING | gold `#D69E2E` | `$W$5` | `$X$5` |
| 5-STAR CALLS | teal `#2E7D6F` | `$W$6` | `$X$6` |

## 7.5 Copy the card four times

**Selecting overlapping shapes is fiddly.** Use this instead:

**Home → Find & Select → Select Objects** → your cursor becomes an arrow → drag a box around the whole card → `Ctrl+C` → `Esc` → `Ctrl+V`.

Then edit the pasted card's label, accent colour, and the two cell references.

✅ **Checkpoint:** cards read `1,000` · `$96,623` · `1,498 hrs` · `3.89` · `307`, with `all 5 reps` beneath each. Click **R02** → sub-lines become `R02 · 218`, `R02 · $20,581`, `R02 · 322 hrs`, `R02 · 3.87`, `R02 · 74`. **Big numbers must not change.**

---

# Part 8 — The Six Charts

Every chart is built from the **`Calc`** sheet, never from a pivot. That's what keeps them stable.

## 8.0 The cleanup routine — apply to all six

After creating each chart:

1. Chart Area → Fill **white**, Outline **light grey**, no shadow
2. Delete **gridlines**
3. Delete the **legend** (unless noted)
4. Delete the **value axis** where data labels are shown — never show both
5. Font: 9 pt grey for labels, 11.5 pt semibold dark for the title
6. Title **left-aligned** in the top-left corner, not centred
7. **Format Data Series → Gap Width ≈ 45%** for bars
8. Plot Area fill → **None**

## 8.1 Calls by month — area + line

**The technique:** two copies of the same numbers. One becomes a filled area, one stays a line.

1. Select `Calc!A5:C17`
2. **Insert → Line → Line with Markers**
3. Right-click a series → **Change Series Chart Type** → make **series 1** an **Area**, leave series 2 as Line with Markers
4. Area: brick fill, **transparency 78%**, no border
5. Line: brick, 2.25 pt
6. Markers: **circle, size 6, white fill, brick border**
7. Data labels on the **line** series only, 8 pt grey
8. Delete the value axis, keep the month axis
9. Title: `Calls by month`

Place it full-width under the KPI cards.

## 8.2 & 8.3 Calls by rep / Revenue by rep — the highlight charts

Build the calls one, then repeat for revenue.

1. Select `Calc!H5:I10` → **Insert → Bar → Clustered Bar**
2. **Right-click → Select Data → Add**:
   - Series name: `Selected`
   - Series values: `=Calc!$K$6:$K$10`
3. ⭐ **Format Data Series → Series Overlap = `100%`** ← this is the whole trick
4. Series 1 (`Calls`) → **grey** `#C6C6C4`
5. Series 2 (`Selected`) → **brick**
6. Data labels on **series 1** only (so all five values always show)
7. Category axis → **Format Axis → tick "Categories in reverse order"** so R01 sits at the top
8. Axis → **Labels → Label Position: Low** (keeps labels on the left)
9. Delete the value axis
10. Title: `Calls by rep`

**For the revenue chart:** same steps with `H5:H10` + `J` values, `L` for the highlight, **orange** accent.

⚠️ **Number format on the revenue labels.** Use:

```text
#,##0.0,
```

The trailing comma divides by 1,000, so `20581` displays as `20.6`. Put the unit in the **title**: `Revenue by rep · $ thousands`.

> **Do not put `$` in a chart data-label format.** Excel treats `$` in chart labels as *the system currency symbol* and swaps it for your local one — on an India-locale machine `$#,##0` renders as `₹`. Escaping it doesn't help. Cell number formats are fine; chart labels are not. Unit in the title is the reliable fix.

✅ **Checkpoint:** both charts show 5 grey bars. Click **R02** → the R02 bar turns brick/orange, others stay grey. Clear → all grey.

## 8.4 Calls by day of week

1. Select `Calc!E5:F12` → **Insert → Bar → Clustered Bar**
2. Brick fill, gap width 45%
3. Data labels on
4. **Categories in reverse order** (so Sunday is at the top)
5. Delete the value axis
6. Title: `Calls by day of week`

## 8.5 Customer mix by city — 100% stacked

1. Select `Calc!N5:P8`
2. **Insert → Bar → 100% Stacked Bar**
3. Female → brick · Male → teal
4. Data labels → **Position: Center**, white, bold, 9 pt
5. Categories in reverse order
6. Delete the value axis
7. **Keep a legend** — or better, delete it and add two small coloured squares with text boxes (`Female` / `Male`) in the chart's top-right
8. Title: `Customer mix by city`

⚠️ **Check you picked the right chart type.** In the Insert gallery, *3-D Clustered Bar* sits right next to *100% Stacked Bar*. If your three bars are **different lengths**, you picked the wrong one — in a 100% stack all bars are identical length. Right-click → Change Chart Type → **100% Stacked Bar** (2-D row).

## 8.6 Satisfaction rating distribution

1. Select `Calc!R5:S11` — **six rows** (0★ through 5★)
2. **Insert → Column → Clustered Column**
3. Gold fill, gap width 60%
4. Data labels above the columns
5. Delete the value axis, keep the category axis
6. Title: `Satisfaction rating distribution`

✅ **Checkpoint:** six columns: 1, 8, 59, 197, 428, 307. Because the axis is hard-coded on Calc, it stays six columns for **every** rep — that's the tutorial's homework problem, solved.

---

# Part 9 — The Customer × Rep Matrix

A 15-row table with a revenue heatmap, where the selected rep's column highlights.

## 9.1 Structure

| Row | Content |
|---|---|
| 60 | Section title in `B60`; hint text right-aligned in `I60` |
| 61 | Headers: `City` `Customer` `R01` `R02` `R03` `R04` `R05` `Total` |
| 62–76 | 15 customers |
| 77 | Totals row |
| **78** | **Hidden helper row** |

## 9.2 Fill in the customers

Column **K** (hidden later) holds the city for **every** row — the formulas need it on all 15 rows. Column **B** shows the city only on the first row of each group, so it reads as a grouped table.

| Rows | City (col K, all rows) | Customers (col C) |
|---|---|---|
| 62–65 | Cincinnati | C0003, C0004, C0011, C0012 |
| 66–71 | Cleveland | C0002, C0007, C0008, C0010, C0013, C0015 |
| 72–76 | Columbus | C0001, C0005, C0006, C0009, C0014 |

Put `Cincinnati` in `B62`, `Cleveland` in `B66`, `Columbus` in `B72` — leave the other B cells blank.

## 9.3 The revenue formula

`D62`, then fill across D:H and down to row 76:

```excel
=IFERROR(GETPIVOTDATA("[Measures].[Total Amount]",Pivots!$A$42,
   "[Customers].[City]","[Customers].[City].&["&$K62&"]",
   "[Customers].[Customer ID]","[Customers].[Customer ID].&["&$C62&"]",
   "[Calls].[Rep ID]","[Calls].[Rep ID].&["&D$61&"]"),0)
```

**Watch the `$` signs — they are what makes one formula fill the whole block:**

- `$K62` / `$C62` — column locked, row moves → each row keeps its own city and customer
- `D$61` — row locked, column moves → each column picks up its own rep header

`I62` (row total): `=SUM($D62:$H62)`
`D77` (column total): `=SUM(D62:D76)` — fill across to `I77`

Format `D62:I77` as `$#,##0`.

✅ **Checkpoint:** first row (C0003) = 891 · 1,332 · 1,282 · 739 · 560 → total **4,804**. Row 77 = 18,415 · 20,581 · 20,872 · 16,651 · 20,104 → **96,623**.

## 9.4 The hidden flag row

`D78`, fill across to `H78`:

```excel
=IF(AND(Calc!$B$2<5,COUNTIF(Pivots!$N$14:$N$19,D$61)>0),1,0)
```

This puts a `1` under each selected rep's column. Conditional formatting will read it.

> **Why a helper row instead of putting the logic straight into the rule?** Conditional-formatting formulas that reach across sheets are fragile and impossible to debug. A visible-then-hidden helper row on the same sheet is reliable, and you can unhide it any time to see what the rule is actually seeing.

## 9.5 Conditional formatting — two rules

### Rule 1 — the heatmap

1. Select **`D62:H76`** (data only, not totals)
2. **Home → Conditional Formatting → Colour Scales → More Rules**
3. **2-Colour Scale**
   - Minimum → **Lowest Value** → near-white `#FFF9F6`
   - Maximum → **Highest Value** → soft brick `#F0B0A0`

### Rule 2 — the selected column

1. Select **`D62:H77`** (include the totals row)
2. **Conditional Formatting → New Rule → Use a formula**
3. Formula:

```excel
=D$78=1
```

4. **Format → Font → Bold**, colour **brick**
5. OK

**Use font formatting, not fill.** The colour scale already owns the fill; two fill rules fight and the higher-priority one wins. Bold + colour coexists with the heatmap.

### Rule 3 — the header

1. Select **`D61:H61`**
2. New Rule → formula `=D$78=1`
3. Format: **brick fill, white bold text**

## 9.6 Hide the scaffolding

- Right-click column **K** → **Hide**
- Right-click row **78** → **Hide**

✅ **Checkpoint:** click **R03** → the R03 column turns bold brick and its header goes brick/white. Clear the slicer → no column highlighted.

---

# Part 10 — Polish

## 10.1 Colours

**Page Layout → Colors → Customize Colors** and set Accent 1–6 to your palette. This makes slicers and tables pick up your colours automatically instead of Office blue.

| Role | Hex |
|---|---|
| Primary (calls) | `#C04433` |
| Accent (revenue) | `#E67E22` |
| Secondary (male) | `#2E7D6F` |
| Rating | `#D69E2E` |
| Duration | `#7A4C6C` |
| Muted (unselected) | `#C6C6C4` |
| Text | `#2D3134` |
| Caption | `#7A7E82` |
| Background | `#F0F0ED` |

**Page Layout → Fonts → Customize Fonts** → Segoe UI (or Aptos on 365).

## 10.2 Slicer styling

Slicer tab → Styles → right-click one → **Duplicate**. In the dialog format these elements:

| Element | Set |
|---|---|
| Whole Slicer | Border → your grey |
| Selected Item With Data | Accent fill, **bold** text |
| Unselected Item With Data | Very light fill |

Then apply your new style.

## 10.3 Print / PDF setup

**Page Layout:**

| Setting | Value |
|---|---|
| Print Area | `$A$1:$J$79` |
| Orientation | Portrait |
| Scaling | **Fit Sheet on One Page** |
| Margins | Narrow (0.2") |
| Centre horizontally | ✅ |

## 10.4 Notes sheet

Add a `Notes` sheet recording: the relationship, the five measures, which pivots connect to the slicer, and why Calc exists. Six months from now this is the difference between fixing it in 10 minutes and rebuilding it.

## 10.5 Final tidy

- Sheet order: `Dashboard` · `Data` · `Pivots` · `Calc` · `Notes`
- Clear the slicer
- Click cell `A1` on Dashboard
- Save

---

# Part 11 — Test It

Run all four. Every value must match.

## Test 1 — Nothing selected

| Check | Expected |
|---|---|
| KPI big numbers | 1,000 · $96,623 · 1,498 hrs · 3.89 · 307 |
| KPI sub-lines | `all 5 reps` |
| Subtitle | `2023 · 1,000 calls · 15 customers · 5 representatives` |
| Rep bars | all grey |
| Matrix | no column highlighted |
| Rating chart | 6 columns summing to 1,000 |

## Test 2 — Click R02

| Check | Expected |
|---|---|
| Big numbers | **unchanged** (1,000 / $96,623 …) |
| Sub-lines | `R02 · 218` · `R02 · $20,581` · `R02 · 322 hrs` · `R02 · 3.87` · `R02 · 74` |
| Subtitle | `2023 · R02 · 218 of 1,000 calls` |
| Rep charts | R02 bar coloured, rest grey |
| Matrix | R02 column bold brick |
| Month chart | totals 218 |
| Customer mix | **unchanged** (it's not connected) |

## Test 3 — Ctrl-click R01 and R03

| Check | Expected |
|---|---|
| Sub-line 1 | `R01, R03 · 396` (189 + 207) |
| Sub-line 2 | `R01, R03 · $39,287` (18,415 + 20,872) |
| Rep charts | **two** bars highlighted |
| Matrix | **two** columns highlighted |

## Test 4 — Refresh

1. On `Data`, change one Purchase Amount
2. **Data → Refresh All**
3. Revenue KPI changes
4. Undo, Refresh All again → back to $96,623

If Test 4 fails, your source isn't a real Table or the pivots point at a fixed range.

---

# Part 12 — Troubleshooting

Real problems, in the order you'll hit them.

### "Add this data to the Data Model" is greyed out
The file is in Protected View, or it's `.xls`. Enable Editing and save as `.xlsx`.

### Relationship won't save
The "one" side isn't unique. `Customers[Customer ID]` must have no duplicates. Also check both columns are the same data type.

### "Add Measure" isn't in the right-click menu
The pivot wasn't built on the Data Model. Rebuild it with the checkbox ticked. Also enable Power Pivot: File → Options → Add-ins → COM Add-ins → Go → tick **Microsoft Power Pivot for Excel**.

### Measure syntax error on `=`
In some entry points the leading `=` is added for you, producing `==COUNTROWS(Calls)`. Drop the `=` and enter just `COUNTROWS(Calls)`.

### A KPI card shows `#REF!` or nothing
Shapes can't display `GETPIVOTDATA` results well and can't show multi-cell results. Point the shape at a **single Calc cell**, and let Calc do the lookup.

### Grand total is wrong / too small
A slicer is filtering a pivot that shouldn't be filtered. Select the pivot → **Analyze → Filter Connections** → untick. This *always* happens after copy-pasting a pivot, because the copy inherits the connection.

### A chart's bars change count when you click a rep
You charted the pivot directly. Rebuild that chart on a hard-coded `Calc` range (Part 6).

### Rating chart totals 999
You built a 1–5 axis. One call rounds to 0. Use `R5:S11` — six rows.

### Revenue labels show `₹` instead of `$`
Chart data labels localise `$`. Use `#,##0.0,` and put `$ thousands` in the title (see 8.3).

### Bars in the "mix" chart are different lengths
You picked *3-D Clustered Bar* instead of *100% Stacked Bar*. Change Chart Type → 2-D row → 100% Stacked Bar.

### The highlight never appears
Check `Calc!K6:K10`. All `#N/A` with a rep selected means `Calc!B2` isn't detecting the selection — verify `SelectedRep` is connected to the slicer and its Grand Total is off.

### The highlight shows on all five reps
You left out `AND($B$2<5, ...)` in K/L.

### Weekday chart is upside down
Format Axis → tick **Categories in reverse order**.

### Pivot doesn't show new rows
The source isn't a Table, or you didn't refresh. `Ctrl+T` first, then **Data → Refresh All**.

---

# Part 13 — Extend It

Ordered easiest → hardest.

1. **Rep photo.** Put pictures on an `Assets` sheet (Insert → Picture → **Place in Cell**), then
   `=IF(Calc!$B$2>1,"",XLOOKUP(Calc!$B$1,Assets[Rep ID],Assets[Photo]))`
2. **Rank line.** `=RANK.AVG(selected_calls, Calc!$I$6:$I$10)` → "Rank 2 of 5 by calls"
3. **Share of calls.** `=Pivots!$A$9/Pivots!$A$4` formatted as `0%`
4. **Revenue per call.** New measure: `=DIVIDE([Total Amount],[Call Count])`
5. **Duration bucket chart** — the column you built in 1.5 and never charted
6. **FY running total.** `RunningTotal = CALCULATE([Total Amount], FILTER(ALL(Calls), Calls[Call Date] <= MAX(Calls[Call Date])))`
7. **Homework — time wasters.** Pivot: Customer ID in Rows, Call Count **and** Total Amount in Values. Look for high calls + low revenue. Judge on **both** numbers, never call count alone.
8. **Homework — training candidates.** Pivot: Rep ID in Rows, Avg Rating + Call Count. Then a monthly trend per rep. A low average from few calls is noise; from 200 calls it's a finding.

---

# Appendix — Every Number in This Project

Use this to verify any step.

## Totals

| Metric | Value |
|---|---|
| Calls | 1,000 |
| Revenue | $96,623 |
| Duration | 89,850 min · 1,498 hrs |
| Avg rating | 3.8854 |
| 5-star calls | 307 |
| Customers | 15 |
| Reps | 5 |
| Duration range | 2 – 176 min |
| Date range | 01-Jan-2023 – 31-Dec-2023 |

## By representative

| Rep | Calls | Revenue |
|---|---|---|
| R01 | 189 | $18,415 |
| R02 | 218 | $20,581 |
| R03 | 207 | $20,872 |
| R04 | 186 | $16,651 |
| R05 | 200 | $20,104 |
| **Total** | **1,000** | **$96,623** |

## By month

| Month | Calls | | Month | Calls |
|---|---|---|---|---|
| Jan | 79 | | Jul | 64 |
| Feb | 66 | | Aug | 50 |
| Mar | 155 | | Sep | 76 |
| Apr | 136 | | Oct | 114 |
| May | 83 | | Nov | 57 |
| Jun | 67 | | Dec | 53 |

## By weekday

| Day | Calls |
|---|---|
| Sunday | 146 |
| Monday | 133 |
| Tuesday | 138 |
| Wednesday | 153 |
| Thursday | 128 |
| Friday | 141 |
| Saturday | 161 |

## Rating distribution

| Rating | Calls |
|---|---|
| 0 ★ | 1 |
| 1 ★ | 8 |
| 2 ★ | 59 |
| 3 ★ | 197 |
| 4 ★ | 428 |
| 5 ★ | 307 |

## Duration buckets

| Bucket | Calls |
|---|---|
| Under 10 mins | 9 |
| 10 to 30 mins | 51 |
| 30 to 60 mins | 179 |
| 1 to 2 hours | 524 |
| More than 2 hours | 237 |

## Customer mix by city

| City | Female | Male | Total |
|---|---|---|---|
| Cincinnati | 144 | 132 | 276 |
| Cleveland | 326 | 63 | 389 |
| Columbus | 129 | 206 | 335 |

## Financial year (July – June)

| FY | Calls |
|---|---|
| 2023 | 586 |
| 2024 | 414 |

## Selection test values

| Selection | Calls | Revenue | Talk time | Avg rating | 5-star |
|---|---|---|---|---|---|
| *(none)* | 1,000 | $96,623 | 1,498 hrs | 3.89 | 307 |
| R02 | 218 | $20,581 | 322 hrs | 3.87 | 74 |
| R01 + R03 | 396 | $39,287 | 576 hrs | 3.89 | 120 |

## Customers by city

| City | Customer IDs |
|---|---|
| Cincinnati | C0003, C0004, C0011, C0012 |
| Cleveland | C0002, C0007, C0008, C0010, C0013, C0015 |
| Columbus | C0001, C0005, C0006, C0009, C0014 |

---

## The five ideas that make this project work

Strip away the clicking and this is what you actually learned:

1. **Separate the layers.** Data → Model → Pivots → Calc → Dashboard. Never calculate where you decorate.
2. **Decide what should *not* react.** Four pivots stay unconnected on purpose. Context is what makes a selection meaningful.
3. **Hard-code your categories.** A chart fed by a fixed list never changes shape. That one idea solves the video's homework problem.
4. **`NA()` + 100% overlap = highlighting.** Excel draws nothing for `#N/A`, so a second series only appears where you want it.
5. **A pivot filtered by a slicer is a sensor.** Count its rows and you know what the user clicked — that's the entire interactivity engine, in one `COUNTIF`.
