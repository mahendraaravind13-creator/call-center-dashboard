# Excel for Data Analysis — Complete Notes (0 → Full)

**How to read this file**

Every topic follows the same 4 lines, so you can scan fast:

- **What** — what the thing is, in plain words
- **Use** — when you would actually reach for it
- **How** — the syntax or the click-path
- **Example** — real code you can type

> Tip: don't read this top-to-bottom like a story. Use the contents below, jump to what you need, and **type it in Excel yourself**. Reading Excel notes teaches you nothing; typing them teaches you everything.

---

## Contents

| # | Section | What it covers |
|---|---|---|
| 1 | [Excel Basics](#1-excel-basics) | Screen, cells, data types |
| 2 | [Shortcuts](#2-shortcuts-that-actually-save-time) | The keys worth memorising |
| 3 | [Filter & Sort](#3-filtering-and-sorting) | Finding rows |
| 4 | [Data Quality](#4-data-quality--checking-before-analysing) | Duplicates, blanks, junk |
| 5 | [Tables](#5-excel-tables--the-most-important-habit) | `Ctrl+T`, structured refs |
| 6 | [Formula Rules](#6-formula-rules-you-must-know) | `=`, `$`, `F4`, errors |
| 7 | [Basic Maths](#7-basic-aggregation) | SUM, COUNT, AVERAGE, MIN, MAX |
| 8 | [Conditional Maths](#8-conditional-aggregation--the-ifs-family) | COUNTIFS, SUMIFS, AVERAGEIFS… |
| 9 | [Logic](#9-logic-functions) | IF, IFS, IFERROR |
| 10 | [Dynamic Arrays](#10-dynamic-arrays-modern-excel) | FILTER, UNIQUE, SORT, TAKE |
| 11 | [Lookups](#11-lookups) | VLOOKUP, INDEX+MATCH, XLOOKUP |
| 12 | [Text](#12-text-functions) | LEFT, TEXT, TEXTJOIN, `&` |
| 13 | [Dates](#13-date-functions) | DATE, YEAR, MONTH, weekday, FY |
| 14 | [Ranking](#14-ranking) | LARGE, RANK.AVG |
| 15 | [**Pivot Tables**](#15-pivot-tables--the-full-breakdown) | Every area, every option |
| 16 | [Slicers](#16-slicers-and-timelines) | Interactive filters |
| 17 | [Charts](#17-charts) | Which chart, and cleaning it |
| 18 | [Conditional Formatting](#18-conditional-formatting) | Colour scales, data bars, formulas |
| 19 | [**Power Query**](#19-power-query--the-full-breakdown) | Every cleaning step |
| 20 | [Data Model](#20-data-model--joining-two-tables) | Relationships |
| 21 | [DAX Measures](#21-dax-measures-power-pivot) | COUNTROWS, CALCULATE |
| 22 | [Dashboards](#22-building-a-dashboard) | Architecture + tricks |
| 23 | [Debugging](#23-debugging-checklist) | When it breaks |
| 24 | [What To Use When](#24-what-to-use-when-decision-tables) | Decision tables |
| 25 | [Master Cheat Sheet](#25-master-cheat-sheet) | Everything on one screen |
| 26 | [Practice Data + Answers](#26-practice-datasets-with-answers) | Test yourself |
| 27 | [Final Checklist](#27-final-checklist) | Tick these off |

---

# 1. Excel Basics

## The screen

| Part | What it is |
|---|---|
| **Ribbon** | The command bar at the top (Home, Insert, Data…) |
| **Column** | Vertical — A, B, C… |
| **Row** | Horizontal — 1, 2, 3… |
| **Cell** | One box, e.g. `A1`. Column letter + row number |
| **Range** | A block of cells, e.g. `A1:C10` |
| **Sheet / Tab** | One page of the workbook |
| **Workbook** | The whole `.xlsx` file |
| **Formula bar** | Shows what is really inside the selected cell |
| **Status bar** | Bottom strip — shows Count/Sum/Average of your selection instantly |

> **Free trick:** select any numbers and just *look at the status bar*. It shows Sum, Average and Count without writing a single formula. Right-click it to add more.

## What a cell can hold

| Type | Looks like | Note |
|---|---|---|
| Number | `1000` | Sits on the **right** by default |
| Text | `Asha` | Sits on the **left** by default |
| Date | `01-Jan-23` | Really a number underneath (1 = 1 Jan 1900) |
| Formula | `=SUM(A1:A9)` | Shows result, formula bar shows the code |
| Boolean | `TRUE` / `FALSE` | Result of a test |
| Error | `#N/A`, `#DIV/0!` | Something went wrong |

> **Diagnostic:** if a number sits on the **left**, Excel thinks it's text. That silently breaks SUM. Fix the data type.

---

# 2. Shortcuts That Actually Save Time

## Selecting

| Shortcut | Does |
|---|---|
| `Ctrl + A` | Select the whole data block |
| `Ctrl + Shift + ↓` | Select down to the last filled row |
| `Ctrl + Shift + →` | Select across to the last filled column |
| `Ctrl + Shift + 8` | Select the current data region |
| `Ctrl + ↓` | Jump to the bottom of the data |
| `Ctrl + Home` | Jump back to A1 |

## Doing

| Shortcut | Does |
|---|---|
| `Ctrl + T` | **Turn range into a Table** (do this always) |
| `Ctrl + Shift + L` | Filter on/off |
| `Alt + =` | AutoSum |
| `Ctrl + 1` | Format Cells dialog |
| `Ctrl + Shift + 3` | Format as Date |
| `Ctrl + Shift + 4` | Format as Currency |
| `Ctrl + Shift + 5` | Format as Percentage |
| `F2` | Edit the cell / see which cells the formula uses |
| `F4` | While editing: cycle `A1 → $A$1 → A$1 → $A1` |
| `F4` | While **not** editing: repeat last action |
| `Alt` | Show ribbon keyboard letters |
| `Ctrl + Shift + Enter` | (Old Excel only) array formula — not needed in 365 |

**Ribbon navigation:** press `Alt`, then follow the letters shown.
Example: `Alt → H → S → C` = Home → Sort & Filter → Clear.

---

# 3. Filtering and Sorting

## Filter

**What:** hides rows you don't want to see. Nothing is deleted.
**Use:** quick exploration — "show me only Chennai".
**How:** `Ctrl + Shift + L`, then click the dropdown arrow on a header.

### Filter types

| Column type | Options you get |
|---|---|
| Text | Equals, Contains, Begins With, Ends With |
| Number | Greater Than, Less Than, Between, **Top 10**, Above Average |
| Date | Last Week / Month / Quarter / Year, Between, specific year |
| Colour | Filter by Cell Colour / Font Colour |

### Key behaviour

- Filters **within one column** = OR (Chennai *or* Hyderabad)
- Filters **across columns** = AND (Chennai *and* Female)

### Handy moves

| Task | How |
|---|---|
| Filter to one value fast | Right-click the value → **Filter → By Selected Cell's Value** |
| Clear one column | Click its dropdown → Clear Filter |
| Clear everything | Home → Sort & Filter → **Clear** |

> **Top 10 filter** is not locked to 10. You can change it to Top 5, Bottom 20, or Top 10 **Percent**.

## Sort

**What:** reorders rows.
**Use:** ranking, or grouping similar rows together.

| Task | How |
|---|---|
| Simple | Click column → A→Z / Z→A / Smallest→Largest |
| **Multi-level** | **Data → Sort** → Add Level |

**Multi-level example** — "Males first, and inside Males sort by salary":

```text
Level 1: Gender   (A → Z)
Level 2: Salary   (Largest → Smallest)
```

This is different from just sorting Salary. Level 1 is the outer group, Level 2 sorts inside it.

> **Danger:** never sort one column alone by selecting just that column. Excel may scramble your rows. Always select the whole table (or use a real Excel Table, which protects you).

---

# 4. Data Quality — Checking BEFORE Analysing

Analysis on dirty data is worse than no analysis. Check these five things every time.

| Problem | How to spot it | What to do |
|---|---|---|
| **Duplicates** | Home → Conditional Formatting → Highlight Cells Rules → **Duplicate Values**, then Filter by Colour | Investigate. An ID should appear once |
| **Blanks / nulls** | Filter dropdown → look for `(Blanks)` | Don't auto-delete. Label as `Missing` |
| **Junk values** (`?`, `N/A`, `-`) | Filter dropdown shows every distinct value | Replace with the correct value |
| **Wrong type** | Numbers sitting on the left = text | Change type (Power Query or Text to Columns) |
| **Hidden spaces** | `"Training "` ≠ `"Training"` | `TRIM()` or Power Query → Trim |

**The golden rule of missing data:**

> Do not delete a record just because a field is blank. First ask *why* it's blank. Flag it (`Missing`, `Need to check`) so it stays visible and countable.

**Clear formatting rules:** Home → Conditional Formatting → **Clear Rules** (from selection or whole sheet).

---

# 5. Excel Tables — The Most Important Habit

**What:** a range converted into a smart object with a name.
**Use:** *always*, before doing anything else.
**How:** click any cell in the data → `Ctrl + T` → tick "My table has headers" → rename it in **Table Design → Table Name**.

## Why it matters

| Without a Table | With a Table |
|---|---|
| `=SUM(F2:F261)` | `=SUM(staff[Salary])` |
| Breaks when rows are added | **Grows automatically** |
| Hard to read | Reads like English |
| Pivot source is fixed | Pivot picks up new rows on refresh |

## Structured references

```excel
staff                  ' the whole table (data only, no headers)
staff[Salary]          ' one column
staff[#Headers]        ' the header row
staff[#All]            ' headers + data
staff[[Emp ID]:[Last Name]]   ' a block of adjacent columns
[@Salary]              ' this row's Salary (used inside the table)
```

**Name your tables properly.** `Table1` tells you nothing. `staff`, `calls`, `customers` tell you everything.

---

# 6. Formula Rules You Must Know

## The basics

- Every formula starts with `=`
- A function looks like `=NAME(argument1, argument2)`
- Arguments are separated by commas
- Text goes in `"double quotes"`; numbers do not

## References — the `$` rule

| Written | Called | When copied it… |
|---|---|---|
| `A1` | Relative | moves both ways |
| `$A$1` | Absolute | never moves |
| `A$1` | Mixed | row locked, column moves |
| `$A1` | Mixed | column locked, row moves |

Press **`F4`** while editing to cycle through these.

**When do you need `$`?** When you copy a formula and one reference must keep pointing at the same cell — like a single threshold cell or a tax rate.

## Errors and what they mean

| Error | Means | Usual fix |
|---|---|---|
| `#N/A` | Lookup found nothing | Wrap in `IFERROR`, or check spelling/spaces |
| `#VALUE!` | Wrong data type (text where a number is needed) | Fix the data type |
| `#REF!` | The referenced cell was deleted | Rewrite the reference |
| `#DIV/0!` | Divided by zero or blank | `=IFERROR(a/b,0)` |
| `#NAME?` | Misspelled function or missing quotes | Check spelling |
| `#SPILL!` | A dynamic array has no room | Clear the cells below/right |
| `#####` | Column too narrow | Widen the column (not an error) |

## Nesting

Nesting = putting one function inside another. The **inside runs first**.

```excel
=TAKE(SORT(staff[Salary],,-1),5)
```

Read it inside-out:

```text
staff[Salary]  →  SORT descending  →  TAKE the first 5
```

---

# 7. Basic Aggregation

| Function | Use | Example |
|---|---|---|
| `SUM` | Add numbers | `=SUM(staff[Salary])` |
| `AVERAGE` | Mean | `=AVERAGE(staff[Salary])` |
| `MEDIAN` | Middle value — better than average when there are extreme values | `=MEDIAN(staff[Salary])` |
| `MIN` | Smallest | `=MIN(staff[Salary])` |
| `MAX` | Largest | `=MAX(staff[Salary])` |
| `COUNT` | Count **numbers only** | `=COUNT(staff[Salary])` |
| `COUNTA` | Count **anything not empty** (incl. text) | `=COUNTA(staff[Emp ID])` |
| `COUNTBLANK` | Count empty cells | `=COUNTBLANK(staff[Gender])` |

## COUNT vs COUNTA — the classic trap

Employee IDs like `E001` are **text**. `COUNT` returns **0** for them.
For headcount always use:

```excel
=COUNTA(staff[Emp ID])
```

## AVERAGE vs MEDIAN

If one person earns ₹50,00,000 and nine earn ₹50,000, the **average** is misleading. The **median** is honest. Report both when salaries or prices are involved.

---

# 8. Conditional Aggregation — the `IFS` Family

This is the most useful group of functions in business Excel.

## The one rule

They all work in **pairs**:

```text
criteria_range , criteria , criteria_range , criteria , ...
```

Range first, then what to match. Always.

## The functions

| Function | Syntax | Starts with |
|---|---|---|
| `COUNTIFS` | `COUNTIFS(range, criteria, ...)` | a criteria range |
| `SUMIFS` | `SUMIFS(sum_range, range, criteria, ...)` | **the numbers to add** |
| `AVERAGEIFS` | `AVERAGEIFS(avg_range, range, criteria, ...)` | **the numbers to average** |
| `MINIFS` | `MINIFS(min_range, range, criteria, ...)` | **the numbers to scan** |
| `MAXIFS` | `MAXIFS(max_range, range, criteria, ...)` | **the numbers to scan** |

> **Notice the odd one out:** `COUNTIFS` begins with a criteria range. The other four begin with the **range you want to do maths on**. That is the #1 thing people get backwards.

## Examples

```excel
' How many people in the department named in A2?
=COUNTIFS(staff[Department],A2)

' How many PERMANENT people in that department?
=COUNTIFS(staff[Department],A2,
          staff[Employee type],"Permanent")

' Total salary for that department
=SUMIFS(staff[Salary],staff[Department],A2)

' Average salary for permanent staff in that department
=AVERAGEIFS(staff[Salary],
            staff[Department],A2,
            staff[Employee type],"Permanent")

' Highest salary among Females
=MAXIFS(staff[Salary],staff[Gender],"Female")
```

## Operators inside criteria

| You want | Write |
|---|---|
| Equals | `A2` or `"Permanent"` |
| Greater than | `">100000"` |
| Greater than a cell | `">"&F1` |
| Not equal | `"<>Permanent"` |
| Starts with | `"A*"` |
| Contains | `"*man*"` |
| Blank / not blank | `""` / `"<>"` |

> Note the `&` when comparing to a **cell**: `">"&F1`, not `">F1"`.

## ⚠️ The mistake everybody makes

```excel
' WRONG — a range where a criterion should be
=COUNTIFS(staff[Department],staff[Employee type],"Permanent",A2:A13)

' RIGHT
=COUNTIFS(staff[Department],A2,staff[Employee type],"Permanent")
```

Say it out loud before you type it:

> "Count rows where **Department = A2** AND **Employee type = Permanent**."

Each "where X = Y" is one pair.

---

# 9. Logic Functions

## IF

**Use:** one condition, two outcomes.

```excel
=IF(condition, value_if_true, value_if_false)
=IF(B2>100000,"High","Normal")
```

## Nested IF vs IFS

**IFS** — cleaner when you have several bands.

```excel
=IFS(
    condition1, result1,
    condition2, result2,
    TRUE,       default
)
```

**The `TRUE` at the end is the "else".** Without it, unmatched rows return `#N/A`.

Real example — the duration buckets used in the call-centre file:

```excel
=IFS(
    [@Duration]<=10, "Under 10 mins",
    [@Duration]<=30, "10 to 30 mins",
    [@Duration]<=60, "30 to 60 mins",
    [@Duration]<=120,"1 to 2 hours",
    TRUE,            "More than 2 hours"
)
```

**Order matters.** IFS stops at the first TRUE. Always go smallest → largest (or the reverse), never jumbled.

## IFERROR

**Use:** hide an ugly error behind a readable message.

```excel
=IFERROR(anything, "Not found")
=IFERROR(VLOOKUP(C4,staff,2,FALSE),"Not found")
```

> **Warning:** `IFERROR` hides *all* errors, including your own typos. Build the formula first, confirm it works, *then* wrap it.

## AND / OR / NOT

```excel
=IF(AND(A2>100000, B2="Female"), "Yes", "No")   ' both must be true
=IF(OR(A2="Mumbai", A2="Delhi"), "Metro", "Other")  ' either
=IF(NOT(A2="Permanent"), "Contract", "Permanent")
```

---

# 10. Dynamic Arrays (Modern Excel)

**What:** one formula that returns **many cells**. The result "spills" out automatically.
**Use:** live extracts and lists that resize themselves.
**Needs:** Excel 365 / Excel 2021 / Excel on the web.

## The spill operator `#`

If a formula in `B4` spills down to `B12`, then:

```excel
=B4#
```

means "the whole spilled block, however big it is now".

Use it so your downstream formulas never break when the list grows:

```excel
=COUNTA(B4#)
```

## The functions

### UNIQUE — remove duplicates

```excel
=UNIQUE(staff[Department])
=SORT(UNIQUE(staff[Department]))     ' unique AND alphabetical
```

### SORT

```excel
=SORT(array, [sort_index], [sort_order], [by_col])
=SORT(staff[Salary])        ' ascending
=SORT(staff[Salary],,-1)    ' descending  (note the two commas)
=SORT(staff,6,-1)           ' sort whole table by column 6, descending
```

> The two commas matter. `-1` must land in the **third** slot (sort_order).

### FILTER — return all matching rows

```excel
=FILTER(array, include, [if_empty])
=FILTER(staff, staff[Salary]>100000)
=FILTER(staff, staff[Salary]>F67, "None found")
```

### FILTER with several conditions — `*` and `+`

| Symbol | Means | Because |
|---|---|---|
| `*` | **AND** | TRUE×TRUE=1, anything×FALSE=0 |
| `+` | **OR** | TRUE+FALSE=1 |

```excel
' Female AND above 100K
=FILTER(staff,(staff[Gender]="Female")*(staff[Salary]>100000))

' Mumbai OR Delhi
=FILTER(staff,(staff[Work location]="Mumbai")+(staff[Work location]="Delhi"))

' Three conditions
=FILTER(staff,
    (staff[Gender]="Female")*
    (staff[Salary]>100000)*
    (staff[Start Date]>=DATE(2020,1,1)))
```

**Wrap every condition in its own brackets.** Missing brackets is the usual cause of a wrong answer here.

### TAKE / DROP

```excel
=TAKE(array, rows, [cols])
=TAKE(SORT(staff[Salary],,-1),5)      ' top 5
=TAKE(SORT(staff[Salary],,-1),I12)    ' top N — N lives in a cell
=DROP(array,1)                        ' remove the first row
```

Negative numbers take from the bottom: `=TAKE(array,-5)` = last 5.

### CHOOSECOLS / CHOOSEROWS

**Use:** keep only the columns you care about.

```excel
=CHOOSECOLS(staff,1,2,3,5,6)

' filtered AND trimmed to 5 columns
=CHOOSECOLS(FILTER(staff,staff[Salary]>100000),1,2,3,5,6)

' matching headers for that extract
=CHOOSECOLS(staff[#Headers],1,2,3,5,6)
```

### TRANSPOSE

Flips rows into columns.

```excel
=TRANSPOSE(XLOOKUP(C4,staff[Last Name],staff))
```

### Others worth knowing

| Function | Use |
|---|---|
| `SEQUENCE(10)` | Make 1,2,3…10 (handy for `LARGE`) |
| `TEXTSPLIT` | Split text into columns by a delimiter |
| `HSTACK` / `VSTACK` | Glue arrays side-by-side / on top |
| `LET` | Name a value once, reuse it, faster |

---

# 11. Lookups

**The question they answer:** "Given an ID/name, fetch this person's other details."

## VLOOKUP (old but everywhere)

```excel
=VLOOKUP(lookup_value, table, column_number, FALSE)
=VLOOKUP(C4, staff, 5, FALSE)      ' 5th column = Department
```

| Rule | Why |
|---|---|
| **Always type `FALSE`** | `FALSE` = exact match. Leave it out and Excel guesses — wrong answers, no error |
| Searches only the **first column** | Can't look leftwards |
| Column number is fragile | Insert a column and every VLOOKUP breaks |

**Improvement:** put the column numbers (`2, 3, 5, 6`) in helper cells and reference them, with the lookup value locked as `$C$4`.

## MATCH — find the position

```excel
=MATCH(value, range, 0)      ' 0 = exact match
=MATCH(C4, staff[Last Name], 0)     ' → e.g. 43
```

## INDEX — fetch by position

```excel
=INDEX(return_range, position)
=INDEX(staff[Emp ID], 43)
```

## INDEX + MATCH — the classic combo

```excel
=INDEX(staff[Department], MATCH(C4, staff[Last Name], 0))
```

Read it as: **MATCH finds the row → INDEX grabs the value from that row.**

**Why bother:** unlike VLOOKUP, the answer column can be **anywhere**, left or right.

## XLOOKUP — use this one

```excel
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
```

```excel
=XLOOKUP(C4, staff[Emp ID], staff[Salary])
=XLOOKUP(C4, staff[Emp ID], staff[Salary], "Not found")
=XLOOKUP(C4, staff[Last Name], staff[Emp ID], "Not found")   ' looks LEFT, fine
=XLOOKUP(C4, staff[Last Name], staff)     ' returns the WHOLE row
```

### Why XLOOKUP beats the others

| | VLOOKUP | INDEX+MATCH | XLOOKUP |
|---|---|---|---|
| Exact match by default | ✗ | ✗ | ✓ |
| Can look left | ✗ | ✓ | ✓ |
| Built-in "not found" | ✗ | ✗ | ✓ |
| Return a whole row | ✗ | ✗ | ✓ |
| Survives inserted columns | ✗ | ✓ | ✓ |
| Easy to read | ✓ | ✗ | ✓ |

### Optional arguments

| `match_mode` | Meaning |
|---|---|
| `0` | Exact (default) |
| `-1` | Exact, else next **smaller** |
| `1` | Exact, else next **larger** |
| `2` | Wildcard (`*`, `?`) |

| `search_mode` | Meaning |
|---|---|
| `1` | First to last (default) |
| `-1` | **Last to first** — use to get the most recent record |

## ⚠️ The limitation of ALL lookups

They return the **first match only**.

Two people named "Tawell"? You get one of them, silently.

**When you need every match, stop using lookups and use `FILTER`:**

```excel
=FILTER(staff, staff[Last Name]=C4)
```

| You need | Use |
|---|---|
| One value | `XLOOKUP` |
| All matching rows | `FILTER` |

---

# 12. Text Functions

| Function | Use | Example | Result |
|---|---|---|---|
| `LEFT` | First N characters | `=LEFT("Asha",1)` | `A` |
| `RIGHT` | Last N characters | `=RIGHT("E001",3)` | `001` |
| `MID` | From the middle | `=MID("ABCDE",2,3)` | `BCD` |
| `LEN` | How many characters | `=LEN("Asha")` | `4` |
| `TRIM` | Remove extra spaces | `=TRIM("  Asha ")` | `Asha` |
| `UPPER`/`LOWER`/`PROPER` | Change case | `=PROPER("asha rao")` | `Asha Rao` |
| `TEXT` | Number → formatted text | `=TEXT(0.186,"0%")` | `19%` |
| `TEXTJOIN` | Join a list with a separator | see below | |
| `&` | Glue text together | `=A2&" "&B2` | `Asha Rao` |
| `SUBSTITUTE` | Replace text | `=SUBSTITUTE(A2,"*","")` | removes `*` |

## TEXTJOIN

```excel
=TEXTJOIN(delimiter, ignore_empty, text...)
=TEXTJOIN(", ", TRUE, SORT(UNIQUE(staff[Department])))
```

→ `Finance, HR, IT, Sales`

Combine with FILTER to list everyone who ties for the top salary in one cell:

```excel
=TEXTJOIN(", ",TRUE,
    FILTER(staff[First Name]&" "&staff[Last Name],
           staff[Salary]=MAX(staff[Salary])))
```

## TEXT — the formatting codes

**Use:** when you put a number inside a sentence, it loses its formatting. `TEXT` puts it back.

| Code | Turns `0.186` / a date into |
|---|---|
| `"0%"` | `19%` |
| `"0.0%"` | `18.6%` |
| `"#,##0"` | `1,000` |
| `"$#,##0"` | `$1,000` |
| `"dddd"` | `Monday` |
| `"ddd"` | `Mon` |
| `"mmm"` | `Jan` |
| `"mmmm yyyy"` | `January 2023` |

```excel
="Selected rep handled "&TEXT(A1,"#,##0")&" calls ("&TEXT(B1,"0%")&")"
```

---

# 13. Date Functions

Excel dates are **numbers**. `1` = 1 Jan 1900. That's why you can subtract them.

| Function | Use | Example |
|---|---|---|
| `TODAY()` | Today's date (updates) | `=TODAY()` |
| `NOW()` | Date + time | `=NOW()` |
| `DATE` | Build a date | `=DATE(2020,1,1)` |
| `YEAR` | Get the year | `=YEAR(A2)` |
| `MONTH` | Get month number | `=MONTH(A2)` → `3` |
| `DAY` | Get day number | `=DAY(A2)` |
| `TEXT(a,"dddd")` | Weekday name | `Monday` |
| `EOMONTH` | End of month | `=EOMONTH(A2,0)` |
| `DATEDIF` | Difference in y/m/d | `=DATEDIF(A2,TODAY(),"y")` |

## Common date patterns

```excel
' Joined in March, any year
=FILTER(staff, MONTH(staff[Start Date])=3)

' Joined 2020 or later
=FILTER(staff, staff[Start Date]>=DATE(2020,1,1))

' Started on a Monday
=FILTER(staff, TEXT(staff[Start Date],"dddd")="Monday")

' Years of service
=DATEDIF(staff[Start Date], TODAY(), "y")
```

## Financial Year

Most companies don't use Jan–Dec. If the FY runs **July → June**:

```excel
=YEAR(A2)+IF(MONTH(A2)>=7,1,0)
```

| Date | FY |
|---|---|
| March 2023 | 2023 |
| July 2023 | 2024 |
| Dec 2023 | 2024 |

> This is exactly the rule used in the call-centre dataset in this folder.

Equivalent longer form:

```excel
=IF(MONTH(A2)<=6, YEAR(A2), YEAR(A2)+1)
```

---

# 14. Ranking

## LARGE / SMALL

**Use:** the Nth biggest/smallest value.

```excel
=LARGE(staff[Salary],1)     ' biggest
=LARGE(staff[Salary],2)     ' 2nd biggest
=SMALL(staff[Salary],1)     ' smallest
```

For a top-5 list, put `1,2,3,4,5` in helper cells and drag, **or** use the modern way:

```excel
=TAKE(SORT(staff[Salary],,-1),5)
```

## RANK.AVG / RANK.EQ

**Use:** "where does this person stand out of everyone?"

```excel
=RANK.AVG(number, ref, [order])
=RANK.AVG(B2, staff[Salary])        ' 0/omitted = descending (1 = highest)
=RANK.AVG(B2, staff[Salary], 1)     ' 1 = ascending
```

| Function | Ties get |
|---|---|
| `RANK.EQ` | The same rank (1,1,3) |
| `RANK.AVG` | The average rank (1.5, 1.5, 3) |

---

# 15. Pivot Tables — The Full Breakdown

**What:** a drag-and-drop summary tool. It groups your rows and does maths on them.
**Use:** any "total/count/average **by** category" question — without writing formulas.
**How:** click any cell in the data → **Insert → PivotTable**.

> Convert to a Table (`Ctrl+T`) **first**. Then the pivot picks up new rows when you refresh.

## 15.1 The four areas

This is the whole concept. Everything else is detail.

```text
┌─────────────────────────────────────┐
│ FILTERS   ← dropdown above the table │
├──────────┬──────────────────────────┤
│ ROWS     │ VALUES                   │
│ ↓ down   │ the numbers              │
│ the side │ COLUMNS → across the top │
└──────────┴──────────────────────────┘
```

| Area | Put here | Result |
|---|---|---|
| **Rows** | Categories you want listed **down** | Customer ID, Department, Month |
| **Columns** | A second category listed **across** | Gender, Year, Rating |
| **Values** | The **numbers** to calculate | Sum of Amount, Count of Calls |
| **Filters** | A field to filter the whole report by | Representative, Region |

**Rule of thumb:** text/categories → Rows or Columns. Numbers → Values.

## 15.2 Summarize Values By

Right-click any value → **Summarize Values By**:

| Option | Use |
|---|---|
| **Sum** | Totals — money, quantity (default for numbers) |
| **Count** | How many rows (default for text) |
| **Average** | Typical value — ratings, prices |
| **Max / Min** | Biggest / smallest |
| **Count Numbers** | Count only numeric cells |
| **Product** | Multiply |
| **StdDev / Var** | Spread of the data |

> **Gotcha:** if a numeric column shows **Count** instead of **Sum**, you have blanks or text in it. Fix the data — don't just switch the setting.

## 15.3 Show Values As — the hidden superpower

Right-click a value → **Show Values As**. This changes how the number is *displayed*, not the data.

| Option | Shows | Use it for |
|---|---|---|
| **No Calculation** | Raw number | Default |
| **% of Grand Total** | Share of everything | Contribution |
| **% of Row Total** | Each row adds to 100% | **Comparing categories fairly when totals differ** |
| **% of Column Total** | Each column adds to 100% | Same, other direction |
| **% Of** | Compared to a chosen item | Index vs a baseline |
| **Difference From** | Gap vs previous/chosen item | Month-on-month change |
| **% Difference From** | Growth % | Growth rates |
| **Running Total In** | Cumulative | **Year-to-date** |
| **% Running Total In** | Cumulative share | Pareto / 80-20 |
| **Rank Smallest → Largest** | Rank number | Leaderboards |
| **Index** | Relative importance | Advanced comparison |

**Why `% of Row Total` matters:** Rep A took 500 calls, Rep B took 100. Raw counts say A is busier on every day. Percentages reveal *when each rep is actually busiest*. Different question, better answer.

## 15.4 Grouping

**Use:** turn many detailed values into a few useful buckets.

| Data | How |
|---|---|
| Dates | Drop into Rows → Excel auto-groups into Years/Quarters/Months. Right-click → **Group** to change |
| Numbers | Right-click → **Group** → set Start, End, By |
| Text | Select the items → right-click → **Group** |

### ⚠️ The grouping trap

Manual groups are frozen to the values that existed **when you made them**. New data outside those boundaries lands in the wrong place — or breaks the group.

**Better:** make a real bucket column with `IFS` in the source data. Then new rows classify themselves forever.

```excel
=IFS([@Duration]<=10,"Under 10 mins",
     [@Duration]<=30,"10 to 30 mins",
     [@Duration]<=60,"30 to 60 mins",
     [@Duration]<=120,"1 to 2 hours",
     TRUE,"More than 2 hours")
```

## 15.5 Sorting and filtering inside a pivot

| Task | How |
|---|---|
| Sort | Click the row-label dropdown → A→Z, or Sort by a value column |
| Custom order | Type the order manually, or drag items |
| **Top 10** | Row label dropdown → **Value Filters → Top 10** (changeable to Top 5, Bottom N, Top N%) |
| Greater than | **Value Filters → Greater Than** |
| Contains | **Label Filters → Contains** |

## 15.6 Report Layout and cosmetics

**Design tab:**

| Setting | What it does |
|---|---|
| **Compact Form** | Default — all row fields in one column |
| **Outline Form** | Each field in its own column, with subtotals |
| **Tabular Form** | Each field its own column, table-like — **best for reusing the output** |
| Repeat All Item Labels | Fills the blanks under grouped labels |
| Subtotals | Show/hide/move |
| Grand Totals | On/off for rows and/or columns |
| Blank Rows | Spacing between groups |

**PivotTable Analyze tab:**

| Setting | What it does |
|---|---|
| **PivotTable Name** | Name it! `SummaryPivot` beats `PivotTable7` |
| Field Settings | Change number format, summary type |
| **Insert Slicer** | Add a visual filter |
| **Insert Timeline** | Date slider |
| **Filter Connections** | Which slicers control *this* pivot |
| **Refresh** | Reload the data |
| **Field Buttons** | Hide the grey buttons on pivot charts |

## 15.7 Refreshing

Pivots do **not** update by themselves.

| Action | Effect |
|---|---|
| Right-click → **Refresh** | Update this pivot |
| **Data → Refresh All** | Update everything, including Power Query |
| Analyze → Options → Data → *Refresh data when opening the file* | Auto-refresh on open |

New rows only appear if the source is a **Table** (or the range was widened).

## 15.8 GETPIVOTDATA

When you click a pivot cell while writing a formula, Excel writes this instead of a plain reference:

```excel
=GETPIVOTDATA("Total Amount",$A$3,"Rep ID","R02")
```

| | Plain reference (`=Pivots!B4`) | GETPIVOTDATA |
|---|---|---|
| Simple to read | ✓ | ✗ |
| Breaks if the pivot moves/resizes | ✓ breaks | ✗ safe |
| Handles missing categories | ✗ | ✓ (wrap in `IFERROR`) |

**To turn it off:** PivotTable Analyze → Options ▾ → untick **Generate GetPivotData**. One-time setting.

**Which to use:** plain references for quick KPI cards; GETPIVOTDATA when the pivot's shape can change (the safer choice for real dashboards).

## 15.9 Copying pivots — the classic bug

Copy-pasting a pivot **also copies its slicer connection**. The copy silently stays filtered, and your "grand total" is wrong.

**Fix:** select the copy → Analyze → **Filter Connections** → untick the slicer → then clear any leftover filter.

## 15.10 Pivot question → setup

| Question | Rows | Columns | Values |
|---|---|---|---|
| Calls per customer | Customer ID | — | Count of Call ID |
| Revenue per customer | Customer ID | — | Sum of Amount |
| Avg rating per rep | Rep ID | — | Average of Rating |
| Monthly trend | Date (grouped) | — | Count of Call ID |
| Busiest weekday per rep | Rep ID | Day of Week | Count of Call ID |
| Duration vs satisfaction | Duration Bucket | Rating Rounded | Count of Call ID |
| Rep workload by month | Rep ID | Month | Count of Call ID |
| Revenue by city & gender | City | Gender | Sum of Amount |

---

# 16. Slicers and Timelines

## Slicer

**What:** big clickable buttons that filter pivots.
**Use:** dashboards — so the reader filters without touching pivot fields.
**How:** select the pivot → **Insert Slicer** → tick a field. (Or right-click the field → Add as Slicer.)

| Task | How |
|---|---|
| Select several | Hold `Ctrl`, or click the multi-select icon |
| Clear | The little "clear filter" ✗ on the slicer |
| **Control several pivots** | Right-click slicer → **Report Connections** → tick each pivot |
| Disconnect one pivot | Report Connections → untick, **or** the pivot's Filter Connections |
| Columns / size | Slicer tab → Columns (e.g. 5 for a horizontal strip) |
| Styling | Slicer tab → Styles → right-click a style → **Duplicate** to customise |

> Slicers only control pivots built on the **same data source / cache**. If a slicer refuses to connect, the pivot came from a different source.

## Timeline

**What:** a slider for dates.
**Use:** filter by month/quarter/year by dragging.
**How:** Insert → Timeline (only works on real date fields).

---

# 17. Charts

## Which chart for which job

| Question | Chart |
|---|---|
| Trend over time | **Line** (or Area) |
| Compare categories | **Bar** (horizontal) or **Column** (vertical) |
| Part of a whole | **100% Stacked Bar** — avoid pies |
| Composition + total | Stacked Column |
| Relationship between two numbers | **Scatter** |
| Distribution | Histogram / Column |
| Progress to target | Bar with a target line |

> **Bar vs Column:** long category names → Bar (horizontal). Time on the axis → Column or Line.

## Cleaning a chart ("chart junk")

Default Excel charts are noisy. Remove, in this order:

1. Gridlines (or make them very light)
2. The value axis **if** you have data labels — never both
3. Legend, if there is only one series or you've coloured the title text
4. Chart border and background fill
5. **Field buttons** on pivot charts: Analyze → Field Buttons → Hide All
6. Tick marks

Then **add**: a clear title that states the finding, data labels, and consistent colours.

## Useful chart settings

| Setting | Where | Use |
|---|---|---|
| **Gap Width** | Format Data Series | Lower = thicker bars (25–50% looks good) |
| **Series Overlap** | Format Data Series | `100%` = bars sit on top of each other (highlight trick) |
| **Categories in reverse order** | Format Axis | Fixes upside-down bar charts |
| Axis Minimum | Format Axis | **Set to 0** — a truncated axis exaggerates differences |
| Number format on labels | Format Data Labels | `#,##0` / `0.0%` |
| Combo chart | Change Chart Type → Combo | Line + Area, or two different scales |

## The Line + Area trick

To get a filled trend with a crisp line on top:

1. Add the **same** measure to the pivot **twice**
2. Make the chart a Line chart
3. Change **series 1** to **Area**
4. Make the area semi-transparent, keep the line solid

## The highlight trick (very useful)

Make the selected item a different colour without changing the chart:

1. Helper column: `=IF(rep=selected_rep, value, NA())`
2. Add it as a **second series**
3. Set **Series Overlap = 100%**

`NA()` values are simply not drawn, so only the selected bar shows the highlight colour on top of the grey one.

---

# 18. Conditional Formatting

**What:** formatting that changes with the value.
**Use:** make patterns visible without reading numbers.
**Where:** Home → Conditional Formatting.

| Type | Use |
|---|---|
| **Highlight Cells Rules** | Duplicates, greater than, text contains |
| **Top/Bottom Rules** | Top 10, above average |
| **Data Bars** | In-cell bar — compare magnitudes at a glance |
| **Colour Scales** | Heatmap — 2 or 3 colours |
| **Icon Sets** | Arrows / traffic lights |
| **New Rule → Use a formula** | Anything custom |

## Formula-based rules

**Use:** highlight a whole row, or react to another cell.

The rule is written **for the top-left cell of your selected range**, then applied relatively.

```excel
' Highlight the whole row when Salary > 100000
=$F2>100000          ' $ locks the column, row moves

' Highlight the column belonging to the selected rep
=D$1=$B$1            ' $ locks the row, column moves
```

| Lock pattern | Effect |
|---|---|
| `$F2` | Whole **row** highlights |
| `F$2` | Whole **column** highlights |
| `$F$2` | One fixed cell decides everything |

**Managing rules:** Conditional Formatting → **Manage Rules**. Order matters — the top rule wins. Use **Stop If True** to prevent overlap.

> **Tip:** a colour scale (fill) and a formula rule (bold/font colour) can coexist on the same cells. Two *fill* rules cannot — the higher one wins.

---

# 19. Power Query — The Full Breakdown

**What:** a data cleaning tool built into Excel. It **records** every step you take.
**Use:** whenever you'll receive the same messy file again. Clean it once, press Refresh forever.
**Where:** the **Data** tab.

> The whole point: **you never clean the same file twice.**

## 19.1 The Editor

| Part | What it is |
|---|---|
| **Ribbon** | Home / Transform / Add Column / View |
| **Queries pane** (left) | Your list of queries |
| **Data preview** (middle) | Sample of the current table |
| **Applied Steps** (right) | **The recording** — every step, in order |
| **Formula bar** | The M code for the selected step |

**Applied Steps is the most important panel.** Click any step to time-travel to that moment.

| Task | How |
|---|---|
| Edit a step | Click the **⚙ gear** icon beside it |
| Delete a step | Click the **✗** |
| Rename a step | Right-click → Rename |
| Reorder | Drag (carefully) |
| **Change the source file** | Click the **Source** step → point at the new path |

## 19.2 Getting data

| Source | Path |
|---|---|
| Excel file | Data → Get Data → From File → **From Excel Workbook** |
| CSV / Text | Data → From Text/CSV |
| Website | Data → **From Web** → paste URL |
| Folder (many files!) | Get Data → From File → **From Folder** |
| Database | Get Data → From Database |

Then the **Navigator** appears → pick the table → click **Transform Data**, *not* Load.

> Always click **Transform Data**. Loading first and cleaning later defeats the purpose.

## 19.3 Every cleaning step you need

| Step | Where | Use |
|---|---|---|
| **Use First Row as Headers** | Home | The real headers are stuck in row 1 |
| **Rename column** | Double-click header | `No.` → `Country` |
| **Remove Columns** | Right-click → Remove | Drop what you don't need |
| **Choose Columns** | Home | Keep only what you need |
| **Remove Rows** | Home → Remove Rows | Top N, blanks, duplicates, errors |
| **Keep Rows** | Home → Keep Rows | Opposite |
| **Filter** | Column dropdown | Drop Totals rows, nulls |
| **Replace Values** | Right-click → Replace Values | `*` → blank, `?` → `Engineering` |
| **Fill Down / Up** | Transform → Fill | Copy a value into the blanks below (merged cells) |
| **Trim** | Transform → Format → Trim | Kill leading/trailing spaces |
| **Clean** | Transform → Format → Clean | Remove invisible characters |
| **Change Type** | Right-click → Change Type | Text → Number/Date. **Do this** |
| **Split Column** | Transform → Split Column | Full Name → First + Last |
| **Merge Columns** | Transform → Merge | Opposite |
| **Conditional Column** | Add Column | An IF ladder, no code |
| **Custom Column** | Add Column | Write M code |
| **Index Column** | Add Column | Add a row number |
| **Group By** | Transform → Group By | Summarise (like a pivot) |
| **Unpivot Columns** | Transform → Unpivot | **Turn a wide report into clean data** |
| **Merge Queries** | Home | Join two tables (like VLOOKUP) |
| **Append Queries** | Home | Stack tables on top of each other |

## 19.4 Data types — do not skip this

Click the icon left of each header and set the right type.

| Type | For |
|---|---|
| Text `ABC` | Names, IDs |
| Whole Number `123` | Counts |
| Decimal Number | Money, ratings |
| **Date** | Dates — unlocks all date transforms |
| Percentage | Ratios |
| True/False | Flags |

**Wrong types are the #1 cause of Power Query pain.** Text dates can't be grouped by month; text numbers can't be summed.

## 19.5 Column Quality (View tab)

Tick these three boxes:

| Tool | Shows |
|---|---|
| **Column quality** | % valid / error / empty |
| **Column distribution** | How many distinct and unique values |
| **Column profile** | Full stats for the selected column |

Use these **before** cleaning to see what's actually broken.

> By default it profiles only the first 1000 rows. Click the status bar to switch to the whole dataset.

## 19.6 Split Column — the middle-name problem

Splitting `First Middle Last` by space from the **left** gives you a mess.

Split by the **right-most delimiter** instead → you reliably get `First Middle` | `Last`.

Transform → Split Column → By Delimiter → Space → **Right-most delimiter**.

## 19.7 Conditional Column

Add Column → Conditional Column. It builds an IF ladder with clicks:

```text
if      Salary < 50000   then "Under 50K"
else if Salary < 100000  then "50K to 100K"
else                          "Above 100K"
```

## 19.8 Custom Column (M language)

Add Column → Custom Column:

```text
if [FTE] = 1 then "Full time" else "Part time"
```

M basics:

| M | Note |
|---|---|
| `[ColumnName]` | Reference a column |
| `if … then … else …` | **lowercase**, and `else` is compulsory |
| `&` | Join text |
| `//` | Comment |

M is **case-sensitive**. `If` fails; `if` works.

## 19.9 Loading and refreshing

**Close & Load To…** gives you options:

| Option | Use |
|---|---|
| Table | Normal — output to a sheet |
| PivotTable Report | Straight to a pivot |
| **Only Create Connection** | No sheet output — for staging queries and the Data Model |
| **Add this data to the Data Model** | Needed for relationships and DAX |

**Refresh:** right-click the output → Refresh, or **Data → Refresh All**.

## 19.10 The lesson about hard-coded filters

If you remove a "Total" row by filtering out the value `1,080`, that filter is **frozen to that number**. Next month's file has a different total and the Total row sneaks back in.

**Better:** filter on something structural — "Country does not equal Total", or remove the **last row** by position.

---

# 20. Data Model — Joining Two Tables

**What:** Excel's built-in mini-database. It lets pivots use **fields from two or more tables at once**.
**Use:** when facts live in one table and attributes in another.

**Example:** `Calls` has the transactions. `Customers` has City and Gender. You want "calls by city" — that needs both.

## How to build it

1. Make both ranges Tables (`Ctrl+T`) and name them
2. Insert → PivotTable → tick **"Add this data to the Data Model"**
3. Do the same for the second table (or load both via Power Query with "Add to Data Model")
4. **Data → Relationships → New** (or PivotTable Analyze → Relationships)
5. Join them on the shared column

```text
Calls[Customer ID]   →   Customers[Customer ID]
   (many side)              (one side)
```

## The rules of a relationship

| Rule | Meaning |
|---|---|
| **One side must be unique** | `Customers[Customer ID]` — each customer appears once |
| **Many side can repeat** | `Calls[Customer ID]` — one customer, many calls |
| Types must match | Text-to-text, number-to-number |
| The "one" table is the **lookup/dimension** | Customers, Products, Calendar |
| The "many" table is the **fact** table | Calls, Sales, Transactions |

Get this backwards and the relationship refuses to save.

## What changes afterwards

The pivot field list gains an **All / Active** toggle, and lists both tables. You can now drag `City` from Customers and `Call Count` from Calls into the same pivot.

---

# 21. DAX Measures (Power Pivot)

**What:** a formula that lives in the Data Model, not in a cell.
**Use:** reusable business calculations that automatically respect whatever filters, slicers and rows the pivot has.

**How:** right-click the table in the field list → **Add Measure** (or Power Pivot → Measures → New Measure).

## Measure vs Calculated Column

| | Calculated Column | Measure |
|---|---|---|
| Calculated | Once per row, stored | On the fly, per pivot cell |
| Uses memory | Yes | Almost none |
| Reacts to slicers | No | **Yes** |
| Use it for | Categories, buckets, flags | Totals, averages, KPIs |

**Rule:** if you want to *slice by it* → column. If you want to *measure it* → measure.

## The core measures

```dax
Call Count     = COUNTROWS(Calls)

Total Amount   = SUM(Calls[Purchase Amount])

Total Duration = SUM(Calls[Call Duration])

Avg Rating     = AVERAGE(Calls[Satisfaction Rating])

5 Star Calls   = CALCULATE([Call Count], Calls[Rating Rounded] = 5)
```

## CALCULATE — the one DAX function to understand

```dax
CALCULATE( <expression>, <filter1>, <filter2>, ... )
```

**Plain English:** "Work out this number, but **change the filters** first."

```dax
5 Star Calls = CALCULATE([Call Count], Calls[Rating Rounded] = 5)
```

= "Count the calls, but only where the rating is 5."

It still respects the pivot row and slicer, then adds its own filter on top. That's why selecting R02 gives *R02's* five-star count, not everyone's.

## Other DAX worth knowing

| Function | Use |
|---|---|
| `DIVIDE(a,b,0)` | Safe division — no `#DIV/0!` |
| `DISTINCTCOUNT(col)` | Count unique values |
| `COUNTROWS(table)` | Count rows |
| `ALL(table)` | Ignore filters — needed for "% of total" |
| `RELATED(col)` | Pull a value across a relationship |
| `IF(cond, a, b)` | Same as Excel |

Percentage-of-total pattern:

```dax
% of All Calls = DIVIDE([Call Count], CALCULATE([Call Count], ALL(Calls)))
```

## Formatting a measure

Set the number format **when you create the measure**. Then every pivot using it is formatted correctly, forever. No repeated manual formatting.

---

# 22. Building a Dashboard

## The architecture — keep layers separate

```text
Raw Data          (Data sheet — never touch by hand)
    ↓
Power Query       (cleaning, if needed)
    ↓
Tables / Data Model + Relationships
    ↓
Pivots + DAX Measures     (Pivots sheet — the calculation layer)
    ↓
Helper / Calc ranges      (Calc sheet — chart-ready numbers)
    ↓
Slicer + Charts + KPI cards   (Dashboard sheet — the only pretty one)
```

**The rule:** never calculate on the dashboard sheet. Never decorate on the pivots sheet.

## Sheets to create

| Sheet | Holds |
|---|---|
| `Data` | The source tables |
| `Pivots` | All pivot tables, named properly |
| `Calc` | Helper formulas that feed charts |
| `Dashboard` | Only visuals |
| `Notes` | How it works (future-you will thank you) |

## KPI cards

1. Insert → Shapes → Rounded Rectangle
2. Select the shape, click the **formula bar**, type `=` and click the pivot cell
3. Format the font, remove the outline

A shape can only show **one cell**. So one card = a few shapes: background + label + big number + sub-caption.

## Which pivots connect to the slicer?

This is the design decision that makes a dashboard actually useful:

| Pivot | Slicer connected? | Why |
|---|---|---|
| Selected-rep KPIs | **Yes** | Should react |
| Trend charts | **Yes** | Should react |
| Overall KPIs | **No** | So you can compare selection vs total |
| All-reps comparison | **No** | You need all bars visible to highlight one |

## Making charts stable

If a category has no rows for the current selection, it vanishes and the chart changes shape — which looks broken.

**Fix:** build a helper range with **all** categories hard-coded, filled by `GETPIVOTDATA` wrapped in `IFERROR`:

```excel
=IFERROR(GETPIVOTDATA("Call Count",Pivots!$A$13,"Month",$A6),0)
```

Missing category → `0` instead of disappearing. The chart never moves.

## Showing what's selected

```excel
' how many items the slicer has selected
=COUNTA(SelectedRepPivotRange)

' the selected rep's share of all calls
=SelectedRepCalls/TotalCalls          ' format as %

' rank
=RANK.AVG(SelectedRepCalls, AllRepCalls)

' hide the summary when several reps are selected
=IF(COUNTA(selected)>1,"",calculation)
```

That last one matters: a "rank" for a multi-selection is meaningless. Blank it out.

## Dynamic image of the selected rep

```excel
=IF(COUNTA(selected)>1, "", XLOOKUP(SelectedRep, Assets[Rep ID], Assets[Image]))
```

Excel 365 can put a picture **in a cell** (Insert → Picture → **Place in Cell**), which makes it lookup-able.

## Design principles

- Answer **business questions**, not "show all the data"
- **5 KPIs maximum** at the top
- One clear message per chart
- Same colours mean the same thing everywhere
- Grey for context, one bright colour for the thing you care about
- Remove chart junk
- Align everything (hold `Alt` while dragging to snap to cells)
- Label units in the **title**, not on every data point

---

# 23. Debugging Checklist

Work down this list in order.

**1. Argument order**
`COUNTIFS` = range, criteria pairs. `SUMIFS` = **sum range first**. Mixing them up is the most common error.

**2. Data types**
Number stored as text? Date stored as text? Check the alignment — text hugs the left.

**3. Hidden spaces**
`"Training "` ≠ `"Training"`. Use `TRIM`, or `=LEN(A2)` to count characters.

**4. Exact match missing**
`VLOOKUP` needs `FALSE`. `MATCH` needs `0`. Without them you get plausible but wrong answers.

**5. `#SPILL!`**
Something is blocking the output. Clear the cells below/right.

**6. Wrong table or column name**
Typo in `staff[Departmnt]` → `#NAME?`.

**7. Missing `$`**
The formula worked in the first cell and broke when copied? You needed absolute references.

**8. Pivot not updating**
Pivots never auto-refresh. Right-click → Refresh. New rows missing? Your source isn't a Table.

**9. Pivot total looks too small**
A slicer is filtering it. Check **Filter Connections**.

**10. Brackets in FILTER**
Each condition needs its own `( )`:
`(a="x")*(b>1)` ✓ — `a="x"*b>1` ✗

---

# 24. What To Use When (Decision Tables)

## By question type

| You need | Use |
|---|---|
| A total | `SUM` |
| Count of rows | `COUNTA` |
| Count with conditions | `COUNTIFS` |
| Sum with conditions | `SUMIFS` |
| Average with conditions | `AVERAGEIFS` |
| Min/max with conditions | `MINIFS` / `MAXIFS` |
| Nth largest | `LARGE` |
| Top N, dynamic | `TAKE(SORT(...))` |
| **All** matching rows | `FILTER` |
| **One** matching value | `XLOOKUP` |
| List without duplicates | `UNIQUE` |
| Unique + sorted | `SORT(UNIQUE(...))` |
| Join list into one cell | `TEXTJOIN` |
| Hide an error | `IFERROR` |
| Bucket a number | `IFS` |
| Rank someone | `RANK.AVG` |

## By tool

| Situation | Tool |
|---|---|
| Quick summary by category | **PivotTable** |
| Same messy file every month | **Power Query** |
| Two tables to combine | **Data Model + relationship** |
| A KPI reused in many pivots | **DAX measure** |
| Let the reader filter | **Slicer** |
| Show a pattern visually | **Conditional formatting** |
| One-off calculation | **Formula** |

## Formula or Pivot?

| Use a **Pivot** when | Use a **Formula** when |
|---|---|
| Exploring, don't know the answer yet | You know exactly what you need |
| Many groupings quickly | Result must sit in a specific cell |
| Non-Excel people will use it | It must update with no Refresh |
| Data is big | You need it inside another formula |

They work together: pivots calculate, formulas pull the numbers onto the dashboard.

---

# 25. Master Cheat Sheet

```excel
' ---------- AGGREGATION ----------
=SUM(range)                =AVERAGE(range)      =MEDIAN(range)
=MIN(range)                =MAX(range)
=COUNT(range)              =COUNTA(range)       =COUNTBLANK(range)

' ---------- CONDITIONAL ----------
=COUNTIFS(range,crit, range2,crit2)
=SUMIFS(sum_range, range,crit, ...)
=AVERAGEIFS(avg_range, range,crit, ...)
=MINIFS(min_range, range,crit, ...)
=MAXIFS(max_range, range,crit, ...)

' ---------- LOGIC ----------
=IF(test, yes, no)
=IFS(test1,r1, test2,r2, TRUE,default)
=IFERROR(formula,"Not found")
=AND(a,b)   =OR(a,b)   =NOT(a)

' ---------- DYNAMIC ARRAYS ----------
=UNIQUE(range)
=SORT(range)              =SORT(range,,-1)        ' descending
=FILTER(array, condition, "None")
=FILTER(t,(a="x")*(b>1))                          ' AND
=FILTER(t,(a="x")+(a="y"))                        ' OR
=TAKE(SORT(range,,-1), 5)
=CHOOSECOLS(array,1,2,3)
=TRANSPOSE(array)
B4#                                               ' spill reference

' ---------- LOOKUP ----------
=XLOOKUP(val, lookup_col, return_col, "Not found")
=XLOOKUP(val, lookup_col, return_col, "NF", 0, -1)   ' last match
=VLOOKUP(val, table, col_num, FALSE)
=INDEX(return_col, MATCH(val, lookup_col, 0))
=FILTER(table, lookup_col=val)                       ' ALL matches

' ---------- TEXT ----------
=LEFT(t,1)      =RIGHT(t,3)     =MID(t,2,3)     =LEN(t)
=TRIM(t)        =PROPER(t)      =SUBSTITUTE(t,"*","")
=TEXT(0.186,"0%")               =A2&" "&B2
=TEXTJOIN(", ",TRUE,array)

' ---------- DATE ----------
=TODAY()        =DATE(2020,1,1)
=YEAR(d)        =MONTH(d)       =DAY(d)
=TEXT(d,"dddd")                 ' weekday name
=TEXT(d,"mmm")                  ' Jan
=DATEDIF(start,TODAY(),"y")     ' years of service
=YEAR(d)+IF(MONTH(d)>=7,1,0)    ' July–June financial year

' ---------- RANK ----------
=LARGE(range,1)      =SMALL(range,1)
=RANK.AVG(value, range)

' ---------- DAX ----------
Call Count   = COUNTROWS(Calls)
Total Amount = SUM(Calls[Purchase Amount])
Avg Rating   = AVERAGE(Calls[Satisfaction Rating])
5 Star Calls = CALCULATE([Call Count], Calls[Rating Rounded]=5)
% of Total   = DIVIDE([Call Count], CALCULATE([Call Count], ALL(Calls)))
```

## The three patterns to know cold

```excel
' 1. Count with two conditions
=COUNTIFS(staff[Department],A2, staff[Employee type],"Permanent")

' 2. Filter with two conditions
=FILTER(staff,(staff[Gender]="Female")*(staff[Salary]>100000))

' 3. Modern lookup
=XLOOKUP(C4, staff[Emp ID], staff[Salary], "Not found")
```

Understand these three and most of the rest is variation.

---

# 26. Practice Datasets (With Answers)

Paste into Excel, press `Ctrl+T`, name the table, then answer without looking.

## Table `staff`

```text
Emp ID | First Name | Gender | Department | Salary | Start Date | FTE | Employee type | Work location
E001   | Asha       | Female | Sales      | 72000  | 2022-02-10 | 1.0 | Permanent  | Mumbai
E002   | Rahul      | Male   | Sales      | 85000  | 2020-03-15 | 0.8 | Permanent  | Delhi
E003   | Neha       | Female | HR         | 68000  | 2021-01-20 | 1.0 | Permanent  | Mumbai
E004   | Arjun      | Male   | IT         | 110000 | 2019-07-05 | 1.0 | Permanent  | Pune
E005   | Priya      | Female | IT         | 125000 | 2023-03-11 | 1.0 | Permanent  | Hyderabad
E006   | Karan      | Male   | HR         | 62000  | 2024-03-18 | 0.5 | Temporary  | Delhi
E007   | Sneha      | Female | Sales      | 98000  | 2022-03-21 | 1.0 | Fixed Term | Mumbai
E008   | Vikram     | Male   | IT         | 118000 | 2021-04-02 | 0.8 | Permanent  | Pune
E009   | Riya       | Female | Finance    | 76000  | 2020-03-25 | 1.0 | Permanent  | Delhi
E010   | Aman       | Male   | Finance    | 92000  | 2018-11-19 | 1.0 | Permanent  | Mumbai
E011   | Isha       | Female | IT         | 105000 | 2020-06-30 | 1.0 | Permanent  | Pune
E012   | Rohit      | Male   | Sales      | 49000  | 2025-02-14 | 0.5 | Temporary  | Chennai
```

### Questions and answers

| # | Question | Formula | Answer |
|---|---|---|---|
| 1 | Total salary | `=SUM(staff[Salary])` | 1,060,000 |
| 2 | Total FTE | `=SUM(staff[FTE])` | 10.6 |
| 3 | Headcount | `=COUNTA(staff[Emp ID])` | 12 |
| 4 | Unique depts | `=SORT(UNIQUE(staff[Department]))` | Finance, HR, IT, Sales |
| 5 | Highest salary | `=MAX(staff[Salary])` | 125,000 |
| 6 | Lowest salary | `=MIN(staff[Salary])` | 49,000 |
| 7 | Average salary | `=AVERAGE(staff[Salary])` | 88,333 |

**Headcount by department** — `=COUNTIFS(staff[Department],A2)`

| Dept | Count |
|---|---|
| Finance | 2 |
| HR | 2 |
| IT | 4 |
| Sales | 4 |

**Permanent headcount** — `=COUNTIFS(staff[Department],A2,staff[Employee type],"Permanent")`

| Dept | Permanent |
|---|---|
| Finance | 2 |
| HR | 1 |
| IT | 4 |
| Sales | 2 |

**Total & average salary by department**

| Dept | `SUMIFS` | `AVERAGEIFS` | `MAXIFS` | `MINIFS` |
|---|---|---|---|---|
| Finance | 168,000 | 84,000 | 92,000 | 76,000 |
| HR | 130,000 | 65,000 | 68,000 | 62,000 |
| IT | 458,000 | 114,500 | 125,000 | 105,000 |
| Sales | 304,000 | 76,000 | 98,000 | 49,000 |

**More**

| Question | Formula | Answer |
|---|---|---|
| Earning > 100K | `=FILTER(staff,staff[Salary]>100000)` | Arjun, Priya, Vikram, Isha |
| Female > 100K | `=FILTER(staff,(staff[Gender]="Female")*(staff[Salary]>100000))` | Priya, Isha |
| Top 5 salaries | `=TAKE(SORT(staff[Salary],,-1),5)` | 125000, 118000, 110000, 105000, 98000 |
| Highest paid person | `=XLOOKUP(MAX(staff[Salary]),staff[Salary],staff[First Name])` | Priya |
| Names starting with A | `=FILTER(staff,LEFT(staff[First Name],1)="A")` | Asha, Arjun, Aman |
| Joined in March | `=FILTER(staff,MONTH(staff[Start Date])=3)` | Rahul, Priya, Karan, Sneha, Riya |
| Median salary, IT | `=MEDIAN(FILTER(staff[Salary],staff[Department]="IT"))` | 114,000 |
| Female ratio, IT | `=COUNTIFS(staff[Department],"IT",staff[Gender],"Female")/COUNTIFS(staff[Department],"IT")` | 50% |

> Check #1 and #2 yourself by selecting the column and reading the status bar. Always verify your own answers.

## Table `calls`

```text
Call Number | Customer ID | Representative | Duration | Date       | Purchase Amount | Satisfaction Rating
C001        | Cust01      | Asha           | 12       | 2023-01-05 | 300             | 4.5
C002        | Cust01      | Rahul          | 45       | 2023-01-12 | 0               | 3.2
C003        | Cust02      | Neha           | 8        | 2023-02-01 | 100             | 4.8
C004        | Cust02      | Asha           | 50       | 2023-02-18 | 50              | 3.5
C005        | Cust03      | Rahul          | 35       | 2023-03-03 | 1200            | 4.7
C006        | Cust03      | Vikram         | 70       | 2023-03-15 | 0               | 2.9
C007        | Cust03      | Rahul          | 60       | 2023-03-20 | 0               | 2.5
C008        | Cust04      | Sneha          | 20       | 2023-04-02 | 800             | 4.9
C009        | Cust04      | Asha           | 25       | 2023-04-10 | 0               | 3.0
C010        | Cust05      | Vikram         | 90       | 2023-05-05 | 2000            | 4.6
C011        | Cust05      | Priya          | 110      | 2023-05-15 | 2500            | 4.9
C012        | Cust06      | Neha           | 15       | 2023-06-01 | 0               | 4.1
C013        | Cust06      | Aman           | 10       | 2023-07-10 | 0               | 4.4
C014        | Cust07      | Rohit          | 5        | 2023-08-22 | 0               | 3.8
C015        | Cust07      | Rahul          | 55       | 2023-09-11 | 0               | 2.7
C016        | Cust07      | Rahul          | 65       | 2023-10-11 | 0               | 2.6
```

## Table `customers`

```text
Customer ID | Gender | Age | City
Cust01      | Female | 30  | Mumbai
Cust02      | Male   | 42  | Delhi
Cust03      | Male   | 28  | Pune
Cust04      | Female | 35  | Mumbai
Cust05      | Female | 31  | Delhi
Cust06      | Female | 45  | Pune
Cust07      | Male   | 26  | Chennai
```

Relate them: `calls[Customer ID]` (many) → `customers[Customer ID]` (one).

### Pivot answers

**Calls per customer** (Rows: Customer ID, Values: Count of Call Number)

| Customer | Calls | Revenue |
|---|---|---|
| Cust01 | 2 | 300 |
| Cust02 | 2 | 150 |
| Cust03 | 3 | 1,200 |
| Cust04 | 2 | 800 |
| Cust05 | 2 | 4,500 |
| Cust06 | 2 | 0 |
| Cust07 | 3 | 0 |

**Average rating per rep** (Rows: Representative, Values: Average of Rating)

| Rep | Calls | Avg rating |
|---|---|---|
| Rahul | 5 | 3.14 |
| Asha | 3 | 3.67 |
| Vikram | 2 | 3.75 |
| Rohit | 1 | 3.80 |
| Aman | 1 | 4.40 |
| Neha | 2 | 4.45 |
| Sneha | 1 | 4.90 |
| Priya | 1 | 4.90 |

### Homework, solved

**Q: Who are the "time wasters"?** (many calls, no money)

Pivot with Customer ID in Rows, and **both** Call Count and Sum of Amount in Values.

| Answer | Why |
|---|---|
| **Cust07** | 3 calls, ₹0 revenue — the clearest case |
| Cust06 | 2 calls, ₹0 |
| Cust03 | 3 calls but ₹1,200 — **not** a time waster |

> The lesson: never judge on call count alone. You need both numbers side by side.

**Q: Which rep needs satisfaction training?**

Pivot: Rep in Rows, Average Rating + Call Count in Values. Then a second pivot with Date in Rows filtered to that rep, to see the trend.

**Answer: Rahul** — lowest average (3.14), highest volume (5 calls), and his recent ratings (2.7, 2.6) are worse than his early ones.

> Always check volume alongside the average. A 2.0 average from one call is noise; from 50 calls it's a problem.

---

# 27. Final Checklist

Tick these off **in Excel**, not on paper.

## Basics
- [ ] Filter, multi-filter, number filter, date filter
- [ ] Sort, and multi-level sort
- [ ] Find duplicates with conditional formatting
- [ ] Filter by colour
- [ ] AutoSum (`Alt + =`)
- [ ] Create and name a Table (`Ctrl + T`)

## Formulas
- [ ] SUM / AVERAGE / MEDIAN / MIN / MAX
- [ ] COUNT vs COUNTA (know the difference)
- [ ] COUNTIFS with two conditions
- [ ] SUMIFS / AVERAGEIFS / MINIFS / MAXIFS
- [ ] IF / IFS with `TRUE` as the else
- [ ] IFERROR
- [ ] `$` and `F4`

## Dynamic arrays
- [ ] UNIQUE, SORT, SORT+UNIQUE
- [ ] FILTER with one condition
- [ ] FILTER with `*` (AND) and `+` (OR)
- [ ] TAKE + SORT for Top N
- [ ] CHOOSECOLS
- [ ] The `#` spill reference

## Lookups
- [ ] VLOOKUP with FALSE
- [ ] INDEX + MATCH
- [ ] XLOOKUP with "not found"
- [ ] XLOOKUP looking left
- [ ] FILTER for multiple matches

## Text and dates
- [ ] LEFT, TRIM, TEXTJOIN, `&`
- [ ] TEXT for `0%` and `dddd`
- [ ] DATE, YEAR, MONTH
- [ ] Financial-year formula

## Pivots
- [ ] Build one from a Table
- [ ] Use all four areas
- [ ] Change Sum → Average
- [ ] Top 10 value filter
- [ ] Group dates by month
- [ ] Running Total In
- [ ] % of Row Total
- [ ] Rename the pivot
- [ ] Refresh after changing data

## Interactivity
- [ ] Insert a slicer
- [ ] Connect it to several pivots
- [ ] Disconnect one pivot
- [ ] Build a pivot chart
- [ ] Hide field buttons

## Power Query
- [ ] Import from a file
- [ ] Promote headers
- [ ] Replace values
- [ ] Fill down
- [ ] Trim
- [ ] Change type
- [ ] Split column
- [ ] Conditional column
- [ ] Edit and delete a step
- [ ] Refresh after the source changes

## Model and dashboard
- [ ] Add two tables to the Data Model
- [ ] Create a relationship
- [ ] Write a DAX measure
- [ ] Write a CALCULATE measure
- [ ] Build a KPI card linked to a pivot
- [ ] Apply a colour scale
- [ ] Build a formula-based highlight rule

---

# The Real Skill

Don't memorise function names. Learn to translate a business question into a tool.

```text
What exactly am I being asked?
        ↓
One value, or many rows, or a summary?
        ↓
One value      → XLOOKUP
Many rows      → FILTER
A summary      → PivotTable
Repeated mess  → Power Query
Two tables     → Data Model + DAX
        ↓
Can I make it update by itself?
```

Everything above is just the vocabulary. **This** is the grammar.

> Worked example:
> *"Count permanent employees by department"*
> → I need a **count** → with **two conditions** → so `COUNTIFS`
> → Department = A2, Employee type = "Permanent"
> → `=COUNTIFS(staff[Department],A2,staff[Employee type],"Permanent")`

---

## See also

The workbook **`Call Center Dashboard.xlsx`** in this folder is a working example of Sections 15, 16, 19, 20, 21 and 22 — the data model, the five DAX measures, nine pivots, the slicer wiring, the highlight trick and the conditional-formatting matrix are all live in it. Open its **Notes** sheet for a map of how those pieces connect.
