# Access Database Analysis Skill

Analyze data inside Microsoft Access databases (`.accdb`/`.mdb`) using Jet/ACE SQL recipes distilled from the *Access Data Analysis Cookbook*. This skill is **read/report-oriented**: profiling, data-quality auditing, aggregation, statistics, and trend analysis — not schema/form/VBA development.

## When to Use This Skill

Use this skill when the user asks you to:
- Profile a table or database (row counts, nulls, distinct values, data types)
- Audit data quality (orphans/unmatched records, duplicates, referential-integrity gaps)
- Summarize or aggregate data (sums, averages, counts, group-bys, crosstabs)
- Compute statistics (mean/median/mode, variance/std dev, correlation, percentile, rank, frequency, growth rate)
- Find trends over time (moving averages, elapsed-time calculations, year/month/day breakdowns)
- Sample data (random rows, top/bottom N, distinct records)
- Export analysis results for charting (PivotTable/PivotChart in Excel)

## Guiding Principle: Analysis Is Read-Only

Every recipe below is a `SELECT`/`TRANSFORM` statement executed with `Invoke-AccessSQL -DbPath $db -SQL "..." -AsJson`. Do **not** run `UPDATE`/`DELETE`/`INSERT`/`ALTER` against the user's data as part of analysis. If a technique in the cookbook is naturally destructive (e.g., a make-table query used just to stage results), instead:
- Prefer a plain `SELECT` and let the caller review JSON output, or
- If a scratch table genuinely helps (e.g., precomputed rank/percentile for reuse), create it with `New-AccessTable`/`Invoke-AccessSQL ... INTO`, tell the user it was created, and offer to drop it (`DROP TABLE`, `-ConfirmDestructive`) when done.

## Workflow

1. **Profile** — understand the shape of the data before analyzing it (Section 1).
2. **Audit** — check for data-quality problems that would skew results (Section 2).
3. **Filter/Sample** — narrow to the relevant rows (Section 3).
4. **Aggregate** — summarize with GROUP BY, crosstabs, calculated fields (Section 4).
5. **Analyze statistically** — mean/median/variance/correlation/percentile/trend (Sections 5–6).
6. **Report** — present JSON results back to the user, or export to Excel for charting (Section 7).

All SQL below uses `{Table}`, `{Field}`, `{Key}` as placeholders — substitute real identifiers. Wrap any identifier containing spaces or a reserved word in `[brackets]`.

---

## 1. Data Profiling & Discovery

Start every analysis task here.

```powershell
Get-AccessObject -DbPath $db -ObjectType table -AsJson
Get-AccessTableInfo -DbPath $db -TableName "{Table}" -AsJson   # fields, types, sizes, indexes
```

SQL profiling templates:

```sql
-- Row count
SELECT Count(*) AS RowCount FROM {Table};

-- Null / non-null counts per column (repeat Sum(IIf(...)) per field)
SELECT
  Sum(IIf([{Field}] Is Null, 1, 0)) AS NullCount,
  Sum(IIf([{Field}] Is Not Null, 1, 0)) AS NonNullCount
FROM {Table};

-- Distinct value count
SELECT Count(*) AS DistinctCount FROM (SELECT DISTINCT {Field} FROM {Table});

-- Min / Max / Avg for a numeric or date column
SELECT Min([{Field}]) AS MinVal, Max([{Field}]) AS MaxVal, Avg([{Field}]) AS AvgVal
FROM {Table};
```

## 2. Data Quality & Integrity Audits

**Unmatched / orphan records** (Cookbook 1.1) — rows in a child table with no match in the parent:

```sql
SELECT Child.*
FROM {ChildTable} Child LEFT JOIN {ParentTable} Parent
  ON Child.{Key} = Parent.{Key}
WHERE Parent.{Key} Is Null;
```

**Duplicate detection** — rows that repeat on one or more "identity" fields:

```sql
SELECT {Field1}, {Field2}, Count(*) AS Occurrences
FROM {Table}
GROUP BY {Field1}, {Field2}
HAVING Count(*) > 1
ORDER BY Count(*) DESC;
```

**Exclude a known-bad set with NOT IN + subquery** (Cookbook 1.4):

```sql
SELECT * FROM {Table}
WHERE {Key} NOT IN (SELECT {Key} FROM {ExclusionTable});
```

**Referential-integrity check** — confirm every FK value exists in the parent (returns 0 rows if clean):

```sql
SELECT DISTINCT Child.{ForeignKey}
FROM {ChildTable} Child LEFT JOIN {ParentTable} Parent
  ON Child.{ForeignKey} = Parent.{Key}
WHERE Child.{ForeignKey} Is Not Null AND Parent.{Key} Is Null;
```

Also useful: `Get-AccessRelationship -DbPath $db -AsJson` to see what relationships/enforcement already exist, and `Search-AccessData -DbPath $db -SearchText "..." -AsJson` to locate a value across tables.

## 3. Filtering, Set Logic & Sampling

- **AND/OR precedence** — group `OR` conditions in parentheses so they don't leak across an `AND` boundary (Cookbook 1.2): `WHERE State="NY" AND (City="Yonkers" OR City="Albany")`.
- **IN** instead of long OR chains: `WHERE {Field} In ("A","B","C")`.
- **Top/bottom N**: `SELECT TOP 10 * FROM {Table} ORDER BY {Field} DESC;` — add `PERCENT` for a percentage slice. Note: `TOP` without an `ORDER BY` is meaningless (order is undefined).
- **Distinct vs. DistinctRow**: `SELECT DISTINCT` collapses duplicate rows on the *selected columns only*; `SELECT DISTINCTROW` (multi-table queries) collapses duplicates based on all underlying table columns, not just the ones displayed.
- **Random sample** — the `Rnd()` argument must reference a per-row field or Access evaluates it once for the whole query: `SELECT TOP 25 * FROM {Table} ORDER BY Rnd(-(Id)*Timer());` or simpler, `ORDER BY Rnd({SomeNumericField})`.
- **Combine similar tables** with `UNION` (dedups) or `UNION ALL` (keeps duplicates):
  ```sql
  SELECT {Field1}, {Field2} FROM {TableA}
  UNION ALL
  SELECT {Field1}, {Field2} FROM {TableB};
  ```
- **Left/Right joins** for "all of A, matched with B where possible": `FROM {A} LEFT JOIN {B} ON {A}.{Key}={B}.{Key}`.

## 4. Aggregation, Grouping & Crosstabs

**Basic aggregates:**

```sql
SELECT {GroupField}, Sum({ValueField}) AS Total, Avg({ValueField}) AS Average,
       Count(*) AS N, Min({ValueField}) AS Lowest, Max({ValueField}) AS Highest
FROM {Table}
GROUP BY {GroupField}
HAVING Count(*) > 1        -- filter on the aggregate, not the raw rows
ORDER BY Total DESC;
```

**Calculated / on-the-fly fields** (Cookbook 1.11, 2.3): give an expression an alias with `:`.

```sql
SELECT [FirstName] & " " & [LastName] AS FullName,
       IIf([Balance] <= 100, "Low Balance", "OK") AS Flag
FROM {Table};
```

**Crosstab / pivot** (Cookbook 2.7) — one row-heading field, one column-heading field, one aggregated value field:

```sql
TRANSFORM Count({KeyField}) AS Total
SELECT {RowField}
FROM {Table}
GROUP BY {RowField}
PIVOT {ColumnField};
```

## 5. Statistical Analysis

Jet/ACE supports these aggregate functions directly in SQL (no VBA needed): `Avg`, `Sum`, `Count`, `Min`, `Max`, `StDev`, `StDevP`, `Var`, `VarP`.

**Mean:** `SELECT Avg({Field}) AS Mean FROM {Table};`

**Mode** (most frequent value):

```sql
SELECT TOP 1 {Field}, Count(*) AS Occurrences
FROM {Table}
GROUP BY {Field}
ORDER BY Count(*) DESC;
```

**Median** (Cookbook 10.2) — no native function; use a self-join count comparison:

```sql
SELECT A.{Field}
FROM {Table} AS A, {Table} AS B
GROUP BY A.{Field}
HAVING Sum(IIf(A.{Field} <= B.{Field}, 1, 0)) >= Count(*) / 2
   AND Sum(IIf(A.{Field} >= B.{Field}, 1, 0)) >= Count(*) / 2;
```

**Variance & standard deviation:** `SELECT Var({Field}) AS Variance, StDev({Field}) AS StdDev FROM {Table};` — use the `P` variants (`VarP`, `StDevP`) when the table *is* the whole population rather than a sample.

**Covariance** (Cookbook 10.4) — how two columns move together:

```sql
SELECT Sum(({FieldX} - (SELECT Avg({FieldX}) FROM {Table})) *
           ({FieldY} - (SELECT Avg({FieldY}) FROM {Table}))) / Count(*) AS Covariance
FROM {Table};
```

**Correlation** — same as above, divided by `StDev({FieldX}) * StDev({FieldY})`; result is bounded -1..1.

**Histogram / frequency buckets** (Cookbook 10.1) — bucket a numeric field with nested `Sum(IIf(...))`:

```sql
SELECT
  Sum(IIf([{Field}] < 25, 1, 0)) AS Under25,
  Sum(IIf([{Field}] Between 25 And 49, 1, 0)) AS From25to49,
  Sum(IIf([{Field}] >= 50, 1, 0)) AS Over50
FROM {Table};
```

**Rank** (Cookbook 10.14) — precompute via a scratch table ordered descending, or compute per-row rank inline:

```sql
SELECT A.{Key}, A.{Field},
  (SELECT Count(*) FROM {Table} B WHERE B.{Field} > A.{Field}) + 1 AS Rank
FROM {Table} A
ORDER BY Rank;
```

**Percentile slice** (Cookbook 10.13): `SELECT TOP 10 PERCENT {Field} FROM {Table} ORDER BY {Field} DESC;`

**Growth rate between periods** (Cookbook 10.9):

```sql
SELECT Year({DateField}) AS Yr, Avg({ValueField}) AS AvgThisYear,
  (SELECT Avg({ValueField}) FROM {Table} b WHERE Year(b.{DateField}) = Year(a.{DateField}) - 1) AS AvgLastYear,
  (AvgThisYear - AvgLastYear) / AvgLastYear * 100 AS GrowthPct
FROM {Table} a
GROUP BY Year({DateField})
ORDER BY Yr;
```

## 6. Time-Series & Trend Analysis

Key date/time functions: `DateDiff("interval", d1, d2)`, `DateAdd("interval", n, d)`, `DatePart("interval", d)`, `Year`, `Month`, `Day`, `Weekday`, `Hour`, `Minute`, `Second`. Interval codes: `yyyy`, `q`, `m`, `y`/`d` (day), `ww`/`w` (week), `h`, `n` (minute — `m` is month), `s`.

**Elapsed time between two columns:**

```sql
SELECT {Key}, DateDiff("n", StartTime, StopTime) AS ElapsedMinutes FROM {Table};
```

**Elapsed business days (exclude a Holidays/Weekends table)** (Cookbook 8.2):

```sql
SELECT {Key}, DateDiff("d", StartDate, StopDate)
  - (SELECT Count(*) FROM Holidays WHERE Holiday Between StartDate And StopDate)
  - (SELECT Count(*) FROM Weekends WHERE Weekend Between StartDate And StopDate) AS BusinessDays
FROM {Table};
```

**Moving average** (Cookbook 9.2) — correlated subquery over a trailing window:

```sql
SELECT A.{DateField}, A.{ValueField},
  (SELECT Avg(B.{ValueField}) FROM {Table} B
     WHERE B.{DateField} Between DateAdd("d", -7, A.{DateField}) And A.{DateField}) AS MovingAvg7
FROM {Table} A
ORDER BY A.{DateField};
```

**Breakdown by calendar part** — group by `Year({DateField})`, `Month({DateField})`, or `DatePart("ww", {DateField})` to build month-over-month / week-over-week summaries.

## 7. Presenting & Exporting Results

AccessPOSH/access-mcp automate SQL and structure, but **not** PivotTables/PivotCharts (those are Access/Excel UI features). To hand off analysis for visualization:

```powershell
# Run the analysis, get JSON, summarize inline in the response
Invoke-AccessSQL -DbPath $db -SQL "SELECT ..." -AsJson

# Or export the source table/query for the user to chart in Excel
Export-AccessToExcel -DbPath $db -ObjectName "{QueryOrTable}" -ExcelPath "C:\out.xlsx" -HasFieldNames -AsJson
```

If a scratch query will be reused, save it with `Set-AccessQuery` instead of re-typing SQL each time; name it clearly (e.g., `qryAnalysis_MonthlyGrowth`) and tell the user it was added so it can be cleaned up later.

## 8. Reporting Findings

When presenting results to the user, always state:
- **The exact SQL used** (so it's reproducible/auditable)
- **Row count of the result set** (sanity check)
- **Caveats** — e.g., nulls excluded, sample vs. population stdev, date range assumed

## Non-Negotiables

- Treat all SQL here as read-only analysis; never mutate live data without explicit user confirmation and `-ConfirmDestructive`.
- Always bracket identifiers with spaces or reserved words: `[Order Date]`, `[Date]`.
- Verify a query runs (`Invoke-AccessSQL ... -AsJson`) and inspect the JSON before summarizing results — never fabricate numbers.
- Call `Close-AccessDatabase` when the analysis session is done to release the COM lock.
- If the user's ask is ambiguous (which table? which date range? sample or population variance?), ask before running long queries against large tables.

**Step 3: Implementation**
- Follow task list sequentially
- Check off tasks as completed
- Run tests to verify before marking "done"
