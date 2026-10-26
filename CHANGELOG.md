# Changelog

All notable changes to Modeleon will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.2.3] — 2026-10-26

### Added

- **A named view, applied.** A model may keep several looks beside its
  own — `model.board = mo.ExcelView('Board pack', font=…, bands=…)`,
  `model.banker = mo.ExcelView('Banker', …)` — and `with
  mo.applied_view(model.board): model.to_excel(path)` draws the book
  under one of them: its fonts, bands, columns and formats over the
  model's own view, whatever it leaves unset coming, as before, from the
  model, a `mo.house_view(...)` around the call and the engine's defaults
  (`mo.default_excel_view`); a `sheet={'order': …}` it declares orders
  the book. The same model, the same numbers and formulas, in as many
  looks as it declares. The period axis stays the model's: `orient` is
  geometry, not a look.
- **Errors caught, and many branches without nesting.** `mo.IFERROR(value,
  fallback)` shows the fallback where a division by zero or a lookup that
  finds nothing would spread an error through the model (`=IFERROR(B7 / B8,
  0)` in every period); `mo.ISERROR` and `mo.ISNUMBER` test a cell, `mo.XOR`
  joins conditions, and `mo.IFS(cond, value, …)` / `mo.SWITCH(expr, case,
  value, …, default)` choose among many branches without nested `IF`s. Each
  is one call per period over a row and runs inside a loop container;
  `XOR`, `IFS` and `SWITCH` are stored as the `_xlfn.` names Excel expects.
  `SWITCH` compares numbers to about 15 significant digits, as the
  spreadsheet does; text given to `XOR` is refused with a pointer to compare
  it instead. The values `IFERROR`, `IFS` and `SWITCH` may return are in
  one unit, inside a loop container too (`mo.IFERROR(usd, eur)` is
  refused; a plain `0` or a line without a unit goes with any), and
  `SWITCH` compares its expression with each case in one unit
  (`mo.SWITCH(hours, limit_usd, …)` is refused).
- **Rounding, powers and logarithms.** `mo.ROUNDUP`, `mo.ROUNDDOWN`,
  `mo.TRUNC`, `mo.CEILING`, `mo.FLOOR`, `mo.POWER`, `mo.SQRT`, `mo.EXP`,
  `mo.LN`, `mo.LOG`, `mo.LOG10` and `mo.SIGN`, one call per period, and
  `mo.PRODUCT`: over one row the product of its cells, written over its
  range (`=PRODUCT(C3:N3)`), over two or more operands with a series among
  them one product per period (`=PRODUCT(C3, C4)` in each column), as `a *
  b` multiplies them, and over scalars only one number. They round the
  number as the cell shows it (`ROUNDUP(0.1 * 3, 1)` is 0.3, not 0.4), read
  a date as its serial number, and give `#NUM!` where no cell could hold
  the result; the same functions work inside a `mo.recurrence` template
  (`"ROUNDUP({prev} * 1.05, 0)"`). A rounded amount keeps its unit;
  `mo.POWER` and `mo.SQRT` raise it as `**` does (`mo.POWER(usd, 2)` is in
  `USD^2`, `mo.SQRT` of a line in `USD^2` is in `USD`, `mo.SQRT(usd)` has
  none); `mo.PRODUCT` multiplies its operands' units (`USD/t` times `t` is
  `USD`), and over one row runs along pure numbers such as growth factors:
  a row in a unit is refused, its cells' product being that unit to the
  power of their count; `mo.CEILING` and `mo.FLOOR` take a number and its
  multiple in one unit; `mo.EXP`, `mo.LN`, `mo.LOG`, `mo.LOG10` and
  `mo.SIGN` carry none. The same holds inside a loop container.
- **Loan schedules and depreciation.** `mo.NPER`, `mo.RATE`, `mo.IPMT`,
  `mo.PPMT`, `mo.CUMIPMT` and `mo.CUMPRINC` — how long a loan runs, what
  rate it carries, and the interest and principal of each payment — and
  `mo.SLN`, `mo.DB`, `mo.DDB` and `mo.SYD`. One call per period, so
  `mo.PPMT(rate, mo.time.index, n, loan)` is a whole amortization row
  (`=PPMT(I!B1, 1, 4, I!B2)`, `=PPMT(I!B1, 2, 4, I!B2)`, …), inside a loop
  container too. `RATE` iterates as the spreadsheet does, retrying from its
  default guesses, and is `#NUM!` where no rate balances the flows. The
  amounts each one reads — a loan and its balloon, a payment, a cost and
  its salvage — are in one unit, the result's (`mo.SLN(cost, salvage,
  life)` with both in USD is in USD; a salvage in EUR is refused), and
  `NPER` and `RATE` carry none.
- **Date arithmetic.** `mo.DATEDIF(start, end, unit)` with all six units,
  Excel's quirks included (`"MD"` from 31 January to 1 March is -2 in a
  common year); `mo.WEEKDAY` and `mo.WEEKNUM` with every return type (21 is
  the ISO week); `mo.NETWORKDAYS` and `mo.WORKDAY`, whose holidays line is
  written as the range every period's call reads (`=NETWORKDAYS(B2, C2,
  Inputs!B5:D5)`).
- **Text.** `mo.LEFT`, `mo.RIGHT`, `mo.MID`, `mo.TRIM`, `mo.FIND`,
  `mo.SEARCH`, `mo.SUBSTITUTE`, `mo.VALUE`, `mo.TEXT` (number and date
  formats: `TEXT(d, "yyyy-mm-dd")`, `TEXT(x, "#,##0.00")`), `mo.REPT` and
  `mo.TEXTJOIN`. An operand reads as the cell would: a number as its General
  text, a date as its serial number, a blank as empty text.
- **Conditional sums and counts.** `mo.SUMIF`, `mo.SUMIFS`, `mo.COUNTIF`,
  `mo.COUNTIFS`, `mo.AVERAGEIF`, `mo.AVERAGEIFS`, `mo.MAXIFS` and
  `mo.MINIFS` over a row, or over a field down the records of a table
  (`=SUMIFS(Deals!C3:C5, Deals!B3:B5, "North")`), with Excel's criteria —
  `">=10"`, `"<>0"`, `"North*"`, `("<", threshold)` joined to a cell as
  `"<" & B2` — and `mo.COUNT` / `mo.COUNTA`: over one row the count of its
  cells (`=COUNT(C5:N5)`), over two or more operands with a series among
  them one count per period (`=COUNT(C5, C9)` in each column), over scalars
  only one number. The lines a `SUMIF`, `SUMIFS`, `AVERAGEIF`,
  `AVERAGEIFS`, `MAXIFS` or `MINIFS` call totals, averages or picks from
  are in one unit, which the result carries, and each criterion is compared
  in its range's unit, one joined to a cell too: `mo.SUMIF(hours, ("<",
  cap), revenue)` with the cap in USD is refused, inside a loop container
  as outside, and an error the joined cell holds is the call's result. A
  criterion a spreadsheet reads one way in some locales and another way in
  others (`'1,000'`, `'1/31/2026'`) is refused when the model is built,
  with the form to write instead (`'1000'`, `'2026-01-31'`). Where the
  cells a conditional call reads cannot form one range — lines on two
  sheets or with a line between them, or a row split by a totals column —
  the call is written as its computed number and `check_compat` lists the
  cell; `mo.COUNT` and `mo.COUNTA` name such cells one by one
  (`=COUNT(East!B1, East!B3)`).
- **Lookups.** `mo.XLOOKUP`, `mo.INDEX`, `mo.MATCH`, `mo.VLOOKUP` and
  `mo.HLOOKUP` over a row, a list of rows, or the records of a table; a key
  that is a series gives one lookup per period (`=INDEX(Inputs!B3:D3,
  MATCH(B2, Inputs!B2:D2, 0))`, then `MATCH(C2, …)`), so a year-varying
  assumption is read from its table instead of being retyped. Approximate
  matches need sorted cells, as in Excel; `XLOOKUP`'s binary searches
  (`search_mode` 2 / -2) are refused, and `None` typed as a key is refused
  (the formula would write it as 0). A key is compared with the cells
  searched in one unit (`mo.XLOOKUP(amount_usd, thresholds_eur, fees)` is
  refused), and the cells a lookup may return — a list of rows, a block's
  column or row, an `if_not_found` — are in one unit, as `mo.CHOOSE`'s
  choices are; a number or a line without a unit goes with any. Where the
  cells searched or returned cannot form one range — lines on two sheets
  or with a line between them, or a row split by a totals column — the
  lookup is written as its computed number and `check_compat` lists the
  cell.
- **Statistics.** `mo.MEDIAN`, `mo.STDEV_S`, `mo.STDEV_P`, `mo.VAR_S`,
  `mo.VAR_P`, `mo.PERCENTILE_INC`, `mo.PERCENTILE_EXC`, `mo.QUARTILE_INC`,
  `mo.QUARTILE_EXC`, `mo.LARGE`, `mo.SMALL`, `mo.RANK_EQ` and `mo.RANK_AVG`
  — Excel's dot is an underscore in the name, and the file stores
  `_xlfn.PERCENTILE.INC` — with the older `mo.STDEV`, `mo.STDEVP`,
  `mo.VAR`, `mo.VARP`, `mo.PERCENTILE`, `mo.QUARTILE` and `mo.RANK` that
  older workbooks carry. Over a row or one field down the records of
  a table: the median of a comps multiple is `=MEDIAN(Comps!C3:C9)`, its
  first quartile `=_xlfn.PERCENTILE.INC(Comps!C3:C9, 0.25)`, and each
  record's rank among its peers one call per record against the same
  range (`=_xlfn.RANK.EQ(C3, C3:C9)`). Several series are combined period
  by period, as `a + b` combines them — the rule every aggregate follows:
  `mo.MEDIAN(plan, actual)` is `=MEDIAN(C5, C9)` in every column,
  with scalars broadcast, series of other lengths refused and a line with
  tracks taken track by track; a list of two or more yearly lines (one
  field down a records table, or sibling lines) is one sample per year —
  the peers' median revenue is `=MEDIAN(Peers!C3:C9)`, then
  `=MEDIAN(Peers!D3:D9)`, …, their first quartile and each record's rank
  among them one call per year over the same column of cells
  (`=_xlfn.RANK.EQ(D3, D3:D9)`). One series alone is reduced over its
  periods (`=MEDIAN(C5:N5)`) and scalars only stay one number; a list of
  one yearly line is refused, as no sample. Inside a `loop=True` container
  a call over two or more values runs one period at a time. Blanks, text
  and booleans in a range are skipped; a `k` or a quart out of range is
  `#NUM!`; `LARGE` takes a fractional `k` up and `SMALL` down; the values
  a call reads are in one unit.
- **A rank, a conditional total or a lookup per record reads its range
  once.** Every call over the same lines — `mo.RANK_EQ(rec.multiple,
  peers)` for each record of a table, or a `SUMIF`, `COUNTIFS`, `MATCH` or
  `XLOOKUP` per record — shares one range, read once and checked before
  each call: a line given another value (replaced, or its list changed in
  place), another unit or another place is read again. Its cells are
  sorted once for `RANK`, `LARGE`, `PERCENTILE` and `MEDIAN`, and indexed
  by value once a second equality in `SUMIF`, `COUNTIFS`, `MATCH`,
  `XLOOKUP` or their kin looks among them; the range's address is worked
  out once per sheet. Each formula still lists all its peers, so the build
  of a very large table keeps a cost per record and peer.
- **Hours.** `default_grain='hour', default_start='2026-01-05T13'` (or a
  `datetime` on the hour, or a date — its midnight): the wall clock of one
  fixed-offset zone, declared once on the window (`default_zone='+01:00'`;
  UTC when omitted; a named zone such as `'Europe/Paris'` is refused with the
  offset to write, since a zone that moves its clocks would skip and repeat
  hours). An hour is `[13:00, 14:00)`: `mo.time.end` is where it stops,
  `mo.time.days` is 1/24 (`mo.time.months` is refused), `regrain` weights a
  mean by it. An hour nests in its day and everything above: `.at('day')`,
  `.at('month')`, `.at('week')`; `timeline={'totals': ['day']}`
  draws a day total after each day's last hour, `05 Jan 2026 | 22:00 23:00 |
  Total 05 Jan 2026`, a live `=SUM` over its hours; a time row's day total
  reads its hours' cells too — the first hour's start, the last hour's
  stop, the `SUM` of their days. The header is a closed
  form off the start cell (`start + i/24`, never a chain that would drift
  over thousands of columns), captions are the hour's start as a clock
  (`dd mmm yyyy hh:mm`), the days row shows its fraction. Keys are
  `YYYY-MM-DDTHH`; the finance spelling is `05 Jan 2026 13:00`.
- **A window wider than a sheet is refused.** A sheet holds 16,384 columns
  and a period is one (a total another): `default_periods` past 16,383 is
  refused where it is declared, and the layout refuses a sheet whose label
  columns, periods and totals together pass the last column
  (`SheetWidthError`), naming the count, what the grain fits (hours about 22
  months) and the two ways out — a coarser grain, or the data kept as rows.
- **An instant keeps its clock in a formula.** A `datetime` off midnight
  inlined into a formula is written `(DATE(2026, 1, 5)+TIME(13, 0, 0))` — in
  brackets, so `end - start` subtracts the whole instant; it was cut to its
  day.
- **An ISO year in its own book.** A window of ISO years
  (`default_grain='iso_year'`), or a weekly book viewed as ISO years,
  writes its header off the start cell: the ISO year of the start cell
  is `YEAR(start + 3)`, each year starts on `DATE(y,1,4) -
  WEEKDAY(DATE(y,1,4),3)` and ends the day before the next one does; an
  anchor chains `start = previous end + 1` and ends each year off its own
  start, so a first year opened on a Monday inside it is short and right.
- **ISO weeks, and the ISO year that totals them.** `default_grain='week'`,
  `default_start='2026-W05'` (or a date that is a Monday): weeks run Monday to
  Sunday and are numbered in their ISO year — `2026-W01` starts on 29 December
  2025, and 2026 has 53 weeks; `mo.time` reads each week (7 days; `months` is
  refused, as a day's is). A week lies inside one ISO year and in no month,
  quarter or calendar year: `.at('iso_year')` rolls it up, and
  `timeline={'totals': ['iso_year']}` draws a total after each ISO year's last
  week; a day sheet may total by `'week'` (with `'iso_year'` above it). A
  weekly header is a live chain off the start cell (`end = start + 6`).
- **A day sheet totals by month.** `ExcelView(timeline={'totals': ['month']})`
  (beside `'quarter'`, `'half'`, `'year'`) on a day window draws a month total
  after each month's last day — `30 31 | Total Jan | 01 02 03 | Total Feb` —
  under a month tier (`Jan`; the year on the tier above when it is declared),
  the days grouped under their month in the outline, each total a live `=SUM`
  over its days (`=` the last day for a stock). Code that reads the totals
  plan (`modeleon.compile.excel.totals.TotalsPlan`) finds a sheet's own periods
  keyed `('period', i)`, where they were `('month', i)` whatever the grain;
  `month_col` stays as an alias of `native_col`.
- **The half-year.** A fifth grain between quarter and year: a window of
  its own (`default_grain='half', default_start='2026-H1'`, or a date —
  a half opened on the 1st of any month is a short first period, as a
  quarter's is), a totals column (`ExcelView(timeline={'totals': ['half']})`
  draws `Total H1` after June and `Total H2` after December, under the
  year's tier and over the quarters'), a target of `.at('half')` on a line
  or on a whole model (`model.at('half')`), with its `mo.time` (181 / 184
  days, 6 months) and its header formulas (`EOMONTH(start, 5)`). Keys are
  `YYYY-Hn`; the finance spelling is `H1 2026`, the compact one `H1-26`.
- **The calendar's answers for a window.**
  `modeleon.core.time.period_bounds('2026-01', 'month', 3)` returns each
  period's first and last day as two lists (1 to 31 January, 1 to 28
  February, 1 to 31 March; for an hour, where it starts and stops), and
  `year_back_index('2026-01', 'month', 15)` the index of the period a year
  earlier in the same window (`None` for the first twelve months, then `0,
  1, 2`). `modeleon.core.calendar` holds a ruler for each grain (hour, day,
  week, month, quarter, half, year, iso_year): `GREGORIAN.ruler('quarter')`
  finds the period holding a day (`start_of`) and where the next one starts
  (`next_start`), writes and reads a period (`label(date(2026, 4, 1),
  'finance')` is `'Q2 2026'`, `parse('2026-Q2')` gives back 1 April) and
  names its grain (`word` is `'Quarterly'`, `letter` `'Q'`). A chart's time
  axis or a growth on last year in your own code can take its dates from
  these instead of doing calendar arithmetic of its own.

### Changed

- **`mo.SUM` and `mo.AVERAGE` over two or more operands work period by
  period.** `mo.SUM(product_a, product_b)` is the two lines' total in every
  period — a subtotal, written `=SUM(C5, C9)` in each column — and
  `mo.AVERAGE(plan, actual)` their mean, with scalars broadcast as `a + b`
  broadcasts them, series of other lengths refused and an operand with
  tracks taken track by track; inside a `loop=True` container they are
  allowed. Before, two series were reduced to one number across both rows
  (`=SUM(B2:F2, B3:F3)`). Over one operand `mo.SUM(row)` stays the total of
  its periods (`=SUM(B2:F2)`), and over scalars only one number (`=SUM(B2,
  B3)`). The total of several rows over all their periods is now
  `mo.SUM(a) + mo.SUM(b)`. Every aggregate follows the rule `mo.MIN` /
  `mo.MAX` follow since 0.2.2: several series are combined period by
  period, one series alone is reduced over its periods, scalars only stay
  one number. The operands of a call are added left to right, as `+` adds
  them (Python's `sum` compensates the rounding), so `mo.SUM(a, b, c)` is
  `a + b + c` to the last digit — over scalars only too, where a value can
  move in its last digit (`mo.SUM(0.1, 0.2, -0.3)`). Lines on other time
  windows and lines along keys are read as `a + b` reads them, as in every
  function (the next entry). A `None` typed into `mo.SUM`,
  `mo.AVERAGE`, `mo.MIN`, `mo.MAX` or `mo.PRODUCT` (or into a list typed
  in) is the 0 the formula holds, as the sheet reads `=MIN(C2, 0)` and as
  the counts read it (`mo.AVERAGE`, `mo.MIN` and `mo.MAX` skipped it);
  text among the values is refused with a message naming
  the function and the value. A per-period `SUM` or `AVERAGE` re-grains
  as `a + b` does — its quarter is the call over its operands' quarters,
  written live in a totals column (`=SUM(E5, E9)`) — and a per-period
  `mo.PRODUCT` as `a * b` does.
- **Every function reads a line on another time window by date, and lines
  along keys key by key — as `a + b` reads them, in Python and in the
  workbook.** A located series on another window (a sub-container's own
  `default_start`, a line's `start=`) is aligned onto the combined window:
  each line continued by its own `extend=` rule — `mo.zero()` for a flow,
  `mo.hold()` for a rate — the result on that window, two grains refused.
  In 0.2.2 `mo.AND`, `mo.OR`, `mo.CHOOSE`, `mo.CONCAT`, `mo.DATE`,
  `mo.DAYS360`, `mo.YEARFRAC` and `mo.DAYS` paired such lines by position —
  January beside a tranche's March, the line's last value standing in for
  the months it lacks — `mo.IF` paired them by position and raised
  `IndexError` past the shorter branch (`mo.IF(rev > 2, rev, fee)` gave
  `[10, 20, 3, 4]` for a fee from March; with the fee's
  `extend=mo.zero()` it is now `[0, 0, 3, 4]`), `mo.MIN` and
  `mo.MAX` refused lines of other lengths (`Cannot broadcast lists of
  different lengths`), and `mo.MOD` and the comparisons paired two windows
  of one length by position and refused others; now each reads the line
  on its date, as do the functions added since (rounding, powers, loan
  schedules, `mo.IFERROR` and its kin, the lookups' keys and positions,
  the statistics, the counts), a `mo.recurrence` template's lines, which
  0.2.2 paired by position, and a loop container, whose rows read such a line by date in every period (a
  one-period line was broadcast into every period of the loop, a line of
  another length refused). A model that relied on the positional reading
  meets a refusal naming the line and the rule to declare: `'fee' does
  not cover the combined window (…). Say what continues it:
  extend=mo.zero() (a flow) or extend=mo.hold() (a rate)`. An unnamed
  formula runs on the window of the lines it reads — `rev * 1` on the
  model's, `fee > 15` on the tranche's — and meets other windows by date;
  it carries no `extend=` rule, so where it does not cover the combined
  window it is refused, naming the line to make of it — and it is placed
  the same whatever the order the model was built in: a line attached to
  the model after a formula was built over it, or a window moved after,
  places the formula as a fresh one is placed. The commonest case
  is a condition on a line of another window, `mo.IF(fee > 15, rev, 0)`,
  which 0.2.2 read by position: its migration is a line with a rule on
  that container, `m.t.big = mo.Variable(m.t.fee > 15, extend=mo.zero())`,
  then `mo.IF(m.t.big, rev, 0)`; a condition on the model's own lines,
  `mo.IF(rev > 2, fee, 0)`, needs nothing. A one-value line on a window
  of several periods is the constant it always was; a line along keys
  runs on no window; a plain line beside a tracked one is read on the
  tracked line's window (`mo.SUM(rev, once, trk)` reads `once` on its one
  date — 0.2.2 refused the call, and `trk + once` read it in every
  period) and a tracked line on its own window only; windows are ordered
  by date, never as text (`'2026-9'` comes before `'2026-10'`: `rev + fee`
  dropped September). A line of dates or text with `extend=mo.zero()` is
  refused where the zero would stand in, pointing to `mo.hold()` (0.2.2
  let the zero stand in without a word). A line of another length with no
  window at all is refused, as `a + b` refuses it; `mo.IF`, `mo.AND`,
  `mo.OR`, `mo.CONCAT`, the dates and a `mo.recurrence` template, which
  paired it by position, now name it (`'pair' cannot be paired with the
  call's other series period by period`); a list typed in keeps its positional reading,
  its last value standing in for the periods it lacks as the cell writes
  it (in `mo.IF` too, which raised `IndexError` there), and `mo.IF` with
  one condition takes the chosen branch whole, as before.
  Lines along the same keys in another order are combined key by key in
  every function, the result keeping their keys, and a line along keys
  against a period series is refused, as `a + b` refuses it — the
  functions zipped them cell by cell. What a workbook written now does
  differently from 0.2.2: where Python aligns two windows — `+`, `-`, `*`,
  `/`, `%`, `**`, the comparisons, every function and a loop container's
  rows — a formula reading a line on a sub-container's own window inside
  the sheet's takes the line's cell of the same date and, for a period
  the line has no column for, its `extend=` rule, written as `0` for a
  flow (`=SUM(C5, 0)`) or as its first or last cell for a rate held;
  0.2.2 wrote the cell in the same position — January read the line's
  first cell, March's — and, past the line's last column, its first cell
  again. Where Python reads one line alone — `fee * 2` over a line on
  another window, set on the model's sheet — the workbook reads it by
  position, as Python does and as 0.2.2 did. A formula whose operands met
  by date on a window of its own, set on a sheet of another window
  (`m.x.y = mo.IFERROR(m.b.bz + m.c.cz, 0)` with `B` from November on a
  sheet from January), is refused by the export as the bare sum was: its
  January column would hold November. A line along keys in another order
  is read at its key's cell (`=B6 * D8`, Bear × Bear), where 0.2.2 wrote
  the cell in the same position (Bear × Bull) beside Python's key-by-key
  value; a function's result over one keyed line keeps its keys and is
  laid out under a caption row of them (`mo.ROUND(growth, 1)`); where the
  line above it runs along periods or other keys that caption row is new
  and the rows below move down one — 0.2.2 dropped the keys and wrote the
  values under the period captions (or refused the line when the keys and
  the periods differed in number). A window
  reaching beyond the sheet's is refused at export, as before.
- **Functions that add, compare or pick take their values in one unit, as
  `+` does.** `mo.SUM(revenue, hours)`, with `revenue` in USD and `hours`
  in hours, is refused, naming both lines (`mo.SUM takes its values in one
  unit; 'revenue' and 'hours' are in USD and hours`), with the advice `+`
  gives; so are `mo.AVERAGE`, `mo.MIN`, `mo.MAX`, the two branches of
  `mo.IF`, the choices of `mo.CHOOSE` and a comparison of two lines
  (`revenue > hours`: `Cannot compare values with incompatible units: USD
  and hours`). Before, these ran in silence and the result had no unit, so
  the next formula could not catch the mix: `mo.SUM(revenue, hours) + cost`
  passed, in USD. A number or a line without a unit goes with any
  (`mo.MAX(revenue, 0)` is in USD). A condition, an index or a case is not
  one of the values: `mo.CHOOSE(scenario, revenue, cost)` with `scenario`
  labelled `#` is in USD, where it had no unit. Inside a loop container,
  over a row read before its line, the check is made when the loop closes.
  The functions this release adds take their values the same way (above).
- **A series where a function took one value is one call per period.**
  `mo.ROUND`'s digits, `mo.EDATE` / `mo.EOMONTH`'s months and every
  argument of `mo.PMT`, `mo.FV` and `mo.PV` may be a line — each period
  computes the call over its own values, written in its column as the
  spreadsheet writes it (`=ROUND(C5, C9)`, `=EDATE(C5, C9)`, `=PMT(C9 /
  100, 12, 1000, 0, 0)`), a scalar standing in every period and lines read
  as above; these raised a bare `TypeError`. `mo.NPV`'s rate and `mo.IRR`
  / `mo.XIRR`'s guess stay one value for the whole row: a series there is
  refused by name, as `mo.XNPV`'s rate already was.
- **A function says how it is written.** Which names the Excel writer
  passes through as `=NAME(...)`, which arguments it writes as ranges, which
  names the file stores with the `_xlfn.` prefix, and which functions a loop
  container runs one period at a time, is one registry
  (`modeleon.functions.registry`) each function fills from its own module;
  the renderer's and the loop container's hand-kept name sets are gone. A
  function name the registry does not know is still written as given, with
  a warning that points at it.
- **A condition's `source_code` follows its formula.** For `mo.IF`,
  `mo.AND`, `mo.OR`, `mo.NOT`, `mo.ISBLANK` and `mo.CHOOSE`, and the new
  `mo.IFERROR`, `mo.ISERROR`, `mo.ISNUMBER`, `mo.XOR`, `mo.IFS` and
  `mo.SWITCH`, `source_code` is the line's `.formula`, rendered when it is
  read, so it follows names given to the lines later and a formula set by
  hand (`flag.formula = '...'`): `mo.IF(a > 0, b, 0.0)` built before `a`
  and `b` were named reads `IF((a > 0), b, 0.0)` once they are. Before, it
  kept the text of the call, `IF((v1088264b0 > 0), v107cf6b70, 0.0)`. A
  text written with `flag.source_code = '...'` still wins; nothing in the
  workbook changes.
- **A recurrence template whose power or function outgrows a cell gives
  `#NUM!`.** `mo.recurrence(1e200, "{prev} ** 2", periods=3)` is
  `[1e+200, '#NUM!', '#NUM!']`, carried on as a division by zero is; 0.2.2
  raised `OverflowError`. The functions new in this release do the same in
  a template (`"EXP({prev} * 1000)"`); an overflow by `*` or `+` still
  gives `inf`, as in Python.
- **A recurrence is in its start's unit.** `mo.recurrence(opening,
  "{prev} * (1 + {g})", g=growth)` from an `opening` in USD is in USD, and
  so is `mo.recurrence_sum(opening, additions)` over USD additions: a
  revenue line built so is labelled `Revenue (USD)` in the workbook and
  the notebook repr, and adding an hours line to it is refused. Before,
  both had no unit unless `unit=` was given: the label read `Revenue`, and
  the sum with hours passed, in hours. A declared `unit=` still decides.
- **Grains nest by a relation, not by their order.** The calendar says
  which grain lies inside which: in `modeleon.core.calendar`,
  `GREGORIAN.nests('day', 'week')` is true and `GREGORIAN.nests('week',
  'month')` false, and `GREGORIAN.verify()` checks each such claim period
  by period. `modeleon.core.time.GRAINS` lists `hour, day, week, month,
  quarter, half, year, iso_year` in that order (ranks 0 to 7). `.at()` on a line or on a whole
  model, and arithmetic between two lines, are refused across grains that
  do not nest (a weekly line's `.at('month')`), with the reason and the way
  out: a grain both share (days). A totals column a sheet cannot draw (a
  month column over weeks, weeks beside months over days) is left out with
  a `TotalsWarning` that names the sheet and says why (`'month' totals are
  not drawn: a week does not lie inside one month (weeks straddle months).
  A week sheet totals by iso_year.`). In 0.2.2 `GRAINS` was `day, month,
  quarter, year`, and a line rolled up to any grain after its own.
- **Large models build faster.** Arithmetic between lines does less work
  for each value, a condition (`mo.IF` and its kin) renders its text only
  when it is read, and under a blend of tracks (`mo.blend`) what a line
  reads is read once for the model, not again for every line built on it:
  in our runs a model of 2,000 lines of arithmetic or conditions builds in
  about three quarters of the time 0.2.2 took. Values, formulas and
  workbooks are the same; writing the workbook takes about as long as
  before.

### Fixed

- **A tracked row reads by position.** `closing[-1]` of a plan/actual row
  raised "Index -1 on a scalar Variable", and `[0]` copied the whole row as
  numbers. Both now read that period of every track, written in the book as
  one live cell per track; a period not entered yet reads 0, as a blank cell
  does in a formula. An index past `[0]` on a one-cell row says to read the
  row by its name instead of pointing at `var.value`.
- **A view goes where it is read.** An unnamed `mo.ExcelView(...)` hung
  under a name that says "view" — `model.excel_view = …`, `model.view = …`,
  `mo.Model(..., view=…)` — became a container nobody reads: the workbook
  kept its look and no total column appeared, or the view's own settings
  were refused as a series (`'totals' has 2 value(s) in a 24-period
  window`). `to_excel` now warns (`UnreadViewWarning`, on the author's
  line) with the slot to use, spelled by the author's own name for the
  model — `model.default_excel_view = mo.ExcelView(...)`. A view the slot
  points at is read, also through `extend`, through the copy a projection
  or a second parent makes, through an ancestor's slot, and under
  `mo.house_view(...)`. A view given a name of its own
  (`model.board_view = mo.ExcelView('Board pack', ...)`, kept to be applied
  with `mo.applied_view`) and a view kept under a name without "view" in it
  draw no warning, and no view's settings are read as a series over the
  model's periods.
- **A horizon given by assignment counts.** `m.default_periods = 36` marks
  the model last given a horizon, as the constructor does, so a line
  reading `mo.time` outside the model's `with` block finds the window; a
  refused change (the lines already read the window) marks nothing.
- **A container cannot hold itself.** `model.m = model` adopted the model
  as its own component and recursed until Python gave up; so did assigning
  an ancestor. Refused with the alias to write (`with model as m:`). A read
  of a component that was never assigned (`'Model' has no component
  'inputs'`) now says how to assign it, and that a row read before its own
  line belongs in a loop container. Both messages spell the container by
  the author's own name for it (`model`, `income`, `model.pnl`), never by
  its declared or display name, so the line they suggest runs.
- **A large table is written in linear time.** A container with thousands
  of children copied its whole list of children for every child it read,
  and checked a new child's name against the list one by one: a records
  table of 10,000 records took two minutes to write and now takes about
  three seconds (3,000 records: 11 s to under a second), and its notebook
  repr and `.code` follow.
- **A blank typed into `LEN`, `UPPER`, `LOWER` or `CONCAT` is written as
  `""`.** It was written as `0`, which the workbook reads as the text "0"
  (`=LEN(0)` is 1 where Python gave 0).
- **`mo.DAYS360` reads a method that varies by period.** A line given as
  `method` was read as one TRUE — any list is — so every period counted
  the European way, where the workbook reads each period's own cell
  (`=DAYS360(B4, B11, B10)`); now each period takes its own method.
- **A list of lines given as one argument says what it means, or is
  refused.** In 0.2.2 `mo.SUM([a, b])` held a Variable as its value and the
  cell its text (`Variable(...)`), as did `mo.AVERAGE`, `mo.MIN` and
  `mo.MAX` over scalar lines; over a list holding a series `mo.MIN` /
  `mo.MAX` raised "The truth value of a list Variable is ambiguous". A
  list of scalar lines is its lines one by one in `mo.SUM`, `mo.AVERAGE`,
  `mo.MIN`, `mo.MAX` and `mo.PRODUCT` (`mo.SUM([p, q])` is `mo.SUM(p, q)`,
  `=SUM(B1, B2)`) and the range they form in `mo.COUNT` / `mo.COUNTA`
  (`=COUNT(B1:B2)`). A list holding a series reads two ways and is
  refused, both spelled with the lines' own names: per period, `mo.SUM(a,
  b)`; the total over all their periods, `mo.SUM(a) + mo.SUM(b)` (each
  function its own: `mo.MIN(mo.MIN(a), mo.MIN(b))`, …). The statistics
  read a list of series lines as one sample per period (above).
- **A grain or totals kind the engine does not know is refused where it is
  given.** `m.default_grain = 'fortnight'` is refused on assignment, as the
  constructor refuses it (`Unknown default_grain 'fortnight'; supported:
  hour, day, week, ...`); it used to be accepted and fail later, in
  `to_excel` or at the first `mo.time` read, on `Unknown time grain
  'fortnight'`, which does not name `default_grain`.
  `ExcelView(timeline={'totals': ['fortnight']})` is refused by name
  (`ExcelView timeline totals may name day, week, month, quarter, half,
  year, iso_year; got 'fortnight'`); it used to be accepted and draw no
  column, without a word. `.at('fortnight')` on a month line names the
  grains it rolls up to (`a month rolls up to quarter, half, year`), where
  it raised a bare `KeyError: 'fortnight'`. A day window's start must be
  `YYYY-MM-DD`: `default_start='2026-W01'` is refused, where it used to
  start the window on Monday 29 December 2025, the first day of that ISO
  week.
- **Months beside years point to the explicit step.** Two units that
  differ only in a month word against a year word (`months` / `yrs`,
  `USD/month` / `USD/yr`) are refused with the step written out both ways
  — `(x / 12).set_unit('yrs')`, `(x * 12).set_unit('months')` — since a
  unit is a label that never rescales: dividing by 12 leaves months in
  months. A refusal with no hint no longer opens its advice with
  "Otherwise".
- **A remainder keeps its unit.** `usd % 12` is in USD, and `usd % hours`
  and `mo.MOD(usd, hours)` are refused, inside a loop container too;
  before, `%` dropped the unit, and both let any two units through,
  leaving none. `100 // usd` is in `1/USD`, as `100 / usd` is.
- **A whole power raises the unit, inside a loop container too.**
  `usd ** 2.0` is in `USD^2`, as `usd ** 2` is, and `(usd * usd) ** 0.5`
  is in `USD`; a power that comes out fractional (`usd ** 0.5`), or an
  exponent that is a line, still leaves none. A unit string reads `^n` as
  a power, so a line declared in `'USD^2'` adds to `usd * usd`. Inside a
  loop container a row takes the unit it takes outside one: `balance ** 2`
  is in `USD^2`, and `balance + balance ** 2` is refused. Before, a power
  written as a float and a root had no unit, so `usd ** 2.0 + usd`
  passed; a loop row built with `**` had none either, so
  `balance + balance ** 2` passed, stamped USD; and `'USD^2'` was a word
  of its own, so adding that line to `usd * usd` was refused as
  `USD^2 and USD^2`.
- **A `pyformula` result has no unit unless declared.** The engine cannot
  see inside the function, and a ratio of two USD lines was stamped USD,
  then added to amounts without a word. Declare the result's unit with
  the new `pyformula(unit='USD')`, or with `set_unit`.
- **A `pyformula` named like a spreadsheet function is not run as that
  function in a loop container.** A `@mo.pyformula` function named `max`
  (or given `name='MAX'`), applied inside a `loop=True` container to a row
  of the loop, is refused when the loop closes, as one of any other name is
  (`MAX(...) cannot be computed again when the loop closes`): build that
  row after the loop's last line. Before, the loop ran `mo.MAX` in its
  place without a word — a function doubling a row of 100, 200 and 310
  gave the row back unchanged.
- **A unit with several slashes reads as written.** `'USD/seat/month'`
  divides at each slash, so a seat count times a price in
  `'USD/seat/month'` is in `USD/month` and adds to a fee in `USD/month`.
  A bracketed group is a unit of its own (`'USD/(seat/month)'` is
  `USD * month/seat`), and a word may carry a note in brackets: hours in
  `'hours (billable)'` times a rate in `'USD/hours (billable)/month'` are
  in `USD/month`. A chain prints as typed; a string whose brackets do not
  balance keeps the old reading and prints as typed. Before, only the
  first `/` was read — `USD` per one symbol `seat/month` — so `seat`
  never cancelled (the product was in `seat * USD/seat/month`, and adding
  the fee was refused), and a note lost its closing bracket
  (`hours (billable`).
- **A unit-keyed number format finds its line by unit equality.**
  `ExcelView(formats={'USD/seat/month': '0.00'})` formats the price a
  model works out, `revenue / seats / months`, whose unit prints
  `USD/(seat * month)`, and `formats={'seat * month': '#,##0'}` formats
  `months * seats`, whose unit prints `month * seat`. Before, a format was
  found by the unit's printed text, so both lines were left in General. A
  pure number matches by its label, so a `'%'` format is never applied
  to a `bps` line.
- **A loop container's rows get the units their formulas give.** When a
  loop closes, a row's unit waits for the rows it reads (`+` takes the
  other side's unit, a product or a power waits for its factors), a lag's
  fill gives the row it lags its unit, and a row nothing settles (a flag,
  an index, a discount factor) is a pure number. With
  `opening = mo.lag(closing, fill=start)` and `start`, `cap` and `inflow`
  in USD, `opening / cap` is a pure number and a closing of
  `opening + inflow + opening ** 2 * k` (`k` in `1/USD`) is in USD; 0.2.2
  read a row that had no unit yet as a pure number, so the ratio printed
  `1/USD` and that closing was refused (`USD and 1/USD`). A unit that
  never settles is refused: `x = mo.lag(x, fill=1.0) * cap` (0.2.2 printed
  `USD^4`) meets `its unit does not settle`, with the two ways out, a
  factor with no unit or a declared unit on the row (`.set_unit(...)`). A
  mix is reported under the row that writes it on every build, where 0.2.2
  sometimes named a row reading it through a lag; and a formula built in
  the loop that fails when the loop closes passes its error to the rows
  built on it, where 0.2.2 handed them the numbers it held while the loop
  was open.
- **Inside a loop container, `line[0]` and a recurrence keep their unit.**
  A row written before the loop closes as `x[k]`, `x[a:b]`,
  `mo.recurrence`, `mo.recurrence_sum` or `mo.cumsum` kept no unit, so
  `balance[0] + fee_usd` passed in silence and `balance / balance[0]`
  printed EUR. They now carry the line's (the start's) unit, as built
  outside a loop, and a mix is refused when the loop closes.
- **A loop refused when it closes hands out no numbers.** A loop
  container whose unit check refused it at close (an `EUR` fee added to a
  `USD` amount) still handed its rows out with the numbers computed from
  the mix and wrote them into the book. Every row, and
  every line built on one, now refuses its value with the refusal, as a
  loop refused before it computes does; the way out is to build the
  container again.
- **A count in `bps` or `pp` beside a share in `%` is refused.** `bps` and
  `pp` are labels that never rescale (50 bps is written 0.005, 2 pp is
  0.02), so a `bps` or `pp` line holding a count (50) is refused beside a
  `%` line holding a fraction (0.05) where the two are added or taken as
  one: `+` and `-`; every function that combines them (`mo.SUM`,
  `mo.AVERAGE`, `mo.MIN`, `mo.MAX`, `mo.IF`, `mo.IFS`, `mo.SWITCH`,
  `mo.CHOOSE`, `mo.IFERROR`, `mo.MOD`, `mo.CEILING`, `mo.FLOOR`, the
  conditional totals, the statistics, the cells a lookup returns); a
  recurrence template (`mo.recurrence(share, '{prev} + {x}', x=counts)`);
  a lag's fill (`mo.lag`, `.shift`); and a loop container, when the loop
  closes, under the row. `rate + step`, with `rate` holding 0.05 in `%`
  and `step` 50 in `bps`, is refused with the way to write it: `A bps line
  holds the fraction (50 bps is 0.005), or the count is converted first:
  step / 10_000.` A share written as points (30 % + 2 pp) and a fraction
  stepped by a fraction (0.05 % + 0.005 bps) add as before. Before,
  nothing made the check, and `rate + step` came out 50.05 without a word.
- **`.shift`, `NPV`, `XNPV`, `PMT`, `FV` and `PV` carry units, and a lag's
  fill is checked.** A lag is in its line's unit, and so must be its fill:
  `mo.lag(balance, fill=opening)`, the balance in USD and the opening in
  EUR, is refused (`mo.lag takes its values in one unit; 'balance' and
  'opening' are in USD and EUR`), and inside a loop container the check is
  made when the loop closes. A number or a line without a unit fills any
  line, and a line without a unit takes any fill. `balance.shift(1)` is in
  USD, as `mo.lag(balance)` is. `mo.NPV` and `mo.XNPV` are in their flows'
  unit, and `mo.PMT`, `mo.FV` and `mo.PV` take their amounts — a present or
  future value, a payment — in one unit, the result's: `mo.PMT(rate, 24,
  loan, balloon)` with the loan in USD and the balloon in EUR is refused;
  the rate and the number of periods are not amounts. Before, a fill or a
  balloon in any unit passed, and `balance.shift(1)` and the five functions
  had no unit: `mo.NPV(rate, flows) + fee` passed with the fee in EUR.
- **`.shift(n, fill_value=<line>)` keeps the fill as a live cell.** A
  line given as the fill was frozen into the formula as its number
  (`=1000.0` in the first cell), so the workbook did not follow an edit of
  the opening cell. The fill is now a reference to its cell (`=inp!B2`),
  as `mo.lag(x, fill=<line>)` keeps it, and `.shift` follows `mo.lag`'s
  unit rule (above): a fill in another unit is refused, in a loop
  container when it closes; a list fill or a fill assigned further down
  is refused, as `mo.lag` refuses them. The 0.2.0 note that `.shift`
  inlines its fill no longer holds.
- **`NPV`, `IRR`, `XNPV` and `XIRR` over a list of lines read each
  line.** `mo.NPV(rate, [y1, y2])` with lines in two units ran in silence
  and stamped no unit, so `usd + NPV(rate, [eur, eur])` passed; the
  book wrote the computed number with a false "reduces a literal list"
  warning, and inside a loop container a list holding a total of a row
  read before its line kept the number it held while the loop was open.
  The list is now
  the lines' cells: they must share one unit (two are refused), the result
  carries it (`NPV`, `XNPV`), the lines are the call's dependencies,
  and the book writes `=NPV(B1, B3:B4)` / `=IRR(B2:B4, 0.1)`. A list
  of numbers typed in stays a value, as before.
- **A recurrence template is read for units.** `mo.recurrence(eur,
  '{prev} + {x}', x=usd_line)` had no unit and ran in silence; the
  template is now read as any formula is: a mix is refused with the
  operators' own message (`recurrence("{prev} + {x}"): Cannot
  add/subtract values with incompatible units: EUR and USD`), a start
  without a unit takes the template's, a template that multiplies the
  line by a dimension every period is refused as a unit that does not
  settle, and inside a loop container the check is made when the loop
  closes, under the row. A declared `unit=` decides what the line
  prints, never what its template may add.
- **`CHOOSE` over a keyed line.** A choice that runs along keys — `mo.Variable({'Bear':
  0.9, 'Base': 1.0, 'Bull': 1.1})`, or a line on one non-time `indexed_by` axis — is, on
  its own, one choice per key in key order: `mo.CHOOSE(pick, growth)` is `mo.CHOOSE(pick,
  growth['Bear'], growth['Base'], growth['Bull'])`, written `=CHOOSE(B1, B3, C3, D3)` over
  the line's cells (one such formula per period when the index is a series; keys that
  are not text spread by position, as `[]` and the workbook read them). Before, the line
  was read as one period series: its keys were lost, a 3-key line in a 4-period window
  was refused for its length, and the workbook held `=CHOOSE(B1, B3)` in every period. Refused
  with an error that names the fix instead of a silent number: a keyed line beside other choices
  (the numbering would depend on its key count), a computed keyed line with no cells of
  its own (name it first), a tracked or keyed index, a series index over tracked
  choices, a line on two axes. A scalar index over tracked choices still switches track
  by track, as before.
- **A line along keys against a period series is refused.** `base * growth`, with
  `base` a period series — a list, a located list (`start=`/`extend=`), a tracked line,
  `mo.time.days` — and `growth` a keyed list (or a line on one finite non-time axis) of
  the same length, was zipped cell by cell — January with Bear, February with Base — and
  the result carried the keys, so the numbers looked right and were wrong; unequal
  lengths died with the generic length error. Every arithmetic operator and comparison
  now raises, naming both lines, the pairing a zip would make, and the fixes that run
  (pick one key, `growth['Bear']`; give the series the same keys or lay it along the
  same axis; or switch with `mo.CHOOSE(pick, growth)`). A keyed list against a line on
  an axis of other labels is refused too. Scalars, one-value bare lists, a raw Python
  list (`keyed * other.value`), keyed × keyed, axis × axis and series × series
  broadcast as before. The layout, the operators and `CHOOSE` use the same test of
  whether a line runs along keys.
- **A wrapping Variable keeps the keys and axes of the line it wraps.**
  `mo.Variable(growth * 2, display_name='Twice')` ran along nothing — the keys of the
  product were dropped, as its unit once was — so the named line lost its caption row
  and its place in `CHOOSE`. A declared `keys=` or `indexed_by=` on the wrapper still wins.
- **A per-period field inside an `orient='columns'` block is refused, not dropped.** The
  matrix holds one cell per field per instance; a list field (a period series, a keyed
  line, a line on an axis) was skipped without a word — its row vanished and every
  aggregate reading it was written as a dead number. The layout now records the refusal
  and `to_excel` raises it before writing a book, naming the
  line, the instance and the block, with the forms that hold such a line (the sheet, or
  `orient='records'` for a period series); `check_compat` reports it as a `layout`
  issue; the notebook repr and every other reader keep their addresses. A one-value
  bare list field is a cell, as a one-value list is a single value everywhere else.
- **A name whose line was overwritten by a plain value can be removed.** After
  `inputs.rate = 0.05` written over the line `rate`, `inputs.remove('rate')`,
  `del inputs.rate`, and moving `inputs` to another container raised
  `AttributeError: 'float' object has no attribute '_qualified_id'` and could leave the
  container half detached. The plain value is removed or carried along like a line.

## [0.2.2] — 2026-10-19

### Added

- **A caption row over a keyed line.** A line that runs along keys instead
  of periods — `mo.Variable({'FY2023': 400, 'FY2024': 450})`, or a line on
  one non-time `indexed_by` axis — laid out on a period-headed sheet gets
  one reserved row directly above it carrying its keys, its cells under
  them from the first period column; consecutive keyed lines with the same
  keys share the row, a hidden line hides it unless a visible sibling shares
  it, and a subtotal column writes nothing over it. Before, the
  line's cells sat under the year captions as if they were periods. The
  notebook repr draws the same row.
- **A unit row under a table's captions.** On a sheet that declares a unit
  column (`ExcelView(meta={'unit': …})`), a records table or a board of
  scalars writes its fields' units in one row under the caption row, in
  the unit column's style, and keeps the captions bare; without a unit
  column the unit is suffixed into the caption (`EV (USD m)`). Before, a
  caption's unit was written one column right, over the neighbour.
- **`check_compat(model, 'excel')` names every number baked into a
  formula.** A cell whose formula is live but carries a computed literal
  (an aggregate over an expression no row holds, a `mo.time` read with no
  time window, a track coordinate without a row) is listed with its line,
  its period and the sub-expression that became a number, beside the cells
  written as plain values; `CompatIssue` gains optional cell fields, its
  positional shape is unchanged.
- **`SUMPRODUCT`.** `mo.SUM(price * volume)` over laid-out rows is written
  `=SUMPRODUCT(C2:F2, C3:F3)` (three rows too); before, the cell held the
  engine's number.
- **A declared sheet order.** `ExcelView(sheet={'order': ['summary',
  'inputs']})` on a model puts those sheets first, as listed, each named by
  the attribute it is assigned to (`model.summary = …`; the tab of the
  model's own lines goes by the model's name). The rest follow in the order
  the model declares them. Every renderer that lists the sheets reads it
  (the workbook, the notebook tabs), and only from the view of the model
  being written, so a container written on its own keeps its own order.
  A single name is an order too, and a repeat is harmless. A name that is
  no sheet is skipped with a `SheetOrderWarning` that lists the sheets
  there are, and the rest of the order still applies. A model without the
  declaration writes the same workbook as before.

### Changed

- **`ExcelView(meta={'unit': False})` shows units nowhere.** The unit flag
  has three states: a caption (or `True`) puts the units in a column of
  their own and, under a table's captions, in a unit row; unset suffixes
  each unit into its row label and table caption (`Revenue (USD m)`);
  `False` hides the unit text everywhere — no column, no suffix, no unit
  row — in the workbook and the notebook repr alike, while the number
  formats keyed by unit (`formats={'%': '0.0%'}`) still apply: the unit
  stays on the line, only its text is hidden. Before, `False` and unset
  wrote the same sheet; a workbook that declared `False` loses its
  suffixes.
- **`mo.MIN` / `mo.MAX` over two or more operands work period by period.**
  `mo.MIN(sweep, balance)` is the smaller of the two in every period and
  `mo.MAX(cash, inputs.min_cash)` the larger: a series, with scalars
  broadcast as `a + b` broadcasts them and an operand with tracks taken
  track by track, written to every column as `=MIN(C2, C3)` / `=MAX(C4,
  Inputs!B2)` — the cap and the floor of a hand-built workbook. Inside a
  `loop=True` container they are allowed, like `IF`. Before, two series
  reduced to one number across both rows (`=MIN(B2:F2, B3:F3)`) and a
  cap had to be spelled `mo.IF(a < b, a, b)`. Over one operand
  `mo.MIN(row)` stays the reduction (`=MIN(B2:F2)`), over scalars only it
  stays one number (`=MIN(B2, B3)`), and a one-operand reduction inside a
  loop stays refused, now pointing at the two-operand form.
- **A `+` chain past Excel's 8,192 characters folds into `SUM`.** A run of
  `+` over plain cell references of one sheet whose text would pass the
  limit is written as `=SUM(first:last)` runs (at most 255 pieces); any
  shorter chain is written as it always was. Before, the cell held the
  engine's number and the writer logged the limit.
- **`meta={'numbering': 'section'}` numbers a sheet as an outline.** A
  first-level section is `1`, a section inside it `1.1`, and a line is
  numbered one level below its section (`1.1.1`). Sections and lines of one
  container count together, so a line on the sheet itself takes the next
  number beside the sections. A section's own `article` wins, and its lines
  count from it. A row that repeats another row (a total shown in a header
  row, a track's copy of a line, a repeated instance in a matrix) takes no
  number; a records table is one item; the rows of a columns matrix are
  numbered like lines. Before, only lines were numbered, flat within each
  first-level section.

### Fixed

- **A compound lag fill is written in brackets.** `x - mo.lag(y, fill=a + b
  - c)` was written `=B2 - I!B1 + I!B2 - I!B3` — the fill bare inside the
  enclosing operator, so the spreadsheet computed 120 where the model
  computed 40. The fill now reads as one operand
  (`=B2 - (I!B1 + I!B2 - I!B3)`) under every operator, under unary minus
  and inside a recurrence template; a simple fill is written as before.
- **A subscript is its cell, never Python text.** `line[-1]` resolves to the
  last period's cell (`=Inp!D2 * 2`), an element of a line with no row of
  its own is its number, and a subscript or method the renderer cannot
  spell is written as the cell's value; before, the cell held the model's
  source text (`=m.inp.ufcf[-1] * 2`).
- **A columns board starts after the sheet's lead columns.** A board
  (`excel_props={'orient': 'columns'}`) on a sheet with a unit, note or
  constants column began at column B regardless, so the unit column
  overwrote its first instance; it now begins where the period columns
  do. A board's or a records container's own direct lines are laid out
  beneath the table (before, they were dropped from the workbook), and a
  line's `header_row` setting is honoured there too.
- **A bare literal in a number format is quoted.** `0.00x` is written
  `0.00"x"` (`#,##0 bn` → `#,##0 "bn"`); date and time codes, `General`,
  `E+`, `AM/PM` and anything already quoted or escaped are left as typed.
- **A workbook carries no green triangles of its writer's own making.**
  Excel marks a suspected mistake with a green triangle in a cell's corner,
  which a reader takes for a note that is not there. Line numbers and year
  labels written as text, and the formulas the writer put in (a quarter's
  total between its months, a range that starts after the opening column),
  are now recorded in each sheet as checked. A formula that computes an
  error keeps its mark, and so do typed values, labels and whatever a
  reader types later into a cell that held no formula.
- **A workbook's formula cells carry the values the engine computed.**
  Every formula cell `to_excel` writes for a row, a header row repeating
  another row, an opening cell or a period caption now holds the value the
  engine computed beside its formula. A reader that does not recalculate
  (a file preview, `openpyxl` with `data_only=True`, pandas) shows the
  model's numbers instead of empty cells, zeros or dates in 1899. Excel
  recalculates on open; LibreOffice, by default, shows the cached values.
  A row total carries the sum of its row. A formula cell with nothing to
  cache (a subtotal column, an empty period, a value no cell can hold) is
  written with no value at all, so a preview shows it blank instead of 0.

## [0.2.1] — 2026-10-06

### Added

- **`mo.time` — the time of the period a formula is computed for.**
  `mo.time.start`, `mo.time.end` (dates), `mo.time.days` (the period's
  calendar day count) and `mo.time.index` (its 1-based number) are
  readable by any formula of a windowed model and written by nobody:

  ```python
  m.debt.interest = mo.lag(m.debt.balance) * m.inputs.rate * mo.time.days / 365
  ```

  A formula carries time, not numbers (`formula` reads `time.days`), so
  the same model is right on any grain: `m.at('quarter')` recomputes
  time for quarters (90, 91, … days) instead of rolling months up.
  Time is read from the innermost open `with` block's window, else the
  model built last; a line whose time was read for another window is
  refused when it joins its container.
- `mo.time.months` — the calendar months in each period (1 a month, 3 a
  quarter, 12 a year), beside `mo.time.days`: an amount a month times it is
  the period's amount on any grain. A header row of it counts the months
  from the period's dates when the header also has the period's end,
  else writes the grain's constant (`=3` for a quarter); a day window
  refuses it.
- **A quarter or year window may start on the 1st of any month.** A yearly
  window started on 1 May 2026 runs May–December 2026 (`mo.time.months` 8,
  `mo.time.days` 245) — a short first period, a stub — then whole calendar
  years; the first period keeps its label (`2026`). Amounts are the
  model's to scale: an annual rate times `mo.time.months / 12` is the
  period's amount. A day inside a month is still refused. A series whose
  first period is short does not combine with one where that period is
  whole, and a mean over a short quarter weighs its days. A coarser view
  whose first bucket starts inside its period (a monthly model from May
  seen by quarter) treats that bucket as a short first period.
- **The workbook's look, as `ExcelView` settings — every one off until a
  view asks for it**, so a model without a view writes the book it always
  did:
  - `sheet={...}`: a title row (`'title': True` or a text), `'gridlines':
    False`, `'same_columns': True` (the same lead columns on every sheet),
    `'columns': [...]` with each lead column's width and look, `'full_rows'`
    (a section's line drawn across the sheet);
  - `cell_types={'styles': {'input': {...}, 'link': {...}, 'formula':
    {...}}}`: inputs, links to another sheet and formulas each in their own
    look, beside the colours `cell_types` already took;
  - `formats={'date': ..., 'number': ..., '<unit>': ...}`: a number format
    by what a line is, where the line names none;
  - `bands={'result': {...}, 'mirror': {...}}`: the look of a line marked
    `excel_props={'result': True}` (its section's result), and of the header
    rows a sheet repeats at its top;
  - `timeline={'opening': 'Pre-start', 'repeat_on_anchor': True}`: a column
    before the first period the time rows are spelled from, and the header
    rows repeated on the anchor sheet too;
  - `meta={'numbering': 'section', 'row_sum_for': ['sum', ...]}`: line
    numbers by section, and a row total for lines whose re-grain rule is
    listed.
- **`mo.house_view(*views)`** — house styles laid under a model's own
  view, most general first (a firm-wide standard, then a team's
  adjustments); each layer changes only what it names, and the model's
  own view stays on top: `with mo.house_view(firm, team):
  model.to_excel("deal.xlsx")`.
- `modeleon.compile.excel.writer.write_laid_out_sheet` and
  `set_book_font` — public entry points for writing a laid-out sheet into
  a workbook an exporter builds itself.
- `Variable.set_unit(unit)` — the unit as a chained setter, beside
  `set_display_name`, `set_description` and `set_style`.
- **`mo.Unit.dimensionless(label)` — a word that labels a pure number.**
  A unit string is always a dimension (any word: `'USD m'`, `'GWh'`), and
  amounts in different dimensions never add. A discount factor, a share or
  a flag is a pure number that should still print its word: declare it
  once and pass the object —

  ```python
  FACTOR = mo.Unit.dimensionless('factor')
  m.val.df = (1 / (1 + m.val.wacc) ** m.val.years).set_unit(FACTOR)
  m.val.pv = m.val.cash * m.val.df          # the cash's unit: the label is inert
  ```

  It prints in the unit column and beside the row's name, and multiplies,
  divides and adds as the number it is — the way `'%'` always has (`'%'`,
  `'x'`, `'bps'`, `'pp'` are the same kind of unit). Pure numbers add to
  one another, the left label printing; a product of two drops the label.
  A label is one word: `'a/b'` or `'m^2'` is refused (a compound unit is a
  dimension). When an addition refuses `USD m` and `USD m * factor`, the
  message now names `'factor'` as the word that is a dimension there and
  shows the declaration. `.code` writes a labelled unit back as
  `mo.Unit.dimensionless('factor')`, and an aggregate over a line with
  `unit='factor'` and one with `unit=FACTOR` names no unit.
- **A window may start at a date.** `default_start` takes a date or a
  date Variable — `m.default_start = m.inputs.financial_close` — as well
  as a label. The window then starts at that period; rows that inherit
  the window move with it, rows with their own `start=` stay on their
  dates. The date must open a period (the 1st of a month for a monthly
  window). A quarter or year window started at a date cell ends each
  period with its calendar quarter or year whatever month the cell holds
  (`=EOMONTH(B9,11-MOD(MONTH(B9)-1,12))`), so a date moved in the book
  moves the periods as the engine does; totals columns and coarser views
  are written for the date the book was built with.
- **`default_header` — the rows a workbook shows above its sheets.** A
  MultiVariable (`m.default_header = m.header`), a list of rows, or `None`
  for no header, set on the model or any container below it. Like
  `default_excel_view` it is a setting, never a row. In the workbook the
  header MultiVariable is the anchor of time: its time rows are live
  formulas off one start cell — the `default_start` Variable's cell, or a
  typed date (`start = previous end + 1`, `end = EOMONTH(start, 0)`, `days
  = end - start + 1`). Every other sheet shows the header's rows at its
  top as references to it, whatever the order the code declared them in; a
  sheet's own list or `None` overrides that. A formula reading
  `mo.time.days` points at its own sheet's header row, else at the anchor,
  and period captions become the period's end date, formatted. Without a
  header, time is spelled off the start inline (a month's days as
  `EOMONTH(start, 0) - EOMONTH(start, -1)`) — a formula, never a number.
  Day, month and period-number rows are formatted as numbers (`0`) unless
  the line or its view names a format. A book that reads no time and
  declares no header is written byte for byte as before.

- **Loop containers — rows that read each other across periods.** A
  container built with `loop=True` lets its rows be read before the
  lines that assign them:

  ```python
  m.debt = mo.MultiVariable("Debt", loop=True)
  with m.debt as d:
      d.interest = mo.lag(d.balance, fill=m.inp.opening) * m.inp.rate
      d.pot      = m.inp.cash - d.interest - m.inp.scheduled
      d.sweep    = mo.IF(d.pot > 0, d.pot, 0.0) * m.inp.share
      d.balance  = mo.lag(d.balance, fill=m.inp.opening) - m.inp.scheduled - d.sweep
  ```

  Row by row that is a circle; period by period it is a plain order of
  evaluation, and the engine settles it the moment the last row is
  assigned — rows on a cycle period by period, everything downstream of
  them as whole series. Inside a loop container rows may come in any
  order. The workbook is the shape a hand-built one has always had: the
  loop closes through the previous column (`C2: =B5 * Inputs!B2`), so
  Excel sees no circular reference, and the opening balance is the `fill`
  — a literal, or the input cell when the fill is a Variable. Values,
  units and `formula` come out as for any other row. A class is a loop
  container for every instance with `class Tranche(mo.MultiVariableClass,
  loop=True)`. A loop is built whole: assigning a row of a loop container
  a second time is refused — build the container again (in a notebook,
  keep its line in the same cell as the loop). Needs `default_periods` on
  the model: the placeholder that holds a row read ahead is a series as
  long as the horizon.

  What is refused, by name: a cycle within one period, as the new
  `SamePeriodCycleError` (its `rows` are the rows of the cycle); a lead
  (`mo.lag(x, -1)`) from one loop row into another; a whole-series
  function over a loop row inside the loop — `SUM`, `MAX`, `MIN`,
  `NPV`, … read the whole series, so a per-period floor is
  spelled `mo.IF(x > 0, x, 0)`; a row read ahead but assigned a
  container, a plain number, a single value or another row not yet
  assigned, as `ForwardReferenceError`; a `fill=` assigned further down;
  `pyformula=` and lambda recurrences over a loop that has not closed. A
  row read ahead but never assigned (usually a misspelt name) is
  reported when the container's `with` block ends, with the line that
  read it. A row that waits on the loop refuses `.value`, `bool()`,
  `float()`, `int()`, iteration and `.at()` until the loop closes; its
  HTML display and `to_dict()` show no values yet; `to_excel` refuses
  while a loop is open. `model.at(grain)` refuses a loop whose rows have
  no re-grain rule and names them; `SamePeriodCycleError.nodes` lists
  every Variable on the cycle.

  Nothing changes for any other container: reading a row a container
  does not have raises `AttributeError` as before, and `hasattr` /
  `getattr(mv, name, default)` answer as before. A component may still be
  named `loop` (`MultiVariable(loop=mo.Variable(...))`): only a bool
  `loop=` declares a loop container. An extension that inspects models
  can mark its own code with `modeleon.core.mv_context.trust_code_in(dir)`,
  so its attribute probes never create rows in a loop container. `detect_cycles()` /
  `assert_acyclic()` no longer count a `lag` read as an edge.
  `ForwardReferenceError` and `SamePeriodCycleError` are exported from
  `modeleon`. A loop row keeps a declared unit (adding two currencies in its
  formula is refused as anywhere else), and a loop container mounted
  under a second parent (`s.debt = m.debt`) stays a loop container.

- **`YEARFRAC`, `DAYS` and `XNPV`.** `mo.YEARFRAC(start, end, basis)` (all
  five bases), `mo.DAYS(end, start)` and `mo.XNPV(rate, flows, dates)`
  compute what Excel computes and are written as live formulas
  (`=YEARFRAC(...)`, `=_xlfn.DAYS(...)`, `=XNPV(...)`). XNPV refuses a
  rate of zero or below and a date earlier than the first, as Excel does
  (`#NUM!`); a blank flow reads as 0 on its date. The values are checked
  against Microsoft Excel 16.95 recalculating the same formulas.

### Changed

- **A blank counts as zero in a formula, as in Excel.** A month nobody
  entered used to make every formula over it blank ("a hole in, a hole
  out"); it now computes over 0: `blank + 5 = 5`, `blank * 5 = 0`,
  `5 / blank = #DIV/0!`, `-blank = 0`, a blank compares as 0 (or as `""`
  against text), a recurrence's balance stays where it was over a blank
  inflow (a zero divisor there is `#DIV/0!`, carried on), a running total
  runs over it, and a lag of it is 0. Functions read it as Excel does:
  `ROUND` / `ABS` / `INT` of a blank are 0, `MOD` by a blank is `#DIV/0!`,
  `IF` / `CHOOSE` landing on a blank show 0, `AND` / `OR` skip it (all
  blank is `#VALUE!`), `LEN` is 0 and `UPPER` / `LOWER` are `""`, `NPV` and
  `IRR` skip it, `XIRR` reads a blank flow as 0 on its date, and a
  blank argument of `PMT` / `FV` / `PV` is 0. A reference to a line
  (`mo.Variable(row)`, `actual=row`) shows the row's blanks as 0, as `=A1`
  does. A line of dates or words keeps its blanks. Only an input nobody
  filled stays blank, so a derived actual row shows 0 in months nobody
  entered, and its quarter is the sum of its months.
- **The blend asks a formula side what was entered.** A side computed by a
  formula is a number even over months nobody entered, so "present" is no
  longer "not blank" for it: the blend reads the formula through to the
  entries it is built from (`modeleon.core.entries` - same-month arithmetic,
  comparisons, `IF` / `ROUND` and the other month-by-month functions,
  `.at(track=...)` of the same track, and, in an expression written for
  one track, loaded actuals in a line without tracks) - nothing entered in a
  month takes the follow side, anything entered the given side, its
  blanks as zeros. A side entered by hand is still present where it is
  not blank; a month with neither side shows 0. The workbook asks the same
  cells, `=IF(COUNTA(<entries>)=0,follow,given)`.
- **`mo.DATE` of three numbers, on a line of its own, is a typed date:
  an input.** `inputs.close = mo.DATE(2026, 7, 1)` (or wrapped,
  `mo.Variable(mo.DATE(2026, 7, 1), display_name=...)`) is written to the
  workbook as the date itself, in a date format and in the input look,
  as `mo.Variable(date(2026, 7, 1))` always was; it was the formula
  `=DATE(2026, 7, 1)`. The value is unchanged. So is a list of them
  (`mo.Variable([mo.DATE(2026, 1, 15), mo.DATE(2026, 2, 15)])`) and a
  track given one (`mo.Variable(plan=mo.DATE(2026, 2, 1), ...)`). Inside
  a formula (`mo.IF(mo.time.start >= mo.DATE(2026, 7, 1), 1, 0)`) it
  stays the `DATE(...)` call, and a date built from another line
  (`mo.DATE(year, 1, 1)`, `mo.EDATE(close, 6)`) stays a formula. A
  model's code spells an untracked line's date back as
  `mo.DATE(2026, 7, 1)`, which builds the same input again, never
  `datetime.date(...)`.
- **A period not entered yet counts as zero in a coarser bucket, as a
  blank cell does in Excel.** A quarter or year of a plan/actual line
  whose actuals stop in August sums the months entered; a wholly future
  quarter is `0`, and a ratio over it `#DIV/0!`, as the workbook shows.
  The engine's bucket was a hole ("not known yet") while the workbook's
  subtotal summed what was entered. Now the two agree for every rule: a
  `geometric` rate compounds a blank month as 0%; `mean`, `min` and `max`
  skip a blank month as Excel's `AVERAGE`, `MIN` and `MAX` do (a mean over
  the entered months' days, written `IF(COUNTA(...)=0,0,...)` where a
  month is blank), and are `0` when no month is entered; `first` / `last`
  read only their own period, a blank there being `0` and a sibling's
  error not theirs. A data line's subtotal and an untracked line's
  `.at()` read are live formulas, no longer written blank (a line with
  tracks read at a coarser grain is written as the engine's values). A
  line of dates or words keeps a blank, having no zero, on every track.
  A blended track (`mo.blend`) is unchanged: it is spliced month by month
  before it is folded, and the workbook no longer splices the folded
  tracks again on a sheet of `.at()` reads or in a model projected to a
  coarser grain.
- A section that names only its grain - `mo.MultiVariable('Debt',
  default_grain='quarter')` under a yearly model - is refused when a line
  reads `mo.time` in it or a series it would place joins it, with what it
  is missing: its time was read off the window above while its lines stood
  at its own grain. A series needs the section's start, `mo.time` its whole
  window. Scalars and lines located on their own (an `.at()` copy) join as
  before, a section of the window's own grain still lives in it, and the
  model may name its grain before its start.
- Two lines of different grains added, subtracted, multiplied or divided:
  the error names both, their grains, and the re-grain that joins them
  (`interest.at('year')`).
- `Unit` is a plain value, no longer a `MultiVariableClass`: a unit
  assigned to a model (`m.USD = mo.Unit('USD m')`) is an attribute, where
  it used to become a section of its own — a sheet named
  after it, with a `Display` row — in the workbook. `Unit(...)` takes only
  its string (an unknown keyword is a `TypeError`, not silently kept), and
  `repr(unit)` shows how it was made (`Unit('USD m')`,
  `Unit.dimensionless('factor')`).
- A unit that is neither a string nor a `Unit` (`unit=5`, a Variable) is a
  `TypeError`; it used to be dropped without a word. A container in its
  place still gives the row no unit.
- **A compare book captions its deviation column "Var"**
  (`ExcelView(tracks='compare')`), and the deviation subtracts the track
  named `plan` when the model has one.
- **`XIRR` accepts a date earlier than the first.** Excel computes such
  a schedule (measured in Excel 16.95), discounting every flow to the
  first date listed, so the engine no longer refuses it. `guess` may be
  a Variable, and a guess of -100% or below is refused (`#NUM!` in
  Excel).
- A row total (`excel_props={'row_sum': True}`) is written in its row's
  number format — unless that format would not show a number: a row drawn
  as a bar (`'"X";;'` marks any positive value) totals to a count, and the
  count stays readable.

### Fixed

- **Unary minus on a line with tracks works track by track**, and `ROUND`,
  `ABS`, `INT` and `MOD` take a line with tracks too (all raised `TypeError`).
  `MOD` over a per-period divisor is element-wise.
- **A line wrapping a line whose actual is entered by hand references it.**
  `mo.Variable(row, plan=[...])` copied the row's expression, which
  describes the row's other tracks only: when the row's months were a
  list of formulas (`[principal * rate / 12 for _ in months]`), the wrapper's
  actual row in the book repeated the plan's formula in every month,
  including those nobody entered. The wrapper now refers to the row
  (`=Savings!C7` on each track's row, and its formula names the row);
  values are unchanged.
- **A line that enters its plan and derives its actual takes the forecast
  in the book where nothing was entered, as the engine does.**
  `mo.Variable(a + b, plan=[...])` computes its actual as the sum of the
  operands' actuals, so its actual row is a formula, and Excel computes
  that formula as 0 over the months nobody entered. The splice asked
  `=IF(actual="",forecast,actual)` of it and took the 0 where the engine
  takes the forecast: a P&L whose actuals stop in Q3 showed Q4 revenue as
  0 in the workbook where the engine had the forecast. The splice now asks
  the entries the actual is built from, as the engine's blend does:
  `=IF(COUNTA(<a's actual>,<b's actual>)=0,forecast,actual)`, or `=IF(<a's
  actual>="",…)` for one entry (an entered error or TRUE/FALSE counts as
  entered). It reads through the formulas between as long as they work
  month by month (arithmetic, comparisons, `IF`, `ROUND` and the like); a
  constant or a line without tracks is no actual's entry, and a part it
  cannot read (a running total, an aggregate over time) adds none. The
  blue font that marks a blended cell holding an actual asks the same
  entries.
- **A formula with a constant step keeps its brackets in subtotals and
  coarser reads.** `y = x * (1 - share)` wrote its quarter as
  `=E3 * 1 - A!B1` (299.75 where the engine has 225): the constant step
  reaches the coarser grain as a copy, and the copy hid the subtraction
  from the bracket check. Every line built on such a step, on any sheet,
  carried the wrong number.
- **Subtotals of a line with tracks are live formulas on every track.**
  A formula line's display row inlined its operands' whole tracked values
  as text (`=TrackValues({...}) - TrackValues({...})`) and each track's
  row (`Margin · actual`) was `#VALUE!`; both now fold by the line's
  formula over its operands' bucket cells on the same track (an operand
  on another sheet is written as its number for that track), and a
  product of two series writes the engine's number for its track. A
  track authored with its own formula (`plan=<expr>`) on a line with no
  rule is refused, as the engine refuses it, in every view. A
  running total of a line with tracks re-grains track by track (it was
  refused), a running total's subtotal after the first reads its own
  previous bucket (`=E11 + I3`) instead of a typed number, and a running
  total inside a larger formula (`cumsum(in) - cumsum(out)`) sums its
  own cells - a range from the first to the last took in the columns
  between them, in subtotals and in month cells alike.
- **`mo.ratio` on a line with tracks folds track by track** - sum over sum
  of the named lines' same-track series, the blend among them, as the
  workbook's subtotal does. It recomputed the formula over each operand's
  own rule (a `mean` denominator gave 6 where the book had 2), and a typed
  ratio line with tracks was refused. A named line without tracks divides
  every track, in the engine and in the workbook, and a track holding one
  value counts it in every period of the bucket.
- **`mo.recurrence` and `mo.recurrence_sum` keep `regrain=`, `unit=`,
  `description=` and `excel_props=`** beside `variables=`: they were
  dropped without a word, so an opening balance declared
  `regrain=mo.up('first')` had no rule and its subtotals were `#VALUE!`.
  Any other keyword beside `variables=` is a `TypeError`; a template may
  still use one of those names as its own variable. A recurrence over
  scalars alone, which cannot count its periods, now says it is a formula
  of the period's time (`base * (1 + g) ** (mo.time.index - 1)`) rather
  than asking for a typed `periods=12`.
- Subtotal columns (`timeline={'totals': [...]}`) are formulas for every
  re-grain rule: `first`, `min`, `max`, `geometric` and a `mo.ratio` were
  written as typed numbers.
- A line read at a coarser grain (`.at('year')`) folds in the workbook as
  the engine folds it: a `mean` weighted by the periods' days
  (`(B2*31+C2*28+...)/365`), a `geometric` rate compounded
  (`EXP(SUMPRODUCT(LN(1+B2:M2)))-1`, `#NUM!` at -100% as the engine's
  value). The same in subtotal columns, where a mean was a plain `AVERAGE`.
- A ratio subtotal reads its numerator and denominator on its own track,
  and a long bucket (a year of days) folds in ranges, within Excel's limits
  of 255 arguments and 8192 characters.
- A line read at a coarser grain reads only its own periods' cells when
  subtotal columns stand between them (`SUM(B4:D4,F4:H4,...)`); one range
  from the first to the last cell summed the subtotals too.
- A unit declared on a row is kept when the model is viewed at a coarser
  grain (`m.at('quarter')`); the recomputed formula's own unit replaced it.
- A view's `font` reached the data rows but not the period header, the
  column captions or the section rows: a model asking for Calibri 10 got
  a book in two sizes. The book's font is now its Normal style, and every
  cell the writer dresses starts from it.
- An error read as TRUE: `IF(margin > 0.2, 50000, 0)` over a month
  with `#DIV/0!` showed 50,000 where the workbook shows `#DIV/0!`, and
  `AND` / `OR` / `NOT` did the same. As in Excel (measured in 16.95), a
  condition that is an error makes `IF`, `AND`, `OR`, `NOT` and `CHOOSE`
  that error — `AND(FALSE, #DIV/0!)` too — while an error in the branch
  not taken changes nothing. `CHOOSE` by an error index is that error,
  not `#VALUE!`.
- `SUM`, `MAX`, `MIN`, `AVERAGE` and `ABS` raised Python's `TypeError`
  on an error cell (`'#DIV/0!'` from a division upstream) and on a hole
  — so the total of every plan/actual row with an actual month not
  entered yet failed. They now read a range as Excel does (measured in
  Excel 16.95): the first error in reading order is the result, a hole
  is skipped as a blank is, and with nothing left SUM, MAX and MIN are 0
  and AVERAGE is `#DIV/0!`. `ABS`, `ROUND`, `INT` and `MOD` pass an error
  cell through.
- `IRR` and `XIRR` could return a rate that is not a root: leading
  zero flows let a large rate shrink the NPV under a fixed tolerance
  (eleven idle months before `-1000, 100 × 11` "converged" at 564% a
  month; Excel and the root say 1.62%), and Newton's step could run
  away on long monthly schedules.
  Both now keep Newton's answer only where the NPV changes sign, and
  fall back to bisection otherwise. XIRR's tolerance is relative to the
  size of the flows, so flows in billions and in thousandths both land
  on the root.
- A product of two series re-grained correctly only when both were
  named rows: through an unnamed intermediate (`a * rate * g`) the
  quarter came out as Σa·rate·Σg instead of Σ(a·rate·g).
- A product over tracks (`qty * price` with plan/actual, or
  `fx * fee * revenue`) could not change grain: `m.at('quarter')` raised.
  Each track now sums as its own flow.
- `IF(condition, series, other)` with a **scalar** condition choosing a
  **series** was typed `scalar` although its value is a list; every
  consumer that dispatches on `var_type` — `cumsum`, `recurrence` over
  such a row, or your own code that stores values by `var_type` — saw
  the whole series in one period, or nothing. The branch now decides
  the shape.
- A running total or a recurrence with no row of its own, used inside a
  larger formula (`mo.IF(mo.cumsum(flag) < 1.5, 1, 0)`,
  `mo.IF(on, opening + mo.cumsum(flow), 0)`, `mo.recurrence(...) * 2`),
  chained on the ENCLOSING row's previous cell in the workbook, so the
  book computed other numbers than Python (`[200, 440, 968]` for
  `[200, 220, 242]`). A running total there is now the `SUM` of its
  source's cells so far; any other roll-forward writes its values; the
  `opening + cumsum(flow)` fold applies only when it is the whole cell.
- `mo.lag` of an unnamed expression (`mo.lag(a.x + a.y) * 2`) wrote its
  fill (`=0 * 2`) in every period of the workbook while Python had the
  right numbers. The cell now holds the expression one period back, as
  one operand (`=(B2 + B3) * 2`); a base
  that reads the row being written (a running total, a recurrence)
  writes its values.
- A projected book (`m.at('quarter').to_excel(...)`) wrote every scalar
  as a reference to its own cell (`=B7` in `B7`), a circular error, and
  every line with a re-grain rule as a range over its own quarterly row
  (`=SUM(B2:D2)` in `B2`). A constant is now its number, a formula over
  constants stays a formula, and a value re-grained from months — a
  quarter's sum, a period-end balance, a ratio of sums, a scalar `SUM` or
  `NPV` over series — is its number.
- Sheet names: an apostrophe inside a name broke every reference to it
  (`'Bob's time'!B2`) and is now doubled (`'Bob''s time'!B2`); a name
  may no longer begin or end with one (Excel refuses the tab); two
  names differing only in case (`Sales`, `SALES`) are two tabs whose
  formulas read the right one — Excel compares names case-insensitively,
  so the second tab used to be renamed behind the formulas and they read
  the first; a sheet written under a changed name (`P/L` as `P_L`) keeps
  its total columns and row outline.
- A `recurrence` template substituted its names inside the sheet names it
  had just written (a variable `s` beside a sheet `Bob's data`), and
  dropped the parentheses of a compound operand (`{prev} * {g}` with
  `g = x + y` became `prev * x + y`).
- A list operand (`x + [1.0, 2.0, 3.0]`) was written as Python text; it
  is now one item per period.

### Known issues

- Tracks (`actual=` / `plan=`) inside a loop are not supported yet.
- Inside a loop container a misspelt row is reported at the end of the
  `with` block (with the line that read it), not at the read; and
  `hasattr(d, name)` is `True` there for any public name.

## [0.2.0] — 2026-09-25

Two eras since 0.1.3. **Time and tracks**: the engine grows a native
time axis and the first finite value axis. **The printed book**: Excel
emission learns styling, totals columns, and formulas that match the
values Python computes.

### Upgrading from 0.1.3

Code written for 0.1.3 and the workbooks it produces can behave
differently.

**API**

- `python_name` is read-only: `set_python_name()` is gone and
  assigning `.python_name` raises `AttributeError`. Name a model's root
  with `mo.Model('name')`.
- `mo.Variable(...)` takes a new `description` parameter right after
  `display_name`, so `formula`, `pyformula`, `unit`, `keys` and
  `excel_props` moved one position. Pass them by keyword.
- `mo.MultiVariable(...)` raises `TypeError` on keyword arguments that
  are neither components nor settings (0.1.3 accepted them silently).
  This includes `row=` / `col=`, which 0.1.3 used to place a section:
  pass `excel_props={'row': N, 'col': N}` instead (0.1.3 ignored these
  two keys in `excel_props`). A `MultiVariableClass` no longer reads
  `row=` / `col=` as placement; they are ordinary parameters, and
  `excel_props={'row': N, 'col': N}` places it too.
- `code`, `description`, `excel_props`, `python_name`, `path`, `id`,
  `sheet_name` and `component_names` are read-only properties of every
  node, so they cannot name a component: `mv.code = mo.Variable(...)`
  and `mo.MultiVariable(..., code=...)` raise `AttributeError` and add
  nothing. As `mo.MultiVariable(...)` keywords, `description`,
  `excel_props`, `excel_layout`, `default_excel_view`, `tracks`,
  `default_grain`, `default_start` and `default_periods` are settings,
  not components.
- `MultiVariable.to_dict()` no longer has a `params` key; it has
  `excel_props` instead.
- `bool(variable)` reflects the value: `Variable(0)`, `Variable(None)`
  and a false comparison (`if seats > 300:`) are falsy. A list-valued
  Variable, even an empty one, raises `ValueError` in a boolean
  context. In 0.1.3 every Variable except an empty list was truthy, so
  such `if` branches always ran. The same applies wherever Python tests
  a Variable's truth: `x in list_of_variables`, `list.index(x)` and
  `list.remove(x)` now compare values (0.1.3 matched the first element
  of any non-empty list); `a or b`, `a and b`, `not a`, `any()` and
  `all()` test the value; and `max()`, `min()`, `sorted()` and chained
  comparisons (`a < b < c`) order by value (0.1.3 returned wrong
  results). Use `is`, or a dict keyed by the Variable (Variables are
  hashable now), for identity.
- Unknown keyword arguments to `mo.Variable(...)` are read as track
  names; 0.1.3 raised `TypeError` at construction. Combined with
  `formula=`, `pyformula=`, `tracks=` or `indexed_by=` they raise
  `ValueError` at once. Otherwise they are checked when the variable
  joins a `mo.Model` or a container that declares `tracks=`: attaching
  it raises `ValueError` unless the name is a declared track. Anywhere
  else the name is not checked, so a misspelled option such as
  `dispaly_name=` is ignored, and arithmetic or a comparison on that
  variable raises `ValueError`.
- Assigning your own attribute on a Variable (`var.note = 'ERP'`) binds
  a track of that name instead of storing an attribute: reading it back
  raises `AttributeError`, and arithmetic on the variable raises
  `ValueError` until it sits under a `mo.Model` that declares the
  track. Keep annotations in `description=`.
- `del mv.x` detaches the component like `mv.remove('x')` (0.1.3 left
  it in place).
- A `MultiVariableClass.compute()` parameter receives exactly what the
  constructor was given, or else its declared default. 0.1.3 took a
  same-named class attribute or node attribute first (`rate = 0.1` on
  the class beat `compute(self, rate=0.05)`; a `display_name`
  parameter got the generated label), so such models can compute
  different values. A parameter named like a read-only node property
  (`code`, `path`, `id`, …) reaches `compute()` but is not stored on the
  instance: `inst.code` returns the node's source.
- A `MultiVariable` passed as a `MultiVariableClass` parameter becomes a
  child of the instance, and the instance's formulas reference that
  child. If the container already sits elsewhere in the model, its rows
  are printed again inside the instance's section as links to the
  originals (0.1.3 referenced the originals). If it sits nowhere else,
  it prints only inside the instance (0.1.3 pasted its values into the
  formulas).
- A custom renderer for `modeleon.compile.Walker` must implement
  `render_regrain` and `render_restrict`. `RenderCtx` has new fields
  (`lossy_inline`, `track_role`, `inlined_value`), all with defaults.
- `from modeleon import *` brings new short names (`up`, `zero`,
  `hold`, `none`, `ratio`, `frozen`, `lag`, `schedule`, `blend`,
  `Model`, `Time`, `DATE`, …) that can shadow your own.
- The notebook view (`_repr_html_`) of a single Variable shows it on a
  sheet of its own with operand values inlined (`=10.0 * 3.0`), instead
  of its row in the model (`=B1 * B2`).

**Values**

- `mo.NPV` follows Excel: the first cash flow is discounted by one
  period, so the value is the old one divided by `1 + rate`. Python and
  the workbook now agree.
- `mo.ROUND` rounds as Excel does: halves away from zero, on the value
  as shown to 15 significant digits (0.1.3 rounded halves to even on the
  binary float): `ROUND(2.5, 0)` is 3, `ROUND(0.125, 2)` is 0.13,
  `ROUND(1.005, 2)` is 1.01, and `ROUND(1250, -2)` is 1300.
- `mo.XIRR` counts a year as 365 days from the first date, as Excel
  does (0.1.3 used 365.25), and a `datetime` counts as its whole day, so
  rates change slightly. A date earlier than the first one now raises
  `ValueError` (Excel returns `#NUM!`), all-zero flows raise instead of
  returning the guess, and very large or very small cash flows converge
  where they used to fail or stop early.
- `None` in arithmetic or a comparison gives `None` (0.1.3 gave
  `'#VALUE!'`, or a plain `True` / `False` for `==` and `!=`), and in a
  list only that period is affected. So `IF(x > 0, a, b)` with an empty
  `x` now returns `b`. The written formulas are unchanged, and Excel
  reads an empty cell as 0, so the workbook can show a number where
  Python has `None`, and an `IF` on an empty input can take the other
  branch: `IF(x <= 0, a, b)` is `b` in Python but `a` in the workbook
  (0.1.3 happened to agree there).
- Dates use Excel's serial-day arithmetic: `date - date` is a number of
  days (an `int` for dates, a `float` once a `datetime` is involved;
  0.1.3 gave a `timedelta`), `date + 90` is a date (0.1.3: `'#VALUE!'`),
  a comparison involving a date compares serial numbers and returns
  `True` / `False` (`date == 45382` can be true), and `date * n` or
  `date / n` is a plain number. A `datetime` keeps its time of day.
  `date ± timedelta` still gives the shifted date, as a `datetime` when
  the `timedelta` has a part-day (0.1.3 dropped it).
- A `timedelta` in any other arithmetic counts as its number of days:
  `td * 2`, `td + td` and `td / 2` are plain numbers (`20.0`, not a
  `timedelta`), `td + 5` is `15.0` (0.1.3: `'#VALUE!'`) and `td == 10`
  is `True`.
- `mo.CONCAT` works element-wise over lists and is written to Excel as
  `&`; whole floats join without `.0`.
- Functions keep the unit shared by their Variable arguments: `IF`,
  `ROUND`, `SUM`, `MIN`, `MAX`, `AVERAGE`, `ABS`, `INT`, `MOD`,
  `mo.cumsum`, `mo.lag` and `@mo.pyformula` functions, but not
  `mo.recurrence` or `NPV` / `PMT` / `FV` / `PV`. `mo.Variable(other)`
  keeps `other`'s unit. These rows gain a unit suffix in their labels,
  e.g. `Total (USD)`. An `@mo.pyformula` function takes that unit even
  when its result is a ratio; clear it with `.unit = None`.
- `mo.recurrence` with a string formula, and `mo.cumsum`, no longer
  raise on a missing value or an error value (such as `'#DIV/0!'`) in
  their input; plain text, or a division by zero inside the formula,
  still raises, and so does a lambda formula. A missing value makes
  that period `None`, and every later period that reads `{prev}`. An
  error value makes that period and every later period `'#VALUE!'`,
  even where the formula does not read `{prev}` (on a tracked line,
  `mo.cumsum` keeps the input's error). The written formulas are
  unchanged, so from that period on the workbook (which reads an empty
  cell as 0) can differ from Python.
- `.formula` spells out an intermediate that is not attached to the
  model instead of showing its generated name: `((a * b)) / 30`, not
  `(v10619ca70) / 30`. An unattached constant still shows a generated
  name.

**Workbooks** — the same model writes a different `.xlsx`:

- Tabs: the first-level children of the emitted root always become
  tabs, and Variables placed directly on the root go to an overview tab
  named after it. `excel_props={'tab': True}` no longer decides which
  containers become tabs: a marked container deeper in the tree prints
  as a section of its tab, and a marked root no longer puts the whole
  model on one sheet. In 0.1.3, once any container was marked, only the
  outermost marked containers were written; Variables on the root and
  unmarked branches were dropped (formulas that used them got their
  values pasted in).
- `excel_props` style keys (`bold`, `bg`, `border_*`, …) take effect,
  and a section's `number_format` applies to the rows inside it.
- The label and value cells of every Variable row carry an explicit
  font from `mo.default_excel_view` (Calibri 11), so they no longer
  follow the workbook's theme font. To keep the workbook default, clear
  it globally with `mo.default_excel_view.font = None`.
- Every sheet laid out across (the default orientation) freezes its
  label column (and any lead columns before the periods), plus its date
  header rows when the sheet has a time axis. Sections become native
  Excel row groups that collapse from their header row. A blank row
  separates a section from the Variable rows above it.
- Date and datetime inputs are written as real spreadsheet dates and a
  `timedelta` input as its number of days (0.1.3 wrote both as text, so
  comparisons in Excel went wrong). A row whose values are all dates
  gets a date number format; the format is chosen per row.

### Added

- **`mo.Model('name')`** — the root of a model; its name fixes every
  node's path. A `mo.Model` assigned into another container becomes an
  ordinary section of it. Every node shares one identity surface:
  `.path`, `.id`, `description=` / `set_description()`,
  `set_display_name()`, `set_style()`.
- **`.code`** — any node renders itself as Modeleon source: values,
  units, display names, `keys`, and formulas whose references stay
  inside the node. Other formulas fall back to their values, and an
  intermediate that is not attached to the model is written as its
  current value (`(a * b) / 30` prints as `(6) / 30`). Not a full round
  trip: descriptions, `excel_props`, the time window, the `mo.Tracks`
  declaration, `indexed_by=`, `regrain=`, `extend=` and `start=` /
  `grain=` are left out, a `MultiVariableClass` instance is written as a
  plain `mo.MultiVariable()`, and a tracked input line is written as
  `TrackValues({...})`, which the snippet cannot run. Functions are
  written without the `mo.` prefix: run the snippet after `import
  modeleon as mo` and `from modeleon import *` (and `import datetime`
  when it holds dates).
- **`indexed_by=[...]`** — a Variable laid out along named axes.
  Arithmetic broadcasts by axis, and the result's shape (`var.shape`)
  follows from its inputs. `mo.lag`, `mo.cumsum`, `mo.recurrence`
  (with `recurrence_sum` and `cohort_retention`), `SUM` / `MAX` /
  `MIN` / `AVERAGE`, `IF` and `.at(grain)` raise `ValueError` on an
  axised operand (see Known issues for the gaps). Only single-axis
  Variables can be written to Excel.
- **`MultiVariable.add()` / `.remove()`**, `.to_excel()` on a single
  Variable, and Variables that are hashable (by identity), so they can
  be dict keys and set members.
- **`MultiVariableClass`**: a `compute()` parameter without a default
  receives the value passed to the constructor (`Cohort(seats=100)`
  calls `compute(seats=100)`; 0.1.3 raised `TypeError`), and
  `description=` sets the instance's description.
- **Functions** `AND`, `OR`, `NOT`, `CHOOSE`, `DATE`, `DAYS360`,
  `ISBLANK`, each written to Excel as the same function. `ISBLANK` is
  true for an empty value and works element-wise over a list. `CHOOSE`
  with a single index over series choices picks per period (the series
  must have the same length, `ValueError` otherwise). `DAYS360` follows
  Excel's US method by default, including a start on the last day of
  February (`DAYS360(2023-02-28, 2023-03-31)` is 30).
- **Ambient time window** — `mo.Model(default_grain=, default_start=,
  default_periods=)` (or on any `MultiVariable`): list-valued lines
  inherit `(start, grain)` from the nearest ancestor that declares one
  (`var.time`). When `default_start` is set, each sheet that holds a
  per-period line gets a header row of period labels (`Jan 2025`,
  `Q1 2025`, `2025`). With `default_periods` set, a list must cover the
  whole window or be a single value, unless it declares `extend=`, has
  its own `start=` / `grain=`, is a `mo.schedule`, is keyed (`keys=` or
  a dict) or is laid out along an axis (`indexed_by=`). `mo.Time(start,
  end or periods=, grain=)` builds a row of period labels.
- **Grain projection** — `model.at('quarter')` (any `MultiVariable`)
  returns the tree re-grained to a coarser grain (`day` → `month` →
  `quarter` → `year`); `var.at('quarter')` re-grains one line by its
  own rule. Rules: `regrain=mo.up('sum'|'last'|'mean'|'first'|'min'|
  'max'|'geometric')` or `mo.up({(from_grain, to_grain): recipe})`
  (`'mean'` is weighted by the days in each period), `mo.frozen(values=
  ...)` for recurrences, and `mo.ratio('num', 'den')` (sum over sum,
  applied by `model.at` only). Under `model.at`, a formula without a
  rule that adds, subtracts or divides lines is recomputed from its
  re-grained inputs; a product of two time-varying lines, a comparison,
  or an `IF` / `MIN` / `MAX` / `ROUND` (and the other element-wise
  functions) is computed at the native grain and summed per bucket.
  `model.at(grain)` raises `ValueError` naming the first input line
  that has no rule, and `var.at(grain)` raises for any line without its
  own rule, formula lines included. `var.set_regrain()` sets a rule
  after construction.
- **`mo.lag(x, periods=1, fill=0)`** — a shift written as a reference
  to the source's cell `periods` periods back; a negative `periods` is
  a lead, and a Variable `fill` stays a live reference.
  **`mo.schedule({'2025-01': 0.15, '2027-07': 0.17})`** — date-keyed
  piecewise constants, filled in when the line joins a model with a
  full window. **`extend=mo.zero()` / `mo.hold()` / `mo.none()`** —
  declared continuation of a partial series.
- **Per-line time windows** — `mo.Variable([...], start='2025-04',
  grain='month')` gives a line its own window (`start=` and `grain=`
  together). In Python, arithmetic between two lines with different
  windows aligns them by date: a side that does not cover the combined
  window must declare `extend=` (otherwise `ValueError`), and lines of
  different grains raise `ValueError`. Excel export does not support
  such lines yet: on a sheet with a dated window (`default_start` set),
  `to_excel` refuses a line whose window differs from the sheet's and
  says how to fix it.
- **Tracks — the finite axis**: parallel named series inside ONE
  Variable. `mo.Tracks('actual', 'budget')` declares user-named tracks
  once per model (names in any language), and
  `mo.Tracks(budget='Budget 2026')` gives a track a display label for
  per-track row labels and `'compare'` headers. Authoring by constructor
  kwargs (`mo.Variable(actual=[...], budget=[...])`), by dotted
  statement (`m.revenue.actual = [...]`), by the pin spelling (a shared
  positional formula plus per-track overrides), or directly with
  `mo.Variable(tracks={'plan': [...], 'actual': [...]})`. Dotted
  coordinate reads (`revenue.actual`) and
  `.at(track=...)` slices are first-class expressions. One formula
  broadcasts across tracks, a mismatch raises, and `lag` / `cumsum` /
  `SUM` / `MAX` / `MIN` / `AVERAGE` / `IF` / `recurrence` work per
  track. Tracks work inside `MultiVariableClass` bodies.
- **The blend line** — `mo.Tracks(..., blend=mo.blend(given=, follow=,
  until=))` adds a synthesized track (named `live`; rename it with
  `name=`) to every tracked line. On a line of typed data it holds the
  given track's values through `until` and the follow track after it,
  each side falling back to the other where it is empty; a formula line
  computes it from its inputs' blended tracks. Recurrence and `cumsum`
  chains on that track continue from its own blended balances (not yet
  inside a `MultiVariableClass` body; see Known issues). `until` is
  optional: without it the given track wins wherever it has a value. A
  blend needs the model's `default_grain`, `default_start` and
  `default_periods`.
- **`mo.ExcelView`** — presentation settings declared on a model or a
  section (`default_excel_view=`), with `mo.default_excel_view` as the
  global default: `orient`, `font`, `cell_types`, `timeline`
  (`header`, `label_format`: `'iso'` / `'finance'` / `'compact'`,
  `totals`, `totals_word`), `bands`, `formats`, `meta`, `nesting`,
  `tracks`. Cell styling (`font`, `cell_types`, `bands`, `formats`, the
  `nesting` indent) is inherited by everything under the declaring
  node; `timeline`, `meta`, `orient` and the `nesting` gap apply to a
  whole tab, so declare them on the model or the tab; `tracks` is read
  from the node you call `to_excel()` on, and `orient='records'` /
  `'columns'` apply to the declaring container only.
- **Tracked lines in Excel** — `ExcelView(tracks=...)`: `'rows'` (the
  default: the display-default row plus one row per other track,
  labeled `name · track`), `'blend'` (the display-default row only), or
  `'compare'` (one row per line; each period becomes a group of
  columns, one per track plus a variance column). The display-default
  row (the blend track, or else the first track the line was written
  with) keeps the line's name, so other formulas reference it. With a
  blend, under `'rows'` and `'compare'` a data line's blended cells are
  live formulas, `=IF(given="", follow, given)` through `until` and
  `=IF(follow="", given, follow)` after it, and such a cell changes
  font color while the given track has a value; under `'blend'` they
  hold values. Other tracks print typed data as values and derived
  tracks as formulas over the same track's rows. The dict form
  `tracks={'mode': ..., 'shown': [...]}` also records a default track
  selection, kept on the view for custom renderers; Excel emission
  ignores it.
- **The financial book** — `ExcelView` styling declared once for the
  whole workbook: `bands={'header': {...}, 'section': {...},
  'subsection': {...}}` (the period header row, and section rows by
  nesting depth), `formats={'number': '#,##0'}` (the default number
  format for value cells), `meta={'article': ..., 'label': ..., 'unit':
  ..., 'constants': ...}` (a row-number column filled from
  `excel_props={'article': '1.1'}`, a units column filled from `unit=`,
  and headers for the name and constants columns),
  `meta={'fields': [('description', 'Comment')]}` (extra lead columns)
  and `nesting={'gap_rows': 0, 'indent': 2}`.
- **Totals columns** — `ExcelView(timeline={'totals': ['quarter',
  'year']})` inserts a total column after each quarter and each year
  (`Jan Feb Mar | Total Q1 | … | Total 2026`; change the word with
  `timeline={'totals_word': 'Gesamt'}`). Lines with a `'sum'`, `'mean'`
  or `'last'` rule get live `SUM`, `AVERAGE` or last-month formulas,
  and a formula line repeats its own formula on the total cells; see
  Known issues for other lines. The header stacks year over quarters over months, and native
  Excel column groups collapse the months into their totals. `SUM`,
  `AVERAGE`, `MIN`, `MAX` and `NPV` over a row that totals columns
  interrupt list the month cells, each with its sheet; `IRR` and
  `XIRR`, which need one range, get the computed value instead of a
  formula.
- **New `excel_props`** — `hidden` (the row, section or tab is written
  in full, formulas intact, but hidden; children inherit it), `row_sum`
  (a live `=SUM` of the row's values in a lead column declared with
  `meta={'fields': [('row_sum', 'Total')]}`), `bg_nonzero` (a
  conditional fill on the row's nonzero numeric cells), `article` (the
  row's number), `header_row` (a Variable's values print on its
  section's header row), `format_by_type` (on a container: colors its
  cells by kind — inputs, formulas, cross-sheet links) and, on
  containers, `orient`.
- **Record layouts** — a container of repeated records (such as
  `MultiVariableClass` instances) declared with `excel_props={'orient':
  'records'}` prints as one table: scalar fields across, records down,
  and one block per per-period field. `orient='columns'` puts the
  records side by side, one row per scalar field; per-period fields are
  not written, and formulas that use them get their values. Without a
  declaration, records print as stacked sections.
- **Formulas match the values** — a printed formula references the
  rows it depends on. When an expression has no equivalent Excel
  formula (for example `mo.cumsum(x) + x` without a row of its own, or
  an aggregate over a list that has no row in the workbook), the cell
  holds the computed value instead. Intermediates built in a Python
  loop and then stored as rows are referenced by their rows instead of
  being expanded into every consumer.
- `to_new_source()` (every node class) and the tracked-line helpers
  `track_expr()`, `shown_track()`, `shown_track_expr()` and
  `display_track_role()` on `Variable`, and
  `MultiVariableClass.__shell__()`, are hooks for renderers and
  tooling, not stable API.

### Fixed

- Excel formulas wrap operands by operator precedence: `a * (1 - b)`
  printed as `=B1 * 1 - B2` and now prints as `=B1 * (1 - B2)`.
- A date inside a formula printed as `2026-01-01`, which Excel reads as
  the number 2024 (so `d - date(2026, 1, 1)` computed a wrong number of
  days), and a `timedelta` as `5 days, 0:00:00`, which Excel cannot
  read; they now print as `DATE(2026, 1, 1)` and `5`.
- An element-wise slice (`x[2:]`, `x[1:] - x[:-1]`) wrote the whole
  range (`=D1:F1`) into every period's cell, giving `#VALUE!` or another
  period's value, and `SUM(x[::2])` summed the whole row; each period
  now references its own cell, and an aggregate over a stepped slice
  lists the cells.
- A row built from a Python list of formulas or references
  (`mo.Variable([x[i] * 2 for i in range(3)])`, `mo.Variable([a, 2,
  b])`) wrote the whole list (`=[B3 * 2, C3 * 2, D3 * 2]`), which Excel
  cannot read, into every cell; each period now gets its own cell
  (`=B3 * 2`, `=C3 * 2`, …; a literal element is written as its value).
- `start + mo.cumsum(flow)` added `start` again in every later period
  (`=B4 + B9 + C5`), so the workbook drifted from the model; later
  periods now read `=B9 + C5`.
- A `mo.recurrence` template written without braces (`'prev * (1 +
  g)'`) printed its text (`=prev * (1 + g)`, a `#NAME?` error); it now
  prints cell references like the `{prev}` form.
- A `mo.recurrence` template with non-ASCII placeholders (`{収入}`)
  raised `ValueError`; it now computes and prints to Excel.
- An aggregate over a list that has no row in the workbook printed only
  its first element (`=SUM(1.0)`), or one element per period inside a
  larger formula (`=SUM(10) + B2`); it now writes the computed value.
- `var.shift(n, fill_value=...)` with a Variable, a string or `None`
  wrote the fill as bare text into the first `n` cells (`=open0`,
  `=n/a`, `=None`; `#NAME?` in Excel); those cells now hold the fill's
  value (a string in quotes, `None` as `0`). Use `mo.lag(x,
  fill=<Variable>)` to keep a live reference.
- Two tabs with the same name: the second one (written as `Name1`)
  started below the first tab's rows, and formulas between the two lost
  their sheet prefix. Each tab now starts at row 1 and cross-tab
  references name the right sheet.
- A tab whose name Excel does not allow (longer than 31 characters, or
  containing `/`, `[`, `]` and the like) was written under a shortened
  name, but formulas referencing it used the original name (`#NAME?`);
  they now use the sheet's written name.
- Sections stored under the same attribute name in different containers
  (`a.rev` and `b.rev`) shared one header position, so a header could
  print below its first row, be overwritten by a row label, or be lost;
  each header now sits above its own rows.
- A marked container inside a tab was written over that tab's rows; it
  now prints as a section.

### Known issues

**Where the workbook can compute a different number than Python**

- A `ROUND` inside a `mo.recurrence` string template
  (`'ROUND({prev} * 1.05, 2)'`) still rounds halves to even in Python,
  while the written `ROUND` rounds them away from zero.
- In totals columns, a formula line whose operand sits on another tab
  that also has totals columns spells that operand's quarter or year as
  one range, which takes in the other tab's total columns, so the total
  is wrong.
- `mo.lag` (or `.shift`) over an expression that has no row of its own
  (`mo.lag(x * 2)`, `mo.lag(mo.cumsum(x))`) writes `=0` in every cell.
  Give the expression its own row first.
- A `mo.recurrence` (or `recurrence_sum`) inside a larger expression
  without a row of its own (`start + mo.recurrence(...)`,
  `mo.recurrence(...) * 2`), or a `mo.cumsum` inside `IF` or `ROUND`,
  repeats the outer operation in every period's formula. Give the
  recurrence or cumsum its own row.
- Without a blend, a tracked line's named row shows the first track it
  was written with, so a formula over lines written in different track
  orders adds different tracks in the workbook. Write every line's
  tracks in the declared order.
- Inside a `MultiVariableClass` body with a blend, `mo.cumsum`,
  `mo.recurrence` and `mo.lag` lines get no blended track, and formulas
  over them can differ between Python and the workbook.
- A `datetime` with a time of day inside a formula is written as
  `DATE(y, m, d)` without the time.
- Lines with their own `start=` on a sheet without a time window are
  written from the first period column; where the window has no
  `default_periods`, a formula that reads past the end of a partial
  series with `extend=` reads the series' first cell; and an expression
  over a line with its own window (`b * 2`, `mo.lag(b)`) does not keep
  that window, so under a `mo.Model` it pairs periods by position.
- `CHOOSE` with a series index over series choices of different lengths
  repeats the last value in Python but not in the workbook.
- `mo.CONCAT`: a literal `None` joins as empty text in Python but is
  written as `0`; a comparison argument is written without parentheses
  (`=B1 & B3 > 3`); a date, boolean or float joins as Python prints it,
  while Excel joins the serial number, `TRUE` and 15 significant
  digits; over lists of different lengths the workbook reads a
  different cell.
- An `IF` whose condition is an error value returns its true branch in
  Python; Excel returns the error. `+` on two text values joins them in
  Python; Excel gives `#VALUE!`.
- `mo.NPV` over a line with an empty period gives `nan` in Python;
  Excel skips the empty cell.
- In totals columns, a missing month inside a bucket is skipped by the
  written `SUM`, while grain projection in Python makes that bucket
  empty, and a `'mean'` line is written as an unweighted `AVERAGE`,
  while Python weights by days.
- Dates from 1900-01-01 to 1900-02-28 compute with serial numbers one
  higher than Excel's (Excel counts a 29 February 1900), and dates
  before 1900 are written as dates, which Excel cannot display.

**Where the workbook shows an error, an empty cell or a fixed number**

- In totals columns, an input line or a recurrence without a
  `regrain=` rule, and any formula that reads it, shows `#VALUE!`, and
  totals of `IF` / `CHOOSE` rows and of lines with a rule other than
  `'sum'`, `'last'` or `'mean'` are written as fixed numbers. In a
  tracked model, the totals cells of formula lines hold a broken
  formula (`=TrackValues(...) * 2`, an error in Excel), and those of
  derived per-track rows show `#VALUE!`.
- `mo.IRR` stops on a fixed tolerance in currency units, so with very
  large cash flows (around 1e12 and up) it can fail to converge where
  Excel returns a rate, and with very small ones it can stop early.
  `mo.XIRR` raises below a rate of -99%, where Excel can still return
  one. `mo.XIRR` over dates given as text (`'2024-01-01'`)
  computes a rate, but the dates are written as text and the workbook
  shows `#VALUE!`.
- A `SUM`, `AVERAGE`, `MIN`, `MAX` or `NPV` over a long row interrupted
  by totals columns, or over a column on an `orient='down'` sheet,
  lists its cells one by one and can exceed Excel's limit of 255
  arguments.
- `IRR` / `XIRR` over a column on a sheet with `orient='down'` are
  written as their computed values, not formulas; such a sheet prints
  no period labels, and its frozen pane is column A instead of the row
  of line names.
- A Python list used directly as an operand (`x + [1, 2, 3]`,
  `mo.IF(c, [1, 2, 3], 0)`, `mo.CHOOSE(i, [1, 2], [3, 4])`) is written
  into the formula as `[1, 2, 3]`, which Excel cannot read; wrap it in
  `mo.Variable([...])`.
- A negative index such as `x[-1]` is written as text Excel cannot
  evaluate; use a non-negative index.
- With `ExcelView(tracks='blend')`, formulas that read a track other
  than the shown one (`rev.budget`, `.at(track=...)`) hold that track's
  values, because its rows are not written.
- `to_excel()` on a section of a tracked model writes no period header
  row on tabs that hold tracked lines; emit the model root.
- `mo.schedule(...)` used in arithmetic before it is attached (`p.b =
  mo.schedule(...) * 2`), or passed as `formula=`, has no value in
  Python (`None`) and writes 0.
- Totals-column cells are bold but do not take the row's `excel_props`
  styling or the view's `font`; `ExcelView(font=...)` styles Variable
  rows only, and `ExcelView(font=None)` on a model does not override
  the global default (set `mo.default_excel_view.font = None`).
- `excel_props={'format_by_type': ...}` on a single Variable and
  `ExcelView(grain=...)` are accepted but have no effect.

**Python-side limitations**

- `mo.up(callable)` is accepted, but re-graining the line raises
  `ValueError` and totals columns do not total it.
- A `mo.schedule` step whose key sorts as text after the window's last
  period label is ignored without an error, even when its date lies
  inside the window (`'2026-02-15'` in a window ending `2026-02`, or an
  unpadded `'2026-3'`).
- `SUM`, `MIN`, `MAX` and `AVERAGE` raise `TypeError` when an argument,
  or a period of one, is `None` (a tracked line with months not yet
  entered included). Unary minus on a date, a tracked line, `None` or
  an error value, and `ABS`, `ROUND`, `INT`, `MOD`, `SUM` and `AVERAGE`
  over dates or `timedelta`s, raise `TypeError`; write `x * -1`.
- `mo.MOD(x, 0)` raises `ZeroDivisionError` and a `mo.recurrence`
  template that divides by zero raises `ValueError`, where Excel shows
  `#DIV/0!`; `mo.MOD` with a list divisor raises `TypeError`.
- `mo.recurrence` does not check a template input passed as a keyword
  (`x=price`) for an axis, and arithmetic between an `indexed_by` line
  and a plain per-period list pairs coordinates with periods without
  an error.
- `.formula` shows a track slice as `x.at(tracks='actual')`; the keyword
  is `track=`.
- `mv.display_name = mo.Variable(...)` adds a component and breaks
  `to_excel()`. A component or `MultiVariableClass` parameter named
  like a node method (`add`, `at`, `remove`, `to_excel`, …) hides that
  method; one named `walk` also makes `to_excel()` of the containing
  model raise `TypeError`.
- `with mo.MultiVariable(...) as mv:` does not add the Variables
  assigned inside the block to `mv`; assign them as attributes
  (`mv.revenue = ...`) or pass them as keywords.

## [0.1.3] — 2026-04-27

Metadata-only release.

- **PyPI Documentation URL** now points to the GitHub README
  (`github.com/modeleonai/modeleon#readme`) instead of the unbuilt
  `modeleon.ai/docs` page. No code or behavior changes.

## [0.1.2] — 2026-04-27

First public release.

(0.1.0 / 0.1.1 existed on PyPI as name-reservation stubs with a
narrow placeholder API. Upgrade directly to 0.1.2 if you have one
of those installed — the API is different.)

### What ships

- **Two primitives.** `Variable` and `MultiVariable`. Operator
  overloading on `Variable` builds a dependency graph; arithmetic,
  comparisons, and method calls all produce live Excel formulas in
  the output, not dead values.
- **Excel output via `model.to_excel(path)`.** Every computed cell
  is a live formula referencing the right cells. Cross-sheet
  references (`assumptions.tax_rate`) resolve as `='Assumptions'!B2`
  in the output.
- **Function library.** `IF`, `SUM`, `MAX`, `MIN`, `AVERAGE`, `ABS`,
  `ROUND`, `INT`, `MOD`, `LEN`, `UPPER`, `LOWER`, `CONCAT`, `YEAR`,
  `MONTH`, `DAY`, `EDATE`, `EOMONTH`, `TODAY`, `IRR`, `NPV`, `XIRR`,
  `PMT`, `FV`, `PV`. Each emits the expected Excel formula
  (`=SUM(B2:F2)`, `=IRR(B2:F2, 0.1)`, etc.).
- **Roll-forward primitives.** `recurrence` for period-over-period
  recurrence (`{prev} * (1 - {churn}) + {new}`), `recurrence_sum`
  for cumulative running totals, `cumsum` for simple running sums.
  *(An earlier version of this entry named these `self_ref` /
  `self_ref_sum` — a pre-release spelling that never shipped; 0.1.2
  already exported `recurrence` / `recurrence_sum`.)*
- **Reusable templates.** `MultiVariableClass` for parameterized
  components (cohorts, schedules) instantiated multiple times.
- **Units with algebra.** `Unit('$')`, `Unit('hr')`, `Unit('$/hr')`
  propagate through arithmetic. Mixed-unit addition raises.
- **JSON serialization.** `mo.to_json(expr)` renders the AST as a
  renderer-neutral dict — useful for diff tools, custom viewers and
  formula inspection.
- **Notebook `_repr_html_`.** Tabbed sheet UI, click-to-reveal
  formulas, trace-precedents on focus, formula/value toggle.

