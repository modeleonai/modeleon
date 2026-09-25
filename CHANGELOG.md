# Changelog

All notable changes to Modeleon will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
  **`mo.schedule({'2025-01': 0.12, '2026-01': 0.16})`** — date-keyed
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
  Variable. `mo.Tracks('факт', 'бюджет')` declares user-named tracks
  once per model, and `mo.Tracks(budget='Budget 2026')` gives a track a
  display label for per-track row labels and `'compare'` headers.
  Authoring by constructor kwargs (`mo.Variable(факт=[...],
  бюджет=[...])`), by dotted statement (`м.выручка.факт = [...]`), by
  the pin spelling (a shared positional formula plus per-track
  overrides), or directly with `mo.Variable(tracks={'plan': [...],
  'actual': [...]})`. Dotted coordinate reads (`выручка.факт`) and
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
  selection, which Excel emission ignores.
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
  `timeline={'totals_word': 'Итого'}`). Lines with a `'sum'`, `'mean'`
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
- A `mo.recurrence` template with non-ASCII placeholders (`{доход}`)
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
- `.formula` shows a track slice as `x.at(tracks='fact')`; the keyword
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
  renderer-neutral dict — useful for diff views, web grids, LLM
  inspection.
- **Notebook `_repr_html_`.** Tabbed sheet UI, click-to-reveal
  formulas, trace-precedents on focus, formula/value toggle.

