# Changelog

All notable changes to Modeleon will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

Two eras since 0.1.3. **Time and tracks**: the engine grows a native
time axis and the first finite value axis. **The printed book**: Excel
emission learns styling, hierarchical totals, and honest formulas.

### Added

- **Ambient time window** — `mo.Model(default_grain=, default_start=,
  default_periods=)`: list-valued lines inherit their native
  `(start, grain)` from the nearest ancestor; date headers derive from
  the window. `mo.Time` axis variable.
- **Grain projection** — `var.at('quarter')`, `project_model(...)`:
  re-grain by per-line rules (`regrain=mo.up('sum'|'last'|'mean')`,
  `mo.frozen`, `mo.ratio`), evaluate-then-regrain for formulas.
- **Tracks — the finite axis**: parallel role-keyed series inside ONE
  Variable. `mo.Tracks('факт', 'бюджет')` declares user-named tracks
  once per model; authoring by constructor kwargs
  (`mo.Variable(факт=[...], бюджет=[...])`), by dotted statement
  (`м.выручка.факт = [...]`), or the pin spelling (a shared positional
  formula plus per-track overrides — given beats derived). Dotted
  coordinate read (`выручка.факт`) and `.at(track=...)` slices as
  first-class AST nodes. One formula broadcasts across tracks; the
  mismatch law is loud; rank-lifting carries lag / cumsum / SUM / MAX /
  MIN / AVERAGE / IF / recurrence per track. Tracks work inside
  `MultiVariableClass` bodies (provisional materialization).
- **The blend (live) line** — `mo.Tracks(..., blend=mo.blend(given=,
  follow=, until=))` synthesizes a spliced series: given-track data up
  to the boundary, the follow track after; recurrence chains re-anchor
  from actual closing balances; holes (un-entered months) propagate as
  quiet absence, never errors. `until` is optional — without it the
  splice is open-ended.
- **`mo.lag`**, **`mo.schedule`** (date-keyed piecewise constants),
  **`extend=mo.zero()/mo.hold()/mo.none()`** (declared continuation of
  partial series), date-aligned arithmetic across differing windows.
- **ExcelView `tracks=` / `grain=`** — the printed form: `'blend'`
  (display-default rows only), `'rows'` (labeled per-track rows),
  `'compare'` (per-track column groups side by side); a declared view
  grain. The dict form `ExcelView(tracks={'mode': ..., 'shown': [...]})`
  keeps `mode` as the emission setting and records `shown`, the default
  track selection for a viewer (presentation metadata — Excel emission
  ignores it).
- **Excel emission of tracked lines** — the display-default (live) row
  keeps the line's name so formulas reference it; other tracks emit as
  labeled value rows.
- **The financial book** — a declarative style cascade on
  `ExcelView`: `bands` (masthead header, section bands painted by
  nesting depth), number formats, a units column declared once
  (`ед.изм.`), № and name column headers, a header for the constants
  column. One declaration styles the whole workbook.
- **Hierarchical totals columns** — subtotal columns woven into the
  time axis (`янв фев мар (Q1) … (год)`) as live formulas; the totals
  header stacks year over quarters over months; native Excel column
  groups fold the coarser tiers away; a view's `grain` accepts a
  multi-select (finest grain = columns, coarser = «Итого» columns).
- **`ISBLANK`** — an optional value can say it is empty.
- **Presentation `excel_props`** — `hidden` (service rows hidden, not
  skipped), `row_sum` («Итого» lead column), `bg_nonzero`
  (value-driven fill).
- **Repeating records print as one sheet** — sibling instances share
  one table instead of a tab per record; `orient='columns'` lays them
  out side by side; dense nesting derives hierarchy from indentation;
  a section's subtotal sits on its own header row.
- **Formulas tell the truth** — the chain law: a printed formula
  references the rows it actually depends on; when a chain can't be
  expressed losslessly the cell falls back to its value explicitly
  instead of printing a wrong formula. Loop-built intermediates
  reference their rows rather than unrolling into consumers.
- **Per-line time windows** — the time axis belongs to the line, not
  only the model; lines with differing windows align by date.

### Fixed

- **Track-authored expressions emit live formulas.**
  `mo.Variable(plan=<row>)` / `forecast=mo.recurrence(...)` kept only
  the track's numbers — the workbook showed dead literals. The
  expansion now recovers the expression for the shown row and every
  per-track subrow. Coordinate law: a track formula resolves tracked
  operands to their SAME-track row (subrow if laid out; the head only
  when it is that track) and degrades one operand to its value
  otherwise — never to the blend row, which computes a different
  number. `x.at(track=)` renders through the same rule.
- **A track authored as a list of references keeps them as
  formulas.** `forecast=[fact, None, …, target_cell]` is a sparse row
  of links: the track's expression is a reference per period where the
  list holds a Variable and a literal elsewhere, so those periods
  become live `=B..` links instead of pasted numbers.
- A tracked `mo.cumsum` printed the Python spelling (`=B3.cumsum()`)
  into every cell when tracks were laid out as columns, instead of the
  live running chain. Now: seed reads the input's cell, later periods
  read the row's own previous-period cell (same track — the per-track
  address list absorbs the column stride) plus the input's current
  cell; an inlined cumsum with no cell of its own falls back to its
  value, never to a formula Excel cannot evaluate.
- A section header that followed Variable rows was glued to the last
  row of the previous run; the walk only left a gap AFTER a section.
  The pre-gap mirrors the view's `nesting.gap_rows`, so dense books
  (`gap_rows=0`) stay untouched.
- A regrain-ruled line now projects its stored tracks, not just the
  live row.
- A data hole inside an aggregation bucket makes the bucket a hole
  instead of a silent partial sum.
- A range formula never swallows an adjacent totals column; a
  bucket's total names itself on the bucket's own tier.
- Recurrence template placeholders and the Excel SelfRef renderer were
  ASCII-only — Cyrillic placeholders (`{доход}`) silently failed to
  substitute. Both are Unicode now.
- Recurrence chains absorb error markers and data holes instead of
  killing the run; `cumsum` carries holes as holes.

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

