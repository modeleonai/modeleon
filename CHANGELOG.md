# Changelog

All notable changes to Modeleon will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

The time-and-tracks era — the engine grows a native time axis and the
first finite value axis.

### Added

- **Ambient time window** — `mo.Model(default_grain=, default_start=,
  default_periods=)`: list-valued lines inherit their native
  `(start, grain)` from the nearest ancestor; date headers derive from
  the window. `mo.Time` axis variable.
- **Grain projection** — `var.at('quarter')`, `project_model(...)`:
  re-grain by per-line rules (`regrain=mo.up('sum'|'last'|'mean')`,
  `mo.frozen`, `mo.ratio`), evaluate-then-regrain for formulas.
- **Tracks — the finite axis** (§ the multi-axis design): parallel
  role-keyed series inside ONE Variable. `mo.Tracks('факт', 'бюджет')`
  declares user-named tracks once per model; authoring by constructor
  kwargs (`mo.Variable(факт=[...], бюджет=[...])`), by dotted statement
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
  quiet absence, never errors.
- **`mo.lag`**, **`mo.schedule`** (date-keyed piecewise constants),
  **`extend=mo.zero()/mo.hold()/mo.none()`** (declared continuation of
  partial series), date-aligned arithmetic across differing windows.
- **ExcelView `tracks=` / `grain=`** — the printed form: `'blend'`
  (display-default rows only), `'rows'` (labeled per-track rows);
  a declared view grain.
- **Excel emission of tracked lines** — the display-default (live) row
  keeps the line's name so formulas reference it; other tracks emit as
  labeled value rows.

### Fixed

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
- **Roll-forward primitives.** `self_ref` for period-over-period
  recurrence (`{prev} * (1 - {churn}) + {new}`), `self_ref_sum`
  for cumulative running totals, `cumsum` for simple running sums.
- **Reusable templates.** `MultiVariableClass` for parameterized
  components (cohorts, schedules) instantiated multiple times.
- **Units with algebra.** `Unit('$')`, `Unit('hr')`, `Unit('$/hr')`
  propagate through arithmetic. Mixed-unit addition raises.
- **JSON serialization.** `mo.to_json(expr)` renders the AST as a
  renderer-neutral dict — useful for diff views, web grids, LLM
  inspection.
- **Notebook `_repr_html_`.** Tabbed sheet UI, click-to-reveal
  formulas, trace-precedents on focus, formula/value toggle.

