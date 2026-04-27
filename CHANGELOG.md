# Changelog

All notable changes to Modeleon will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.2] — 2026-04-26

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

