# Modeleon

[![PyPI](https://img.shields.io/pypi/v/modeleon.svg?color=0d9a9f)](https://pypi.org/project/modeleon/)
[![Python](https://img.shields.io/badge/python-3.12+-0d9a9f.svg)](https://python.org)
[![License](https://img.shields.io/badge/license-Apache_2.0-0d9a9f.svg)](LICENSE)

**Write Python. Ship real Excel formulas.**

> **Variable** — a value.  
> **MultiVariable** — a concept, modeled as a set of Variables and sub-concepts.  
> **Excel** — one cell-traceable snapshot of a MultiVariable.

```python
import modeleon as mo

forecast = mo.MultiVariable("Forecast")
forecast.pnl = mo.MultiVariable("P&L")
with forecast.pnl as pnl:
    pnl.revenue  = mo.Variable(1_000_000, unit="$")
    pnl.cogs_pct = mo.Variable(0.6, display_name="COGS %")
    pnl.cogs     = pnl.revenue * pnl.cogs_pct
    pnl.profit   = pnl.revenue - pnl.cogs

forecast.to_excel("forecast.xlsx")
```

The resulting `forecast.xlsx`:

|   | A         | B          |
|---|-----------|------------|
| 1 | Revenue   | `1000000`  |
| 2 | COGS %    | `0.6`      |
| 3 | Cogs      | `=B1*B2`   |
| 4 | Profit    | `=B1-B3`   |

`B3` is `=B1*B2` — a real Excel formula, not the baked value `600000`.

---

## Install

```bash
pip install modeleon
```

Python 3.12+.

---

## Variable

A value or formula. Arithmetic tracks dependencies automatically.

```python
revenue = mo.Variable(1_000_000)
cogs    = revenue * mo.Variable(0.6)
cogs.value         # 600_000.0
cogs.formula       # 'revenue * 0.6'
cogs.dependencies  # {'revenue'}
```

Scalar, list, or unit-bearing — all the same Variable.

---

## MultiVariable

A concept. Variables are its attributes; sub-MultiVariables are its inner concepts.

```python
forecast = mo.MultiVariable("Forecast")
forecast.assumptions = mo.MultiVariable("Assumptions", tax_rate=mo.Variable(0.25))
forecast.pnl = mo.MultiVariable("P&L")
forecast.pnl.revenue = mo.Variable(1_000_000)
forecast.pnl.taxes   = forecast.pnl.revenue * forecast.assumptions.tax_rate
#                                              ↑ cross-MV reference, fully resolved
```

Reusable concepts subclass `MultiVariableClass`:

```python
class Cohort(mo.MultiVariableClass):
    def compute(self, start, churn, price):
        self.users   = mo.recurrence(start, "{prev} * (1 - {c})", c=churn)
        self.revenue = self.users * price
```

---

## Excel as a snapshot

`.to_excel()` walks the MultiVariable tree and writes a workbook where every Variable is a cell, every formula is `=A1*B2`, every cross-sheet reference is `=Sheet!Cell`.

Standard library functions emit live formulas, not values:

```python
mo.IRR(cf)       # =IRR(B2:F2, 0.1)
mo.NPV(0.1, cf)  # =NPV(0.1, B2:F2)
mo.IF(rev > 0, rev * 0.25, 0)   # =IF(B1>0, B1*0.25, 0)
mo.EOMONTH(start, 1)            # =EOMONTH(B1, 1)
```

---

## License

Apache 2.0. See [LICENSE](LICENSE) and [NOTICE](NOTICE).

[modeleon.ai](https://modeleon.ai) · [PyPI](https://pypi.org/project/modeleon/)
