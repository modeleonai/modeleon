# Modeleon Examples

Runnable financial models showing common patterns. Each example produces a real `.xlsx` file with live Excel formulas.

Run any example:

```bash
python examples/01_income_statement.py
```

## Examples

| File | What it shows |
|------|---------------|
| [01_income_statement.py](01_income_statement.py) | Simple single-sheet model: revenue, COGS, gross profit |
| [02_multi_sheet.py](02_multi_sheet.py) | Multiple tabs with cross-sheet references (Assumptions ↔ P&L) |
| [06_irr_npv.py](06_irr_npv.py) | Financial functions: IRR and NPV for cash-flow analysis |

Every example ends with `mo.to_excel("output.xlsx")` — open the file in Excel to see live formulas.

Time-series examples (`Timeline`, `TimeVariable`, `TimeDependentVariable`)
ship with the `modeleon-pro` package.
