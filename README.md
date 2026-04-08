# Modeleon

[![License](https://img.shields.io/badge/License-Apache_2.0-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/Python-3.12+-green.svg)](https://python.org)
[![PyPI](https://img.shields.io/pypi/v/modeleon.svg)](https://pypi.org/project/modeleon/)

**Financial Model Engineering — Python code compiles to live Excel formulas.**

Write financial models in Python. Get real, auditable Excel formulas — not dead values.

```python
from modeleon import Model

model = Model("Revenue Forecast")

price = model.var("Price", 50)
volume = model.var("Volume", 1000)
revenue = model.var("Revenue", price * volume)

model.to_excel("forecast.xlsx")
```

`forecast.xlsx` — real formulas, not values:

|   | A | B |
|---|---|---|
| 1 | Price | 50 |
| 2 | Volume | 1000 |
| 3 | Revenue | `=B1*B2` |

Every cell is a live formula you can audit, extend, and trust.

> **Coming soon.** Star this repo to follow progress.

[modeleon.ai](https://modeleon.ai) | [PyPI](https://pypi.org/project/modeleon/)

## License

Apache 2.0
