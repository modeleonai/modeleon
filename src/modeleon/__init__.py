# SPDX-License-Identifier: Apache-2.0
"""
Modeleon - Financial Modeling Engine

Python code compiles to a hierarchical graph, then to live Excel formulas.

Stability policy (0.x). The top-level API re-exported from
``modeleon`` — :class:`Variable`, :class:`MultiVariable`, the shipped
functions, :func:`to_json` — is stable within minor releases.
Attributes prefixed with a single underscore (``_value``, ``_expr``,
``_source_code``, ``_set_expr``) are internal storage; use the public
properties (``.value``, ``.expr``, ``.source_code``, ``.formula``)
instead. Extension authors writing custom helpers or renderers should
import from :mod:`modeleon.extend`, which stabilizes that surface.

Core primitives (just two):

- ``Variable`` — a single value or formula. Operator overloading tracks
  dependencies.
- ``MultiVariable`` — a container grouping Variables and nested
  MultiVariables. Pass ``excel_props={'tab': True}`` to mark it as an Excel tab;
  omit ``role`` for a plain grouping MV. First-depth sub-MultiVariables
  of whatever you emit become tabs automatically when no explicit
  ``excel_props={'tab': True}`` markers exist.

Typical usage::

    import modeleon as mo

    model = mo.MultiVariable("Acme Forecast")
    model.income = mo.MultiVariable("Income Statement", excel_props={'tab': True})
    model.income.revenue = mo.Variable(1_000_000, display_name='Revenue')
    model.income.cogs = model.income.revenue * mo.Variable(0.6, display_name='COGS %')
    model.income.gross_profit = model.income.revenue - model.income.cogs

    model.to_excel("forecast.xlsx")

Every user script builds a MultiVariable tree and calls ``.to_excel()``
on its root. No module-level emit function — explicitness over magic.

Multi-period structure is a plain list-valued Variable —
``mo.Variable([1.0, 1.1, 1.2], display_name='Growth')`` — which
``recurrence`` expands across the horizon.

Package layout:
    core/       - Variable, MultiVariable, QPath, Unit, Expr AST
    functions/  - IF, SUM, MAX, MIN, AVERAGE, ABS, ROUND, INT, MOD,
                  recurrence, IRR, NPV, XIRR, PMT, FV, PV, …
    compile/    - AST renderers + layout + .xlsx writer (Excel, JSON)
    display/    - notebook HTML rendering
    plugins.py  - extension point registry
"""

__version__ = "0.1.2"

# Plugin system — extensions register via entry points
from modeleon.plugins import load_plugins as _load_plugins

# Core DSL
from modeleon.core.variable import Variable
from modeleon.core.multi_variable import (
    MultiVariableClass,
    MultiVariable,
)
from modeleon.core.errors import CircularDependencyError, CrossScopeReferenceWarning
from modeleon.core.mv_context import ModelStructureWarning
from modeleon.core.unit import Unit

# AST backends (Excel is implicit via ``model.to_excel(path)``; others
# available for introspection and alternative renderers)
from modeleon.compile import check_compat, to_json

# Pre-built DSL helpers
from modeleon.functions import (
    val,
    pyformula,
    # Aggregates
    SUM, MAX, MIN, AVERAGE,
    # Math
    ABS, ROUND, INT, MOD,
    # Conditional
    IF,
    # Dates
    YEAR, MONTH, DAY, EDATE, EOMONTH, TODAY,
    # Text
    LEN, UPPER, LOWER, CONCAT,
    # Recurrence / cohort
    cumsum, recurrence, recurrence_sum, cohort_retention,
    # Financial
    IRR, NPV, XIRR, PMT, FV, PV,
)

_load_plugins()

__all__ = [
    # Core DSL
    "Variable",
    "MultiVariable",
    "MultiVariableClass",
    "Unit",
    # Exceptions / warnings
    "CircularDependencyError",
    "CrossScopeReferenceWarning",
    "ModelStructureWarning",
    # AST introspection (alternative renderer — see compile/json_backend)
    "to_json",
    # Backend compatibility inspection (see compile/compat)
    "check_compat",
    # Excel-cased aggregates and helpers (match =SUM(...)=MAX(...) in output)
    "IF",
    "SUM",
    "MAX",
    "MIN",
    "AVERAGE",
    "ABS",
    "ROUND",
    "INT",
    "MOD",
    # Dates
    "YEAR",
    "MONTH",
    "DAY",
    "EDATE",
    "EOMONTH",
    "TODAY",
    # Text
    "LEN",
    "UPPER",
    "LOWER",
    "CONCAT",
    # Domain helpers (no Excel equivalent — stay lowercase)
    "cumsum",
    "cohort_retention",
    "recurrence",
    "recurrence_sum",
    "val",
    "pyformula",
    # Financial
    "IRR",
    "NPV",
    "XIRR",
    "PMT",
    "FV",
    "PV",
]
