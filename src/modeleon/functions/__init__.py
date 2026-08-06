# SPDX-License-Identifier: Apache-2.0
"""Pre-built DSL functions.

Organized by topic:

- :mod:`aggregate`    — ``SUM``, ``MAX``, ``MIN``, ``AVERAGE``
- :mod:`mathfn`       — ``ABS``, ``ROUND``, ``INT``, ``MOD``
- :mod:`conditional`  — ``IF``, ``AND``, ``OR``, ``NOT``, ``CHOOSE``
- :mod:`dates`        — ``YEAR``, ``MONTH``, ``DAY``, ``DATE``, ``EDATE``, ``EOMONTH``, ``DAYS360``, ``TODAY``
- :mod:`text`         — ``LEN``, ``UPPER``, ``LOWER``, ``CONCAT``
- :mod:`recurrence`   — ``recurrence``, ``recurrence_sum``, ``cumsum``
- :mod:`cohort`       — ``cohort_retention``
- :mod:`financial`    — ``IRR``, ``NPV``, ``XIRR``, ``PMT``, ``FV``, ``PV``

The :class:`Unit` type lives in :mod:`modeleon.core.unit` — it's a
primitive used by :class:`Variable` itself, not a formula helper.

Backend compatibility convention. Most helpers build a plain
:class:`~modeleon.core.expr.FuncCall` with no backend hint —
every renderer emits them natively (``SUM`` → ``=SUM(...)`` in Excel,
``SUM(...)`` in JSON, ``df.sum()`` in a future pandas renderer).
Functions that are native to a specific target — today just the Excel
financial helpers (``IRR``, ``NPV``, ``XIRR``, ``PMT``, ``FV``, ``PV``)
and the Excel date shift helpers (``EDATE``, ``EOMONTH``) — pass
``render_backends=frozenset({'excel'})`` through
:func:`._helpers.make_func_var` / ``_wrap_with_formula``. Other
renderers fall back to inlining the Python-computed ``_value``.
Inspect compat for a given target with :func:`modeleon.check_compat`.
"""

from ._helpers import pyformula, val
from .aggregate import AVERAGE, MAX, MIN, SUM
from .cohort import cohort_retention
from .conditional import AND, CHOOSE, IF, NOT, OR
from .dates import DATE, DAY, DAYS360, EDATE, EOMONTH, MONTH, TODAY, YEAR
from .financial import FV, IRR, NPV, PMT, PV, XIRR
from .mathfn import ABS, INT, MOD, ROUND
from .lag import lag
from .schedule import schedule
from .recurrence import cumsum, recurrence, recurrence_sum
from .text import CONCAT, LEN, LOWER, UPPER


__all__ = [
    # Aggregates
    "SUM", "MAX", "MIN", "AVERAGE",
    # Math
    "ABS", "ROUND", "INT", "MOD",
    # Conditional / logical
    "IF", "AND", "OR", "NOT", "CHOOSE",
    # Dates
    "YEAR", "MONTH", "DAY", "DATE", "EDATE", "EOMONTH", "DAYS360", "TODAY",
    # Text
    "LEN", "UPPER", "LOWER", "CONCAT",
    # Financial
    "IRR", "NPV", "XIRR", "PMT", "FV", "PV",
    # Domain helpers (no Excel equivalent — stay lowercase)
    "cumsum", "lag",
    "schedule", "recurrence", "recurrence_sum", "cohort_retention",
    "val", "pyformula",
]
