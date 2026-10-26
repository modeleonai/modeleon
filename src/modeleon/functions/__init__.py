# SPDX-License-Identifier: Apache-2.0
"""Pre-built DSL functions.

Organized by topic:

- :mod:`aggregate`    — ``SUM``, ``MAX``, ``MIN``, ``AVERAGE``
- :mod:`criteria`     — ``SUMIF``, ``SUMIFS``, ``COUNTIF``, ``COUNTIFS``,
  ``AVERAGEIF``, ``AVERAGEIFS``, ``MAXIFS``, ``MINIFS``, ``COUNT``,
  ``COUNTA``
- :mod:`mathfn`       — ``ABS``, ``ROUND``, ``INT``, ``MOD``, ``ROUNDUP``,
  ``ROUNDDOWN``, ``TRUNC``, ``CEILING``, ``FLOOR``, ``POWER``, ``SQRT``,
  ``EXP``, ``LN``, ``LOG``, ``LOG10``, ``SIGN``, ``PRODUCT``
- :mod:`conditional`  — ``IF``, ``AND``, ``OR``, ``NOT``, ``CHOOSE``,
  ``ISBLANK``
- :mod:`logical`      — ``IFERROR``, ``ISERROR``, ``ISNUMBER``, ``XOR``,
  ``IFS``, ``SWITCH``
- :mod:`lookup`       — ``MATCH``, ``INDEX``, ``XLOOKUP``, ``VLOOKUP``,
  ``HLOOKUP``
- :mod:`dates`        — ``YEAR``, ``MONTH``, ``DAY``, ``DATE``, ``EDATE``, ``EOMONTH``, ``DAYS360``,
  ``YEARFRAC``, ``DAYS``, ``TODAY``, ``DATEDIF``, ``WEEKDAY``, ``WEEKNUM``,
  ``NETWORKDAYS``, ``WORKDAY``
- :mod:`text`         — ``LEN``, ``UPPER``, ``LOWER``, ``CONCAT``, ``LEFT``,
  ``RIGHT``, ``MID``, ``TRIM``, ``FIND``, ``SEARCH``, ``SUBSTITUTE``,
  ``VALUE``, ``TEXT``, ``REPT``, ``TEXTJOIN``
- :mod:`recurrence`   — ``recurrence``, ``recurrence_sum``, ``cumsum``
- :mod:`lag`          — ``lag``
- :mod:`schedule`     — ``schedule``
- :mod:`cohort`       — ``cohort_retention``
- :mod:`financial`    — ``IRR``, ``NPV``, ``XIRR``, ``XNPV``, ``PMT``, ``FV``, ``PV``,
  ``NPER``, ``IPMT``, ``PPMT``, ``RATE``, ``CUMIPMT``, ``CUMPRINC``,
  ``SLN``, ``DB``, ``DDB``, ``SYD``
- :mod:`statistics`   — ``MEDIAN``, ``STDEV_S``, ``STDEV_P``, ``VAR_S``, ``VAR_P``,
  ``PERCENTILE_INC``, ``PERCENTILE_EXC``, ``QUARTILE_INC``, ``QUARTILE_EXC``,
  ``LARGE``, ``SMALL``, ``RANK_EQ``, ``RANK_AVG``, and the older names
  ``STDEV``, ``STDEVP``, ``VAR``, ``VARP``, ``PERCENTILE``, ``QUARTILE``, ``RANK``

How each name is shaped — one call per period, a reduction, a call
over ranges — and how the workbook writes it is the function
registry's to say (:mod:`registry`); every function registers itself
from its own module. Every aggregate — ``SUM``, ``AVERAGE``, ``MIN``,
``MAX``, ``PRODUCT``, ``COUNT``, ``COUNTA``, the statistics — follows one
rule: several series are combined period by period, as ``a + b``
combines them (``SUM(a, b)`` is ``=SUM(B2, B3)`` in every column); one
series alone is reduced over its periods (``SUM(a)`` is ``=SUM(B2:F2)``);
scalars only stay one number.

The :class:`Unit` type lives in :mod:`modeleon.core.unit` — it's a
primitive used by :class:`Variable` itself, not a formula helper.

Backend compatibility convention. Most helpers build a plain
:class:`~modeleon.core.expr.FuncCall` with no backend hint —
every renderer emits them natively (``SUM`` → ``=SUM(...)`` in Excel,
a ``funccall`` node named ``SUM`` in JSON).
Functions that are native to a specific target — the Excel
financial helpers (``IRR``, ``NPV``, ``XIRR``, ``PMT``, ``FV``, ``PV``)
and the Excel date helpers (``EDATE``, ``EOMONTH``, ``YEARFRAC``,
``DAYS``, ``DATEDIF``, ``WEEKDAY``, ``WEEKNUM``, ``NETWORKDAYS``,
``WORKDAY``) — pass ``render_backends=frozenset({'excel'})`` through
:func:`._helpers.make_func_var` / ``_wrap_with_formula``. Other
renderers fall back to inlining the Python-computed ``_value``.
Inspect compat for a given target with :func:`modeleon.check_compat`.
"""

from ._helpers import pyformula, val
from .aggregate import AVERAGE, MAX, MIN, SUM
from .cohort import cohort_retention
from .conditional import AND, CHOOSE, IF, ISBLANK, NOT, OR
from .criteria import (
    AVERAGEIF, AVERAGEIFS, COUNT, COUNTA, COUNTIF, COUNTIFS, MAXIFS, MINIFS, SUMIF, SUMIFS,
)
from .dates import (
    DATE, DATEDIF, DAY, DAYS, DAYS360, EDATE, EOMONTH, MONTH, NETWORKDAYS, TODAY, WEEKDAY,
    WEEKNUM, WORKDAY, YEAR, YEARFRAC,
)
from .financial import (
    CUMIPMT, CUMPRINC, DB, DDB, FV, IPMT, IRR, NPER, NPV, PMT, PPMT, PV, RATE, SLN, SYD, XIRR,
    XNPV,
)
from .logical import IFERROR, IFS, ISERROR, ISNUMBER, SWITCH, XOR
from .lookup import HLOOKUP, INDEX, MATCH, VLOOKUP, XLOOKUP
from .mathfn import (
    ABS, CEILING, EXP, FLOOR, INT, LN, LOG, LOG10, MOD, POWER, PRODUCT, ROUND, ROUNDDOWN,
    ROUNDUP, SIGN, SQRT, TRUNC,
)
from .lag import lag
from .schedule import schedule
from .recurrence import cumsum, recurrence, recurrence_sum
from .statistics import (
    LARGE, MEDIAN, PERCENTILE, PERCENTILE_EXC, PERCENTILE_INC, QUARTILE, QUARTILE_EXC,
    QUARTILE_INC, RANK, RANK_AVG, RANK_EQ, SMALL, STDEV, STDEV_P, STDEV_S, STDEVP, VAR,
    VAR_P, VAR_S, VARP,
)
from .text import (
    CONCAT, FIND, LEFT, LEN, LOWER, MID, REPT, RIGHT, SEARCH, SUBSTITUTE, TEXT, TEXTJOIN, TRIM,
    UPPER, VALUE,
)


__all__ = [
    # Aggregates
    "SUM", "MAX", "MIN", "AVERAGE",
    # Conditional aggregates and counts
    "SUMIF", "SUMIFS", "COUNTIF", "COUNTIFS", "AVERAGEIF", "AVERAGEIFS", "MAXIFS", "MINIFS",
    "COUNT", "COUNTA",
    # Math
    "ABS", "ROUND", "INT", "MOD", "ROUNDUP", "ROUNDDOWN", "TRUNC", "CEILING", "FLOOR",
    "POWER", "SQRT", "EXP", "LN", "LOG", "LOG10", "SIGN", "PRODUCT",
    # Conditional / logical
    "IF", "AND", "OR", "NOT", "CHOOSE", "ISBLANK",
    "IFERROR", "ISERROR", "ISNUMBER", "XOR", "IFS", "SWITCH",
    # Lookup / reference
    "MATCH", "INDEX", "XLOOKUP", "VLOOKUP", "HLOOKUP",
    # Dates
    "YEAR", "MONTH", "DAY", "DATE", "EDATE", "EOMONTH", "DAYS360", "TODAY",
    "YEARFRAC", "DAYS", "DATEDIF", "WEEKDAY", "WEEKNUM", "NETWORKDAYS", "WORKDAY",
    # Text
    "LEN", "UPPER", "LOWER", "CONCAT",
    "LEFT", "RIGHT", "MID", "TRIM", "FIND", "SEARCH", "SUBSTITUTE", "VALUE", "TEXT", "REPT",
    "TEXTJOIN",
    # Financial
    "IRR", "NPV", "XIRR", "XNPV", "PMT", "FV", "PV",
    "NPER", "IPMT", "PPMT", "RATE", "CUMIPMT", "CUMPRINC", "SLN", "DB", "DDB", "SYD",
    # Statistics
    "MEDIAN", "STDEV_S", "STDEV_P", "VAR_S", "VAR_P",
    "PERCENTILE_INC", "PERCENTILE_EXC", "QUARTILE_INC", "QUARTILE_EXC",
    "LARGE", "SMALL", "RANK_EQ", "RANK_AVG",
    "STDEV", "STDEVP", "VAR", "VARP", "PERCENTILE", "QUARTILE", "RANK",
    # Domain helpers (no Excel equivalent — stay lowercase)
    "cumsum", "lag",
    "schedule", "recurrence", "recurrence_sum", "cohort_retention",
    "val", "pyformula",
]
