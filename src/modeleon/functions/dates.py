# SPDX-License-Identifier: Apache-2.0
"""Date helpers — ``YEAR``, ``MONTH``, ``DAY``, ``EDATE``, ``EOMONTH``, ``TODAY``.

**Invariant:** every call returns a :class:`Variable` with a tracked
formula — even when inputs are plain ``date`` objects or ISO strings. The
Excel cell is the product (``=EOMONTH(...)``), not just the value.
"""

from __future__ import annotations

from calendar import monthrange
from datetime import date, datetime
from typing import Any, Callable, Union

from dateutil.relativedelta import relativedelta

from ..core.variable import Variable
from ._helpers import make_func_var, operand, src

DateOperand = Union[Variable, date, datetime, str]


_EXCEL_ONLY = frozenset({'excel'})
"""Render-backend set for Excel-native date helpers (``EDATE``,
``EOMONTH``). Other renderers fall back to inlining ``_value``. The
simple accessors (``YEAR`` / ``MONTH`` / ``DAY`` / ``TODAY``) stay
universal — every render target has some equivalent primitive."""


def _extract_date(x):
    """Return a ``date`` from a Variable, date, datetime, or ISO string."""
    if isinstance(x, Variable):
        return _extract_date(x._value)
    if isinstance(x, datetime):
        return x.date()
    if isinstance(x, date):
        return x
    if isinstance(x, str):
        return datetime.strptime(x, '%Y-%m-%d').date()
    raise TypeError(
        f"Expected a date, datetime, ISO string, or Variable — got {type(x).__name__}."
    )


def _broadcast(fn: Callable[..., Any], *raws: Any) -> tuple[Any, str]:
    """Apply ``fn`` over scalar or list operands, element-wise.

    Scalars repeat; lists shorter than the longest clamp to their last
    element (matches length-1 broadcast elsewhere). Returns
    ``(value, var_type)`` where ``var_type`` is ``'list'`` if any operand
    is a list, else ``'scalar'``.
    """
    lists = [r for r in raws if isinstance(r, list)]
    if not lists:
        return fn(*raws), 'scalar'
    n = max(len(r) for r in lists)

    def at(r: Any, i: int) -> Any:
        if isinstance(r, list):
            return r[i] if i < len(r) else r[-1]
        return r

    return [fn(*[at(r, i) for r in raws]) for i in range(n)], 'list'


def _unary_date_func(
    func_name: str,
    py_fn: Callable[[date], int],
) -> Callable[[DateOperand], Variable]:
    """Build a Variable-aware date helper like YEAR/MONTH/DAY.

    Always returns a Variable, regardless of whether input is a Variable,
    ``date``, ISO string, or ``datetime``.
    """
    def helper(value: DateOperand) -> Variable:
        raw, expr = operand(value)
        if isinstance(raw, list):
            calc: Any = [py_fn(_extract_date(v)) for v in raw]
            var_type = 'list'
        else:
            calc = py_fn(_extract_date(raw))
            var_type = 'scalar'
        return make_func_var(
            func_name, [expr], calc, 'int', var_type,
            source_code=f"{func_name}({src(value)})",
        )

    return helper


YEAR = _unary_date_func("YEAR", lambda d: d.year)
YEAR.__doc__ = "Year of a date (renders ``=YEAR(...)``). Always returns a Variable."

MONTH = _unary_date_func("MONTH", lambda d: d.month)
MONTH.__doc__ = "Month 1-12 of a date (renders ``=MONTH(...)``). Always returns a Variable."

DAY = _unary_date_func("DAY", lambda d: d.day)
DAY.__doc__ = "Day of month 1-31 (renders ``=DAY(...)``). Always returns a Variable."


def EDATE(start_date: DateOperand, months: Union[Variable, int]) -> Variable:
    """Shift a date by ``months`` (renders ``=EDATE(...)``). Always returns a Variable.

    ``EDATE("2025-01-15", 3)`` → Variable with value ``date(2025, 4, 15)`` and
    formula ``EDATE("2025-01-15", 3)``. Negative values go backward.
    """
    raw_start, start_expr = operand(start_date)
    raw_months, months_expr = operand(months)
    m = int(raw_months)

    def shift(d):
        return _extract_date(d) + relativedelta(months=m)

    if isinstance(raw_start, list):
        calc: Any = [shift(v) for v in raw_start]
        var_type = 'list'
    else:
        calc = shift(raw_start)
        var_type = 'scalar'

    return make_func_var(
        "EDATE", [start_expr, months_expr], calc, 'datetime', var_type,
        source_code=f"EDATE({src(start_date)}, {m})",
        render_backends=_EXCEL_ONLY,
    )


def EOMONTH(start_date: DateOperand, months: Union[Variable, int] = 0) -> Variable:
    """End-of-month date after shifting by ``months`` (renders ``=EOMONTH(...)``).

    ``EOMONTH("2025-01-15", 0)`` → Variable with value ``date(2025, 1, 31)``.
    ``months=1`` → ``date(2025, 2, 28)``. Always returns a Variable.
    """
    raw_start, start_expr = operand(start_date)
    raw_months, months_expr = operand(months)
    m = int(raw_months)

    def end_of(d):
        shifted = _extract_date(d) + relativedelta(months=m)
        last_day = monthrange(shifted.year, shifted.month)[1]
        return date(shifted.year, shifted.month, last_day)

    if isinstance(raw_start, list):
        calc: Any = [end_of(v) for v in raw_start]
        var_type = 'list'
    else:
        calc = end_of(raw_start)
        var_type = 'scalar'

    return make_func_var(
        "EOMONTH", [start_expr, months_expr], calc, 'datetime', var_type,
        source_code=f"EOMONTH({src(start_date)}, {m})",
        render_backends=_EXCEL_ONLY,
    )


def DATE(
    year: Union[Variable, int],
    month: Union[Variable, int],
    day: Union[Variable, int],
) -> Variable:
    """Construct a date from year / month / day (renders ``=DATE(...)``).

    Excel-style overflow: out-of-range month or day rolls over, so
    ``DATE(2025, 15, 10)`` → ``date(2026, 3, 10)`` and the common
    ``DATE(YEAR(d), MONTH(d) + 3, DAY(d))`` quarter-step advances the year
    past December. Always returns a Variable; scalar or element-wise over
    list inputs.
    """
    y_raw, y_expr = operand(year)
    m_raw, m_expr = operand(month)
    d_raw, d_expr = operand(day)

    def build(y: Any, m: Any, d: Any) -> date:
        return date(int(y), 1, 1) + relativedelta(months=int(m) - 1, days=int(d) - 1)

    calc, var_type = _broadcast(build, y_raw, m_raw, d_raw)
    return make_func_var(
        "DATE", [y_expr, m_expr, d_expr], calc, 'datetime', var_type,
        source_code=f"DATE({src(year)}, {src(month)}, {src(day)})",
    )


def DAYS360(
    start_date: DateOperand,
    end_date: DateOperand,
    method: Union[Variable, bool] = False,
) -> Variable:
    """Days between two dates on a 360-day year (renders ``=DAYS360(...)``).

    ``method=False`` (default) uses the US/NASD convention; ``method=True``
    uses the European convention (day 31 always becomes 30). The core of
    30/360 interest accrual::

        days = DAYS360(prev_period_date, this_period_date, True)
        interest = balance * rate / 360 * days

    Always returns a Variable; scalar, or element-wise over list inputs
    (e.g. ``DAYS360(dates.shift(1), dates, True)`` across a schedule).
    """
    start_raw, start_expr = operand(start_date)
    end_raw, end_expr = operand(end_date)
    method_raw, method_expr = operand(method)
    european = bool(method_raw)

    def count(s: Any, e: Any) -> int:
        d1 = _extract_date(s)
        d2 = _extract_date(e)
        day1, day2 = d1.day, d2.day
        if european:
            if day1 == 31:
                day1 = 30
            if day2 == 31:
                day2 = 30
        else:
            if day1 == 31:
                day1 = 30
            if day2 == 31 and day1 == 30:
                day2 = 30
        return (d2.year - d1.year) * 360 + (d2.month - d1.month) * 30 + (day2 - day1)

    calc, var_type = _broadcast(count, start_raw, end_raw)
    return make_func_var(
        "DAYS360", [start_expr, end_expr, method_expr], calc, 'int', var_type,
        source_code=f"DAYS360({src(start_date)}, {src(end_date)}, {src(method)})",
    )


def TODAY() -> Variable:
    """Current date (renders ``=TODAY()``).

    Python value is ``date.today()`` at call time; the Excel cell is
    ``=TODAY()`` which recalculates live when the workbook opens.
    """
    return make_func_var(
        "TODAY", [], date.today(), 'datetime', 'scalar',
        source_code="TODAY()",
    )
