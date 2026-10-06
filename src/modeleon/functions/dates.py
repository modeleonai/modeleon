# SPDX-License-Identifier: Apache-2.0
"""Date helpers — ``YEAR``, ``MONTH``, ``DAY``, ``EDATE``, ``EOMONTH``,
``DATE``, ``DAYS360``, ``YEARFRAC``, ``DAYS``, ``TODAY``.

**Invariant:** every call returns a :class:`Variable` with a tracked
formula — even when inputs are plain ``date`` objects or ISO strings. The
Excel cell is the product (``=EOMONTH(...)``), not just the value.

The day counts (``DAYS360``, ``YEARFRAC``, ``DAYS``) compute what Excel
computes, quirks included, so the workbook recalculates to the engine's
numbers. Dates from 1 March 1900 on: Excel's serial calendar counts a
29 February 1900 that never existed, and nothing here imitates it.
"""

from __future__ import annotations

import math
import numbers
from calendar import isleap, monthrange
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


def _refuse_blanks(func_name: str, *operands: Any, needs: str = "a date") -> None:
    """A blank where the function needs a value is refused by name: the
    line and the period, so the author knows which cell to fill. Excel
    reads the blank as 0 and computes on with the year 1900."""
    for value in operands:
        raw = value._value if isinstance(value, Variable) else value
        if isinstance(raw, list):
            at = next((i for i, v in enumerate(raw) if _is_blank(v)), None)
            if at is None:
                continue
        elif not _is_blank(raw):
            continue
        else:
            at = None
        name = (value._display_name or value.python_name) if isinstance(value, Variable) else None
        line = f"'{name}'" if name else "the input"
        if at is None:
            raise TypeError(f"{func_name} needs {needs}: {line} is blank.")
        raise TypeError(f"{func_name} needs {needs} in every period: "
                        f"{line} is blank in {_period_name(value, at)}.")


def _is_blank(v: Any) -> bool:
    return v is None or (isinstance(v, Variable) and v._value is None)


def _period_name(value: Any, i: int) -> str:
    """The ``i``-th period of ``value``'s time as the timeline header writes
    it (``Feb 2026``); its number where the line has no time."""
    loc = value.time if isinstance(value, Variable) else None
    if loc is not None and loc.start and loc.grain:
        from ..core.time import _period_labels
        try:
            return _period_labels(loc.start, None, i + 1, loc.grain, 'finance')[-1]
        except (ValueError, TypeError):
            pass
    return f"period {i + 1}"


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
        _refuse_blanks(func_name, value)
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
    _refuse_blanks("EDATE", start_date)
    _refuse_blanks("EDATE", months, needs="a number")
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
    _refuse_blanks("EOMONTH", start_date)
    _refuse_blanks("EOMONTH", months, needs="a number")
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
    _refuse_blanks("DATE", year, month, day, needs="a number")
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


def _is_month_end(d: date) -> bool:
    return d.day == monthrange(d.year, d.month)[1]


def _thirty_360(d1: date, day1: int, d2: date, day2: int) -> int:
    """Days between two dates whose day-of-month fields are already adjusted."""
    return (d2.year - d1.year) * 360 + (d2.month - d1.month) * 30 + (day2 - day1)


def _days360(d1: date, d2: date, european: bool) -> int:
    """Excel's ``DAYS360(d1, d2, method)``.

    European (``method`` TRUE): a 31st becomes the 30th, on either date.

    U.S. (``method`` FALSE, the default) is Excel's own rule, NOT the full
    NASD/SIA 30/360 US set:

    - the start date on the last day of ANY month becomes the 30th — the
      last day of February included, always (SIA moves February only for an
      end-of-month security);
    - the end date moves only when it is the 31st: to the 30th when the
      adjusted start is the 30th, otherwise it stays the 31st (Excel's help
      phrases this as "the 1st of the next month" — the same count). The end
      on the last day of February or of a 30-day month does NOT move,
      although Microsoft's current help page says "the last day of a month":
      DAYS360(1-Jan-2011, 28-Feb-2011) is 57 in Excel, not 60;
    - SIA's "both dates on the last day of February → the end becomes the
      30th" is absent: DAYS360(28-Feb-2011, 29-Feb-2012) is 359, not 360,
      and DAYS360(28-Feb-2011, 28-Feb-2011) is -2.

    The dates are never swapped: a start after the end counts negative
    with the same adjustments. The same rules as Apache POI's ``Days360``
    (after its bug 60029) and LibreOffice Calc's source; agrees with Excel
    16.0 (build 20326) on every pair tested. The Python ``formulas``
    package still gives DAYS360(28-Feb-2011, 31-Mar-2011) = 31 where Excel
    gives 30.
    """
    day1, day2 = d1.day, d2.day
    if european:
        day1, day2 = min(day1, 30), min(day2, 30)
    else:
        if _is_month_end(d1):
            day1 = 30
        if day2 == 31 and day1 == 30:
            day2 = 30
    return _thirty_360(d1, day1, d2, day2)


def DAYS360(
    start_date: DateOperand,
    end_date: DateOperand,
    method: Union[Variable, bool] = False,
) -> Variable:
    """Days between two dates on a 360-day year (renders ``=DAYS360(...)``).

    ``method=False`` (default) uses Excel's U.S. (NASD) method — the start
    on the last day of a month, February too, counts as the 30th; an end
    on the 31st counts as the 30th only when the start is the 30th.
    ``method=True`` uses the European method (day 31 always becomes 30).
    Exactly Excel's arithmetic, see :func:`_days360`. The core of 30/360
    interest accrual::

        days = DAYS360(prev_period_date, this_period_date, True)
        interest = balance * rate / 360 * days

    Always returns a Variable; scalar, or element-wise over list inputs
    (e.g. ``DAYS360(dates.shift(1), dates, True)`` across a schedule).
    """
    _refuse_blanks("DAYS360", start_date, end_date)
    start_raw, start_expr = operand(start_date)
    end_raw, end_expr = operand(end_date)
    method_raw, method_expr = operand(method)
    european = bool(method_raw)

    def count(s: Any, e: Any) -> int:
        return _days360(_extract_date(s), _extract_date(e), european)

    calc, var_type = _broadcast(count, start_raw, end_raw)
    return make_func_var(
        "DAYS360", [start_expr, end_expr, method_expr], calc, 'int', var_type,
        source_code=f"DAYS360({src(start_date)}, {src(end_date)}, {src(method)})",
    )


def _yearfrac_30_360_us(d1: date, d2: date) -> int:
    """YEARFRAC basis 0 numerator, ``d1 <= d2`` — Excel's own 30/360 US,
    which is NOT ``DAYS360(d1, d2)``. Two cases differ, both with the start
    on the last day of February:

    - end on the 31st: DAYS360 counts the end as the 30th, YEARFRAC keeps
      the 31st — YEARFRAC(28-Feb-2011, 31-Mar-2011, 0) is 31/360 where
      DAYS360 gives 30. This is the case Microsoft's help page flags as a
      known issue ("may return an incorrect result … start_date is the last
      day in February"); Excel computes it this way, so the engine does;
    - end on the last day of February too: YEARFRAC counts both as the 30th,
      so 28-Feb-2010 to 28-Feb-2011 is exactly 1, where DAYS360 leaves the
      end alone and gives 358.

    Order of the rules as in Apache POI's ``YearFracCalculator`` (after
    D. Wheeler's reverse-engineering of Excel); agrees with Excel 16.0
    (build 20326) on every pair tested.
    """
    day1, day2 = d1.day, d2.day
    if day1 == 31 and day2 == 31:
        day1 = day2 = 30
    elif day1 == 31:
        day1 = 30
    elif day1 == 30 and day2 == 31:
        day2 = 30
    elif d1.month == 2 and _is_month_end(d1):
        day1 = 30
        if d2.month == 2 and _is_month_end(d2):
            day2 = 30
    return _thirty_360(d1, day1, d2, day2)


def _actual_year_length(d1: date, d2: date) -> float:
    """YEARFRAC basis 1 denominator, ``d1 < d2`` — Excel's actual/actual.

    Dates at most a year apart (the same year, or the next year no later
    in the calendar than the start) divide by 366 when a 29 February can
    fall in the span — both dates in one leap year, the start in January or
    February of a leap year, or the end on 29 February or after it in a leap
    year — and by 365 otherwise. Further apart, by the AVERAGE length of
    every calendar year the span touches, first and last included:
    YEARFRAC(1-Jan-2012, 1-Jan-2014, 1) divides 731 days by (366+365+365)/3.
    """
    if d1.year == d2.year:
        return 366.0 if isleap(d1.year) else 365.0
    within_a_year = d1.year + 1 == d2.year and (
        d1.month > d2.month or (d1.month == d2.month and d1.day >= d2.day)
    )
    if within_a_year:
        if isleap(d1.year):
            counts_feb29 = d1.month <= 2
        elif isleap(d2.year):
            counts_feb29 = d2.month > 2 or (d2.month == 2 and d2.day == 29)
        else:
            counts_feb29 = False
        return 366.0 if counts_feb29 else 365.0
    years = range(d1.year, d2.year + 1)
    return sum(366 if isleap(y) else 365 for y in years) / len(years)


def _yearfrac(d1: date, d2: date, basis: int) -> float:
    """Excel's ``YEARFRAC(d1, d2, basis)``: the order of the dates does not
    matter, equal dates give 0."""
    if d1 > d2:
        d1, d2 = d2, d1
    if d1 == d2:
        return 0.0
    if basis == 0:
        return _yearfrac_30_360_us(d1, d2) / 360
    if basis == 1:
        return (d2 - d1).days / _actual_year_length(d1, d2)
    if basis == 2:
        return (d2 - d1).days / 360
    if basis == 3:
        return (d2 - d1).days / 365
    return _days360(d1, d2, european=True) / 360


def _yearfrac_basis(value: Any) -> int:
    """Excel truncates the basis to an integer and returns ``#NUM!`` outside
    0–4 and ``#VALUE!`` for TRUE/FALSE; the engine raises the same cases."""
    if isinstance(value, bool) or not isinstance(value, numbers.Real):
        raise TypeError(
            f"YEARFRAC basis must be a number 0-4, got {value!r} "
            f"(Excel returns #VALUE!)."
        )
    basis = math.trunc(value)
    if not 0 <= basis <= 4:
        raise ValueError(
            f"YEARFRAC basis must be 0-4 — 0 US 30/360, 1 actual/actual, "
            f"2 actual/360, 3 actual/365, 4 European 30/360 — got {value!r} "
            f"(Excel returns #NUM!)."
        )
    return basis


def YEARFRAC(
    start_date: DateOperand,
    end_date: DateOperand,
    basis: Union[Variable, int] = 0,
) -> Variable:
    """Fraction of a year between two dates (renders ``=YEARFRAC(...)``).

    ``basis`` as in Excel: 0 (default) US (NASD) 30/360, 1 actual/actual,
    2 actual/360, 3 actual/365, 4 European 30/360. The order of the dates
    does not matter. Excel's own arithmetic, quirks included — basis 0 is
    not ``DAYS360/360`` at the end of February (:func:`_yearfrac_30_360_us`)
    and basis 1 averages year lengths only over spans longer than a year
    (:func:`_actual_year_length`). An accrual on an actual/365 basis::

        interest = balance * rate * YEARFRAC(prev_date, this_date, 3)

    Always returns a Variable; scalar, or element-wise over list inputs
    (dates and basis alike).
    """
    _refuse_blanks("YEARFRAC", start_date, end_date)
    start_raw, start_expr = operand(start_date)
    end_raw, end_expr = operand(end_date)
    basis_raw, basis_expr = operand(basis)

    def fraction(s: Any, e: Any, b: Any) -> float:
        return _yearfrac(_extract_date(s), _extract_date(e), _yearfrac_basis(b))

    calc, var_type = _broadcast(fraction, start_raw, end_raw, basis_raw)
    return make_func_var(
        "YEARFRAC", [start_expr, end_expr, basis_expr], calc, 'float', var_type,
        source_code=f"YEARFRAC({src(start_date)}, {src(end_date)}, {src(basis)})",
        render_backends=_EXCEL_ONLY,
    )


#: DAYS is an Excel 2013 function: in the file it must be written
#: ``_xlfn.DAYS`` (bare, Excel shows ``#NAME?`` — measured in Excel 16.0,
#: build 20326). The Excel renderer adds the prefix (``_FILE_PREFIX``); the
#: function keeps its plain name everywhere a person reads the formula.
_DAYS_FUNC_NAME = "DAYS"


def DAYS(end_date: DateOperand, start_date: DateOperand) -> Variable:
    """Days from ``start_date`` to ``end_date`` (renders ``=DAYS(end, start)``).

    Excel's argument order — the END date first; ``end - start``, negative
    when the end is earlier. Always returns a Variable; scalar, or
    element-wise over list inputs. In the xlsx the call is stored as
    ``_xlfn.DAYS(...)``, which Excel shows as ``=DAYS(...)``.
    """
    _refuse_blanks("DAYS", end_date, start_date)
    end_raw, end_expr = operand(end_date)
    start_raw, start_expr = operand(start_date)

    def count(e: Any, s: Any) -> int:
        return (_extract_date(e) - _extract_date(s)).days

    calc, var_type = _broadcast(count, end_raw, start_raw)
    return make_func_var(
        _DAYS_FUNC_NAME, [end_expr, start_expr], calc, 'int', var_type,
        source_code=f"DAYS({src(end_date)}, {src(start_date)})",
        render_backends=_EXCEL_ONLY,
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
