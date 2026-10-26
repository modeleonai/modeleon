# SPDX-License-Identifier: Apache-2.0
"""Date helpers — ``YEAR``, ``MONTH``, ``DAY``, ``EDATE``, ``EOMONTH``,
``DATE``, ``DAYS360``, ``YEARFRAC``, ``DAYS``, ``TODAY``, ``DATEDIF``,
``WEEKDAY``, ``WEEKNUM``, ``NETWORKDAYS``, ``WORKDAY``.

**Invariant:** every call returns a :class:`Variable` with a tracked
formula — even when inputs are plain ``date`` objects or ISO strings. The
Excel cell is the product (``=EOMONTH(...)``), not just the value.

The day counts (``DAYS360``, ``YEARFRAC``, ``DAYS``, ``DATEDIF``,
``NETWORKDAYS``) compute what Excel computes, quirks included, so the
workbook recalculates to the engine's numbers. Dates from 1 March 1900 on: Excel's serial calendar counts a
29 February 1900 that never existed, and nothing here imitates it.

LibreOffice Calc recalculates three of these cells differently, and the
engine keeps Excel's numbers: ``WEEKNUM`` (every mode but 21) numbers
the last days of a year 1 where Excel reaches 53 or 54; ``DATEDIF``'s
``"YD"`` moves the start's month and day into the end's year where Excel
moves the end's into the start's; a ``DATEDIF`` with the start after the
end is ``#NUM!`` in Excel and ``#VALUE!`` in LibreOffice.
"""

from __future__ import annotations

import math
import numbers
from calendar import isleap, monthrange
from datetime import date, datetime, timedelta
from typing import Any, Callable, Sequence, Union

from dateutil.relativedelta import relativedelta

from ..core.variable import Variable
from ._helpers import (
    align_keyed, align_located, elementwise, is_error, is_series, make_func_var, operand,
    placed, refuse_unpaired, src,
)
from .registry import register

DateOperand = Union[Variable, date, datetime, str]
Holidays = Union[Variable, Sequence[DateOperand], date, datetime, str, None]


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


def _by_date(func_name: str, lines: Sequence[Any],
             raws: list[Any]) -> tuple[list[Any], Any, list[Any] | None]:
    """The operands' values read as ``a + b`` reads them: lines on other
    time windows by date, lines along keys key by key; with the combined
    window and the keys the result runs on."""
    raws, keys = align_keyed(func_name, lines, raws)
    raws, window = align_located(lines, raws)
    refuse_unpaired(lines, raws, window)
    return raws, window, keys


def _broadcast(fn: Callable[..., Any], *raws: Any) -> tuple[Any, str]:
    """Apply ``fn`` over scalar or list operands, element-wise.

    Scalars repeat; lists shorter than the longest clamp to their last
    element (matches length-1 broadcast elsewhere). Returns
    ``(value, var_type)`` where ``var_type`` is ``'list'`` if any operand
    is a list, else ``'scalar'``. Lines are read by date first
    (:func:`_by_date`).
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
        keys = align_keyed(func_name, [value], [raw])[1]
        if isinstance(raw, list):
            calc: Any = [py_fn(_extract_date(v)) for v in raw]
            var_type = 'list'
        else:
            calc = py_fn(_extract_date(raw))
            var_type = 'scalar'
        out = make_func_var(
            func_name, [expr], calc, 'int', var_type,
            source_code=f"{func_name}({src(value)})",
        )
        return placed(out, None, keys)

    return helper


YEAR = _unary_date_func("YEAR", lambda d: d.year)
YEAR.__doc__ = "Year of a date (renders ``=YEAR(...)``). Always returns a Variable."

MONTH = _unary_date_func("MONTH", lambda d: d.month)
MONTH.__doc__ = "Month 1-12 of a date (renders ``=MONTH(...)``). Always returns a Variable."

DAY = _unary_date_func("DAY", lambda d: d.day)
DAY.__doc__ = "Day of month 1-31 (renders ``=DAY(...)``). Always returns a Variable."


# One period of ``EDATE`` / ``EOMONTH`` whose months are a series: the
# months read whole, their fraction dropped, as the spreadsheet reads them.
def _edate(start: Any, months: Any) -> date:
    return _extract_date(start) + relativedelta(months=int(months))


def _eomonth(start: Any, months: Any) -> date:
    shifted = _edate(start, months)
    return date(shifted.year, shifted.month, monthrange(shifted.year, shifted.month)[1])


def EDATE(start_date: DateOperand, months: Union[Variable, int]) -> Variable:
    """Shift a date by ``months`` (renders ``=EDATE(...)``). Always returns a Variable.

    ``EDATE("2025-01-15", 3)`` → Variable with value ``date(2025, 4, 15)`` and
    formula ``EDATE("2025-01-15", 3)``. Negative values go backward.
    ``months`` may vary by period: a series of months is one call per
    period, ``=EDATE(C5, C9)`` in every column.
    """
    _refuse_blanks("EDATE", start_date)
    _refuse_blanks("EDATE", months, needs="a number")
    if is_series(months):
        return elementwise("EDATE", _edate, [start_date, months], value_type='datetime',
                           exceptions='raise', render_backends=_EXCEL_ONLY)
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

    out = make_func_var(
        "EDATE", [start_expr, months_expr], calc, 'datetime', var_type,
        source_code=f"EDATE({src(start_date)}, {m})",
        render_backends=_EXCEL_ONLY,
    )
    return placed(out, None, align_keyed("EDATE", [start_date], [raw_start])[1])


def EOMONTH(start_date: DateOperand, months: Union[Variable, int] = 0) -> Variable:
    """End-of-month date after shifting by ``months`` (renders ``=EOMONTH(...)``).

    ``EOMONTH("2025-01-15", 0)`` → Variable with value ``date(2025, 1, 31)``.
    ``months=1`` → ``date(2025, 2, 28)``. Always returns a Variable.
    ``months`` may vary by period, one call per period as for
    :func:`EDATE`.
    """
    _refuse_blanks("EOMONTH", start_date)
    _refuse_blanks("EOMONTH", months, needs="a number")
    if is_series(months):
        return elementwise("EOMONTH", _eomonth, [start_date, months], value_type='datetime',
                           exceptions='raise', render_backends=_EXCEL_ONLY)
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

    out = make_func_var(
        "EOMONTH", [start_expr, months_expr], calc, 'datetime', var_type,
        source_code=f"EOMONTH({src(start_date)}, {m})",
        render_backends=_EXCEL_ONLY,
    )
    return placed(out, None, align_keyed("EOMONTH", [start_date], [raw_start])[1])


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
        return _date_overflow(int(y), int(m), int(d))

    (y_raw, m_raw, d_raw), window, keys = _by_date(
        "DATE", (year, month, day), [y_raw, m_raw, d_raw])
    calc, var_type = _broadcast(build, y_raw, m_raw, d_raw)
    out = make_func_var(
        "DATE", [y_expr, m_expr, d_expr], calc, 'datetime', var_type,
        source_code=f"DATE({src(year)}, {src(month)}, {src(day)})",
    )
    return placed(out, window, keys)


def _date_overflow(year: int, month: int, day: int) -> date:
    """``DATE(year, month, day)`` as the spreadsheet reads it: a month or a
    day past its range rolls over, so ``DATE(2025, 2, 29)`` is 1 March
    2025 and month 0 is the December before."""
    return date(year, 1, 1) + relativedelta(months=month - 1, days=day - 1)


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
    (e.g. ``DAYS360(dates.shift(1), dates, True)`` across a schedule) —
    the method too, each period's own.
    """
    _refuse_blanks("DAYS360", start_date, end_date)
    start_raw, start_expr = operand(start_date)
    end_raw, end_expr = operand(end_date)
    method_raw, method_expr = operand(method)

    def count(s: Any, e: Any, european: Any) -> int:
        return _days360(_extract_date(s), _extract_date(e), bool(european))

    (start_raw, end_raw, method_raw), window, keys = _by_date(
        "DAYS360", (start_date, end_date, method), [start_raw, end_raw, method_raw])
    calc, var_type = _broadcast(count, start_raw, end_raw, method_raw)
    out = make_func_var(
        "DAYS360", [start_expr, end_expr, method_expr], calc, 'int', var_type,
        source_code=f"DAYS360({src(start_date)}, {src(end_date)}, {src(method)})",
    )
    return placed(out, window, keys)


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

    (start_raw, end_raw, basis_raw), window, keys = _by_date(
        "YEARFRAC", (start_date, end_date, basis), [start_raw, end_raw, basis_raw])
    calc, var_type = _broadcast(fraction, start_raw, end_raw, basis_raw)
    out = make_func_var(
        "YEARFRAC", [start_expr, end_expr, basis_expr], calc, 'float', var_type,
        source_code=f"YEARFRAC({src(start_date)}, {src(end_date)}, {src(basis)})",
        render_backends=_EXCEL_ONLY,
    )
    return placed(out, window, keys)


#: DAYS is an Excel 2013 function: in the file it must be written
#: ``_xlfn.DAYS`` (bare, Excel shows ``#NAME?`` — measured in Excel 16.0,
#: build 20326). The registry entry below carries the prefix
#: (``file_prefix='_xlfn.'``) for the Excel renderer; the function keeps
#: its plain name everywhere a person reads the formula.
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

    (end_raw, start_raw), window, keys = _by_date(
        "DAYS", (end_date, start_date), [end_raw, start_raw])
    calc, var_type = _broadcast(count, end_raw, start_raw)
    out = make_func_var(
        _DAYS_FUNC_NAME, [end_expr, start_expr], calc, 'int', var_type,
        source_code=f"DAYS({src(end_date)}, {src(start_date)})",
        render_backends=_EXCEL_ONLY,
    )
    return placed(out, window, keys)


def TODAY() -> Variable:
    """Current date (renders ``=TODAY()``).

    Python value is ``date.today()`` at call time; the Excel cell is
    ``=TODAY()`` which recalculates live when the workbook opens.
    """
    return make_func_var(
        "TODAY", [], date.today(), 'datetime', 'scalar',
        source_code="TODAY()",
    )


# ─── Day counts and calendar positions ────────────────────────────────
#
# ``DATEDIF``, ``WEEKDAY``, ``WEEKNUM``, ``NETWORKDAYS`` and ``WORKDAY`` are
# one call per period through the shared :func:`elementwise` helper over
# their date operands — a scalar date stands in every period, a series
# gives a series, tracks lift, an error cell passes through — under this
# module's law for what the function needs: a blank, or a value that is
# not a date, is refused by name at build time, never computed on as the
# year 1900. An error the DATA decides (``DATEDIF`` with the start after
# the end) is the period's ``#NUM!``, so an ``IF`` can guard it; a wrong
# constant (a unit, a return type) is refused at build time with Excel's
# error named. The holidays of ``NETWORKDAYS`` / ``WORKDAY`` are one range
# every period's call reads, never a series. No result is in an operand's
# unit: a day count, a weekday, a week number and a date carry none.


_DATEDIF_UNITS = ('Y', 'M', 'D', 'MD', 'YM', 'YD')

#: The weekday a week begins on (Monday 0 … Sunday 6) for each return
#: type ``WEEKDAY`` and ``WEEKNUM`` share: 1 and 17 Sunday, 2, 3 and 11
#: Monday, 12 to 16 Tuesday to Saturday.
_WEEK_STARTS = {1: 6, 2: 0, 3: 0, 11: 0, 12: 1, 13: 2, 14: 3, 15: 4, 16: 5, 17: 6}
_WEEKDAY_TYPES = frozenset(_WEEK_STARTS)
_WEEKNUM_TYPES = frozenset(_WEEK_STARTS) - {3} | {21}


def _no_unit(out: Variable) -> Variable:
    out._unit = None
    return out


def _whole_number(func_name: str, value: Any, what: str) -> int:
    """A number the function reads whole, truncated as Excel truncates
    it; TRUE/FALSE and text are refused (Excel returns ``#VALUE!``)."""
    if isinstance(value, bool) or not isinstance(value, numbers.Real):
        raise TypeError(
            f"{func_name} {what} must be a number, got {value!r} (Excel returns #VALUE!)."
        )
    return math.trunc(value)


def _return_type(func_name: str, value: Any, allowed: frozenset[int], spelled: str) -> int:
    kind = _whole_number(func_name, value, "return type")
    if kind not in allowed:
        raise ValueError(
            f"{func_name} return type must be {spelled}, got {value!r} "
            f"(Excel returns #NUM!)."
        )
    return kind


def _datedif_unit(value: Any) -> str:
    if not isinstance(value, str) or value.upper() not in _DATEDIF_UNITS:
        raise ValueError(
            f"DATEDIF unit must be one of 'Y', 'M', 'D', 'MD', 'YM', 'YD' — whole "
            f"years, whole months, days, days past the whole months, months past "
            f"the whole years, days past the whole years — got {value!r} "
            f"(Excel returns #NUM!)."
        )
    return value.upper()


def _datedif(start: Any, end: Any, unit: Any) -> Any:
    """Excel's ``DATEDIF(start, end, unit)``, its quirks included.

    ``"Y"`` and ``"M"`` are complete years and months — a year or a month
    is complete only once the end's day is reached, so 31 January to 28
    February is 0 months and 29 February to the next 28 February 0 years.
    ``"YM"`` is the months past the whole years.

    ``"MD"``, the days past the whole months, is the count Microsoft's
    help flags as possibly negative: with the start's day past the end's,
    the start's day is placed in the month before the end's and rolls
    over as ``DATE()`` rolls it, so 31 January to 1 March of a common year
    is 1 March minus "31 February" = 3 March, -2 (the rule LibreOffice
    extracted from Excel by trial and error).

    ``"YD"``, the days past the whole years, places the END's month and
    day in the START's year — or the year after, when they fall before
    the start's — and counts from the start: 3 January 2015 to 28 July
    2016 is 206 through the 28-day February of 2015, where a count that
    ends at the end date gives 207 (Excel's value is documented at
    exceljet.net; LibreOffice counts the other way and notes in its
    source that Excel answers 323 where it answers 324 for 22 June 1988
    to 11 May 2012). The rule rests on those measured values, none of
    which has a 29 February at either end; where one does, it is
    applied as written — 29 February 2024 to 1 March 2025 is 1 day past
    the whole year (LibreOffice: 0), and an end on 29 February placed in
    a common year rolls over to 1 March — and the day may differ in
    Excel.

    A start after the end is ``#NUM!``, never swapped.
    """
    s, e = _extract_date(start), _extract_date(end)
    kind = _datedif_unit(unit)
    if s > e:
        return '#NUM!'
    if kind == 'D':
        return (e - s).days
    years = e.year - s.year - int((e.month, e.day) < (s.month, s.day))
    if kind == 'Y':
        return years
    months = (e.year - s.year) * 12 + e.month - s.month - int(e.day < s.day)
    if kind == 'M':
        return months
    if kind == 'YM':
        return months - 12 * years
    if kind == 'MD':
        if e.day >= s.day:
            return e.day - s.day
        return (e - _date_overflow(e.year, e.month - 1, s.day)).days
    later = int((e.month, e.day) < (s.month, s.day))
    return (_date_overflow(s.year + later, e.month, e.day) - s).days


def _weekday(value: Any, return_type: Any = 1) -> int:
    d = _extract_date(value)
    kind = _return_type("WEEKDAY", return_type, _WEEKDAY_TYPES, "1, 2, 3 or 11-17")
    number = (d.weekday() - _WEEK_STARTS[kind]) % 7
    return number if kind == 3 else number + 1


def _weeknum(value: Any, return_type: Any = 1) -> int:
    d = _extract_date(value)
    kind = _return_type("WEEKNUM", return_type, _WEEKNUM_TYPES, "1, 2, 11-17 or 21")
    if kind == 21:
        return d.isocalendar()[1]
    first = date(d.year, 1, 1)
    # Week 1 is the week that holds 1 January, however few of its days
    # fall in the year.
    into_week = (first.weekday() - _WEEK_STARTS[kind]) % 7
    return ((d - first).days + into_week) // 7 + 1


def _networkdays(start: Any, end: Any, holidays: frozenset[date]) -> int:
    s, e = _extract_date(start), _extract_date(end)
    sign = 1
    if s > e:
        s, e, sign = e, s, -1
    weeks, rest = divmod((e - s).days + 1, 7)
    count = weeks * 5 + sum(1 for i in range(rest) if (s + timedelta(days=i)).weekday() < 5)
    count -= sum(1 for h in holidays if s <= h <= e and h.weekday() < 5)
    return sign * count


def _workday(start: Any, days: Any, holidays: frozenset[date]) -> date:
    d = _extract_date(start)
    left = _whole_number("WORKDAY", days, "days")
    step = timedelta(days=1 if left > 0 else -1)
    left = abs(left)
    while left:
        d += step
        if d.weekday() < 5 and d not in holidays:
            left -= 1
    return d


def _holidays(func_name: str, holidays: Holidays) -> tuple[frozenset[date], Any, Any]:
    """The holidays as a set of dates, their AST node (None when none were
    given) and the first error marker among them (None when none). A
    blank among them is skipped, as an empty cell of the range is, and
    so is a 0 — what a line extended with ``mo.zero()`` fills with, a
    serial before every date, which the spreadsheet never counts."""
    if holidays is None:
        return frozenset(), None, None
    if isinstance(holidays, Variable):
        from ..core.shape import reject_axised
        from ..core.tracks import TrackValues
        reject_axised(holidays, func_name)
        if isinstance(holidays._value, TrackValues):
            name = holidays._display_name or holidays.python_name or "the holidays"
            raise ValueError(
                f"{func_name}: {name!r} carries tracks ({', '.join(holidays._value.roles)}) — "
                f"the holidays are one list of dates. Read one track first (.at(track=...))."
            )
    elif isinstance(holidays, (list, tuple)):
        if any(isinstance(h, Variable) for h in holidays):
            raise TypeError(
                f"{func_name} holidays: a list of lines is not a range — put the dates in "
                f"one line (mo.Variable([...])) and pass that line."
            )
        holidays = list(holidays)
    raw, node = operand(holidays)
    values = raw if isinstance(raw, list) else [raw]
    error = next((v for v in values if is_error(v)), None)
    dates = set()
    for i, v in enumerate(values):
        if v is None or is_error(v) or _is_zero(v):
            continue
        try:
            dates.add(_extract_date(v))
        except (TypeError, ValueError):
            if isinstance(holidays, Variable):
                label = holidays._display_name or holidays.python_name
                where = f"'{label}'" if label else "the holidays line"
                at = f" in {_period_name(holidays, i)}" if isinstance(raw, list) else ""
            elif isinstance(raw, list):
                where, at = "the list", f" at position {i + 1}"
            else:
                where, at = "the holidays", ""
            raise TypeError(
                f"{func_name} holidays must be dates: {where} holds {v!r}{at}."
            ) from None
    return frozenset(dates), node, error


def _is_zero(v: Any) -> bool:
    return isinstance(v, numbers.Real) and not isinstance(v, bool) and v == 0


def _with_holidays(
    func_name: str,
    kernel: Callable[..., Any],
    cell_args: list[Any],
    holidays: Holidays,
    value_type: str,
) -> Variable:
    """One call per period over ``cell_args`` with the holidays read
    whole by every call — the range position of the written formula,
    not a series the helper would spread over the periods."""
    dates, node, error = _holidays(func_name, holidays)

    def call(*cells: Any) -> Any:
        return error if error is not None else kernel(*cells, dates)

    args = cell_args if node is None else [*cell_args, holidays]
    source_code = f"{func_name}({', '.join(src(a) for a in args)})"
    out = elementwise(func_name, call, cell_args, value_type=value_type, exceptions='raise',
                      source_code=source_code, render_backends=_EXCEL_ONLY)
    if node is not None:
        # The same call with the holidays among its operands: on the window
        # and along the keys the per-period call's lines were read on.
        window = out.time if getattr(out, '_grain', None) is not None else None
        out = placed(make_func_var(
            func_name, [operand(a)[1] for a in cell_args] + [node], out._value, value_type,
            out.var_type, source_code=source_code, render_backends=_EXCEL_ONLY,
        ), window, out._keys)
    return _no_unit(out)


def DATEDIF(start_date: DateOperand, end_date: DateOperand, unit: Union[Variable, str]) -> Variable:
    """Whole years, months or days between two dates (renders ``=DATEDIF(...)``).

    ``unit`` as in Excel, in either case: ``"Y"`` complete years, ``"M"``
    complete months, ``"D"`` days, ``"YM"`` months past the whole years,
    ``"MD"`` days past the whole months, ``"YD"`` days past the whole
    years. Excel's own arithmetic, quirks included — ``"MD"`` can be
    negative (31 January to 1 March is -1 in a leap year, -2 otherwise)
    and ``"YD"`` counts from the start date through the start's own
    year, see :func:`_datedif`. A start after the end is the period's
    ``#NUM!``, so ``mo.IF(end >= start, mo.DATEDIF(start, end, "M"), 0)``
    guards it. Months of service to a leaving date::

        months = mo.DATEDIF(joined, mo.IF(mo.ISBLANK(left), mo.TODAY(), left), "M")

    Always returns a Variable; scalar, or one call per period over list
    inputs. A blank date is refused by name.
    """
    _refuse_blanks("DATEDIF", start_date, end_date)
    _refuse_blanks("DATEDIF", unit, needs="a unit")
    return _no_unit(elementwise(
        "DATEDIF", _datedif, [start_date, end_date, unit], value_type='int',
        exceptions='raise', render_backends=_EXCEL_ONLY,
    ))


def WEEKDAY(value: DateOperand, return_type: Union[Variable, int, None] = None) -> Variable:
    """Day of the week as a number (renders ``=WEEKDAY(...)``).

    ``return_type`` as in Excel: 1 (default) Sunday is 1 to Saturday 7, 2
    Monday is 1 to Sunday 7, 3 Monday is 0 to Sunday 6, 11 to 17 a week
    beginning on Monday to Sunday with that day 1. Written without the
    type when none is given. Always returns a Variable; scalar, or one
    call per period over a list of dates.
    """
    _refuse_blanks("WEEKDAY", value)
    args: list[Any] = [value]
    if return_type is not None:
        _refuse_blanks("WEEKDAY", return_type, needs="a number")
        args.append(return_type)
    return _no_unit(elementwise(
        "WEEKDAY", _weekday, args, value_type='int', exceptions='raise',
        render_backends=_EXCEL_ONLY,
    ))


def WEEKNUM(value: DateOperand, return_type: Union[Variable, int, None] = None) -> Variable:
    """Week of the year (renders ``=WEEKNUM(...)``).

    ``return_type`` as in Excel: 1 (default) weeks begin on Sunday, 2 on
    Monday, 11 to 17 on Monday to Sunday — week 1 is the week that holds
    1 January, so the last days of a year reach week 53 — and 21 the ISO
    week, where the first days of a year can belong to week 52 or 53 of
    the year before. Written without the type when none is given. Always
    returns a Variable; scalar, or one call per period over a list of
    dates. LibreOffice Calc gives the last days of a year week 1 in every
    mode but 21 - it counts the week that runs into the next year as that
    year's first; the engine computes Excel's 53 or 54.
    """
    _refuse_blanks("WEEKNUM", value)
    args: list[Any] = [value]
    if return_type is not None:
        _refuse_blanks("WEEKNUM", return_type, needs="a number")
        args.append(return_type)
    return _no_unit(elementwise(
        "WEEKNUM", _weeknum, args, value_type='int', exceptions='raise',
        render_backends=_EXCEL_ONLY,
    ))


def NETWORKDAYS(
    start_date: DateOperand,
    end_date: DateOperand,
    holidays: Holidays = None,
) -> Variable:
    """Working days from ``start_date`` to ``end_date``, both included
    (renders ``=NETWORKDAYS(...)``).

    Saturdays and Sundays are not working days, nor are the ``holidays``:
    a line of dates in the model (``inputs.holidays``, written as the
    range the formula reads; a line shorter than the model's window keeps
    its own length with ``extend=mo.none()``), a plain list of dates, or
    one date. A holiday on a weekend or outside the span changes nothing;
    a blank or a 0 among the holidays is skipped, as an empty cell is. A
    start after the end counts negative. Always returns a Variable;
    scalar, or one call per period over list dates with the same holidays
    throughout. A plain list of holidays has no cells of its own, so the
    workbook carries the computed value in its place (the renderer warns
    when it does); name the holidays as a line to keep the formula live.
    """
    _refuse_blanks("NETWORKDAYS", start_date, end_date)
    return _with_holidays("NETWORKDAYS", _networkdays, [start_date, end_date], holidays, 'int')


def WORKDAY(
    start_date: DateOperand,
    days: Union[Variable, int, float],
    holidays: Holidays = None,
) -> Variable:
    """The date ``days`` working days after ``start_date`` (renders ``=WORKDAY(...)``).

    Saturdays, Sundays and the ``holidays`` (as for :func:`NETWORKDAYS`)
    are stepped over; negative ``days`` go backward; a fraction is
    truncated; 0 is the start date itself, weekend or not. Always
    returns a Variable holding a date; scalar, or one call per period
    over a list start or a list of day counts.
    """
    _refuse_blanks("WORKDAY", start_date)
    _refuse_blanks("WORKDAY", days, needs="a number")
    return _with_holidays("WORKDAY", _workday, [start_date, days], holidays, 'datetime')


register("YEAR", shape="elementwise")
register("MONTH", shape="elementwise")
register("DAY", shape="elementwise")
register("DATE", shape="elementwise")
register("EDATE", shape="elementwise")
register("EOMONTH", shape="elementwise")
register("DAYS360", shape="elementwise")
register("TODAY", shape="elementwise")
# Computed over a series by their helpers; not run period by period
# inside a loop container.
register("YEARFRAC", shape="scalar")
register(_DAYS_FUNC_NAME, shape="scalar", file_prefix="_xlfn.")
register("DATEDIF", shape="elementwise")
register("WEEKDAY", shape="elementwise")
register("WEEKNUM", shape="elementwise")
# One call per period over the dates (or the day count), the holidays a
# range every period's call reads: ``=NETWORKDAYS(B2, C2, Inputs!B5:D5)``.
register("NETWORKDAYS", shape="range", range_args=[2])
register("WORKDAY", shape="range", range_args=[2])
