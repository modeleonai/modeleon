# SPDX-License-Identifier: Apache-2.0
"""Time — the discrete period axis and grain calendar.

``mo.Time`` builds a period-labeled axis Variable. An axis is just a
Variable (see :mod:`modeleon.core.shape`), so it rides the existing shape
machinery::

    time = mo.Time('2024-01', periods=3)           # monthly
    revenue = mo.Variable([100, 110, 121], indexed_by=[time])

Grains: ``day`` (``YYYY-MM-DD``), ``month`` (``YYYY-MM``), ``quarter``
(``YYYY-Qn``), ``year`` (``YYYY``). Extent is ``start`` plus exactly one of
``end`` or ``periods``. The per-grain calendar helpers (parse / step /
format / count) are the foundation the re-grain projection reuses.
"""

from __future__ import annotations

from collections import namedtuple
from datetime import date, timedelta
from typing import Any, List, Optional

from .regrain import LINE_UNKNOWN, apply_recipe, hole_value
from .variable import Variable

#: Supported grains, finest to coarsest. ``week`` is not supported.
GRAINS = ('day', 'month', 'quarter', 'year')

#: Coarsening order — re-grain only goes left-to-right.
_ORDER = {'day': 0, 'month': 1, 'quarter': 2, 'year': 3}

#: Period-label styles for the Excel timeline header. ``iso`` is the engine's
#: native form (matches ``mo.Time`` axis labels); ``finance`` / ``compact`` are
#: presentation variants.
TIME_LABEL_STYLES = ('iso', 'finance', 'compact')

#: Default header style when an ``ExcelView`` declares none.
DEFAULT_TIME_LABEL_STYLE = 'finance'

_MONTH_ABBR = ('', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
               'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec')

#: A model's default *projection* window — ``start`` + ``grain`` + number of
#: ``periods``. Resolved up the MV tree by :func:`resolve_default_window`.
#: Distinct from a Variable's native ``.time`` location: ``start``/``grain``
#: also seed each line's native location (they cascade to ``.time``), but
#: ``periods`` is window-only — a line's period count is just its data length.
#: ``start`` is always a period LABEL. When the window starts at a date
#: Variable (``m.default_start = m.inputs.financial_close``),
#: ``start_source`` is that Variable — the cell a workbook's time is
#: anchored to — and ``start`` is its date's label.
#: ``first_day`` is set when a quarter or year window starts inside its
#: first period (on the 1st of a month): that period — a stub — runs from
#: this day to its calendar end and keeps its label; every later one is
#: whole.
DefaultWindow = namedtuple('DefaultWindow',
                           ['start', 'grain', 'periods', 'start_source', 'first_day'],
                           defaults=(None, None))


def resolve_default_window(node: Any) -> Optional['DefaultWindow']:
    """The default projection window declared nearest up ``node``'s MV tree.

    Walks ``node`` then its ``_owner`` / ``_parent`` chain for the first that
    declares ``default_grain`` and returns its ``(default_start, default_grain,
    default_periods)``. ``None`` if none is declared. Consumed by ``to_excel``
    as the window to project onto when the caller passes no explicit
    ``start`` / ``periods`` / ``grain``.
    """
    seen: set = set()
    n = node
    while n is not None and id(n) not in seen:
        seen.add(id(n))
        grain = getattr(n, 'default_grain', None)
        if grain is not None:
            slots = n.__dict__ if hasattr(n, '__dict__') else {}
            return DefaultWindow(
                getattr(n, 'default_start', None), grain,
                getattr(n, 'default_periods', None),
                slots.get('_default_start_source'), slots.get('_default_first_day'),
            )
        n = getattr(n, '_owner', None) or getattr(n, '_parent', None)
    return None


def section_window_refusal(node: Any, reads_time: bool = True) -> Optional[str]:
    """Why ``node``'s window cannot be used, or ``None``: the nearest
    container declaring ``default_grain`` names its grain but not its
    own window, under a window of another grain. Its time would be read
    off the window above while its lines stand at its own grain.

    Reading ``mo.time`` needs the whole window (``default_start`` and
    ``default_periods``); a series placed in it needs its start
    (``reads_time=False``). The model at the top may name its grain
    before its start (the start is often an input, set later), so only a
    container with a parent is refused; one of the same grain as the
    window above lives in it."""
    seen: set = set()
    n = node
    declaring = None
    while n is not None and id(n) not in seen:
        seen.add(id(n))
        if getattr(n, 'default_grain', None) is not None:
            declaring = n
            break
        n = getattr(n, '_owner', None) or getattr(n, '_parent', None)
    if declaring is None:
        return None
    if getattr(declaring, 'default_start', None) is not None and (
            getattr(declaring, 'default_periods', None) or not reads_time):
        return None
    parent = getattr(declaring, '_owner', None) or getattr(declaring, '_parent', None)
    if parent is None:
        return None
    above = resolve_default_window(parent)
    grain = declaring.default_grain
    if above is None or above.grain is None or above.grain == grain:
        return None
    label = (getattr(declaring, '_display_name', None)
             or getattr(declaring, '_python_name', None) or 'A section')
    counts = f"{above.grain}s from {above.start}" if above.start else f"{above.grain}s"
    if above.periods:
        counts += f", {above.periods} of them"
    return (
        f"{label!r} names its grain ({grain}) but not its own window, and the "
        f"window above it counts in {counts}. Give it its own window - "
        f"default_start and default_periods beside default_grain - or drop "
        f"default_grain to live in the window above."
    )


def window_start_label(value: Any, grain: str) -> str:
    """The period label a window starting at ``value`` begins with.

    ``value`` is a label (``'2026-01'``) or a date — the day a deal is
    anchored to. A date starts on the 1st of a month. On a quarter or year
    grain it may fall inside the period — a valuation on 1 May: the first
    period is then the rest of that quarter or year (a stub,
    :func:`stub_first_day`) and keeps its label. A day inside a month is
    refused: a period's months are whole.
    """
    from datetime import datetime
    if isinstance(value, str):
        _parse(value, grain)
        return value
    if isinstance(value, datetime):
        if value.time() != datetime.min.time():
            raise ValueError(
                f"A window starts at a day, not at a time of day; got "
                f"{value.isoformat()}."
            )
        value = value.date()
    if not isinstance(value, date):
        raise TypeError(
            f"A window starts at a date or a period label like '2026-01'; "
            f"got {value!r}."
        )
    if grain == 'day':
        return value.isoformat()
    if grain == 'month' and value.day == 1:
        return f"{value.year:04d}-{value.month:02d}"
    if grain == 'quarter' and value.day == 1 and value.month in (1, 4, 7, 10):
        return f"{value.year:04d}-Q{(value.month - 1) // 3 + 1}"
    if grain == 'year' and (value.month, value.day) == (1, 1):
        return f"{value.year:04d}"
    if grain in ('quarter', 'year') and value.day == 1:
        # A stub: the first period is the rest of this quarter / year.
        return _containing_label(value, grain)
    raise ValueError(
        f"A {grain} window starts on the 1st of a month; {value.isoformat()} "
        f"falls inside a month. Start the window on the 1st and keep the "
        f"exact day as an input — read mo.time.days for the part of a period "
        f"after it."
    )


def _containing_label(day: date, grain: str) -> str:
    """The label of the ``grain`` period that holds ``day``."""
    if grain == 'quarter':
        return f"{day.year:04d}-Q{(day.month - 1) // 3 + 1}"
    if grain == 'year':
        return f"{day.year:04d}"
    if grain == 'month':
        return f"{day.year:04d}-{day.month:02d}"
    return day.isoformat()


def spelled_first_day(window: Any) -> Optional[date]:
    """The day a workbook spells a quarter / year ``window``'s time from so
    that its first period may be a stub, or ``None`` for the whole-period
    spelling: the stub's day, or — the window starts at a date cell, which
    may be moved inside a period in the book — its first period's first
    day. A start written into the formulas cannot move: whole spelling."""
    if window is None or window.grain not in ('quarter', 'year') or window.start is None:
        return None
    if getattr(window, 'first_day', None) is not None:
        return window.first_day
    if getattr(window, 'start_source', None) is not None:
        return _first_day(_parse(window.start, window.grain), window.grain)
    return None


def coarse_first_day(native_start: str, native_grain: str, native_first_day: Optional[date],
                     target_grain: str) -> Optional[date]:
    """The day the first ``target_grain`` period of a series re-grained from
    ``native_start`` starts on, when that is inside the coarse period — the
    coarse first period is then short, as every line and time row of the
    re-grained view must agree. ``None`` when it starts on the coarse
    period's first day, or the target is finer than a quarter."""
    if target_grain not in ('quarter', 'year') or native_start is None:
        return None
    anchor = _parse(native_start, native_grain)
    first = native_first_day or _first_day(anchor, native_grain)
    if first.day != 1:
        return None
    label = _target_label(anchor, native_grain, target_grain)
    if _first_day(_parse(label, target_grain), target_grain) == first:
        return None
    return first


def stub_first_day(value: Any, grain: str) -> Optional[date]:
    """The day a window's short first period starts on, or ``None`` when
    its first period is whole: a quarter or year window starting at a date
    inside its first period (always the 1st of a month —
    :func:`window_start_label` refuses any other day)."""
    from datetime import datetime
    if grain not in ('quarter', 'year'):
        return None
    if isinstance(value, datetime):
        value = value.date()
    if not isinstance(value, date):
        return None
    if _first_day(_parse(_containing_label(value, grain), grain), grain) == value:
        return None
    return value


def _parse(start: str, grain: str) -> Any:
    """Parse a ``start`` label into a per-grain anchor (the first period)."""
    if grain == 'day':
        try:
            return date.fromisoformat(start)
        except (ValueError, TypeError):
            raise ValueError(f"Time day start must be 'YYYY-MM-DD'; got {start!r}.")
    if grain == 'month':
        return _parse_year_month(start)
    if grain == 'quarter':
        return _parse_year_quarter(start)
    if grain == 'year':
        try:
            return int(start)
        except (ValueError, TypeError):
            raise ValueError(f"Time year start must be 'YYYY'; got {start!r}.")
    raise ValueError(f"Unknown time grain {grain!r}; supported: {', '.join(GRAINS)}.")


def _parse_year_month(start: str) -> tuple:
    parts = start.split('-') if isinstance(start, str) else []
    if len(parts) != 2:
        raise ValueError(f"Time month start must be 'YYYY-MM'; got {start!r}.")
    try:
        year, month = int(parts[0]), int(parts[1])
    except ValueError:
        raise ValueError(f"Time month start must be 'YYYY-MM'; got {start!r}.")
    if not 1 <= month <= 12:
        raise ValueError(f"Month must be 01-12 in {start!r}.")
    return (year, month)


def _parse_year_quarter(start: str) -> tuple:
    parts = start.upper().split('-Q') if isinstance(start, str) else []
    if len(parts) != 2:
        raise ValueError(f"Time quarter start must be 'YYYY-Qn'; got {start!r}.")
    try:
        year, quarter = int(parts[0]), int(parts[1])
    except ValueError:
        raise ValueError(f"Time quarter start must be 'YYYY-Qn'; got {start!r}.")
    if not 1 <= quarter <= 4:
        raise ValueError(f"Quarter must be Q1-Q4 in {start!r}.")
    return (year, quarter)


def _step(anchor: Any, grain: str) -> Any:
    """Advance an anchor by one period of ``grain``."""
    if grain == 'day':
        return anchor + timedelta(days=1)
    if grain == 'month':
        year, month = anchor
        month += 1
        if month > 12:
            month, year = 1, year + 1
        return (year, month)
    if grain == 'quarter':
        year, quarter = anchor
        quarter += 1
        if quarter > 4:
            quarter, year = 1, year + 1
        return (year, quarter)
    return anchor + 1  # year


def _format(anchor: Any, grain: str, style: str = 'iso') -> str:
    if grain == 'day':
        if style == 'finance':
            return f"{anchor.day:02d} {_MONTH_ABBR[anchor.month]} {anchor.year:04d}"
        if style == 'compact':
            return f"{anchor.day:02d}-{_MONTH_ABBR[anchor.month]}-{anchor.year % 100:02d}"
        return anchor.isoformat()
    if grain == 'month':
        year, month = anchor
        if style == 'finance':
            return f"{_MONTH_ABBR[month]} {year:04d}"
        if style == 'compact':
            return f"{_MONTH_ABBR[month]}-{year % 100:02d}"
        return f"{year:04d}-{month:02d}"
    if grain == 'quarter':
        year, quarter = anchor
        if style == 'finance':
            return f"Q{quarter} {year:04d}"
        if style == 'compact':
            return f"Q{quarter}-{year % 100:02d}"
        return f"{year:04d}-Q{quarter}"
    # year
    if style == 'compact':
        return f"FY{anchor % 100:02d}"
    return f"{anchor:04d}"


def _count(start_anchor: Any, end_anchor: Any, grain: str) -> int:
    """Inclusive period count from ``start_anchor`` to ``end_anchor``."""
    if grain == 'day':
        return (end_anchor - start_anchor).days + 1
    if grain == 'month':
        (sy, sm), (ey, em) = start_anchor, end_anchor
        return (ey - sy) * 12 + (em - sm) + 1
    if grain == 'quarter':
        (sy, sq), (ey, eq) = start_anchor, end_anchor
        return (ey - sy) * 4 + (eq - sq) + 1
    return end_anchor - start_anchor + 1  # year


def _period_labels(start: str, end: Optional[str], periods: Optional[int],
                   grain: str, style: str = 'iso') -> List[str]:
    anchor = _parse(start, grain)
    if periods is not None:
        count = periods
    else:
        count = _count(anchor, _parse(end, grain), grain)  # type: ignore[arg-type]
        if count <= 0:
            raise ValueError(f"Time end {end!r} is before start {start!r}.")
    labels: List[str] = []
    current = anchor
    for _ in range(count):
        labels.append(_format(current, grain, style))
        current = _step(current, grain)
    return labels


class Time(Variable):
    """The discrete time axis — a period-labeled axis Variable.

    ``mo.Time('2024-01', '2024-12')`` and ``mo.Time('2024-01', periods=12)``
    both build monthly labels; pass ``grain='quarter'`` (etc.) for other
    grains. The result is a ``Variable``, used as an axis via
    ``indexed_by=[time]``.
    """

    def __init__(self, start: str, end: Optional[str] = None, *,
                 periods: Optional[int] = None, grain: str = 'month',
                 display_name: str = 'Time'):
        if grain not in GRAINS:
            raise ValueError(
                f"Unknown time grain {grain!r}; supported: {', '.join(GRAINS)}."
            )
        if (end is None) == (periods is None):
            got = 'both' if end is not None else 'neither'
            raise ValueError(
                f"Time needs exactly one of `end` or `periods`; got {got}."
            )
        if periods is not None and periods <= 0:
            raise ValueError(f"Time `periods` must be positive; got {periods}.")

        labels = _period_labels(start, end, periods, grain)
        super().__init__(labels, value_type='string', var_type='list',
                         display_name=display_name)
        self._axis_kind = 'time'
        self._grain = grain
        self._start = start


# --- coarsening / bucketing (re-grain projection foundation) ----------------

def _days_in_period(anchor: Any, grain: str) -> int:
    """Calendar day-count of one period at ``grain`` — the re-grain weight.

    Periods are not equal (Feb vs Jul; leap years), so any averaging recipe
    must weight by real days, not by period count.
    """
    if grain == 'day':
        return 1
    if grain == 'month':
        year, month = anchor
        nxt = date(year + 1, 1, 1) if month == 12 else date(year, month + 1, 1)
        return (nxt - date(year, month, 1)).days
    if grain == 'quarter':
        year, quarter = anchor
        first_month = (quarter - 1) * 3 + 1
        end_year, end_month = (year + 1, 1) if quarter == 4 else (year, first_month + 3)
        return (date(end_year, end_month, 1) - date(year, first_month, 1)).days
    return (date(anchor + 1, 1, 1) - date(anchor, 1, 1)).days  # year


def period_weights(grain: str, start: str, n: int,
                   first_day: Optional[date] = None) -> List[int]:
    """Day-count weights for ``n`` consecutive periods of ``grain`` from
    ``start``. A pure function of the calendar — every series over the same
    window shares one weight vector, which is what keeps weighted recipes
    aligned across lines. A short first period (``first_day``) weighs
    only its own days."""
    anchor = _parse(start, grain)
    weights: List[int] = []
    for _ in range(n):
        weights.append(_days_in_period(anchor, grain))
        anchor = _step(anchor, grain)
    if first_day is not None and weights:
        weights[0] = time_field_series('days', start, grain, 1, first_day)[0]
    return weights

def _to_year_month(anchor: Any, grain: str) -> tuple:
    """Normalize a native-grain anchor to ``(year, month)`` (month = the
    period's first month) for coarse-label computation."""
    if grain == 'day':
        return (anchor.year, anchor.month)
    if grain == 'month':
        return anchor
    if grain == 'quarter':
        year, quarter = anchor
        return (year, (quarter - 1) * 3 + 1)
    return (anchor, 1)  # year


def _target_label(anchor: Any, native_grain: str, target_grain: str) -> str:
    """The coarse-grain label a native period falls into."""
    year, month = _to_year_month(anchor, native_grain)
    if target_grain == 'month':
        return f"{year:04d}-{month:02d}"
    if target_grain == 'quarter':
        return f"{year:04d}-Q{(month - 1) // 3 + 1}"
    if target_grain == 'year':
        return f"{year:04d}"
    raise ValueError(
        f"Cannot coarsen to {target_grain!r}; re-grain only goes to a coarser "
        f"grain (day -> month -> quarter -> year)."
    )


def bucket_ranges(native_grain: str, target_grain: str, native_start: str,
                  n: int) -> List[tuple]:
    """Group ``n`` native periods (from ``native_start``) into ``target_grain``
    buckets. Returns ``[(target_label, lo, hi), ...]`` half-open index ranges.

    The grouping is a pure function of ``(native_grain, target_grain,
    native_start)`` — independent of any Variable's values — which is what lets
    several series share one bucketing (the alignment that makes a ratio-of-sums
    correct).
    """
    if _ORDER[target_grain] <= _ORDER[native_grain]:
        raise ValueError(
            f"Re-grain coarsens only: {native_grain} -> {target_grain} is not a "
            f"coarser grain (day -> month -> quarter -> year)."
        )
    anchor = _parse(native_start, native_grain)
    labels: List[str] = []
    current = anchor
    for _ in range(n):
        labels.append(_target_label(current, native_grain, target_grain))
        current = _step(current, native_grain)
    ranges: List[tuple] = []
    i = 0
    while i < n:
        j = i
        while j < n and labels[j] == labels[i]:
            j += 1
        ranges.append((labels[i], i, j))
        i = j
    return ranges


def regrain_series(values: List[Any], native_grain: str, target_grain: str,
                   native_start: str, recipe: str,
                   first_day: Optional[date] = None, *, blank: Any = LINE_UNKNOWN) -> tuple:
    """Roll ``values`` up from ``native_grain`` to ``target_grain`` using a named
    ``recipe`` (sum / last / mean / …). Returns ``(target_labels, coarse_values)``.

    Weight-aware: each bucket carries its periods' calendar day-counts
    (:func:`period_weights`), so ``mean`` is time-weighted — a 31-day month
    outweighs a 28-day one, and non-uniform calendars change *weights*, never
    recipe code. ``blank`` is what a hole counts as when the caller knows
    the whole line (:func:`~modeleon.core.regrain.hole_value`); left out,
    ``values`` decide.
    """
    ranges = bucket_ranges(native_grain, target_grain, native_start, len(values))
    weights = period_weights(native_grain, native_start, len(values), first_day)
    if blank is LINE_UNKNOWN:
        blank = hole_value(values)
    out_labels: List[str] = []
    out_values: List[Any] = []
    for label, lo, hi in ranges:
        out_labels.append(label)
        out_values.append(apply_recipe(values[lo:hi], recipe, weights[lo:hi], blank=blank))
    return out_labels, out_values


# --- the time of a period (``mo.time``) ------------------------------------

#: The fields of ``mo.time``: the period's first and last day, its calendar
#: day and month counts and its 1-based number in the window.
TIME_FIELDS = ('start', 'end', 'days', 'months', 'index')

#: The calendar months one period of a grain holds. A day holds none: a
#: day window refuses ``mo.time.months`` (``mo.time.days`` is its measure).
MONTHS_IN_PERIOD = {'day': 0, 'month': 1, 'quarter': 3, 'year': 12}


def _first_day(anchor: Any, grain: str) -> date:
    """The first calendar day of the period ``anchor`` names."""
    if grain == 'day':
        return anchor
    if grain == 'month':
        return date(anchor[0], anchor[1], 1)
    if grain == 'quarter':
        return date(anchor[0], (anchor[1] - 1) * 3 + 1, 1)
    return date(anchor, 1, 1)  # year


def time_field_series(field: str, start: str, grain: str, n: int,
                      first_day: Optional[date] = None) -> List[Any]:
    """``mo.time.<field>`` over ``n`` periods of ``grain`` from ``start``.

    A short first period (``first_day``, the 1st of a month inside the
    first period) starts on that day: its ``days`` and ``months`` are only
    the part of the period from it; its ``end`` and every later period are
    the calendar's."""
    if field not in TIME_FIELDS:
        raise AttributeError(
            f"mo.time has no field {field!r}; it has {', '.join(TIME_FIELDS)}."
        )
    anchor = _parse(start, grain)
    out: List[Any] = []
    for i in range(n):
        first = _first_day(anchor, grain)
        days = _days_in_period(anchor, grain)
        end = first + timedelta(days=days - 1)
        months = MONTHS_IN_PERIOD[grain]
        if i == 0 and first_day is not None:
            nxt = end + timedelta(days=1)
            first, days = first_day, (end - first_day).days + 1
            months = (nxt.year - first_day.year) * 12 + nxt.month - first_day.month
        out.append({'start': first,
                    'end': end,
                    'days': days,
                    'months': months,
                    'index': i + 1}[field])
        anchor = _step(anchor, grain)
    return out


def time_variable(field: str, window: 'DefaultWindow') -> Variable:
    """A fresh Variable holding ``mo.time.<field>`` over ``window``.

    Its expression is the :class:`~modeleon.core.expr.TimeRef` itself, so
    every formula built from it carries time, not the numbers. It is
    located at the window it was read for: re-grained, aligned by date
    and written out as time, never as a row somebody typed. It has no
    unit — a day count must not stamp ``days`` on every balance it
    touches."""
    from .expr import TimeRef
    from .mv_context import _TIME_READ
    _TIME_READ[0] = True
    first_day = getattr(window, 'first_day', None)
    values = time_field_series(field, window.start, window.grain, int(window.periods),
                               first_day)
    var = Variable(
        values,
        value_type='datetime' if field in ('start', 'end') else 'int',
        var_type='list', start=window.start, grain=window.grain,
    )
    var._first_day = first_day
    var._set_expr(TimeRef(field))
    return var


def time_field_of(var: Any) -> Optional[str]:
    """The ``mo.time`` field ``var`` IS — a row whose expression is time
    itself (``h.days = mo.time.days``, or a copy / wrapper of one) — or
    ``None`` for any other row, including one that merely reads time."""
    from .expr import MethodCall, TimeRef, VarRef
    seen: set = set()
    while var is not None and id(var) not in seen:
        seen.add(id(var))
        expr = getattr(var, '_expr', None)
        if isinstance(expr, TimeRef):
            return expr.field
        if isinstance(expr, VarRef):
            var = expr.var
        elif isinstance(expr, MethodCall) and expr.method == 'copy' \
                and isinstance(expr.base, VarRef):
            var = expr.base.var
        else:
            return None
    return None


class AmbientTime:
    """``mo.time`` — the start, end, day and month counts and number of
    the period a formula is computed for.

    Time is not a row anyone writes: every line of a windowed model lives
    in periods, and any formula may read the one it is in::

        m.debt.interest = mo.lag(m.debt.balance) * rate * mo.time.days / 365
        m.opex.rent = a.rent_per_month * mo.time.months

    Each field reads the window of the innermost open ``with`` block
    (else the model built last), exactly as a period count is found for a
    loop. Where the time shows up in a workbook — a header row of the
    sheet, a sheet of its own, a row of live dates — is the workbook's
    business (``default_header``), never the formula's.
    """

    __slots__ = ()

    @property
    def start(self) -> Variable:
        """The first day of each period."""
        return self._read('start')

    @property
    def end(self) -> Variable:
        """The last day of each period."""
        return self._read('end')

    @property
    def days(self) -> Variable:
        """The calendar day count of each period (31, 28, …; 90 a quarter)."""
        return self._read('days')

    @property
    def months(self) -> Variable:
        """The calendar months in each period: 1 a month, 3 a quarter, 12 a
        year — a length of time, like ``days``: an amount a month times it
        is the period's amount on any grain."""
        return self._read('months')

    @property
    def index(self) -> Variable:
        """The period's number in the window: 1, 2, 3, …"""
        return self._read('index')

    def _read(self, field: str) -> Variable:
        from .mv_context import ambient_window
        window = ambient_window()
        if window is None:
            raise ValueError(
                f"mo.time.{field} is the time of a period, and there is no "
                f"model window here. Read it inside a model's `with` block "
                f"(mo.Model(..., default_grain='month', default_start='2026-01', "
                f"default_periods=12))."
            )
        missing = [k for k, v in (('default_start', window.start),
                                  ('default_periods', window.periods)) if not v]
        if missing:
            raise ValueError(
                f"mo.time.{field} needs the window's {' and '.join(missing)}: "
                f"set {'them' if len(missing) > 1 else 'it'} before the first "
                f"line that reads time."
            )
        if field == 'months' and window.grain == 'day':
            raise ValueError(
                "mo.time.months counts the calendar months of a period: 1 a "
                "month, 3 a quarter, 12 a year. A day window has none - read "
                "mo.time.days."
            )
        return time_variable(field, window)

    def __repr__(self) -> str:
        return "mo.time (start, end, days, months, index of the period)"


#: The time of the period a formula is computed for — see :class:`AmbientTime`.
time = AmbientTime()
