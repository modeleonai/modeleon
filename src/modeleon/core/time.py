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

from .regrain import apply_recipe
from .variable import Variable

#: Supported grains, finest to coarsest. ``week`` is a deferred grain.
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
DefaultWindow = namedtuple('DefaultWindow', ['start', 'grain', 'periods'])


def resolve_default_window(node: Any) -> Optional['DefaultWindow']:
    """The default projection window declared nearest up ``node``'s MV tree.

    Walks ``node`` then its ``_owner`` / ``_parent`` chain for the first that
    declares ``default_grain`` and returns its ``(default_start, default_grain,
    default_periods)``. ``None`` if none is declared. Consumed by ``to_excel``
    (the projection slice) as the window to project onto when the caller passes
    no explicit ``start`` / ``periods`` / ``grain``.
    """
    seen: set = set()
    n = node
    while n is not None and id(n) not in seen:
        seen.add(id(n))
        grain = getattr(n, 'default_grain', None)
        if grain is not None:
            return DefaultWindow(
                getattr(n, 'default_start', None), grain,
                getattr(n, 'default_periods', None),
            )
        n = getattr(n, '_owner', None) or getattr(n, '_parent', None)
    return None


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


def period_weights(grain: str, start: str, n: int) -> List[int]:
    """Day-count weights for ``n`` consecutive periods of ``grain`` from
    ``start``. A pure function of the calendar — every series over the same
    window shares one weight vector, which is what keeps weighted recipes
    aligned across lines."""
    anchor = _parse(start, grain)
    weights: List[int] = []
    for _ in range(n):
        weights.append(_days_in_period(anchor, grain))
        anchor = _step(anchor, grain)
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
                   native_start: str, recipe: str) -> tuple:
    """Roll ``values`` up from ``native_grain`` to ``target_grain`` using a named
    ``recipe`` (sum / last / mean / …). Returns ``(target_labels, coarse_values)``.

    Weight-aware: each bucket carries its periods' calendar day-counts
    (:func:`period_weights`), so ``mean`` is time-weighted — a 31-day month
    outweighs a 28-day one, and non-uniform calendars change *weights*, never
    recipe code.
    """
    ranges = bucket_ranges(native_grain, target_grain, native_start, len(values))
    weights = period_weights(native_grain, native_start, len(values))
    out_labels: List[str] = []
    out_values: List[Any] = []
    for label, lo, hi in ranges:
        out_labels.append(label)
        out_values.append(apply_recipe(values[lo:hi], recipe, weights[lo:hi]))
    return out_labels, out_values
