# SPDX-License-Identifier: Apache-2.0
"""Time — the discrete period axis and grain calendar.

``mo.Time`` builds a period-labeled axis Variable. An axis is just a
Variable (see :mod:`modeleon.core.shape`), so it rides the existing shape
machinery::

    time = mo.Time('2024-01', periods=3)           # monthly
    revenue = mo.Variable([100, 110, 121], indexed_by=[time])

Grains: ``hour`` (``YYYY-MM-DDTHH``), ``day`` (``YYYY-MM-DD``), ``week``
(``YYYY-Www``), ``month`` (``YYYY-MM``), ``quarter`` (``YYYY-Qn``),
``half`` (``YYYY-Hn``), ``year`` (``YYYY``), ``iso_year``. Extent is
``start`` plus exactly one of
``end`` or ``periods``. The per-grain calendar helpers (parse / step /
format / count) are the foundation the re-grain projection reuses.
"""

from __future__ import annotations

import functools
from collections import namedtuple
from datetime import date, datetime, time as clock_time, timedelta
from typing import Any, List, Optional

from .calendar import GREGORIAN, MONTH_ABBR, as_day
from .regrain import LINE_UNKNOWN, apply_recipe, hole_value
from .variable import Variable

#: Supported grains, finest to coarsest — the gregorian calendar's rulers
#: (:mod:`modeleon.core.calendar`).
GRAINS = GREGORIAN.grains

#: The columns one sheet holds (Excel's 16,384, ``XFD``). A period is a
#: column, so no window may be wider than a sheet; the label column takes
#: one, so a window refuses past :data:`WINDOW_MAX_PERIODS` where it is
#: declared, and the layout counts the rest (units, buckets).
SHEET_COLUMNS = 16384
WINDOW_MAX_PERIODS = SHEET_COLUMNS - 1


def article(grain: str) -> str:
    """``'an'`` before ``hour``, ``'a'`` before every other ruler's name —
    in the engine's messages."""
    return 'an' if grain == 'hour' else 'a'


class SheetWidthError(ValueError):
    """A window wider than a sheet: refused by name, never truncated."""


def width_refusal(grain: str, periods: int, lead: int = 1,
                  columns: Optional[int] = None) -> str:
    """Why ``periods`` periods of ``grain`` — ``columns`` columns with
    their totals, ``lead`` label columns before them — do not fit one
    sheet: the count, the sheet's width, what this grain fits, and the
    two ways out — a coarser grain, or the data kept as rows."""
    columns = periods if columns is None else columns
    fits = SHEET_COLUMNS - lead
    span = ''
    if GREGORIAN.knows(grain):
        ruler = GREGORIAN.ruler(grain)
        days = fits * ruler.measure(ruler.start_of(date(2026, 1, 1)))
        span = (f"; {grain}s fit about {round(days / 30.44)} months" if days < 3650
                else f"; {grain}s fit about {round(days / 365.25)} years")
    coarser = next((g for g in GREGORIAN.rolls_up_to(grain)), None) if GREGORIAN.knows(grain) else None
    by = f"View it by {coarser} (default_grain='{coarser}', or .at('{coarser}') on its lines)" \
        if coarser else "View it by a coarser grain"
    totals = " with its totals" if columns != periods else ""
    return (
        f"A window of {periods} {grain}s needs {columns} columns{totals} beside {lead} "
        f"label column{'s' if lead != 1 else ''}, and a sheet holds {SHEET_COLUMNS}{span}. "
        f"{by}, or keep the {grain}s as the rows of a table and chart them."
    )


#: A window's zone: a fixed offset from UTC, spelled ``+01:00`` — a zone
#: that moves its clocks would skip and repeat hours.
_ZONE_NAMES = {'UTC': '+00:00', 'Z': '+00:00', 'GMT': '+00:00'}


def zone_offset(value: Any) -> str:
    """``default_zone`` as ``+HH:MM``: a fixed offset (``'+01:00'``,
    ``'-03:30'``, ``'UTC'`` / ``'Z'``). A named zone (``'Europe/Paris'``) is
    refused with the offset to write instead: an hour is the wall clock of
    one zone whose clock never moves."""
    import re
    if isinstance(value, str):
        text = value.strip()
        if text.upper() in _ZONE_NAMES:
            return _ZONE_NAMES[text.upper()]
        m = re.fullmatch(r'([+-])(\d{2}):?(\d{2})', text)
        if m and int(m.group(2)) <= 14 and m.group(3) in ('00', '15', '30', '45'):
            return f"{m.group(1)}{m.group(2)}:{m.group(3)}"
    hint = ''
    if isinstance(value, str) and '/' in value:
        try:
            from zoneinfo import ZoneInfo
            now = datetime.now(ZoneInfo(value.strip()))
            off = now.utcoffset()
            if off is not None:
                minutes = int(off.total_seconds() // 60)
                sign = '+' if minutes >= 0 else '-'
                hint = (f" {value.strip()} is {sign}{abs(minutes) // 60:02d}:"
                        f"{abs(minutes) % 60:02d} today; write that, if its clocks never move.")
        except Exception:  # noqa: BLE001 — a hint only
            hint = ''
    raise ValueError(
        f"default_zone is a fixed offset from UTC - '+01:00', '-03:30' or 'UTC'; "
        f"got {value!r}. A named zone can move its clocks, and an hour window "
        f"never skips or repeats an hour.{hint}"
    )

#: Coarsening order — re-grain only goes left-to-right.
_ORDER = {name: GREGORIAN.rank(name) for name in GRAINS}

#: Period-label styles for the Excel timeline header. ``iso`` is the engine's
#: native form (matches ``mo.Time`` axis labels); ``finance`` / ``compact`` are
#: presentation variants.
TIME_LABEL_STYLES = ('iso', 'finance', 'compact')

#: Default header style when an ``ExcelView`` declares none.
DEFAULT_TIME_LABEL_STYLE = 'finance'

_MONTH_ABBR = MONTH_ABBR

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
#: ``zone`` is the window's fixed offset (``'+01:00'``) when it declares
#: one — read only where an instant leaves the engine; the engine itself
#: computes on the wall clock.
DefaultWindow = namedtuple('DefaultWindow',
                           ['start', 'grain', 'periods', 'start_source', 'first_day', 'zone'],
                           defaults=(None, None, None))


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
                slots.get('default_zone'),
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
    refused: a period's months are whole. A clock ruler (an hour) starts
    at an instant on the hour — a ``datetime`` on the wall clock of the
    window's zone, or a date (its midnight).
    """
    if isinstance(value, str):
        _parse(value, grain)
        return value
    ruler = GREGORIAN.ruler(grain) if GREGORIAN.knows(grain) else None
    if ruler is not None and ruler.sub_day and isinstance(value, date):
        if isinstance(value, datetime) and value.tzinfo is not None:
            raise ValueError(
                f"An hour window starts at a wall-clock time of its zone, not at "
                f"{value.isoformat()}: write the time without its offset and "
                f"declare the zone once - default_zone='+01:00'."
            )
        first = ruler.start_of(value)
        if first != (value if isinstance(value, datetime)
                     else datetime.combine(value, clock_time())):
            raise ValueError(ruler.refuse_start(value))
        return ruler.label(first)
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
    ruler = GREGORIAN.ruler(grain)
    first = ruler.start_of(value)
    if first == value:
        return ruler.label(first)
    if _stub_unit_start(ruler, value):
        # A stub: the first period is the rest of this quarter / year.
        return ruler.label(first)
    raise ValueError(ruler.refuse_start(value))


def _containing_label(day: date, grain: str) -> str:
    """The label of the ``grain`` period that holds ``day`` (an instant
    on a clock ruler)."""
    if not GREGORIAN.knows(grain):
        return day.isoformat()
    ruler = GREGORIAN.ruler(grain)
    return ruler.label(ruler.start_of(day))


def _has_stub(grain: Any) -> bool:
    """Whether a ``grain`` window may open inside its first period (on a
    boundary of the ruler's ``stub_unit``) — a quarter or year window
    starting on the 1st of any month."""
    return GREGORIAN.knows(grain) and GREGORIAN.ruler(grain).stub_unit is not None


def _stub_unit_start(ruler: Any, day: date) -> bool:
    """Whether ``day`` opens a period of ``ruler``'s stub unit — the 1st
    of a month under a quarter or year ruler. An instant opens one only
    at its midnight."""
    if ruler.stub_unit is None:
        return False
    if isinstance(day, datetime):
        if day.time() != clock_time():
            return False
        day = day.date()
    return GREGORIAN.ruler(ruler.stub_unit).start_of(day) == day


def spelled_first_day(window: Any) -> Optional[date]:
    """The day a workbook spells a quarter / year ``window``'s time from so
    that its first period may be a stub, or ``None`` for the whole-period
    spelling: the stub's day, or — the window starts at a date cell, which
    may be moved inside a period in the book — its first period's first
    day. A start written into the formulas cannot move: whole spelling."""
    if window is None or not _has_stub(window.grain) or window.start is None:
        return None
    if getattr(window, 'first_day', None) is not None:
        return window.first_day
    if getattr(window, 'start_source', None) is not None:
        return _first_day(_parse(window.start, window.grain), window.grain)
    return None


def _coarse_first_day(native_start: str, native_grain: str, native_first_day: Optional[date],
                      target_grain: str) -> Optional[date]:
    """The day the first ``target_grain`` period of a series re-grained from
    ``native_start`` starts on, when that is inside the coarse period — the
    coarse first period is then short, as every line and time row of the
    re-grained view must agree. ``None`` when it starts on the coarse
    period's first day, or the target is finer than a quarter."""
    if not _has_stub(target_grain) or native_start is None:
        return None
    anchor = _parse(native_start, native_grain)
    first = native_first_day or _first_day(anchor, native_grain)
    if not _stub_unit_start(GREGORIAN.ruler(target_grain), first):
        return None
    first = as_day(first)
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
    if not _has_stub(grain):
        return None
    if isinstance(value, datetime):
        value = value.date()
    if not isinstance(value, date):
        return None
    if _first_day(_parse(_containing_label(value, grain), grain), grain) == value:
        return None
    return value


def _parse(start: str, grain: str) -> Any:
    """Parse a ``start`` label into a per-grain anchor (the first period).

    The anchor is the period's first day spelled the way this module
    has always carried it — ``date`` for a day, ``(year, month)`` for a
    month, ``(year, quarter)`` for a quarter, ``year`` for a year; the
    ruler itself speaks in first days (:func:`_first_day` /
    :func:`_anchor_of` convert)."""
    return _anchor_of(GREGORIAN.ruler(grain).parse(start), grain)


def _parse_year_month(start: str) -> tuple:
    return _anchor_of(GREGORIAN.ruler('month').parse(start), 'month')


def _parse_year_quarter(start: str) -> tuple:
    return _anchor_of(GREGORIAN.ruler('quarter').parse(start), 'quarter')


def _anchor_of(first: date, grain: str) -> Any:
    """The per-grain anchor of the period starting on ``first`` — the
    inverse of :func:`_first_day`."""
    if grain == 'month':
        return (first.year, first.month)
    if grain == 'quarter':
        return (first.year, (first.month - 1) // 3 + 1)
    if grain == 'half':
        return (first.year, (first.month - 1) // 6 + 1)
    if grain == 'year':
        return first.year
    return first


def _step(anchor: Any, grain: str) -> Any:
    """Advance an anchor by one period of ``grain``."""
    ruler = GREGORIAN.ruler(grain)
    return _anchor_of(ruler.next_start(_first_day(anchor, grain)), grain)


def _format(anchor: Any, grain: str, style: str = 'iso') -> str:
    return GREGORIAN.ruler(grain).label(_first_day(anchor, grain), style)


def _count(start_anchor: Any, end_anchor: Any, grain: str) -> int:
    """Inclusive period count from ``start_anchor`` to ``end_anchor``."""
    return GREGORIAN.ruler(grain).count(_first_day(start_anchor, grain),
                                         _first_day(end_anchor, grain))


def _period_labels(start: str, end: Optional[str], periods: Optional[int],
                   grain: str, style: str = 'iso') -> List[str]:
    anchor = _parse(start, grain)
    if periods is not None:
        count = periods
    else:
        count = _count(anchor, _parse(end, grain), grain)  # type: ignore[arg-type]
        if count <= 0:
            raise ValueError(f"Time end {end!r} is before start {start!r}.")
    # Step between labels, never past the last: the period after a window
    # that ends in December 9999 is one no date can hold.
    labels: List[str] = []
    current = anchor
    for i in range(count):
        if i:
            current = _step(current, grain)
        labels.append(_format(current, grain, style))
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

def _days_in_period(anchor: Any, grain: str) -> float:
    """Calendar day-count of one period at ``grain`` — the re-grain weight
    (whole days, an ``int``; 1/24 for an hour).

    Periods are not equal (Feb vs Jul; leap years), so any averaging recipe
    must weight by real days, not by period count.
    """
    return GREGORIAN.ruler(grain).measure(_first_day(anchor, grain))


def _period_weights(grain: str, start: str, n: int,
                    first_day: Optional[date] = None) -> List[float]:
    """Day-count weights for ``n`` consecutive periods of ``grain`` from
    ``start``. A pure function of the calendar — every series over the same
    window shares one weight vector, which is what keeps weighted recipes
    aligned across lines. A short first period (``first_day``) weighs
    only its own days."""
    anchor = _parse(start, grain)
    weights: List[float] = []
    for _ in range(n):
        weights.append(_days_in_period(anchor, grain))
        anchor = _step(anchor, grain)
    if first_day is not None and weights:
        weights[0] = time_field_series('days', start, grain, 1, first_day)[0]
    return weights

def _target_label(anchor: Any, native_grain: str, target_grain: str) -> str:
    """The coarse-grain label a native period falls into — the period of
    the target ruler that holds the native period's first day."""
    if not GREGORIAN.knows(target_grain) or GREGORIAN.rank(target_grain) == 0:
        # The finest ruler is no coarser grain of anything.
        raise ValueError(
            f"Cannot coarsen to {target_grain!r}; re-grain only goes to a coarser "
            f"grain ({' -> '.join(GRAINS)})."
        )
    target = GREGORIAN.ruler(target_grain)
    return target.label(target.start_of(_first_day(anchor, native_grain)))


def regrain_refusal(native_grain: str, target_grain: str) -> str:
    """Why a ``native_grain`` line cannot be rolled up to ``target_grain``:
    the target is not coarser, or the two rulers do not nest (a week
    straddles two months) — never a guessed split."""
    up = GREGORIAN.rolls_up_to(native_grain) if GREGORIAN.knows(native_grain) else ()
    a = article(native_grain)
    to = (f"{a} {native_grain} rolls up to {', '.join(up)}" if up
          else f"{a} {native_grain} rolls up to nothing coarser")
    if (native_grain == target_grain or not GREGORIAN.knows(target_grain)
            or not GREGORIAN.knows(native_grain)
            or GREGORIAN.nests(target_grain, native_grain)):
        return (f"Re-grain coarsens only: {native_grain} -> {target_grain} is not a "
                f"coarser grain ({to}).")
    return (f"{a.capitalize()} {native_grain} does not lie inside one {target_grain}: "
            f"{native_grain}s straddle {target_grain}s, so {a} {native_grain} line cannot "
            f"be rolled up to {target_grain}s without splitting {a} {native_grain} ({to}). "
            f"Model the line in a grain both share (days), or total it by a grain it "
            f"rolls up to.")


def _bucket_ranges(native_grain: str, target_grain: str, native_start: str,
                   n: int) -> List[tuple]:
    """Group ``n`` native periods (from ``native_start``) into ``target_grain``
    buckets. Returns ``[(target_label, lo, hi), ...]`` half-open index ranges.

    The grouping is a pure function of ``(native_grain, target_grain,
    native_start)`` — independent of any Variable's values — which is what lets
    several series share one bucketing (the alignment that makes a ratio-of-sums
    correct).
    """
    if not GREGORIAN.nests(native_grain, target_grain):
        raise ValueError(regrain_refusal(native_grain, target_grain))
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


# --- the calendar, remembered -----------------------------------------------
#
# ``bucket_ranges`` / ``period_weights`` / ``coarse_first_day`` are pure
# functions of the calendar and their arguments, and a projection asks them
# the same question for every line of a model, mostly with a handful of
# distinct argument tuples. Each remembers a bounded number of answers and
# hands back a fresh list, so no caller can edit the remembered one; an
# unhashable argument takes the plain path.


@functools.lru_cache(maxsize=4096)
def _bucket_ranges_cached(native_grain: str, target_grain: str, native_start: Any,
                          n: int) -> tuple:
    return tuple(_bucket_ranges(native_grain, target_grain, native_start, n))


def bucket_ranges(native_grain: str, target_grain: str, native_start: str,
                  n: int) -> List[tuple]:
    """Group ``n`` native periods (from ``native_start``) into ``target_grain``
    buckets. Returns ``[(target_label, lo, hi), ...]`` half-open index ranges.

    The grouping is a pure function of ``(native_grain, target_grain,
    native_start)`` — independent of any Variable's values — which is what lets
    several series share one bucketing (the alignment that makes a ratio-of-sums
    correct). Remembered per argument tuple; the list is the caller's own.
    """
    try:
        return list(_bucket_ranges_cached(native_grain, target_grain, native_start, n))
    except TypeError:
        return _bucket_ranges(native_grain, target_grain, native_start, n)


@functools.lru_cache(maxsize=4096)
def _period_weights_cached(grain: str, start: Any, n: int, first_day: Optional[date]) -> tuple:
    return tuple(_period_weights(grain, start, n, first_day))


def period_weights(grain: str, start: str, n: int,
                   first_day: Optional[date] = None) -> List[float]:
    """Day-count weights for ``n`` consecutive periods of ``grain`` from
    ``start`` — see :func:`_period_weights`; remembered per argument tuple,
    the list is the caller's own."""
    try:
        return list(_period_weights_cached(grain, start, n, first_day))
    except TypeError:
        return _period_weights(grain, start, n, first_day)


@functools.lru_cache(maxsize=4096)
def _coarse_first_day_cached(native_start: Any, native_grain: str,
                             native_first_day: Optional[date], target_grain: str) -> Optional[date]:
    return _coarse_first_day(native_start, native_grain, native_first_day, target_grain)


def coarse_first_day(native_start: str, native_grain: str, native_first_day: Optional[date],
                     target_grain: str) -> Optional[date]:
    """The day the first ``target_grain`` period of a series re-grained from
    ``native_start`` starts on, when that is inside the coarse period — see
    :func:`_coarse_first_day`; remembered per argument tuple (a date is
    immutable)."""
    try:
        return _coarse_first_day_cached(native_start, native_grain, native_first_day, target_grain)
    except TypeError:
        return _coarse_first_day(native_start, native_grain, native_first_day, target_grain)


# --- the time of a period (``mo.time``) ------------------------------------

#: The fields of ``mo.time``: the period's first and last day (on a clock
#: ruler: the instants it starts and stops at), its calendar day and month
#: counts and its 1-based number in the window.
TIME_FIELDS = ('start', 'end', 'days', 'months', 'index')

#: The calendar months one period of a grain holds. A day holds none: a
#: day window refuses ``mo.time.months`` (``mo.time.days`` is its measure).
MONTHS_IN_PERIOD = {r.name: (r.months or 0) for r in GREGORIAN.rulers}


def _first_day(anchor: Any, grain: str) -> date:
    """The first calendar day of the period ``anchor`` names — the
    inverse of :func:`_anchor_of`."""
    if grain == 'month':
        return date(anchor[0], anchor[1], 1)
    if grain == 'quarter':
        return date(anchor[0], (anchor[1] - 1) * 3 + 1, 1)
    if grain == 'half':
        return date(anchor[0], (anchor[1] - 1) * 6 + 1, 1)
    if grain == 'year':
        return date(anchor, 1, 1)
    return anchor  # a day's anchor is the day


def time_field_series(field: str, start: str, grain: str, n: int,
                      first_day: Optional[date] = None) -> List[Any]:
    """``mo.time.<field>`` over ``n`` periods of ``grain`` from ``start``.

    A short first period (``first_day``, the 1st of a month inside the
    first period) starts on that day: its ``days`` and ``months`` are only
    the part of the period from it; its ``end`` and every later period are
    the calendar's. A clock ruler's period ends where it stops (13:00 →
    14:00) and lasts its fraction of a day."""
    if field not in TIME_FIELDS:
        raise AttributeError(
            f"mo.time has no field {field!r}; it has {', '.join(TIME_FIELDS)}."
        )
    anchor = _parse(start, grain)
    ruler = GREGORIAN.ruler(grain)
    out: List[Any] = []
    if ruler.sub_day:
        for i in range(n):
            first = _first_day(anchor, grain)
            out.append({'start': first,
                        'end': ruler.next_start(first),
                        'days': ruler.measure(first),
                        'months': MONTHS_IN_PERIOD[grain],
                        'index': i + 1}[field])
            anchor = _step(anchor, grain)
        return out
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


def period_bounds(start: str, grain: str, n: int,
                  first_day: Optional[date] = None) -> tuple:
    """``(starts, ends)`` — the first and last day of each of ``n``
    periods of ``grain`` from ``start``; a short first period starts on
    ``first_day``. A caller that needs each period's dates reads them here
    and does no calendar arithmetic of its own."""
    return (time_field_series('start', start, grain, n, first_day),
            time_field_series('end', start, grain, n, first_day))


def year_back_index(start: str, grain: str, n: int) -> List[Optional[int]]:
    """For each of ``n`` periods of ``grain`` from ``start``, the index
    of the period one calendar year earlier in the same window, or
    ``None`` when the window does not reach it — what a growth-on-last-
    year compares against, asked of the calendar rather than counted."""
    ruler = GREGORIAN.ruler(grain)
    labels = _period_labels(start, None, n, grain, 'iso')
    index = {label: i for i, label in enumerate(labels)}
    out: List[Optional[int]] = []
    for label in labels:
        back = ruler.year_back(ruler.parse(label))
        out.append(None if back is None else index.get(ruler.label(back)))
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
    whole = not (field == 'days' and GREGORIAN.ruler(window.grain).sub_day)
    var = Variable(
        values,
        value_type='datetime' if field in ('start', 'end') else ('int' if whole else 'float'),
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
        """The calendar day count of each period (31, 28, …; 90 a quarter;
        1/24 an hour)."""
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
        if field == 'months' and GREGORIAN.ruler(window.grain).months is None:
            raise ValueError(
                "mo.time.months counts the calendar months of a period: 1 a "
                f"month, 3 a quarter, 12 a year. {article(window.grain).capitalize()} "
                f"{window.grain} window has none - read mo.time.days."
            )
        return time_variable(field, window)

    def __repr__(self) -> str:
        return "mo.time (start, end, days, months, index of the period)"


#: The time of the period a formula is computed for — see :class:`AmbientTime`.
time = AmbientTime()
