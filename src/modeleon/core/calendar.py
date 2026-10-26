# SPDX-License-Identifier: Apache-2.0
"""The calendar — rulers, and the family they belong to.

Time is a line of instants. A **ruler** is a grain laid on it: the rule
that cuts the line into periods — ``month`` cuts on the 1st of every
month, ``quarter`` on the 1st of January, April, July and October,
``day`` on every day, ``hour`` on every hour of the wall clock. A ruler answers four questions and nothing else: which
period holds a day (:meth:`Ruler.start_of`), where the next period
starts (:meth:`Ruler.next_start`), how a period is written
(:meth:`Ruler.label`) and read back (:meth:`Ruler.parse`). Everything
the engine does with time — a window's labels, the buckets a coarser
grain groups a window into, the day-count weights of a mean, the fields
of ``mo.time``, a short first period — is derived from those answers in
:mod:`modeleon.core.time`; nothing there asks what a grain is called.

A **calendar** is a family of rulers that agree on how they nest: in
the gregorian calendar a day lies in one month, a month in one quarter,
a quarter in one year, so a coarser ruler's periods are unions of whole
finer ones and a roll-up is exact. :data:`GREGORIAN` is the calendar
every model runs on; its rulers — hour, day, ISO week, month,
quarter, half, year, ISO year — are the engine's grains.

The rulers of whole days speak in ``datetime.date``. A **clock ruler**
(``sub_day``: the hour) speaks in naive ``datetime.datetime`` — the
wall clock of one fixed-offset zone, so no hour is ever skipped or
repeated; the zone itself is the window's (``default_zone``), kept with
the window for a caller that writes an instant out with its offset (a
workbook's dates carry no zone).
"""

from __future__ import annotations

import re
from datetime import date, datetime, time, timedelta
from typing import Dict, FrozenSet, Iterable, Optional, Tuple

#: Three-letter month names, 1-based (``MONTH_ABBR[1] == 'Jan'``).
MONTH_ABBR = ('', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
              'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec')

#: A day label is exactly ``YYYY-MM-DD``. ``date.fromisoformat`` also reads
#: ``2026-W01`` (an ISO week) and ``20260105``; a week label must not pass
#: as the Monday it names — a day window would then start on a date the
#: label never spells.
_DAY_LABEL = re.compile(r'^\d{4}-\d{2}-\d{2}$')

#: An hour label is exactly ``YYYY-MM-DDTHH`` — the wall-clock hour of
#: the window's zone.
_HOUR_LABEL = re.compile(r'^(\d{4})-(\d{2})-(\d{2})T(\d{2})$')


def as_day(value: date) -> date:
    """The day an instant falls on — a ``datetime`` cut to its date; a
    ``date`` is returned as it is. A ruler of whole days answers about
    days, whatever instant it is asked about."""
    return value.date() if isinstance(value, datetime) else value


def as_instant(value: date) -> datetime:
    """``value`` as a wall-clock instant: a ``date`` is its midnight."""
    if isinstance(value, datetime):
        return value
    return datetime.combine(value, time())


class Ruler:
    """One grain of a calendar: the rule that cuts the line of days into
    periods. A period is named by its first day; every answer below is
    about first days.

    ``months`` is the whole calendar months one period holds (``None``
    for a ruler finer than a month — ``mo.time.months`` is refused
    there). ``stub_unit`` names the finer ruler at whose boundaries a
    window may open INSIDE a period — a quarter or year window starting
    on the 1st of any month keeps the period's label and shortens it;
    ``None`` means a window starts only on a period's first day.
    """

    name: str = ''
    rank: int = 0
    months: Optional[int] = None
    stub_unit: Optional[str] = None
    grammar: str = ''
    #: How the grain is named to a reader, and its short form ('Monthly',
    #: 'M') — defined once here for any caller that names a grain.
    word: str = ''
    letter: str = ''
    #: A ruler of a fixed number of days per period (a day, a week):
    #: the workbook spells its time by adding days to the start cell.
    stride_days: Optional[int] = None
    #: A clock ruler: its periods are stretches of the wall clock shorter
    #: than a day, named by their first instant (a ``datetime``); its
    #: ``mo.time.end`` is where a period STOPS (13:00 → 14:00) — a period
    #: of hours has no last day.
    sub_day: bool = False
    #: Hours in one period of a clock ruler (an hour: 1): the workbook
    #: spells its time as the start cell plus i/24, one closed form.
    stride_hours: Optional[int] = None
    #: An ISO week-numbering year: 52 or 53 weeks from the Monday of the
    #: week holding 4 January — the workbook spells its bounds with
    #: ``DATE(y,1,4) - WEEKDAY(DATE(y,1,4),3)``.
    week_years: bool = False

    def refuse_start(self, value: date) -> str:
        """Why a window of this ruler cannot start on ``value`` — a day
        inside its first period that no stub may open on."""
        return (
            f"A {self.name} window starts on the 1st of a month; {value.isoformat()} "
            f"falls inside a month. Start the window on the 1st and keep the "
            f"exact day as an input — read mo.time.days for the part of a period "
            f"after it."
        )

    def parse(self, label: str) -> date:
        """The first day of the period ``label`` names, or a refusal that
        spells the grammar."""
        raise NotImplementedError

    def label(self, first: date, style: str = 'iso') -> str:
        """How the period starting on ``first`` is written: ``iso`` (the
        engine's own key), ``finance`` or ``compact``."""
        raise NotImplementedError

    def start_of(self, day: date) -> date:
        """The first day of the period that holds ``day``."""
        raise NotImplementedError

    def next_start(self, first: date) -> date:
        """The first day of the period after the one starting on ``first``."""
        raise NotImplementedError

    def measure(self, first: date) -> float:
        """The calendar day count of the period starting on ``first`` —
        the weight a time-weighted recipe uses: a whole number of days
        (an ``int``) for every ruler but a clock ruler's (1/24 an hour)."""
        return (self.next_start(first) - first).days

    def count(self, first: date, last: date) -> int:
        """Inclusive period count from the period starting on ``first``
        to the one starting on ``last``."""
        n, cur = 1, first
        while cur < last:
            cur = self.next_start(cur)
            n += 1
        return n

    def year_back(self, first: date) -> Optional[date]:
        """The first day of the period a year before the one starting on
        ``first`` — what a growth on last year compares against: the
        period holding the same date a year earlier (29 February asks
        for the 28th)."""
        try:
            back = first.replace(year=first.year - 1)
        except ValueError:
            back = first.replace(year=first.year - 1, day=28)
        return self.start_of(back)

    def __repr__(self) -> str:
        return f"Ruler({self.name!r})"


class _Hour(Ruler):
    """The hours of the wall clock: a period is ``[13:00, 14:00)``, keyed
    ``2026-10-21T13``. Its instants are naive ``datetime`` values — the
    clock of the window's fixed-offset zone, where every day has 24 hours
    (a zone that moves its clocks is refused where it is declared).
    ``mo.time.end`` is the hour's stop, ``mo.time.days`` its 1/24 of a
    day. An hour nests in its day, and so in everything a day nests in."""

    name, rank, grammar, word, letter = 'hour', 0, "'YYYY-MM-DDTHH'", 'Hourly', 'Hr'
    sub_day = True
    stride_hours = 1

    def parse(self, label: str) -> date:
        m = _HOUR_LABEL.match(label) if isinstance(label, str) else None
        if m is None:
            raise ValueError(f"Time hour start must be 'YYYY-MM-DDTHH'; got {label!r}.")
        year, month, day, hour = (int(g) for g in m.groups())
        if not 0 <= hour <= 23:
            raise ValueError(f"An hour is 00-23 in {label!r}.")
        try:
            return datetime(year, month, day, hour)
        except ValueError:
            raise ValueError(f"Time hour start must be 'YYYY-MM-DDTHH'; got {label!r}.") from None

    def label(self, first: date, style: str = 'iso') -> str:
        first = as_instant(first)
        clock = f"{first.hour:02d}:00"
        if style == 'finance':
            return f"{first.day:02d} {MONTH_ABBR[first.month]} {first.year:04d} {clock}"
        if style == 'compact':
            return f"{first.day:02d}-{MONTH_ABBR[first.month]}-{first.year % 100:02d} {clock}"
        return f"{first.year:04d}-{first.month:02d}-{first.day:02d}T{first.hour:02d}"

    def start_of(self, day: date) -> date:
        return as_instant(day).replace(minute=0, second=0, microsecond=0)

    def next_start(self, first: date) -> date:
        return as_instant(first) + timedelta(hours=1)

    def measure(self, first: date) -> float:
        return 1 / 24

    def count(self, first: date, last: date) -> int:
        return int((as_instant(last) - as_instant(first)) // timedelta(hours=1)) + 1

    def refuse_start(self, value: date) -> str:
        hour = self.start_of(value)
        return (
            f"An hour window starts on the hour; {as_instant(value).isoformat()} is "
            f"inside the hour from {hour.strftime('%H:%M')}. Start it at "
            f"{hour.isoformat()} and keep the exact moment as an input."
        )


class _Day(Ruler):
    name, rank, grammar, word, letter = 'day', 1, "'YYYY-MM-DD'", 'Daily', 'D'
    stride_days = 1

    def parse(self, label: str) -> date:
        if not isinstance(label, str) or not _DAY_LABEL.match(label):
            raise ValueError(f"Time day start must be 'YYYY-MM-DD'; got {label!r}.")
        try:
            return date.fromisoformat(label)
        except ValueError:
            raise ValueError(f"Time day start must be 'YYYY-MM-DD'; got {label!r}.") from None

    def label(self, first: date, style: str = 'iso') -> str:
        first = as_day(first)
        if style == 'finance':
            return f"{first.day:02d} {MONTH_ABBR[first.month]} {first.year:04d}"
        if style == 'compact':
            return f"{first.day:02d}-{MONTH_ABBR[first.month]}-{first.year % 100:02d}"
        return first.isoformat()

    def start_of(self, day: date) -> date:
        return as_day(day)

    def next_start(self, first: date) -> date:
        return first + timedelta(days=1)

    def count(self, first: date, last: date) -> int:
        return (last - first).days + 1


#: An ISO week label: ``2026-W05`` — the ISO year, then the week.
_WEEK_LABEL = re.compile(r'^(\d{4})-W(\d{2})$')

_WEEKDAYS = ('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday')


class _Week(Ruler):
    """ISO 8601 weeks: Monday to Sunday, numbered within their ISO year
    (week 1 holds the year's first Thursday), keyed ``2026-W05``. A week
    lies inside one ISO year and in no month, quarter or calendar year."""

    name, rank, grammar, word, letter = 'week', 2, "'YYYY-Www'", 'Weekly', 'W'
    stride_days = 7

    def parse(self, label: str) -> date:
        m = _WEEK_LABEL.match(label) if isinstance(label, str) else None
        if m is None:
            raise ValueError(f"Time week start must be 'YYYY-Www' (an ISO week); got {label!r}.")
        try:
            return date.fromisocalendar(int(m.group(1)), int(m.group(2)), 1)
        except ValueError:
            raise ValueError(
                f"{label!r} names no ISO week: the year {m.group(1)} has "
                f"{_iso_weeks_in(int(m.group(1)))} weeks."
            ) from None

    def label(self, first: date, style: str = 'iso') -> str:
        year, week, _ = first.isocalendar()
        if style == 'finance':
            return f"W{week:02d} {year:04d}"
        if style == 'compact':
            return f"W{week:02d}-{year % 100:02d}"
        return f"{year:04d}-W{week:02d}"

    def start_of(self, day: date) -> date:
        day = as_day(day)
        return day - timedelta(days=day.weekday())

    def next_start(self, first: date) -> date:
        return first + timedelta(days=7)

    def count(self, first: date, last: date) -> int:
        return (last - first).days // 7 + 1

    def year_back(self, first: date) -> Optional[date]:
        """The same week number of the ISO year before (W01 2026 against
        W01 2025); a week 53 whose year before had 52 has none."""
        year, week, _ = first.isocalendar()
        try:
            return date.fromisocalendar(year - 1, week, 1)
        except ValueError:
            return None

    def refuse_start(self, value: date) -> str:
        monday = self.start_of(value)
        return (
            f"A week window starts on a Monday; {value.isoformat()} is a "
            f"{_WEEKDAYS[value.weekday()]}. Start it on {monday.isoformat()}, the "
            f"Monday of that week, and keep the exact day as an input."
        )


def _iso_weeks_in(year: int) -> int:
    """52 or 53 — the weeks of an ISO year (28 December is always in its last)."""
    return date(year, 12, 28).isocalendar()[1]


class _IsoYear(Ruler):
    """The ISO week-numbering year: 52 or 53 whole weeks from the Monday
    of its week 1 — what a weekly sheet totals by. Keyed by its number
    (``2026``); it starts on 29 December 2025 and is not the calendar
    year, so it nests weeks and days and nothing coarser than a day."""

    name, rank, grammar, word, letter = 'iso_year', 7, "'YYYY'", 'Yearly (ISO)', 'IY'
    stub_unit = 'week'
    week_years = True

    def parse(self, label: str) -> date:
        try:
            year = int(label)
        except (ValueError, TypeError):
            raise ValueError(f"Time iso_year start must be 'YYYY'; got {label!r}.") from None
        return date.fromisocalendar(year, 1, 1)

    def label(self, first: date, style: str = 'iso') -> str:
        year = first.isocalendar()[0]
        # The ISO year of the period's first day is the period's: a stub
        # starts on a Monday inside it.
        if style == 'compact':
            return f"{year % 100:02d}"
        return f"{year:04d}"

    def start_of(self, day: date) -> date:
        return date.fromisocalendar(day.isocalendar()[0], 1, 1)

    def next_start(self, first: date) -> date:
        return date.fromisocalendar(first.isocalendar()[0] + 1, 1, 1)

    def year_back(self, first: date) -> Optional[date]:
        return date.fromisocalendar(first.isocalendar()[0] - 1, 1, 1)

    def refuse_start(self, value: date) -> str:
        return (
            f"An iso_year window starts on a Monday; {value.isoformat()} is a "
            f"{_WEEKDAYS[value.weekday()]}. Start it on the Monday of that week, "
            f"{(value - timedelta(days=value.weekday())).isoformat()}."
        )


class _Months(Ruler):
    """A ruler of ``months`` whole months, its periods opening on the
    1st of a month in step from January."""

    def start_of(self, day: date) -> date:
        k = self.months or 1
        return date(day.year, ((day.month - 1) // k) * k + 1, 1)

    def next_start(self, first: date) -> date:
        year, month = first.year, first.month + (self.months or 1)
        while month > 12:
            month, year = month - 12, year + 1
        return date(year, month, 1)

    def count(self, first: date, last: date) -> int:
        k = self.months or 1
        return ((last.year - first.year) * 12 + (last.month - first.month)) // k + 1


class _Month(_Months):
    name, rank, months, grammar, word, letter = 'month', 3, 1, "'YYYY-MM'", 'Monthly', 'M'

    def parse(self, label: str) -> date:
        parts = label.split('-') if isinstance(label, str) else []
        if len(parts) != 2:
            raise ValueError(f"Time month start must be 'YYYY-MM'; got {label!r}.")
        try:
            year, month = int(parts[0]), int(parts[1])
        except ValueError:
            raise ValueError(f"Time month start must be 'YYYY-MM'; got {label!r}.") from None
        if not 1 <= month <= 12:
            raise ValueError(f"Month must be 01-12 in {label!r}.")
        return date(year, month, 1)

    def label(self, first: date, style: str = 'iso') -> str:
        year, month = first.year, first.month
        if style == 'finance':
            return f"{MONTH_ABBR[month]} {year:04d}"
        if style == 'compact':
            return f"{MONTH_ABBR[month]}-{year % 100:02d}"
        return f"{year:04d}-{month:02d}"


class _Quarter(_Months):
    name, rank, months, stub_unit, grammar = 'quarter', 4, 3, 'month', "'YYYY-Qn'"
    word, letter = 'Quarterly', 'Q'

    def parse(self, label: str) -> date:
        parts = label.upper().split('-Q') if isinstance(label, str) else []
        if len(parts) != 2:
            raise ValueError(f"Time quarter start must be 'YYYY-Qn'; got {label!r}.")
        try:
            year, quarter = int(parts[0]), int(parts[1])
        except ValueError:
            raise ValueError(f"Time quarter start must be 'YYYY-Qn'; got {label!r}.") from None
        if not 1 <= quarter <= 4:
            raise ValueError(f"Quarter must be Q1-Q4 in {label!r}.")
        return date(year, (quarter - 1) * 3 + 1, 1)

    def label(self, first: date, style: str = 'iso') -> str:
        year, quarter = first.year, (first.month - 1) // 3 + 1
        if style == 'finance':
            return f"Q{quarter} {year:04d}"
        if style == 'compact':
            return f"Q{quarter}-{year % 100:02d}"
        return f"{year:04d}-Q{quarter}"


class _Half(_Months):
    name, rank, months, stub_unit, grammar = 'half', 5, 6, 'month', "'YYYY-Hn'"
    word, letter = 'Half-yearly', 'H'

    def parse(self, label: str) -> date:
        parts = label.upper().split('-H') if isinstance(label, str) else []
        if len(parts) != 2:
            raise ValueError(f"Time half start must be 'YYYY-Hn'; got {label!r}.")
        try:
            year, half = int(parts[0]), int(parts[1])
        except ValueError:
            raise ValueError(f"Time half start must be 'YYYY-Hn'; got {label!r}.") from None
        if not 1 <= half <= 2:
            raise ValueError(f"Half must be H1-H2 in {label!r}.")
        return date(year, (half - 1) * 6 + 1, 1)

    def label(self, first: date, style: str = 'iso') -> str:
        year, half = first.year, (first.month - 1) // 6 + 1
        if style == 'finance':
            return f"H{half} {year:04d}"
        if style == 'compact':
            return f"H{half}-{year % 100:02d}"
        return f"{year:04d}-H{half}"


class _Year(_Months):
    name, rank, months, stub_unit, grammar = 'year', 6, 12, 'month', "'YYYY'"
    word, letter = 'Yearly', 'Y'

    def parse(self, label: str) -> date:
        try:
            year = int(label)
        except (ValueError, TypeError):
            raise ValueError(f"Time year start must be 'YYYY'; got {label!r}.") from None
        return date(year, 1, 1)

    def label(self, first: date, style: str = 'iso') -> str:
        if style == 'compact':
            return f"FY{first.year % 100:02d}"
        return f"{first.year:04d}"


class Calendar:
    """A named family of rulers, finest first, and how they nest: a
    pair ``(fine, coarse)`` is in ``nesting`` when every period of
    ``coarse`` is a union of whole periods of ``fine`` — then a roll-up
    from ``fine`` to ``coarse`` is exact. Rulers that do not nest either
    way (a week and a month) never meet in a roll-up: the engine refuses
    rather than splitting a week between two months."""

    def __init__(self, name: str, rulers: Tuple[Ruler, ...],
                 nesting: Iterable[Tuple[str, str]] = ()) -> None:
        self.name = name
        self.rulers = tuple(rulers)
        self._by_name: Dict[str, Ruler] = {r.name: r for r in self.rulers}
        self._nesting: FrozenSet[Tuple[str, str]] = frozenset(nesting)

    @property
    def grains(self) -> Tuple[str, ...]:
        """The rulers' names, finest to coarsest."""
        return tuple(r.name for r in self.rulers)

    def ruler(self, name: str) -> Ruler:
        """The ruler called ``name``, or a refusal naming the family."""
        try:
            return self._by_name[name]
        except KeyError:
            raise ValueError(
                f"Unknown time grain {name!r}; supported: {', '.join(self.grains)}."
            ) from None

    def knows(self, name: object) -> bool:
        return isinstance(name, str) and name in self._by_name

    def rank(self, name: str) -> int:
        """The ruler's place in the family's order, finest first — an
        order to list them in, not a statement that they nest."""
        return self.ruler(name).rank

    def nests(self, fine: str, coarse: str) -> bool:
        """Whether every period of ``coarse`` is a union of whole
        periods of ``fine``."""
        return (fine, coarse) in self._nesting

    def coarser(self, kind: object, grain: object) -> bool:
        """Whether ``kind`` names a ruler that ``grain`` nests in — a
        bucket a ``grain`` sheet can total by; an unknown name on
        either side never is."""
        if not (self.knows(kind) and self.knows(grain)):
            return False
        return self.nests(grain, kind)   # type: ignore[arg-type]

    def comparable(self, a: str, b: str) -> bool:
        """Whether ``a`` and ``b`` are one ruler or one nests in the other."""
        return a == b or self.nests(a, b) or self.nests(b, a)

    def rolls_up_to(self, fine: str) -> Tuple[str, ...]:
        """The rulers ``fine`` nests in, finest first."""
        return tuple(r.name for r in self.rulers if self.nests(fine, r.name))

    def verify(self, span: int = 420) -> None:
        """Check every ruler and every nesting claim on a probe span of
        ``span`` periods from 2023: labels read back as the period they
        name, periods tile the line with no gap or overlap, and each
        claimed nesting holds for every fine period. A clock ruler is
        probed with its own tick (an hour) and over ``span`` days of its
        periods, so its nestings cross month and year ends. Raises naming
        the first bad period."""
        anchor = date(2023, 1, 2)          # a Monday, inside every ruler's first period

        def tick(r: Ruler) -> timedelta:
            return timedelta(hours=1) if r.sub_day else timedelta(days=1)

        def start(r: Ruler) -> date:
            return r.start_of(as_instant(anchor) if r.sub_day else anchor)

        for r in self.rulers:
            first = start(r)
            for _ in range(span):
                nxt = r.next_start(first)
                if not first < nxt:
                    raise AssertionError(f"{r.name}: the period of {first} does not end after it")
                if r.start_of(nxt - tick(r)) != first or r.start_of(nxt) != nxt:
                    raise AssertionError(f"{r.name}: the period of {first} overlaps or leaves a gap")
                if r.parse(r.label(first)) != first:
                    raise AssertionError(f"{r.name}: {r.label(first)!r} does not read back as {first}")
                first = nxt
        for fine, coarse in self._nesting:
            f, c = self.ruler(fine), self.ruler(coarse)
            first = start(f)
            for _ in range(span * 24 if f.sub_day else span):
                last = f.next_start(first) - tick(f)
                if c.start_of(first) != c.start_of(last):
                    raise AssertionError(
                        f"{fine} does not nest in {coarse}: the {fine} from {first} "
                        f"to {last} straddles two"
                    )
                first = f.next_start(first)

    def __repr__(self) -> str:
        return f"Calendar({self.name!r}, {self.grains})"


HOUR = _Hour()
DAY = _Day()
WEEK = _Week()
MONTH = _Month()
QUARTER = _Quarter()
HALF = _Half()
YEAR = _Year()
ISO_YEAR = _IsoYear()


def _closure(pairs: Iterable[Tuple[str, str]]) -> FrozenSet[Tuple[str, str]]:
    """The transitive closure of a nesting relation."""
    out = set(pairs)
    while True:
        more = {(a, d) for (a, b) in out for (c, d) in out if b == c} - out
        if not more:
            return frozenset(out)
        out |= more


#: The gregorian calendar — the one every model runs on. Hours nest in
#: their day, days in everything; months in quarters, halves and years;
#: ISO weeks only in their ISO year (a week straddles months, quarters
#: and calendar years).
GREGORIAN = Calendar(
    'gregorian',
    (HOUR, DAY, WEEK, MONTH, QUARTER, HALF, YEAR, ISO_YEAR),
    _closure([
        ('hour', 'day'),
        ('day', 'week'), ('day', 'month'), ('month', 'quarter'), ('quarter', 'half'),
        ('half', 'year'), ('week', 'iso_year'),
    ]),
)
