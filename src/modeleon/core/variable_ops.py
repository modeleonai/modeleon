# SPDX-License-Identifier: Apache-2.0
"""Arithmetic and comparison mixin for :class:`Variable`.

Split out from ``variable.py`` to keep the main file focused on
identity, construction, and representation. This module owns:

- Broadcasting (scalar/list, list/list with length-1 fallback)
- Safe-op semantics (Excel-style ``#DIV/0!`` and error propagation)
- Key alignment and propagation for keyed (dict-sourced) Variables
- Unit propagation through arithmetic
- The shared ``_binary_op`` / ``_comparison_op`` scaffolds
- All binary, unary, and comparison dunders

``_VariableArithmetic`` is a **mixin** — it assumes ``self`` is a
:class:`Variable` and relies on ``self._value``, ``self._keys``,
``self._as_operand``, plus the ``expr`` property setter (which updates
``_expr`` and ``_dependency_refs`` together). Variable inherits from
this class so the operator surface is available everywhere Variable
is, while the file on disk stays coherent.

:class:`Variable` is imported lazily, once (``_variable_class``), to
avoid the circular import that would arise if we imported at module
scope (``variable.py`` imports this module; this module needs Variable
to construct results).
"""

from __future__ import annotations

import functools
import math
import os
import weakref
from datetime import date, datetime, time, timedelta
from typing import Any, List, Optional, Sequence, Tuple, TYPE_CHECKING

from .expr import BinOp, Compare, Expr, Literal, UnaryOp, VarRef
from .shape import keys_along
from .tracks import TrackValues

if TYPE_CHECKING:
    from .unit import Unit
    from .variable import Variable


#: :class:`~modeleon.core.variable.Variable`, bound on first use.
#: ``variable.py`` imports this module before it defines the class, so it
#: cannot be imported at load time; a function-level import, though, looks
#: the module up again on every call, which costs more than the arithmetic
#: of a scalar operator. Every operator below reads it.
_Variable: Optional[type[Variable]] = None


def _variable_class() -> type[Variable]:
    """The :class:`Variable` class, imported once (see ``_Variable``)."""
    global _Variable
    if _Variable is None:
        from .variable import Variable
        _Variable = Variable
    return _Variable


# Excel's day 0: a date's serial number counts the days since then.
_EXCEL_EPOCH_ORDINAL = date(1899, 12, 30).toordinal()
_ONE_DAY = timedelta(days=1)
_COMPARISON_SYMBOLS = frozenset({'==', '!=', '<', '<=', '>', '>='})


#: A blank compares as "" against text, as Excel's ``A1=""`` does.
_COMPARISONS = frozenset(('==', '!=', '<', '<=', '>', '>='))

def _to_serial(x: Any) -> Any:
    """A date as Excel's serial number; a timedelta as its days.

    A ``date`` is a whole number of days since 1899-12-30. A ``datetime``
    adds its time of day as a fraction, so it is always a float (noon on
    2024-03-31 is 45382.5). Anything else comes back unchanged.
    """
    if isinstance(x, datetime):
        seconds = x.hour * 3600 + x.minute * 60 + x.second + x.microsecond / 1e6
        return x.toordinal() - _EXCEL_EPOCH_ORDINAL + seconds / 86400
    if isinstance(x, date):
        return x.toordinal() - _EXCEL_EPOCH_ORDINAL
    if isinstance(x, timedelta):
        return x / _ONE_DAY
    return x


def _date_compare(a: Any, b: Any, op: Any) -> Any:
    """Compare the way Excel does when one side is a date.

    A date compares as its serial number, so it can be compared with a
    plain number too. Across types Excel never finds two values equal and
    orders every number (a date included) before any text, and text before
    TRUE/FALSE. The result is a Python ``bool``.
    """
    def key(x: Any) -> Any:
        if isinstance(x, bool):
            return (2, x)
        if isinstance(x, str):
            return (1, x)
        serial = _to_serial(x)
        if isinstance(serial, (int, float)):
            return (0, serial)
        return None

    ka, kb = key(a), key(b)
    if ka is None or kb is None:
        return '#VALUE!'
    return bool(op(ka, kb))


def _date_serial_op(a: Any, b: Any, op: Any, op_symbol: Optional[str]) -> Any:
    """Excel-style arithmetic and comparisons where a date is involved.

    Excel stores a date as its serial number (see ``_to_serial``) and
    every operator works on that number. This mirrors it on Python
    ``date`` / ``datetime`` values, so Python computes what the written
    workbook computes:

    - ``date ± number`` (or ``number + date``) → the date that many days
      later or earlier. A fraction of a day gives a ``datetime``. A
      ``timedelta`` counts as its days, fraction included.
    - ``date - date`` → the days between them: an ``int`` for two pure
      dates, a ``float`` when a ``datetime`` brings a time of day.
    - any other operation (``date * n``, ``n / date``, ``n - date``,
      ``date + date``, ``date % 7``, …) → a plain number computed on
      the serial, as Excel does.
    - comparisons → ``True`` / ``False`` (see ``_date_compare``).
    - text, or a day count that no date can hold → ``#VALUE!``.
    """
    if op_symbol in _COMPARISON_SYMBOLS:
        return _date_compare(a, b, op)
    a_is_date, b_is_date = isinstance(a, date), isinstance(b, date)
    shifts = ((op_symbol == '+' and a_is_date != b_is_date)
              or (op_symbol == '-' and a_is_date and not b_is_date))
    if shifts:
        when, offset = (a, b) if a_is_date else (b, a)
        try:
            if isinstance(offset, timedelta):
                step = offset
            elif isinstance(offset, (int, float)):
                step = timedelta(days=offset)
            else:
                return '#VALUE!'
            if op_symbol == '-':
                step = -step
            # A pure date moved by part of a day gains a time of day.
            if not isinstance(when, datetime) and step % _ONE_DAY:
                when = datetime.combine(when, time())
            return when + step
        except (ValueError, OverflowError):
            return '#VALUE!'

    sa, sb = _to_serial(a), _to_serial(b)
    if not isinstance(sa, (int, float)) or not isinstance(sb, (int, float)):
        return '#VALUE!'
    try:
        return op(sa, sb)
    except ZeroDivisionError:
        # A zero serial (1899-12-30, or an empty timedelta) divides like 0.
        return '#DIV/0!'
    except Exception:
        return '#VALUE!'


def _on_open_loop(v: Any) -> bool:
    """A row read before its line, or built on one: its loop has not
    closed (see :mod:`modeleon.core.loops`), so its unit is provisional."""
    return bool(getattr(v, '_awaits', None)) or getattr(v, '_forward_of', None) is not None


def _provisional_unit(unit_fn: Any, left: Any, right: Any, *operands: Any) -> Any:
    """``unit_fn(left, right)`` — except over a row whose loop has not
    closed: its unit is not known yet, so a mismatch is not an error
    here; the loop derives every unit again when it closes, and reports
    a real mismatch there."""
    if any(_on_open_loop(v) for v in operands):
        try:
            return unit_fn(left, right)
        except ValueError:
            return None
    return unit_fn(left, right)


def _grain_rank(grain: Any) -> int:
    """The calendar's rank of ``grain`` (finest 0, the hour); an unknown
    word ranks with the finest. Read lazily: ``core.time`` imports this module's
    neighbours."""
    from .calendar import GREGORIAN
    return GREGORIAN.rank(grain) if GREGORIAN.knows(grain) else 0


def _mixed_first_periods(left: Any, lt: Any, right: Any, rt: Any) -> str:
    """The refusal when two series start the same-named periods on
    different days: one window's first period is short (it starts inside
    its quarter or year), the other's is not or starts on another day."""
    def name(v: Any) -> str:
        label = v._display_name or getattr(v, '_python_name', None)
        if label:
            return repr(label)
        formula = getattr(v, 'formula', None)
        return f"the line {formula}" if formula else "an unnamed line"

    def first(v: Any, t: Any) -> str:
        if t.first_day is not None:
            return (f"{name(v)} starts its first period, {t.start}, on "
                    f"{t.first_day.isoformat()}")
        return f"{name(v)} starts with the whole period {t.start}"
    return (
        f"{name(left)} and {name(right)} do not line up: {first(left, lt)}; "
        f"{first(right, rt)}. Periods of the same name would hold different "
        f"days - build both in the same window."
    )


def _mixed_grains(left: Any, left_grain: str, right: Any, right_grain: str) -> str:
    """Two lines of different grains in one formula: name both, their
    grains, and the re-grain that joins them."""
    adjective = {'hour': 'hourly', 'day': 'daily', 'week': 'weekly', 'month': 'monthly',
                 'quarter': 'quarterly', 'half': 'half-yearly', 'year': 'yearly',
                 'iso_year': 'ISO-yearly'}

    def describe(v: Any, grain: str) -> str:
        label = (getattr(v, '_display_name', None)
                 or getattr(v, '_python_name', None))
        word = adjective.get(grain, grain)
        return f"{label!r} ({grain}s)" if label else f"{'an' if word == 'hourly' else 'a'} {word} line"

    from .calendar import GREGORIAN
    if (GREGORIAN.knows(left_grain) and GREGORIAN.knows(right_grain)
            and not GREGORIAN.comparable(left_grain, right_grain)):
        # Weeks and months: neither rolls up into the other.
        return (
            f"{describe(left, left_grain)} and {describe(right, right_grain)} meet "
            f"in one formula at grains that do not nest: {left_grain}s straddle "
            f"{right_grain}s, so neither line can be re-grained into the other "
            f"without splitting a period. Build both at a grain they share "
            f"(days), or keep them apart."
        )
    finer = left if _grain_rank(left_grain) < _grain_rank(right_grain) else right
    coarse = right_grain if finer is left else left_grain
    code = getattr(finer, '_python_name', None)
    hint = f"{code}.at('{coarse}')" if code else f".at('{coarse}')"
    return (
        f"{describe(left, left_grain)} and {describe(right, right_grain)} meet "
        f"in one formula at different grains. Re-grain the finer line first: "
        f"{hint} gives it in {coarse}s (date alignment only aligns within one "
        f"grain)."
    )


def _onto_window(var: Any, vals: list, off: int, total: int, union_start: str,
                 grain: str, union_first: Any, other: str) -> list:
    """``vals`` of ``var`` materialized onto a combined window of ``total``
    periods from ``union_start``, its own first value ``off`` periods in:
    the cells outside its own extent follow its DECLARED ``extend=``
    rule — zero (a flow is absent), hold (a rate stays in force; held
    backward from the first value when extended into the past), or
    none/undeclared → an error naming both extents and the rules to declare (``other``
    spells the other side's). Zero continues numbers: a line of dates or
    text it would continue with 0 is refused, pointing to hold. A time
    field is defined in every period, not only the window's."""
    from .expr import TimeRef
    time_expr = getattr(var, '_expr', None)
    if isinstance(time_expr, TimeRef):
        field = time_expr.field
        if field == 'index':
            return [i - off + 1 for i in range(total)]
        from .time import time_field_series
        return time_field_series(field, union_start, grain, total, union_first)
    rule = getattr(var, '_extend', None)
    if off == 0 and len(vals) == total:
        return list(vals)
    kind = _not_numbers(vals) if rule == 'zero' else None
    if kind is not None:
        plural, one = kind
        raise ValueError(
            f"{_line_name(var)} holds {plural}, and extend=mo.zero() would continue it "
            f"with 0 — no {one} — in the periods of the combined window it lacks "
            f"(it spans {len(vals)} period(s) from {var.time.start}; the other "
            f"side spans {other}). Hold its {one} there instead: extend=mo.hold()."
        )
    out = []
    for i in range(total):
        j = i - off
        if 0 <= j < len(vals):
            out.append(vals[j])
        elif rule == 'zero':
            out.append(0.0)
        elif rule == 'hold':
            out.append(vals[0] if j < 0 else vals[-1])
        else:
            raise ValueError(_uncovered(var, len(vals), var.time.start, other))
    return out


def _uncovered(var: Any, n: int, start: str, other: str, reads: bool = False) -> str:
    """The refusal of a series with no rule that does not cover the
    combined window: a line is told to declare one; an unnamed formula,
    which cannot (``reads``: placed by the lines it reads), to become a
    line that does."""
    extent = (f"it spans {n} period(s) from {start}"
              f"{', the periods of the lines it reads' if reads else ''}; "
              f"the other side spans {other}")
    named = var._display_name or getattr(var, '_python_name', None)
    if named or var._expr is None:
        return (f"{_line_name(var)} does not cover the combined window ({extent}). "
                f"Say what continues it: extend=mo.zero() (a flow) or extend=mo.hold() "
                f"(a rate) — never a silent guess.{_regrained(var)}")
    home = _home_of(var) if reads else None
    where = (f"a line of {home[0]}, the container {_line_name(home[1])} runs on, "
             if home else "a line ")
    line = f"{home[0]}.<name> = " if home else ""
    return (
        f"{_line_name(var)} does not cover the combined window ({extent}). A formula "
        f"carries no extend= rule: make it {where}that says what continues it — "
        f"{line}mo.Variable({_formula_code(var)}, extend=mo.zero()) (a flow) or "
        f"extend=mo.hold() (a rate) — never a silent guess."
    )


def _home_of(var: Any) -> Optional[Tuple[str, Any]]:
    """The container a formula's periods come from — the path of the one
    the first located line it reads stands in — with that line; None
    where none stands in one."""
    pending = [var]
    seen = set()
    while pending:
        node = pending.pop()
        for ref in node._expr.iter_refs():
            if id(ref) in seen:
                continue
            seen.add(id(ref))
            place = _period_ref(ref)
            if place == 'formula':
                pending.append(ref)
            elif place is not None and place != _NO_WINDOW_YET:
                owner = getattr(ref, '_owner', None) or _forward_home(ref)
                return (str(owner.path), ref) if owner is not None else None
    return None


def _regrained(var: Any) -> str:
    """For a re-grained copy of a line with an ``extend=`` rule — the copy
    keeps none: the recipe of a partial coarse period is not the rule's —
    the way to read it by date."""
    from .expr import Regrain
    expr = getattr(var, '_expr', None)
    source = expr.source if isinstance(expr, Regrain) else None
    rule = getattr(source, '_extend', None)
    if rule not in ('zero', 'hold'):
        return ""
    return (
        f" A re-grained line keeps no rule of the line it re-grains "
        f"({_line_name(source)} declares extend=mo.{rule}()): combine the lines "
        f"at their own grain and re-grain the result (.at('{var._grain}')), or "
        f"give the re-grained line a rule of its own."
    )


def _not_numbers(vals: list) -> Optional[Tuple[str, str]]:
    """``('dates', 'date')`` or ``('text', 'text')`` for a series of dates
    or text — its first value neither blank nor an error says — else
    None."""
    for v in vals:
        if v is None or (isinstance(v, str) and v.startswith('#')):
            continue
        if isinstance(v, (date, datetime)):
            return ('dates', 'date')
        if isinstance(v, str):
            return ('text', 'text')
        return None
    return None


# ─── Lines on other time windows ───────────────────────────────────
#
# One reading for the operators and for every function that is one call
# per period (``functions._helpers.align_located``): series on other
# windows meet by DATE, each continued by its own ``extend=`` rule.

#: :func:`_reads_on` of a formula whose periods have no one date.
_NO_ONE_WINDOW = 'no one window'
#: :func:`_place_of` of a series that has no window yet — an input not
#: adopted, a formula over such inputs: read by position until it has one.
_NO_WINDOW_YET = 'no window yet'
_UNREAD = object()
#: Counts the changes of windows (:func:`_windows_changed`): a placement
#: kept on a formula at an earlier count was derived from windows that
#: may have moved since, and is walked again. The count starts at random,
#: so a placement unpickled from another process is never trusted.
_WINDOWS_EPOCH = int.from_bytes(os.urandom(6), 'big')


class _Readers(list):
    """The formulas placed through a line, held weakly
    (:func:`_placed_through`). A copy or a pickle of the line starts with
    none, and the placements kept before it are walked again
    (:func:`_readers_anew`): a copied formula's placement was derived
    through the line's original, whose adoption the copy never hears of."""

    __slots__ = ()

    def __reduce__(self) -> Tuple[Any, ...]:
        return (_readers_anew, ())


def _readers_anew() -> "_Readers":
    _windows_changed()
    return _Readers()


@functools.lru_cache(maxsize=4096)
def _starts_on(label: str, grain: str) -> Any:
    """The first day of the period ``label`` names: windows are ordered by
    date, never as text (``'2026-9'`` comes before ``'2026-10'``)."""
    from .calendar import GREGORIAN
    return GREGORIAN.ruler(grain).parse(label)


def _forward_home(line: Any) -> Any:
    """The loop container a row read before its line stands in for (a
    placeholder), or None."""
    forward = line.__dict__.get('_forward_of')
    return forward[0] if forward else None


def _window_of(line: Any) -> Any:
    """Where ``line`` runs: its ``time``; for a row of a loop container
    read before its line, the container's window, which the row holds a
    period of for every period."""
    loc = line.time
    if loc is None:
        home = _forward_home(line)
        if home is not None and isinstance(line._value, list):
            from .time import resolve_default_window
            from .variable import TimeLoc
            window = resolve_default_window(home)
            if window is not None and window.grain is not None:
                return TimeLoc(window.start, window.grain, window.first_day)
    return loc


def _series_length(value: Any) -> Optional[int]:
    """How many periods a value runs along: a list's length, a tracked
    value's track length; None for a scalar."""
    if isinstance(value, list):
        return len(value)
    if isinstance(value, TrackValues):
        return value.time_length
    return None


def _a_constant(var: Any) -> bool:
    """A one-value line on a window of more periods than one: the extent
    check admits it as a constant, read in every period as a number is. One located on
    its own (``start=``, a re-grain), continued by a rule, or on a
    one-period window holds that one period."""
    if var._grain is not None or var._extend is not None or var._schedule is not None:
        return False
    from .time import resolve_default_window
    owner = (getattr(var, '_owner', None) or getattr(var, '_parent', None)
             or _forward_home(var))
    window = resolve_default_window(owner) if owner is not None else None
    return window is None or window.periods != 1


def _period_ref(ref: Any) -> Any:
    """How a line a formula reads places it: ``(loc, n)`` for a series of
    ``n`` periods on a window, ``'formula'`` for an unnamed formula read
    through, ``_NO_WINDOW_YET`` for an input with no window yet, None for
    nothing that runs on a window (a scalar, a line along keys or an
    axis, a constant)."""
    value = ref._value
    n: Optional[int]
    if isinstance(value, list):
        n = len(value)
        if not n or ref._keys:
            return None
    elif isinstance(value, TrackValues):
        n = value.time_length
        if not n:
            return None
    else:
        return None
    if ref._indexed_by:
        return None
    loc = _window_of(ref)
    if loc is not None and loc.start is not None:
        return None if n == 1 and _a_constant(ref) else (loc, n)
    return 'formula' if ref._expr is not None else _NO_WINDOW_YET


def _same_window(a: Any, b: Any) -> bool:
    """Two places, ``(loc, n)``, on one window: one grain, period count
    and first day, and starts that name one period."""
    (la, na), (lb, nb) = a, b
    if la.grain != lb.grain or na != nb or la.first_day != lb.first_day:
        return False
    return la.start == lb.start or _starts_on(la.start, la.grain) == _starts_on(lb.start, lb.grain)


def _windows_changed() -> None:
    """A window moved — a container's ``default_start``, ``default_grain``
    or ``default_periods`` set, a line or a container detached: every
    placement kept on a formula (:func:`_reads_on`) is walked again when
    it is next read."""
    global _WINDOWS_EPOCH
    _WINDOWS_EPOCH += 1


def _kept_placement(var: Any) -> Any:
    """The placement kept on the unnamed formula ``var`` (:func:`_reads_on`,
    :func:`_keep_runs_on`) — or ``_UNREAD``: none kept, or one kept before
    a window moved."""
    kept = var.__dict__.get('_reads_on')
    if kept is not None and kept[1] == _WINDOWS_EPOCH:
        return kept[0]
    return _UNREAD


def _placed_through(line: Any, node: Any) -> None:
    """Note on ``line`` — a series with no window yet, an unnamed formula
    — that the placement kept on ``node`` was derived through it: once
    ``line`` is adopted it runs on a window, and ``node`` is placed again
    (:func:`_forget_placements`)."""
    readers = line.__dict__.get('_placed_through')
    if readers is None:
        readers = line.__dict__['_placed_through'] = _Readers()
    readers.append(weakref.ref(node))
    n = len(readers)
    if n >= 64 and not (n & (n - 1)):
        # Formulas discarded since, dropped each time the list doubles.
        readers[:] = [ref for ref in readers if ref() is not None]


def _forget_placements(line: Any) -> None:
    """``line`` was adopted: drop the placement kept on every formula
    placed through it (:func:`_placed_through`), and through those — each
    is walked again, on the window ``line`` runs on now."""
    stack = [line]
    seen: set = set()
    while stack:
        readers = stack.pop().__dict__.pop('_placed_through', None)
        if not readers:
            continue
        for ref in readers:
            node = ref()
            if node is not None and id(node) not in seen:
                seen.add(id(node))
                node.__dict__.pop('_reads_on', None)
                stack.append(node)


def _scan(node: Any, found: dict, open_: set) -> Tuple[Any, list]:
    """One pass over the lines the unnamed formula ``node`` reads:
    ``(place, [])`` — its :func:`_reads_on` — or ``(None, pending)``, the
    unnamed formulas it reads that are not placed yet (``found``, by id,
    or kept on them); one still open (``open_``) reads ``node`` itself
    and places nothing. Placed, ``node`` is noted on each line it was
    placed through (:func:`_placed_through`)."""
    window: Any = None
    pending: list = []
    through: list = []
    for ref in node._expr.iter_refs():
        place = _period_ref(ref)
        if place == 'formula':
            place = found.get(id(ref), _kept_placement(ref))
            if place is _UNREAD:
                if id(ref) not in open_:
                    pending.append(ref)
                continue
            through.append(ref)
        elif place == _NO_WINDOW_YET:
            through.append(ref)
            continue
        if place is None:
            continue
        if place == _NO_ONE_WINDOW:
            window = _NO_ONE_WINDOW
            break
        if window is None:
            window = place
        elif place != window and not _same_window(place, window):
            window = _NO_ONE_WINDOW
            break
    if window != _NO_ONE_WINDOW:
        if pending:
            return None, pending
        if window is not None and _series_length(node._value) != window[1]:
            window = _NO_ONE_WINDOW
    for ref in through:
        _placed_through(ref, node)
    return window, []


def _reads_on(var: Any) -> Any:
    """Where an unnamed formula's periods run: ``(loc, n)``, the window
    every located series it reads runs on — through the unnamed formulas
    it reads — when they share one and the formula holds one value per
    period of it; None when it reads none (its periods are positions
    only); ``_NO_ONE_WINDOW`` otherwise (lines of other windows paired by
    position, or another count of values). Kept on each formula placed,
    so a chain of formulas is walked once — for as long as the windows it
    was derived from stand: the placement is dropped when a line it was
    derived through is adopted (:func:`_forget_placements`) and walked
    again after any window moved (:func:`_windows_changed`), so a formula
    is placed the same whatever the order the model was built in."""
    known = _kept_placement(var)
    if known is not _UNREAD:
        return known
    found: dict = {}
    open_: set = set()
    stack = [var]
    while stack:
        node = stack.pop()
        if id(node) in found:
            continue
        place, pending = _scan(node, found, open_)
        if pending:
            open_.add(id(node))
            stack.append(node)
            stack.extend(pending)
            continue
        open_.discard(id(node))
        found[id(node)] = place
        if not _on_open_loop(node):
            node.__dict__['_reads_on'] = (place, _WINDOWS_EPOCH)
    return found[id(var)]


def _place_of(line: Any, value: Any, homes: Optional[dict] = None) -> Any:
    """Where the series ``value`` of the operand ``line`` runs, to read it
    by date: ``(loc, n, own)`` — the window and its period count, the
    line's own (``own``) or, for an unnamed formula, the one of the lines
    it reads (:func:`_reads_on`); None where it runs on no window (a
    scalar, a value typed in, a line along keys or an axis, a constant:
    read by position); ``_NO_WINDOW_YET`` for a series that has none yet
    (an input not adopted, a formula that reads no located line: read by
    position until it has one); ``_NO_ONE_WINDOW`` for a formula whose
    periods have no one date. ``homes`` keeps, by container, the window
    its lines share, so one call resolves it once."""
    n: Optional[int]
    if isinstance(value, list):
        n = len(value)
        if not n:
            return None
    elif isinstance(value, TrackValues):
        n = value.time_length
        if not n:
            return None
    else:
        return None
    if not isinstance(line, _Variable or _variable_class()):
        return None
    if line._indexed_by or (line._keys and isinstance(value, list)):
        return None
    home = None if line._grain is not None else (
        line.__dict__.get('_owner') or line.__dict__.get('_parent'))
    if homes is not None and home is not None:
        loc = homes.get(id(home), _UNREAD)
        if loc is _UNREAD:
            loc = homes[id(home)] = line.time
    elif home is None and line._grain is None and line._expr is not None:
        # An unnamed formula stands in no container: no window of its own.
        loc = _window_of(line) if '_forward_of' in line.__dict__ else None
    else:
        loc = _window_of(line)
    if loc is not None and loc.start is not None:
        if n == 1 and _a_constant(line):
            return None
        return (loc, n, True)
    if line._expr is None:
        return _NO_WINDOW_YET
    read = _reads_on(line)
    if read is None:
        return _NO_WINDOW_YET
    if read == _NO_ONE_WINDOW:
        return read
    return (read[0], n, False) if read[1] == n else _NO_ONE_WINDOW


def _keep_runs_on(result: Any, place: Any, through: Sequence[Any]) -> None:
    """Keep on an operator's ``result`` read by position where its periods
    run — the window its operands share (:func:`_meet`) — so a formula
    built on it is placed without walking back through it; noted on the
    operands it was derived ``through`` (:func:`_placed_through`)."""
    if place is None:
        return
    fields = result.__dict__
    if place == _NO_ONE_WINDOW or _series_length(fields['_value']) == place[1]:
        fields['_reads_on'] = (place, _WINDOWS_EPOCH)
        for line in through:
            _placed_through(line, result)


def _formula_code(line: Any) -> str:
    formula = getattr(line, 'formula', None)
    return formula if formula else "..."


def _off_its_window(line: Any, loc: Any, n: int, own: bool, other: str) -> str:
    """The refusal of a series that holds no rule — an unnamed formula
    placed by the lines it reads, a tracked line — and does not cover the
    combined window."""
    if own:
        return (
            f"{_line_name(line)} holds tracks on {n} period(s) from {loc.start}, "
            f"and the other side spans {other}: a tracked line is read on its own "
            f"window only — the lines beside it must lie within it."
        )
    return _uncovered(line, n, loc.start, other, reads=True)


def _no_one_window(line: Any, total: int, start: str) -> str:
    return (
        f"{_line_name(line)} has no one window to read by date: it pairs lines of "
        f"other windows by position, or holds another count of values than the "
        f"lines it reads, while the other lines meet on {total} period(s) from "
        f"{start}. Make it a line on the window it runs on — "
        f"mo.Variable({_formula_code(line)}, start=..., grain=...)."
    )


def _on_dates(lines: Sequence[Any], values: list) -> Tuple[list, Any]:
    """``values`` — ``values[i]`` the value of the operand ``lines[i]`` —
    read as ``a + b`` reads two windows: each series on a window of its
    own materialized onto the combined window by its declared ``extend=``
    rule (:func:`_onto_window`). Returns the values and the combined
    window, a ``TimeLoc`` — or the values as given and None when fewer
    than two series run on windows, or all on one (the positional zip is
    then by date). An unnamed formula, placed by the lines it reads
    (:func:`_place_of`), and a tracked line hold no rule: each must cover
    the combined window. Two grains, a short first period inside the
    combined window and a formula whose periods have no one date are
    refused."""
    out, window, _, _ = _meet(lines, values)
    return out, window


def _meet(lines: Sequence[Any], values: list) -> Tuple[list, Any, Any, list]:
    """:func:`_on_dates`, with where a result read by position runs: the
    window its dated operands share, ``(loc, n)`` — ``_NO_ONE_WINDOW``
    beside a formula whose periods have no one date, None beside none —
    or the combined window when one was formed; and the operands that
    placement was derived through (unnamed formulas, series with no
    window yet), for :func:`_keep_runs_on`."""
    dated: List[Tuple[int, Any, Any, int, bool]] = []
    adrift: Any = None
    through: list = []
    homes: dict = {}
    for i, line in enumerate(lines):
        place = _place_of(line, values[i], homes)
        if place is None:
            continue
        if place == _NO_WINDOW_YET:
            through.append(line)
            continue
        if place == _NO_ONE_WINDOW:
            adrift = line if adrift is None else adrift
            through.append(line)
            continue
        loc, n, own = place
        if not own:
            through.append(line)
        dated.append((i, line, loc, n, own))
    if not dated:
        return values, None, (None if adrift is None else _NO_ONE_WINDOW), through
    _, first, t0, n0, _ = dated[0]
    if all((loc == t0 and n == n0) or _same_window((loc, n), (t0, n0))
           for _, _, loc, n, _ in dated[1:]):
        # One window — positions are dates.
        return values, None, ((t0, n0) if adrift is None else _NO_ONE_WINDOW), through
    days: dict = {}
    for _, _, loc, _, _ in dated:
        days.setdefault(loc.grain, set()).add(loc.first_day)
    if any(len(d) > 1 for d in days.values()):
        for j, (_, left, lt, _, _) in enumerate(dated):
            for _, right, rt, _, _ in dated[j + 1:]:
                if lt.grain == rt.grain and lt.first_day != rt.first_day:
                    # One side's first period is short. It may open the
                    # combined window, never sit inside it: on the other
                    # side the same period starts on another day.
                    st, ot = (lt, rt) if lt.first_day is not None else (rt, lt)
                    if (ot.first_day is not None or _starts_on(st.start, st.grain)
                            >= _starts_on(ot.start, ot.grain)):
                        raise ValueError(_mixed_first_periods(left, lt, right, rt))
    for _, line, loc, _, _ in dated[1:]:
        if loc.grain != t0.grain:
            raise ValueError(_mixed_grains(first, t0.grain, line, loc.grain))
    from .time import _count, _parse
    from .variable import TimeLoc
    grain = t0.grain
    starts = [_starts_on(loc.start, grain) for _, _, loc, _, _ in dated]
    k0 = min(range(len(dated)), key=starts.__getitem__)
    union = dated[k0][2]
    anchor = _parse(union.start, grain)
    offsets = [0 if s == starts[k0] else _count(anchor, _parse(loc.start, grain), grain) - 1
               for s, (_, _, loc, _, _) in zip(starts, dated)]
    total = max(off + n for off, (_, _, _, n, _) in zip(offsets, dated))
    if adrift is not None:
        raise ValueError(_no_one_window(adrift, total, union.start))
    out = list(values)
    for k, (off, (i, line, loc, n, own)) in enumerate(zip(offsets, dated)):
        if len(dated) == 2:
            _, _, other_loc, other_n, _ = dated[1 - k]
            other = f"{other_n} period(s) from {other_loc.start}"
        else:
            other = f"{total} period(s) from {union.start} between them"
        if not own or isinstance(values[i], TrackValues):
            if off or n != total:
                raise ValueError(_off_its_window(line, loc, n, own, other))
            continue
        out[i] = _onto_window(line, values[i], off, total, union.start, grain,
                              union.first_day, other)
    window = TimeLoc(union.start, grain, union.first_day)
    return out, window, window, through


def _on_loop_window(var: Any, x: list, row: Any, loop: Any, n: int) -> Optional[list]:
    """The series ``x`` of the line ``var``, read inside a loop container
    whose rows (``row`` one of them) run on ``loop`` (a ``TimeLoc``) over
    ``n`` periods: by DATE, as the operators read it — each period the
    line's value of that date, the periods it lacks continued by its own
    ``extend=`` rule (:func:`_onto_window`). None where it is read as it
    is: on the loop's window, or on none (a constant, a value typed in, a
    line along keys). A line without a rule, an unnamed formula that
    does not cover the window, another grain and a formula whose periods
    have no one date are refused as the operators refuse them."""
    place = _place_of(var, x)
    if place is None or place == _NO_WINDOW_YET:
        return None
    if place == _NO_ONE_WINDOW:
        raise ValueError(_no_one_window(var, n, loop.start))
    loc, k, own = place
    if _same_window((loc, k), (loop, n)):
        return None
    grain = loop.grain
    if loc.grain != grain:
        raise ValueError(_mixed_grains(var, loc.grain, row, grain))
    if loc.first_day != loop.first_day:
        # A short first period may open the loop's window, never sit
        # inside it (as in :func:`_meet`).
        st, ot = (loc, loop) if loc.first_day is not None else (loop, loc)
        if ot.first_day is not None or _starts_on(st.start, grain) >= _starts_on(ot.start, grain):
            raise ValueError(_mixed_first_periods(row, loop, var, loc))
    from .time import _count, _parse
    if _starts_on(loc.start, grain) >= _starts_on(loop.start, grain):
        off = _count(_parse(loop.start, grain), _parse(loc.start, grain), grain) - 1
    else:
        off = 1 - _count(_parse(loc.start, grain), _parse(loop.start, grain), grain)
    other = f"{n} period(s) from {loop.start}, the loop's window"
    if own:
        return _onto_window(var, x, off, n, loop.start, grain, loop.first_day, other)
    if off <= 0 and n - off <= k:
        return list(x[-off:n - off])
    raise ValueError(_off_its_window(var, loc, k, own, other))


def _line_name(v: Any) -> str:
    """How a message names a line: its display or python name, else its
    formula (``mo.time.days`` as typed), else "an unnamed line"."""
    label = getattr(v, '_display_name', None) or getattr(v, '_python_name', None)
    if label:
        return repr(label)
    formula = getattr(v, 'formula', None)
    if formula:
        return f"the line {'mo.' + formula if formula.startswith('time.') else formula}"
    return "an unnamed line"


def _line_code(v: Any) -> Optional[str]:
    """The line as code a message can suggest typing: its python name,
    else its formula in brackets; ``None`` when neither exists."""
    name = getattr(v, '_python_name', None)
    if name:
        return name
    formula = getattr(v, 'formula', None)
    return f"({formula})" if formula else None


def _period_count(series: Any) -> Optional[int]:
    """How many periods ``series`` runs along, or ``None`` when it is not a
    period series: a list of several values, a list of any length that
    is located in time (``start=``/``grain=``, ``extend=``, a schedule —
    a one-value located flow still meets every period), or a tracked
    line over periods."""
    values = getattr(series, '_value', None)
    if isinstance(values, TrackValues):
        n = values.time_length
        return n if n is not None and n > 1 else None
    if not isinstance(values, list):
        return None
    located = (
        getattr(series, '_grain', None) is not None
        or getattr(series, '_extend', None) is not None
        or getattr(series, '_schedule', None) is not None
    )
    if len(values) > 1 or (located and len(values) >= 1):
        return len(values)
    return None


def _keys_against_periods(keyed: Any, labels: List[Any], series: Any, n: int) -> str:
    """A line that runs along keys met a period series in one formula:
    name both, show the pairing a cell-by-cell zip would make, give the
    fixes that run (one key, the switch, or both lines along the same
    keys / axis)."""
    shown = ", ".join(str(k) for k in labels)
    pairs = f"{labels[0]} with the first period"
    if len(labels) > 1:
        pairs += f", {labels[1]} with the second"
    if len(labels) > 2:
        pairs += ", …"
    code = _line_code(keyed)
    first = labels[0]
    if code is None:
        pick = f"pick one key of {_line_name(keyed)}"
    elif isinstance(first, str) and getattr(keyed, '_keys', None):
        pick = f'pick one key, {code}["{first}"]'
    else:
        pick = f"pick one cell, {code}[0]"
    # The series can be given the keys only when it is a bare input a
    # person typed; a computed line, a time property or an unnamed line
    # gets the switch instead, which reads one key by a number.
    typeable = (
        getattr(series, '_python_name', None) and getattr(series, '_expr', None) is None
    )
    if getattr(keyed, '_keys', None):
        along = f"its keys ({shown})"
        literal = ", ".join(f"{k!r}: …" for k in labels)
        second = (f"give {_line_name(series)} the same keys, mo.Variable({{{literal}}})"
                  if typeable else
                  f"switch with mo.CHOOSE(pick, {code or 'the line'}), one number picking a key")
    else:
        axis = keyed._indexed_by[0]
        axis_code = getattr(axis, '_python_name', None)
        axis_label = axis_code or getattr(axis, '_display_name', None)
        along = f"the axis {axis_label!r} ({shown})" if axis_label else f"an axis ({shown})"
        second = (f"lay {_line_name(series)} along the same axis, "
                  f"indexed_by=[{axis_code or '…'}]"
                  if typeable else
                  f"switch with mo.CHOOSE(pick, {code or 'the line'}), one number picking a cell")
    values_word = "1 value" if n == 1 else f"{n} values"
    return (f"{_line_name(keyed)} runs along {along} and {_line_name(series)} runs along "
            f"periods ({values_word}): paired cell by cell ({pairs}) the numbers would "
            f"look right and be wrong. {pick[0].upper() + pick[1:]}; or {second}.")


def _keys_against_keys(left: Any, lk: List[Any], right: Any, rk: List[Any]) -> str:
    """A keyed list met a line on an axis whose labels differ: nobody
    aligns the pair (keys align with keys, axes with axes), so a zip by
    position would pair unrelated labels."""
    def along(v: Any, labels: List[Any]) -> str:
        shown = ", ".join(str(k) for k in labels)
        if getattr(v, '_keys', None):
            return f"its keys ({shown})"
        axis = v._indexed_by[0]
        name = getattr(axis, '_python_name', None) or getattr(axis, '_display_name', None)
        return f"the axis {name!r} ({shown})" if name else f"an axis ({shown})"
    return (f"{_line_name(left)} runs along {along(left, lk)} and {_line_name(right)} runs "
            f"along {along(right, rk)}: paired cell by cell the labels would not match. "
            f"Lay both lines along the same axis (indexed_by=[…]), or give both the same "
            f"keys in the same order.")


#: Unit words read as a year or a month, for the per-period hint.
_YEAR_WORDS = frozenset({"year", "years", "yr", "yrs", "y", "год", "года", "лет"})
_MONTH_WORDS = frozenset({"month", "months", "mo", "mon", "мес", "месяц", "месяца"})


def _month_and_year(left: Any, right: Any, left_over: dict, right_over: dict) -> Optional[str]:
    """The hint for two units that differ only in a month word against a
    year word at the same power (months / yrs, USD/month / USD/yr): a
    label never rescales, so the step between them is written out."""
    if len(left_over) != 1 or len(right_over) != 1:
        return None
    (lw, lexp), = left_over.items()
    (rw, rexp), = right_over.items()
    if lexp != rexp or lexp not in (1, -1):
        return None
    if lw.lower() in _MONTH_WORDS and rw.lower() in _YEAR_WORDS:
        months, years = str(left), str(right)
    elif lw.lower() in _YEAR_WORDS and rw.lower() in _MONTH_WORDS:
        months, years = str(right), str(left)
    else:
        return None
    to_years, to_months = ("/ 12", "* 12") if lexp == 1 else ("* 12", "/ 12")
    return (
        f" A unit is a label and never rescales: a line in {months} does not "
        f"become {years} when its numbers are divided or multiplied by 12. "
        f"Write the step: (x {to_years}).set_unit('{years}') for {months} "
        f"into {years}, or (x {to_months}).set_unit('{months}') for {years} "
        f"into {months}."
    )


def _incompatible_units(left: Any, right: Any) -> str:
    """Why two units do not add — and, when one side is the other times a
    word or two, that those words are dimensions: a string unit always
    is, and a word that labels a pure number is declared as one."""
    lc, rc = left._unit_components, right._unit_components
    if str(left) == str(right) and bool(lc) != bool(rc):
        def what(u: Any) -> str:
            return "a dimension" if u._unit_components else "a labelled pure number"
        pair = f"{left} ({what(left)}) and {right} ({what(right)})"
    else:
        pair = f"{left} and {right}"
    message = f"Cannot add/subtract values with incompatible units: {pair}."
    # One side is the other times words, each to the first power: what is
    # left of each side once the other's symbols (at the same exponent) are
    # taken away is empty on one side and plain words on the other.
    left_over = {k: v for k, v in lc.items() if rc.get(k) != v}
    right_over = {k: v for k, v in rc.items() if lc.get(k) != v}
    smaller, larger = ((lc, right_over) if not left_over else
                       (rc, left_over) if not right_over else (None, None))
    convert = ("convert one side to the other's unit before the operation, "
               "or drop the unit on one side if the arithmetic is intentional.")
    hinted = True
    if (smaller and larger and all(v == 1 for v in larger.values())
            and not set(larger) & set(smaller)):
        words = sorted(larger)
        named = " and ".join(f"'{w}'" for w in words)
        verb = "is a dimension" if len(words) == 1 else "are dimensions"
        example = words[0]
        message += (
            f" {named} {verb} here: a string unit always is. If it labels a "
            f"pure number (a factor, a share, a flag), declare it once - "
            f"{example.upper() if example.isidentifier() else 'NAME'} = "
            f"mo.Unit.dimensionless('{example}') - and set that object as the "
            f"unit of the lines that carry it."
        )
    elif (smaller and larger and all(v == -1 for v in larger.values())
            and not set(larger) & set(smaller)):
        words = sorted(larger)
        named = " and ".join(f"'{w}'" for w in words)
        message += (f" One side is per {named}: it becomes an amount once "
                    f"multiplied by a quantity in {named}")
        word = words[0]
        if len(words) == 1 and word.lower() in _YEAR_WORDS:
            message += (
                f" - the share of the year a period covers: "
                f"year_share = (mo.time.months / 12).set_unit('{word}') for a "
                f"fee or a salary, (mo.time.days / 365) for interest by the "
                f"day - then rate * year_share."
            )
        elif len(words) == 1 and word.lower() in _MONTH_WORDS:
            message += (
                f" - the months a period covers: "
                f"mo.time.months.set_unit('{word}'), then rate * that."
            )
        else:
            message += "."
    elif (time_hint := _month_and_year(left, right, left_over, right_over)):
        message += time_hint
        convert = "drop the unit on one side if the arithmetic is intentional."
    else:
        hinted = False
    # 'Otherwise' only follows a hint; alone, the advice opens the sentence.
    return message + (f" Otherwise {convert}" if hinted
                      else f" {convert[0].upper()}{convert[1:]}")


#: Words that label a count of a small step - basis points, percentage
#: points. A label never rescales: 50 bps is written 0.005, 2 pp is 0.02.
_STEP_WORDS = {'bps': ('50 bps', '0.005', '10_000'), 'pp': ('2 pp', '0.02', '100')}


def _numbers(value: Any) -> list:
    """Every number in a value - a scalar, a list, each track's cells."""
    if isinstance(value, TrackValues):
        return [n for v in value._tracks.values() for n in _numbers(v)]
    if isinstance(value, (list, tuple)):
        return [n for v in value for n in _numbers(v)]
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return []
    return [value]


def _own_numbers(var: Any) -> list:
    return _numbers(var._value)


def _holds_a_count(var: Any, numbers: Any = None) -> bool:
    """A bps or pp line holding a count (50 where 0.005 is meant).
    ``numbers`` reads the numbers a line is judged by (all of its own by
    default)."""
    unit = var._get_unit()
    if unit is None or unit._unit_components or unit._display_symbol not in _STEP_WORDS:
        return False
    return any(abs(n) >= 1 for n in (numbers or _own_numbers)(var))


def _holds_a_fraction(var: Any, numbers: Any = None) -> bool:
    """A % line holding fractions only (0.05, never 5)."""
    unit = var._get_unit()
    if unit is None or unit._unit_components or unit._display_symbol != '%':
        return False
    shares = (numbers or _own_numbers)(var)
    return bool(shares) and not any(abs(n) >= 1 for n in shares)


def _refuse_a_count_added_to_a_share(left: Any, right: Any, numbers: Any = None,
                                     where: Optional[str] = None) -> None:
    """A bps or pp line holding a count (50) meets a % line holding
    fractions (0.05) in a sum: a label never rescales, so the result is
    wrong under any reading - 0.05 + 50 is 50.05, not 0.055. ``numbers``
    reads the numbers a side is judged by (all of its own by default);
    ``where`` names the function that adds, compares or picks the two
    (``'mo.MAX'``) in place of ``+``."""
    for counted, share in ((left, right), (right, left)):
        if not (_holds_a_count(counted, numbers) and _holds_a_fraction(share, numbers)):
            continue
        word = counted._get_unit()._display_symbol
        example, fraction, step = _STEP_WORDS[word]
        code = getattr(counted, '_python_name', None) or _line_name(counted)
        met = (f"added to {_line_name(share)} (%)" if where is None
               else f"beside {_line_name(share)} (%) in {where}")
        wrong = "the sum would be wrong" if where is None else "the two cannot be read as one"
        raise ValueError(
            f"{_line_name(counted)} is in {word} and holds a count, {met}, which "
            f"holds a fraction: a label never rescales, so {wrong}. A {word} line "
            f"holds the fraction ({example} is {fraction}), or the count is "
            f"converted first: {code} / {step}."
        )


def _refuse_counts_among_shares(lines: Sequence[Any], numbers: Any = None,
                                where: Optional[str] = None) -> None:
    """:func:`_refuse_a_count_added_to_a_share` over every two of the
    ``lines`` a function adds, compares or picks, in one pass: a line is a
    count, a fraction or neither on its own, so the pair to refuse is the
    first count and the first fraction among them."""
    counted = next((v for v in lines if _holds_a_count(v, numbers)), None)
    if counted is None:
        return
    share = next((v for v in lines if _holds_a_fraction(v, numbers)), None)
    if share is not None:
        _refuse_a_count_added_to_a_share(counted, share, numbers, where)


class _VariableArithmetic:
    """Operator-overloading mixin for :class:`Variable`.

    The attributes/methods declared below are provided by the concrete
    :class:`Variable` subclass — declared here under ``TYPE_CHECKING``
    so type-checkers know the mixin expects them on ``self``. They have
    no runtime effect.
    """

    # Attributes provided by the Variable subclass — type-only annotations.
    _value: Any
    _expr: Optional[Expr]
    _keys: Optional[List[Any]]
    _unit: Optional["Unit"]
    var_type: str
    value_type: str

    if TYPE_CHECKING:
        def _as_operand(self, parenthesize: bool = False) -> Expr: ...
        def _set_expr(self, expr: Expr) -> None: ...

    # ─── Broadcasting ───────────────────────────────────────────
    #
    # COMPUTE CHOKEPOINT. Every arithmetic value produced by the engine
    # flows through this function (operator overloads in
    # ``_VariableArithmetic`` call it). Do not add parallel compute
    # paths — keep all value arithmetic routed through here.

    @staticmethod
    def _reject_pending_roles(*operands) -> None:
        """Refuse eager compute on a role-kwarg Variable that has not
        materialized. Values compute eagerly, at operator time;
        an unmaterialized operand is ``None`` and would bake a permanent
        ``#VALUE!`` into every dependent — loud beats silently wrong.
        """
        for v in operands:
            if (getattr(v, '_role_kwargs', None) is not None
                    and not getattr(v, '_roles_materialized', False)):
                label = (getattr(v, '_display_name', None)
                         or getattr(v, '_python_name', None) or 'variable')
                raise ValueError(
                    f"{label!r} authors track coordinates that have not "
                    f"materialized yet — attach it to the model (values "
                    f"compute at operator time, so the attach must come "
                    f"BEFORE the formulas that read it)."
                )

    @staticmethod
    def _reject_keys_against_periods(left: Any, right: Any) -> None:
        """Refuse a line that runs along keys against a period series.

        Zipped cell by cell, a keyed list (or a line on one finite
        non-time axis) and a series — a list, a located list, a tracked
        line — pair Bear with January: numbers that look right and are
        wrong. Keyed against keyed is alignment's business
        (``_align_keyed_values``), an axis against an axis the axis
        expansion's; a keyed list against an axis of other labels is
        nobody's and is refused too. A scalar, a one-value bare list and
        a raw Python list (``keyed * other.value``, positional on
        purpose) broadcast as before.
        """
        lk, rk = keys_along(left), keys_along(right)
        if lk is not None and rk is not None:
            left_dict = bool(getattr(left, '_keys', None))
            right_dict = bool(getattr(right, '_keys', None))
            if left_dict != right_dict and [str(k) for k in lk] != [str(k) for k in rk]:
                raise ValueError(_keys_against_keys(left, lk, right, rk))
            return
        if lk is None and rk is None:
            return
        if lk is not None:
            keyed, labels, series = left, lk, right
        else:
            keyed, labels, series = right, rk or [], left
        n = _period_count(series)
        if n is None:
            return
        if (getattr(series, '_indexed_by', None) or ()) and (
            getattr(keyed, '_indexed_by', None) or ()
        ):
            return  # both on axes — the axis expansion crosses them
        raise ValueError(_keys_against_periods(keyed, labels, series, n))

    @staticmethod
    def _broadcast_operation(left, right, op, safe_divide: bool = False,
                             op_symbol: Optional[str] = None):
        """Element-wise op with NumPy-style broadcasting.

        Scalar/scalar → scalar. List/scalar → element-wise list.
        Scalar/list → element-wise list. List/list → element-wise;
        length-1 operands broadcast over the longer list.

        Excel-style error markers (``'#DIV/0!'``, ``'#NAME?'``, …)
        propagate through operations. With ``safe_divide=True``, division
        by zero yields ``'#DIV/0!'`` instead of raising.

        ``op_symbol`` names the operator (``'+'``, ``'<'``, …). Dates need
        it: ``date + 5`` is a date, ``date * 5`` is a number and
        ``date < 5`` is a bool, all computed on the date's serial number.
        """
        # Whether a blank reads as "" (a comparison against text) is one
        # question per call, not one per cell.
        blank_reads_text = op_symbol in _COMPARISONS

        def safe_op(a, b):
            ta, tb = type(a), type(b)
            if (ta is float or ta is int) and (tb is float or tb is int):
                # Two plain numbers — most cells of a model. None of the
                # markers below can be one of them (not text, not blank,
                # not a date), so only the zero divisor is left to ask.
                if safe_divide and b == 0:
                    return '#DIV/0!'
                try:
                    return op(a, b)
                except Exception:
                    return '#VALUE!'
            # An Excel-style error marker propagates, the left one first.
            if isinstance(a, str) and a.startswith('#'):
                return a
            if isinstance(b, str) and b.startswith('#'):
                return b
            # A blank counts as zero, as in Excel: a formula over a
            # month nobody entered computes over 0 (blank + 5 = 5,
            # 5 / blank = #DIV/0!). Only an input cell stays blank;
            # whether an actual was entered is the blend's question,
            # asked of its entries (``core.entries``).
            if a is None:
                a = '' if isinstance(b, str) and blank_reads_text else 0
            if b is None:
                b = '' if isinstance(a, str) and blank_reads_text else 0
            if safe_divide and b == 0:
                return '#DIV/0!'
            if isinstance(a, (date, timedelta)) or isinstance(b, (date, timedelta)):
                return _date_serial_op(a, b, op, op_symbol)
            try:
                return op(a, b)
            except Exception:
                return '#VALUE!'

        left_is_list = isinstance(left, list)
        right_is_list = isinstance(right, list)

        if not left_is_list and not right_is_list:
            return safe_op(left, right)
        elif left_is_list and not right_is_list:
            return [safe_op(x, right) for x in left]
        elif not left_is_list and right_is_list:
            return [safe_op(left, x) for x in right]
        else:
            if len(left) != len(right):
                if len(left) == 1:
                    return [safe_op(left[0], x) for x in right]
                elif len(right) == 1:
                    return [safe_op(x, right[0]) for x in left]
                raise ValueError(
                    f"Cannot broadcast lists of different lengths: "
                    f"{len(left)} vs {len(right)}. "
                    f"Element-wise ops need equal lengths, or one side of "
                    f"length 1 which broadcasts over the other. "
                    f"Common fix: align the two list-valued inputs to the "
                    f"same period count."
                )
            return [safe_op(a, b) for a, b in zip(left, right)]

    @staticmethod
    def _align_keyed_values(left, right):
        """Reorder right's list to match left's key order when both are keyed.

        If only one side is keyed, or keys match already, pass through.
        Raises if key sets disagree — silently reordering a mismatch would
        produce a silently wrong result.
        """
        variable_cls = _Variable or _variable_class()
        left_val = left._value if isinstance(left, variable_cls) else left
        right_val = right._value if isinstance(right, variable_cls) else right

        left_keys = getattr(left, '_keys', None) if isinstance(left, variable_cls) else None
        right_keys = getattr(right, '_keys', None) if isinstance(right, variable_cls) else None

        if (left_keys is not None and right_keys is not None
                and isinstance(right_val, list)):
            if set(left_keys) != set(right_keys):
                raise ValueError(
                    f"Cannot align Variables with different key sets: "
                    f"{left_keys} vs {right_keys}. "
                    f"Element-wise ops on keyed Variables need matching keys "
                    f"so each value has a partner. Either pick the same keys "
                    f"for both sides, or drop keys on one side (pass through "
                    f"``.value``) and use positional broadcasting instead."
                )
            if left_keys != right_keys:
                key_to_idx = {k: i for i, k in enumerate(right_keys)}
                right_val = [right_val[key_to_idx[k]] for k in left_keys]

        return left_val, right_val

    @staticmethod
    def _propagate_keys(left, right):
        """Result keys come from whichever operand has them; left wins if both."""
        variable_cls = _Variable or _variable_class()
        left_keys = getattr(left, '_keys', None) if isinstance(left, variable_cls) else None
        right_keys = getattr(right, '_keys', None) if isinstance(right, variable_cls) else None
        if left_keys is not None:
            return list(left_keys)
        if right_keys is not None:
            return list(right_keys)
        return None

    # ─── Units ──────────────────────────────────────────────────
    #
    # A Unit is a value: its algebra always builds a new one and nothing
    # changes a unit once built. So when the result of ``*`` / ``/`` IS
    # one operand's unit — the other side is unitless or dimensionless
    # (``$ * %``, ``$ / 12``) — that operand's own unit is the answer,
    # exactly as ``+`` / ``-`` already return theirs. Returning it avoids
    # building an equal copy (and a dimensionless Unit for the bare side)
    # on every operation; the algebra itself is unchanged.

    @staticmethod
    def _unit_mul(left_unit, right_unit):
        """Resulting unit for multiplication."""
        if left_unit is None and right_unit is None:
            return None
        left_c = left_unit._unit_components if left_unit is not None else None
        right_c = right_unit._unit_components if right_unit is not None else None
        if not right_c:
            return left_unit if left_c else None
        if not left_c:
            return right_unit
        result = left_unit * right_unit
        return result if result._unit_components else None

    @staticmethod
    def _unit_div(left_unit, right_unit):
        """Resulting unit for division."""
        if left_unit is None and right_unit is None:
            return None
        left_c = left_unit._unit_components if left_unit is not None else None
        right_c = right_unit._unit_components if right_unit is not None else None
        if not right_c:
            return left_unit if left_c else None
        from .unit import Unit
        lu = left_unit if left_unit is not None else Unit.dimensionless()
        result = lu / right_unit
        return result if result._unit_components else None

    @staticmethod
    def _unit_mod(left_unit, right_unit):
        """Resulting unit for ``%`` - ``MOD``'s law: the remainder keeps
        the unit both sides share, a number goes with any, two units are
        refused."""
        try:
            return _VariableArithmetic._unit_add(left_unit, right_unit)
        except ValueError as exc:
            raise ValueError(str(exc).replace(
                "Cannot add/subtract", "Cannot take the remainder of", 1)) from None

    @staticmethod
    def _unit_add(left_unit, right_unit):
        """Resulting unit for addition/subtraction.

        Both-None → None. One-sided (scalar + unitful) inherits the
        unit. Both-sides-present must agree — otherwise raise, because
        silently returning one side would mask a real modeling bug.
        """
        if left_unit is None and right_unit is None:
            return None
        if left_unit is not None and right_unit is not None:
            if left_unit != right_unit:
                raise ValueError(_incompatible_units(left_unit, right_unit))
            return left_unit
        return left_unit if left_unit is not None else right_unit

    @staticmethod
    def _unit_pow(unit, exponent):
        """Resulting unit for ``x ** n``: each dimension's power times
        ``n``, when every one comes out whole — ``hr ** 2`` is ``hr^2``,
        ``hr ** 2.0`` too, ``(USD^2) ** 0.5`` is ``USD``. Otherwise None:
        a fractional power of a dimension (``USD ** 0.5``) has no unit to
        write, and an exponent that is not a number (a line, a list of
        them) may differ by period.
        """
        if (unit is None or isinstance(exponent, bool)
                or not isinstance(exponent, (int, float)) or not math.isfinite(exponent)):
            return None
        powers = {}
        for symbol, power in unit._unit_components.items():
            raised = power * exponent
            if abs(raised - round(raised)) > 1e-9:
                return None
            powers[symbol] = round(raised)
        from .unit import Unit
        result = Unit(_components=powers)
        # ``hr ** 0`` → dimensionless; collapse the empty-components
        # Unit to ``None`` to match the ``mul`` / ``div`` convention.
        return result if result._unit_components else None

    def _get_unit(self):
        """Get unit from self, or None for scalars."""
        return getattr(self, '_unit', None)

    # ─── Binary operators ──────────────────────────────────────

    def _binary_op(self, other, op_symbol: str, op_func, unit_fn,
                   safe_divide: bool = False, parenthesize: bool = False,
                   reversed: bool = False) -> "Variable":
        """Shared implementation for all binary arithmetic operators.

        Args:
            other: Right operand — another Variable, a scalar, or a list.
            op_symbol: Text inserted between operands in the formula string.
            op_func: Element-wise function applied for value computation.
            unit_fn: ``(left_unit, right_unit) -> result_unit`` or ``None``.
                     ``None`` for ``**``, whose exponent sets the unit
                     afterwards.
            safe_divide: Route through the safe-division path so 1/0 becomes
                         ``"#DIV/0!"`` instead of raising.
            parenthesize: Wrap compound-unnamed-formula operands in parens
                          to preserve precedence (``(a+b) / (c+d)``).
            reversed: ``True`` for reflected ops (``__rsub__``, etc.).
                      In the formula string, operand order flips; dependency
                      refs still record ``[self, other]`` in that order.
        """
        variable_cls = _Variable or _variable_class()

        self._reject_pending_roles(self, other)

        # ``result`` is a plain Variable built right here, and its class
        # declares no descriptor for any field written below: its own
        # ``__dict__`` takes each write exactly as ``__setattr__`` would,
        # without a Python-level call per field.
        if isinstance(other, variable_cls):
            left, right = (other, self) if reversed else (self, other)
            expr_node = BinOp(op_symbol, left._as_operand(parenthesize),
                              right._as_operand(parenthesize))
            result = variable_cls()
            result._set_expr(expr_node)
            fields = result.__dict__
            if op_symbol in ('+', '-') and not any(
                    getattr(v, '_awaits', None) or getattr(v, '_forward_of', None) is not None
                    for v in (left, right)):
                _refuse_a_count_added_to_a_share(left, right)
            self._reject_keys_against_periods(left, right)
            lv, rv = self._align_keyed_values(left, right)
            expanded = self._expand_axised_operands(left, right, lv, rv)
            if expanded is not None:
                lv, rv = expanded
            (lv, rv), window, runs_on, through = _meet((left, right), [lv, rv])
            if window is not None:
                # The result is located at the union window explicitly —
                # downstream alignment then knows where it starts — and
                # marked as read by date there (a renderer reads its
                # operands' cells of the same date).
                fields['_start'], fields['_grain'], fields['_first_day'] = window
                fields['_read_by_date'] = True
            if isinstance(lv, TrackValues) or isinstance(rv, TrackValues):
                # The tracks broadcast law: zip per role;
                # a plain operand broadcasts into every track;
                # differing role sets die loudly inside ``combine``.
                fields['_value'] = TrackValues.combine(
                    lv, rv,
                    lambda a, b: self._broadcast_operation(
                        a, b, op_func, safe_divide=safe_divide, op_symbol=op_symbol
                    ),
                )
            else:
                fields['_value'] = self._broadcast_operation(
                    lv, rv, op_func, safe_divide=safe_divide, op_symbol=op_symbol
                )
            if unit_fn is not None:
                fields['_unit'] = _provisional_unit(
                    unit_fn, left._get_unit(), right._get_unit(), left, right)
            if window is None:
                _keep_runs_on(result, runs_on, through)
        else:
            self_u = self._get_unit()

            def _combine(a, b):
                if isinstance(a, TrackValues) or isinstance(b, TrackValues):
                    return TrackValues.combine(
                        a, b,
                        lambda x, y: self._broadcast_operation(
                            x, y, op_func, safe_divide=safe_divide, op_symbol=op_symbol
                        ),
                    )
                return self._broadcast_operation(
                    a, b, op_func, safe_divide=safe_divide, op_symbol=op_symbol
                )

            if reversed:
                expr_node = BinOp(op_symbol, Literal(other), self._as_operand(parenthesize))
                val = _combine(other, self._value)
            else:
                expr_node = BinOp(op_symbol, self._as_operand(parenthesize), Literal(other))
                val = _combine(self._value, other)
            result = variable_cls()
            result._set_expr(expr_node)
            fields = result.__dict__
            fields['_value'] = val
            if unit_fn is not None:
                # Scalars are treated as dimensionless when computing the output unit
                fields['_unit'] = _provisional_unit(
                    unit_fn, *((None, self_u) if reversed else (self_u, None)), self)

        value = fields['_value']
        if isinstance(value, list) and len(value) > 1:
            fields['var_type'] = "list"
        elif type(value).__name__ == 'TrackValues' and value.time_length:
            fields['var_type'] = "list"
        # A date-valued result (e.g. ``start + 90``, ``end - 1``) must carry
        # the datetime value type so the writer formats the cell as a date.
        sample = value[0] if isinstance(value, list) and value else value
        if isinstance(sample, (date, datetime)):
            fields['value_type'] = "datetime"
        k_left, k_right = (other, self) if reversed else (self, other)
        fields['_keys'] = self._propagate_keys(k_left, k_right)
        fields['_indexed_by'] = self._propagate_indexed_by(k_left, k_right)
        return result

    @staticmethod
    def _align_by_date(left: Any, right: Any, lv: Any, rv: Any):
        """Date-aligned zip — calendar alignment of two operands on other
        windows (:func:`_on_dates`, which every function that is one call
        per period shares).

        Engages ONLY when both operands are series on windows that
        disagree (different starts or different lengths at one grain): a
        line's own window, or for an unnamed formula the one of the lines
        it reads. Same-window operands keep the byte-identical positional
        path. Different GRAINS refuse — that is re-grain territory, not
        alignment.

        Each side materializes onto the union window: cells outside an
        operand's own extent follow its DECLARED ``extend=`` rule — zero
        (a flow is absent), hold (a rate stays in force; held backward
        from the first value when extended into the past), or
        none/undeclared → an error naming both extents and the rules to declare. A plain
        line beside a tracked one is read on the tracked line's window.
        Returns ``(lv', rv', union_start, grain, union_first)`` or
        ``None`` when not applicable.
        """
        out, window = _on_dates((left, right), [lv, rv])
        if window is None:
            return None
        return out[0], out[1], window.start, window.grain, window.first_day

    @staticmethod
    def _expand_axised_operands(left: Any, right: Any, lv: Any, rv: Any):
        """Identity-aligned expansion of two AXISED operands.

        Engages only when BOTH operands declare ``_indexed_by`` and the
        axis tuples differ (different sets → cross-product; same set in
        a different order → transpose). Each flat row-major value is
        expanded to the union shape (left-first axis order — the same
        order ``_propagate_indexed_by`` emits) by replication, so the
        subsequent element-wise zip computes the true cross-product
        instead of the historical 2-cell diagonal that used to hide
        under a 4-cell label.

        Returns ``(lv', rv')`` or ``None`` when not applicable (either
        side lacks declared axes — scalars broadcast fine and plain
        time-lists keep their positional semantics — or
        the axis tuples are identical, where the plain zip is already
        correct and cheaper).
        """
        la = getattr(left, '_indexed_by', ()) or ()
        ra = getattr(right, '_indexed_by', ()) or ()
        if not la or not ra:
            return None
        if tuple(id(a) for a in la) == tuple(id(a) for a in ra):
            return None
        if lv is None or rv is None:
            return None

        union = _VariableArithmetic._propagate_indexed_by(left, right)
        dims = [
            len(a._value) if isinstance(getattr(a, '_value', None), list) else 1
            for a in union
        ]
        total = 1
        for d in dims:
            total *= d
        pos_by_id = {id(a): i for i, a in enumerate(union)}

        def expand(axes: Tuple[Any, ...], flat: Any) -> list:
            vals = flat if isinstance(flat, list) else [flat]
            pos = [pos_by_id[id(a)] for a in axes]
            own_dims = [dims[p] for p in pos]
            out = []
            for cell in range(total):
                # Union coordinates, row-major (last axis fastest).
                rem, coords = cell, [0] * len(dims)
                for i in range(len(dims) - 1, -1, -1):
                    coords[i] = rem % dims[i]
                    rem //= dims[i]
                # Row-major index into the operand's own axis order.
                idx = 0
                for p, d in zip(pos, own_dims):
                    idx = idx * d + coords[p]
                out.append(vals[idx])
            return out

        return expand(la, lv), expand(ra, rv)

    @staticmethod
    def _propagate_indexed_by(left: Any, right: Any) -> Tuple[Any, ...]:
        """Union the operands' ``_indexed_by`` axes, preserving the
        order of first appearance (left wins ties).

        Operator-built Variables inherit shape from this union — e.g.
        ``revenue * cogs_pct`` of shapes ``(months, scenarios)`` and
        ``()`` produces a result of shape ``(months, scenarios)``;
        ``a * b`` over two distinct axes produces the cross-product
        shape ``(left_axes..., right_axes...)``.

        Compares axes by identity (``id``) — Variable's ``__eq__`` is
        overridden to return a comparison expression, so ``axis in
        seen`` would always go through that operator and produce
        nonsense. Identity is the right relation for "same axis Variable".
        """
        la = getattr(left, '_indexed_by', ()) if left is not None else ()
        ra = getattr(right, '_indexed_by', ()) if right is not None else ()
        seen: list = list(la)
        seen_ids = {id(a) for a in seen}
        for axis in ra:
            if id(axis) not in seen_ids:
                seen.append(axis)
                seen_ids.add(id(axis))
        return tuple(seen)

    # ─── Binary dunders ────────────────────────────────────────

    def __add__(self, other) -> "Variable":
        return self._binary_op(other, "+", lambda a, b: a + b, self._unit_add)

    def __radd__(self, other) -> "Variable":
        return self.__add__(other)

    def __sub__(self, other) -> "Variable":
        return self._binary_op(other, "-", lambda a, b: a - b, self._unit_add)

    def __rsub__(self, other) -> "Variable":
        return self._binary_op(other, "-", lambda a, b: a - b, self._unit_add, reversed=True)

    def __mul__(self, other) -> "Variable":
        return self._binary_op(other, "*", lambda a, b: a * b, self._unit_mul)

    def __rmul__(self, other) -> "Variable":
        return self.__mul__(other)

    def __truediv__(self, other) -> "Variable":
        return self._binary_op(other, "/", lambda a, b: a / b, self._unit_div,
                               safe_divide=True, parenthesize=True)

    def __rtruediv__(self, other) -> "Variable":
        return self._binary_op(other, "/", lambda a, b: a / b, self._unit_div,
                               safe_divide=True, parenthesize=True, reversed=True)

    def __pow__(self, other) -> "Variable":
        result = self._binary_op(other, "**", lambda a, b: a ** b, None)
        # A constant exponent raises the unit (``_unit_pow``); one that
        # is a line leaves none.
        variable_cls = _Variable or _variable_class()
        if not isinstance(other, variable_cls):
            result._unit = self._unit_pow(self._get_unit(), other)
        return result

    def __rpow__(self, other) -> "Variable":
        return self._binary_op(other, "**", lambda a, b: a ** b, None, reversed=True)

    def __mod__(self, other) -> "Variable":
        return self._binary_op(other, "%", lambda a, b: a % b, self._unit_mod)

    def __rmod__(self, other) -> "Variable":
        return self._binary_op(other, "%", lambda a, b: a % b, self._unit_mod,
                               reversed=True)

    def __floordiv__(self, other) -> "Variable":
        return self._binary_op(other, "//", lambda a, b: a // b, self._unit_div)

    def __rfloordiv__(self, other) -> "Variable":
        return self._binary_op(other, "//", lambda a, b: a // b, self._unit_div,
                               reversed=True)

    # ─── Unary dunders ─────────────────────────────────────────

    def __neg__(self) -> "Variable":
        """-a (unary negation)"""
        variable_cls = _Variable or _variable_class()
        result = variable_cls()
        result._set_expr(UnaryOp("-", self._as_operand(parenthesize=True)))

        def neg(x):
            if x is None:
                return 0                # -blank is 0, as in Excel
            if isinstance(x, str) and x.startswith('#'):
                return x
            try:
                return -x
            except TypeError:
                return '#VALUE!'

        def neg_series(v):
            return [neg(x) for x in v] if isinstance(v, list) else neg(v)

        if isinstance(self._value, TrackValues):
            result._value = TrackValues.lift(neg_series, self._value)
        else:
            result._value = neg_series(self._value)
        result._unit = self._get_unit()
        result._keys = list(self._keys) if self._keys is not None else None
        result._indexed_by = getattr(self, '_indexed_by', ())
        if isinstance(result._value, list) and len(result._value) > 1:
            result.var_type = 'list'
        elif type(result._value).__name__ == 'TrackValues' and result._value.time_length:
            result.var_type = 'list'
        return result

    def __pos__(self) -> "Variable":
        """+a is a no-op — return self."""
        return self  # type: ignore[return-value]  # mixin self is the concrete Variable at runtime

    # ─── Comparisons ───────────────────────────────────────────
    #
    # COMPUTE CHOKEPOINT for comparisons. Parallel to
    # ``_broadcast_operation`` for arithmetic — do not add parallel
    # compute paths.

    def _comparison_op(self, other, op_symbol: str, op_func) -> "Variable":
        """Shared implementation for all comparison operators."""
        variable_cls = _Variable or _variable_class()

        # ``self`` is always a :class:`Variable` at runtime — the mixin is
        # only mounted on Variable. The casts below tell mypy that.
        self_var: "Variable" = self  # type: ignore[assignment]
        self._reject_pending_roles(self, other)

        def _combine(a, b):
            if isinstance(a, TrackValues) or isinstance(b, TrackValues):
                return TrackValues.combine(
                    a, b, lambda x, y: self._broadcast_operation(
                        x, y, op_func, op_symbol=op_symbol
                    )
                )
            return self._broadcast_operation(a, b, op_func, op_symbol=op_symbol)

        # A plain Variable built right here: its fields go straight into
        # its ``__dict__``, as in ``_binary_op``.
        if isinstance(other, variable_cls):
            # Over a row whose loop is open the units are provisional: the
            # loop checks the comparison when it closes.
            lu, ru = self_var._get_unit(), other._get_unit()
            if (lu is not None and ru is not None and lu != ru
                    and not _on_open_loop(self_var) and not _on_open_loop(other)):
                raise ValueError(_incompatible_units(lu, ru).replace(
                    "Cannot add/subtract", "Cannot compare", 1))
            result = variable_cls()
            result._set_expr(Compare(op_symbol, VarRef(self_var), VarRef(other)))
            fields = result.__dict__
            self._reject_keys_against_periods(self, other)
            lv, rv = self._align_keyed_values(self, other)
            expanded = self._expand_axised_operands(self, other, lv, rv)
            if expanded is not None:
                lv, rv = expanded
            # Two windows meet by date, as in ``_binary_op``.
            (lv, rv), window, runs_on, through = _meet((self, other), [lv, rv])
            if window is not None:
                fields['_start'], fields['_grain'], fields['_first_day'] = window
                fields['_read_by_date'] = True
            fields['_value'] = _combine(lv, rv)
            if window is None:
                _keep_runs_on(result, runs_on, through)
        else:
            result = variable_cls()
            result._set_expr(Compare(op_symbol, VarRef(self_var), Literal(other)))
            fields = result.__dict__
            fields['_value'] = _combine(self._value, other)

        value = fields['_value']
        if isinstance(value, list) and len(value) > 1:
            fields['var_type'] = 'list'
        elif type(value).__name__ == 'TrackValues' and value.time_length:
            fields['var_type'] = 'list'
        fields['_keys'] = self._propagate_keys(self, other)
        fields['_indexed_by'] = self._propagate_indexed_by(self, other)
        return result

    # ─── Comparison dunders ────────────────────────────────────

    # ``__eq__`` / ``__ne__`` deliberately return a :class:`Variable`
    # (a comparison expression) rather than ``bool`` — the engine treats
    # ``a == b`` as a formula-building operator. The override mismatch
    # against ``object`` is intentional, not a bug.
    def __eq__(self, other) -> "Variable":  # type: ignore[override]
        return self._comparison_op(other, '==', lambda a, b: a == b)

    def __ne__(self, other) -> "Variable":  # type: ignore[override]
        return self._comparison_op(other, '!=', lambda a, b: a != b)

    def __lt__(self, other) -> "Variable":
        return self._comparison_op(other, '<', lambda a, b: a < b)

    def __le__(self, other) -> "Variable":
        return self._comparison_op(other, '<=', lambda a, b: a <= b)

    def __gt__(self, other) -> "Variable":
        return self._comparison_op(other, '>', lambda a, b: a > b)

    def __ge__(self, other) -> "Variable":
        return self._comparison_op(other, '>=', lambda a, b: a >= b)
