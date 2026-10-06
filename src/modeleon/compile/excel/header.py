# SPDX-License-Identifier: Apache-2.0
"""Where a workbook's time lives, and how a cell reaches it.

A model's time is its window; a formula reads it as ``mo.time.days`` and
knows nothing of sheets. A workbook has no such thing as time, so the
writer materializes it:

* **The anchor.** The MultiVariable a model names as its
  ``default_header`` holds the time rows (``h.days = mo.time.days``) —
  on its own sheet they are real formulas off one start cell: the
  window's start Variable (``m.default_start = m.inputs.financial_close``)
  or a typed date.
* **Mirrors.** Every other sheet shows the header rows at its top as
  references to the anchor (``='Time'!C4``), unless it declares its own
  list of rows or ``None``.
* **References.** A formula reading ``mo.time.days`` on a sheet points
  at that sheet's own header row when it has one, else at the anchor —
  never at another sheet's header. A row mirrored into the header is
  read the same way.
* **No header** anywhere: time is spelled inline off the start
  (``EOMONTH(start, 2) - EOMONTH(start, 1)`` for a period's days), still
  a live formula, never a number.

The model stays one dependency (``interest → time.days``); only the
workbook's wiring depends on where the header is.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Any, Callable, Dict, List, Optional

from ...core.multi_variable import MultiVariableBase
from ...core.time import time_field_of
from ...core.variable import Variable

#: Months in one period of each grain.
STEP = {'month': 1, 'quarter': 3, 'year': 12}

#: No ``default_header`` anywhere up the tree.
UNSET = object()

#: The time rows that hold a value at the opening, the point just before
#: the first period: the end of "period 0" — the day before the start —
#: and its number, 0. The start and the days of a period that is not
#: there have no value.
OPENING_FIELDS = ('end', 'index')


def resolve_default_header(node: Any) -> Any:
    """The nearest ``default_header`` up ``node``'s tree — a MultiVariable,
    a list of rows, ``None`` (no header) or :data:`UNSET`."""
    seen: set = set()
    n = node
    while n is not None and id(n) not in seen:
        seen.add(id(n))
        own = getattr(n, '__dict__', {})
        if 'default_header' in own:
            return own['default_header']
        n = getattr(n, '_owner', None) or getattr(n, '_parent', None)
    return UNSET


def header_rows(value: Any) -> List[Variable]:
    """The rows a ``default_header`` value shows, top to bottom."""
    if value is None or value is UNSET:
        return []
    if isinstance(value, list):
        return list(value)
    rows: List[Variable] = []
    for _, comp in value.walk():
        if isinstance(comp, Variable) and not isinstance(comp, MultiVariableBase):
            rows.append(comp)
    return rows


def resolve_anchor(node: Any) -> Optional[MultiVariableBase]:
    """The MultiVariable that anchors the workbook's time: the nearest
    MV-valued ``default_header`` up the tree — or, where the model names
    its header as a LIST of rows (``m.default_header = [ruler.end,
    ruler.index, phases.operations_flag]``: the rows every sheet opens
    on), the section holding the first time row of that list. The list
    says what the sheets show; the section it reaches keeps the time."""
    seen: set = set()
    n = node
    while n is not None and id(n) not in seen:
        seen.add(id(n))
        value = getattr(n, '__dict__', {}).get('default_header', UNSET)
        if isinstance(value, MultiVariableBase):
            return value
        if isinstance(value, list) and _is_model_root(n):
            for row in value:
                if time_field_of(row) is not None:
                    owner = getattr(row, '_owner', None)
                    if isinstance(owner, MultiVariableBase):
                        return owner
        n = getattr(n, '_owner', None) or getattr(n, '_parent', None)
    return None


def _is_model_root(node: Any) -> bool:
    """A model — where a book-wide header is declared. A sheet's own list
    only says what that sheet shows."""
    from ...core.model import Model
    return isinstance(node, Model)


@dataclass
class SheetHeader:
    """One sheet's header as laid out: ``mirrors`` are ``(source, mirror)``
    pairs placed at the top; ``home`` are header rows that live on this
    sheet (the anchor's own sheet); ``key`` is the window they count."""
    mirrors: List[tuple] = field(default_factory=list)
    home: List[Variable] = field(default_factory=list)
    key: Optional[tuple] = None


@dataclass
class TimeAnchor:
    """One anchor's time rows as laid out: ``fields`` maps a time field to
    the row's id; ``ids`` are those rows. ``key`` is the window they
    count — ``(grain, start, periods)``."""
    key: tuple
    fields: Dict[str, str] = field(default_factory=dict)
    ids: set = field(default_factory=set)


@dataclass
class SheetTime:
    """How a formula on one sheet reaches time: ``fields`` maps a time
    field to the id of this sheet's header row carrying it; ``alias`` maps
    a mirrored row's id to its mirror's; ``anchor`` is the anchor of this
    sheet's window, when there is one."""
    fields: Dict[str, str] = field(default_factory=dict)
    alias: Dict[str, str] = field(default_factory=dict)
    header_key: Optional[tuple] = None
    anchor: Optional[TimeAnchor] = None
    #: The window of the sheet's periods (its captions).
    key: Optional[tuple] = None


@dataclass
class BookTime:
    """The workbook's time: every sheet's view of it. A formula reads time
    through the window it lives in — this sheet's header and the anchor
    only when their window is the formula's; otherwise off its window's
    own start."""
    sheets: Dict[str, SheetTime] = field(default_factory=dict)
    placed: set = field(default_factory=set)


def window_key(window: Any) -> Optional[tuple]:
    if window is None or window.grain is None or window.start is None:
        return None
    first_day = getattr(window, 'first_day', None)
    if first_day is not None:
        # A short first period is another window than the whole one.
        return (window.grain, window.start, window.periods, first_day)
    return (window.grain, window.start, window.periods)


def date_literal(d: date) -> str:
    return f"DATE({d.year},{d.month},{d.day})"


def _offset(start0: str, k: int, first_day: Optional[date]) -> str:
    """Months from the start of a quarter / year to the start cell's month:
    a number when the start is a date written into the formula, a formula
    when it is a cell (an edited cell moves the first period's end)."""
    if first_day is not None and start0.startswith('DATE('):
        return str((first_day.month - 1) % k)
    return f"MOD(MONTH({start0})-1,{k})"


def period_end_of(start: str, k: int) -> str:
    """The last day of the quarter / year whose month ``start`` falls in —
    ``start`` the 1st of a month."""
    return f"EOMONTH({start},{k - 1}-MOD(MONTH({start})-1,{k}))"


def months_between(start: str, end: str) -> str:
    """The calendar months from ``start`` (a 1st of a month) to ``end``."""
    return f"(YEAR({end} + 1) - YEAR({start})) * 12 + MONTH({end} + 1) - MONTH({start})"


def spell_inline(field_: str, i: int, grain: str, start0: str,
                 first_day: Optional[date] = None) -> str:
    """``mo.time.<field>`` of period ``i``, off the start cell alone.
    ``first_day`` set: the window's first period is short — it starts at
    ``start0`` and ends with its quarter / year."""
    if grain == 'day':
        day = f"({start0}+{i})" if i else start0
        return {'start': day, 'end': day, 'days': '1', 'months': '0',
                'index': str(i + 1)}[field_]
    k = STEP[grain]
    if first_day is not None and grain in ('quarter', 'year'):
        off = _offset(start0, k, first_day)

        def end_at(j: int) -> str:
            n = (j + 1) * k - 1
            shift = f"{n - int(off)}" if off.isdigit() else f"{n}-{off}"
            return f"EOMONTH({start0},{shift})"
        end = end_at(i)
        prev_end = end_at(i - 1) if i else None
        months0 = str(k - int(off)) if off.isdigit() else f"({k}-{off})"
        return {
            'start': start0 if i == 0 else f"({prev_end}+1)",
            'end': end,
            'days': f"({end}-{start0}+1)" if i == 0 else f"({end}-{prev_end})",
            'months': months0 if i == 0 else str(k),
            'index': str(i + 1),
        }[field_]
    end = f"EOMONTH({start0},{(i + 1) * k - 1})"
    prev_end = f"EOMONTH({start0},{i * k - 1})"
    return {
        'start': start0 if i == 0 else f"({prev_end}+1)",
        'end': end,
        'days': f"({end}-{prev_end})",
        # The months of a period are its grain's: a constant, not a date
        # formula - periods open on a month's first day.
        'months': str(k),
        'index': str(i + 1),
    }[field_]


def spell_anchor(field_: str, i: int, grain: str, start0: str,
                 cell: Callable[[str, int], Optional[str]],
                 first_day: Optional[date] = None) -> str:
    """An anchor row's own cell for period ``i``: chained off the other
    anchor rows the way a hand-built timing sheet is (``start = previous
    end + 1``, ``end = EOMONTH(start, …)``, ``days = end - start + 1``),
    falling back to the start cell where a row is missing. ``cell(f,
    -1)`` is the row's opening cell, when the book has an opening
    column: then the first period chains off it like every other one
    (``start = opening end + 1``) — one formula across the row.

    A window whose first period is short (``first_day``) ends each period
    with its quarter / year whatever month it starts in, and counts its
    months from its dates — again one formula across the row, which stays
    right when the start cell is edited in the workbook."""
    if first_day is not None and grain in ('quarter', 'year'):
        stub = _spell_anchor_stub(field_, i, grain, start0, cell, first_day)
        if stub is not None:
            return stub
    if grain == 'day':
        if field_ == 'start':
            if i == 0:
                opening = cell('end', -1)
                return f"{opening}+1" if opening else start0
            prev = cell('start', i - 1)
            return f"{prev}+1" if prev else spell_inline('start', i, grain, start0)
        if field_ == 'end':
            return cell('start', i) or spell_inline('end', i, grain, start0)
        if field_ == 'days':
            return '1'
        if field_ == 'months':
            return '0'
    k = STEP.get(grain, 1)
    if field_ == 'start':
        if i == 0:
            opening = cell('end', -1)
            return f"{opening}+1" if opening else start0
        prev_end = cell('end', i - 1)
        if prev_end:
            return f"{prev_end}+1"
        prev_start = cell('start', i - 1)
        if prev_start:
            return f"EOMONTH({prev_start},{k - 1})+1"
        return spell_inline('start', i, grain, start0)
    if field_ == 'end':
        own_start = cell('start', i)
        if own_start:
            return f"EOMONTH({own_start},{k - 1})"
        prev_end = cell('end', i - 1) if i else None
        if prev_end:
            return f"EOMONTH({prev_end},{k})"
        return spell_inline('end', i, grain, start0)
    if field_ == 'days':
        own_start, own_end = cell('start', i), cell('end', i)
        if own_start and own_end:
            return f"{own_end}-{own_start}+1"
        if own_end:
            return f"{own_end}-EOMONTH({own_end},{-k})"
        if own_start:
            return f"EOMONTH({own_start},{k - 1})-{own_start}+1"
        return spell_inline('days', i, grain, start0)
    if field_ == 'months':
        return _months_cell(i, start0, cell) or str(k)
    # index
    prev = cell('index', i - 1)
    if prev:
        return f"{prev}+1"
    return str(i + 1)


def _spell_anchor_stub(field_: str, i: int, grain: str, start0: str,
                       cell: Callable[[str, int], Optional[str]],
                       first_day: date) -> Optional[str]:
    """``spell_anchor`` for a window whose first period is short; ``None``
    where the whole-period chain is already right (start and index)."""
    k = STEP[grain]
    own_start, own_end = cell('start', i), cell('end', i)
    if field_ == 'end':
        if own_start:
            return period_end_of(own_start, k)
        prev_end = cell('end', i - 1) if i else None
        if prev_end:
            return f"EOMONTH({prev_end},{k})"
        return spell_inline('end', i, grain, start0, first_day)
    if field_ == 'start' and i > 0 and not cell('end', i - 1):
        prev_start = cell('start', i - 1)
        if prev_start:
            return f"{period_end_of(prev_start, k)}+1"
        return spell_inline('start', i, grain, start0, first_day)
    if field_ == 'days':
        if own_start and own_end:
            return f"{own_end}-{own_start}+1"
        if own_end:
            return (f"{own_end}-{start0}+1" if i == 0
                    else f"{own_end}-EOMONTH({own_end},{-k})")
        if own_start:
            return f"{period_end_of(own_start, k)}-{own_start}+1"
        return spell_inline('days', i, grain, start0, first_day)
    if field_ == 'months':
        return (_months_cell(i, start0, cell)
                or spell_inline('months', i, grain, start0, first_day))
    return None


def _months_cell(i: int, start0: str, cell: Callable[[str, int], Optional[str]]) -> Optional[str]:
    """An anchor row's months, from the period's own dates the way its
    days are (``end - start + 1``): the months from its first day to its
    last. The first day is the start row's cell, else the day after the
    previous period's end (the opening column's, for the first period),
    else — the first period — the window's start. ``None`` without an end
    row: nothing in the book holds the period's last day."""
    own_end = cell('end', i)
    if not own_end:
        return None
    start = cell('start', i)
    if not start:
        prev_end = cell('end', i - 1)
        start = f"{prev_end} + 1" if prev_end else (start0 if i == 0 else None)
    return months_between(start, own_end) if start else None


def build_book_time(engine: Any, roots: list, sheet_names: list,
                    addresses: Dict[str, Any]) -> Optional[BookTime]:
    """The workbook's time from a laid-out book, or ``None`` when no
    sheet has a window (nothing in the book lives in periods)."""
    from ...core.time import resolve_default_window
    if not any(window_key(w) for w in engine.window_by_sheet.values()):
        return None
    book = BookTime(placed=set(addresses))
    anchors: Dict[int, TimeAnchor] = {}
    for root, name in zip(roots, sheet_names):
        st = SheetTime(key=window_key(engine.window_by_sheet.get(id(root))))
        anchor_mv = resolve_anchor(getattr(root, '_sheet_source', None) or root)
        if anchor_mv is not None:
            if id(anchor_mv) not in anchors:
                anchor = TimeAnchor(key=window_key(resolve_default_window(anchor_mv)))
                for row in header_rows(anchor_mv):
                    f = time_field_of(row)
                    if f is not None and row.id in addresses:
                        anchor.fields.setdefault(f, row.id)
                        anchor.ids.add(row.id)
                anchors[id(anchor_mv)] = anchor
            st.anchor = anchors[id(anchor_mv)]
        plan = engine.header_by_sheet.get(id(root))
        if plan is not None:
            st.header_key = plan.key
            for row in plan.home:
                f = time_field_of(row)
                if f is not None and row.id in addresses:
                    st.fields.setdefault(f, row.id)
            home = {id(row) for row in plan.home}
            for src, mirror in plan.mirrors:
                if id(src) in home:
                    # The anchor repeating its own rows at its top: a
                    # formula there still reads the row itself.
                    continue
                st.alias[src.id] = mirror.id
                f = time_field_of(src)
                if f is not None:
                    st.fields.setdefault(f, mirror.id)
        book.sheets[name] = st
    return book


def mirror_of(src: Variable) -> Variable:
    """A header row on another sheet: a reference to ``src``, carrying its
    values so every renderer can fall back to them."""
    from ...core.expr import VarRef
    mirror = Variable(display_name=src.display_name)
    value = src._value
    mirror._value = list(value) if isinstance(value, list) else value
    mirror.var_type = src.var_type
    mirror.value_type = src.value_type
    mirror._excel_props = dict(getattr(src, '_excel_props', {}) or {})
    # What the row counts and what it is: its unit column and its format
    # follow the row it repeats.
    mirror._unit = getattr(src, '_unit', None)
    mirror._regrain = getattr(src, '_regrain', None)
    mirror._set_expr(VarRef(src))
    mirror._mirror_of = src
    return mirror
