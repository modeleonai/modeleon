# SPDX-License-Identifier: Apache-2.0
"""Statistics — ``MEDIAN``, ``STDEV_S``, ``STDEV_P``, ``VAR_S``, ``VAR_P``,
``PERCENTILE_INC``, ``PERCENTILE_EXC``, ``QUARTILE_INC``, ``QUARTILE_EXC``,
``LARGE``, ``SMALL``, ``RANK_EQ``, ``RANK_AVG`` — and the names an older
workbook carries for seven of them: ``STDEV``, ``STDEVP``, ``VAR``,
``VARP``, ``PERCENTILE``, ``QUARTILE``, ``RANK``.

Each is written as the spreadsheet writes it: ``=MEDIAN(Comps!C3:C7)``,
``=_xlfn.PERCENTILE.INC(Comps!C3:C7, 0.25)``, ``=_xlfn.RANK.EQ(C3,
C3:C7)``. A name Excel spells with a dot is written in Python with an
underscore — ``PERCENTILE_INC`` is ``PERCENTILE.INC`` — and, Excel 2010's,
stored with the ``_xlfn.`` prefix; the older names are the same functions
under their own names (``STDEV`` is ``STDEV_S``, ``RANK`` is ``RANK_EQ``).

Several series are combined period by period, as ``a + b`` combines
them; one series alone is reduced over its periods; scalars only stay one
number. ``MEDIAN`` and the spreads (``STDEV_*``, ``VAR_*``) over ONE
argument reduce its cells to one number: a row, written as the range of
its cells (``=MEDIAN(B2:M2)``), or a list of scalar lines whose cells
stand one under another — one field down the records of a table, written
as a vertical range, ``Comps!C3:C7``. Over two or more values with a
series among them — rows, a list typed in (one value per period, as the
operators read it), the lines of a list of lines — the call is one per
period over that period's cells: ``MEDIAN(a, b)`` is ``=MEDIAN(C5, C9)``
in every column; a scalar, a scalar line and a one-value series stand in
every period, series of other lengths are refused, and a line with tracks
lifts per track. Over scalars only it is one number, ``=MEDIAN(B2, B3)``.
The others read ONE range (``PERCENTILE``'s array, ``RANK``'s ref): a
row, or a list of lines. A list of scalar lines is one sample; a list of
two or more lines with a series among them — a field of yearly lines down
a records table, or sibling lines — is one sample per period, and the
call is one per period over that period's cells, written as the run they
form in its column (``=_xlfn.PERCENTILE.INC(Peers!C3:C7, 0.25)``), with
``RANK``'s number this period's value of the line ranked. A list holding
one series alone is refused: the line itself is read over its periods.
Cells that form no run cannot be one range, so the cell carries the
computed value and ``check_compat`` lists it — never a comma list, which
these functions read as further arguments. Inside a loop container a call
over two or more values runs one period at a time. ``k``, ``quart``,
``RANK``'s number and its order are a line — one value, or a series for
one call per period — or a value typed in (a list typed in is such a
series, one value per period, never the spreadsheet's array of ``k``),
read as the spreadsheet reads a number: a blank is 0,
TRUE 1, a date its serial, numeric text its number, other text
``#VALUE!``. ``RANK`` of each record among its peers is one call per
record against the same vertical range:
``mo.RANK_EQ(c.multiple, [c.multiple for c in peers])``.

The edges, as the spreadsheet reads them. In a range or a line, numbers
and dates count (a date as its serial) and a blank, text or a boolean is
skipped; typed into the call, a number, TRUE / FALSE, numeric text and
None (the 0 the formula holds) count, other text is ``#VALUE!``. An error
is the result: in the reductions the first one in reading order; in the
others, as LibreOffice reads them (Excel documents no order), a ``k`` or
a quart that is an error or out of its range before the range's error —
``LARGE`` / ``SMALL`` hold ``k`` against the count after the range — and
``RANK`` its order, its range, then its number. No number at all
is ``#NUM!`` — ``#N/A`` for ``RANK``, which finds nothing to rank among —
and ``STDEV_S`` / ``VAR_S`` need two numbers, ``STDEV_P`` / ``VAR_P`` one
(``#DIV/0!``). ``PERCENTILE_INC``'s ``k`` lies in [0, 1], interpolating
between the cells at ``(n - 1) * k``; ``PERCENTILE_EXC``'s inside (0, 1)
and at ``(n + 1) * k`` between 1 and n; ``QUARTILE_INC``'s quart is 0 to
4 and ``QUARTILE_EXC``'s 1 to 3, a fraction truncated, each the percentile
at quart / 4 (2 is the median); out of its range a ``k`` or a quart is
``#NUM!``. ``LARGE`` takes a fractional ``k`` up and ``SMALL`` down
(``LARGE(x, 2.5)`` is the third largest, ``SMALL(x, 2.5)`` the second
smallest) — LibreOffice's reading; Excel documents none — and ``k`` past
the count is ``#NUM!``. ``RANK`` finds its number by exact equality,
``order`` 0 (the default) ranking the largest first and any other the
smallest; ties share the best rank (``RANK_EQ``) or their mean
(``RANK_AVG``); a number the range lacks is ``#N/A``. A position within
15 significant digits of a whole one is that position, as the sheet
reads it: ``100 * 0.29``, 28.999999999999996, is the 30th cell.
The middle, a percentile, a quartile and the k-th of dates is a date; a
spread of dates is a number of days. The values a call reads are in one
unit, as ``+`` needs them — two are refused, ``RANK``'s number against
its range too (``k``, a quart and an order are no values) — and the
result is in it: a variance in its square, a rank in none.

LibreOffice departs from Excel in five places, and the engine keeps
Excel's rule: there a boolean in a range is a number; numeric text typed
into ``MEDIAN`` or a spread is ``#VALUE!``, and so is numeric text as
``LARGE`` / ``SMALL``'s ``k``, typed or in a cell, which LibreOffice
reads from a number only — the engine reads that ``k`` as Excel reads a
value argument, though Excel documents nothing for it; a ``k`` or a
quart out of its range and a range without numbers are written as
``#VALUE!``, not ``#NUM!`` (``RANK``'s: not ``#N/A``); a negative
fractional quart is rounded down, not truncated; the spreads read their
arguments from the last, so of two errors the later one shows.

Not covered: one line with tracks alone, and a parameter with tracks
(slice a coordinate first), the spreadsheet's array of ``k``
(``LARGE(x, {1, 2, 3})``, one result per ``k``), and ``MODE``,
``PERCENTRANK``, ``AVEDEV``.
"""

from __future__ import annotations

import bisect
import math
import numbers
from datetime import date, timedelta
from typing import Any, Callable, Sequence

from ..core.expr import Expr, ListExpr, Literal, VarRef
from ..core.multi_variable import MultiVariableBase
from ..core.shape import keys_along, reject_axised
from ..core.tracks import TrackValues
from ..core.unit import Unit
from ..core.variable import Variable
from ..core.variable_ops import _to_serial
from ._helpers import (
    UnitLaw, _per_period, align_keyed, align_located, make_func_var, placed,
    refuse_mixed_units, same_unit, shared_unit, src, unit_law,
)
from .criteria import (
    _ERRORS, _cells, _checked, _from_serial, _label, _Lines, _Range, _range, _series_value,
    _shared, _unit_items,
)
from .registry import per_period_arity, register
from .text import _number_of

#: An optional argument the call did not give — left out of the formula too.
_OMITTED: Any = object()

#: The most arguments a spreadsheet function takes.
_MAX_ARGS = 255

# ─── What a statistic reads ─────────────────────────────────────────


# A position computed off a number is read to 15 significant digits, as
# the spreadsheet reads it: 2.9999999999999996 is 3.
def _approx(x: Any) -> Any:
    if isinstance(x, float) and math.isfinite(x):
        return float(f"{x:.15g}")
    return x


# What a statistic reads of a range, in order — numbers, dates and
# durations — or the first error among them; a blank, text and a boolean
# are skipped.
def _counted(cells: Sequence[Any]) -> list[Any] | str:
    out: list[Any] = []
    for cell in cells:
        if isinstance(cell, str):
            if cell in _ERRORS:
                return cell
        elif isinstance(cell, (date, timedelta)):
            out.append(cell)
        elif isinstance(cell, numbers.Real) and not isinstance(cell, bool):
            if not math.isfinite(cell):
                return '#NUM!'
            out.append(cell)
    return out


# The counted cells sorted, with their serials worked out once for every
# call that ranks a number among them.
class _Ranked(list[Any]):
    __slots__ = ('_serials',)
    _serials: list[Any]

    def serials(self) -> list[Any]:
        try:
            return self._serials
        except AttributeError:
            self._serials = [_to_serial(c) for c in self]
            return self._serials


# What a range function picks from: its cells counted and sorted, or the
# first error among them.
def _ranked_of(cells: Sequence[Any]) -> Any:
    counted = _counted(cells)
    return counted if isinstance(counted, str) else _Ranked(sorted(counted, key=_to_serial))


# A k, a quart, a number or an order as the spreadsheet reads a number. An
# error a cell holds is that error; typed into the call, such text is text.
def _number(value: Any, typed: bool) -> Any:
    if not typed and isinstance(value, str) and value in _ERRORS:
        return value
    try:
        return _number_of(value)
    except TypeError:
        return '#VALUE!'
    except OverflowError:
        return '#NUM!'


# A value typed into a reduction's call: a date stays a date (its serial
# counts); any other is read as a number, None as the 0 the formula holds.
def _typed(value: Any) -> Any:
    if isinstance(value, (date, timedelta)):
        return value
    return _number(value, True)


def _lines(arg: Any) -> list[Variable]:
    if isinstance(arg, Variable):
        return [arg]
    if isinstance(arg, (list, tuple)):
        return [item for item in arg if isinstance(item, Variable)]
    return []


# A number computed between cells is a date when they are dates.
def _as_cell(number: float, among: Sequence[Any]) -> Any:
    if not math.isfinite(number):
        return '#NUM!'
    if all(isinstance(c, date) for c in among):
        return _from_serial(number)
    return number


def _type_of(value: Any, otherwise: str) -> str:
    fibers = list(value.values()) if isinstance(value, TrackValues) else [value]
    for fiber in fibers:
        for v in (fiber if isinstance(fiber, list) else [fiber]):
            if not isinstance(v, str):
                return 'datetime' if isinstance(v, date) else otherwise
    return otherwise


#: Where a result runs: the combined window its operands were read on by
#: date and the keys they ran along (``_aligned``); either may be None.
_Place = tuple[Any, list[Any] | None]
_NOWHERE: _Place = (None, None)


def _result(func: str, nodes: list[Expr], texts: list[str], value: Any,
            value_type: str, unit: Any, place: _Place = _NOWHERE) -> Variable:
    if isinstance(value, TrackValues):
        var_type = 'list' if value.time_length else 'scalar'
    else:
        var_type = 'list' if isinstance(value, list) else 'scalar'
    with _checked(nodes):
        out = make_func_var(
            func, nodes, value, value_type, var_type,
            source_code=f"{func}({', '.join(texts)})",
        )
    out._unit = unit
    return placed(out, *place)


# The values a call reads per period, read as ``a + b`` reads its operands:
# series on other time windows by date (each continued by its own extend=
# rule), lines along keys key by key; ``owners[i]`` is the line ``raws[i]``
# came from (None for a value typed in). With where the result runs.
def _aligned(where: str, owners: Sequence[Any], raws: list[Any]) -> tuple[list[Any], _Place]:
    raws, keys = align_keyed(where, owners, raws)
    raws, window = align_located(owners, raws)
    return raws, (window, keys)


# Where every line of a range read once runs, worked out once: the window
# and the length of its series when all are plain series of one window —
# None where one runs along keys, holds tracks or another window, which
# each call aligns (``_aligned``).
def _range_window(shared: _Lines) -> tuple[Any, int] | None:
    first: tuple[Any, int] | None = None
    for line, raw in zip(shared.lines, shared.raws):
        if not isinstance(raw, list) or keys_along(line) is not None:
            return None
        loc = line.time
        if loc is None or loc.start is None:
            return None
        if first is None:
            first = (loc, len(raw))
        elif (loc, len(raw)) != first:
            return None
    return first


# The values of a call over a range of lines, read with the range's as
# ``_aligned`` reads them. Beside a range read once whose lines are plain
# series of one window, values that add no other window and no keys leave
# nothing to align — each call of a rank per record over a table reads its
# own line, not every peer's window again.
def _with_range(where: str, shared: _Lines | None, ranged: Sequence[Any],
                owners: Sequence[Any], cells: list[Any],
                raws: list[Any]) -> tuple[list[Any], _Place]:
    if shared is not None:
        window = shared.keep('window', lambda: _range_window(shared))
        if window is not None and all(_on_window(window, owner, raw)
                                      for owner, raw in zip(owners, raws)):
            return [*cells, *raws], _NOWHERE
    return _aligned(where, [*ranged, *owners], [*cells, *raws])


def _on_window(window: tuple[Any, int], owner: Any, raw: Any) -> bool:
    if not isinstance(owner, Variable):
        return True                     # a value typed in is read as it is
    if keys_along(owner) is not None or isinstance(raw, TrackValues):
        return False
    if not isinstance(raw, list):
        return True
    loc = owner.time
    return loc is not None and (loc, len(raw)) == window


# A variance is in the square of its cells' unit; a label squared is none.
# What is no unit — none, or one a loop has not derived yet — stays as is.
def _squared(unit: Any) -> Any:
    if not isinstance(unit, Unit):
        return unit
    unit = unit ** 2
    return None if unit.is_dimensionless else unit


# ─── Several series: one call per period ────────────────────────────


def _is_series(x: Any) -> bool:
    return isinstance(x, Variable) and _series_value(x._value)


# A list or a tuple holding a line is a list of lines (with values typed
# in among them): a range, each line one value of the sample.
def _of_lines(arg: Any) -> bool:
    return isinstance(arg, (list, tuple)) and any(isinstance(i, Variable) for i in arg)


# Is a call over these values one call per period: as many as the
# function's per_period_arity or more, a series among them? A reduction's
# values are its arguments — a list typed in a series, as the operators
# read it — and the lines of a list of lines; a range function's are the
# lines its range lists.
def _by_period(where: str, values: Sequence[Any], typed_series: bool) -> bool:
    count, shaped = 0, False
    for v in values:
        if _of_lines(v):
            count += len(v)
            shaped = shaped or any(_is_series(i) for i in v)
        else:
            count += 1
            shaped = shaped or _is_series(v) or (typed_series and isinstance(v, list) and bool(v))
    return shaped and count >= (per_period_arity(where) or 2)


# A list of lines that holds one series alone is no sample per period: the
# line itself is read over its periods.
def _refuse_one_series(where: str, arg: Any, example: str) -> None:
    if _of_lines(arg) and len(arg) == 1 and _is_series(arg[0]):
        line = _label(arg[0])
        raise ValueError(
            f"{where}(): [{line}] lists one line, and a list of lines with a series in "
            f"it is one sample per period, which needs two or more values; "
            f"{example.format(line=line)}"
        )


# The values a call reads per period, in reading order — a line's cells, a
# list of lines' lines and values, a value typed in (a list typed in a
# series of them) — with whether each was typed in, the AST nodes, the
# names in the source, the lines and the line each value came from.
def _per_period_values(args: Sequence[Any], where: str) -> tuple[
        list[Any], list[bool], list[Expr], list[str], list[Variable], list[Any]]:
    raws: list[Any] = []
    typed: list[bool] = []
    nodes: list[Expr] = []
    texts: list[str] = []
    lines: list[Variable] = []
    owners: list[Any] = []
    for a in args:
        if isinstance(a, Variable):
            reject_axised(a, where)
            raws.append(a._value)
            typed.append(False)
            nodes.append(VarRef(a))
            texts.append(src(a))
            lines.append(a)
            owners.append(a)
        elif isinstance(a, (list, tuple)) and not a:
            raise ValueError(f"{where}(): a range needs at least one cell; got {a!r}.")
        elif isinstance(a, tuple) or _of_lines(a):
            items: list[Expr] = []
            labels: list[str] = []
            for item in a:
                if isinstance(item, Variable):
                    reject_axised(item, where)
                    raws.append(item._value)
                    items.append(VarRef(item))
                    labels.append(_label(item))
                    lines.append(item)
                    owners.append(item)
                elif isinstance(item, (list, tuple, dict, set, MultiVariableBase)):
                    raise TypeError(
                        f"{where}(): an item of a list of lines is a line or one value; "
                        f"got {type(item).__name__}."
                    )
                else:
                    # A value among lines is the array's own, read as a cell.
                    raws.append(item)
                    items.append(Literal(item))
                    labels.append(repr(item))
                    owners.append(None)
                typed.append(False)
            nodes.append(ListExpr(items))
            texts.append(f"[{', '.join(labels)}]")
        elif isinstance(a, (dict, set, MultiVariableBase)):
            raise TypeError(
                f"{where}(): an argument is a row, a list of lines, a line or a value "
                f"typed in; got {type(a).__name__}."
            )
        else:
            raws.append(list(a) if isinstance(a, list) else a)
            typed.append(True)
            nodes.append(Literal(list(a) if isinstance(a, list) else a))
            texts.append(repr(a))
            owners.append(None)
    return raws, typed, nodes, texts, lines, owners


# ``call`` over each period's values; a scalar or a one-value series stands
# in every period, other lengths are refused, a line with tracks lifts per
# track.
def _each_period(raws: list[Any], call: Callable[[list[Any]], Any]) -> Any:
    if any(isinstance(r, TrackValues) for r in raws):
        return TrackValues.lift(lambda *fibers: _per_period(list(fibers), call), *raws)
    return _per_period(raws, call)


# ─── The reductions ─────────────────────────────────────────────────


# Every cell a reduction reads, in reading order, with the AST nodes, the
# names in the source and the one unit its lines may be in.
def _operands(args: tuple[Any, ...], where: str) -> tuple[list[Any], list[Expr],
                                                          list[str], Any]:
    cells: list[Any] = []
    nodes: list[Expr] = []
    texts: list[str] = []
    for a in args:
        if isinstance(a, (Variable, list, tuple)):
            rng = _range(a, where)
            cells.extend(rng.values)
            nodes.append(rng.node)
            texts.append(rng.text)
        elif isinstance(a, (dict, set, MultiVariableBase)):
            raise TypeError(
                f"{where}(): an argument is a row, a list of lines, a line or a value "
                f"typed in; got {type(a).__name__}."
            )
        else:
            cells.append(_typed(a))
            nodes.append(Literal(a))
            texts.append(repr(a))
    lines = [line for a in args for line in _lines(a)]
    refuse_mixed_units(where, lines)
    return cells, nodes, texts, shared_unit(lines)


# reduce(the counted cells, or the first error among them) — over every
# cell of every argument, or, over several series, once per period over
# that period's cells; and the AST nodes, the names in the source, the
# unit the lines share and where the result runs.
def _reduction(where: str, args: tuple[Any, ...], reduce: Callable[[list[Any]], Any]
               ) -> tuple[Any, list[Expr], list[str], Any, _Place]:
    if not args:
        raise ValueError(
            f"{where}() needs at least one argument. "
            f"Syntax: mo.{where}(row)  or  mo.{where}(a, b, c). Got 0 arguments."
        )
    if len(args) > _MAX_ARGS:
        raise ValueError(
            f"{where}() takes at most {_MAX_ARGS} arguments, as the spreadsheet does; "
            f"got {len(args)}. Pass the lines as one list — one range."
        )
    shared = _shared(args[0]) if len(args) == 1 else None
    if shared is not None:
        return _one_reduction(where, args[0], shared, reduce)
    if _by_period(where, args, typed_series=True):
        raws, typed, nodes, texts, lines, owners = _per_period_values(args, where)
        refuse_mixed_units(where, lines)
        raws, place = _aligned(where, owners, raws)

        def one(values: list[Any]) -> Any:
            counted = _counted([_typed(v) if t else v for v, t in zip(values, typed)])
            return counted if isinstance(counted, str) else reduce(counted)

        return _each_period(raws, one), nodes, texts, shared_unit(lines), place
    for a in args:
        _refuse_one_series(where, a, f"mo.{where}({{line}}) reduces the line over its periods.")
    cells, nodes, texts, unit = _operands(args, where)
    counted = _counted(cells)
    return ((counted if isinstance(counted, str) else reduce(counted)), nodes, texts, unit,
            _NOWHERE)


# A reduction over one range read once (``_shared``): what it reads and
# what it gives, as ``_reduction`` reads and refuses them, the value worked
# out once per function — each call gets a list of its own.
def _one_reduction(where: str, arg: Any, shared: _Lines, reduce: Callable[[list[Any]], Any]
                   ) -> tuple[Any, list[Expr], list[str], Any, _Place]:
    def reduced(cells: list[Any]) -> Any:
        counted = _counted(cells)
        return counted if isinstance(counted, str) else reduce(counted)

    lines = shared.unit_lines()[0]
    if (not shared.row and shared.series
            and len(shared.lines) >= (per_period_arity(where) or 2)):
        rng = shared.keep('per period', lambda: _per_period_range(shared, where))
        refuse_mixed_units(where, lines)
        if any(isinstance(raw, TrackValues) for raw in rng.values):
            value = _each_period(list(rng.values), reduced)
            return value, [rng.node], [rng.text], rng.unit, _NOWHERE
        # The lines read on one window and along one set of keys, worked
        # out once: the same for every call over them.
        cells, place = shared.keep(
            'aligned', lambda: _aligned(where, shared.lines, list(rng.values)))
        value = shared.keep(('reduced', where), lambda: _each_period(list(cells), reduced))
        value = list(value) if isinstance(value, list) else value
        return value, [rng.node], [rng.text], rng.unit, place
    _refuse_one_series(where, arg, f"mo.{where}({{line}}) reduces the line over its periods.")
    rng = shared.cells(where)
    refuse_mixed_units(where, lines)
    value = shared.keep(('reduced', where), lambda: reduced(rng.values))
    return value, [rng.node], [rng.text], rng.unit, _NOWHERE


# A list of lines read once as one sample per period: each line's value as
# it holds it, as ``_per_period_values`` reads a list of lines.
def _per_period_range(shared: _Lines, where: str) -> _Range:
    for var in shared.lines:
        reject_axised(var, where)
    return _Range(list(shared.raws), shared.node, shared.text, shared_unit(shared.lines),
                  shared.first_type, shared)


# Each period's cells of a list of lines, counted and sorted, worked out
# once for every call over them — None when a line holds tracks or the lines'
# lengths differ, which each call reads (and refuses) as before.
def _ranked_per_period(raws: list[Any]) -> list[Any] | None:
    if any(isinstance(raw, TrackValues) for raw in raws):
        return None
    lengths = {len(raw) for raw in raws if isinstance(raw, list)}
    if not lengths or lengths - {1, max(lengths)}:
        return None
    ranked: list[Any] = _per_period(list(raws), _ranked_of)
    return ranked


def _median(ranked: list[Any]) -> Any:
    n = len(ranked)
    if n % 2:
        return ranked[n // 2]
    lo, hi = ranked[n // 2 - 1], ranked[n // 2]
    return _as_cell((_to_serial(lo) + _to_serial(hi)) / 2, (lo, hi))


# Two passes over exact sums, so 1e15 + 1, + 2, + 3 spread by exactly 1.
def _variance(values: list[Any], ddof: int) -> Any:
    n = len(values)
    if n <= ddof:
        return '#DIV/0!'
    try:
        mean = math.fsum(values) / n
        out = math.fsum((v - mean) ** 2 for v in values) / (n - ddof)
    except OverflowError:
        return '#NUM!'
    return out if math.isfinite(out) else '#NUM!'


def _spread(where: str, args: tuple[Any, ...], ddof: int, root: bool) -> Variable:
    def spread(counted: list[Any]) -> Any:
        value = _variance([_to_serial(c) for c in counted], ddof)
        return math.sqrt(value) if root and not isinstance(value, str) else value

    value, nodes, texts, unit, place = _reduction(where, args, spread)
    return _result(where, nodes, texts, value, 'float', unit if root else _squared(unit), place)


def _middle(counted: list[Any]) -> Any:
    return _median(sorted(counted, key=_to_serial)) if counted else '#NUM!'


def MEDIAN(*args: Any) -> Variable:
    """The middle number of the cells (renders ``=MEDIAN(...)``): of an
    even count, the mean of the two middle ones; ``#NUM!`` with no number.
    Over ONE row or list of scalar lines, its cells; over two or more values
    with a series among them, each period's cells — ``MEDIAN(a, b)`` is one
    call per period, as ``a + b`` is; over scalars only, one number.

    Example::

        mid = mo.MEDIAN([c.a.multiple, c.b.multiple, c.c.multiple])
        # =MEDIAN(Comps!C3:C5)
        mo.MEDIAN(revenue)                      # =MEDIAN(B2:M2)
        mo.MEDIAN([p.revenue for p in peers])   # each year: =MEDIAN(Peers!C3:C7)
        mo.MEDIAN(plan, actual)                 # each period: =MEDIAN(C5, C9)
    """
    where = "MEDIAN"
    value, nodes, texts, unit, place = _reduction(where, args, _middle)
    return _result(where, nodes, texts, value, _type_of(value, 'float'), unit, place)


def STDEV_S(*args: Any) -> Variable:
    """The standard deviation of a sample (renders ``=STDEV.S(...)``,
    stored as ``_xlfn.STDEV.S``): over ``n - 1``; ``#DIV/0!`` with fewer
    than two numbers. ``mo.STDEV_S(returns)`` → ``=_xlfn.STDEV.S(B2:M2)``."""
    return _spread("STDEV_S", args, 1, True)


def STDEV_P(*args: Any) -> Variable:
    """The standard deviation of a whole population (renders
    ``=STDEV.P(...)``, stored as ``_xlfn.STDEV.P``): over ``n``;
    ``#DIV/0!`` with no number."""
    return _spread("STDEV_P", args, 0, True)


def VAR_S(*args: Any) -> Variable:
    """The variance of a sample (renders ``=VAR.S(...)``, stored as
    ``_xlfn.VAR.S``): over ``n - 1``; ``#DIV/0!`` with fewer than two
    numbers. In the square of the cells' unit."""
    return _spread("VAR_S", args, 1, False)


def VAR_P(*args: Any) -> Variable:
    """The variance of a whole population (renders ``=VAR.P(...)``,
    stored as ``_xlfn.VAR.P``): over ``n``; ``#DIV/0!`` with no number."""
    return _spread("VAR_P", args, 0, False)


def STDEV(*args: Any) -> Variable:
    """:func:`STDEV_S` under the name an older workbook carries (renders
    ``=STDEV(...)``)."""
    return _spread("STDEV", args, 1, True)


def STDEVP(*args: Any) -> Variable:
    """:func:`STDEV_P` under the name an older workbook carries (renders
    ``=STDEVP(...)``)."""
    return _spread("STDEVP", args, 0, True)


def VAR(*args: Any) -> Variable:
    """:func:`VAR_S` under the name an older workbook carries (renders
    ``=VAR(...)``)."""
    return _spread("VAR", args, 1, False)


def VARP(*args: Any) -> Variable:
    """:func:`VAR_P` under the name an older workbook carries (renders
    ``=VARP(...)``)."""
    return _spread("VARP", args, 0, False)


# ─── One range ──────────────────────────────────────────────────────


# A k, a quart, a number or an order: a line — one value, or a series for
# one call per period — or a value typed in, and whether it was. Empty text
# typed into a line is the blank cell the workbook holds.
def _param(value: Any, where: str, name: str) -> tuple[Any, Expr, str, bool]:
    if isinstance(value, Variable):
        reject_axised(value, where)
        raw = value._value
        if hasattr(raw, 'roles'):
            raise ValueError(
                f"{where}(): {src(value)} has tracks — slice a coordinate first "
                f"(.at(track='...'))."
            )
        cells = _cells(value, raw)
        return (cells if isinstance(raw, list) else cells[0]), VarRef(value), src(value), False
    if isinstance(value, (dict, set, tuple, MultiVariableBase)) or (
            isinstance(value, list) and any(isinstance(v, Variable) for v in value)):
        raise TypeError(
            f"{where}(): {name} is one value, or a line (one call per period); "
            f"got {type(value).__name__}."
        )
    return value, Literal(value), repr(value), True


# pick(the range's cells sorted ascending — or its error —, the parameters
# read as numbers) per period; the range is argument ``at`` of the call:
# read whole by every period's call, or, a list of lines with series among
# them, this period's cells of each. The values compared or picked among
# are in one unit: the range's lines and those of ``compared`` (RANK's
# number); a k or an order is not one. Series on other time windows are
# read by date, lines along keys key by key (``_aligned``).
def _one_range(where: str, array: Any, at: int, params: list[tuple[str, Any]],
               pick: Callable[..., Any], compared: Sequence[Any] = ()
               ) -> tuple[Any, _Range, list[Expr], list[str], _Place]:
    shared = _shared(array)
    if shared is not None:
        by_period = (not shared.row and shared.series
                     and len(shared.lines) >= (per_period_arity(where) or 2))
    else:
        by_period = _by_period(where, [array], typed_series=False)
    ranged_owners: list[Any] = []
    if by_period:
        if shared is not None:
            rng = shared.keep('per period', lambda: _per_period_range(shared, where))
            ranged_owners = list(shared.lines)
        else:
            cells, _, items, labels, lines, ranged_owners = _per_period_values([array], where)
            rng = _Range(cells, items[0], labels[0], shared_unit(lines), lines[0].value_type)
    else:
        shown = [name for name, _ in params]
        shown.insert(at, "{line}")
        _refuse_one_series(
            where, array, f"mo.{where}({', '.join(shown)}) reads the line over its periods.")
        rng = shared.cells(where) if shared is not None else _range(array, where)
    ranged = list(shared.unit_lines()[0]) if shared is not None else _lines(array)
    refuse_mixed_units(where, [line for a in compared for line in _lines(a)] + ranged)
    raws: list[Any] = []
    typed: list[bool] = []
    nodes: list[Expr] = []
    texts: list[str] = []
    owners: list[Any] = []
    for name, value in params:
        raw, node, text, is_typed = _param(value, where, name)
        raws.append(raw)
        typed.append(is_typed)
        nodes.append(node)
        texts.append(text)
        owners.append(value if isinstance(value, Variable) else None)
    nodes.insert(at, rng.node)
    texts.insert(at, rng.text)

    def numbers(values: list[Any]) -> list[Any]:
        return [_number(v, t) for v, t in zip(values, typed)]

    if by_period:
        n = len(rng.values)
        values, place = _with_range(where, shared, ranged_owners, owners, rng.values, raws)
        sorted_by = (shared.keep('ranked per period', lambda: _ranked_per_period(rng.values))
                     if shared is not None and place == _NOWHERE else None)
        if sorted_by is not None:
            def picked(values: list[Any]) -> Any:
                return pick(values[0], *numbers(values[1:]))

            return _per_period([sorted_by, *raws], picked), rng, nodes, texts, place

        def each(values: list[Any]) -> Any:
            return pick(_ranked_of(values[:n]), *numbers(values[n:]))

        return _each_period(values, each), rng, nodes, texts, place

    ranked = (shared.keep('ranked', lambda: _ranked_of(rng.values)) if shared is not None
              else _ranked_of(rng.values))

    def one(values: list[Any]) -> Any:
        return pick(ranked, *numbers(values))

    raws, place = _aligned(where, owners, raws)
    return _per_period(raws, one), rng, nodes, texts, place


# The cell at 0-based position ``at`` of the sorted cells, or the point
# between two neighbours on the line joining them.
def _interpolated(ranked: list[Any], at: float) -> Any:
    lo = math.floor(_approx(at))
    part = at - lo
    if part <= 0 or lo + 1 >= len(ranked):
        return ranked[lo]
    a, b = ranked[lo], ranked[lo + 1]
    x, y = _to_serial(a), _to_serial(b)
    return _as_cell(x + part * (y - x), (a, b))


def _inclusive(ranked: list[Any] | str, k: Any) -> Any:
    if isinstance(k, str):
        return k
    if not 0 <= k <= 1:
        return '#NUM!'
    if isinstance(ranked, str):
        return ranked
    if not ranked:
        return '#NUM!'
    return _interpolated(ranked, k * (len(ranked) - 1))


# The exclusive percentile lies at (n + 1) * k, between 1 and n.
def _between(ranked: list[Any], k: Any) -> Any:
    n = len(ranked)
    rank = k * (n + 1)
    if rank < 1 or rank > n:
        return '#NUM!'
    return _interpolated(ranked, rank - 1)


def _exclusive(ranked: list[Any] | str, k: Any) -> Any:
    if isinstance(k, str):
        return k
    if not 0 < k < 1:
        return '#NUM!'
    if isinstance(ranked, str):
        return ranked
    if not ranked:
        return '#NUM!'
    return _between(ranked, k)


def _quartile(ranked: list[Any] | str, quart: Any, *, inclusive: bool) -> Any:
    if isinstance(quart, str):
        return quart
    q = math.trunc(_approx(quart))
    if not (0 <= q <= 4 if inclusive else 1 <= q <= 3):
        return '#NUM!'
    if isinstance(ranked, str):
        return ranked
    if not ranked:
        return '#NUM!'
    if q == 2:
        return _median(ranked)
    if inclusive:
        return _interpolated(ranked, 0.25 * q * (len(ranked) - 1))
    return _between(ranked, 0.25 * q)


def _kth(ranked: list[Any] | str, k: Any, *, largest: bool) -> Any:
    if isinstance(k, str):
        return k
    if isinstance(ranked, str):
        return ranked
    n = len(ranked)
    if not n:
        return '#NUM!'
    at = math.ceil(_approx(k)) if largest else math.floor(_approx(k))
    if not 1 <= at <= n:
        return '#NUM!'
    return ranked[n - at] if largest else ranked[at - 1]


def _ranked_type(rng: _Range) -> str:
    return rng.value_type if rng.value_type in ('int', 'float') else 'float'


def PERCENTILE_INC(array: Any, k: Any) -> Variable:
    """The ``k``-th percentile of the cells, ``k`` in [0, 1] (renders
    ``=PERCENTILE.INC(...)``, stored as ``_xlfn.PERCENTILE.INC``): between
    the cells at ``(n - 1) * k`` of the sorted numbers; ``#NUM!`` for a
    ``k`` outside [0, 1] or no number.

    Example::

        q1 = mo.PERCENTILE_INC([c.a.multiple, c.b.multiple, c.c.multiple], 0.25)
        # =_xlfn.PERCENTILE.INC(Comps!C3:C5, 0.25)
        mo.PERCENTILE_INC([p.revenue for p in peers], 0.25)
        # each year: =_xlfn.PERCENTILE.INC(Peers!C3:C5, 0.25)
    """
    where = "PERCENTILE_INC"
    value, rng, nodes, texts, place = _one_range(where, array, 0, [("k", k)], _inclusive)
    return _result(where, nodes, texts, value, _type_of(value, 'float'), rng.unit, place)


def PERCENTILE_EXC(array: Any, k: Any) -> Variable:
    """The ``k``-th percentile of the cells, ``k`` inside (0, 1), the ends
    left out (renders ``=PERCENTILE.EXC(...)``, stored as
    ``_xlfn.PERCENTILE.EXC``): between the cells at ``(n + 1) * k``, which
    must lie between 1 and n; else ``#NUM!``."""
    where = "PERCENTILE_EXC"
    value, rng, nodes, texts, place = _one_range(where, array, 0, [("k", k)], _exclusive)
    return _result(where, nodes, texts, value, _type_of(value, 'float'), rng.unit, place)


def QUARTILE_INC(array: Any, quart: Any) -> Variable:
    """A quartile of the cells (renders ``=QUARTILE.INC(...)``, stored as
    ``_xlfn.QUARTILE.INC``): ``quart`` 0 the smallest, 1 the first
    quartile, 2 the median, 3 the third, 4 the largest — the inclusive
    percentile at quart / 4; a fraction is truncated; ``#NUM!`` outside
    0 to 4."""
    where = "QUARTILE_INC"
    value, rng, nodes, texts, place = _one_range(
        where, array, 0, [("quart", quart)],
        lambda ranked, q: _quartile(ranked, q, inclusive=True))
    return _result(where, nodes, texts, value, _type_of(value, 'float'), rng.unit, place)


def QUARTILE_EXC(array: Any, quart: Any) -> Variable:
    """A quartile of the cells, the ends left out (renders
    ``=QUARTILE.EXC(...)``, stored as ``_xlfn.QUARTILE.EXC``): ``quart`` 1
    to 3, the exclusive percentile at quart / 4; a fraction is truncated;
    ``#NUM!`` outside 1 to 3 or where too few cells hold that
    percentile."""
    where = "QUARTILE_EXC"
    value, rng, nodes, texts, place = _one_range(
        where, array, 0, [("quart", quart)],
        lambda ranked, q: _quartile(ranked, q, inclusive=False))
    return _result(where, nodes, texts, value, _type_of(value, 'float'), rng.unit, place)


def PERCENTILE(array: Any, k: Any) -> Variable:
    """:func:`PERCENTILE_INC` under the name an older workbook carries
    (renders ``=PERCENTILE(...)``)."""
    where = "PERCENTILE"
    value, rng, nodes, texts, place = _one_range(where, array, 0, [("k", k)], _inclusive)
    return _result(where, nodes, texts, value, _type_of(value, 'float'), rng.unit, place)


def QUARTILE(array: Any, quart: Any) -> Variable:
    """:func:`QUARTILE_INC` under the name an older workbook carries
    (renders ``=QUARTILE(...)``)."""
    where = "QUARTILE"
    value, rng, nodes, texts, place = _one_range(
        where, array, 0, [("quart", quart)],
        lambda ranked, q: _quartile(ranked, q, inclusive=True))
    return _result(where, nodes, texts, value, _type_of(value, 'float'), rng.unit, place)


def LARGE(array: Any, k: Any) -> Variable:
    """The ``k``-th largest cell (renders ``=LARGE(...)``): ``k`` 1 is the
    largest; a fractional ``k`` is taken up; ``#NUM!`` for ``k`` below 1
    or past the count. The cell comes back as it holds it — a date stays
    a date.

    Example::

        runner_up = mo.LARGE(revenue, 2)       # =LARGE(B2:M2, 2)
    """
    where = "LARGE"
    value, rng, nodes, texts, place = _one_range(
        where, array, 0, [("k", k)], lambda ranked, n: _kth(ranked, n, largest=True))
    return _result(where, nodes, texts, value, _type_of(value, _ranked_type(rng)), rng.unit,
                   place)


def SMALL(array: Any, k: Any) -> Variable:
    """The ``k``-th smallest cell (renders ``=SMALL(...)``): ``k`` 1 is the
    smallest; a fractional ``k`` is taken down; ``#NUM!`` for ``k`` below
    1 or past the count."""
    where = "SMALL"
    value, rng, nodes, texts, place = _one_range(
        where, array, 0, [("k", k)], lambda ranked, n: _kth(ranked, n, largest=False))
    return _result(where, nodes, texts, value, _type_of(value, _ranked_type(rng)), rng.unit,
                   place)


# The arguments are read as LibreOffice reads them: the order, the range,
# then the number.
def _ranking(average: bool) -> Callable[..., Any]:
    def rank(ranked: list[Any] | str, number: Any, order: Any = 0) -> Any:
        if isinstance(order, str):
            return order
        if isinstance(ranked, str):
            return ranked
        if not ranked:
            return '#N/A'
        if isinstance(number, str):
            return number
        keys = (ranked.serials() if isinstance(ranked, _Ranked)
                else [_to_serial(c) for c in ranked])
        below = bisect.bisect_left(keys, number)
        equal = bisect.bisect_right(keys, number) - below
        if not equal:
            return '#N/A'
        first = (below if order != 0 else len(keys) - below - equal) + 1
        return first + (equal - 1) / 2 if average else first
    return rank


def _rank(where: str, number: Any, ref: Any, order: Any, average: bool) -> Variable:
    params: list[tuple[str, Any]] = [("number", number)]
    if order is not _OMITTED:
        params.append(("order", order))
    value, _rng, nodes, texts, place = _one_range(
        where, ref, 1, params, _ranking(average), (number,))
    return _result(where, nodes, texts, value, 'float' if average else 'int', None, place)


def RANK_EQ(number: Any, ref: Any, order: Any = _OMITTED) -> Variable:
    """The rank of ``number`` among the cells of ``ref`` (renders
    ``=RANK.EQ(...)``, stored as ``_xlfn.RANK.EQ``): the largest is 1
    unless ``order`` is non-zero, when the smallest is; ties share the
    best rank; ``#N/A`` when ``ref`` lacks the number.

    Example::

        for c in peers:                      # one call per record, one range
            c.rank = mo.RANK_EQ(c.multiple, [p.multiple for p in peers])
        # =_xlfn.RANK.EQ(C3, C3:C7), =_xlfn.RANK.EQ(C4, C3:C7), …
        # A yearly line among its peers' — each year's rank:
        c.rank = mo.RANK_EQ(c.revenue, [p.revenue for p in peers])
        # =_xlfn.RANK.EQ(C3, C3:C7), =_xlfn.RANK.EQ(D3, D3:D7), …
    """
    return _rank("RANK_EQ", number, ref, order, False)


def RANK_AVG(number: Any, ref: Any, order: Any = _OMITTED) -> Variable:
    """The rank of ``number`` among the cells of ``ref`` (renders
    ``=RANK.AVG(...)``, stored as ``_xlfn.RANK.AVG``), ties sharing the
    mean of their ranks: 5, 5, 3 rank 1.5, 1.5, 3."""
    return _rank("RANK_AVG", number, ref, order, True)


def RANK(number: Any, ref: Any, order: Any = _OMITTED) -> Variable:
    """:func:`RANK_EQ` under the name an older workbook carries (renders
    ``=RANK(...)``)."""
    return _rank("RANK", number, ref, order, False)


# The unit a call gives, the law a loop container derives its rows' units
# by again: the values a reduction reads (a list's lines each) share one
# unit, the call's — a variance's its square; the values a range function
# compares or picks among (RANK's number and its range) share one, the
# call's — a rank's none.
def _values_law(where: str, square: bool = False) -> UnitLaw:
    def law(args: Sequence[Any], unit_of: Callable[[Any], Any]) -> Any:
        unit = same_unit(where, [n for a in args for n in _unit_items(a, unit_of)], unit_of)
        return _squared(unit) if square else unit
    return law


def _range_law(where: str, at: int) -> UnitLaw:
    def law(args: Sequence[Any], unit_of: Callable[[Any], Any]) -> Any:
        nodes = [*args[:at], *_unit_items(args[at], unit_of)] if at < len(args) else list(args)
        unit = same_unit(where, nodes, unit_of)
        return None if at else unit
    return law


for _name in ("MEDIAN", "STDEV_S", "STDEV_P", "STDEV", "STDEVP"):
    unit_law(_name, _values_law(_name))
for _name in ("VAR_S", "VAR_P", "VAR", "VARP"):
    unit_law(_name, _values_law(_name, square=True))
for _name in ("PERCENTILE_INC", "PERCENTILE_EXC", "QUARTILE_INC", "QUARTILE_EXC", "PERCENTILE",
              "QUARTILE", "LARGE", "SMALL"):
    unit_law(_name, _range_law(_name, 0))
for _name in ("RANK_EQ", "RANK_AVG", "RANK"):
    unit_law(_name, _range_law(_name, 1))

# Two values or more, a series among them — a reduction's operands, the
# lines a range lists — make a call one per period.
register("MEDIAN", shape="reduction", per_period_arity=2)
register("STDEV_S", shape="reduction", excel="STDEV.S", file_prefix="_xlfn.", per_period_arity=2)
register("STDEV_P", shape="reduction", excel="STDEV.P", file_prefix="_xlfn.", per_period_arity=2)
register("VAR_S", shape="reduction", excel="VAR.S", file_prefix="_xlfn.", per_period_arity=2)
register("VAR_P", shape="reduction", excel="VAR.P", file_prefix="_xlfn.", per_period_arity=2)
register("STDEV", shape="reduction", per_period_arity=2)
register("STDEVP", shape="reduction", per_period_arity=2)
register("VAR", shape="reduction", per_period_arity=2)
register("VARP", shape="reduction", per_period_arity=2)
register("PERCENTILE_INC", shape="range", range_args=(0,), excel="PERCENTILE.INC",
         file_prefix="_xlfn.", per_period_arity=2)
register("PERCENTILE_EXC", shape="range", range_args=(0,), excel="PERCENTILE.EXC",
         file_prefix="_xlfn.", per_period_arity=2)
register("QUARTILE_INC", shape="range", range_args=(0,), excel="QUARTILE.INC",
         file_prefix="_xlfn.", per_period_arity=2)
register("QUARTILE_EXC", shape="range", range_args=(0,), excel="QUARTILE.EXC",
         file_prefix="_xlfn.", per_period_arity=2)
register("PERCENTILE", shape="range", range_args=(0,), per_period_arity=2)
register("QUARTILE", shape="range", range_args=(0,), per_period_arity=2)
register("LARGE", shape="range", range_args=(0,), per_period_arity=2)
register("SMALL", shape="range", range_args=(0,), per_period_arity=2)
register("RANK_EQ", shape="range", range_args=(1,), excel="RANK.EQ", file_prefix="_xlfn.",
         per_period_arity=2)
register("RANK_AVG", shape="range", range_args=(1,), excel="RANK.AVG", file_prefix="_xlfn.",
         per_period_arity=2)
register("RANK", shape="range", range_args=(1,), per_period_arity=2)
