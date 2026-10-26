# SPDX-License-Identifier: Apache-2.0
"""Lookup and reference — ``MATCH``, ``INDEX``, ``XLOOKUP``, ``VLOOKUP``,
``HLOOKUP``.

Each finds a cell in a range or a block and is written as the
spreadsheet writes it: ``=INDEX(Inputs!B3:D3, MATCH(C5, Inputs!B2:D2,
0))``, ``=_xlfn.XLOOKUP(C5, Inputs!B2:D2, Inputs!B3:D3)``,
``=VLOOKUP("Beta", Deals!A3:C5, 3, FALSE)``. A range argument is what
the conditional aggregates take (:mod:`criteria`): a row, written as
the range of its cells, or a list of scalar lines whose cells stand
one under another, written as a vertical range. A block argument is
two-dimensional: a container of records laid out as a table
(``orient='records'``) — one row per record, the record's label first
and its scalar fields after it, in the order the table has them,
``Deals!A3:C5`` — or a list of rows laid out one under another
(``[years, inflation]`` → ``Inputs!B2:D3``). Cells that form no
rectangle (records laid out as sections, a table with a unit column,
rows with a totals column between their cells) cannot be one range:
the cell carries the computed value and ``check_compat`` lists it.

The key may be a series: the result is a series, one lookup per
period, and the cell per period is the call over that period's key
cell and the whole range. A position (``INDEX``'s, ``VLOOKUP``'s
column) may be a series too, and so may ``XLOOKUP``'s ``if_not_found``;
series on other time windows are read by date, as ``a + b`` reads them
(each continued by its own ``extend=`` rule, the result on the
combined window), and lines along keys key by key. A mode
(``match_type``, ``range_lookup``, ``match_mode``, ``search_mode``) is
one value, typed in.

The edges, as the spreadsheet reads them. Text matches case-
insensitively, the whole cell, with the wildcards ``*`` and ``?``
(``~`` taking the next character literally) only where Excel reads
them: an exact ``MATCH`` (``match_type`` 0), an exact ``VLOOKUP`` /
``HLOOKUP`` (``range_lookup`` FALSE) and ``XLOOKUP`` with
``match_mode`` 2. A number never matches text (``'10'`` is not 10
here, unlike in a ``SUMIF`` criterion), a boolean matches a boolean only, a date
is its serial. A blank cell matches nothing, and a key read from a
blank cell finds nothing (``#N/A``); ``None`` typed as the key is
refused — it has no cell, and the formula would write it as 0. A key
that is an error marker gives that error; an error among the cells
searched is skipped; the cell found carries its own content — an
error marker stays one, a blank reads as 0, as a formula on a blank
cell does. The approximate modes — ``MATCH`` with ``match_type`` 1 or
-1, ``VLOOKUP`` / ``HLOOKUP`` with ``range_lookup`` TRUE (the default,
as in Excel) — need the cells sorted, ascending for 1 / TRUE and
descending for -1, as Excel requires: the engine takes the last cell
not above (not below) the key, which is Excel's answer over sorted
cells and nobody's over unsorted ones. ``XLOOKUP``'s ``match_mode``
-1 / 1 need no order: the nearest cell below / above the key, the
first among equals (``search_mode`` -1: the last); its binary searches,
``search_mode`` 2 and -2, are refused — the engine reads every cell.

A position is read as the spreadsheet reads a number: a boolean is 1
or 0, a date its serial, text that reads as a number that number, a
fraction is truncated. A position 0 — the spreadsheet's whole row or
column — is ``#VALUE!``, as is a negative one, a blank or other text;
past the range, ``#REF!``. One case differs: in a line of periods,
for a 0 read from a cell the sheet takes the range's cell in the
period's own column (it intersects the whole row with the column),
where the engine gives ``#VALUE!``. ``VLOOKUP`` and ``HLOOKUP`` check
the column or row number before they search, so its error wins over
the key's: past the block is ``#REF!`` whether or not the key is
found.

The result is typed by the line the cell comes from and carries its
unit — never the keys' — and ``MATCH``, a position, has none; a result
that can come from cells of different lines is typed by them when
they agree, else by the value found, and carries a unit only when
every one of them has it. Cells it can come from in two units are
refused, as ``CHOOSE`` refuses them, and so is an ``if_not_found`` in
another unit than the cells returned; the key is compared with the
cells searched as ``==`` compares two lines, so a key in another unit
is refused too. A number or a line without a unit goes with any.
Empty text typed into a line is a blank cell, as the workbook holds it.
``XLOOKUP`` is in Excel 2021 and Microsoft 365, not in Excel 2019: the
file stores it as
``_xlfn.XLOOKUP``; a mode given without ``if_not_found`` is written
after ``NA()``, Excel's own not-found value, as the argument before
it.

``XLOOKUP(key, line)`` over a line that runs along keys of its own
(``mo.Variable({2026: 0.05, 2027: 0.04})``) looks the key up among the
line's keys: the keys are typed into the call (``XLOOKUP(year, [2026,
2027], inflation)``) and have no cells the workbook can name, so the
cell carries the computed value.

LibreOffice departs from Excel in three places, and the engine keeps
Excel's rule: there a boolean is a number (``MATCH(TRUE, …)`` finds a
1), text that reads as a number matches the number (``MATCH("2027",
years, 0)`` finds 2027), and a column or row number past the block is
``#VALUE!``, not ``#REF!``.

Not covered: a block with a record that lacks a field (the table
leaves the cell empty; no cell can be named for it), ``INDEX`` over a
one-dimensional range with two positions, a key with tracks (slice a
coordinate first), and ``MATCH`` or ``XLOOKUP`` over a block (one-
dimensional in Excel too).
"""

from __future__ import annotations

import math
from dataclasses import dataclass
from datetime import date, timedelta
from typing import Any, Callable, Sequence

from ..core.expr import Expr, FuncCall, ListExpr, Literal, VarRef
from ..core.multi_variable import MultiVariableBase
from ..core.shape import keys_along, reject_axised
from ..core.unit import unit_key
from ..core.variable import Variable
from ..core.variable_ops import _to_serial
from ._helpers import (
    _per_period, align_keyed, align_located, is_error, make_func_var, placed, src,
)
from .conditional import _blank0, _value_type_of
from .criteria import (
    _Range, _as_number, _cells, _checked, _index, _label, _range, _same_length, _wildcard,
)
from .registry import register

#: An optional argument the call did not give — left out of the formula too.
_OMITTED: Any = object()

# ─── One cell against the key ───────────────────────────────────────


def _kind(value: Any) -> str | None:
    """What a cell can be matched as: ``'number'`` (a date and a
    duration among them), ``'text'``, ``'bool'``; None for a blank or an
    error marker, which match nothing."""
    if isinstance(value, bool):
        return 'bool'
    if isinstance(value, (int, float, date, timedelta)):
        return 'number'
    if isinstance(value, str):
        return None if is_error(value) else 'text'
    return None


def _ordered(value: Any) -> Any:
    """The value as it compares within its kind: text case-insensitively,
    a date as its serial."""
    if isinstance(value, str):
        return value.casefold()
    if isinstance(value, (date, timedelta)):
        return _to_serial(value)
    return value


def _exact_test(key: Any, wildcard: bool) -> Callable[[Any], bool]:
    kind = _kind(key)
    if kind is None:
        return lambda v: False
    if kind == 'text':
        if wildcard:
            pattern = _wildcard(key)
            return lambda v: _kind(v) == 'text' and pattern.fullmatch(v) is not None
        folded = key.casefold()
        return lambda v: _kind(v) == 'text' and v.casefold() == folded
    target = _ordered(key)
    return lambda v: _kind(v) == kind and _ordered(v) == target


def _exact(cells: list[Any], key: Any, *, wildcard: bool, last: bool = False) -> int | None:
    """The position of the cell equal to the key — the first, or the
    last — or None."""
    test = _exact_test(key, wildcard)
    hits = [i for i, v in enumerate(cells) if test(v)]
    if not hits:
        return None
    return hits[-1] if last else hits[0]


# ``_exact`` over a range read once: the cells its index finds for the key,
# each read by the same test; every cell where the index finds none, or for
# a pattern, which no index answers.
def _exact_in(rng: _Range, key: Any, *, wildcard: bool, last: bool = False) -> int | None:
    kind = _kind(key)
    if kind is None:
        return None
    pattern = wildcard and kind == 'text' and any(ch in '*?~' for ch in key)
    index = None if pattern else _index(rng)
    near: list[int] | None = None
    if index is not None:
        if kind == 'bool':
            near = index.bools[key]
        elif kind == 'number':
            near = index.exact.get(_ordered(key), [])
        elif wildcard:
            near = index.like(key)
        else:
            near = index.folded.get(key.casefold(), [])
    if near is None:
        return _exact(rng.values, key, wildcard=wildcard, last=last)
    test = _exact_test(key, wildcard)
    hits = [i for i in sorted(set(near)) if test(rng.values[i])]
    if not hits:
        return None
    return hits[-1] if last else hits[0]


def _nearest(cells: list[Any], key: Any, *, below: bool, last: bool = False) -> int | None:
    """The position of the cell nearest the key on one side of it — the
    largest not above it, or the smallest not below it — among the
    cells of the key's kind, the first among equals or the last; None
    when no cell is on that side."""
    kind = _kind(key)
    target = _ordered(key)
    best: Any = None
    best_at: int | None = None
    for i, v in enumerate(cells):
        if kind is None or _kind(v) != kind:
            continue
        o = _ordered(v)
        if (o > target) if below else (o < target):
            continue
        closer = best is None or ((o > best) if below else (o < best))
        if closer or (o == best and last):
            best, best_at = o, i
    return best_at


def _sorted_match(cells: list[Any], key: Any, *, ascending: bool) -> int | None:
    """``MATCH`` 1 / -1 and the approximate ``VLOOKUP`` / ``HLOOKUP``: over
    cells sorted ascending (descending), the last cell not above (not
    below) the key, among the cells of the key's kind; None when the
    first such cell is already past it."""
    kind = _kind(key)
    target = _ordered(key)
    found = None
    for i, v in enumerate(cells):
        if kind is None or _kind(v) != kind:
            continue
        o = _ordered(v)
        if (o <= target) if ascending else (o >= target):
            found = i
    return found


def _position(n: Any, size: int) -> int | str:
    """A 1-based position typed or read from a cell as a 0-based index,
    read as the spreadsheet reads a number: an error marker stays one; a
    boolean is 1 or 0, a date its serial, text that reads as a number
    that number; 0, a negative number, a blank or other text is
    ``#VALUE!``; past ``size`` is ``#REF!``; a fraction is truncated."""
    if is_error(n):
        return n
    if isinstance(n, bool):
        n = int(n)
    elif isinstance(n, (date, timedelta)):
        n = _to_serial(n)
    elif isinstance(n, str):
        n = _as_number(n)
    if not isinstance(n, (int, float)) or (isinstance(n, float) and not math.isfinite(n)):
        return '#VALUE!'
    k = int(n)
    if k < 1:
        return '#VALUE!'
    if k > size:
        return '#REF!'
    return k - 1


# ─── The arguments ──────────────────────────────────────────────────


@dataclass
class _Block:
    """A two-dimensional range: the cells' values by row, the line each
    cell comes from (None for a record's label or a typed value), the
    AST node and the name in the source. ``labelled`` when the first
    column holds the records' labels."""

    rows: list[list[Any]]
    lines: list[list[Variable | None]]
    node: Expr
    text: str
    labelled: bool = False

    @property
    def height(self) -> int:
        return len(self.rows)

    @property
    def width(self) -> int:
        return len(self.rows[0])


def _series_value(raw: Any) -> bool:
    return isinstance(raw, list) or hasattr(raw, 'roles')


def _refuse_tracks(var: Variable, where: str) -> None:
    if hasattr(var._value, 'roles'):
        raise ValueError(
            f"{where}(): {src(var)} has tracks — slice a coordinate first "
            f"(.at(track='...'))."
        )


def _is_block(arg: Any) -> bool:
    """A container of records, or a list with a row among its items
    (list-valued lines, or typed lists), is a block; a line or a list of
    scalar lines is a range."""
    if isinstance(arg, MultiVariableBase):
        return True
    if isinstance(arg, (list, tuple)) and arg:
        return any(
            isinstance(item, list)
            or (isinstance(item, Variable) and _series_value(item._value))
            for item in arg
        )
    return False


def _records_block(mv: MultiVariableBase, where: str) -> _Block:
    """The table a container of records is laid out as: one row per
    record — its label, then its scalar fields in first-appearance order
    across the records, a field a record lacks left blank; a per-period
    field is a block of its own below the table, not a column of it."""
    from ..compile.excel.view import ExcelView
    records = [
        rec for rec in mv._components.values()
        if isinstance(rec, MultiVariableBase) and not isinstance(rec, ExcelView)
    ]
    name = getattr(mv, 'python_name', None) or mv.display_name
    if not records:
        raise ValueError(
            f"{where}(): {name} holds no records — a block is a container of "
            f"records laid out as a table (orient='records'), or a list of rows."
        )
    fields: list[str] = []
    seen: set[str] = set()
    for rec in records:
        for fname, f in rec._components.items():
            if not isinstance(f, Variable) or fname in seen:
                continue
            seen.add(fname)
            if not _series_value(f._value):
                fields.append(fname)
    if not fields:
        raise ValueError(
            f"{where}(): the records of {name} have no scalar fields — a "
            f"per-period field is a block of its own below the table."
        )
    rows: list[list[Any]] = []
    lines: list[list[Variable | None]] = []
    items: list[Expr] = []
    for rec in records:
        row: list[Any] = [rec.display_name]
        row_lines: list[Variable | None] = [None]
        nodes: list[Expr] = [Literal(rec.display_name)]
        for fname in fields:
            f = rec._components.get(fname)
            if isinstance(f, Variable):
                reject_axised(f, where)
                row.append(_cells(f, f._value)[0])
                row_lines.append(f)
                nodes.append(VarRef(f))
            else:
                row.append(None)
                row_lines.append(None)
                nodes.append(Literal(None))
        rows.append(row)
        lines.append(row_lines)
        items.append(ListExpr(nodes))
    return _Block(rows, lines, ListExpr(items), name, labelled=True)


def _rows_block(arg: Sequence[Any], where: str) -> _Block:
    """A list of rows — list-valued lines, or typed lists — of one length."""
    rows: list[list[Any]] = []
    lines: list[list[Variable | None]] = []
    items: list[Expr] = []
    labels: list[str] = []
    for item in arg:
        if isinstance(item, Variable):
            reject_axised(item, where)
            _refuse_tracks(item, where)
            if not isinstance(item._value, list):
                raise ValueError(
                    f"{where}(): a block is a list of rows; {_label(item)} is one "
                    f"cell. A list of scalar lines is a range — one position "
                    f"addresses it (mo.INDEX(lines, k))."
                )
            cells = _cells(item, item._value)
            rows.append(cells)
            lines.append([item] * len(cells))
            items.append(VarRef(item))
            labels.append(_label(item))
        elif isinstance(item, list):
            rows.append(list(item))
            lines.append([None] * len(item))
            items.append(Literal(list(item)))
            labels.append(repr(item))
        else:
            raise TypeError(
                f"{where}(): a block is a list of rows (list-valued lines); got "
                f"{type(item).__name__} among them."
            )
    widths = {len(row) for row in rows}
    if len(widths) != 1:
        raise ValueError(
            f"{where}(): the rows of a block have the same number of cells; got "
            f"{', '.join(f'{label} with {len(row)}' for label, row in zip(labels, rows))}."
        )
    return _Block(rows, lines, ListExpr(items), f"[{', '.join(labels)}]")


def _block(arg: Any, where: str) -> _Block:
    if isinstance(arg, MultiVariableBase):
        return _records_block(arg, where)
    if isinstance(arg, (list, tuple)) and arg:
        return _rows_block(arg, where)
    raise TypeError(
        f"{where}(): a block is a container of records laid out as a table, or a "
        f"list of rows; got {type(arg).__name__}."
    )


def _key(value: Any, where: str) -> tuple[Any, Expr]:
    """The key, or a position: a scalar line, a series, or a typed value."""
    if isinstance(value, Variable):
        reject_axised(value, where)
        _refuse_tracks(value, where)
        return value._value, VarRef(value)
    if isinstance(value, (dict, set, tuple, MultiVariableBase)):
        raise TypeError(
            f"{where}(): a key is one value, or a line (one lookup per period); "
            f"got {type(value).__name__}."
        )
    return value, Literal(value)


def _lookup_value(value: Any, where: str) -> tuple[Any, Expr]:
    """The value looked up — never None typed into the call, which has no
    cell: the formula would write it as 0 and look 0 up."""
    raw, node = _key(value, where)
    if isinstance(value, Variable):
        return raw, node
    if any(v is None for v in (raw if isinstance(raw, list) else [raw])):
        what = "is None" if raw is None else f"{src(value)} holds None"
        raise ValueError(
            f"{where}(): the key {what}, and None has no cell — the formula would "
            f"write it as 0 and look 0 up. Type the value to look up, or pass a "
            f"line whose cells may be left blank: a blank cell finds nothing (#N/A)."
        )
    return raw, node


def _mode(value: Any, where: str, name: str, allowed: tuple[int, ...], default: int) -> int:
    """A mode typed into the call, or its default when omitted."""
    if value is _OMITTED:
        return default
    if isinstance(value, bool) or not isinstance(value, int) or value not in allowed:
        raise ValueError(
            f"{where}(): {name} is one of {', '.join(str(a) for a in allowed)}, typed "
            f"in; got {value!r}."
        )
    return value


def _flag(value: Any, where: str) -> bool:
    """``range_lookup``: TRUE (approximate) unless given as FALSE / 0."""
    if value is _OMITTED:
        return True
    if isinstance(value, (bool, int)):
        return bool(value)
    raise ValueError(
        f"{where}(): range_lookup is TRUE or FALSE, typed in; got {value!r}."
    )


def _span(n: Any, size: int) -> list[int]:
    """The indexes a position can address: the one it holds, or every one
    for a series, a position not given (None) or one the range lacks."""
    at = None if isinstance(n, list) else _position(n, size)
    return [at] if isinstance(at, int) else list(range(size))


def _type_of(value: Any) -> str:
    """The value type a cell's content gives; a blank or an error marker
    reads as a number."""
    if isinstance(value, date):
        return 'datetime'
    if value is None or is_error(value):
        return 'float'
    return _value_type_of(value)


def _picked(block: _Block, r: Any, c: Any, value: Any) -> tuple[str, Any]:
    """The value type and unit of a block pick: those of the cells it can
    come from — the one cell when both positions are typed in, a column's
    or a row's when one is, every cell otherwise — when the cells share
    them (a record's label is text, with no unit; a field a record lacks
    counts for neither); else the type of the value found, a series as a
    number, and no unit."""
    types: set[str] = set()
    units: set[Any] = set()
    unit: Any = None
    for i in _span(r, block.height):
        for j in _span(c, block.width):
            line, cell = block.lines[i][j], block.rows[i][j]
            if line is None and cell is None:
                continue
            if line is None:
                types.add('string' if block.labelled and j == 0 else _type_of(cell))
                units.add(None)
                continue
            types.add(line.value_type)
            units.add(None if line._unit is None else unit_key(line._unit))
            unit = line._unit if unit is None else unit
    if len(types) == 1:
        value_type = types.pop()
    else:
        value_type = 'float' if isinstance(value, list) else _type_of(value)
    return value_type, unit if len(units) == 1 else None


def _read(where: str, lines: Sequence[Any],
          raws: list[Any]) -> tuple[list[Any], Any, list[Any] | None]:
    """The per-period operands' values read as ``a + b`` reads them —
    series on other time windows by date, lines along keys key by key —
    with the combined window and the keys the result runs on."""
    raws, keys = align_keyed(where, lines, raws)
    raws, window = align_located(lines, raws)
    return raws, window, keys


def _result(func: str, nodes: list[Expr], texts: list[str], value: Any,
            value_type: str, unit: Any) -> Variable:
    """The cell with its formula: a series when the key or a position is
    one; the unit is the returned range's, never the keys' (a position
    has none)."""
    with _checked(nodes):
        out = make_func_var(
            func, nodes, value, value_type, 'list' if isinstance(value, list) else 'scalar',
            source_code=f"{func}({', '.join(texts)})",
        )
    out._unit = unit
    return out


# ─── The functions ──────────────────────────────────────────────────


def MATCH(lookup_value: Any, lookup_range: Any, match_type: Any = _OMITTED) -> Variable:
    """The 1-based position of ``lookup_value`` in ``lookup_range``
    (renders ``=MATCH(...)``): ``match_type`` 0 is the first exact match
    (text with wildcards), 1 (the default) the last cell not above the
    key over cells sorted ascending, -1 the last cell not below it over
    cells sorted descending; ``#N/A`` when none.

    Example::

        at = mo.MATCH(year, inputs.years, 0)          # =MATCH(C5, Inputs!B2:D2, 0)
        rate = mo.INDEX(inputs.inflation, at)        # =INDEX(Inputs!B3:D3, MATCH(...))
    """
    where = "MATCH"
    key_raw, key_node = _lookup_value(lookup_value, where)
    rng = _range(lookup_range, where)
    mode = _mode(match_type, where, "match_type", (1, 0, -1), 1)

    def one(values: list[Any]) -> Any:
        key = values[0]
        if is_error(key):
            return key
        if key is None:
            return '#N/A'
        if mode == 0:
            at = _exact_in(rng, key, wildcard=True)
        else:
            at = _sorted_match(rng.values, key, ascending=mode == 1)
        return '#N/A' if at is None else at + 1

    nodes = [key_node, rng.node]
    texts = [src(lookup_value), rng.text]
    if match_type is not _OMITTED:
        nodes.append(Literal(mode))
        texts.append(str(mode))
    (key_raw,), window, keys = _read(where, [lookup_value], [key_raw])
    return placed(_result(where, nodes, texts, _per_period([key_raw], one), 'int', None),
                  window, keys)


def INDEX(range_or_block: Any, row_num: Any, column_num: Any = _OMITTED) -> Variable:
    """The cell at a 1-based position of a range, or at a row and a
    column of a block (renders ``=INDEX(...)``): ``mo.INDEX(rates, 2)``,
    ``mo.INDEX(deals, 2, 3)`` — the third column of the second record,
    the label column first; ``mo.INDEX([years, rates], 2, k)``."""
    where = "INDEX"
    if _is_block(range_or_block):
        if column_num is _OMITTED:
            raise ValueError(
                f"{where}() over a block takes a row and a column number: "
                f"mo.INDEX(table, 2, 3)."
            )
        block = _block(range_or_block, where)
        r_raw, r_node = _key(row_num, where)
        c_raw, c_node = _key(column_num, where)

        def pick(values: list[Any]) -> Any:
            r, c = values
            at = _position(r, block.height)
            if isinstance(at, str):
                return at
            col = _position(c, block.width)
            if isinstance(col, str):
                return col
            return _blank0(block.rows[at][col])

        (r_raw, c_raw), window, keys = _read(where, [row_num, column_num], [r_raw, c_raw])
        value = _per_period([r_raw, c_raw], pick)
        value_type, unit = _picked(block, r_raw, c_raw, value)
        return placed(_result(where, [block.node, r_node, c_node],
                              [block.text, src(row_num), src(column_num)], value, value_type,
                              unit), window, keys)
    if column_num is not _OMITTED:
        raise ValueError(
            f"{where}() over a range takes one position (mo.INDEX(row, k)); a row "
            f"and a column number address a block — a table of records, or a "
            f"list of rows."
        )
    rng = _range(range_or_block, where)
    k_raw, k_node = _key(row_num, where)

    def one(values: list[Any]) -> Any:
        at = _position(values[0], len(rng.values))
        return at if isinstance(at, str) else _blank0(rng.values[at])

    (k_raw,), window, keys = _read(where, [row_num], [k_raw])
    return placed(_result(where, [rng.node, k_node], [rng.text, src(row_num)],
                          _per_period([k_raw], one), rng.value_type or 'float', rng.unit),
                  window, keys)


def XLOOKUP(lookup_value: Any, lookup_range: Any, return_range: Any = None,
            if_not_found: Any = _OMITTED, match_mode: Any = _OMITTED,
            search_mode: Any = _OMITTED) -> Variable:
    """The cell of ``return_range`` at the position of ``lookup_value``
    in ``lookup_range`` (renders ``=XLOOKUP(...)``, stored as
    ``_xlfn.XLOOKUP``): an exact match unless ``match_mode`` is -1 (or
    the next smaller), 1 (or the next larger) or 2 (wildcards); the
    first found, or the last with ``search_mode`` -1; ``if_not_found``
    in place of ``#N/A``. Over a line that runs along keys of its own,
    ``mo.XLOOKUP(year, inflation)`` looks the key up among its keys
    (``if_not_found`` and the modes by keyword there).

    Example::

        rate = mo.XLOOKUP(year, inputs.years, inputs.inflation)
        # =_xlfn.XLOOKUP(C5, Inputs!B2:D2, Inputs!B3:D3)
    """
    where = "XLOOKUP"
    key_raw, key_node = _lookup_value(lookup_value, where)
    if return_range is None:
        keys = keys_along(lookup_range) if isinstance(lookup_range, Variable) else None
        if keys is None:
            raise ValueError(
                f"{where}() takes the key, the cells to search and the cells to "
                f"return: mo.XLOOKUP(year, years, rates). Over one line, that line "
                f"runs along keys of its own (mo.Variable({{2026: 0.05, ...}})) and "
                f"its keys are searched."
            )
        search = _Range(list(keys), Literal(list(keys)), repr(list(keys)), None, None)
        back = _Range(_cells(lookup_range, lookup_range._value), VarRef(lookup_range),
                      src(lookup_range), getattr(lookup_range, '_unit', None),
                      lookup_range.value_type)
    else:
        search = _range(lookup_range, where)
        back = _range(return_range, where)
        _same_length(where, search, back)
    mode = _mode(match_mode, where, "match_mode", (0, -1, 1, 2), 0)
    if type(search_mode) is int and abs(search_mode) == 2:
        raise ValueError(
            f"{where}(): search_mode {search_mode} is a binary search over sorted "
            f"cells, which the engine does not do — it reads every cell. Use 1 for "
            f"the first match or -1 for the last."
        )
    order = _mode(search_mode, where, "search_mode", (1, -1), 1)
    last = order < 0
    fallback = None if if_not_found is _OMITTED else _key(if_not_found, where)

    def one(values: list[Any]) -> Any:
        key = values[0]
        if is_error(key):
            return key
        at = None
        if key is not None:
            if mode in (0, 2):
                at = _exact_in(search, key, wildcard=mode == 2, last=last)
            else:
                at = _nearest(search.values, key, below=mode == -1, last=last)
        if at is None:
            return '#N/A' if fallback is None else _blank0(values[1])
        return _blank0(back.values[at])

    raws = [key_raw] if fallback is None else [key_raw, fallback[0]]
    lines = [lookup_value] if fallback is None else [lookup_value, if_not_found]
    raws, window, keys = _read(where, lines, raws)
    nodes = [key_node, search.node, back.node]
    texts = [src(lookup_value), search.text, back.text]
    if fallback is not None:
        nodes.append(fallback[1])
        texts.append(src(if_not_found))
    elif match_mode is not _OMITTED or search_mode is not _OMITTED:
        nodes.append(FuncCall('NA', []))
    if match_mode is not _OMITTED or search_mode is not _OMITTED:
        nodes.append(Literal(mode))
        texts.append(f"match_mode={mode}")
    if search_mode is not _OMITTED:
        nodes.append(Literal(order))
        texts.append(f"search_mode={order}")
    return placed(_result(where, nodes, texts, _per_period(raws, one),
                          back.value_type or 'float', back.unit), window, keys)


def VLOOKUP(lookup_value: Any, block: Any, col_index_num: Any,
            range_lookup: Any = _OMITTED) -> Variable:
    """The cell in column ``col_index_num`` of the row of ``block`` whose
    first column holds ``lookup_value`` (renders ``=VLOOKUP(...)``): over
    a table of records the first column is the record's label; an
    approximate match over a sorted first column unless ``range_lookup``
    is FALSE, as in Excel.

    Example::

        amount = mo.VLOOKUP("Beta", deals, 3, False)   # =VLOOKUP("Beta", Deals!A3:C5, 3, FALSE)
    """
    where = "VLOOKUP"
    key_raw, key_node = _lookup_value(lookup_value, where)
    table = _block(block, where)
    c_raw, c_node = _key(col_index_num, where)
    approx = _flag(range_lookup, where)
    first = [row[0] for row in table.rows]

    def one(values: list[Any]) -> Any:
        key, c = values
        # The column number is checked before the key is searched, as the
        # spreadsheet checks an argument: its error wins over the key's.
        col = _position(c, table.width)
        if isinstance(col, str):
            return col
        if is_error(key):
            return key
        if key is None:
            return '#N/A'
        at = (_sorted_match(first, key, ascending=True) if approx
              else _exact(first, key, wildcard=True))
        return '#N/A' if at is None else _blank0(table.rows[at][col])

    (key_raw, c_raw), window, keys = _read(where, [lookup_value, col_index_num],
                                           [key_raw, c_raw])
    value = _per_period([key_raw, c_raw], one)
    value_type, unit = _picked(table, None, c_raw, value)
    nodes = [key_node, table.node, c_node]
    texts = [src(lookup_value), table.text, src(col_index_num)]
    if range_lookup is not _OMITTED:
        nodes.append(Literal(approx))
        texts.append(str(approx))
    return placed(_result(where, nodes, texts, value, value_type, unit), window, keys)


def HLOOKUP(lookup_value: Any, block: Any, row_index_num: Any,
            range_lookup: Any = _OMITTED) -> Variable:
    """The cell in row ``row_index_num`` of the column of ``block`` whose
    first row holds ``lookup_value`` (renders ``=HLOOKUP(...)``): over a
    list of rows, ``mo.HLOOKUP(year, [years, inflation], 2)``; an
    approximate match over a sorted first row unless ``range_lookup``
    is FALSE, as in Excel."""
    where = "HLOOKUP"
    key_raw, key_node = _lookup_value(lookup_value, where)
    table = _block(block, where)
    r_raw, r_node = _key(row_index_num, where)
    approx = _flag(range_lookup, where)
    first = table.rows[0]

    def one(values: list[Any]) -> Any:
        key, r = values
        # The row number is checked before the key is searched, as the
        # spreadsheet checks an argument: its error wins over the key's.
        row = _position(r, table.height)
        if isinstance(row, str):
            return row
        if is_error(key):
            return key
        if key is None:
            return '#N/A'
        at = (_sorted_match(first, key, ascending=True) if approx
              else _exact(first, key, wildcard=True))
        return '#N/A' if at is None else _blank0(table.rows[row][at])

    (key_raw, r_raw), window, keys = _read(where, [lookup_value, row_index_num],
                                           [key_raw, r_raw])
    value = _per_period([key_raw, r_raw], one)
    value_type, unit = _picked(table, r_raw, None, value)
    nodes = [key_node, table.node, r_node]
    texts = [src(lookup_value), table.text, src(row_index_num)]
    if range_lookup is not _OMITTED:
        nodes.append(Literal(approx))
        texts.append(str(approx))
    return placed(_result(where, nodes, texts, value, value_type, unit), window, keys)


register("MATCH", shape="range", range_args=(1,))
register("INDEX", shape="range", range_args=(0,))
register("XLOOKUP", shape="range", range_args=(1, 2), file_prefix="_xlfn.")
register("VLOOKUP", shape="range", range_args=(1,))
register("HLOOKUP", shape="range", range_args=(1,))
# Written by the renderer as given: ``NA()``, the not-found value an
# ``XLOOKUP`` with a mode and no ``if_not_found`` is written with.
register("NA", shape="scalar", bare=True)


# ─── Units ──────────────────────────────────────────────────────────

# A lookup compares its key with the cells it searches, as ``==`` compares
# two lines, and takes one of the cells it returns, as ``CHOOSE`` takes one
# of its values: a line searched in another unit than the key is refused,
# and so are cells to return in two units, or an ``if_not_found`` in
# another unit than theirs; a number or an unlabelled line goes with any.
# Each law gives the unit the call is built with (``_result``) — the
# returned cells', never the key's; ``MATCH``, a position, has none.

from ._helpers import same_unit, unit_law  # noqa: E402 - the laws below


def _key_against(func: str, key: Any, searched: Sequence[Any],
                 unit_of: Callable[[Any], Any]) -> None:
    """The key compared with each line among the cells searched: a line in
    another unit is refused, as comparing the two lines is."""
    from ..core.unit import UNSETTLED
    from ..core.variable_ops import _incompatible_units
    from .criteria import _named

    own = unit_of(key)
    if own is None or own is UNSETTLED:
        return
    for line in searched:
        unit = unit_of(line)
        if unit is not None and unit is not UNSETTLED and unit != own:
            detail = _incompatible_units(own, unit).split(": ", 1)[1]
            raise ValueError(
                f"mo.{func} compares the key {_named(key)} with {_named(line)}; "
                f"they are in {detail}"
            )


def _rows_of(block: Any) -> list[Any]:
    """A block's rows as its node holds them: a record (its label, then its
    fields), a row line, or a list typed in."""
    return list(block.items) if isinstance(block, ListExpr) else []


def _pick_unit(func: str, block: Any, r: Any, c: Any, unit_of: Callable[[Any], Any]) -> Any:
    """The unit of a pick from a block, as ``_picked`` gives it — that of the
    cells the positions can address (``r`` / ``c`` a position's node, None
    for any) when every one of them has it, else none — with cells in two
    units refused, as ``CHOOSE`` refuses them; unless a position that is
    one value is one the block lacks, when the call is an error in every
    period and takes no cell."""
    from ..core.unit import UNSETTLED

    def held(node: Any) -> Any:
        return (node.value if isinstance(node, Literal)
                else node.var._value if isinstance(node, VarRef) else [])

    rows = _rows_of(block)
    if not rows:
        return None
    top = rows[0].items if isinstance(rows[0], ListExpr) else held(rows[0])
    spans: list[Any] = []
    erring = False
    for node, size in ((r, len(rows)), (c, len(top) if isinstance(top, list) else 1)):
        n = [] if node is None else held(node)
        spans.append(_span(n, size))
        erring = erring or not (isinstance(n, list) or isinstance(_position(n, size), int))
    lines: list[Any] = []
    seen: set[int] = set()
    bare = False                     # a label or a value typed in: no unit
    for i in spans[0]:
        for j in spans[1]:
            cell = rows[i].items[j] if isinstance(rows[i], ListExpr) else rows[i]
            if isinstance(cell, VarRef):
                if id(cell.var) not in seen:
                    seen.add(id(cell.var))
                    lines.append(cell)
            elif isinstance(cell, Literal):
                value = cell.value[j] if isinstance(cell.value, list) else cell.value
                bare = bare or value is not None
    if not erring:
        same_unit(func, lines, unit_of)
    units = [unit_of(n) for n in lines]
    if any(u is UNSETTLED for u in units):
        return UNSETTLED
    keys = {None if u is None else unit_key(u) for u in units} | ({None} if bare else set())
    return next((u for u in units if u is not None), None) if len(keys) == 1 else None


def _match_unit(args: Sequence[Any], unit_of: Callable[[Any], Any]) -> Any:
    """``MATCH`` compares the key with the lines of the range; a position
    has no unit."""
    from .criteria import _lines_of
    if len(args) >= 2:
        _key_against("MATCH", args[0], _lines_of(args[1]), unit_of)
    return None


def _index_unit(args: Sequence[Any], unit_of: Callable[[Any], Any]) -> Any:
    """``INDEX`` takes a cell of a range — the unit its lines share — or of a
    block, at the rows and columns its positions can address."""
    from .criteria import _lines_of
    if len(args) == 2:
        return same_unit("INDEX", _lines_of(args[0]), unit_of)
    if len(args) == 3:
        return _pick_unit("INDEX", args[0], args[1], args[2], unit_of)
    return None


def _xlookup_unit(args: Sequence[Any], unit_of: Callable[[Any], Any]) -> Any:
    """``XLOOKUP`` compares the key with the lines searched and takes a cell
    of the lines returned, or ``if_not_found`` in their unit: theirs."""
    from .criteria import _lines_of
    if len(args) < 3:
        return None
    _key_against("XLOOKUP", args[0], _lines_of(args[1]), unit_of)
    returned = _lines_of(args[2])
    unit = same_unit("XLOOKUP", returned, unit_of)
    if len(args) > 3:                # if_not_found, or NA() before a mode
        same_unit("XLOOKUP", [*returned, args[3]], unit_of)
    return unit


def _vlookup_unit(args: Sequence[Any], unit_of: Callable[[Any], Any]) -> Any:
    """``VLOOKUP`` compares the key with the block's first column and takes a
    cell of the column its number addresses, in any row."""
    if len(args) < 3:
        return None
    first = [row.items[0] if isinstance(row, ListExpr) else row for row in _rows_of(args[1])]
    _key_against("VLOOKUP", args[0], first, unit_of)
    return _pick_unit("VLOOKUP", args[1], None, args[2], unit_of)


def _hlookup_unit(args: Sequence[Any], unit_of: Callable[[Any], Any]) -> Any:
    """``HLOOKUP`` compares the key with the block's first row and takes a
    cell of the row its number addresses, in any column."""
    from .criteria import _lines_of
    if len(args) < 3:
        return None
    rows = _rows_of(args[1])
    if rows:
        _key_against("HLOOKUP", args[0], _lines_of(rows[0]), unit_of)
    return _pick_unit("HLOOKUP", args[1], args[2], None, unit_of)


unit_law("MATCH", _match_unit)
unit_law("INDEX", _index_unit)
unit_law("XLOOKUP", _xlookup_unit)
unit_law("VLOOKUP", _vlookup_unit)
unit_law("HLOOKUP", _hlookup_unit)
