# SPDX-License-Identifier: Apache-2.0
"""Aggregation helpers — reduce Variables to a scalar.

``SUM``, ``MAX``, ``MIN``, ``AVERAGE`` — all Excel-cased to match the cells
they compile to (``=SUM(B2:F2)``, ``=AVERAGE(B2:F2)``) and all variadic like
Excel: ``SUM(a, b, c)``, ``SUM(range)``, ``SUM(1, 2, 3)`` all work.

``MIN`` / ``MAX`` over two or more operands with a series among them are
not a reduction but one call per period — the cap and the floor a
workbook writes in every column (``=MIN(C2, C3)``); see :func:`MIN`.

**Invariant:** every call returns a :class:`Variable` with a tracked
formula — even when all operands are plain numbers. ``mo.SUM(1, 2, 3)``
produces ``=SUM(1, 2, 3)`` in Excel, not a bare ``6``. The formula is
the product.
"""

from __future__ import annotations

import builtins
from typing import Any, List, Union

from ..core.expr import Literal, VarRef
from ..core.variable import Variable
from ._helpers import make_func_var, operand, src

Number = Union[int, float]
AggOperand = Union[Variable, Number, List[Number]]


def _flatten_args(args: tuple, func_name: str) -> tuple[List[Any], list]:
    """Return (flat values, AST arg nodes) from Variables, scalars, and lists."""
    if not args:
        raise ValueError(
            f"{func_name}() needs at least one argument. "
            f"Syntax: mo.{func_name}(a, b, c)  or  mo.{func_name}(list_var). "
            f"Got 0 arguments."
        )

    flat: List[Any] = []
    ast_args: list = []
    from ..core.shape import reject_axised
    from ..core.tracks import TrackValues
    for a in args:
        if isinstance(a, Variable) and isinstance(a._value, TrackValues):
            raise ValueError(
                f"{func_name}() over a Variable with tracks reduces per "
                f"coordinate — use the lifted form on ONE operand "
                f"(handled by the caller), or slice a coordinate first "
                f"(.at(track='...'))."
            )
        if isinstance(a, Variable):
            reject_axised(a, func_name)
            flat.extend(a._value if isinstance(a._value, list) else [a._value])
            ast_args.append(VarRef(a))
        elif isinstance(a, list):
            flat.extend(a)
            ast_args.append(Literal(a))
        else:
            flat.append(a)
            ast_args.append(Literal(a))
    return flat, ast_args


def _first_variable_value_type(args: tuple, fallback: str = 'float') -> str:
    for a in args:
        if isinstance(a, Variable):
            return a.value_type
    return fallback


def _common_unit(args: tuple):
    """The single unit shared by every unit-bearing Variable operand —
    an aggregation of $ is $. Mixed or absent units → None (never
    guess)."""
    from ..core.unit import unit_key
    units = {
        unit_key(a._unit) for a in args
        if isinstance(a, Variable) and getattr(a, '_unit', None) is not None
    }
    if len(units) != 1:
        return None
    for a in args:
        if isinstance(a, Variable) and getattr(a, '_unit', None) is not None:
            return a._unit
    return None


def _result(func_name: str, args: tuple, ast_args: list, value: Any,
            value_type: str, var_type: str) -> Variable:
    source_code = f"{func_name}({', '.join(src(a) for a in args)})"
    out = make_func_var(
        func_name, ast_args, value, value_type, var_type,
        source_code=source_code,
    )
    unit = _common_unit(args)
    if unit is not None:
        out._unit = unit
    return out


def _aggregate(func_name: str, args: tuple, value: Any, value_type: str) -> Variable:
    _, ast_args = _flatten_args(args, func_name)
    return _result(func_name, args, ast_args, value, value_type, 'scalar')


def _is_series(a: Any) -> bool:
    """A period-shaped operand: a list, or a Variable holding one (a row,
    or a row per track)."""
    from ..core.tracks import TrackValues
    if isinstance(a, Variable):
        return isinstance(a._value, (list, TrackValues))
    return isinstance(a, list)


def _per_period(values: List[Any], excel: Any) -> Any:
    """``excel`` over the operands of one period at a time; a scalar or
    a one-value series stands in every period, as the operators
    broadcast it. Scalars only: one call."""
    lists = [v for v in values if isinstance(v, list)]
    if not lists:
        return excel(values)
    n = builtins.max(len(v) for v in lists)
    for v in lists:
        if len(v) not in (1, n):
            raise ValueError(
                f"Cannot broadcast lists of different lengths: {len(v)} vs {n}. "
                f"Element-wise ops need equal lengths, or one side of length 1 "
                f"which broadcasts over the other."
            )

    def at(v: Any, i: int) -> Any:
        if isinstance(v, list):
            return v[i] if len(v) == n else v[0]
        return v

    return [excel([at(v, i) for v in values]) for i in range(n)]


def _elementwise(func_name: str, args: tuple, reduce: Any) -> Variable:
    """``MIN`` / ``MAX`` over two or more operands with a series among
    them: one call per period, so the cell per period is ``=MIN(C2,
    C3)``. Tracks lift per coordinate, as under the operators."""
    from ..core.shape import reject_axised
    from ..core.tracks import TrackValues
    raws: List[Any] = []
    ast_args: list = []
    for a in args:
        if isinstance(a, Variable):
            reject_axised(a, func_name)
        raw, node = operand(a)
        raws.append(raw)
        ast_args.append(node)
    excel = _excel_reduce(reduce)
    if any(isinstance(r, TrackValues) for r in raws):
        value: Any = TrackValues.lift(lambda *fibers: _per_period(list(fibers), excel), *raws)
        var_type = 'list' if value.time_length else 'scalar'
    else:
        value = _per_period(raws, excel)
        var_type = 'list'
    return _result(func_name, args, ast_args, value,
                   _first_variable_value_type(args, 'float'), var_type)


def _lifted_reduce(func_name: str, var: Variable, reducer) -> Variable:
    """Per-coordinate reduction of ONE Variable with tracks:
    ``SUM(revenue)`` → Tracks of per-role totals. The AST stays the
    ordinary FuncCall; the value lifts."""
    from ..core.expr import VarRef
    from ..core.tracks import TrackValues
    from ._helpers import make_func_var
    lifted = TrackValues.lift(
        # A single value per track stands as itself; a blank one reduces
        # as an empty range does (0, or #DIV/0! for AVERAGE).
        lambda track: (reducer(track) if isinstance(track, list)
                       else reducer([]) if track is None else track),
        var._value,
    )
    label = getattr(var, "python_name", None) or var.path.leaf
    out = make_func_var(
        func_name, [VarRef(var)], lifted, 'float', 'scalar',
        source_code=f"{func_name}({label})",
    )
    if getattr(var, '_unit', None) is not None:
        out._unit = var._unit
    return out


def _excel_reduce(reduce: Any, empty: Any = 0) -> Any:
    """``reduce`` over one range as the spreadsheet reads it: an error in
    the range is the result — the first one, in reading order — and a hole
    (a cell not entered yet) is skipped, as a blank is. With nothing left
    the result is ``empty``: 0 for SUM, MAX and MIN, ``#DIV/0!`` for
    AVERAGE."""
    def excel(values: List[Any]) -> Any:
        for value in values:
            if isinstance(value, str) and value.startswith("#"):
                return value
        known = [value for value in values if value is not None]
        return reduce(known) if known else empty
    return excel


def _mean(values: List[Any]) -> Any:
    return builtins.sum(values) / len(values)


def SUM(*args: AggOperand) -> Variable:
    """Sum of Variables, scalars, or lists (renders ``=SUM(...)``).

    Accepts ``SUM(a, b, c)``, ``SUM(list_var)``, ``SUM([1, 2, 3])``, or
    ``SUM(1, 2, 3)``. Always returns a Variable with a tracked formula —
    the Excel cell is the product, not the value.

    Example::

        total = mo.SUM(revenues)         # =SUM(B2:F2)
        total = mo.SUM(q1, q2, q3, q4)   # =SUM(q1, q2, q3, q4)
        total = mo.SUM(1, 2, 3)          # =SUM(1, 2, 3), value = 6
    """
    from ..core.tracks import TrackValues
    if (len(args) == 1 and isinstance(args[0], Variable)
            and isinstance(args[0]._value, TrackValues)):
        return _lifted_reduce("SUM", args[0], _excel_reduce(builtins.sum))
    flat, _ = _flatten_args(args, "SUM")
    return _aggregate("SUM", args, _excel_reduce(builtins.sum)(flat),
                      _first_variable_value_type(args, 'float'))


def MAX(*args: AggOperand) -> Variable:
    """Maximum (renders ``=MAX(...)``).

    Over ONE operand, the largest value of the row: ``mo.MAX(revenue)``
    → ``=MAX(B2:F2)``. Over two or more operands with a series among
    them, the larger PER PERIOD — a floor — with scalars broadcast as
    ``a + b`` broadcasts them: ``mo.MAX(cash, inp.min_cash)`` →
    ``=MAX(C4, Inputs!B2)`` in every column, inside a loop container
    too. Over scalars only, one number: ``mo.MAX(a, b)`` → ``=MAX(B2, B3)``.
    """
    if len(args) >= 2 and any(_is_series(a) for a in args):
        return _elementwise("MAX", args, builtins.max)
    from ..core.tracks import TrackValues
    if (len(args) == 1 and isinstance(args[0], Variable)
            and isinstance(args[0]._value, TrackValues)):
        return _lifted_reduce("MAX", args[0], _excel_reduce(builtins.max))
    flat, _ = _flatten_args(args, "MAX")
    return _aggregate("MAX", args, _excel_reduce(builtins.max)(flat),
                      _first_variable_value_type(args, 'float'))


def MIN(*args: AggOperand) -> Variable:
    """Minimum (renders ``=MIN(...)``).

    Over ONE operand, the smallest value of the row: ``mo.MIN(cash)`` →
    ``=MIN(B2:F2)``. Over two or more operands with a series among them,
    the smaller PER PERIOD — a cap — with scalars broadcast as ``a + b``
    broadcasts them: ``mo.MIN(sweep, balance)`` → ``=MIN(C2, C3)`` in
    every column, inside a loop container too. Over scalars only, one
    number: ``mo.MIN(a, b)`` → ``=MIN(B2, B3)``.
    """
    if len(args) >= 2 and any(_is_series(a) for a in args):
        return _elementwise("MIN", args, builtins.min)
    from ..core.tracks import TrackValues
    if (len(args) == 1 and isinstance(args[0], Variable)
            and isinstance(args[0]._value, TrackValues)):
        return _lifted_reduce("MIN", args[0], _excel_reduce(builtins.min))
    flat, _ = _flatten_args(args, "MIN")
    return _aggregate("MIN", args, _excel_reduce(builtins.min)(flat),
                      _first_variable_value_type(args, 'float'))


def AVERAGE(*args: AggOperand) -> Variable:
    """Mean across Variables, scalars, or lists (renders ``=AVERAGE(...)``)."""
    from ..core.tracks import TrackValues
    if (len(args) == 1 and isinstance(args[0], Variable)
            and isinstance(args[0]._value, TrackValues)):
        return _lifted_reduce("AVERAGE", args[0], _excel_reduce(_mean, "#DIV/0!"))
    flat, _ = _flatten_args(args, "AVERAGE")
    value = _excel_reduce(_mean, "#DIV/0!")(flat)
    return _aggregate("AVERAGE", args, value, 'float')
