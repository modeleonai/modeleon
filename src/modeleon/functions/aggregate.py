# SPDX-License-Identifier: Apache-2.0
"""Aggregation helpers — reduce Variables to a scalar.

``SUM``, ``MAX``, ``MIN``, ``AVERAGE`` — all Excel-cased to match the cells
they compile to (``=SUM(B2:F2)``, ``=AVERAGE(B2:F2)``) and all variadic like
Excel: ``SUM(a, b, c)``, ``SUM(range)``, ``SUM(1, 2, 3)`` all work.

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
from ._helpers import make_func_var, src

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
    an aggregation of ₸ is ₸. Mixed or absent units → None (never
    guess)."""
    units = {
        str(a._unit) for a in args
        if isinstance(a, Variable) and getattr(a, '_unit', None) is not None
    }
    if len(units) != 1:
        return None
    for a in args:
        if isinstance(a, Variable) and getattr(a, '_unit', None) is not None:
            return a._unit
    return None


def _aggregate(func_name: str, args: tuple, value: Any, value_type: str) -> Variable:
    _, ast_args = _flatten_args(args, func_name)
    source_code = f"{func_name}({', '.join(src(a) for a in args)})"
    out = make_func_var(
        func_name, ast_args, value, value_type, 'scalar',
        source_code=source_code,
    )
    unit = _common_unit(args)
    if unit is not None:
        out._unit = unit
    return out


def _lifted_reduce(func_name: str, var: Variable, reducer) -> Variable:
    """Per-coordinate reduction of ONE Variable with tracks (§15.3.5):
    ``SUM(выручка)`` → Tracks of per-role totals. The AST stays the
    ordinary FuncCall; the value lifts."""
    from ..core.expr import VarRef
    from ..core.tracks import TrackValues
    from ._helpers import make_func_var
    lifted = TrackValues.lift(
        lambda track: reducer(track) if isinstance(track, list) else track,
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
        return _lifted_reduce("SUM", args[0], builtins.sum)
    flat, _ = _flatten_args(args, "SUM")
    return _aggregate("SUM", args, builtins.sum(flat),
                      _first_variable_value_type(args, 'float'))


def MAX(*args: AggOperand) -> Variable:
    """Maximum across Variables, scalars, or lists (renders ``=MAX(...)``)."""
    from ..core.tracks import TrackValues
    if (len(args) == 1 and isinstance(args[0], Variable)
            and isinstance(args[0]._value, TrackValues)):
        return _lifted_reduce("MAX", args[0], builtins.max)
    flat, _ = _flatten_args(args, "MAX")
    return _aggregate("MAX", args, builtins.max(flat),
                      _first_variable_value_type(args, 'float'))


def MIN(*args: AggOperand) -> Variable:
    """Minimum across Variables, scalars, or lists (renders ``=MIN(...)``)."""
    from ..core.tracks import TrackValues
    if (len(args) == 1 and isinstance(args[0], Variable)
            and isinstance(args[0]._value, TrackValues)):
        return _lifted_reduce("MIN", args[0], builtins.min)
    flat, _ = _flatten_args(args, "MIN")
    return _aggregate("MIN", args, builtins.min(flat),
                      _first_variable_value_type(args, 'float'))


def AVERAGE(*args: AggOperand) -> Variable:
    """Mean across Variables, scalars, or lists (renders ``=AVERAGE(...)``)."""
    from ..core.tracks import TrackValues
    if (len(args) == 1 and isinstance(args[0], Variable)
            and isinstance(args[0]._value, TrackValues)):
        return _lifted_reduce("AVERAGE", args[0], lambda f: builtins.sum(f) / len(f) if f else 0)
    flat, _ = _flatten_args(args, "AVERAGE")
    value = builtins.sum(flat) / len(flat) if flat else 0
    return _aggregate("AVERAGE", args, value, 'float')
