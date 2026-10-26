# SPDX-License-Identifier: Apache-2.0
"""Aggregation helpers — ``SUM``, ``MAX``, ``MIN``, ``AVERAGE``.

All Excel-cased to match the cells they compile to (``=SUM(B2:F2)``,
``=AVERAGE(B2:F2)``) and all variadic like Excel: ``SUM(a, b, c)``,
``SUM(range)``, ``SUM(1, 2, 3)`` all work.

One rule for every aggregate: several series are combined period by
period, as ``a + b`` combines them; one series alone is reduced over its
periods; scalars only stay one number. ``SUM(revenue)`` is the total of
the row, ``=SUM(B2:F2)``; ``SUM(product_a, product_b)`` is the two lines'
total in every period, ``=SUM(B2, B3)`` in each column — a subtotal, as
``MAX(cash, floor)`` is a floor and ``MIN(sweep, balance)`` a cap; the
total of several rows over all their periods is ``SUM(a) + SUM(b)``. A
list of scalar lines given as one argument is its lines, one by one
(``SUM([p, q])`` is ``SUM(p, q)``); a list holding a series is refused,
since it reads as either of those two — the call names both.

**Invariant:** every call returns a :class:`Variable` with a tracked
formula — even when all operands are plain numbers. ``mo.SUM(1, 2, 3)``
produces ``=SUM(1, 2, 3)`` in Excel, not a bare ``6``. The formula is
the product.
"""

from __future__ import annotations

import builtins
import functools
import operator
from typing import Any, List, Union

from ..core.expr import Literal, VarRef
from ..core.variable import Variable
from ._helpers import (
    _typed_blank,
    elementwise,
    first_variable_value_type,
    is_series,
    make_func_var,
    src,
)
from .registry import per_period_arity, register

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
            # A ``None`` typed in is the 0 the formula holds; a line's
            # blank cell (above) stays a blank the reduction skips.
            flat.extend(_typed_blank(a, 0))
            ast_args.append(Literal(a))
        else:
            flat.append(_typed_blank(a, 0))
            ast_args.append(Literal(a))
    return flat, ast_args


def _result(func_name: str, args: tuple, ast_args: list, value: Any,
            value_type: str, var_type: str) -> Variable:
    source_code = f"{func_name}({', '.join(src(a) for a in args)})"
    return make_func_var(
        func_name, ast_args, value, value_type, var_type,
        source_code=source_code, any_value_type=True,
    )


def _aggregate(func_name: str, args: tuple, value: Any, value_type: str) -> Variable:
    _, ast_args = _flatten_args(args, func_name)
    return _result(func_name, args, ast_args, value, value_type, 'scalar')


#: Each function's one number over every cell of two rows, for the
#: refusal of a list holding a series: what it is called, how it is
#: spelled (``{more}`` / ``{also}`` / ``{times}`` mark where further lines
#: go).
_OVER_ALL = {
    "SUM": ("the total", "mo.SUM({a}) + mo.SUM({b}){more}"),
    "AVERAGE": ("the mean", "(mo.SUM({a}) + mo.SUM({b}){more}) / "
                            "(mo.COUNT({a}) + mo.COUNT({b}){more})"),
    "MIN": ("the smallest", "mo.MIN(mo.MIN({a}), mo.MIN({b}){also})"),
    "MAX": ("the largest", "mo.MAX(mo.MAX({a}), mo.MAX({b}){also})"),
    "PRODUCT": ("the product", "mo.PRODUCT({a}) * mo.PRODUCT({b}){times}"),
    "COUNT": ("the count", "mo.COUNT({a}) + mo.COUNT({b}){more}"),
    "COUNTA": ("the count", "mo.COUNTA({a}) + mo.COUNTA({b}){more}"),
}


def _lists_lines(arg: Any) -> bool:
    return isinstance(arg, (list, tuple)) and any(isinstance(i, Variable) for i in arg)


def _refuse_listed_series(func_name: str, args: tuple) -> None:
    """A list of lines given as one argument with a series in it reads as
    two calls — one per period over the lines' cells, or one number over
    every cell — and Python cannot know which was meant: refused, both
    spelled out with the lines' own names."""
    for arg in args:
        if not _lists_lines(arg):
            continue
        lines = [i for i in arg if isinstance(i, Variable)]
        series = next((i for i in lines if is_series(i)), None)
        if series is None:
            continue
        names = [src(i) for i in lines]
        listed = f"[{', '.join(src(i) if isinstance(i, Variable) else repr(i) for i in arg)}]"
        noun, spelled = _OVER_ALL[func_name]
        if len(names) >= 2:
            more = len(names) > 2
            per = (f"the lines one by one, mo.{func_name}({names[0]}, {names[1]}"
                   f"{', …' if more else ''})")
            over = spelled.format(a=names[0], b=names[1], more=" + …" if more else "",
                                  also=", …" if more else "", times=" * …" if more else "")
        else:
            per = f"the line itself, {names[0]}"
            over = f"mo.{func_name}({names[0]})"
        raise ValueError(
            f"mo.{func_name}({listed}) gives its lines as one list, and {src(series)} in "
            f"it is a series, so the call reads two ways: per period, {per}; {noun} over "
            f"all their periods, {over}. Write the one you mean."
        )


def _listed(func_name: str, args: tuple) -> tuple:
    """The operands of a call given a list of scalar lines (and values) as
    one argument: its items, one by one — ``mo.SUM([p, q])`` is
    ``mo.SUM(p, q)``, written ``=SUM(B1, B2)``. A list holding a series is
    refused (``_refuse_listed_series``)."""
    _refuse_listed_series(func_name, args)
    if not any(_lists_lines(a) for a in args):
        return args
    out: List[Any] = []
    for arg in args:
        if not _lists_lines(arg):
            out.append(arg)
            continue
        for item in arg:
            if isinstance(item, (list, tuple, dict, set)):
                raise TypeError(
                    f"{func_name}(): an item of a list of lines is a line or one value; "
                    f"got {type(item).__name__}."
                )
            out.append(item)
    return tuple(out)


def _by_period(func_name: str, args: tuple) -> bool:
    """Is the call one per period: as many operands as the registry's
    ``per_period_arity`` or more, a series among them?"""
    arity = per_period_arity(func_name)
    return arity is not None and len(args) >= arity and any(is_series(a) for a in args)


def _elementwise(func_name: str, args: tuple, reduce: Any, empty: Any = 0,
                 value_type: str | None = None) -> Variable:
    """Two or more operands with a series among them: one call per
    period, so the cell per period is ``=SUM(C2, C3)``. The period's
    operands reduce as one range does (an error is the result, a blank is
    skipped, nothing left is ``empty``), a ``None`` typed into the call
    being the 0 the formula holds; located series on other windows are
    aligned by date, tracks lift per coordinate and lines along the same
    keys key by key, as under the operators. A text operand is refused
    at build time (``TypeError``), not read as ``#VALUE!``: the
    spreadsheet skips text in a referenced cell, so no error value would
    match the workbook."""
    excel = _excel_reduce(reduce, empty, func_name)
    return elementwise(func_name, lambda *values: excel(list(values)), args,
                       value_type=value_type, exceptions='raise', typed_blank=0)


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
    return make_func_var(
        func_name, [VarRef(var)], lifted, 'float', 'scalar',
        source_code=f"{func_name}({label})", any_value_type=True,
    )


# How a refusal names what the function does with its numbers.
_VERBS = {"SUM": "adds", "AVERAGE": "averages", "MIN": "compares", "MAX": "compares",
          "PRODUCT": "multiplies"}


def _excel_reduce(reduce: Any, empty: Any = 0, name: str | None = None) -> Any:
    """``reduce`` over one range as the spreadsheet reads it: an error in
    the range is the result — the first one, in reading order — and a hole
    (a cell not entered yet) is skipped, as a blank is. With nothing left
    the result is ``empty``: 0 for SUM, MAX and MIN, ``#DIV/0!`` for
    AVERAGE. Given the function's ``name``, text among the values is
    refused at build time (``TypeError``) with a message naming both: the
    spreadsheet skips text in a range, so no error value would match the
    workbook."""
    def excel(values: List[Any]) -> Any:
        for value in values:
            if isinstance(value, str) and value.startswith("#"):
                return value
        if name is not None:
            for value in values:
                if isinstance(value, str):
                    raise TypeError(
                        f"{name} {_VERBS.get(name, 'reads')} numbers; {value!r} is not "
                        f"one. The spreadsheet skips text in a range, so no error value "
                        f"would match the workbook — give it the number."
                    )
        known = [value for value in values if value is not None]
        return reduce(known) if known else empty
    return excel


def _mean(values: List[Any]) -> Any:
    return builtins.sum(values) / len(values)


# The operands of one call added left to right, as ``a + b + c`` adds
# them: ``sum()`` compensates the rounding, and a period's subtotal would
# differ in its last digit from the same lines added with ``+``.
def _added(values: List[Any]) -> Any:
    return functools.reduce(operator.add, values)


def _mean_added(values: List[Any]) -> Any:
    return _added(values) / len(values)


def SUM(*args: AggOperand) -> Variable:
    """Sum of Variables, scalars, or lists (renders ``=SUM(...)``).

    Over ONE operand, the total of the row over its periods:
    ``mo.SUM(revenue)`` → ``=SUM(B2:F2)``. Over two or more operands with
    a series among them, the total PER PERIOD — a subtotal of lines, as
    ``a + b`` adds them — with scalars broadcast and a line with tracks
    taken track by track: ``mo.SUM(product_a, product_b)`` → ``=SUM(B2,
    B3)`` in every column, inside a loop container too. Over scalars
    only, one number. The total of several rows over all their periods
    is ``mo.SUM(a) + mo.SUM(b)``. Always returns a Variable with a
    tracked formula — the Excel cell is the product, not the value.

    Example::

        total = mo.SUM(revenues)         # =SUM(B2:F2)
        subtotal = mo.SUM(a, b, c)       # each period: =SUM(B2, B3, B4)
        total = mo.SUM(q1, q2, q3, q4)   # scalar lines: =SUM(B1, B2, B3, B4)
        total = mo.SUM(1, 2, 3)          # =SUM(1, 2, 3), value = 6
    """
    args = _listed("SUM", args)
    if _by_period("SUM", args):
        return _elementwise("SUM", args, _added)
    from ..core.tracks import TrackValues
    if (len(args) == 1 and isinstance(args[0], Variable)
            and isinstance(args[0]._value, TrackValues)):
        return _lifted_reduce("SUM", args[0], _excel_reduce(builtins.sum, name="SUM"))
    flat, _ = _flatten_args(args, "SUM")
    add = builtins.sum if len(args) == 1 else _added
    return _aggregate("SUM", args, _excel_reduce(add, name="SUM")(flat),
                      first_variable_value_type(args, 'float'))


def MAX(*args: AggOperand) -> Variable:
    """Maximum (renders ``=MAX(...)``).

    Over ONE operand, the largest value of the row: ``mo.MAX(revenue)``
    → ``=MAX(B2:F2)``. Over two or more operands with a series among
    them, the larger PER PERIOD — a floor — with scalars broadcast as
    ``a + b`` broadcasts them: ``mo.MAX(cash, inp.min_cash)`` →
    ``=MAX(C4, Inputs!B2)`` in every column, inside a loop container
    too. Over scalars only, one number: ``mo.MAX(a, b)`` → ``=MAX(B2, B3)``.
    """
    args = _listed("MAX", args)
    if _by_period("MAX", args):
        return _elementwise("MAX", args, builtins.max)
    from ..core.tracks import TrackValues
    if (len(args) == 1 and isinstance(args[0], Variable)
            and isinstance(args[0]._value, TrackValues)):
        return _lifted_reduce("MAX", args[0], _excel_reduce(builtins.max, name="MAX"))
    flat, _ = _flatten_args(args, "MAX")
    return _aggregate("MAX", args, _excel_reduce(builtins.max, name="MAX")(flat),
                      first_variable_value_type(args, 'float'))


def MIN(*args: AggOperand) -> Variable:
    """Minimum (renders ``=MIN(...)``).

    Over ONE operand, the smallest value of the row: ``mo.MIN(cash)`` →
    ``=MIN(B2:F2)``. Over two or more operands with a series among them,
    the smaller PER PERIOD — a cap — with scalars broadcast as ``a + b``
    broadcasts them: ``mo.MIN(sweep, balance)`` → ``=MIN(C2, C3)`` in
    every column, inside a loop container too. Over scalars only, one
    number: ``mo.MIN(a, b)`` → ``=MIN(B2, B3)``.
    """
    args = _listed("MIN", args)
    if _by_period("MIN", args):
        return _elementwise("MIN", args, builtins.min)
    from ..core.tracks import TrackValues
    if (len(args) == 1 and isinstance(args[0], Variable)
            and isinstance(args[0]._value, TrackValues)):
        return _lifted_reduce("MIN", args[0], _excel_reduce(builtins.min, name="MIN"))
    flat, _ = _flatten_args(args, "MIN")
    return _aggregate("MIN", args, _excel_reduce(builtins.min, name="MIN")(flat),
                      first_variable_value_type(args, 'float'))


def AVERAGE(*args: AggOperand) -> Variable:
    """Mean across Variables, scalars, or lists (renders ``=AVERAGE(...)``).

    Over ONE operand, the mean of the row over its periods:
    ``mo.AVERAGE(price)`` → ``=AVERAGE(B2:F2)``. Over two or more operands
    with a series among them, the mean PER PERIOD, with scalars
    broadcast and a line with tracks taken track by track:
    ``mo.AVERAGE(plan, actual)`` → ``=AVERAGE(B2, B3)`` in every column; a
    period whose cells are all blank is ``#DIV/0!``. Over scalars only,
    one number.
    """
    args = _listed("AVERAGE", args)
    if _by_period("AVERAGE", args):
        return _elementwise("AVERAGE", args, _mean_added, "#DIV/0!", 'float')
    from ..core.tracks import TrackValues
    if (len(args) == 1 and isinstance(args[0], Variable)
            and isinstance(args[0]._value, TrackValues)):
        return _lifted_reduce("AVERAGE", args[0], _excel_reduce(_mean, "#DIV/0!", "AVERAGE"))
    flat, _ = _flatten_args(args, "AVERAGE")
    mean = _mean if len(args) == 1 else _mean_added
    value = _excel_reduce(mean, "#DIV/0!", "AVERAGE")(flat)
    return _aggregate("AVERAGE", args, value, 'float')


register("SUM", shape="reduction", per_period_arity=2)
register("AVERAGE", shape="reduction", per_period_arity=2)
register("MAX", shape="reduction", per_period_arity=2)
register("MIN", shape="reduction", per_period_arity=2)
# ``COUNT``'s helper lives beside the conditional counts (functions/criteria.py).
register("COUNT", shape="reduction", per_period_arity=2)
# Written by the renderer without a helper of its own, for ``SUM(a * b)``
# over two rows — its arrays are ranges only (a comma list would be read
# as more arrays).
register("SUMPRODUCT", shape="range", bare=True)
