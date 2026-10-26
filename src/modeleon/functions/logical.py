# SPDX-License-Identifier: Apache-2.0
"""Logical and error helpers — ``IFERROR``, ``ISERROR``, ``ISNUMBER``,
``XOR``, ``IFS``, ``SWITCH``.

Each is one call per period through the shared :func:`elementwise`
helper: a series operand gives a series, scalars and a one-value series
broadcast as ``a + b`` does, other length mismatches are refused, tracked
operands lift per coordinate, and the cell per period is the call over
that period's cells (``=IFERROR(C4, 0)``). ``IF``, ``AND``, ``OR``,
``NOT``, ``ISBLANK`` and ``CHOOSE`` live in :mod:`conditional`.

An operand is read as the cell the workbook holds for it: ``None`` is a
blank cell, and so is an empty text typed into a line (the workbook
writes it as an empty cell); a formula's empty text, or one given to the
call, is text.

The error functions read the engine's error markers — ``'#DIV/0!'``,
``'#VALUE!'``, ``'#N/A'``, …, any text starting with ``#``, as
arithmetic carries one: ``IFERROR`` replaces a marker with its fallback,
``ISERROR`` names one, ``ISNUMBER`` is FALSE on one. A branch not taken
is never read, as the spreadsheet never reads it: ``IFS(a == 0, 0, True,
1 / a)`` is 0 where ``a`` is 0. ``XOR``, ``IFS`` and ``SWITCH`` came to
Excel after 2007: the file stores them as ``_xlfn.XOR`` / ``_xlfn.IFS``
/ ``_xlfn.SWITCH``, shown without the prefix.
"""

from __future__ import annotations

import math
import numbers
from datetime import date, timedelta
from typing import Any, Callable, Iterable, Sequence

from ..core.variable import Variable
from ..core.variable_ops import _to_serial
from ._helpers import elementwise, is_error, same_unit, shared_unit, unit_law
from .conditional import _blank0
from .registry import register

#: ``SWITCH`` without a default, when no case matches.
_NO_DEFAULT = object()


# ─── Reading the operands ───────────────────────────────────────────


def _typed_line(arg: Any) -> bool:
    # Values typed in, not computed: the workbook writes the line's
    # empty text as an empty cell.
    return isinstance(arg, Variable) and arg._expr is None


def _as_cells(pyfunc: Callable[..., Any], args: Sequence[Any]) -> Callable[..., Any]:
    # ``pyfunc`` over one period's operands as their cells read: an empty
    # text typed into a line is the blank cell the workbook writes.
    typed = [_typed_line(a) for a in args]
    if not any(typed):
        return pyfunc

    def read(*values: Any) -> Any:
        return pyfunc(*[None if t and v == '' else v for t, v in zip(typed, values)])
    return read


def _values(arg: Any) -> list[Any]:
    # Every value an operand gives — a scalar, a series, a series per
    # track — read as its cells.
    from ..core.tracks import TrackValues
    raw = arg._value if isinstance(arg, Variable) else arg
    parts = list(raw.values()) if isinstance(raw, TrackValues) else [raw]
    flat = [v for part in parts for v in (part if isinstance(part, list) else [part])]
    return [None if v == '' else v for v in flat] if _typed_line(arg) else flat


def _type_of(values: Iterable[Any]) -> str:
    # Typed by the first value that is neither a blank nor an error marker.
    sample = next((v for v in values if v is not None and not is_error(v)), None)
    if isinstance(sample, bool):
        return 'bool'
    if isinstance(sample, int):
        return 'int'
    if isinstance(sample, str):
        return 'string'
    if isinstance(sample, date):
        return 'datetime'
    return 'float'


def _shown(value: Any) -> Any:
    # A formula that lands on a blank cell shows 0.
    from ..core.tracks import TrackValues
    if isinstance(value, TrackValues):
        return TrackValues.lift(_shown, value)
    if isinstance(value, list):
        return [_blank0(v) for v in value]
    return _blank0(value)


def _boolean(out: Variable) -> Variable:
    # A TRUE / FALSE carries no unit; the formula is the readable source.
    out._source_from_formula()
    out._unit = None
    return out


def _picked(out: Variable, takes: Sequence[Any], sample: list[Any]) -> Variable:
    # A value picked among ``takes``, the operands the result may be: typed
    # by ``sample``'s first value that is not blank and not an error, a
    # picked blank shows 0, and the unit is the one ``takes`` share — a
    # condition or a case compared is never the result. Two units among
    # ``takes`` were refused by the function's law, as it was built.
    out._source_from_formula()
    out.value_type = _type_of(sample)
    out._value = _shown(out._value)
    out._unit = shared_unit(takes) if out.value_type in ('int', 'float') else None
    return out


# ─── Errors ─────────────────────────────────────────────────────────


def _iferror(value: Any, fallback: Any) -> Any:
    return fallback if is_error(value) else value


def IFERROR(value: Any, fallback: Any) -> Variable:
    """``value``, or ``fallback`` where ``value`` is an error (renders ``=IFERROR(...)``).

    The way to catch a ``#DIV/0!`` before it cascades: ``mo.IFERROR(profit
    / revenue, 0)`` is the margin where revenue is nonzero and 0 where it
    is not. Both operands are read in every period (both are the cell's
    precedents); a blank shows 0, as a formula landing on a blank cell
    does. The result is typed by the value where it is not an error — a
    text fallback does not make a row of numbers text — and keeps the
    unit the operands share: ``IFERROR(cost, 0)`` is still in dollars.
    """
    args = [value, fallback]
    out = elementwise("IFERROR", _as_cells(_iferror, args), args, propagate_errors=False)
    return _picked(out, args, _values(value) + _values(fallback))


def ISERROR(value: Any) -> Variable:
    """TRUE where ``value`` is an error marker (renders ``=ISERROR(...)``).

    Any error — ``#DIV/0!``, ``#VALUE!``, ``#N/A``, ``#NUM!``, … — as the
    engine carries it: a text value starting with ``#``. Never an error
    itself. Element-wise over a series.
    """
    return _boolean(elementwise("ISERROR", is_error, [value], value_type='bool',
                                propagate_errors=False))


def _isnumber(value: Any) -> bool:
    if isinstance(value, bool):
        return False
    return isinstance(value, (numbers.Real, date, timedelta))


def ISNUMBER(value: Any) -> Variable:
    """TRUE where ``value`` is a number (renders ``=ISNUMBER(...)``).

    A date is a number (the spreadsheet stores it as a serial number), a
    duration is its day count; a boolean, text (a number written as text
    included), a blank and an error are not. Never an error itself.
    Element-wise over a series.
    """
    return _boolean(elementwise("ISNUMBER", _isnumber, [value], value_type='bool',
                                propagate_errors=False))


# ─── XOR ────────────────────────────────────────────────────────────


def _xor(*values: Any) -> Any:
    # A blank or an empty text is skipped, as the spreadsheet skips such a
    # cell; a number is TRUE when nonzero; nothing left is #VALUE!.
    counted = [v for v in values if v is not None and v != '']
    if not counted:
        return '#VALUE!'
    return sum(bool(v) for v in counted) % 2 == 1


def XOR(*conditions: Any) -> Variable:
    """Exclusive OR — TRUE when an odd number of conditions hold (renders ``=XOR(...)``).

    Element-wise over series operands. A blank or an empty text is
    skipped, as the spreadsheet skips such a cell; a number counts as TRUE
    when nonzero; the first error among a period's operands is the
    period's result; nothing to count is ``#VALUE!``. Text is refused
    (``TypeError``): the spreadsheet skips text in a cell it references
    but reads text typed into the formula as ``#VALUE!``, so no value
    would match the workbook — compare it instead, ``XOR(answer == "yes",
    …)``. Stored in the file as ``_xlfn.XOR(...)``.
    """
    if not conditions:
        raise TypeError("XOR takes at least one condition.")
    for condition in conditions:
        text = next((v for v in _values(condition)
                     if isinstance(v, str) and v and not is_error(v)), None)
        if text is not None:
            raise TypeError(
                f"XOR counts conditions, and {text!r} is text: the spreadsheet skips "
                f"text in a cell it references but reads text typed into the formula "
                f"as #VALUE!, so no value would match the workbook. Compare the text "
                f"instead — XOR(answer == 'yes', …)."
            )
    return _boolean(elementwise("XOR", _xor, list(conditions), value_type='bool'))


# ─── IFS and SWITCH ─────────────────────────────────────────────────


def _truth(value: Any) -> Any:
    # A condition as the spreadsheet reads it: an error stays one, a blank
    # is FALSE, a number is TRUE when nonzero, the text TRUE or FALSE (in
    # any case) is that boolean, other text — an empty one too — #VALUE!.
    if is_error(value):
        return value
    if value is None:
        return False
    if isinstance(value, str):
        word = value.casefold()
        return word == 'true' if word in ('true', 'false') else '#VALUE!'
    return bool(value)


def _ifs(*operands: Any) -> Any:
    for condition, value in zip(operands[0::2], operands[1::2]):
        truth = _truth(condition)
        if is_error(truth):
            return truth
        if truth:
            return value
    return '#N/A'


def IFS(*pairs: Any) -> Variable:
    """The value of the first condition that holds (renders ``=IFS(...)``).

    ``mo.IFS(score >= 90, "A", score >= 80, "B", True, "C")`` — condition
    and value pairs, read in order; a final ``True`` is the catch-all.
    No condition holding is ``#N/A``, as in the spreadsheet. A condition
    that is an error makes the result that error; a value or a condition
    after the first that holds is never read. A blank condition is FALSE,
    a number is TRUE when nonzero, the text ``"TRUE"`` or ``"FALSE"`` is
    that boolean, other text is ``#VALUE!``; a picked blank shows 0.
    Element-wise over series operands. The result is typed by the first
    value taken that is not blank and not an error, and keeps a unit the
    values share — a condition's unit is not the result's. Stored in the
    file as ``_xlfn.IFS(...)``.
    """
    if not pairs or len(pairs) % 2:
        raise TypeError(
            f"IFS takes condition and value pairs — IFS(cond1, value1, cond2, "
            f"value2, …); got {len(pairs)} operand{'s' if len(pairs) != 1 else ''}."
        )
    args = list(pairs)
    out = elementwise("IFS", _as_cells(_ifs, args), args, propagate_errors=False)
    return _picked(out, args[1::2], _values(out))


#: Numbers closer than this share of both are equal, as the spreadsheet
#: compares them — about 15 significant digits.
_EPSILON = 2.0 ** -48
#: A double holds every whole number up to this one exactly.
_WHOLE = 2.0 ** 53 - 1


def _equal_numbers(a: float, b: float) -> bool:
    # Within 2**-48 of both, so a computed 0.1 * 3 is the 0.3 typed into a
    # case; zero equals only zero, and two whole numbers only each other.
    if a == b:
        return True
    if a == 0.0 or b == 0.0:
        return False
    d, a, b = abs(a - b), abs(a), abs(b)
    if not math.isfinite(d):
        return False
    if all(x <= _WHOLE and x.is_integer() for x in (d, a, b)):
        return False
    return d < a * _EPSILON and d < b * _EPSILON


def _key(value: Any) -> tuple[str, Any]:
    # What a cell compares as: text regardless of case, a boolean only with
    # a boolean, a date or a duration by its serial number.
    if isinstance(value, bool):
        return ('bool', value)
    if isinstance(value, str):
        return ('text', value.casefold())
    serial = _to_serial(value)
    if isinstance(serial, numbers.Real):
        return ('number', serial)
    return ('other', value)


_BLANK_MATCHES = (('number', 0), ('text', ''), ('bool', False))


def _matches(expression: Any, case: Any) -> bool:
    if expression is None or case is None:
        other = case if expression is None else expression
        return other is None or _key(other) in _BLANK_MATCHES
    kind, left = _key(expression)
    case_kind, right = _key(case)
    if kind != case_kind:
        return False
    if kind == 'number':
        return _equal_numbers(float(left), float(right))
    return bool(left == right)


def _switch(expression: Any, *rest: Any) -> Any:
    cases, default = (rest[:-1], rest[-1]) if len(rest) % 2 else (rest, _NO_DEFAULT)
    if is_error(expression):
        return expression
    for case, value in zip(cases[0::2], cases[1::2]):
        if is_error(case):
            return case
        if _matches(expression, case):
            return value
    return '#N/A' if default is _NO_DEFAULT else default


def SWITCH(expression: Any, *cases: Any) -> Variable:
    """The value of the first case equal to ``expression`` (renders ``=SWITCH(...)``).

    ``mo.SWITCH(inputs.scenario, "base", 0.05, "bull", 0.08, 0.03)`` —
    the expression, then case and value pairs, then an optional default
    (an odd number of operands after the expression); no case matching
    and no default is ``#N/A``, as in the spreadsheet. A case is equal as
    cells are equal: numbers to about 15 significant digits (a computed
    ``0.1 * 3`` is ``0.3``), text regardless of case, a date by its serial
    number, a blank as 0, as empty text and as FALSE, TRUE/FALSE only
    with each other, and never across kinds (1 is not "1"). An expression
    or a case that is an error makes the result that error; a case after
    the match is never read; a picked blank shows 0. Element-wise over
    series operands. The result is typed by the first value taken that is
    not blank and not an error, and keeps a unit the values share — the
    expression and the cases are compared, never taken. Stored in the
    file as ``_xlfn.SWITCH(...)``.
    """
    if len(cases) < 2:
        raise TypeError(
            f"SWITCH takes an expression, then case and value pairs — "
            f"SWITCH(expr, case1, value1, …, [default]); got "
            f"{len(cases)} operand{'s' if len(cases) != 1 else ''} after the expression."
        )
    args = [expression, *cases]
    takes = [*cases[1::2], *(cases[-1:] if len(cases) % 2 else ())]
    out = elementwise("SWITCH", _as_cells(_switch, args), args, propagate_errors=False)
    return _picked(out, takes, _values(out))


register("IFERROR", shape="elementwise")
register("ISERROR", shape="elementwise")
register("ISNUMBER", shape="elementwise")
register("XOR", shape="elementwise", file_prefix="_xlfn.")
register("IFS", shape="elementwise", file_prefix="_xlfn.")
register("SWITCH", shape="elementwise", file_prefix="_xlfn.")


def _ifs_unit(args: Any, unit_of: Callable[[Any], Any]) -> Any:
    """``IFS`` takes the value after each condition."""
    return same_unit("IFS", args[1::2], unit_of)


def _switch_unit(args: Any, unit_of: Callable[[Any], Any]) -> Any:
    """``SWITCH`` compares its expression with each case, as ``==`` does,
    and takes the value after each case, or the default."""
    pairs = (len(args) - 1) // 2
    _refuse_a_case_in_another_unit(args[0], args[1:1 + 2 * pairs:2], unit_of)
    default = list(args[-1:]) if (len(args) - 1) % 2 else []
    return same_unit("SWITCH", [*args[2::2], *default], unit_of)


def _refuse_a_case_in_another_unit(expression: Any, cases: Sequence[Any],
                                   unit_of: Callable[[Any], Any]) -> None:
    """A case in another unit than the expression is refused, as comparing
    the two lines is; a number or an unlabelled line goes with any."""
    from ..core.expr import VarRef
    from ..core.unit import UNSETTLED
    from ..core.variable_ops import _incompatible_units, _line_name

    def named(node: Any) -> str:
        return _line_name(node.var) if isinstance(node, VarRef) else "a formula"

    own = unit_of(expression)
    if own is None or own is UNSETTLED:
        return
    for case in cases:
        unit = unit_of(case)
        if unit is not None and unit is not UNSETTLED and unit != own:
            detail = _incompatible_units(own, unit).split(": ", 1)[1]
            raise ValueError(
                f"mo.SWITCH compares {named(expression)} with the case "
                f"{named(case)}; they are in {detail}"
            )


unit_law("IFS", _ifs_unit)
unit_law("SWITCH", _switch_unit)
