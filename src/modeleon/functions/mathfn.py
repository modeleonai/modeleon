# SPDX-License-Identifier: Apache-2.0
"""Element-wise math helpers — ``ABS``, ``ROUND``, ``INT``, ``MOD``; the
rounding family ``ROUNDUP``, ``ROUNDDOWN``, ``TRUNC``, ``CEILING``,
``FLOOR``; the powers and logarithms ``POWER``, ``SQRT``, ``EXP``, ``LN``,
``LOG``, ``LOG10``; ``SIGN``; and ``PRODUCT``, a reduction over one
series and one call per period over several.

Named ``mathfn`` (not ``math``) to avoid shadowing Python's stdlib ``math``.

**Invariant:** every call returns a :class:`Variable` with a tracked
formula. ``mo.ABS(-5)`` produces ``=ABS(-5)`` in Excel, not a bare ``5`` —
the formula is the product.

The functions after ``MOD`` read a cell as the spreadsheet does: a blank
is 0 (``ROUNDUP`` of a blank is 0, ``EXP`` of one is 1), a boolean is 1
or 0, a date its serial number (``FLOOR(date, 7)`` is a serial on a
multiple of 7), numeric text is its number (``'4'`` is 4) and other text
``#VALUE!``, a number outside the function's domain is ``#NUM!``
(``SQRT`` of a negative, ``LN`` of zero), one too large for a cell is
``#NUM!`` too, and an error cell passes through. A rounded amount keeps
its unit, a remainder and a multiple share one unit with what they divide
by; ``POWER`` and ``SQRT`` raise the unit as ``**`` does (``SQRT`` of
``USD^2`` is ``USD``; a power that is not whole in every dimension, or
that is a line, leaves none); ``PRODUCT`` multiplies its operands'
units; an exponential, a logarithm and a sign carry none.
"""

from __future__ import annotations

import builtins
import decimal
import math
import sys
from datetime import date, timedelta
from typing import Any, Union

from ..core.expr import Literal, VarRef
from ..core.unit import UNSETTLED
from ..core.variable import Variable
from ._helpers import (
    align_keyed, align_located, elementwise, is_series, make_func_var, operand, placed,
    series_node, src, unit_law,
)
from .aggregate import (
    _by_period, _elementwise, _excel_reduce, _flatten_args, _lifted_reduce, _listed,
)
from .registry import per_period_arity, register

Number = Union[int, float]
Numeric = Union[Variable, Number]


def _apply(py_fn, raw) -> tuple[Any, str]:
    """Apply ``py_fn`` element-wise over list values, and track by track
    over a line with tracks. A blank counts as zero (``ROUND`` of a blank
    cell is 0 in the spreadsheet); an error cell (``'#DIV/0!'``) passes
    through, as arithmetic already carries one; a zero divisor is
    ``#DIV/0!``."""
    from ..core.tracks import TrackValues

    def one(v: Any) -> Any:
        if isinstance(v, str) and v.startswith("#"):
            return v
        try:
            return py_fn(0 if v is None else v)
        except ZeroDivisionError:
            return '#DIV/0!'
        except TypeError:
            return '#VALUE!'

    def series(v: Any) -> Any:
        return [one(x) for x in v] if isinstance(v, list) else one(v)

    if isinstance(raw, TrackValues):
        return (TrackValues.lift(series, raw),
                'list' if raw.time_length is not None else 'scalar')
    if isinstance(raw, list):
        return [one(v) for v in raw], 'list'
    return one(raw), 'scalar'


def ABS(value: Numeric) -> Variable:
    """Absolute value (renders ``=ABS(...)``). Scalar or element-wise on lists.

    Accepts a Variable or a plain number. Always returns a Variable.
    """
    raw, expr = operand(value)
    calc, var_type = _apply(builtins.abs, raw)
    value_type = value.value_type if isinstance(value, Variable) else 'float'
    out = make_func_var(
        "ABS", [expr], calc, value_type, var_type,
        source_code=f"ABS({src(value)})",
    )
    return placed(out, None, align_keyed("ABS", [value], [raw])[1])


#: No limit on digits or exponent, so a very long whole number or an
#: absurd ``digits`` (``ROUND(x, -500)`` is 0) never raises a decimal error.
_ROUND_CONTEXT = decimal.Context(
    prec=decimal.MAX_PREC, rounding=decimal.ROUND_HALF_UP,
    Emax=decimal.MAX_EMAX, Emin=decimal.MIN_EMIN,
)
#: The 15 significant digits a cell shows, a half in the 16th going
#: away from zero.
_SHOWN_CONTEXT = decimal.Context(prec=15, rounding=decimal.ROUND_HALF_UP)


def _spreadsheet_round(v, digits: int, rounding: str = decimal.ROUND_HALF_UP):
    """Round ``v`` the way the spreadsheet's ``ROUND`` does.

    Two differences from Python's ``round``: a half goes AWAY FROM ZERO
    (2.5 -> 3, -2.5 -> -3; Python sends it to the even digit, 2), and
    the number rounded is the one a cell shows — the float to 15
    significant digits — not its binary value. ``1.005`` is held as
    1.00499999…, so Python rounds it to 1.0 while the sheet shows 1.005
    and rounds it to 1.01. (``ROUND_HALF_UP`` in :mod:`decimal` is
    away from zero.) A number of 16 or more whole digits has no decimals
    left to show and is rounded as held, so its whole digits never move.
    Anything that is not a finite number goes to Python's ``round`` as
    before. ``rounding`` is the :mod:`decimal` mode: ``ROUND_HALF_UP`` is
    ``ROUND``, ``ROUND_UP`` (away from zero) ``ROUNDUP``, ``ROUND_DOWN``
    (toward zero) ``ROUNDDOWN`` and ``TRUNC`` — each on the shown number.
    """
    if isinstance(v, bool) or not isinstance(v, (int, float)):
        return builtins.round(v, digits)
    if isinstance(v, int):
        if digits >= 0:
            return v                            # a whole number has no decimals
        shown = decimal.Decimal(v)
    elif not math.isfinite(v):
        return builtins.round(v, digits)
    elif abs(v) >= 1e15:
        shown = decimal.Decimal(v)              # exactly as held
    else:
        shown = _SHOWN_CONTEXT.plus(decimal.Decimal(v))
    if digits >= -shown.as_tuple().exponent:
        rounded = shown                         # no digit past the cut-off
    else:
        step = decimal.Decimal(1).scaleb(-digits, context=_ROUND_CONTEXT)
        rounded = shown.quantize(step, rounding=rounding, context=_ROUND_CONTEXT)
    return int(rounded) if isinstance(v, int) else float(rounded)


def ROUND(value: Numeric, digits: Numeric = 0) -> Variable:
    """Round a Variable or scalar to ``digits`` decimal places (renders ``=ROUND(...)``).

    Computes what the spreadsheet computes: halves round away from zero
    (``ROUND(2.5, 0) == 3``), on the number as a cell shows it
    (``ROUND(1.005, 2) == 1.01``). Negative ``digits`` round to tens,
    hundreds, … (``ROUND(1250, -2) == 1300``). ``digits`` may vary by
    period: a series of digits is one call per period, ``=ROUND(C5, C9)``
    in every column.
    """
    value_type = value.value_type if isinstance(value, Variable) else 'float'
    if is_series(digits):
        return elementwise("ROUND", _round_to, [value, digits], blank=0, value_type=value_type)
    raw, expr = operand(value)
    raw_digits, digits_expr = operand(digits)
    d = int(raw_digits or 0)            # blank digits read as 0
    calc, var_type = _apply(lambda v: _spreadsheet_round(v, d), raw)
    out = make_func_var(
        "ROUND", [expr, digits_expr], calc, value_type, var_type,
        source_code=f"ROUND({src(value)}, {d})",
    )
    return placed(out, None, align_keyed("ROUND", [value], [raw])[1])


def INT(value: Numeric) -> Variable:
    """Floor a Variable or scalar to the nearest integer (renders ``=INT(...)``).

    Matches Excel's ``INT`` (floor toward -∞), not Python's ``int()``
    (truncate toward 0). So ``INT(-2.5) == -3``.
    """
    raw, expr = operand(value)
    calc, var_type = _apply(math.floor, raw)
    out = make_func_var(
        "INT", [expr], calc, 'int', var_type,
        source_code=f"INT({src(value)})",
    )
    return placed(out, None, align_keyed("INT", [value], [raw])[1])


def MOD(number: Numeric, divisor: Numeric) -> Variable:
    """Remainder after division (renders ``=MOD(number, divisor)``).

    For the usual case, the ``%`` operator already compiles to ``MOD(...)``.
    Use this when you want the explicit function call form in the Excel output.
    Element-wise as ``%``: lines on other time windows are read by date,
    lines along keys key by key.
    """
    from ..core.tracks import TrackValues
    from ..core.variable_ops import _VariableArithmetic

    num_val, num_expr = operand(number)
    div_val, div_expr = operand(divisor)
    raws, keys = align_keyed("MOD", (number, divisor), [num_val, div_val])
    (num_val, div_val), window = align_located((number, divisor), raws)

    def mod(a: Any, b: Any) -> Any:
        # Element-wise with broadcasting, as the ``%`` operator: a blank is
        # 0, a zero divisor #DIV/0!, the result takes the divisor's sign.
        return _VariableArithmetic._broadcast_operation(
            a, b, lambda x, y: x % y, safe_divide=True, op_symbol='%')

    if isinstance(num_val, TrackValues) or isinstance(div_val, TrackValues):
        calc: Any = TrackValues.combine(num_val, div_val, mod)
        var_type = 'list' if calc.time_length is not None else 'scalar'
    else:
        calc = mod(num_val, div_val)
        var_type = 'list' if isinstance(calc, list) else 'scalar'
    out = make_func_var(
        "MOD", [num_expr, div_expr], calc, 'float', var_type,
        source_code=f"MOD({src(number)}, {src(divisor)})",
    )
    return placed(out, window, keys)


# ─── The rounding family, the powers and logarithms, the sign ─────────
#
# Each is one call per period through the shared :func:`elementwise`
# helper, over a kernel that takes one period's numbers. The recurrence
# template evaluator reuses the kernels, so ``"ROUNDUP({prev}, 0)"``
# computes what the cell computes.


def _number(v: Any) -> int | float:
    """A cell's number, as the spreadsheet reads an operand: a boolean is
    1 or 0, a date its serial number, numeric text its number; other text
    raises ``TypeError`` (the helper's ``#VALUE!``), a number no cell
    holds ``OverflowError`` (its ``#NUM!``)."""
    if type(v) is int or (type(v) is float and math.isfinite(v)):
        return v                                # the common case, without the reader's checks
    from .text import _number_of                # late, so neither module needs the other first
    return _number_of(v)


def _in_a_cell(out: int | float) -> Any:
    """``out`` as a cell holds it: a negative zero is 0, a number too
    large for a cell is ``#NUM!``."""
    if isinstance(out, int):
        return out if abs(out) <= sys.float_info.max else '#NUM!'
    return out + 0.0 if math.isfinite(out) else '#NUM!'


def _rounded(value: Any, digits: Any, rounding: str) -> Any:
    return _in_a_cell(_spreadsheet_round(_number(value), int(_number(digits)), rounding))


# ``ROUND`` of one period whose digits are a series: the value read as
# ``ROUND`` reads it over one number of digits, the digits as the
# spreadsheet reads a number, their fraction dropped.
def _round_to(value: Any, digits: Any) -> Any:
    return _spreadsheet_round(value, int(_number(digits)))


def _roundup(value: Any, digits: Any) -> Any:
    return _rounded(value, digits, decimal.ROUND_UP)


def _rounddown(value: Any, digits: Any) -> Any:
    return _rounded(value, digits, decimal.ROUND_DOWN)


def _trunc(value: Any, digits: Any = 0) -> Any:
    return _rounddown(value, digits)


def _multiple(x: int | float, significance: int | float, rounding: str) -> Any:
    """``x`` moved to a multiple of ``significance`` the way ``rounding``
    says. The quotient is the double the sheet divides to, read as a cell
    shows it: ``1.5 / 0.1`` is held as 15.000000000000002 and shown as 15,
    so ``CEILING(1.5, 0.1)`` stays 1.5. One of 16 or more whole digits has
    no decimals left to show and is taken as held, as ``ROUND`` takes such
    a number, so an exact multiple stays (``CEILING(1000000000000001, 1)``);
    one too large for a cell is ``#NUM!``, one too small for it 0."""
    held = x / significance
    if not math.isfinite(held):
        return '#NUM!'
    quotient = decimal.Decimal(held)
    if abs(held) < 1e15:
        quotient = _SHOWN_CONTEXT.plus(quotient)
    steps = quotient.to_integral_value(rounding=rounding)
    if isinstance(x, int) and isinstance(significance, int):
        return _in_a_cell(int(steps) * significance)
    return _in_a_cell(float(steps * _SHOWN_CONTEXT.plus(decimal.Decimal(significance))))


def _ceiling(value: Any, significance: Any) -> Any:
    x, step = _number(value), _number(significance)
    if step == 0:
        return 0 if isinstance(x, int) and isinstance(step, int) else 0.0
    if x > 0 and step < 0:
        return '#NUM!'
    return _multiple(x, step, decimal.ROUND_CEILING)


def _floor(value: Any, significance: Any) -> Any:
    x, step = _number(value), _number(significance)
    if step == 0 and x == 0:
        return 0 if isinstance(x, int) and isinstance(step, int) else 0.0
    if step == 0:
        raise ZeroDivisionError                 # the helper's #DIV/0!
    if x > 0 and step < 0:
        return '#NUM!'
    return _multiple(x, step, decimal.ROUND_FLOOR)


#: A whole number of more bits than this is past the largest a cell
#: holds (2 ** 1024): ``POWER`` of whole numbers beyond it is ``#NUM!``
#: without computing the number.
_POWER_BITS = 1100


def _approx_equal(a: float, b: float) -> bool:
    """The spreadsheet's test for the same number: equal to within the
    last four bits of the smaller, or half its 15th significant digit."""
    if a == b:
        return True
    if a == 0 or b == 0 or (a < 0) != (b < 0):
        return False
    least = min(abs(a), abs(b))
    return abs(a - b) < max(least * 2.0 ** -48, 10.0 ** math.floor(math.log10(least)) * 5e-15)


def _odd_root(x: int | float, y: float) -> Any:
    """A negative number to a fractional power, as the spreadsheet takes
    it: real only as an odd root, ``y`` the inverse of an odd whole
    number to within the 15th digit (``POWER(-8, 1/3)`` is -2,
    ``POWER(-32, -0.2)`` -0.5); any other fraction is ``#NUM!``."""
    nearest = 1 / y + (0.5 if y > 0 else -0.5)
    if not abs(nearest) < 2.0 ** 63:
        return '#NUM!'
    n = int(nearest)                            # toward zero
    if n % 2 == 0 or not _approx_equal(1 / n, y):
        return '#NUM!'
    return _in_a_cell(-((-x) ** y))


def _power(base: Any, exponent: Any) -> Any:
    x, y = _number(base), _number(exponent)
    if x == 0 and y == 0:
        return '#NUM!'
    if x < 0 and not (isinstance(y, int) or y.is_integer()):
        return _odd_root(x, y)
    if (isinstance(x, int) and isinstance(y, int) and y > 0 and abs(x) > 1
            and y * math.log2(abs(x)) > _POWER_BITS):
        return '#NUM!'                          # a whole number no cell can hold
    return _in_a_cell(x ** y)   # 0 ** -1 is the helper's #DIV/0!, a float too large its #NUM!


def _sqrt(value: Any) -> Any:
    x = _number(value)
    return '#NUM!' if x < 0 else math.sqrt(x)


def _exp(value: Any) -> Any:
    return math.exp(_number(value))             # too large is the helper's #NUM!


def _ln(value: Any) -> Any:
    x = _number(value)
    return '#NUM!' if x <= 0 else math.log(x)


def _log10(value: Any) -> Any:
    x = _number(value)
    return '#NUM!' if x <= 0 else math.log10(x)


def _log(value: Any, base: Any = 10) -> Any:
    x, b = _number(value), _number(base)
    if x <= 0 or b <= 0:
        return '#NUM!'
    if b == 1:
        return '#DIV/0!'
    if b == 10:
        return math.log10(x)                    # exact at the powers of ten
    if b == 2:
        return math.log2(x)
    return math.log(x) / math.log(b)


def _sign(value: Any) -> int:
    x = _number(value)
    return (x > 0) - (x < 0)


def _whole(x: Any) -> bool:
    """An operand that is whole in every period: a plain whole number, or
    a Variable of them."""
    if isinstance(x, Variable):
        return x.value_type == 'int'
    return isinstance(x, int) and not isinstance(x, bool)


def _whole_or_float(*args: Any) -> str:
    """``'int'`` when the result is whole by construction — every operand
    is — else ``'float'``."""
    return 'int' if all(_whole(a) for a in args) else 'float'


def ROUNDUP(value: Numeric, digits: Numeric) -> Variable:
    """Round away from zero to ``digits`` decimal places (renders ``=ROUNDUP(...)``).

    ``ROUNDUP(3.2, 0) == 4``, ``ROUNDUP(-3.2, 0) == -4``,
    ``ROUNDUP(31415.92654, -2) == 31500``; on the number as a cell shows
    it, so ``ROUNDUP(0.1 + 0.2, 1) == 0.3``. Scalar or element-wise on
    lists; ``digits`` may vary by period. A whole number stays whole.
    """
    return elementwise("ROUNDUP", _roundup, [value, digits], blank=0,
                       value_type=_whole_or_float(value))


def ROUNDDOWN(value: Numeric, digits: Numeric) -> Variable:
    """Round toward zero to ``digits`` decimal places (renders ``=ROUNDDOWN(...)``).

    ``ROUNDDOWN(3.9, 0) == 3``, ``ROUNDDOWN(-3.9, 0) == -3``,
    ``ROUNDDOWN(31415.92654, -2) == 31400``; on the number as a cell shows
    it, so ``ROUNDDOWN(4.35 * 100, 0) == 435``, not 434.
    """
    return elementwise("ROUNDDOWN", _rounddown, [value, digits], blank=0,
                       value_type=_whole_or_float(value))


def TRUNC(value: Numeric, digits: Numeric | None = None) -> Variable:
    """Drop the decimals past ``digits`` (default 0), toward zero (renders ``=TRUNC(...)``).

    ``TRUNC(8.9) == 8``, ``TRUNC(-8.9) == -8``, ``TRUNC(1234.5678, 2) ==
    1234.56``: ``ROUNDDOWN`` with an optional second argument, written
    ``=TRUNC(x)`` when none is given.
    """
    args = [value] if digits is None else [value, digits]
    return elementwise("TRUNC", _trunc, args, blank=0, value_type=_whole_or_float(value))


def CEILING(value: Numeric, significance: Numeric) -> Variable:
    """Round away from zero to a multiple of ``significance`` (renders ``=CEILING(...)``).

    The classic function's rules: ``CEILING(2.5, 1) == 3``,
    ``CEILING(0.234, 0.01) == 0.24``, an exact multiple stays; a negative
    number with a negative significance rounds away from zero
    (``CEILING(-2.5, -2) == -4``), with a positive one toward zero
    (``CEILING(-2.5, 2) == -2``); a positive number with a negative
    significance is ``#NUM!``; a significance of 0 gives 0. The quotient
    is read as a cell shows it (``CEILING(1.5, 0.1) == 1.5``); one too
    large for a cell is ``#NUM!``. A date is its serial number.
    """
    return elementwise("CEILING", _ceiling, [value, significance], blank=0,
                       value_type=_whole_or_float(value, significance))


def FLOOR(value: Numeric, significance: Numeric) -> Variable:
    """Round toward zero to a multiple of ``significance`` (renders ``=FLOOR(...)``).

    The classic function's rules: ``FLOOR(3.7, 2) == 2``, ``FLOOR(0.234,
    0.01) == 0.23``, an exact multiple stays; a negative number with a
    positive significance rounds away from zero (``FLOOR(-2.5, 2) ==
    -4``), with a negative one toward zero (``FLOOR(-2.5, -2) == -2``); a
    positive number with a negative significance is ``#NUM!``; a
    significance of 0 is ``#DIV/0!``, except that ``FLOOR(0, 0)`` is 0.
    """
    return elementwise("FLOOR", _floor, [value, significance], blank=0,
                       value_type=_whole_or_float(value, significance))


def POWER(base: Numeric, exponent: Numeric) -> Variable:
    """``base`` to the power ``exponent`` (renders ``=POWER(...)``).

    The ``**`` operator already writes ``^``; this is the function call in
    the cell. ``POWER(5, 2) == 25``, ``POWER(4, 5 / 4) == 5.656854…``. A
    negative base to a fractional power is real only as an odd root —
    ``POWER(-8, 1 / 3) == -2``, the exponent the inverse of an odd whole
    number to within the 15th digit — and ``#NUM!`` otherwise
    (``POWER(-8, 0.5)``); ``0 ** 0`` is ``#NUM!``, ``0 ** -1`` is
    ``#DIV/0!``, a result too large for a cell ``#NUM!``, one too small
    for it 0. The unit follows ``**``: each dimension's power times the
    exponent (``POWER(t, 2)`` in hr is hr^2, ``POWER(area, 0.5)`` in m^2
    is m); none when a power comes out fractional or the exponent is a
    line.
    """
    plain = not isinstance(exponent, Variable)
    whole = _whole(base) and plain and _whole(exponent) and exponent >= 0
    return elementwise("POWER", _power, [base, exponent], blank=0,
                       value_type='int' if whole else 'float')


def SQRT(value: Numeric) -> Variable:
    """Square root (renders ``=SQRT(...)``); a negative number is ``#NUM!``.
    The unit is halved as ``** 0.5`` halves it: ``SQRT`` of a variance in
    USD^2 is in USD; of a unit with an odd power, none."""
    return elementwise("SQRT", _sqrt, [value], blank=0, value_type='float')


def EXP(value: Numeric) -> Variable:
    """``e`` to the power ``value`` (renders ``=EXP(...)``): ``EXP(0) == 1``;
    a result too large for a cell is ``#NUM!``."""
    return elementwise("EXP", _exp, [value], blank=0, value_type='float')


def LN(value: Numeric) -> Variable:
    """Natural logarithm (renders ``=LN(...)``); zero or a negative number is ``#NUM!``."""
    return elementwise("LN", _ln, [value], blank=0, value_type='float')


def LOG(value: Numeric, base: Numeric | None = None) -> Variable:
    """Logarithm to ``base``, 10 when none is given (renders ``=LOG(x)`` /
    ``=LOG(x, base)``).

    ``LOG(100) == 2``, ``LOG(8, 2) == 3``; zero or a negative number or
    base is ``#NUM!``, a base of 1 is ``#DIV/0!``.
    """
    args = [value] if base is None else [value, base]
    return elementwise("LOG", _log, args, blank=0, value_type='float')


def LOG10(value: Numeric) -> Variable:
    """Logarithm to base 10 (renders ``=LOG10(...)``); zero or a negative number is ``#NUM!``."""
    return elementwise("LOG10", _log10, [value], blank=0, value_type='float')


def SIGN(value: Numeric) -> Variable:
    """The sign of a number, 1, 0 or -1 (renders ``=SIGN(...)``)."""
    return elementwise("SIGN", _sign, [value], blank=0, value_type='int')


def _product(values: list[Any]) -> Any:
    """``math.prod`` over a range's numbers, a date as its serial number;
    one that is not a number is refused at build time, as ``SUM`` refuses
    it."""
    factors = []
    for v in values:
        if isinstance(v, (date, timedelta)):
            v = _number(v)
        elif not isinstance(v, (int, float)):
            raise TypeError(
                f"PRODUCT multiplies numbers; {v!r} is not one. The spreadsheet "
                f"skips text in a range, so no error value would match the workbook — "
                f"give it the number."
            )
        factors.append(v)
    return _in_a_cell(math.prod(factors))


def PRODUCT(*args: Union[Variable, Number, list[Number]]) -> Variable:
    """Product of Variables, scalars, or lists (renders ``=PRODUCT(...)``).

    Over ONE operand, the product of the row's cells over its periods:
    ``PRODUCT(growth_factors)`` → ``=PRODUCT(B2:F2)``. Over two or more
    operands with a series among them, the product PER PERIOD, as ``a *
    b`` multiplies them — scalars broadcast, a line with tracks taken
    track by track: ``PRODUCT(price, volume)`` → ``=PRODUCT(B2, B3)`` in
    every column. Over scalars only, one number. The cells are read as
    LibreOffice reads them: a blank is skipped, a boolean is 1 or 0, a
    date its serial number, an error cell is the result, nothing left is
    0, a product too large for a cell is ``#NUM!``. Excel skips a boolean
    held in a referenced cell and counts only one typed into the call;
    the values here do not say where a boolean came from, so a referenced
    FALSE makes 0 as ``SUM`` counts it, where Excel would skip it. Text is
    refused at build time, as ``SUM`` refuses it. The result is a whole
    number only over whole numbers (a product of dates is a number, not a
    date). Its unit is the product of its operands' (``PRODUCT(price,
    volume)`` in USD/t and t is USD); one row in a dimension alone is
    refused, its cells' product being that dimension to the power of
    their count — a row of pure numbers (growth factors) is what a
    product over the periods runs along.
    """
    args = _listed("PRODUCT", args)
    if _by_period("PRODUCT", args):
        return _elementwise("PRODUCT", args, _product, 0, _whole_or_float(*args))
    from ..core.tracks import TrackValues
    product = _excel_reduce(_product)
    if (len(args) == 1 and isinstance(args[0], Variable)
            and isinstance(args[0]._value, TrackValues)):
        out = _lifted_reduce("PRODUCT", args[0], product)
    else:
        flat, ast_args = _flatten_args(args, "PRODUCT")
        out = make_func_var(
            "PRODUCT", ast_args, product(flat), _whole_or_float(*args),
            'scalar', source_code=f"PRODUCT({', '.join(src(a) for a in args)})",
        )
    return out


# ─── Units ──────────────────────────────────────────────────────────


def _rounded_unit(args: Any, unit_of: Any) -> Any:
    """A rounded value keeps its unit; the digits are a count."""
    return unit_of(args[0]) if args else None


def _power_unit(args: Any, unit_of: Any) -> Any:
    """The unit ``**`` gives: a constant exponent raises each dimension;
    an exponent that is a line may differ by period, so none."""
    from ..core.variable_ops import _VariableArithmetic
    if not args:
        return None
    base = unit_of(args[0])
    if base is UNSETTLED:
        return UNSETTLED
    exponent = args[1] if len(args) > 1 else None
    return _VariableArithmetic._unit_pow(
        base, exponent.value if isinstance(exponent, Literal) else None)


def _sqrt_unit(args: Any, unit_of: Any) -> Any:
    from ..core.variable_ops import _VariableArithmetic
    if not args:
        return None
    base = unit_of(args[0])
    return UNSETTLED if base is UNSETTLED else _VariableArithmetic._unit_pow(base, 0.5)


def _cell_count(var: Any) -> int:
    from ..core.tracks import TrackValues
    raw = var._value
    if isinstance(raw, TrackValues):
        return raw.time_length or 1
    return len(raw) if isinstance(raw, list) else 1


def _product_unit(args: Any, unit_of: Any) -> Any:
    """The product of the operands' units; a row in a dimension refused
    where its cells are multiplied together — one call per period
    multiplies each operand's unit once, as ``a * b`` does."""
    from ..core.variable_ops import _VariableArithmetic, _line_name
    arity = per_period_arity("PRODUCT") or 2
    by_period = len(args) >= arity and any(series_node(node) for node in args)
    unit = None
    waiting = False
    for node in args:
        factor = unit_of(node)
        if factor is UNSETTLED:
            waiting = True
            continue
        if factor is None:
            continue
        words = factor._unit_components
        if (words and not by_period and isinstance(node, VarRef)
                and _cell_count(node.var) > 1):
            hint = ""
            if len(words) == 1 and next(iter(words.values())) == 1:
                word = next(iter(words))
                name = word.upper() if word.isidentifier() else "NAME"
                hint = (f" If '{word}' labels a pure number, declare it once - {name} = "
                        f"mo.Unit.dimensionless('{word}') - and set that object as the "
                        f"row's unit.")
            raise ValueError(
                f"mo.PRODUCT multiplies the cells of {_line_name(node.var)}, a row in "
                f"{factor}: their product is in {factor} to the power of their count. "
                f"A product runs along a row of pure numbers (growth factors, 1 + a "
                f"rate); amounts multiply line by line: a * b.{hint}"
            )
        unit = _VariableArithmetic._unit_mul(unit, factor)
    return UNSETTLED if waiting else unit


register("ABS", shape="elementwise")
register("ROUND", shape="elementwise")
register("INT", shape="elementwise")
register("MOD", shape="elementwise")
register("ROUNDUP", shape="elementwise")
register("ROUNDDOWN", shape="elementwise")
register("TRUNC", shape="elementwise")
register("CEILING", shape="elementwise")
register("FLOOR", shape="elementwise")
register("POWER", shape="elementwise")
register("SQRT", shape="elementwise")
register("EXP", shape="elementwise")
register("LN", shape="elementwise")
register("LOG", shape="elementwise")
register("LOG10", shape="elementwise")
register("SIGN", shape="elementwise")
register("PRODUCT", shape="reduction", per_period_arity=2)
unit_law("ROUND", _rounded_unit)
unit_law("ROUNDUP", _rounded_unit)
unit_law("ROUNDDOWN", _rounded_unit)
unit_law("TRUNC", _rounded_unit)
unit_law("POWER", _power_unit)
unit_law("SQRT", _sqrt_unit)
unit_law("PRODUCT", _product_unit)
