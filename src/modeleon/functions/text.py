# SPDX-License-Identifier: Apache-2.0
"""Text helpers — ``LEN``, ``UPPER``, ``LOWER``, ``CONCAT``, ``LEFT``,
``RIGHT``, ``MID``, ``TRIM``, ``FIND``, ``SEARCH``, ``SUBSTITUTE``,
``VALUE``, ``TEXT``, ``TEXTJOIN``, ``REPT``.

**Invariant:** every call returns a :class:`Variable` with a tracked
formula — even when all inputs are plain strings. ``mo.UPPER("hello")``
produces ``=UPPER("hello")`` in Excel, not a bare ``"HELLO"``.

The helpers from ``LEFT`` on read an operand as the spreadsheet reads
the cell: a blank is empty text (0 where a number is asked for), TRUE
and FALSE are the words, a number is its General text as Excel writes
it — 15 significant digits, written out in full up to 1E20 and down to
a fraction whose spelling fits 20 characters (``0.00001``,
``1234567890123460``), E notation past them (``1E+20``,
``1.23456789012346E-05``) — and a date is its serial number: ``LEFT``
of a date cell gives the digits of 46023, as the workbook will;
``TEXT(d, 'yyyy-mm-dd')`` spells it. A blank given as a literal where
text is read (``None``, or a gap in a list) is written to the workbook
as empty text, which is what it reads as. A count or a position
(``LEFT``'s ``num_chars``, ``MID``'s ``start_num``) is truncated to a
whole number; numeric text is read as its number, other text is
``#VALUE!``. The first spreadsheet error among a period's operands
(``#VALUE!``, ``#DIV/0!``, …) is that period's result; text that only
starts with ``#`` — a ``'#12'`` label, a ``'#,##0'`` format, the
``'#'`` of a ``REPT`` bar — is text, while text that spells an error
exactly (``'#N/A'``) is read as that error: the engine carries errors as
their text. A result carries no unit: a label, a position and a number
read from text are not amounts.
"""

from __future__ import annotations

import decimal
import functools
import logging
import math
import numbers
import re
import sys
from datetime import date, datetime, timedelta
from typing import Any, Callable, Iterable, Union

from ..core.expr import Literal, VarRef
from ..core.variable import Variable
from ._helpers import (
    align_keyed, align_located, elementwise, is_series, make_func_var, placed,
    refuse_unpaired, src, text_operand,
)
from .registry import register

logger = logging.getLogger(__name__)

TextOperand = Union[Variable, str]


def _unary_text_func(
    func_name: str,
    py_fn: Callable[[Any], Any],
    result_value_type: str,
) -> Callable[[TextOperand], Variable]:
    """Build a Variable-aware string helper like LEN / UPPER / LOWER.

    Always returns a Variable, regardless of whether input is a Variable or
    a plain string.
    """
    def helper(value: TextOperand) -> Variable:
        raw, expr = text_operand(_text_arg(value))
        # A blank cell reads as empty text, as in Excel (LEN of a blank
        # is 0, UPPER of a blank is "").
        if isinstance(raw, list):
            calc: Any = [py_fn('' if v is None else v) for v in raw]
            var_type = 'list'
        else:
            calc = py_fn('' if raw is None else raw)
            var_type = 'scalar'
        out = make_func_var(
            func_name, [expr], calc, result_value_type, var_type,
            source_code=f"{func_name}({src(value)})",
        )
        return placed(out, None, align_keyed(func_name, [value], [raw])[1])

    return helper


LEN = _unary_text_func("LEN", lambda s: len(str(s)), 'int')
LEN.__doc__ = "Character length of a string Variable or plain string (renders ``=LEN(...)``)."

UPPER = _unary_text_func("UPPER", lambda s: str(s).upper(), 'string')
UPPER.__doc__ = "Uppercase a string Variable or plain string (renders ``=UPPER(...)``)."

LOWER = _unary_text_func("LOWER", lambda s: str(s).lower(), 'string')
LOWER.__doc__ = "Lowercase a string Variable or plain string (renders ``=LOWER(...)``)."


def _concat_piece(v: Any) -> str:
    """Render one operand value for concatenation, Excel-style.

    ``None`` → empty; an integer-valued float (``2022.0``) → ``"2022"``
    so period labels read ``"1Q 2022"`` not ``"1Q 2022.0"``.
    """
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def CONCAT(*args: TextOperand) -> Variable:
    """Concatenate Variables and/or plain strings (renders ``=CONCAT(...)``).

    Element-wise over list (period-indexed) operands, so
    ``mo.CONCAT(quarter, "Q ", year, " ", phase)`` over per-period
    Variables yields per-period labels (``"1Q 2022 A"``, …); lines on
    other time windows are read by date, as ``a + b`` reads them, and
    lines along keys key by key. Always returns a Variable carrying the
    ``=CONCAT(...)`` formula.
    """
    if not args:
        raise ValueError(
            "CONCAT() needs at least one argument. "
            "Syntax: mo.CONCAT(prefix, name)  or  mo.CONCAT(\"Q1 \", \"2025\"). "
            "Got 0 arguments."
        )

    raws, expr_args = [], []
    for a in args:
        raw, expr = text_operand(_text_arg(a))
        raws.append(raw)
        expr_args.append(expr)
    raws, keys = align_keyed("CONCAT", args, raws)
    raws, window = align_located(args, raws)
    refuse_unpaired(args, raws, window)

    if any(isinstance(r, list) for r in raws):
        n = max(len(r) for r in raws if isinstance(r, list))

        def at(r: Any, i: int) -> Any:
            if isinstance(r, list):
                return r[i] if i < len(r) else r[-1]
            return r

        value: Any = ["".join(_concat_piece(at(r, i)) for r in raws) for i in range(n)]
        var_type = 'list'
    else:
        value = "".join(_concat_piece(r) for r in raws)
        var_type = 'scalar'

    source_code = "CONCAT(" + ", ".join(src(a) for a in args) + ")"
    out = make_func_var(
        "CONCAT", expr_args, value, 'string', var_type,
        source_code=source_code,
    )
    return placed(out, window, keys)


# ─── Reading an operand as the spreadsheet does ─────────────────────

#: The longest text a cell holds.
_MAX_TEXT = 32767
#: The spreadsheet's error values, as arithmetic carries them. Text that
#: only starts with ``#`` is text to these helpers (module docstring).
_ERRORS = frozenset({'#DIV/0!', '#VALUE!', '#NUM!', '#N/A', '#NAME?', '#REF!', '#NULL!'})
#: Day 0 of the spreadsheet's calendar: 1 January 1900 is serial 1.
_SERIAL_EPOCH = date(1899, 12, 30)
#: The 15 significant digits a cell shows, a half in the 16th going away
#: from zero; and room for every digit of a number written out in full.
_SHOWN = decimal.Context(prec=15, rounding=decimal.ROUND_HALF_UP)
_WIDE = decimal.Context(prec=decimal.MAX_PREC, rounding=decimal.ROUND_HALF_UP,
                        Emax=decimal.MAX_EMAX, Emin=decimal.MIN_EMIN)


def _serial(d: date) -> int | float:
    """A date as the spreadsheet holds it — its day count from the epoch,
    a time of day as a fraction. Dates from 1 March 1900 on: the
    29 February 1900 Excel counts is not imitated."""
    if isinstance(d, datetime):
        days = (d.replace(tzinfo=None) - datetime(1899, 12, 30)) / timedelta(days=1)
        return int(days) if days.is_integer() else days
    return (d - _SERIAL_EPOCH).days


def _error_among(values: Iterable[Any]) -> Any:
    """The first spreadsheet error among the values, else None."""
    return next((v for v in values if isinstance(v, str) and v in _ERRORS), None)


def _shown(x: float) -> decimal.Decimal:
    """The number as a cell shows it, to 15 significant digits; a number
    no cell holds (infinite, not a number) is ``OverflowError`` (``#NUM!``)."""
    if not math.isfinite(x):
        raise OverflowError(f"{x!r} is not a number a cell holds")
    return _SHOWN.plus(decimal.Decimal(x))


def _general(x: float) -> str:
    """A number as General text, as Excel writes it (module docstring)."""
    if x == 0:
        return '0'
    shown = _shown(x)
    digits = ''.join(map(str, shown.as_tuple().digits)).rstrip('0')
    point = shown.adjusted()                # the power of ten of the first digit
    if point > 19 or (point < 0 and 1 - point + len(digits) > 20):
        mantissa = digits[0] + ('.' + digits[1:] if len(digits) > 1 else '')
        body = f"{mantissa}E{'+' if point >= 0 else '-'}{abs(point):02d}"
    elif point < 0:
        body = '0.' + '0' * (-point - 1) + digits
    elif len(digits) > point + 1:
        body = digits[:point + 1] + '.' + digits[point + 1:]
    else:
        body = digits + '0' * (point + 1 - len(digits))
    return ('-' if x < 0 else '') + body


def _text_arg(a: Any) -> Any:
    """A literal blank where text is read, as the empty text it reads as:
    the workbook spells a literal None as 0, whose text is "0". A
    Variable stays as it is — its blank is a blank cell."""
    if a is None:
        return ''
    if isinstance(a, list):
        return ['' if v is None else v for v in a]
    return a


def _text_of(v: Any) -> str:
    """The text the spreadsheet reads a value as (module docstring)."""
    if v is None:
        return ''
    if isinstance(v, str):
        return v
    if isinstance(v, bool):
        return 'TRUE' if v else 'FALSE'
    if isinstance(v, date):
        v = _serial(v)
    if isinstance(v, timedelta):            # a duration is its count of days
        v = v / timedelta(days=1)
    if isinstance(v, numbers.Real):
        return _general(float(v))           # a cell holds a double, whole or not
    return str(v)


#: Numeric text: an optional sign, spaces allowed after it, digits with
#: or without thousands separators (the first group of any length, the
#: others of three), a decimal part, an exponent. ASCII digits only.
_NUMBER_TEXT = re.compile(
    r'[+-]?\s*(?:[0-9]+(?:,[0-9]{3})*(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?')
#: A date written the one way every locale reads it — ``2026-01-31`` —
#: and a time of day after it.
_ISO_DATE = re.compile(r'([0-9]{4}-[0-9]{2}-[0-9]{2})(?:\s+(.+))?')
#: A time of day: hours, minutes and seconds (``12:30``, ``12:30:45.5``),
#: AM or PM after them (``1:30 PM``, ``12 AM``).
_TIME = re.compile(
    r'([0-9]+)(?::([0-9]{1,2})(?::([0-9]{1,2}(?:\.[0-9]*)?))?)?(?:\s*([AP]M))?', re.IGNORECASE)


def _day_fraction(text: str) -> float | None:
    """The fraction of a day time text spells (``'12:00'`` is 0.5), or
    None. Minutes and seconds stay below 60; with AM or PM the hours
    are 0 to 12, without them any count (``'25:00'`` is a day and an hour)."""
    match = _TIME.fullmatch(text)
    if match is None or (match[2] is None and match[4] is None):    # a bare number
        return None
    hours, minutes, seconds = int(match[1]), int(match[2] or 0), float(match[3] or 0)
    if minutes > 59 or seconds >= 60:
        return None
    if match[4] is not None:
        if hours > 12:
            return None
        hours = hours % 12 + (12 if match[4].upper() == 'PM' else 0)
    return (hours * 3600 + minutes * 60 + seconds) / 86400


def _parse_number(text: str) -> int | float:
    """The number numeric text spells — ``'1,234'``, ``'12%'``,
    ``'(1,234)'``, ``'1e3'``, a time of day (the fraction of a day), an
    ISO date with or without a time (its serial) — or ``TypeError`` (the
    spreadsheet's ``#VALUE!``); a number too large for a cell is
    ``OverflowError`` (``#NUM!``). Other date text, a currency sign and
    the words TRUE and FALSE are not read: the spreadsheet would, by its
    locale."""
    s = text.strip()
    iso = _ISO_DATE.fullmatch(s)
    if iso is not None:
        try:
            day = _serial(date.fromisoformat(iso[1]))
        except ValueError:
            raise TypeError(f"{text!r} is not a date") from None
        fraction = 0.0 if iso[2] is None else _day_fraction(iso[2])
        if fraction is None:
            raise TypeError(f"{text!r} is not a date and time")
        return day if iso[2] is None else day + fraction
    # A plus sign may lead a time, as it leads a number; a time of day
    # is never negative.
    fraction = _day_fraction(s[1:] if s[:1] == '+' else s)
    if fraction is not None:
        return fraction
    percent = s.endswith('%')               # read before the parentheses: '(5%)' is no number
    if percent:
        s = s[:-1].strip()
    negative = len(s) > 1 and s[0] == '(' and s[-1] == ')'
    if negative:
        s = s[1:-1].strip()
    if _NUMBER_TEXT.fullmatch(s) is None or (negative and s[:1] in ('+', '-')):
        raise TypeError(f"{text!r} is not a number")
    plain = ''.join(s.split()).replace(',', '')
    exponent = 'e' in plain.lower()
    if percent and exponent:
        raise TypeError(f"{text!r} is not a number")
    # An integral text is an int while a cell holds it exactly — 15 digits.
    whole = not (percent or exponent or '.' in plain)
    x: int | float = (int(plain) if whole and len(plain.lstrip('+-').lstrip('0')) <= 15
                      else float(plain))
    if abs(x) > sys.float_info.max:
        raise OverflowError(f"{text!r} is too large for a cell")
    if negative:
        x = -x
    return x / 100 if percent else x


def _number_of(v: Any) -> int | float:
    """The number the spreadsheet reads an operand as: a blank is 0,
    TRUE 1, a date its serial, a duration its days, numeric text its
    number."""
    if v is None:
        return 0
    if isinstance(v, bool):
        return int(v)
    if isinstance(v, str):
        return _parse_number(v)
    if isinstance(v, date):
        return _serial(v)
    if isinstance(v, timedelta):
        return v / timedelta(days=1)
    if isinstance(v, float) and not math.isfinite(v):
        raise OverflowError(f"{v!r} is not a number a cell holds")
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, numbers.Real):
        return float(v)
    raise TypeError(f"{v!r} is not a number")


def _count_of(v: Any) -> int:
    """A count or a position, truncated to a whole number as the
    spreadsheet truncates it (``LEFT(text, 2.9)`` takes 2)."""
    return math.trunc(_number_of(v))


def _flag_of(v: Any) -> bool:
    """TRUE or FALSE as the spreadsheet reads a logical operand: a number
    is TRUE unless it is zero, a blank is FALSE, the words TRUE and
    FALSE are themselves; other text is ``TypeError`` (``#VALUE!``)."""
    if isinstance(v, str) and v.strip().upper() in ('TRUE', 'FALSE'):
        return v.strip().upper() == 'TRUE'
    return _number_of(v) != 0


# ─── One cell at a time ─────────────────────────────────────────────


def _left(text: Any, num_chars: Any = 1) -> Any:
    t, k = _text_of(text), _count_of(num_chars)
    return '#VALUE!' if k < 0 else t[:k]


def _right(text: Any, num_chars: Any = 1) -> Any:
    t, k = _text_of(text), _count_of(num_chars)
    return '#VALUE!' if k < 0 else t[max(len(t) - k, 0):]


def _mid(text: Any, start_num: Any, num_chars: Any) -> Any:
    t, s, k = _text_of(text), _count_of(start_num), _count_of(num_chars)
    return '#VALUE!' if s < 1 or k < 0 else t[s - 1:s - 1 + k]


def _trim(text: Any) -> str:
    # Spaces only — not tabs, not the non-breaking space — and a run of
    # them inside the text collapses to one, as Excel's TRIM collapses it.
    return ' '.join(part for part in _text_of(text).split(' ') if part)


def _wildcard_pattern(find: str) -> str:
    """``SEARCH``'s pattern as a regular expression: ``?`` one character,
    ``*`` any run, ``~`` makes the character after it literal (whatever
    it is); a last ``~`` is itself."""
    out: list[str] = []
    i = 0
    while i < len(find):
        c = find[i]
        if c == '~' and i + 1 < len(find):
            out.append(re.escape(find[i + 1]))
            i += 2
            continue
        out.append('.' if c == '?' else '.*' if c == '*' else re.escape(c))
        i += 1
    return ''.join(out)


def _position(find: Any, within: Any, start: Any, wildcards: bool) -> Any:
    f, w, s = _text_of(find), _text_of(within), _count_of(start)
    if s < 1 or s > len(w):
        return '#VALUE!'
    if not f:
        # Excel's rule: empty find_text is found at start_num. LibreOffice
        # says #VALUE!.
        return s
    if wildcards:
        match = re.compile(_wildcard_pattern(f), re.IGNORECASE | re.DOTALL).search(w, s - 1)
        at = match.start() if match else -1
    else:
        at = w.find(f, s - 1)
    return at + 1 if at >= 0 else '#VALUE!'


def _find(find_text: Any, within_text: Any, start_num: Any = 1) -> Any:
    return _position(find_text, within_text, start_num, wildcards=False)


def _search(find_text: Any, within_text: Any, start_num: Any = 1) -> Any:
    return _position(find_text, within_text, start_num, wildcards=True)


def _substitute(text: Any, old_text: Any, new_text: Any, instance_num: Any = None) -> Any:
    t, old, new = _text_of(text), _text_of(old_text), _text_of(new_text)
    if instance_num is None:
        out = t.replace(old, new) if old else t
    else:
        k = _count_of(instance_num)
        if k < 1:
            return '#VALUE!'
        at = -1
        for _ in range(k):          # the k-th occurrence, counted without overlap
            at = t.find(old, at + len(old) if at >= 0 else 0) if old else -1
            if at < 0:
                return t
        out = t[:at] + new + t[at + len(old):]
    return out if len(out) <= _MAX_TEXT else '#VALUE!'


def _value(text: Any) -> Any:
    if isinstance(text, bool):
        return '#VALUE!'
    return _number_of(text)


def _rept(text: Any, number_times: Any) -> Any:
    t, k = _text_of(text), _count_of(number_times)
    return '#VALUE!' if k < 0 or len(t) * k > _MAX_TEXT else t * k


def _join(delimiter: Any, ignore_empty: Any, pieces: list[Any]) -> Any:
    words = [_text_of(v) for v in pieces]
    if _flag_of(ignore_empty):
        words = [w for w in words if w]
    out = _text_of(delimiter).join(words)
    return out if len(out) <= _MAX_TEXT else '#VALUE!'


#: The number formats ``TEXT`` writes: ``0``, ``00``, ``0.00``, ``#,##0``,
#: ``#,##0.00``, any of them with a ``%``.
_NUMBER_FORMAT = re.compile(
    r'(?P<thousands>#,##)?(?P<whole>0+)(?:\.(?P<decimals>0+))?(?P<percent>%)?')
_MONTHS = ('January', 'February', 'March', 'April', 'May', 'June', 'July',
           'August', 'September', 'October', 'November', 'December')
_WEEKDAYS = ('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday')
#: The date codes ``TEXT`` writes, longest first so ``mmm`` is read
#: before ``mm``; English names, as Excel's English locale writes them.
_DATE_FIELDS: dict[str, Callable[[date], str]] = {
    'yyyy': lambda d: f"{d.year:04d}",
    'yy': lambda d: f"{d.year % 100:02d}",
    'mmmmm': lambda d: _MONTHS[d.month - 1][0],
    'mmmm': lambda d: _MONTHS[d.month - 1],
    'mmm': lambda d: _MONTHS[d.month - 1][:3],
    'mm': lambda d: f"{d.month:02d}",
    'm': lambda d: str(d.month),
    'dddd': lambda d: _WEEKDAYS[d.weekday()],
    'ddd': lambda d: _WEEKDAYS[d.weekday()][:3],
    'dd': lambda d: f"{d.day:02d}",
    'd': lambda d: str(d.day),
}
_DATE_SEPARATORS = frozenset(' -/.,:')


def _date_pieces(fmt: str) -> list[str] | None:
    """A date format as its codes and separators (``'dd.mm.yyyy'`` →
    ``['dd', '.', 'mm', '.', 'yyyy']``), or None when it is not one."""
    pieces: list[str] = []
    i = 0
    while i < len(fmt):
        code = next((c for c in _DATE_FIELDS if fmt.startswith(c, i)), None)
        if code is not None:
            pieces.append(code)
            i += len(code)
        elif fmt[i] in _DATE_SEPARATORS:
            pieces.append(fmt[i])
            i += 1
        else:
            return None
    return pieces if any(p in _DATE_FIELDS for p in pieces) else None


def _format_number(x: float, thousands: bool, whole: int, decimals: int, percent: bool) -> str:
    # The number as a cell shows it, 15 significant digits and zeros past
    # them, rounded half away from zero to the format's decimals.
    shown = _shown(x)
    if percent:
        shown = shown.scaleb(2, context=_WIDE)
    rounded = shown.quantize(decimal.Decimal(1).scaleb(-decimals), context=_WIDE)
    integer, dot, fraction = format(rounded.copy_abs(), 'f').partition('.')
    integer = integer.zfill(whole)          # the padding counts digits: '#,##000000' is 001,234
    if thousands:
        head = len(integer) % 3 or 3
        integer = ','.join([integer[:head]]
                           + [integer[i:i + 3] for i in range(head, len(integer), 3)])
    # The sign is the number's, as the spreadsheet shows it: -0.4 to "0" is "-0".
    return ('-' if x < 0 else '') + integer + dot + fraction + ('%' if percent else '')


def _text(value: Any, format_text: Any, warned: set[str]) -> Any:
    fmt = _text_of(format_text)
    v = value
    if isinstance(v, str):
        try:
            v = _parse_number(v)
        except TypeError:
            return v                    # text that is not a number is given back as it is
    if isinstance(v, bool):
        return 'TRUE' if v else 'FALSE'
    if not fmt:
        return ''                       # an empty format shows nothing
    x = float(_number_of(v))            # a blank is 0, a date its serial number
    low = fmt.lower()
    if low in ('general', '@'):
        return _general(x)
    number = _NUMBER_FORMAT.fullmatch(fmt)
    if number is not None:
        return _format_number(x, bool(number['thousands']), len(number['whole']),
                              len(number['decimals'] or ''), bool(number['percent']))
    pieces = _date_pieces(low)
    if pieces is not None:
        d = v if isinstance(v, date) else _SERIAL_EPOCH + timedelta(days=math.floor(x))
        return ''.join(_DATE_FIELDS[p](d) if p in _DATE_FIELDS else p for p in pieces)
    if fmt not in warned:
        warned.add(fmt)
        logger.warning(
            "TEXT: the format %r is not one the engine writes (0, 0.00, #,##0, 0%%, "
            "yyyy-mm-dd, mmm yyyy, …). The workbook's cell shows the spreadsheet's "
            "result; the Python value is the number as General text.", fmt,
        )
    return _general(x)


def _per_cell(func_name: str, pyfunc: Callable[..., Any], args: list[Any],
              value_type: str = 'string') -> Variable:
    """One call per period through the shared helper. The first error
    among a period's operands is its result and text that only starts
    with ``#`` is text (:data:`_ERRORS`); the result carries no unit
    (module docstring)."""
    @functools.wraps(pyfunc)               # keeps the helper's arity check on ``pyfunc``
    def guarded(*values: Any) -> Any:
        error = _error_among(values)
        return pyfunc(*values) if error is None else error

    out = elementwise(func_name, guarded, args, value_type=value_type, propagate_errors=False)
    out._unit = None
    return out


# ─── The helpers ────────────────────────────────────────────────────


def LEFT(text: Any, num_chars: Any = None) -> Variable:
    """The first ``num_chars`` characters, one by default (renders ``=LEFT(...)``).

    Element-wise over a series. A count past the end gives the whole
    text; a negative count is ``#VALUE!``.
    """
    text = _text_arg(text)
    return _per_cell("LEFT", _left, [text] if num_chars is None else [text, num_chars])


def RIGHT(text: Any, num_chars: Any = None) -> Variable:
    """The last ``num_chars`` characters, one by default (renders ``=RIGHT(...)``).

    Element-wise over a series. A count past the start gives the whole
    text; a negative count is ``#VALUE!``.
    """
    text = _text_arg(text)
    return _per_cell("RIGHT", _right, [text] if num_chars is None else [text, num_chars])


def MID(text: Any, start_num: Any, num_chars: Any) -> Variable:
    """``num_chars`` characters from position ``start_num``, 1-based (renders ``=MID(...)``).

    Element-wise over a series. A start past the end gives empty text;
    a start below 1 or a negative count is ``#VALUE!``.
    """
    return _per_cell("MID", _mid, [_text_arg(text), start_num, num_chars])


def TRIM(text: Any) -> Variable:
    """Text without its leading, trailing and repeated spaces (renders ``=TRIM(...)``).

    Spaces only, as in Excel: a tab or a non-breaking space stays, and a
    run of spaces inside the text collapses to one. Element-wise over a
    series.
    """
    return _per_cell("TRIM", _trim, [_text_arg(text)])


def FIND(find_text: Any, within_text: Any, start_num: Any = None) -> Variable:
    """1-based position of ``find_text`` in ``within_text`` (renders ``=FIND(...)``).

    Case-sensitive, no wildcards; the search starts at ``start_num``
    (1 by default). Absent text, a start below 1 or past the end is
    ``#VALUE!``. Empty ``find_text`` is found at ``start_num``
    (``FIND('', 'abc')`` is 1), as Excel documents it; LibreOffice gives
    ``#VALUE!``. Element-wise over a series.
    """
    args = [_text_arg(find_text), _text_arg(within_text)]
    return _per_cell("FIND", _find, args if start_num is None else [*args, start_num],
                     value_type='int')


def SEARCH(find_text: Any, within_text: Any, start_num: Any = None) -> Variable:
    """1-based position of ``find_text`` in ``within_text``, case-insensitive
    (renders ``=SEARCH(...)``).

    Excel's wildcards: ``?`` is one character, ``*`` any run, ``~``
    makes the character after it literal (``~?``, ``~*``, ``~~``; a
    last ``~`` is itself). The search starts at ``start_num`` (1 by
    default). No match, a start below 1 or past the end is ``#VALUE!``;
    empty ``find_text`` is found at ``start_num`` as in :func:`FIND`.
    Element-wise over a series.
    """
    args = [_text_arg(find_text), _text_arg(within_text)]
    return _per_cell("SEARCH", _search, args if start_num is None else [*args, start_num],
                     value_type='int')


def SUBSTITUTE(text: Any, old_text: Any, new_text: Any, instance_num: Any = None) -> Variable:
    """``text`` with ``old_text`` replaced by ``new_text`` (renders ``=SUBSTITUTE(...)``).

    Every occurrence, or only the ``instance_num``-th (1-based, counted
    without overlap: the second ``aa`` in ``aaa`` is not there). Case-
    sensitive; an instance below 1 is ``#VALUE!``, one past the last
    occurrence leaves the text as it is. Element-wise over a series.
    """
    args = [_text_arg(text), _text_arg(old_text), _text_arg(new_text)]
    return _per_cell("SUBSTITUTE", _substitute,
                     args if instance_num is None else [*args, instance_num])


def VALUE(text: Any) -> Variable:
    """The number numeric text spells (renders ``=VALUE(...)``).

    ``'1,234'`` is 1234, ``'12%'`` 0.12, ``'(1,234)'`` -1234, ``'1e3'``
    1000, spaces around the number (and after its sign) ignored; a comma
    is a thousands separator — the first group of any length, the others
    of three. A time of day is the fraction of a day it spells —
    ``'12:00'`` is 0.5, ``'1:30 PM'`` 0.5625; hours, minutes and seconds
    (``'12:30:45.5'``), AM or PM after them — and ``'2026-01-31'`` is
    the date's serial number, a time after it (``'2026-01-31 18:00'``)
    adding its fraction. A number stays itself, a blank is 0, a date its
    serial number. Other text — words, empty text, TRUE, a currency
    sign, date text in any other form (which the spreadsheet would read
    by its locale: ``'1/2'``, a ``T`` between the date and the time), a
    negative time, minutes or seconds of 60 or more — is ``#VALUE!``; a
    number too large for a cell (``'1e999'``) is ``#NUM!``. Element-wise
    over a series.
    """
    return _per_cell("VALUE", _value, [text], value_type='float')


def TEXT(value: Any, format_text: Any) -> Variable:
    """A number as text in a format (renders ``=TEXT(value, format)``).

    The formats the engine writes, as the spreadsheet writes them:

    - ``0``, ``0.0``, ``0.00``, …, ``00``, ``000`` (zero-padded),
      ``#,##0``, ``#,##0.00``, …, and any of them with ``%`` (``0%``,
      ``0.0%``): halves round away from zero on the number as a cell
      shows it (``TEXT(2.5, '0')`` is ``3``, ``TEXT(1.005, '0.00')`` is
      ``1.01``); a negative number keeps its sign, ``TEXT(-0.4, '0')``
      being ``-0``;
    - ``General`` and ``@``: the number as General text (15 significant
      digits, E notation from 1E20 and for a fraction whose spelling
      passes 20 characters);
    - a date format of ``yyyy``, ``yy``, ``mmmmm`` (the month's first
      letter), ``mmmm``, ``mmm``, ``mm``, ``m``, ``dddd``, ``ddd``,
      ``dd``, ``d`` with ``-``, ``/``, ``.``, ``,``, ``:`` or spaces
      between (``yyyy-mm-dd``, ``dd.mm.yyyy``, ``mmm yyyy``, ``d mmmm
      yyyy``), English month and day names, over a date or a serial
      number. Codes are case-insensitive; hours, minutes and seconds are
      not written.

    Every format shows the number to 15 significant digits, zeros past
    them (``TEXT(123456789012345678, '0')`` is ``123456789012346000``);
    an empty format shows nothing. Any other format is written to the
    workbook as given, with a warning; its Python value is the number as
    General text. Numeric text is formatted as its number, other text is
    given back as it is, a blank reads as 0, TRUE and FALSE are the
    words. Element-wise over a series.
    """
    warned: set[str] = set()
    return _per_cell("TEXT", lambda v, f: _text(v, f, warned), [value, _text_arg(format_text)])


def REPT(text: Any, number_times: Any) -> Variable:
    """``text`` repeated ``number_times`` times (renders ``=REPT(...)``).

    A count is truncated to a whole number; a negative one, or a result
    longer than a cell holds (32,767 characters — Excel's limit, which
    LibreOffice does not apply), is ``#VALUE!``. Element-wise over a
    series.
    """
    return _per_cell("REPT", _rept, [_text_arg(text), number_times])


def TEXTJOIN(delimiter: Any, ignore_empty: Any, *texts: Any) -> Variable:
    """Texts joined by a delimiter (renders ``=TEXTJOIN(...)``, stored
    as ``_xlfn.TEXTJOIN`` — an Excel 2019 name).

    One text for the whole call: a series operand is joined along its
    periods, as the workbook's ``TEXTJOIN(", ", TRUE, B2:F2)`` joins the
    row, and a scalar operand is one piece. ``ignore_empty`` TRUE skips
    blanks and empty text; FALSE keeps them as empty pieces between
    delimiters (a number is TRUE unless it is zero, a blank FALSE, the
    words themselves are read, other text is ``#VALUE!``). Numbers,
    TRUE/FALSE and dates read as the spreadsheet reads them (module
    docstring); the first error among the operands is the result.
    """
    from ..core.shape import reject_axised
    from ..core.tracks import TrackValues
    if not texts:
        raise ValueError(
            "TEXTJOIN() needs a delimiter, ignore_empty and at least one text. "
            "Syntax: mo.TEXTJOIN(', ', True, names)  or  mo.TEXTJOIN('-', True, a, b). "
            "Got the delimiter and ignore_empty only."
        )
    delimiter = _text_arg(delimiter)
    texts = tuple(_text_arg(t) for t in texts)
    for role, a in (("delimiter", delimiter), ("ignore_empty", ignore_empty)):
        if is_series(a):
            raise ValueError(
                f"TEXTJOIN: the {role} is one value for the whole call, not a series. "
                f"Give one text ({role}=', ') or one cell."
            )
    raws: list[Any] = []
    ast_args: list[Any] = []
    for a in (delimiter, ignore_empty, *texts):
        if isinstance(a, Variable):
            reject_axised(a, "TEXTJOIN")
            if isinstance(a._value, TrackValues):
                raise ValueError(
                    "TEXTJOIN() over a Variable with tracks joins one track at a "
                    "time — slice a track first (.at(track='...'))."
                )
            ast_args.append(VarRef(a))
            raws.append(a._value)
        else:
            ast_args.append(Literal(a))
            raws.append(a)
    delim, ignore, *rows = raws
    pieces = [v for row in rows for v in (row if isinstance(row, list) else [row])]
    value = _error_among((delim, ignore, *pieces))
    if value is None:
        try:
            value = _join(delim, ignore, pieces)
        except TypeError:
            value = '#VALUE!'
        except OverflowError:
            value = '#NUM!'
    source_code = "TEXTJOIN(" + ", ".join(src(a) for a in (delimiter, ignore_empty, *texts)) + ")"
    return make_func_var("TEXTJOIN", ast_args, value, 'string', 'scalar',
                         source_code=source_code)


register("LEN", shape="elementwise")
register("UPPER", shape="elementwise")
register("LOWER", shape="elementwise")
# The Excel renderer writes CONCAT as the ``&`` operator, not as a call.
register("CONCAT", shape="elementwise", passthrough=False)
register("LEFT", shape="elementwise")
register("RIGHT", shape="elementwise")
register("MID", shape="elementwise")
register("TRIM", shape="elementwise")
register("FIND", shape="elementwise")
register("SEARCH", shape="elementwise")
register("SUBSTITUTE", shape="elementwise")
register("VALUE", shape="elementwise")
register("TEXT", shape="elementwise")
register("REPT", shape="elementwise")
# TEXTJOIN joins whole rows: a series operand is a range, a scalar one a
# cell. An Excel 2019 name — the file stores it as ``_xlfn.TEXTJOIN``.
register("TEXTJOIN", shape="range", file_prefix="_xlfn.")
