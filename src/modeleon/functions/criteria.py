# SPDX-License-Identifier: Apache-2.0
"""Conditional aggregates — ``SUMIF``, ``SUMIFS``, ``COUNTIF``, ``COUNTIFS``,
``AVERAGEIF``, ``AVERAGEIFS``, ``MAXIFS``, ``MINIFS`` — and the counts
``COUNT`` and ``COUNTA``.

Each reads whole ranges and gives one number, written as the spreadsheet
writes it: ``=SUMIF(B2:F2, ">=10", B3:F3)``. The counts follow the rule
every aggregate follows: over two or more operands with a series among
them they are one call per period, as ``a + b`` combines its operands
(``=COUNT(B2, B3)`` in every column), over one series the count of its
cells, over scalars only one number. A range argument is a row
— a list-valued Variable, written as the range of its cells — or a
Python list of scalar lines whose cells stand one under another in a
column (a field down the records of a table, a block of scalar lines),
written as a vertical range, ``Deals!C3:C5``; one element of a row,
``flow[0]``, is such a line too. Cells that form no run cannot be one
range: the cell carries the computed value and ``check_compat`` lists it
— never a comma list, which these functions read as further arguments.

The criteria is read as the spreadsheet reads it, once, in
:func:`parse_criteria` and :func:`matches`: a number, a date or a boolean
means equality; a string carries an optional operator — ``'>=10'``,
``'<>0'``, ``'<5'`` — and without one means equality; what follows the
operator is read as a cell entry: a number (``'10'``, ``'-2.5'``,
``'1e3'``, ``'5%'``), ``TRUE`` / ``FALSE``, a date written ``2026-01-31``
(a time of day may follow: ``2026-01-31 12:00``), else text. An entry
that a spreadsheet reads as a number or a date in some locales and as
text in others — ``'1,000'``, ``'$100'``, ``'(5)'``, ``'12:00'``,
``'1/31/2026'``, ``'Jan 31, 2026'`` — is refused when the model is built,
with the form to write instead. Text compares case-insensitively, whole
cell, with the wildcards ``*`` (any run of characters) and ``?`` (one
character), ``~`` taking the character after it literally. A scalar
Variable is its value (a cell holding ``'>=10'`` is that criteria) and
is written as that cell; ``('>=', threshold)`` or
``mo.CONCAT('>=', threshold)`` joins an operator to a cell and is
written as ``">=" & B2``. A per-period criteria is not one: the
conditional sum for one period is a call over that period's column of
the records, the vertical case above.

The edges, as the spreadsheet reads them. Two numbers are equal when
they differ by less than 2**-48 of the smaller (``0.1 + 0.2`` meets
``0.3``), as the spreadsheet compares them. A date is a number — its
serial — so a date cell meets a numeric criteria and a number cell a
date one. Text that reads as a number meets an equality (the text
``'10'`` equals 10) but never an ordering (``'>5'``); text equal to the
criteria's own text meets it too (the text ``TRUE`` meets ``'TRUE'``);
a boolean is never a number. ``''`` matches a blank cell and empty
text, ``'='`` only a blank cell, ``'<>'`` whatever is not blank; an
ordering with nothing after it compares text with the empty text
(``'>='`` meets every text, ``'<'`` none). ``'<>x'`` matches blanks,
numbers and errors too. An error cell matches only its own text
(``'#DIV/0!'``) or a ``'<>'``; a criteria cell holding an error is the
result. ``None`` as the criteria, typed or read from a cell left blank,
is the number 0, as a reference to an empty cell reads; joined to an
operator, ``('=', None)``, it is the operator alone, as a blank cell
joined to one is. Empty text typed into a line (``''``) is a blank cell
in the workbook and is read as one; the empty text a formula gives is a
cell: ``'<>'`` and ``COUNTA`` count it, ``''`` matches it, ``'='`` does
not. In the summed, averaged or ranked range a blank, a text and a
boolean are skipped, a date is its serial (``MAXIFS`` over dates gives
the date, ``AVERAGEIF`` the mean date, ``SUMIF`` a number of days), and
an error among the MATCHED cells is the result; nothing matched sums to
0, averages to ``#DIV/0!`` and ranks (``MAXIFS``, ``MINIFS``) to 0. Text
orders case-insensitively by code point (``'<m'``); a spreadsheet's
collation may place accented text elsewhere. Ranges must hold the same
number of cells (Excel's ``#VALUE!`` — refused when the model is
built). A line with tracks is refused: slice a coordinate first.

Units, as the operators read them: the cells a call totals, averages or
ranks are in one unit, the result's (a sum of USD is USD), and lines in
two are refused, as ``SUM`` refuses them; a criteria in a unit compares
with its range in that unit — ``('>=', cap)`` with ``cap`` in hours
over amounts in USD is refused, as ``amount >= cap`` is. A typed
criteria and an unlabelled line go with any; a count has none.

LibreOffice departs from Excel in three places, and the engine keeps
Excel's rule: there a boolean cell is a number (``TRUE`` meets ``'>0'``
and is counted and summed), numeric text in a criteria range meets a
numeric criteria only as text (the text ``'10'`` meets ``"10"``, not
the number 10), and a wildcard matches a number's digits.

Not covered: a criteria over a two-dimensional block (a table's period
block as one range), union ranges (``(B2:B5, D2:D5)``), and a criteria
that is a series.
"""

from __future__ import annotations

import bisect
import builtins
import math
import operator
import re
import unicodedata
import weakref
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from functools import partial
from itertools import chain
from typing import Any, Callable, Iterator, Sequence

from ..core.expr import Expr, FuncCall, ListExpr, Literal, Subscript, VarRef
from ..core.shape import reject_axised
from ..core.unit import UNSETTLED
from ..core.variable import Variable
from ..core.variable_ops import _to_serial
from ._helpers import (
    UnitLaw, _per_period, align_keyed, align_located, is_error, make_func_var, placed,
    operand_unit, same_unit,
    shared_unit, src, unit_law,
)
from .aggregate import _excel_reduce, _mean, _refuse_listed_series
from .registry import per_period_arity, register

Number = int | float
RangeArg = Variable | Sequence[Any]

# ─── The criteria, as the spreadsheet reads it ──────────────────────

#: Longest first, so ``'<>'`` and ``'>='`` are read before ``'<'`` and ``'>'``.
_OPERATORS = ('<>', '>=', '<=', '=', '>', '<')

_NUMBER = re.compile(r'[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+|%)?', re.ASCII)
#: A date as every spreadsheet reads it, a time of day after it or not.
_ISO_DATE = re.compile(
    r'(\d{4})-(\d{1,2})-(\d{1,2})(?: +(\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d{1,6}))?)?)?',
    re.ASCII)
#: The ISO form with a ``T`` before the time, which not every spreadsheet reads.
_ISO_T = re.compile(r'\d{4}-\d{1,2}-\d{1,2}T\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?',
                    re.ASCII | re.IGNORECASE)
_DIGITS = frozenset('0123456789')
_LETTERS = re.compile(r'[^\W\d_]+')
#: What a number, a date or a time is written with besides digits: the
#: words of a date or a time of day, and the marks around the digits (a
#: currency sign besides).
_DATE_WORDS = frozenset({
    'jan', 'january', 'feb', 'february', 'mar', 'march', 'apr', 'april', 'may', 'jun',
    'june', 'jul', 'july', 'aug', 'august', 'sep', 'sept', 'september', 'oct', 'october',
    'nov', 'november', 'dec', 'december', 'am', 'pm',
})
_MARKS = frozenset(",./:-+()%'")
#: The error values a cell holds — text that only starts with ``#`` is text.
_ERRORS = frozenset({'#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A'})


def _as_number(text: str) -> Number | None:
    """The number a cell entry reads as, else None: ``'10'``, ``'-2.5'``,
    ``'1e3'``, ``'5%'``; never ``'nan'``, ``'1_000'``, ``'1e3%'`` or ``''``."""
    entry = text.strip()
    if not _NUMBER.fullmatch(entry):
        return None
    if entry.endswith('%'):
        return float(entry[:-1]) / 100
    try:
        return int(entry)
    except ValueError:
        return float(entry)


def _as_date(found: re.Match[str]) -> date | datetime | None:
    """The date (with its time of day, if one follows) a ``2026-01-31``
    entry stands for; None for a day before 1900 or a time past the day."""
    year, month, day, hour, minute, second, fraction = found.groups()
    try:
        when = datetime(int(year), int(month), int(day), int(hour or 0), int(minute or 0),
                        int(second or 0), int((fraction or '0').ljust(6, '0')))
    except ValueError:
        return None
    if when.year < 1900:
        return None
    return when if hour is not None else when.date()


def _no_such_day(found: re.Match[str]) -> bool:
    """``2026-02-30``: no day of the calendar, so text to every spreadsheet."""
    year, month, day = (int(part) for part in found.groups()[:3])
    try:
        date(year, month, day)
    except ValueError:
        return True
    return False


def _read_by_locale(entry: str) -> bool:
    """Does a spreadsheet read this entry as a number or a date in some
    locales — digits among separators, a currency sign, parentheses, a
    month's name, AM / PM: ``'1,000'``, ``'$100'``, ``'(5)'``, ``'12:00'``,
    ``'1/31/2026'``, ``'Jan 31, 2026'``, ``'2026-01-31T12:00'``?"""
    if not any(ch in _DIGITS for ch in entry):
        return False
    if _ISO_T.fullmatch(entry):
        return True
    if any(word.casefold() not in _DATE_WORDS for word in _LETTERS.findall(entry)):
        return False
    return all(ch in _DIGITS or ch in _MARKS or ch.isspace() or unicodedata.category(ch) == 'Sc'
               for ch in _LETTERS.sub('', entry))


def _read_entry(text: str) -> Any:
    """What a criteria's comparand stands for: a number when the text reads
    as one, ``TRUE`` / ``FALSE`` as booleans, a ``2026-01-31`` date as
    that date, else the text itself; ValueError for an entry read as a
    number or a date in some spreadsheets only (a day before 1900 among
    them)."""
    number = _as_number(text)
    if number is not None:
        return number
    entry = text.strip()
    if entry.upper() in ('TRUE', 'FALSE'):
        return entry.upper() == 'TRUE'
    found = _ISO_DATE.fullmatch(entry)
    if found is not None:
        when = _as_date(found)
        if when is not None:
            return when
        if _no_such_day(found):
            return text
    if _read_by_locale(entry):
        raise ValueError(
            f"the criteria entry {entry!r} reads as a number or a date in some "
            f"spreadsheets and as text in others. Write a number plainly ('1000', "
            f"'-5'), a date as 2026-01-31 (or join the operator to a date value: "
            f"('>=', mo.DATE(2026, 1, 31))), a time of day as a fraction of a day "
            f"(12:00 is 0.5); to match the text itself, put '?' for one of its "
            f"marks ('10?20')."
        )
    return text


def _parse(criteria: Any) -> tuple[str, Any, str | None]:
    """``(operator, comparand, text)``: the comparand as :func:`parse_criteria`
    reads it, with the criteria's own text after the operator when it was
    given as text (None otherwise)."""
    if criteria is None:
        return '=', 0, None
    if isinstance(criteria, (bool, int, float, date, timedelta)):
        return '=', criteria, None
    text = str(criteria)
    if text == '':
        return '=', '', None
    op, rest = '=', text
    for candidate in _OPERATORS:
        if text.startswith(candidate):
            op, rest = candidate, text[len(candidate):]
            break
    if rest == '':
        return op, None, None
    return op, _read_entry(rest), rest


def parse_criteria(criteria: Any) -> tuple[str, Any]:
    """``(operator, comparand)`` as Excel reads a criteria argument.

    A number, a date or a boolean is ``('=', itself)``; a string carries
    an optional leading operator (``'<>'``, ``'>='``, ``'<='``, ``'='``,
    ``'>'``, ``'<'``) and the rest is read as a cell entry — a number
    when it reads as one, ``TRUE`` / ``FALSE`` as booleans, a
    ``2026-01-31`` date (a time of day may follow) as that date, else
    text. ``''`` is the blank criterion ``('=', '')`` — a blank cell or
    empty text; an operator with nothing after it has None for its
    comparand (``'='`` a blank cell only, ``'<>'`` whatever is not
    blank); ``None`` — a cell left blank — is ``('=', 0)``, as a reference
    to an empty cell reads. An entry read as a number in some locales
    only (``'1,000'``, ``'12:00'``, ``'1/31/2026'``) is a ValueError.
    """
    op, comparand, _text = _parse(criteria)
    return op, comparand


def _wildcard(pattern: str) -> re.Pattern[str]:
    """``*`` any run of characters, ``?`` one, ``~`` the character after
    it taken literally (a last ``~`` is itself); anything else literal;
    the whole cell, case-insensitively."""
    parts: list[str] = []
    i = 0
    while i < len(pattern):
        ch = pattern[i]
        if ch == '~' and i + 1 < len(pattern):
            parts.append(re.escape(pattern[i + 1]))
            i += 2
            continue
        parts.append('.*' if ch == '*' else '.' if ch == '?' else re.escape(ch))
        i += 1
    return re.compile(''.join(parts), re.IGNORECASE | re.DOTALL)


def _compare(a: Any, b: Any, op: str) -> bool:
    if op == '=':
        return bool(a == b)
    if op == '<>':
        return bool(a != b)
    if op == '>':
        return bool(a > b)
    if op == '<':
        return bool(a < b)
    if op == '>=':
        return bool(a >= b)
    return bool(a <= b)


_E48 = 2.0 ** -48
_E53 = 2.0 ** 53


def _whole(x: float) -> bool:
    return x < _E53 and x == math.floor(x)


def _approx_equal(a: Number, b: Number) -> bool:
    """Equal as the spreadsheet compares two numbers: the same, or apart by
    less than 2**-48 of the smaller — never two whole numbers that are not
    the same."""
    if a == b:
        return True
    if isinstance(a, int) and isinstance(b, int):
        return False
    x, y = float(a), float(b)
    if x == 0.0 or y == 0.0:
        return False
    d = abs(x - y)
    if not math.isfinite(d):
        return False
    x, y = abs(x), abs(y)
    if _whole(d) and _whole(x) and _whole(y):
        return False
    return d < x * _E48 and d < y * _E48


def _numbers_compare(a: Number, b: Number, op: str) -> bool:
    """``a op b`` with equality as :func:`_approx_equal` sees it."""
    same = _approx_equal(a, b)
    if op == '=':
        return same
    if op == '<>':
        return not same
    if op == '>':
        return a > b and not same
    if op == '<':
        return a < b and not same
    if op == '>=':
        return a > b or same
    return a < b or same


def _cell_number(value: Any) -> Number | None:
    """A cell as a number: a number is itself, a date its serial (a
    duration its days); text, a boolean, a blank and an error are not."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return value
    if isinstance(value, (date, timedelta)):
        serial: Number = _to_serial(value)
        return serial
    return None


_EPOCH = datetime(1899, 12, 30)


def _from_serial(serial: float) -> date | datetime:
    """The date a serial stands for; a datetime when it carries a time
    of day (the mean of dates seldom lands on a whole day)."""
    when = _EPOCH + timedelta(days=serial)
    return when.date() if when.time() == time() else when


def _is_text(value: Any) -> bool:
    return isinstance(value, str) and not is_error(value)


def _equal(comparand: Any, text: str | None) -> Callable[[Any], bool]:
    """Whether a cell equals a number, a date or a boolean: a number or a
    date by its serial; text that reads as a number as that number (a
    boolean never does); text equal to the criteria's own ``text``, when
    it was given as text, in any case."""
    folded = text.casefold() if text is not None else None

    def same_text(v: Any) -> bool:
        return folded is not None and _is_text(v) and v.casefold() == folded

    if isinstance(comparand, bool):
        return lambda v: v == comparand if isinstance(v, bool) else same_text(v)
    target = _to_serial(comparand)

    def same(v: Any) -> bool:
        number = _cell_number(v)
        if number is None and _is_text(v):
            if same_text(v):
                return True
            number = _as_number(v)
        return number is not None and _approx_equal(number, target)

    return same


def _matcher(op: str, comparand: Any, text: str | None = None) -> Callable[[Any], bool]:
    """Whether one cell's value meets ``op comparand``; ``text`` is the
    criteria's own text after the operator, when it was given as text."""
    if comparand is None or (isinstance(comparand, str) and comparand == ''):
        # Nothing after the operator. '' alone matches a blank cell and
        # empty text; '=' a blank cell only; '<>' whatever is not blank,
        # empty text included, as COUNTA counts it; an ordering compares
        # text with the empty text.
        if op == '=':
            if comparand == '':
                return lambda v: v is None or v == ''
            return lambda v: v is None
        if op == '<>':
            return lambda v: v is not None
        return lambda v: _is_text(v) and _compare(v.casefold(), '', op)
    if isinstance(comparand, str):
        if is_error(comparand):
            token = comparand.upper()
            if op == '=':
                return lambda v: isinstance(v, str) and v.upper() == token
            if op == '<>':
                return lambda v: not (isinstance(v, str) and v.upper() == token)
            return lambda v: False
        if op in ('=', '<>'):
            pattern = _wildcard(comparand)

            def equal(v: Any) -> bool:
                return _is_text(v) and pattern.fullmatch(v) is not None

            return equal if op == '=' else (lambda v: not equal(v))
        folded = comparand.casefold()
        return lambda v: _is_text(v) and _compare(v.casefold(), folded, op)
    if op in ('=', '<>'):
        same = _equal(comparand, text)
        return same if op == '=' else (lambda v: not same(v))
    if isinstance(comparand, bool):
        return lambda v: isinstance(v, bool) and _compare(v, comparand, op)
    # An ordering reads numbers and dates only (a date criteria is its
    # serial, as ``DATE(...)`` gives one); text that reads as a number
    # takes no part in it.
    target = _to_serial(comparand)

    def ordered(v: Any) -> bool:
        number = _cell_number(v)
        return number is not None and _numbers_compare(number, target, op)

    return ordered


def matches(value: Any, criteria: Any) -> bool:
    """Does one cell's ``value`` meet ``criteria``, read by
    :func:`parse_criteria`? ``matches('North', 'n*')`` is True,
    ``matches(None, '<>')`` False, ``matches('10', 10)`` True — text that
    reads as a number equals it — and ``matches('10', '>5')`` False."""
    return _matcher(*_parse(criteria))(value)


# ─── The arguments ──────────────────────────────────────────────────


@dataclass
class _Range:
    values: list[Any]
    node: Expr
    text: str
    unit: Any
    value_type: str | None
    #: The range read once, when the range is one (see ``_shared``).
    shared: _Lines | None = None


@dataclass
class _Criterion:
    test: Callable[[Any], bool]
    node: Expr
    text: str
    #: The error a criteria cell holds, which is then the result.
    error: str | None = None
    #: ``(operator, comparand, text)`` the test is built from (see
    #: ``_matcher``), which an index of the range answers; None for none.
    parsed: tuple[str, Any, str | None] | None = None


def _label(var: Variable) -> str:
    """How a line in a list is named in ``_source_code``: the record it
    belongs to and its own name (``a.amount``), one element of a row as
    ``a.flow[0]``, else its name."""
    return _label_of(var)[0]


# The name, and whether it is read off the line's own path alone — a path
# the line keeps (``_qualified_id``), replaced whenever the line moves or
# is named (one element of a row is named by the row's).
def _label_of(var: Variable) -> tuple[str, bool]:
    parts = str(var.path).split('.')
    if len(parts) >= 2 and parts[0] != '__floating__' and var.python_name:
        return '.'.join(parts[-2:]), var._qualified_id is not None
    inner = getattr(var, '_expr', None)
    if isinstance(inner, Subscript) and isinstance(inner.base, VarRef):
        return f"{_label(inner.base.var)}[{inner.key!r}]", False
    return src(var), var._qualified_id is not None


def _series_value(raw: Any) -> bool:
    return isinstance(raw, list) or hasattr(raw, 'roles')


def _typed_in(var: Variable) -> bool:
    """Does the line hold values typed in, not a formula? One element of
    a row (``flow[0]``) is the row's cell, so the row answers."""
    inner = var._expr
    while isinstance(inner, Subscript) and isinstance(inner.base, VarRef):
        var = inner.base.var
        inner = var._expr
    return not var.is_formula


def _cells(var: Variable, raw: Any) -> list[Any]:
    """The cells of a line as the workbook holds them: empty text typed
    in is written as nothing, so it is a blank; a formula's empty text
    stays a cell."""
    cells = list(raw) if isinstance(raw, list) else [raw]
    if _typed_in(var):
        return [None if isinstance(c, str) and c == '' else c for c in cells]
    return cells


def _range(arg: Any, where: str) -> _Range:
    """A range argument: a Variable (a row, or one cell) or a list of
    scalar lines; its cells' values, its AST node, its name in the
    source, the unit and value type its own lines share."""
    shared = _shared(arg)
    if shared is not None:
        return shared.cells(where)
    if isinstance(arg, (list, tuple)):
        if not arg:
            raise ValueError(f"{where}(): a range needs at least one cell; got [].")
        values: list[Any] = []
        items: list[Expr] = []
        labels: list[str] = []
        for item in arg:
            if isinstance(item, Variable):
                reject_axised(item, where)
                raw = item._value
                if _series_value(raw):
                    raise _series_in_range(where, item)
                values.extend(_cells(item, raw))
                items.append(VarRef(item))
                labels.append(_label(item))
            else:
                values.append(item)
                items.append(Literal(item))
                labels.append(repr(item))
        lines = [item for item in arg if isinstance(item, Variable)]
        return _Range(values, ListExpr(items), f"[{', '.join(labels)}]",
                      shared_unit(lines), lines[0].value_type if lines else None)
    raise TypeError(
        f"{where}(): a range is a list-valued line (a row) or a list of scalar "
        f"lines; got {type(arg).__name__}."
    )


def _tracked_range(where: str) -> ValueError:
    return ValueError(
        f"{where}() over a line with tracks reduces per coordinate — "
        f"slice a coordinate first (.at(track='...'))."
    )


def _series_in_range(where: str, item: Variable) -> ValueError:
    return ValueError(
        f"{where}(): a range is a row or a list of scalar lines; "
        f"{_label(item)} in the list is a series. Pass the row "
        f"itself, or one element of it ({_label(item)}[0])."
    )


# ─── A range read once ──────────────────────────────────────────────

# A row, or a list of lines, is read once and shared by every call over
# the same lines — the same objects, in the same order: one AST node that
# every call's formula holds, so a renderer writes the range once, and
# what the calls read of the lines (their cells, their names in the
# source, the unit they share) and derive from the cells (the numbers
# sorted, the cells indexed by value), each read or derived once and kept
# with the node. Each call first checks that every line still holds the
# value it held when read — the same object, and of a list each cell the
# same object, the list as long — and the same unit and path, and reads
# the lines again when one does not: the engine replaces a line's value
# when it computes it again (a loop closing, a schedule laid out on
# adoption), and a list can be changed in place (the caller's own list a
# line was given, the one ``value`` hands out). A list holding a value
# typed in among its lines is read by each call, as it is short.

_VALUE = operator.attrgetter('_value')
_UNIT = operator.attrgetter('_unit')
_PATH = operator.attrgetter('_qualified_id')

#: The ranges read, by the identities of their lines, each a weak
#: reference to its node; what was read lives on the node and goes with it.
_ROWS: dict[int, Any] = {}
_LISTS: dict[tuple[int, ...], Any] = {}


def _holds_empty_text(raw: Any) -> bool:
    if isinstance(raw, list):
        return any(isinstance(c, str) and c == '' for c in raw)
    return isinstance(raw, str) and raw == ''


class _Lines:
    __slots__ = ('lines', 'node', 'row', 'raws', 'lists', 'sizes', 'flat', 'units', 'rooted',
                 'paths', 'named', 'blanks', 'first_type', 'text', 'series', 'kept')

    def __init__(self, lines: tuple[Variable, ...], node: Expr, row: bool) -> None:
        self.lines = lines
        self.node = node
        self.row = row
        self._read()

    # A copy or a pickle of the node takes none of it: the lines are read
    # again where it is used.
    def __reduce__(self) -> tuple[Any, ...]:
        return (type(None), ())

    def _read(self) -> None:
        lines = self.lines
        self.raws = tuple(map(_VALUE, lines))
        # The lists among the values — their lengths and their cells as read.
        self.lists = tuple(raw for raw in self.raws if isinstance(raw, list))
        self.sizes = tuple(map(len, self.lists))
        self.flat = tuple(chain.from_iterable(self.lists))
        self.units = tuple(map(_UNIT, lines))
        self.first_type = lines[0].value_type
        rooted: list[Variable] = []
        named: list[tuple[Variable, str]] = []
        labels: list[str] = []
        for var in lines:
            label, by_path = ((src(var), var._qualified_id is not None) if self.row
                              else _label_of(var))
            labels.append(label)
            if by_path:
                rooted.append(var)
            else:
                named.append((var, label))
        self.rooted = tuple(rooted)
        self.paths = tuple(map(_PATH, rooted))
        self.named = tuple(named)
        self.text = labels[0] if self.row else f"[{', '.join(labels)}]"
        self.blanks = tuple((var, _typed_in(var)) for var, raw in zip(lines, self.raws)
                            if _holds_empty_text(raw))
        self.series = any(map(_series_value, self.raws))
        self.kept: dict[Any, Any] = {}

    def _unchanged(self) -> bool:
        lines = self.lines
        if not (all(map(operator.is_, map(_VALUE, lines), self.raws))
                and tuple(map(len, self.lists)) == self.sizes
                and all(map(operator.is_, chain.from_iterable(self.lists), self.flat))
                and all(map(operator.is_, map(_UNIT, lines), self.units))
                and all(map(operator.is_, map(_PATH, self.rooted), self.paths))
                and lines[0].value_type == self.first_type):
            return False
        name = src if self.row else _label
        return (all(name(var) == label for var, label in self.named)
                and all(_typed_in(var) is typed for var, typed in self.blanks))

    def check(self) -> None:
        if not self._unchanged():
            self._read()

    # What ``make`` derives from the lines as they are now, derived once.
    def keep(self, key: Any, make: Callable[[], Any]) -> Any:
        try:
            return self.kept[key]
        except KeyError:
            value = self.kept[key] = make()
            return value

    # The range as ``_range`` reads it, refused as it refuses.
    def cells(self, where: str) -> _Range:
        rng = self.kept.get('cells')
        if rng is None:
            rng = self.kept['cells'] = self._cells(where)
        return rng

    def _cells(self, where: str) -> _Range:
        if self.row:
            var = self.lines[0]
            reject_axised(var, where)
            raw = self.raws[0]
            if hasattr(raw, 'roles'):
                raise _tracked_range(where)
            return _Range(_cells(var, raw), self.node, self.text, self.units[0],
                          self.first_type, self)
        values: list[Any] = []
        for var, raw in zip(self.lines, self.raws):
            reject_axised(var, where)
            if _series_value(raw):
                raise _series_in_range(where, var)
            values.extend(_cells(var, raw))
        return _Range(values, self.node, self.text, shared_unit(self.lines), self.first_type,
                      self)

    # The first line of each unit among the lines, with its node: what a
    # check of their units reads — the first unit, the first other unit and
    # each label of a pure number are first among them.
    def unit_lines(self) -> tuple[list[Variable], list[Expr]]:
        from ..core.unit import unit_key

        def make() -> tuple[list[Variable], list[Expr]]:
            nodes: list[Expr] = (list(self.node.items) if isinstance(self.node, ListExpr)
                                 else [self.node])
            seen: set[Any] = set()
            lines: list[Variable] = []
            picked: list[Expr] = []
            for var, unit, node in zip(self.lines, self.units, nodes):
                if unit is None:
                    continue
                key = unit_key(unit)
                if key not in seen:
                    seen.add(key)
                    lines.append(var)
                    picked.append(node)
            return lines, picked
        result: tuple[list[Variable], list[Expr]] = self.keep('units', make)
        return result


# The range a row or a list of lines is, read once (see above); None for
# anything else, read by each call.
def _shared(arg: Any) -> _Lines | None:
    if isinstance(arg, Variable):
        return _find(_ROWS, id(arg), (arg,), True)
    if isinstance(arg, (list, tuple)) and arg:
        return _find(_LISTS, tuple(map(id, arg)), arg, False)
    return None


def _find(table: dict[Any, Any], key: Any, arg: Sequence[Any], row: bool) -> _Lines | None:
    ref = table.get(key)
    node = ref() if ref is not None else None
    if node is not None:
        shared: _Lines = node.__dict__['_lines']
        shared.check()
        return shared
    if not row and not all(isinstance(item, Variable) for item in arg):
        return None
    lines = tuple(arg)
    node = VarRef(lines[0]) if row else ListExpr([VarRef(var) for var in lines])
    shared = _Lines(lines, node, row)
    node.__dict__['_lines'] = shared
    table[key] = weakref.ref(node, partial(_forget, table, key))
    return shared


def _forget(table: dict[Any, Any], key: Any, ref: Any) -> None:
    if table.get(key) is ref:
        del table[key]


#: The ranges a call has just checked, by node, while it builds its result:
#: the unit law run then reads their lines by their own units (the only
#: reading a result's law makes), from what they were read as.
_CHECKED: dict[int, _Lines] = {}


# The ranges among a call's argument nodes — each checked by the call that
# holds it — taken as checked while the call builds its result.
@contextmanager
def _checked(nodes: Sequence[Any]) -> Iterator[None]:
    added: list[int] = []
    for node in nodes:
        shared = getattr(node, '__dict__', {}).get('_lines')
        if shared is not None and id(node) not in _CHECKED:
            _CHECKED[id(node)] = shared
            added.append(id(node))
    try:
        yield
    finally:
        for key in added:
            _CHECKED.pop(key, None)


# A range's lines as a unit law reads them: a row is one, a list of lines
# each of its items — or, for a range the call has just checked, read by
# the lines' own units, the first line of each unit among them, over which
# every unit check answers the same.
def _unit_items(node: Any, unit_of: Callable[[Any], Any]) -> list[Any]:
    shared = _CHECKED.get(id(node))
    if shared is not None and shared.node is node and unit_of is operand_unit:
        return list(shared.unit_lines()[1])
    return list(node.items) if isinstance(node, ListExpr) else [node]


# ─── The cells indexed by value ─────────────────────────────────────

# A range's cells by value, for an equality — a criteria's and an exact
# lookup's: the cells that can meet the value asked for, found without
# reading every cell, each then read by the test itself — so the index
# narrows the reading and never decides. A number or a date is filed by
# its serial (text that reads as a number by that number too), in order,
# so the cells within a hair of a value — the spreadsheet's equality of
# two numbers — are one slice; text by its case folded, and by its letters
# folded as the spreadsheet's text pattern compares them, one by one (text
# with a letter that folds otherwise is read for every value); a boolean
# by itself.

#: Twice the spreadsheet's 2**-48: a slice this wide holds every number
#: equal to the value.
_FUZZ = 2.0 ** -47

#: How a character folds, as the spreadsheet's text pattern compares it
#: (see ``_fold``).
_FOLDED: dict[str, str | None] = {}


# A text as its letters fold for a pattern without wildcards — two texts a
# pattern of one matches fold alike — or None when a letter does not fold
# to one: its lower case, upper case, lower case again (the letters that
# share an upper case, ``i`` and ``ı``, fold to one). An ASCII letter's is
# its lower case.
def _fold(text: str) -> str | None:
    if text.isascii():
        return text.lower()
    out: list[str] = []
    for ch in text:
        if ch in _FOLDED:
            folded = _FOLDED[ch]
        else:
            low = ch.lower()
            up = low.upper() if len(low) == 1 else ''
            key = up.lower() if len(up) == 1 else ''
            folded = _FOLDED[ch] = key if len(key) == 1 else None
        if folded is None:
            return None
        out.append(folded)
    return ''.join(out)


class _Index:
    __slots__ = ('serials', 'at', 'odd', 'bools', 'folded', 'keyed', 'loose', 'exact')

    def __init__(self, values: list[Any]) -> None:
        numbers: list[tuple[Any, int]] = []
        self.odd: list[int] = []
        self.bools: dict[bool, list[int]] = {True: [], False: []}
        self.folded: dict[str, list[int]] = {}
        self.keyed: dict[str, list[int]] = {}
        self.loose: list[int] = []
        #: The numbers and dates by serial, as a lookup compares them.
        self.exact: dict[Any, list[int]] = {}
        for i, v in enumerate(values):
            if isinstance(v, bool):
                self.bools[v].append(i)
                continue
            number = _cell_number(v)
            if number is not None:
                self.exact.setdefault(number, []).append(i)
            elif _is_text(v):
                self.folded.setdefault(v.casefold(), []).append(i)
                key = _fold(v)
                if key is None:
                    self.loose.append(i)
                else:
                    self.keyed.setdefault(key, []).append(i)
                number = _as_number(v)
            if number is None:
                continue
            if isinstance(number, float) and not math.isfinite(number):
                self.odd.append(i)
            else:
                numbers.append((number, i))
        numbers.sort(key=operator.itemgetter(0))
        self.serials = [n for n, _i in numbers]
        self.at = [i for _n, i in numbers]

    # The cells a number or a date can equal: those within a hair of it.
    def near(self, value: Any) -> list[int] | None:
        target = _to_serial(value)
        try:
            fuzz = abs(target) * _FUZZ
            lo, hi = target - fuzz, target + fuzz
        except OverflowError:
            return None
        if not (math.isfinite(lo) and math.isfinite(hi)):
            return None
        return (self.at[bisect.bisect_left(self.serials, lo):
                        bisect.bisect_right(self.serials, hi)] + self.odd)

    # The text cells a pattern of ``text`` without wildcards can match.
    def like(self, text: str) -> list[int] | None:
        if any(ch in '*?~' for ch in text):
            return None
        key = _fold(text)
        return None if key is None else self.keyed.get(key, []) + self.loose

    # Whether ``meeting`` finds the cells ``op comparand`` can meet: an
    # equality, to anything but a blank, an error or a pattern.
    @staticmethod
    def answers(op: str, comparand: Any) -> bool:
        if op != '=' or comparand is None:
            return False
        return not isinstance(comparand, str) or not (
            comparand == '' or is_error(comparand) or any(ch in '*?~' for ch in comparand))

    # The cells ``op comparand`` can meet, given as ``text`` (see
    # ``_matcher``); None when every cell is read.
    def meeting(self, op: str, comparand: Any, text: str | None) -> list[int] | None:
        if not self.answers(op, comparand):
            return None
        same = self.folded.get(text.casefold(), []) if text is not None else []
        if isinstance(comparand, bool):
            return self.bools[comparand] + same
        if isinstance(comparand, str):
            return self.like(comparand)
        near = self.near(comparand)
        return None if near is None else near + same


# The index of a range read once, for an equality it answers — built for
# the second such equality over the range: the first reads every cell, as
# one alone over a range does.
def _index(rng: _Range) -> _Index | None:
    shared = rng.shared
    if shared is None:
        return None
    kept = shared.kept
    index: _Index | None = kept.get('index')
    if index is None:
        if not kept.get('asked'):
            kept['asked'] = True
            return None
        index = kept['index'] = _Index(rng.values)
    return index


def _scalar_criteria(var: Variable, where: str) -> Any:
    reject_axised(var, where)
    raw = var._value
    if _series_value(raw):
        raise ValueError(
            f"{where}(): a criteria is one value — a number, text, or a scalar "
            f"line; {src(var)} is a series. A conditional sum per period is a "
            f"call over that period's cells, with this period's value as the "
            f"criteria."
        )
    return _cells(var, raw)[0]


def _error_value(raw: Any) -> str | None:
    """The error a cell holds, else None (text that only starts with ``#``
    is text)."""
    return raw if isinstance(raw, str) and raw in _ERRORS else None


def _timed(value: Any) -> bool:
    """A datetime with a time of day, which ``DATE(...)`` cannot write."""
    return isinstance(value, datetime) and value.time() != time()


def _never(value: Any) -> bool:
    return False


_Parsed = tuple[str, Any, str | None]


def _test(criteria: Any, where: str,
          line: Variable | None = None) -> tuple[Callable[[Any], bool], _Parsed]:
    """The matcher of one criteria, with what it was read as; an entry
    read as a number in some locales only is refused, naming the call (and
    the line holding it)."""
    try:
        parsed = _parse(criteria)
        return _matcher(*parsed), parsed
    except ValueError as exc:
        held = f" It is the value of {src(line)}." if line is not None else ""
        raise ValueError(f"{where}(): {exc}{held}") from None


def _joined(op: str, raw: Any, where: str,
            line: Variable) -> tuple[Callable[[Any], bool], _Parsed]:
    """``op & cell``, joined as the spreadsheet joins them: a blank cell or
    empty text leaves the operator alone, a boolean is ``TRUE`` /
    ``FALSE``, text is read after the operator, and a number or a date
    is its value."""
    if raw is None or (isinstance(raw, str) and raw == ''):
        return _matcher(op, None), (op, None, None)
    if isinstance(raw, bool):
        raw = 'TRUE' if raw else 'FALSE'
    if isinstance(raw, str):
        return _test(op + raw, where, line)
    return _matcher(op, raw), (op, raw, None)


def _entry_text(value: Any) -> str:
    """A typed value joined to an operator, as the joined text spells it."""
    if value is None:
        return ''
    if isinstance(value, bool):
        return 'TRUE' if value else 'FALSE'
    if isinstance(value, datetime):
        return repr(_to_serial(value))
    return str(value)


def _criterion(criteria: Any, where: str) -> _Criterion:
    """A criteria argument: a number, text, date or boolean (a literal); a
    scalar Variable (its cell); ``(operator, value)`` (the operator joined
    to the value, ``">=" & B2``). A cell holding an error makes that
    error the result."""
    if isinstance(criteria, Variable):
        raw = _scalar_criteria(criteria, where)
        node: Expr = VarRef(criteria)
        text = src(criteria)
        error = _error_value(raw)
        if error is not None:
            return _Criterion(_never, node, text, error)
        test, parsed = _test(raw, where, criteria)
        return _Criterion(test, node, text, parsed=parsed)
    if isinstance(criteria, tuple):
        if len(criteria) != 2 or criteria[0] not in _OPERATORS:
            raise ValueError(
                f"{where}(): a criteria pair is (operator, value) with the operator "
                f"one of {', '.join(repr(op) for op in _OPERATORS)}; got {criteria!r}."
            )
        op, value = criteria
        if isinstance(value, Variable):
            raw = _scalar_criteria(value, where)
            node = FuncCall('CONCAT', [Literal(op), VarRef(value)])
            text = f"CONCAT({op!r}, {src(value)})"
            error = _error_value(raw)
            if error is not None:
                return _Criterion(_never, node, text, error)
            test, parsed = _joined(op, raw, where, value)
            return _Criterion(test, node, text, parsed=parsed)
        if isinstance(value, (date, timedelta)) and not _timed(value):
            node = FuncCall('CONCAT', [Literal(op), Literal(value)])
            return _Criterion(_matcher(op, value), node, f"CONCAT({op!r}, {value!r})",
                              parsed=(op, value, None))
        joined = f"{op}{_entry_text(value)}"
        test, parsed = _test(joined, where)
        return _Criterion(test, Literal(joined), repr(joined), parsed=parsed)
    if isinstance(criteria, (list, dict, set)):
        raise ValueError(
            f"{where}(): a criteria is one value — a number, text, a scalar line "
            f"or (operator, value); got {type(criteria).__name__}."
        )
    # A time of day is written as the serial: ``DATE(...)`` would drop it.
    written = _to_serial(criteria) if _timed(criteria) else criteria
    test, parsed = _test(criteria, where)
    return _Criterion(test, Literal(written), repr(criteria), parsed=parsed)


def _same_length(where: str, first: _Range, other: _Range) -> None:
    if len(first.values) != len(other.values):
        raise ValueError(
            f"{where}(): every range has the same number of cells — {first.text} "
            f"has {len(first.values)}, {other.text} has {len(other.values)}."
        )


def _pairs(args: tuple[Any, ...], where: str, example: str) -> list[tuple[Any, Any]]:
    if len(args) < 2 or len(args) % 2:
        raise ValueError(
            f"{where}() takes criteria ranges and criteria in pairs: {example}. "
            f"Got {len(args)} argument{'s' if len(args) != 1 else ''} after the first."
        )
    return [(args[i], args[i + 1]) for i in range(0, len(args), 2)]


def _mask(pairs: list[tuple[Any, Any]], where: str, like: _Range | None,
          nodes: list[Expr], texts: list[str]
          ) -> tuple[list[tuple[_Range, _Criterion]], str | None]:
    """Every ``(range, criteria)`` pair read, and the error of the first
    criteria cell holding one; the ranges must be as long as ``like`` (the
    aggregated range) and as each other."""
    read: list[tuple[_Range, _Criterion]] = []
    error: str | None = None
    for rng_arg, crit_arg in pairs:
        rng = _range(rng_arg, where)
        if like is not None:
            _same_length(where, like, rng)
        crit = _criterion(crit_arg, where)
        error = error or crit.error
        read.append((rng, crit))
        nodes.extend([rng.node, crit.node])
        texts.extend([rng.text, crit.text])
        like = like or rng
    return read, error


# The positions every pair admits, in order: the cells an index finds for
# a pair are read by its test, and every other pair's test reads those
# alone.
def _admitted(read: list[tuple[_Range, _Criterion]]) -> list[int]:
    found: list[int] | None = None
    rest: list[tuple[_Range, _Criterion]] = []
    for rng, crit in read:
        parsed = crit.parsed
        index = (_index(rng) if parsed is not None and _Index.answers(parsed[0], parsed[1])
                 else None)
        near = index.meeting(*parsed) if index is not None and parsed is not None else None
        if near is None:
            rest.append((rng, crit))
            continue
        values, test = rng.values, crit.test
        own = [i for i in sorted(set(near)) if test(values[i])]
        if found is None:
            found = own
        else:
            kept = set(own)
            found = [i for i in found if i in kept]
    if found is None:
        rng, crit = rest.pop(0)
        values, test = rng.values, crit.test
        found = [i for i, v in enumerate(values) if test(v)]
    for rng, crit in rest:
        values, test = rng.values, crit.test
        found = [i for i in found if test(values[i])]
    return found


def _picked(values: list[Any], at: list[int]) -> list[Any]:
    return [values[i] for i in at]


def _numbers(values: list[Any]) -> list[Any]:
    """The cells a sum, an average or a ranking reads: numbers, dates
    and errors; a blank, text or boolean is skipped."""
    return [v for v in values
            if is_error(v) or isinstance(v, (date, timedelta))
            or (isinstance(v, (int, float)) and not isinstance(v, bool))]


def _serials(values: list[Any]) -> list[Any]:
    return [_to_serial(v) for v in values]


# A date reduces as its serial; the largest and the smallest come back
# as the cell holds them (a date stays a date).
_sum = _excel_reduce(lambda values: builtins.sum(_serials(values)))
_average = _excel_reduce(lambda values: _mean(_serials(values)), '#DIV/0!')
_largest = _excel_reduce(lambda values: builtins.max(values, key=_to_serial))
_smallest = _excel_reduce(lambda values: builtins.min(values, key=_to_serial))


def _mean_of(cells: list[Any]) -> Any:
    """The mean of the picked cells; of dates, the mean date."""
    value = _average(cells)
    if cells and isinstance(value, (int, float)) and all(isinstance(c, date) for c in cells):
        return _from_serial(value)
    return value


def _mean_type(value: Any) -> str:
    return 'datetime' if isinstance(value, date) else 'float'


def _summed_type(rng: _Range) -> str:
    """A sum is a number: of its range's type when that is a number's, a
    float otherwise (a sum of dates is a number of days)."""
    return rng.value_type if rng.value_type in ('int', 'float') else 'float'


def _ranked_type(rng: _Range) -> str:
    """The largest or the smallest is one of the range's numbers or
    dates: of its range's type when that is a number's or a date's, a
    float otherwise."""
    return rng.value_type if rng.value_type in ('int', 'float', 'datetime') else 'float'


def _result(func: str, nodes: list[Expr], texts: list[str], value: Any,
            value_type: str) -> Variable:
    """One number with its formula; its unit is the function's law (see
    Units below)."""
    with _checked(nodes):
        return make_func_var(
            func, nodes, value, value_type, 'scalar',
            source_code=f"{func}({', '.join(texts)})",
        )


# ─── The functions ──────────────────────────────────────────────────


def SUMIF(range: RangeArg, criteria: Any, sum_range: RangeArg | None = None) -> Variable:
    """Sum of the cells of ``sum_range`` (else ``range``) where ``range``
    meets ``criteria`` (renders ``=SUMIF(...)``).

    Example::

        north = mo.SUMIF(sales.region, "North", sales.amount)   # =SUMIF(B2:F2, "North", B3:F3)
        big = mo.SUMIF(sales.amount, ">=1000")                  # =SUMIF(B3:F3, ">=1000")
        over = mo.SUMIF(sales.amount, (">=", inputs.floor))     # =SUMIF(B3:F3, ">=" & Inputs!B1)
        per_deal = mo.SUMIF([d.a.region, d.b.region], "North",  # =SUMIF(Deals!B3:B4, "North",
                            [d.a.amount, d.b.amount])           #         Deals!C3:C4)
    """
    where = "SUMIF"
    rng = _range(range, where)
    crit = _criterion(criteria, where)
    agg = rng if sum_range is None else _range(sum_range, where)
    _same_length(where, rng, agg)
    value = crit.error or _sum(_numbers(_picked(agg.values, _admitted([(rng, crit)]))))
    nodes = [rng.node, crit.node] + ([] if sum_range is None else [agg.node])
    texts = [rng.text, crit.text] + ([] if sum_range is None else [agg.text])
    return _result(where, nodes, texts, value, _summed_type(agg))


def SUMIFS(sum_range: RangeArg, *criteria: Any) -> Variable:
    """Sum of the cells of ``sum_range`` where every ``(range, criteria)``
    pair is met (renders ``=SUMIFS(...)``).

    Example::

        mo.SUMIFS(sales.amount, sales.region, "North", sales.year, 2026)
        # =SUMIFS(B3:F3, B2:F2, "North", B4:F4, 2026)
    """
    where = "SUMIFS"
    agg = _range(sum_range, where)
    nodes: list[Expr] = [agg.node]
    texts = [agg.text]
    pairs = _pairs(criteria, where, "mo.SUMIFS(sum_range, range1, criteria1, range2, criteria2)")
    read, error = _mask(pairs, where, agg, nodes, texts)
    value = error or _sum(_numbers(_picked(agg.values, _admitted(read))))
    return _result(where, nodes, texts, value, _summed_type(agg))


def COUNTIF(range: RangeArg, criteria: Any) -> Variable:
    """Number of cells of ``range`` that meet ``criteria`` (renders
    ``=COUNTIF(...)``): ``mo.COUNTIF(sales.region, "N*")``."""
    where = "COUNTIF"
    rng = _range(range, where)
    crit = _criterion(criteria, where)
    value = crit.error or len(_admitted([(rng, crit)]))
    return _result(where, [rng.node, crit.node], [rng.text, crit.text], value, 'int')


def COUNTIFS(*criteria: Any) -> Variable:
    """Number of positions where every ``(range, criteria)`` pair is met
    (renders ``=COUNTIFS(...)``): ``mo.COUNTIFS(region, "North", amount, ">0")``."""
    where = "COUNTIFS"
    nodes: list[Expr] = []
    texts: list[str] = []
    pairs = _pairs(criteria, where, "mo.COUNTIFS(range1, criteria1, range2, criteria2)")
    read, error = _mask(pairs, where, None, nodes, texts)
    value = error or len(_admitted(read))
    return _result(where, nodes, texts, value, 'int')


def AVERAGEIF(range: RangeArg, criteria: Any,
              average_range: RangeArg | None = None) -> Variable:
    """Mean of the cells of ``average_range`` (else ``range``) where
    ``range`` meets ``criteria`` (renders ``=AVERAGEIF(...)``); of dates,
    the mean date; ``#DIV/0!`` when no number is picked."""
    where = "AVERAGEIF"
    rng = _range(range, where)
    crit = _criterion(criteria, where)
    agg = rng if average_range is None else _range(average_range, where)
    _same_length(where, rng, agg)
    value = crit.error or _mean_of(_numbers(_picked(agg.values, _admitted([(rng, crit)]))))
    nodes = [rng.node, crit.node] + ([] if average_range is None else [agg.node])
    texts = [rng.text, crit.text] + ([] if average_range is None else [agg.text])
    return _result(where, nodes, texts, value, _mean_type(value))


def AVERAGEIFS(average_range: RangeArg, *criteria: Any) -> Variable:
    """Mean of the cells of ``average_range`` where every ``(range,
    criteria)`` pair is met (renders ``=AVERAGEIFS(...)``); ``#DIV/0!``
    when no number is picked."""
    where = "AVERAGEIFS"
    agg = _range(average_range, where)
    nodes: list[Expr] = [agg.node]
    texts = [agg.text]
    pairs = _pairs(criteria, where,
                   "mo.AVERAGEIFS(average_range, range1, criteria1, range2, criteria2)")
    read, error = _mask(pairs, where, agg, nodes, texts)
    value = error or _mean_of(_numbers(_picked(agg.values, _admitted(read))))
    return _result(where, nodes, texts, value, _mean_type(value))


def MAXIFS(max_range: RangeArg, *criteria: Any) -> Variable:
    """Largest cell of ``max_range`` where every ``(range, criteria)`` pair
    is met (renders ``=MAXIFS(...)``); 0 when none is."""
    where = "MAXIFS"
    agg = _range(max_range, where)
    nodes: list[Expr] = [agg.node]
    texts = [agg.text]
    pairs = _pairs(criteria, where, "mo.MAXIFS(max_range, range1, criteria1, range2, criteria2)")
    read, error = _mask(pairs, where, agg, nodes, texts)
    value = error or _largest(_numbers(_picked(agg.values, _admitted(read))))
    return _result(where, nodes, texts, value, _ranked_type(agg))


def MINIFS(min_range: RangeArg, *criteria: Any) -> Variable:
    """Smallest cell of ``min_range`` where every ``(range, criteria)`` pair
    is met (renders ``=MINIFS(...)``); 0 when none is."""
    where = "MINIFS"
    agg = _range(min_range, where)
    nodes: list[Expr] = [agg.node]
    texts = [agg.text]
    pairs = _pairs(criteria, where, "mo.MINIFS(min_range, range1, criteria1, range2, criteria2)")
    read, error = _mask(pairs, where, agg, nodes, texts)
    value = error or _smallest(_numbers(_picked(agg.values, _admitted(read))))
    return _result(where, nodes, texts, value, _ranked_type(agg))


# ─── The counts ─────────────────────────────────────────────────────


def _count_operands(args: tuple[Any, ...], where: str) -> tuple[list[tuple[Any, bool]],
                                                                 list[Expr], list[str]]:
    """``(value, typed)`` per cell — ``typed`` for a value given in the
    call itself, not read from a cell or a list (an array) — with the
    AST nodes and names."""
    if not args:
        raise ValueError(
            f"{where}() needs at least one argument. "
            f"Syntax: mo.{where}(a, b, c)  or  mo.{where}(row). Got 0 arguments."
        )
    cells: list[tuple[Any, bool]] = []
    nodes: list[Expr] = []
    texts: list[str] = []
    for a in args:
        if isinstance(a, (Variable, list, tuple)):
            rng = _range(a, where)
            cells.extend((v, False) for v in rng.values)
            nodes.append(rng.node)
            texts.append(rng.text)
        else:
            cells.append((a, True))
            nodes.append(Literal(a))
            texts.append(repr(a))
    return cells, nodes, texts


def _counts_as_number(value: Any, typed: bool) -> bool:
    """``COUNT``: a number, a date or a duration in a cell; typed into the
    call, a boolean, text that reads as a number and ``None`` (written as
    the 0 it stands for) count too; never an error."""
    if isinstance(value, bool) or value is None:
        return typed
    if isinstance(value, (int, float, date, timedelta)):
        return True
    return typed and isinstance(value, str) and _as_number(value) is not None


# Is a count one call per period: as many values as the registry's
# per_period_arity or more — each line of a list of lines one value — a
# series among them (a line holding a list or a row per track, a list
# typed in)?
def _counts_by_period(args: tuple[Any, ...], where: str) -> bool:
    count, shaped = 0, False
    for a in args:
        if isinstance(a, (list, tuple)) and any(isinstance(i, Variable) for i in a):
            count += len(a)
            shaped = shaped or any(_series_value(i._value) for i in a if isinstance(i, Variable))
        else:
            count += 1
            shaped = shaped or (_series_value(a._value) if isinstance(a, Variable)
                                else isinstance(a, list) and bool(a))
    return shaped and count >= (per_period_arity(where) or 2)


# A line's cells as the workbook holds them (``_cells``), one per period,
# or one per period of each track.
def _line_cells(var: Variable) -> Any:
    from ..core.tracks import TrackValues

    def fiber(raw: Any) -> Any:
        cells = _cells(var, raw)
        return cells if isinstance(raw, list) else cells[0]

    raw = var._value
    return TrackValues.lift(fiber, raw) if isinstance(raw, TrackValues) else fiber(raw)


# A count over two or more values with a series among them: one call per
# period over this period's cells, as ``a + b`` combines its operands — a
# line's cell of the period, a value typed in (a list typed in is one per
# period), the cells of a list of scalar lines in every period; scalars
# and one-value series broadcast, other lengths are refused, located
# lines on other windows are aligned by date (each continued by its own
# extend= rule, the cells the sheet holds for it), a line with tracks is
# counted track by track, lines along the same keys key by key.
def _count_each_period(where: str, args: tuple[Any, ...],
                       counts: Callable[[Any, bool], bool]) -> Variable:
    from ..core.tracks import TrackValues
    raws: list[Any] = []
    typed: list[bool] = []
    nodes: list[Expr] = []
    texts: list[str] = []
    for a in args:
        if isinstance(a, Variable):
            reject_axised(a, where)
            raws.append(_line_cells(a))
            typed.append(False)
            nodes.append(VarRef(a))
            texts.append(src(a))
        elif isinstance(a, tuple) or (isinstance(a, list)
                                      and any(isinstance(i, Variable) for i in a)):
            rng = _range(a, where)
            raws.append(tuple(rng.values))
            typed.append(False)
            nodes.append(rng.node)
            texts.append(rng.text)
        elif isinstance(a, list):
            if not a:
                raise ValueError(f"{where}(): a range needs at least one cell; got [].")
            raws.append(list(a))
            typed.append(True)
            nodes.append(Literal(list(a)))
            texts.append(repr(a))
        else:
            raws.append(a)
            typed.append(True)
            nodes.append(Literal(a))
            texts.append(repr(a))

    def one(values: list[Any]) -> int:
        return builtins.sum(
            1 for value, t in zip(values, typed)
            for cell in (value if isinstance(value, tuple) else (value,)) if counts(cell, t)
        )

    raws, keys = align_keyed(where, args, raws)
    raws, window = align_located(args, raws)
    value: Any
    if any(isinstance(raw, TrackValues) for raw in raws):
        value = TrackValues.lift(lambda *fibers: _per_period(list(fibers), one), *raws)
        var_type = 'list' if value.time_length else 'scalar'
    else:
        value = _per_period(raws, one)
        var_type = 'list' if isinstance(value, list) else 'scalar'
    with _checked(nodes):
        out = make_func_var(where, nodes, value, 'int', var_type,
                            source_code=f"{where}({', '.join(texts)})")
    return placed(out, window, keys)


def _counta_counts(value: Any, typed: bool) -> bool:
    return typed or value is not None


def COUNT(*args: Any) -> Variable:
    """Number of cells holding a number (renders ``=COUNT(...)``): a date
    is one, a blank, text, a boolean or an error is not; a boolean,
    numeric text or ``None`` (the 0 the formula holds) typed into the
    call itself counts, as the spreadsheet counts it. Over ONE operand,
    the count of the row's cells: ``mo.COUNT(row)`` → ``=COUNT(B2:F2)``.
    Over two or more operands with a series among them, the count of
    each period's cells: ``mo.COUNT(plan, actual)`` → ``=COUNT(B2, B3)`` in
    every column. Over scalars only, one number."""
    _refuse_listed_series("COUNT", args)
    if _counts_by_period(args, "COUNT"):
        return _count_each_period("COUNT", args, _counts_as_number)
    cells, nodes, texts = _count_operands(args, "COUNT")
    value = builtins.sum(1 for v, typed in cells if _counts_as_number(v, typed))
    return _result("COUNT", nodes, texts, value, 'int')


def COUNTA(*args: Any) -> Variable:
    """Number of cells that are not blank (renders ``=COUNTA(...)``): text,
    numbers, booleans, errors and a formula's empty text all count, and
    a value typed into the call always does; empty text typed into a
    line is a blank cell. Over ONE operand, the row's cells; over two or
    more operands with a series among them, each period's cells
    (``=COUNTA(B2, B3)`` in every column); over scalars only, one
    number."""
    _refuse_listed_series("COUNTA", args)
    if _counts_by_period(args, "COUNTA"):
        return _count_each_period("COUNTA", args, _counta_counts)
    cells, nodes, texts = _count_operands(args, "COUNTA")
    value = builtins.sum(1 for v, typed in cells if _counta_counts(v, typed))
    return _result("COUNTA", nodes, texts, value, 'int')


# ─── Units ──────────────────────────────────────────────────────────


def _lines_of(node: Any) -> list[Any]:
    """A range's lines: a row is one, a list of lines each of its items —
    in the call that has just checked them, the first line of each unit
    among them (see ``_unit_items``)."""
    return _unit_items(node, operand_unit)


def _criteria_unit(node: Any, unit_of: Callable[[Any], Any]) -> tuple[Any, Any]:
    """The unit a criteria compares in, and the node carrying it: a line's
    own; for an operator joined to a cell — ``('>=', cap)``, or a line
    that is ``mo.CONCAT('>=', cap)`` — the cell's. A typed one has none."""
    if isinstance(node, VarRef):
        unit = unit_of(node)
        if unit is not None:
            return unit, node
        inner = getattr(node.var, '_expr', None)
        if not (isinstance(inner, FuncCall) and inner.func == 'CONCAT'):
            return None, None
        node = inner
    if isinstance(node, FuncCall) and node.func == 'CONCAT':
        for arg in node.args:
            unit = unit_of(arg)
            if unit is not None:
                return unit, arg
    return None, None


def joined_criteria(node: Any) -> tuple[str, Any] | None:
    """``(operator, value node)`` when ``node`` is an operator joined to a
    value - the ``CONCAT('>=', cap)`` a ``('>=', cap)`` criteria is written
    as - else None. A loop container runs the call again from its formula
    and hands the pair back, so the criteria is read as the author gave it."""
    if (isinstance(node, FuncCall) and node.func == 'CONCAT' and len(node.args) == 2
            and isinstance(node.args[0], Literal) and node.args[0].value in _OPERATORS):
        return node.args[0].value, node.args[1]
    return None


def _named(node: Any) -> str:
    from ..core.variable_ops import _line_name
    return _line_name(node.var) if isinstance(node, VarRef) else "a formula"


def _conditional_unit(func: str, total: Any, pairs: list[tuple[Any, Any]],
                      unit_of: Callable[[Any], Any]) -> Any:
    """Each criteria compared in its range's unit, then the unit the
    totalled lines share (none for a count)."""
    from ..core.variable_ops import _incompatible_units
    for rng, criteria in pairs:
        unit, held = _criteria_unit(criteria, unit_of)
        if unit is None or unit is UNSETTLED:
            continue
        for line in _lines_of(rng):
            own = unit_of(line)
            if own is not None and own is not UNSETTLED and own != unit:
                detail = _incompatible_units(own, unit).split(": ", 1)[1]
                raise ValueError(
                    f"mo.{func} compares {_named(line)} with the criteria "
                    f"{_named(held)}; they are in {detail}"
                )
    return None if total is None else same_unit(func, _lines_of(total), unit_of)


def _one_criteria_law(func: str) -> UnitLaw:
    """``SUMIF`` / ``AVERAGEIF`` (range, criteria, the range totalled —
    else the range itself); ``COUNTIF`` (range, criteria), a count."""
    def law(args: Any, unit_of: Callable[[Any], Any]) -> Any:
        if len(args) < 2:
            return None
        total = None if func == "COUNTIF" else args[2] if len(args) > 2 else args[0]
        return _conditional_unit(func, total, [(args[0], args[1])], unit_of)
    return law


def _criteria_pairs_law(func: str) -> UnitLaw:
    """``SUMIFS`` and its kin (the range totalled, then range and criteria
    pairs); ``COUNTIFS`` (the pairs alone), a count."""
    def law(args: Any, unit_of: Callable[[Any], Any]) -> Any:
        if not args:
            return None
        total, pairs = (None, list(args)) if func == "COUNTIFS" else (args[0], list(args[1:]))
        return _conditional_unit(func, total, list(zip(pairs[0::2], pairs[1::2])), unit_of)
    return law


register("SUMIF", shape="range")
register("SUMIFS", shape="range")
register("COUNTIF", shape="range")
register("COUNTIFS", shape="range")
register("AVERAGEIF", shape="range")
register("AVERAGEIFS", shape="range")
register("MAXIFS", shape="range", file_prefix="_xlfn.")
register("MINIFS", shape="range", file_prefix="_xlfn.")
register("COUNTA", shape="reduction", per_period_arity=2)
unit_law("SUMIF", _one_criteria_law("SUMIF"))
unit_law("AVERAGEIF", _one_criteria_law("AVERAGEIF"))
unit_law("COUNTIF", _one_criteria_law("COUNTIF"))
unit_law("SUMIFS", _criteria_pairs_law("SUMIFS"))
unit_law("AVERAGEIFS", _criteria_pairs_law("AVERAGEIFS"))
unit_law("MAXIFS", _criteria_pairs_law("MAXIFS"))
unit_law("MINIFS", _criteria_pairs_law("MINIFS"))
unit_law("COUNTIFS", _criteria_pairs_law("COUNTIFS"))
# ``COUNT`` is registered beside ``SUM`` (functions/aggregate.py).
