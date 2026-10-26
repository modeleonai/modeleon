# SPDX-License-Identifier: Apache-2.0
"""The operator fast paths agree with the general rules they shortcut.

Every operator builds its result by writing fields straight into the
result's ``__dict__``, takes two plain numbers through a short cell rule,
and asks for the right operand's location only when the left one has one.
Each shortcut is exact only while an assumption holds; these tests pin
the assumptions and compare each shortcut with the long way round.
"""

from __future__ import annotations

import enum
import inspect
import math
import operator
from datetime import date, datetime, timedelta
from decimal import Decimal
from fractions import Fraction

import pytest

import modeleon as mo
from modeleon.core.expr import BinOp, VarRef
from modeleon.core.variable import Variable
from modeleon.core.variable_ops import (
    _COMPARISONS,
    _VariableArithmetic,
    _date_serial_op,
)


def _general_cell_rule(a, b, op, safe_divide, op_symbol):
    """One cell, the long way: an error marker propagates (the left one
    first), a blank reads as 0 (or "" against text in a comparison), a zero
    divisor is #DIV/0!, a date computes on its serial number, and anything
    that fails is #VALUE!."""
    if isinstance(a, str) and a.startswith('#'):
        return a
    if isinstance(b, str) and b.startswith('#'):
        return b
    if a is None:
        a = '' if isinstance(b, str) and op_symbol in _COMPARISONS else 0
    if b is None:
        b = '' if isinstance(a, str) and op_symbol in _COMPARISONS else 0
    if safe_divide and b == 0:
        return '#DIV/0!'
    if isinstance(a, (date, timedelta)) or isinstance(b, (date, timedelta)):
        return _date_serial_op(a, b, op, op_symbol)
    try:
        return op(a, b)
    except Exception:
        return '#VALUE!'


class _Small(enum.IntEnum):
    ZERO = 0
    ONE = 1


class _Float(float):
    pass


class _Int(int):
    pass


_OPS = {
    '+': operator.add, '-': operator.sub, '*': operator.mul, '/': operator.truediv,
    '//': operator.floordiv, '%': operator.mod, '**': operator.pow,
    '==': operator.eq, '!=': operator.ne, '<': operator.lt, '<=': operator.le,
    '>': operator.gt, '>=': operator.ge,
}

# Plain numbers (the fast path) next to everything that must NOT take it:
# bools, int and float subclasses, other numeric types, blanks, text, error
# markers and dates.
_CELLS = [
    0, 1, -1, 7, 2**63, 0.0, -0.0, 1.5, -2.25, math.inf, -math.inf, math.nan,
    True, False, _Small.ZERO, _Small.ONE, _Float(1.5), _Int(3),
    Decimal('1.5'), Decimal('0'), Fraction(1, 3), complex(1, 1),
    None, '', 'abc', '#DIV/0!', '#N/A',
    date(2026, 1, 31), datetime(2026, 1, 1, 12), timedelta(days=2), timedelta(0),
]


def _huge_power(a, b) -> bool:
    """Whether ``a ** b`` would build an enormous exact number."""
    exact = (int, Decimal, Fraction)
    if isinstance(b, exact) and not isinstance(b, bool) and abs(b) > 64:
        return True
    return isinstance(a, exact) and not isinstance(a, bool) and abs(a) > 10**6


def _same(x, y) -> bool:
    if isinstance(x, float) and isinstance(y, float) and math.isnan(x) and math.isnan(y):
        return True
    if type(x) is int and type(y) is int:
        return x == y  # 7 ** a date's serial has more digits than repr allows
    return type(x) is type(y) and repr(x) == repr(y)


@pytest.mark.parametrize('op_symbol', sorted(_OPS))
@pytest.mark.parametrize('safe_divide', [False, True])
def test_every_cell_matches_the_general_rule(op_symbol, safe_divide):
    op = _OPS[op_symbol]
    for a in _CELLS:
        for b in _CELLS:
            if op_symbol == '**' and _huge_power(a, b):
                continue  # 7 ** 2**63 is a big number, not a test
            got = _VariableArithmetic._broadcast_operation(
                a, b, op, safe_divide=safe_divide, op_symbol=op_symbol)
            want = _general_cell_rule(a, b, op, safe_divide, op_symbol)
            assert _same(got, want), (op_symbol, safe_divide, a, b, got, want)


def test_lists_broadcast_through_the_same_rule():
    left = [1, 2.0, None, '#N/A', 0, True]
    right = [0, 0.0, 3, 4, -0.0, 2]
    got = _VariableArithmetic._broadcast_operation(
        left, right, operator.truediv, safe_divide=True, op_symbol='/')
    want = [_general_cell_rule(a, b, operator.truediv, True, '/')
            for a, b in zip(left, right)]
    assert all(_same(g, w) for g, w in zip(got, want)), (got, want)


def _all_subclasses(cls):
    for sub in cls.__subclasses__():
        yield sub
        yield from _all_subclasses(sub)


def _data_descriptor(cls, name) -> bool:
    try:
        attr = inspect.getattr_static(cls, name)
    except AttributeError:
        return False
    return hasattr(type(attr), '__set__') or hasattr(type(attr), '__delete__')


def test_no_operator_field_is_a_descriptor():
    """A field written straight into ``__dict__`` must be a plain field: a
    property of that name would be read instead of what was written."""
    m = mo.Model('m', default_grain='month', default_start='2026-01', default_periods=3)
    m.s = mo.MultiVariable('S')
    m.s.a = mo.Variable([1.0, 2.0, 3.0])
    m.s.part = mo.Variable([5.0], start='2026-02', grain='month', extend=mo.zero())
    m.s.when = mo.time.start
    results = [m.s.a + m.s.part, m.s.a * 2, 2 - m.s.a, m.s.a > 1, -m.s.a, m.s.a ** 2,
               m.s.when + 3]
    fields = {name for r in results for name in vars(r)}
    assert '_first_day' in fields  # the date-aligned result was built
    for cls in [Variable, *_all_subclasses(Variable)]:
        shadowed = sorted(n for n in fields if _data_descriptor(cls, n))
        assert not shadowed, (cls.__qualname__, shadowed)


def test_no_variable_class_intercepts_attribute_writes():
    """The ``expr`` setter writes its fields straight into ``__dict__`` for
    every Variable class, so none may observe writes on its own."""
    for cls in _all_subclasses(Variable):
        assert cls.__setattr__ is Variable.__setattr__, cls.__qualname__


def test_the_expr_setter_fills_a_subclass_like_a_variable():
    class Line(mo.Variable):
        pass

    m = mo.Model('m', default_grain='month', default_start='2026-01', default_periods=3)
    m.s = mo.MultiVariable('S')
    m.s.a = mo.Variable([1.0, 2.0, 3.0])
    m.s.days = mo.time.days
    node = BinOp('*', VarRef(m.s.a), VarRef(m.s.days))
    plain, sub = mo.Variable(), Line()
    plain._set_expr(node)
    sub.expr = node
    for v in (plain, sub):
        assert v._expr is node
        assert v._dependency_refs == [m.s.a, m.s.days]
        assert v._reads_time is True
        assert v._awaits is None
    sub.expr = None
    assert sub._dependency_refs == [] and sub._reads_time is False


def test_an_unlocated_left_side_never_asks_the_right_side():
    """``_align_by_date`` answers ``None`` for an unlocated left side before
    it looks the right side up — the same answer it gave when it looked."""
    stub = mo.Model('stub', default_grain='quarter', default_start=date(2026, 2, 1),
                    default_periods=2)
    stub.s = mo.MultiVariable('S')
    stub.s.days = mo.time.days
    whole = mo.Model('whole', default_grain='quarter', default_start='2026-Q1',
                     default_periods=2)
    whole.s = mo.MultiVariable('S')
    whole.s.days = mo.time.days
    floating = mo.Variable([1.0, 2.0])
    align = _VariableArithmetic._align_by_date
    assert stub.s.days.time.first_day is not None
    assert floating.time is None
    assert align(floating, stub.s.days, floating._value, stub.s.days._value) is None
    assert align(stub.s.days, floating, stub.s.days._value, floating._value) is None
    assert (floating + stub.s.days).value == [a + b for a, b in
                                             zip([1.0, 2.0], stub.s.days.value)]
    # Both located, first periods on different days: refused either way round.
    for left, right in ((stub.s.days, whole.s.days), (whole.s.days, stub.s.days)):
        with pytest.raises(ValueError, match='do not line up'):
            left + right
