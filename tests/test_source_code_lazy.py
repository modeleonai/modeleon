# SPDX-License-Identifier: Apache-2.0
"""The source text of a condition is its formula, rendered when someone
reads it — not at every call.

``IF``, ``AND``, ``OR``, ``NOT``, ``CHOOSE`` and the logical helpers say
"the readable source is the formula"; rendering that text at construction
inlines every anonymous operand, and nested conditions render the inner
text twice, for a value almost nobody reads. ``Variable._source_code`` is
therefore derived on read: a text written on the instance wins, else the
formula (the condition family), else the one text the Variables of a list
share (a ``Variable([a, b, ...])`` compound), else None.
"""

import copy
import pickle

import pytest

import modeleon as mo
from modeleon.core.expr import FuncCall
from modeleon.functions.logical import IFERROR, IFS, ISERROR, ISNUMBER, SWITCH, XOR


def _model():
    m = mo.Model('m', default_grain='month', default_start='2026-01', default_periods=3)
    m.x = mo.MultiVariable('X')
    m.x.a = mo.Variable([1.0, 5.0, 9.0])
    m.x.b = mo.Variable([2.0, 2.0, 2.0])
    return m


@pytest.fixture
def renders(monkeypatch):
    """Count the renders of a call node — the whole-expression render a
    condition's source text costs."""
    count = {'n': 0}
    original = FuncCall.to_string

    def counting(self):
        count['n'] += 1
        return original(self)

    monkeypatch.setattr(FuncCall, 'to_string', counting)
    return count


CONDITIONS = {
    'IF': lambda m: mo.IF(m.x.a > 4.0, m.x.a, 0.0),
    'AND': lambda m: mo.AND(m.x.a > 4.0, m.x.b > 1.0),
    'OR': lambda m: mo.OR(m.x.a > 4.0, False),
    'NOT': lambda m: mo.NOT(m.x.a > 4.0),
    'CHOOSE': lambda m: mo.CHOOSE(2, m.x.a, m.x.b),
    'IFERROR': lambda m: IFERROR(m.x.a / m.x.b, 0.0),
    'ISERROR': lambda m: ISERROR(m.x.a),
    'ISNUMBER': lambda m: ISNUMBER(m.x.a),
    'XOR': lambda m: XOR(m.x.a > 4.0, True),
    'IFS': lambda m: IFS(m.x.a > 7.0, 'big', True, 'small'),
    'SWITCH': lambda m: SWITCH(m.x.a, 1.0, 'one', 'other'),
}


class TestTheSourceIsTheFormulaRenderedOnRead:
    def test_an_if_renders_nothing_until_read(self, renders):
        m = _model()
        out = mo.IF(m.x.a > 4.0, m.x.a, 0.0)
        assert renders['n'] == 0
        assert '_source_code' not in vars(out)
        assert out.source_code == "IF((a > 4.0), a, 0.0)"
        assert renders['n'] >= 1

    @pytest.mark.parametrize('name', sorted(CONDITIONS))
    def test_every_reader_gets_the_formula(self, name, renders):
        m = _model()
        out = CONDITIONS[name](m)
        assert renders['n'] == 0, name
        text = out.formula
        assert text and text.startswith(f"{name}(")
        # The public property, the internal attribute and the guarded read
        # the renderers use all return the same string.
        assert out.source_code == text
        assert out._source_code == text
        assert getattr(out, '_source_code', None) == text
        assert hasattr(out, '_source_code')

    def test_a_nested_condition_renders_the_inner_text_once_per_read(self, renders):
        m = _model()
        inner = mo.IF(m.x.a > 4.0, 1.0, 0.0)
        outer = mo.IF(m.x.b > 1.0, inner, 2.0)
        assert renders['n'] == 0
        assert outer.source_code == "IF((b > 1.0), (IF((a > 4.0), 1.0, 0.0)), 2.0)"
        assert renders['n'] == 2  # the outer call and the inlined inner one

    def test_a_written_text_wins_over_the_formula(self):
        m = _model()
        out = mo.IF(m.x.a > 4.0, m.x.a, 0.0)
        out.source_code = "mine(a)"
        assert out.source_code == out._source_code == "mine(a)"
        # The ``None`` an adopted typed-dates line writes clears the text.
        out._source_code = None
        assert out.source_code is None and out._source_code is None

    def test_a_resolved_call_text_stays_eager(self):
        # Helpers that resolve scalars into the text (``PMT(0.005, ...)``)
        # still write it at the call; a plain input has none.
        r = mo.PMT(0.005, 360, 300000)
        assert '_source_code' in vars(r)
        assert r.source_code == "PMT(0.005, 360, 300000, 0, 0)"
        assert mo.Variable(1.0).source_code is None

    def test_copies_carry_the_text(self):
        m = _model()
        out = mo.IF(m.x.a > 4.0, m.x.a, 0.0)
        assert mo.Variable(out).source_code == out.formula
        assert mo.Variable(formula=out).source_code == out.formula
        assert pickle.loads(pickle.dumps(out)).source_code == out.formula
        assert copy.deepcopy(out).source_code == out.formula


class TestACompoundSharesItsItemsText:
    def test_the_one_text_the_items_share_decided_on_read(self, renders):
        m = _model()
        out = mo.IF(m.x.a > 4.0, m.x.a, 0.0)
        same = mo.Variable([out, out])
        with_scalars = mo.Variable([out, 1.0, out])
        assert renders['n'] == 0
        assert same.source_code == out.formula
        assert with_scalars.source_code == out.formula
        assert '_source_code' not in vars(same)

    def test_different_or_missing_texts_give_none(self):
        m = _model()
        out = mo.IF(m.x.a > 4.0, m.x.a, 0.0)
        other = mo.IF(m.x.a > 1.0, 1.0, 0.0)
        assert mo.Variable([out, other]).source_code is None
        assert mo.Variable([out, m.x.a]).source_code is None
        assert mo.Variable([m.x.a, m.x.b]).source_code is None
        assert mo.Variable([1.0, 2.0]).source_code is None

    def test_eager_texts_share_too(self):
        a = mo.PMT(0.005, 360, 300000)
        b = mo.PMT(0.005, 360, 300000)
        assert mo.Variable([a, b]).source_code == "PMT(0.005, 360, 300000, 0, 0)"
        assert mo.Variable([a, mo.PMT(0.01, 360, 300000)]).source_code is None


def test_a_loop_row_listing_itself_shares_no_text():
    """A row of a loop container whose line lists the row itself is
    rewritten in place, so its items reach the row again. The model is
    refused; reading the row's text afterwards must answer None, not
    recurse."""
    m = mo.Model('m', default_grain='month', default_start='2026-01', default_periods=3)
    m.d = mo.MultiVariable('D', loop=True)
    with pytest.raises(Exception):
        with m.d:
            placeholder = m.d.row
            m.d.row = mo.Variable([placeholder, 2.0, 3.0])
    assert m.d.row._source_code is None
    assert mo.Variable(m.d.row).source_code is None
