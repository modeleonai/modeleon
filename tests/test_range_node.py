# SPDX-License-Identifier: Apache-2.0
"""A range is read once.

Every call over the same lines — the rank of each record among its peers,
a ``SUMIFS`` per record over the records' fields, a ``MATCH`` per record —
reads the range once and shares it: one AST node that every call's formula
holds, the cells, their names in the source, the unit they share and what
is derived from them (the numbers sorted, the cells indexed by value). The
workbook writes the range once per sheet. A call reads the lines again when
one holds another value — replaced, or a list changed in place — or
another unit or path than it held when read.

The scale tests count the work, never the clock: a table twice as long
costs twice the work, not four times.
"""

import copy
import datetime as dt
import gc
import pickle

import pytest
from openpyxl import load_workbook

import modeleon as mo
from modeleon.compile.excel.renderer import ExcelRenderer
from modeleon.core.expr import ListExpr
from modeleon.functions import criteria, lookup, statistics


class Comp(mo.MultiVariableClass):
    def compute(self, multiple, rev=None):
        self.multiple = mo.Variable(multiple, display_name="Multiple", unit="x")
        if rev is not None:
            self.rev = mo.Variable(rev, display_name="Rev", unit="USD")


class Tx(mo.MultiVariableClass):
    def compute(self, region, amount, day):
        self.region = mo.Variable(region, display_name="Region")
        self.amount = mo.Variable(amount, display_name="Amount", unit="USD")
        self.day = mo.Variable(day, display_name="Day")


def _model():
    return mo.Model("m", default_grain="year", default_start="2026", default_periods=3)


def _comps(m, n):
    m.comps = mo.MultiVariable("Comps", excel_props={"tab": True, "orient": "records"})
    with m.comps as c:
        for k in range(n):
            setattr(c, f"r{k}", Comp(multiple=float(k * 7 % n), rev=[k, 2.0 * k, n - k],
                                     display_name=f"R{k}"))
    return [m.comps._component(f"r{k}") for k in range(n)]


def _txs(m, n, groups):
    m.tx = mo.MultiVariable("Tx", excel_props={"tab": True, "orient": "records"})
    with m.tx as t:
        for k in range(n):
            setattr(t, f"x{k}", Tx(region=f"G{k % groups}", amount=float(k % 13),
                                   day=dt.date(2026, 1, 1 + k % 28), display_name=f"x{k}"))
    return [m.tx._component(f"x{k}") for k in range(n)]


def _read_by_each_call(monkeypatch):
    """From here on every call reads its range itself, as before a range
    was read once: the answers to compare with."""
    monkeypatch.setattr(criteria, "_shared", lambda arg: None)
    monkeypatch.setattr(statistics, "_shared", lambda arg: None)


@pytest.fixture
def counted(monkeypatch):
    """The work, by kind: the cells read off a line, the cells a statistic
    counts, the cells a criteria or a lookup tests, the cells a renderer
    looks up for a range, the lists of lines walked for their lines."""
    work = {"cells": 0, "counted": 0, "tested": 0, "rendered": 0, "walked": 0}

    real_cells = criteria._cells

    def cells(var, raw):
        work["cells"] += 1
        return real_cells(var, raw)

    real_counted = statistics._counted

    def counted_(cells_):
        work["counted"] += len(cells_)
        return real_counted(cells_)

    def testing(real):
        def make(*args, **kwargs):
            test = real(*args, **kwargs)

            def tested(v):
                work["tested"] += 1
                return test(v)
            return tested
        return make

    real_cell_of = ExcelRenderer._cell_of

    def cell_of(self, *args, **kwargs):
        work["rendered"] += 1
        return real_cell_of(self, *args, **kwargs)

    monkeypatch.setattr(criteria, "_cells", cells)
    monkeypatch.setattr(statistics, "_counted", counted_)
    monkeypatch.setattr(criteria, "_matcher", testing(criteria._matcher))
    monkeypatch.setattr(lookup, "_exact_test", testing(lookup._exact_test))
    real_walk = ListExpr.iter_refs

    def walked(self):
        work["walked"] += 1
        return real_walk(self)

    monkeypatch.setattr(ExcelRenderer, "_cell_of", cell_of)
    monkeypatch.setattr(ListExpr, "iter_refs", walked)
    return work


def _rows(path, sheet):
    return [[c.value for c in row] for row in load_workbook(path)[sheet].iter_rows()]


# ══════════════════════════════════════════════════════════════════════
# The work grows with the table, not with its square
# ══════════════════════════════════════════════════════════════════════


class TestScale:
    def _ranked(self, counted, n, tmp_path):
        m = _model()
        with m:
            recs = _comps(m, n)
            peers = [r.multiple for r in recs]
            for k in counted:
                counted[k] = 0
            with m.comps:
                for r in recs:
                    r.rank = mo.RANK_EQ(r.multiple, peers).set_display_name("Rank")
            built = dict(counted)
            counted["rendered"] = counted["walked"] = 0
            m.to_excel(tmp_path / f"rank{n}.xlsx")
        return built, (counted["rendered"], counted["walked"])

    def test_a_rank_per_record_reads_and_writes_its_range_once(self, counted, tmp_path):
        (small, written_small), (large, written_large) = (
            self._ranked(counted, 60, tmp_path), self._ranked(counted, 120, tmp_path))
        # The range is read once: its cells once, counted and sorted once.
        assert (small["cells"], large["cells"]) == (60, 120)
        assert (small["counted"], large["counted"]) == (60, 120)
        # Written once per sheet: each peer's cell looked up once, and the
        # list walked once for what its lines read.
        assert (written_small, written_large) == ((60, 1), (120, 1))

    def test_a_rank_per_record_and_year_sorts_each_year_once(self, counted):
        def ranked(n):
            m = _model()
            with m:
                recs = _comps(m, n)
                revs = [r.rev for r in recs]
                counted["counted"] = 0
                with m.comps:
                    for r in recs:
                        r.rank = mo.RANK_EQ(r.rev, revs).set_display_name("Rank")
            return counted["counted"]

        assert (ranked(40), ranked(80)) == (3 * 40, 3 * 80)

    def test_a_sumifs_per_record_tests_only_the_cells_of_its_key(self, counted, tmp_path):
        def summed(n):
            m = _model()
            with m:
                recs = _txs(m, n, groups=n // 4)
                amounts, regions = [r.amount for r in recs], [r.region for r in recs]
                counted["tested"] = counted["cells"] = 0
                with m.tx:
                    for r in recs:
                        r.total = mo.SUMIFS(amounts, regions, r.region).set_display_name("Tot")
                built = dict(counted)
                counted["rendered"] = 0
                m.to_excel(tmp_path / f"sumifs{n}.xlsx")
            return built, counted["rendered"]

        (small, written_small), (large, written_large) = summed(60), summed(120)
        # The first call tests every cell, as an equality alone does; each
        # after tests the four cells of its key. Each line is read once (its
        # region once more, as the criteria of its own call).
        assert (small["tested"], large["tested"]) == (60 + 4 * 59, 120 + 4 * 119)
        assert (small["cells"], large["cells"]) == (3 * 60, 3 * 120)
        assert (written_small, written_large) == (2 * 60, 2 * 120)

    def test_a_match_per_record_tests_only_the_cells_of_its_key(self, counted):
        def matched(n):
            m = _model()
            with m:
                recs = _txs(m, n, groups=n)
                regions = [r.region for r in recs]
                counted["tested"] = 0
                with m.tx:
                    for r in recs:
                        r.at = mo.MATCH(r.region, regions, 0).set_display_name("At")
                        r.back = mo.XLOOKUP(r.region, regions, regions).set_display_name("Back")
                assert [r.at.value for r in recs] == list(range(1, n + 1))
            return counted["tested"]

        # The first call tests every cell; each after, the one of its key.
        assert (matched(60), matched(120)) == (60 + 2 * 60 - 1, 120 + 2 * 120 - 1)

    def test_an_equality_alone_reads_every_cell_and_a_second_indexes_them(self, counted):
        lines = [mo.Variable(f"t{k}") for k in range(40)]
        counted["tested"] = 0
        first = mo.COUNTIF(lines, "t7")
        assert counted["tested"] == 40
        shared = criteria._shared(lines)
        assert "index" not in shared.kept
        pattern = mo.COUNTIF(lines, "t1*")       # a pattern, which no index answers
        found = mo.MATCH("t1?", lines, 0)
        assert "index" not in shared.kept
        counted["tested"] = 0
        second = mo.MATCH("T9", lines, 0)
        assert "index" in shared.kept
        assert counted["tested"] == 1
        assert (first.value, pattern.value, found.value, second.value) == (1, 11, 11, 10)

    def test_a_median_per_record_is_worked_out_once(self, counted):
        def mid(n):
            m = _model()
            with m:
                recs = _comps(m, n)
                peers = [r.multiple for r in recs]
                counted["counted"] = 0
                with m.comps:
                    for r in recs:
                        r.gap = (r.multiple - mo.MEDIAN(peers)).set_display_name("Gap")
            return counted["counted"]

        assert (mid(60), mid(120)) == (60, 120)


# ══════════════════════════════════════════════════════════════════════
# One node, shared by every call over the same lines
# ══════════════════════════════════════════════════════════════════════


class TestOneNode:
    def test_the_calls_hold_one_node_and_a_new_list_of_the_same_lines_finds_it(self):
        m = _model()
        with m:
            recs = _comps(m, 5)
            peers = [r.multiple for r in recs]
            with m.comps:
                for r in recs:
                    r.rank = mo.RANK_EQ(r.multiple, list(peers))
            nodes = {id(r.rank._expr.args[1]) for r in recs}
            assert len(nodes) == 1
            again = mo.PERCENTILE_INC(tuple(peers), 0.5)
            assert again._expr.args[0] is recs[0].rank._expr.args[1]
            other = mo.RANK_EQ(recs[0].multiple, peers[::-1])
            assert other._expr.args[1] is not again._expr.args[0]

    def test_a_copy_or_a_pickle_takes_the_formula_and_reads_the_lines_again(self):
        m = _model()
        with m:
            recs = _comps(m, 4)
            peers = [r.multiple for r in recs]
            with m.comps:
                for r in recs:
                    r.rank = mo.RANK_EQ(r.multiple, peers)
        for twin in (copy.deepcopy(recs[1].rank), pickle.loads(pickle.dumps(recs[1].rank))):
            node = twin._expr.args[1]
            assert node.__dict__.get("_lines") is None
            assert [item.var.value for item in node.items] == [0.0, 3.0, 2.0, 1.0]
            assert twin.value == recs[1].rank.value == 1

    def test_a_range_no_formula_holds_any_more_is_forgotten(self):
        m = _model()
        with m:
            recs = _comps(m, 4)
            key = tuple(id(r.multiple) for r in recs)
            rank = mo.RANK_EQ(recs[0].multiple, [r.multiple for r in recs])
            assert key in criteria._LISTS
            del rank
            gc.collect()
            assert key not in criteria._LISTS


class TestTheWriterWalksAListOnce:
    def test_a_line_off_the_window_in_a_shared_list_is_named_for_each_formula(
            self, tmp_path):
        m = _model()
        # Read by date, as ``a + b`` reads it: a year later, no rule to
        # continue it — refused when the call is built.
        late = mo.Variable([5.0, 6.0, 7.0], start="2027", grain="year")
        # Two years from the sheet's first, continued by zero: read by date,
        # but a line the book does not write, whose values would stand in
        # the sheet's columns — refused by the export, for each formula.
        odd = mo.Variable([5.0, 6.0], start="2026", grain="year", extend=mo.zero())
        with m:
            m.s = mo.MultiVariable("S")
            m.s.a = mo.Variable([1.0, 2.0, 3.0])
            m.s.b = mo.Variable([3.0, 1.0, 2.0])
            with pytest.raises(ValueError, match="does not cover the combined window"):
                mo.RANK_EQ(m.s.a, [m.s.a, m.s.b, late])
            lines = [m.s.a, m.s.b, odd]
            m.s.r1 = mo.RANK_EQ(m.s.a, lines)
            m.s.r2 = mo.RANK_EQ(m.s.b, lines)
        with pytest.raises(ValueError, match="time window of its own") as refused:
            m.to_excel(tmp_path / "odd.xlsx")
        assert "(m.s.r1) reads" in str(refused.value)
        assert "(m.s.r2) reads" in str(refused.value)

    def test_a_list_met_again_on_its_own_way_down_is_not_taken_as_walked(self):
        from modeleon.compile.excel.writer import _reads_unwritten_off_window
        from modeleon.core.expr import FuncCall, VarRef

        a = mo.Variable()
        inner = ListExpr([VarRef(a)])
        a._set_expr(FuncCall("MEDIAN", [inner]))
        outer = ListExpr([VarRef(a)])
        clean = {}
        assert _reads_unwritten_off_window(FuncCall("F", [outer]), None, {}, set(),
                                           clean) is None
        assert clean == {}
        b = mo.Variable()
        b._set_expr(FuncCall("MEDIAN", [ListExpr([VarRef(mo.Variable(1.0))])]))
        shared = ListExpr([VarRef(b)])
        assert _reads_unwritten_off_window(FuncCall("F", [shared]), None, {}, set(),
                                           clean) is None
        assert shared in clean.values()


# ══════════════════════════════════════════════════════════════════════
# A line that changed is read again
# ══════════════════════════════════════════════════════════════════════


class TestReadAgain:
    def test_a_value_replaced_is_read_again(self):
        m = _model()
        with m:
            recs = _comps(m, 4)                     # 0, 3, 2, 1
            peers = [r.multiple for r in recs]
            assert mo.RANK_EQ(recs[1].multiple, peers).value == 1
            # As a loop closing, or a schedule laid out on adoption, writes it.
            recs[0].multiple._value = 9.0
            assert mo.RANK_EQ(recs[1].multiple, peers).value == 2
            assert mo.LARGE(peers, 1).value == 9.0
            assert mo.SUMIF(peers, 9.0).value == 9.0
            assert mo.MATCH(9.0, peers, 0).value == 1

    def test_a_row_changed_in_place_is_read_again(self):
        m = _model()
        with m:
            m.s = mo.MultiVariable("S")
            data = [1, 8.0, 9.0]
            m.s.u = mo.Variable(data)               # the line holds the caller's list
            m.s.k = mo.Variable(8.0)
            assert mo.MEDIAN(m.s.u).value == 8.0
            assert mo.COUNTIF(m.s.u, 8.0).value == 1
            assert mo.RANK_EQ(m.s.k, m.s.u).value == 2
            data[1] = 100.0
            assert mo.MEDIAN(m.s.u).value == 9.0
            assert mo.COUNTIF(m.s.u, 8.0).value == 0
            assert mo.RANK_EQ(m.s.k, m.s.u).value == "#N/A"
            assert mo.MATCH(100.0, m.s.u, 0).value == 2
            assert mo.COUNTIF(m.s.u, 1).value == 1
            m.s.u.value[0] = True                   # equal to 1, yet a boolean
            assert mo.COUNTIF(m.s.u, 1).value == 0
            assert mo.MEDIAN(m.s.u).value == 54.5
            assert mo.LARGE(m.s.u, 3).value == "#NUM!"

    def test_a_series_changed_in_place_is_read_again(self):
        m = _model()
        with m:
            m.s = mo.MultiVariable("S")
            m.s.a = mo.Variable([1.0, 2.0, 3.0])
            m.s.b = mo.Variable([4.0, 5.0, 6.0])
            m.s.c = mo.Variable([7.0, 0.0, 9.0])
            lines = [m.s.a, m.s.b, m.s.c]
            assert mo.RANK_EQ(m.s.a, lines).value == [3, 2, 3]
            assert mo.MEDIAN(lines).value == [4.0, 2.0, 6.0]
            m.s.a.value[0] = 100.0
            assert mo.RANK_EQ(m.s.a, lines).value == [1, 2, 3]
            assert mo.MEDIAN(lines).value == [7.0, 2.0, 6.0]
            # A cell moved from one line to the next: the same cells in the
            # same order, in lines of other lengths — refused as they now
            # are, as ``a + b`` refuses them.
            m.s.b.value.insert(0, m.s.a.value.pop())
            with pytest.raises(ValueError, match="does not cover the combined window"):
                m.s.a + m.s.b
            with pytest.raises(ValueError, match="does not cover the combined window"):
                mo.MEDIAN(lines)

    def test_a_unit_set_is_read_again(self):
        m = _model()
        with m:
            recs = _comps(m, 3)
            peers = [r.multiple for r in recs]
            assert str(mo.MEDIAN(peers)._unit) == "x"
            recs[2].multiple.unit = "USD"
            with pytest.raises(ValueError, match="one unit"):
                mo.MEDIAN(peers)
            recs[2].multiple.unit = "x"
            assert str(mo.LARGE(peers, 1)._unit) == "x"

    def test_a_line_adopted_or_moved_is_named_again(self):
        a, b = mo.Variable(1.0), mo.Variable(2.0)
        first = mo.RANK_EQ(a, [a, b])
        assert first.source_code == f"RANK_EQ({a.path.leaf}, [{a.path.leaf}, {b.path.leaf}])"
        m = _model()
        with m:
            m.s = mo.MultiVariable("S")
            m.s.a, m.s.b = a, b
            second = mo.RANK_EQ(m.s.a, [a, b])
        assert second.source_code == "RANK_EQ(a, [s.a, s.b])"

    def test_a_line_moved_to_another_container_is_named_again(self):
        m = _model()
        with m:
            m.s = mo.MultiVariable("S")
            m.t = mo.MultiVariable("T")
            m.s.a, m.s.b = mo.Variable(1.0), mo.Variable(2.0)
            lines = [m.s.a, m.s.b]
            assert mo.LARGE(lines, 1).source_code == "LARGE([s.a, s.b], 1)"
            line = m.s.remove("a")
            m.t.a = line
            assert mo.LARGE(lines, 1).source_code == "LARGE([t.a, s.b], 1)"

    def test_a_list_changed_is_another_range(self):
        m = _model()
        with m:
            recs = _comps(m, 4)
            peers = [r.multiple for r in recs]
            assert mo.LARGE(peers, 1).value == 3.0
            peers.append(mo.Variable(7.0))
            assert mo.LARGE(peers, 1).value == 7.0
            peers[-1] = mo.Variable(-1.0)
            assert mo.SMALL(peers, 1).value == -1.0

    def test_empty_text_typed_in_then_given_a_formula(self):
        from modeleon.core.expr import Literal
        a, blank = mo.Variable(1.0), mo.Variable("")
        line = [a, blank]
        assert mo.COUNTIF(line, "=").value == 1        # typed in: a blank cell
        blank.expr = Literal("")
        assert mo.COUNTIF(line, "=").value == 0        # a formula's empty text

    def test_the_first_line_typed_again(self):
        a, b = mo.Variable(1), mo.Variable(2)
        assert mo.LARGE([a, b], 1).value_type == "int"
        a.value_type = "float"
        assert mo.LARGE([a, b], 1).value_type == "float"


class TestALoopClosing:
    """A call made while a loop is open reads provisional rows; when the
    loop closes the call runs again over the same lines, now holding their
    values — the range is read again, as a range read by each call is."""

    @staticmethod
    def _model():
        m = mo.Model("m", default_grain="year", default_start="2026", default_periods=4)
        with m:
            m.d = mo.MultiVariable("D", loop=True)
            with m.d as d:
                d.a = mo.Variable([1.0, 5.0, 2.0, 4.0])
                d.c = mo.lag(d.b, fill=0.0) + 1.0
                d.mid = mo.MEDIAN([d.a, d.c])
                d.rank = mo.RANK_EQ(d.a, [d.a, d.c])
                d.top = mo.LARGE([d.a, d.c], 1)
                d.b = d.a + d.c
        return m

    def test_the_rows_are_read_again(self, monkeypatch):
        shared = self._model()
        _read_by_each_call(monkeypatch)
        plain = self._model()
        assert shared.d.c.value == plain.d.c.value == [1.0, 3.0, 9.0, 12.0]
        for name in ("mid", "rank", "top"):
            assert getattr(shared.d, name).value == getattr(plain.d, name).value, name
        assert shared.d.mid.value == [1.0, 4.0, 5.5, 8.0]
        assert shared.d.rank.value == [1, 1, 2, 2]
        assert shared.d.top.value == [1.0, 5.0, 9.0, 12.0]


# ══════════════════════════════════════════════════════════════════════
# An index answers as reading every cell does
# ══════════════════════════════════════════════════════════════════════

_CELLS = [
    10, 10.0, 10 + 1e-15, 10 + 1e-13, -0.0, 0, 0.1 + 0.2, 0.3, 1e300, 1e-300, 3, 2.5,
    dt.date(2026, 1, 31), dt.datetime(2026, 1, 31, 12, 0), dt.timedelta(days=3),
    True, False, "10", "1e1", "10%", " 10 ", "North", "north", "NORTH", "nörth",
    "ı", "I", "i", "İ", "ß", "SS", "ss", "ﬅ", "st", "\u212a", "k", "Σ", "σ", "ς",
    "\u0432", "\u1c80", "\u0412", "TRUE", "true", "x*", "x?y", "~", "#N/A", "#text", "", None,
]
_CRITERIA = [
    10, "10", "=10", 10.0, "1e1", True, "TRUE", "true", False, "north", "NORTH", "i",
    "I", "ı", "İ", "ss", "ß", "st", "k", "σ", "ς", "\u0432", "\u0412", dt.date(2026, 1, 31),
    ("=", dt.date(2026, 1, 31)), "#N/A", "", "=", "<>", ">5", "<>north", "n*", "x~*",
    0.3, 0.1 + 0.2, -0.0, 3, "3", None, 1e300,
]


def _answers(lines, amounts):
    out = []
    for crit in _CRITERIA:
        out.append(("SUMIF", crit, mo.SUMIF(lines, crit, amounts).value))
        out.append(("COUNTIF", crit, mo.COUNTIF(lines, crit).value))
        out.append(("AVERAGEIF", crit, mo.AVERAGEIF(lines, crit, amounts).value))
        out.append(("SUMIFS", crit, mo.SUMIFS(amounts, lines, crit, amounts, ">0").value))
        out.append(("COUNTIFS", crit, mo.COUNTIFS(amounts, ">0", lines, crit).value))
        out.append(("MAXIFS", crit, mo.MAXIFS(amounts, lines, crit).value))
    for key in _CELLS + ["x", "nörth", "NöRTH"]:
        if key is None:
            continue
        line = mo.Variable(key)
        out.append(("MATCH", key, mo.MATCH(line, lines, 0).value))
        for mode in (0, 2):
            for search in (1, -1):
                out.append(("XLOOKUP", (key, mode, search),
                            mo.XLOOKUP(line, lines, amounts, "none",
                                       match_mode=mode, search_mode=search).value))
    return out


def _same(a, b):
    return a == b or (isinstance(a, float) and isinstance(b, float) and a != a and b != b)


class TestTheIndexAnswersAsEveryCell:
    def test_criteria_and_lookups_over_awkward_cells(self, monkeypatch):
        lines = [mo.Variable(v) for v in _CELLS]
        amounts = [mo.Variable(float(k + 1)) for k in range(len(_CELLS))]
        indexed = _answers(lines, amounts)
        _read_by_each_call(monkeypatch)
        read = _answers(lines, amounts)
        assert len(indexed) == len(read)
        wrong = [(i, r) for i, r in zip(indexed, read) if not _same(i[2], r[2])]
        assert not wrong

    def test_a_letter_folds_as_the_text_pattern_compares_it(self):
        """Two letters a pattern of one matches fold alike — every letter
        the pattern engine compares without regard to case."""
        _sre = pytest.importorskip("_sre")
        casefix = pytest.importorskip("re._casefix")
        import sys
        lower = {}
        groups = {}
        for i in range(sys.maxunicode + 1):
            if 0xD800 <= i <= 0xDFFF:
                continue
            low = _sre.unicode_tolower(i)
            lower[i] = low
            groups.setdefault(low, []).append(i)
        for p, low in lower.items():
            if not _sre.unicode_iscased(p):
                continue
            key = criteria._fold(chr(p))
            if key is None:
                continue
            for same in (low, *casefix._EXTRA_CASES.get(low, ())):
                for c in groups.get(same, ()):
                    other = criteria._fold(chr(c))
                    assert other is None or other == key, (chr(p), chr(c))


# ══════════════════════════════════════════════════════════════════════
# The workbook is the one a range read by each call writes
# ══════════════════════════════════════════════════════════════════════


def _records_model():
    m = _model()
    with m:
        recs = _comps(m, 8)
        peers = [r.multiple for r in recs]
        revs = [r.rev for r in recs]
        with m.comps:
            for r in recs:
                r.rank = mo.RANK_EQ(r.multiple, peers).set_display_name("Rank")
                r.avg = mo.RANK_AVG(r.multiple, peers, 1).set_display_name("Avg")
                r.p = mo.PERCENTILE_INC(peers, 0.25).set_display_name("P25")
                r.gap = (r.multiple - mo.MEDIAN(peers)).set_display_name("Gap")
                r.year_rank = mo.RANK_EQ(r.rev, revs).set_display_name("YearRank")
                r.same = mo.COUNTIF(peers, r.multiple).set_display_name("Same")
                r.over = mo.SUMIFS(peers, peers, (">", r.multiple)).set_display_name("Over")
                r.at = mo.MATCH(r.multiple, peers, 0).set_display_name("At")
                r.back = mo.INDEX(peers, r.at).set_display_name("Back")
                r.look = mo.XLOOKUP(r.multiple, peers, peers).set_display_name("Look")
        m.s = mo.MultiVariable("Summary", excel_props={"tab": True})
        m.s.top = mo.LARGE(peers, 2)
        m.s.mid = mo.MEDIAN(revs)
    return m


class TestTheSameWorkbook:
    def test_every_cell(self, monkeypatch, tmp_path):
        _records_model().to_excel(tmp_path / "shared.xlsx")
        _read_by_each_call(monkeypatch)
        _records_model().to_excel(tmp_path / "each.xlsx")
        shared = load_workbook(tmp_path / "shared.xlsx")
        each = load_workbook(tmp_path / "each.xlsx")
        assert shared.sheetnames == each.sheetnames
        for name in shared.sheetnames:
            assert _rows(tmp_path / "shared.xlsx", name) == _rows(tmp_path / "each.xlsx", name)
        rows = _rows(tmp_path / "shared.xlsx", "Comps")
        header, first, last = rows[1], rows[2], rows[9]
        assert first[header.index("Rank")] == "=_xlfn.RANK.EQ(B3, B3:B10)"
        assert last[header.index("Rank")] == "=_xlfn.RANK.EQ(B10, B3:B10)"
        assert first[header.index("Over (x)")] == '=SUMIFS(B3:B10, B3:B10, ">" & B3)'
        assert first[header.index("At")] == "=MATCH(B3, B3:B10, 0)"

    def test_the_source_and_the_units(self, monkeypatch):
        shared = _records_model()
        _read_by_each_call(monkeypatch)
        each = _records_model()
        for a, b in zip(shared.walk(), each.walk()):
            va, vb = a[1], b[1]
            if isinstance(va, mo.Variable):
                assert va.source_code == vb.source_code, a[0]
                assert va.value == vb.value, a[0]
                assert str(va._unit) == str(vb._unit), a[0]
                assert va.value_type == vb.value_type, a[0]
