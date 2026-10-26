# SPDX-License-Identifier: Apache-2.0
"""Lookup and reference: a key found in a range or a block as the
spreadsheet finds it, one cell with its formula — a series when the
key is one."""

import datetime as dt

import pytest

import modeleon as mo
from modeleon.core.expr import FuncCall, ListExpr, Literal, VarRef
from modeleon.functions.lookup import (
    HLOOKUP,
    INDEX,
    MATCH,
    VLOOKUP,
    XLOOKUP,
    _exact,
    _nearest,
    _position,
    _sorted_match,
)
from modeleon.functions.registry import (
    admits_in_loop,
    file_prefix,
    is_passthrough,
    is_range_taking,
    range_positions,
    spec,
)

# ══════════════════════════════════════════════════════════════════════
# One cell against the key
# ══════════════════════════════════════════════════════════════════════


class TestTheMatchingCore:
    @pytest.mark.parametrize("cells, key, wildcard, at", [
        ([10, 20, 30], 20, False, 1),
        ([10, 20, 30], 25, False, None),
        ([10, "20", 30], 20, False, None),           # a number never matches text
        (["10", 20], "10", False, 0),
        (["North", "south"], "SOUTH", False, 1),     # text, case-insensitively
        (["Northern", "North"], "North", False, 1),  # the whole cell
        (["Northern", "North"], "North*", True, 0),  # wildcards where Excel reads them
        (["Northern", "North"], "North*", False, None),
        (["a~b", "ab"], "a~b", True, 1),
        (["a~b", "ab"], "a~b", False, 0),
        ([True, 1], True, False, 0),                 # a boolean matches a boolean only
        ([1, True], True, False, 1),
        ([None, 0, ""], None, False, None),          # a blank matches nothing
        ([None, 0], 0, False, 1),
        ([None, ""], "", False, 1),
        (["#N/A", 5], 5, False, 1),                  # an error among the cells is skipped
        (["#N/A", 5], "#N/A", False, None),
        ([dt.date(2026, 1, 1), 46023], dt.date(2026, 1, 1), False, 0),
        ([46023, dt.date(2026, 1, 1)], dt.date(2026, 1, 1), False, 0),  # a date is its serial
        ([dt.datetime(2026, 1, 1, 12)], dt.date(2026, 1, 1), False, None),
    ])
    def test_an_exact_match(self, cells, key, wildcard, at):
        assert _exact(cells, key, wildcard=wildcard) == at

    def test_the_first_or_the_last_among_equals(self):
        assert _exact([2, 1, 2], 2, wildcard=False) == 0
        assert _exact([2, 1, 2], 2, wildcard=False, last=True) == 2

    @pytest.mark.parametrize("cells, key, below, last, at", [
        ([10, 30, 20], 25, True, False, 2),     # the largest not above, in any order
        ([10, 30, 20], 25, False, False, 1),    # the smallest not below
        ([10, 30, 20], 20, True, False, 2),     # an equal cell is nearest of all
        ([10, 30, 20], 5, True, False, None),
        ([10, 30, 20], 35, False, False, None),
        ([20, 10, 20], 25, True, False, 0),     # ties: the first, or the last
        ([20, 10, 20], 25, True, True, 2),
        (["b", "d"], "C", True, False, 0),      # text orders case-insensitively
        (["b", "d"], "C", False, False, 1),
        ([10, "x", None, "#N/A", 20], 15, True, False, 0),   # another kind, a blank, an error
    ])
    def test_the_nearest_on_one_side(self, cells, key, below, last, at):
        assert _nearest(cells, key, below=below, last=last) == at

    @pytest.mark.parametrize("cells, key, ascending, at", [
        ([10, 20, 30], 25, True, 1),
        ([10, 20, 30], 30, True, 2),
        ([10, 20, 20, 30], 20, True, 2),       # the last among equals, as Excel finds it
        ([10, 20, 30], 5, True, None),
        ([30, 20, 10], 25, False, 0),
        ([30, 20, 10], 15, False, 1),
        ([30, 20, 10], 35, False, None),
        ([10, "x", 30], 25, True, 0),          # another kind is skipped
        (["a", "c", "e"], "d", True, 1),
    ])
    def test_over_sorted_cells(self, cells, key, ascending, at):
        assert _sorted_match(cells, key, ascending=ascending) == at

    @pytest.mark.parametrize("n, size, at", [
        (1, 3, 0), (3, 3, 2), (2.9, 3, 1),
        (0, 3, "#VALUE!"), (-1, 3, "#VALUE!"), (None, 3, "#VALUE!"),
        ("2", 3, 1), ("2.9", 3, 1), ("two", 3, "#VALUE!"), ("", 3, "#VALUE!"),
        (True, 3, 0), (False, 3, "#VALUE!"),         # a boolean is 1 or 0
        (dt.date(2026, 1, 1), 3, "#REF!"),           # a date is its serial
        (float("inf"), 3, "#VALUE!"), ("1e999", 3, "#VALUE!"), ("1" * 400, 3, "#REF!"),
        (4, 3, "#REF!"), ("#DIV/0!", 3, "#DIV/0!"),
    ])
    def test_a_position(self, n, size, at):
        assert _position(n, size) == at


# ══════════════════════════════════════════════════════════════════════
# The functions
# ══════════════════════════════════════════════════════════════════════


class Deal(mo.MultiVariableClass):
    def compute(self, region, amount):
        self.region = mo.Variable(region, display_name="Region")
        self.amount = mo.Variable(amount, display_name="Amount", unit="USD")
        self.flow = mo.Variable([amount * 0.1, amount * 0.2, amount * 0.3], display_name="Flow")


def _model():
    m = mo.Model("m", default_grain="month", default_start="2026-01", default_periods=3)
    with m:
        m.i = mo.MultiVariable("I")
        m.i.years = mo.Variable([2026, 2027, 2028], display_name="years")
        m.i.rates = mo.Variable([0.05, 0.04, 0.03], display_name="rates", unit="USD")
        m.i.names = mo.Variable(["Alpha", "Beta", "Gamma"], display_name="names")
        m.i.infl = mo.Variable({2026: 0.05, 2027: 0.04, 2028: 0.03}, display_name="infl")
        m.i.year = mo.Variable([2026, 2028, 2030], display_name="year")
        m.i.pick = mo.Variable(2027, display_name="pick")
        m.deals = mo.MultiVariable("Deals", excel_props={"tab": True, "orient": "records"})
        with m.deals as d:
            d.a = Deal(region="North", amount=100.0, display_name="Alpha")
            d.b = Deal(region="South", amount=200.0, display_name="Beta")
            d.c = Deal(region="North", amount=300.0, display_name="Gamma")
    return m


class TestMATCH:
    def test_exact_sorted_ascending_and_descending(self):
        m = _model()
        assert MATCH(2027, m.i.years, 0).value == 2
        assert MATCH(2027.5, m.i.years).value == 2
        assert MATCH(2027.5, m.i.years, 1).value == 2
        assert MATCH(2025, m.i.years).value == "#N/A"
        desc = mo.Variable([30, 20, 10])
        assert MATCH(25, desc, -1).value == 1
        assert MATCH(15, desc, -1).value == 2
        assert MATCH(35, desc, -1).value == "#N/A"

    def test_text_case_insensitively_with_wildcards_and_never_a_number(self):
        m = _model()
        assert MATCH("beta", m.i.names, 0).value == 2
        assert MATCH("g*", m.i.names, 0).value == 3
        assert MATCH("2027", m.i.years, 0).value == "#N/A"
        assert MATCH(2027, mo.Variable(["2026", "2027"]), 0).value == "#N/A"

    def test_a_blank_key_an_error_key_and_an_error_among_the_cells(self):
        m = _model()
        assert MATCH(mo.Variable(None), m.i.years, 0).value == "#N/A"
        assert MATCH(mo.Variable([2027, None]), m.i.years, 0).value == [2, "#N/A"]
        assert MATCH("#DIV/0!", m.i.years, 0).value == "#DIV/0!"
        assert MATCH(2028, mo.Variable(["#N/A", 2027, 2028]), 0).value == 3

    def test_over_a_list_of_lines_and_a_keyed_line_s_cells(self):
        m = _model()
        d = m.deals
        assert MATCH("Beta", [d.a.region, d.b.region, d.c.region], 0).value == "#N/A"
        assert MATCH("South", [d.a.region, d.b.region, d.c.region], 0).value == 2
        assert MATCH(0.04, m.i.infl, 0).value == 2

    def test_a_series_key_is_one_lookup_per_period(self):
        m = _model()
        out = MATCH(m.i.year, m.i.years, 0)
        assert out.value == [1, 3, "#N/A"] and out.var_type == "list"
        assert MATCH([2028, 2026], m.i.years, 0).value == [3, 1]

    def test_a_position_is_an_int_without_a_unit(self):
        m = _model()
        out = MATCH(0.04, m.i.rates, 0)
        assert out.value == 2 and out.value_type == "int" and out._unit is None
        assert out.var_type == "scalar"

    def test_the_formula_and_the_source(self):
        m = _model()
        out = MATCH(m.i.pick, m.i.years, 0)
        assert isinstance(out._expr, FuncCall) and out._expr.func == "MATCH"
        assert [type(a) for a in out._expr.args] == [VarRef, VarRef, Literal]
        assert out.formula == "MATCH(pick, years, 0)"
        assert out._source_code == "MATCH(pick, years, 0)"
        assert MATCH(m.i.pick, m.i.years).formula == "MATCH(pick, years)"

    def test_the_mode_is_typed_in(self):
        m = _model()
        with pytest.raises(ValueError, match="match_type is one of 1, 0, -1"):
            MATCH(2027, m.i.years, 2)
        with pytest.raises(ValueError, match="match_type"):
            MATCH(2027, m.i.years, m.i.pick)

    def test_none_typed_as_the_key_is_refused(self):
        # It has no cell: the formula would write it as 0 and look 0 up.
        m = _model()
        with pytest.raises(ValueError, match="the key is None, and None has no cell"):
            MATCH(None, m.i.years, 0)
        with pytest.raises(ValueError, match=r"the key \[2027, None\] holds None"):
            MATCH([2027, None], m.i.years, 0)
        for call in (lambda: XLOOKUP(None, m.i.years, m.i.rates),
                     lambda: XLOOKUP(None, m.i.infl),
                     lambda: VLOOKUP(None, m.deals, 3, False),
                     lambda: HLOOKUP(None, [m.i.years, m.i.rates], 2, False)):
            with pytest.raises(ValueError, match="a blank cell finds nothing"):
                call()


class TestINDEX:
    def test_a_position_in_a_row_and_in_a_list_of_lines(self):
        m = _model()
        d = m.deals
        assert INDEX(m.i.rates, 2).value == 0.04
        assert INDEX([d.a.amount, d.b.amount, d.c.amount], 3).value == 300.0
        assert INDEX(m.i.rates, 2.9).value == 0.04

    @pytest.mark.parametrize("k, result", [
        (0, "#VALUE!"), (-1, "#VALUE!"), (None, "#VALUE!"), ("two", "#VALUE!"),
        (False, "#VALUE!"), (4, "#REF!"), ("#N/A", "#N/A"),
    ])
    def test_a_position_the_range_does_not_have(self, k, result):
        assert INDEX(_model().i.rates, k).value == result

    def test_a_position_is_read_as_the_spreadsheet_reads_a_number(self):
        from modeleon.core.unit import unit_key
        m = _model()
        assert INDEX(m.i.rates, "2").value == 0.04
        assert INDEX(m.i.rates, True).value == 0.05
        assert INDEX(m.i.rates, mo.Variable("3")).value == 0.03
        assert INDEX(m.deals, "2", "3").value == 200.0
        assert VLOOKUP("Beta", m.deals, "3", False).value == 200.0
        assert HLOOKUP(2027, [m.i.years, m.i.rates], "2", False).value == 0.04
        # Typed by the column the text names, and in that column's unit.
        out = VLOOKUP("Beta", m.deals, "3", False)
        assert out.value_type == "float"
        assert unit_key(out._unit) == unit_key(m.deals.a.amount._unit)

    def test_the_cell_found_carries_its_own_content(self):
        cells = mo.Variable([1.0, None, "#DIV/0!", "x"])
        assert INDEX(cells, 2).value == 0
        assert INDEX(cells, 3).value == "#DIV/0!"
        assert INDEX(cells, 4).value == "x"

    def test_a_series_position_is_one_pick_per_period(self):
        m = _model()
        out = INDEX(m.i.rates, mo.Variable([1, 3, 5]))
        assert out.value == [0.05, 0.03, "#REF!"] and out.var_type == "list"

    def test_type_and_unit_come_from_the_range(self):
        from modeleon.core.unit import unit_key
        m = _model()
        out = INDEX(m.i.rates, 1)
        assert out.value_type == "float" and unit_key(out._unit) == unit_key(m.i.rates._unit)
        assert INDEX(m.i.years, 1).value_type == "int"
        assert INDEX(m.i.names, 1).value_type == "string"

    def test_over_a_table_of_records_the_label_column_first(self):
        from modeleon.core.unit import unit_key
        m = _model()
        assert INDEX(m.deals, 2, 1).value == "Beta"
        assert INDEX(m.deals, 2, 2).value == "South"
        assert INDEX(m.deals, 3, 3).value == 300.0
        assert INDEX(m.deals, 4, 3).value == "#REF!"
        assert INDEX(m.deals, 1, 4).value == "#REF!"
        assert INDEX(m.deals, 0, 1).value == "#VALUE!"
        assert INDEX(m.deals, 2, 1).value_type == "string"
        amount = INDEX(m.deals, 2, 3)
        assert amount.value_type == "float"
        assert unit_key(amount._unit) == unit_key(m.deals.a.amount._unit)

    def test_over_a_list_of_rows(self):
        m = _model()
        assert INDEX([m.i.years, m.i.rates], 2, 3).value == 0.03
        assert INDEX([m.i.years, m.i.rates], 1, 3).value == 2028
        out = INDEX([m.i.years, m.i.rates], 2, mo.Variable([1, 2, 3]))
        assert out.value == [0.05, 0.04, 0.03] and out.var_type == "list"
        assert INDEX([m.i.years, m.i.rates], 2, 1).value_type == "float"
        assert INDEX([[1, 2], [3, 4]], 2, 1).value == 3

    def test_the_formula_names_the_block(self):
        m = _model()
        out = INDEX(m.deals, 2, 3)
        assert [type(a) for a in out._expr.args] == [ListExpr, Literal, Literal]
        assert all(isinstance(row, ListExpr) for row in out._expr.args[0].items)
        assert out._source_code == "INDEX(deals, 2, 3)"
        rows = INDEX([m.i.years, m.i.rates], 2, 3)
        assert [type(a) for a in rows._expr.args[0].items] == [VarRef, VarRef]
        assert rows._source_code == "INDEX([i.years, i.rates], 2, 3)"

    def test_one_position_over_a_block_and_two_over_a_range_are_refused(self):
        m = _model()
        with pytest.raises(ValueError, match="a row and a column number"):
            INDEX(m.deals, 2)
        with pytest.raises(ValueError, match="one position"):
            INDEX(m.i.rates, 1, 2)


class TestXLOOKUP:
    def test_exact_by_default(self):
        m = _model()
        assert XLOOKUP(2027, m.i.years, m.i.rates).value == 0.04
        assert XLOOKUP("beta", m.i.names, m.i.rates).value == 0.04
        assert XLOOKUP(2030, m.i.years, m.i.rates).value == "#N/A"
        assert XLOOKUP("B*", m.i.names, m.i.rates).value == "#N/A"
        assert XLOOKUP("B*", m.i.names, m.i.rates, match_mode=2).value == 0.04
        assert XLOOKUP(dt.date(2026, 1, 1), mo.Variable([dt.date(2026, 1, 1)]),
                       mo.Variable([7.0])).value == 7.0

    def test_if_not_found_a_value_or_a_cell(self):
        m = _model()
        assert XLOOKUP(2030, m.i.years, m.i.rates, 0).value == 0
        assert XLOOKUP(2030, m.i.years, m.i.rates, "none").value == "none"
        assert XLOOKUP(2030, m.i.years, m.i.rates, m.i.pick).value == 2027
        assert XLOOKUP(2027, m.i.years, m.i.rates, 0).value == 0.04

    def test_the_nearest_below_or_above_in_any_order(self):
        keys = mo.Variable([10, 30, 20])
        vals = mo.Variable(["a", "b", "c"])
        assert XLOOKUP(25, keys, vals, match_mode=-1).value == "c"
        assert XLOOKUP(25, keys, vals, match_mode=1).value == "b"
        assert XLOOKUP(5, keys, vals, match_mode=-1).value == "#N/A"
        assert XLOOKUP(5, keys, vals, 0, -1).value == 0
        dup = mo.Variable([20, 10, 20])
        assert XLOOKUP(20, dup, vals).value == "a"
        assert XLOOKUP(20, dup, vals, search_mode=-1).value == "c"
        assert XLOOKUP(25, dup, vals, match_mode=-1, search_mode=-1).value == "c"

    def test_a_series_key_is_one_lookup_per_period(self):
        m = _model()
        out = XLOOKUP(m.i.year, m.i.years, m.i.rates, 0)
        assert out.value == [0.05, 0.03, 0] and out.var_type == "list"
        assert XLOOKUP(m.i.year, m.i.years, m.i.rates).value == [0.05, 0.03, "#N/A"]

    def test_over_a_keyed_line_the_keys_are_searched(self):
        m = _model()
        assert XLOOKUP(2027, m.i.infl).value == 0.04
        assert XLOOKUP(2029, m.i.infl).value == "#N/A"
        assert XLOOKUP(2029, m.i.infl, if_not_found=0).value == 0
        assert XLOOKUP(2029, m.i.infl, match_mode=-1).value == 0.03
        out = XLOOKUP(m.i.year, m.i.infl, if_not_found=0)
        assert out.value == [0.05, 0.03, 0]
        fy = mo.Variable({"FY2023": 400, "FY2024": 450})
        assert XLOOKUP("fy2024", fy).value == 450
        with pytest.raises(ValueError, match="runs along keys"):
            XLOOKUP(2027, m.i.rates)

    def test_the_keyed_form_types_the_keys_into_the_call(self):
        m = _model()
        out = XLOOKUP(m.i.pick, m.i.infl)
        assert [type(a) for a in out._expr.args] == [VarRef, Literal, VarRef]
        assert out._expr.args[1].value == [2026, 2027, 2028]
        assert out.formula == "XLOOKUP(pick, [2026, 2027, 2028], infl)"
        assert out._source_code == "XLOOKUP(pick, [2026, 2027, 2028], infl)"

    def test_a_mode_without_if_not_found_is_written_after_na(self):
        m = _model()
        out = XLOOKUP(m.i.pick, m.i.years, m.i.rates, match_mode=-1)
        assert out.formula == "XLOOKUP(pick, years, rates, NA(), -1)"
        assert out._source_code == "XLOOKUP(pick, years, rates, match_mode=-1)"
        out = XLOOKUP(m.i.pick, m.i.years, m.i.rates, search_mode=-1)
        assert out.formula == "XLOOKUP(pick, years, rates, NA(), 0, -1)"
        assert out._source_code == "XLOOKUP(pick, years, rates, match_mode=0, search_mode=-1)"
        out = XLOOKUP(m.i.pick, m.i.years, m.i.rates, 0, 1, -1)
        assert out.formula == "XLOOKUP(pick, years, rates, 0, 1, -1)"
        assert XLOOKUP(m.i.pick, m.i.years, m.i.rates).formula == "XLOOKUP(pick, years, rates)"

    def test_type_and_unit_come_from_the_returned_range_never_the_keys(self):
        from modeleon.core.unit import unit_key
        m = _model()
        out = XLOOKUP(2027, m.i.years, m.i.rates)
        assert out.value_type == "float" and unit_key(out._unit) == unit_key(m.i.rates._unit)
        back = XLOOKUP(0.04, m.i.rates, m.i.years)
        assert back.value_type == "int" and back._unit is None
        assert XLOOKUP(2027, m.i.infl).value_type == "float"

    def test_refusals(self):
        m = _model()
        with pytest.raises(ValueError, match="same number of cells"):
            XLOOKUP(2027, m.i.years, mo.Variable([1, 2]))
        with pytest.raises(ValueError, match="match_mode is one of"):
            XLOOKUP(2027, m.i.years, m.i.rates, match_mode=3)
        with pytest.raises(ValueError, match="search_mode is one of 1, -1"):
            XLOOKUP(2027, m.i.years, m.i.rates, search_mode=0)
        for binary in (2, -2):
            with pytest.raises(ValueError, match="binary search .* Use 1 for the first"):
                XLOOKUP(2027, m.i.years, m.i.rates, search_mode=binary)
        with pytest.raises(ValueError, match="search_mode is one of 1, -1, typed in"):
            XLOOKUP(2027, m.i.years, m.i.rates, search_mode=mo.Variable(2))
        with pytest.raises(TypeError, match="a key is one value"):
            XLOOKUP({2027}, m.i.years, m.i.rates)


class TestVLOOKUP:
    def test_over_a_table_of_records_by_label(self):
        m = _model()
        assert VLOOKUP("Beta", m.deals, 3, False).value == 200.0
        assert VLOOKUP("beta", m.deals, 2, False).value == "South"
        assert VLOOKUP("Beta", m.deals, 1, False).value == "Beta"
        assert VLOOKUP("G*", m.deals, 3, False).value == 300.0
        assert VLOOKUP("Delta", m.deals, 3, False).value == "#N/A"

    def test_approximate_by_default_over_a_sorted_first_column(self):
        m = _model()
        assert VLOOKUP("Bob", m.deals, 3).value == 200.0
        assert VLOOKUP("Bob", m.deals, 3, True).value == 200.0
        assert VLOOKUP("Aardvark", m.deals, 3).value == "#N/A"
        assert VLOOKUP("Bob", m.deals, 3, False).value == "#N/A"

    def test_the_column_number(self):
        m = _model()
        assert VLOOKUP("Beta", m.deals, 0, False).value == "#VALUE!"
        assert VLOOKUP("Beta", m.deals, 4, False).value == "#REF!"
        assert VLOOKUP("Beta", m.deals, "#N/A", False).value == "#N/A"
        # Checked before the key is searched: its error wins, found or not.
        assert VLOOKUP("Delta", m.deals, 4, False).value == "#REF!"
        assert VLOOKUP("Delta", m.deals, 0, False).value == "#VALUE!"
        assert VLOOKUP("#DIV/0!", m.deals, 4, False).value == "#REF!"
        assert VLOOKUP("#DIV/0!", m.deals, 0, False).value == "#VALUE!"
        assert VLOOKUP("#DIV/0!", m.deals, 3, False).value == "#DIV/0!"
        out = VLOOKUP("Beta", m.deals, mo.Variable([1, 2, 3]), False)
        assert out.value == ["Beta", "South", 200.0] and out.var_type == "list"

    def test_a_series_key_a_blank_key_and_an_error_key(self):
        m = _model()
        keys = mo.Variable(["Alpha", "Delta", "Gamma"])
        assert VLOOKUP(keys, m.deals, 3, False).value == [100.0, "#N/A", 300.0]
        assert VLOOKUP(mo.Variable(None), m.deals, 3, False).value == "#N/A"
        assert VLOOKUP("#DIV/0!", m.deals, 3, False).value == "#DIV/0!"

    def test_over_a_list_of_rows(self):
        m = _model()
        # The first column of a list of rows is the rows' first cells.
        assert VLOOKUP(2026, [m.i.years, m.i.rates], 3, False).value == 2028
        assert VLOOKUP(0.05, [m.i.years, m.i.rates], 2, False).value == 0.04
        assert VLOOKUP(2027, [m.i.years, m.i.rates], 1, False).value == "#N/A"

    def test_type_and_unit_come_from_the_column(self):
        from modeleon.core.unit import unit_key
        m = _model()
        assert VLOOKUP("Beta", m.deals, 1, False).value_type == "string"
        assert VLOOKUP("Beta", m.deals, 2, False).value_type == "string"
        amount = VLOOKUP("Beta", m.deals, 3, False)
        assert amount.value_type == "float"
        assert unit_key(amount._unit) == unit_key(m.deals.a.amount._unit)
        # A position read from a cell picks the column by the cell's value.
        assert unit_key(VLOOKUP("Beta", m.deals, mo.Variable(3), False)._unit) == unit_key(
            m.deals.a.amount._unit)
        assert VLOOKUP("Beta", m.deals, mo.Variable(2), False).value_type == "string"

    def test_a_record_that_lacks_a_field_leaves_the_cell_blank(self):
        class Custom(mo.MultiVariableClass):
            def compute(self, salary, bonus=None):
                self.salary = mo.Variable(salary, display_name="Salary")
                if bonus is not None:
                    self.bonus = mo.Variable(bonus, display_name="Bonus")

        m = mo.Model("m")
        with m:
            m.staff = mo.MultiVariable("Staff", excel_props={"tab": True, "orient": "records"})
            with m.staff as s:
                s.a = Custom(salary=1, display_name="A")
                s.b = Custom(salary=2, bonus=9, display_name="B")
        assert VLOOKUP("A", m.staff, 3, False).value == 0
        assert VLOOKUP("B", m.staff, 3, False).value == 9
        assert INDEX(m.staff, 1, 3).value == 0

    def test_the_formula_and_the_flag(self):
        m = _model()
        out = VLOOKUP("Beta", m.deals, 3, False)
        assert [type(a) for a in out._expr.args] == [Literal, ListExpr, Literal, Literal]
        assert out._expr.args[3].value is False
        assert out._source_code == "VLOOKUP('Beta', deals, 3, False)"
        assert VLOOKUP("Beta", m.deals, 3)._source_code == "VLOOKUP('Beta', deals, 3)"
        assert VLOOKUP("Bob", m.deals, 3, 0).value == "#N/A"

    def test_refusals(self):
        m = _model()
        with pytest.raises(ValueError, match="range_lookup is TRUE or FALSE"):
            VLOOKUP("Beta", m.deals, 3, m.i.pick)
        with pytest.raises(ValueError, match="holds no records"):
            VLOOKUP("x", m.i, 1)
        with pytest.raises(TypeError, match="a block is"):
            VLOOKUP("x", m.i.rates, 1)
        with pytest.raises(ValueError, match="a block is a list of rows"):
            INDEX([m.i.years, m.i.pick], 1, 1)
        with pytest.raises(ValueError, match="same number of cells"):
            INDEX([m.i.years, mo.Variable([1, 2])], 1, 1)
        with pytest.raises(ValueError, match="no scalar fields"):
            class Only(mo.MultiVariableClass):
                def compute(self):
                    self.flow = mo.Variable([1.0, 2.0, 3.0])
            t = mo.MultiVariable("T", excel_props={"orient": "records"})
            t.a = Only()
            VLOOKUP("a", t, 1)


class TestAResultFromDifferentLines:
    """A result that can come from cells of different lines is typed by them
    when they agree, else by the value found; given a unit only when every one
    of them carries the same unit."""

    def test_the_cells_agree(self):
        from modeleon.core.unit import unit_key
        m = _model()
        amount = INDEX(m.deals, mo.Variable([1, 2, 3]), 3)
        assert amount.value_type == "float"
        assert unit_key(amount._unit) == unit_key(m.deals.a.amount._unit)
        assert INDEX(m.deals, mo.Variable([1, 2, 3]), 2).value_type == "string"

    def test_the_cells_disagree(self):
        m = _model()
        rows = [m.i.years, m.i.rates]
        out = VLOOKUP(2027.5, rows, 2)                 # a year or a rate, by the key
        assert out.value == 0.04 and out.value_type == "float" and out._unit is None
        out = VLOOKUP(2026, rows, 2, False)
        assert out.value == 2027 and out.value_type == "int" and out._unit is None
        out = HLOOKUP("alpha", m.deals, 3, False)      # a label, a region or an amount
        assert out.value == "Gamma" and out.value_type == "string" and out._unit is None
        out = INDEX(rows, mo.Variable([1, 2, 1]), 2)   # a series: a number
        assert out.value == [2027, 0.04, 2027]
        assert out.value_type == "float" and out._unit is None
        out = INDEX(m.deals, 2, mo.Variable([1, 2, 3]))
        assert out.value == ["Beta", "South", 200.0]
        assert out.value_type == "float" and out._unit is None


class TestHLOOKUP:
    def test_over_a_list_of_rows(self):
        m = _model()
        assert HLOOKUP(2027, [m.i.years, m.i.rates], 2, False).value == 0.04
        assert HLOOKUP(2027.5, [m.i.years, m.i.rates], 2).value == 0.04
        assert HLOOKUP(2025, [m.i.years, m.i.rates], 2).value == "#N/A"
        assert HLOOKUP(2027, [m.i.years, m.i.rates], 1, False).value == 2027
        assert HLOOKUP(2027, [m.i.years, m.i.rates], 3, False).value == "#REF!"
        assert HLOOKUP(2027, [m.i.years, m.i.rates], 0, False).value == "#VALUE!"
        assert HLOOKUP(2025, [m.i.years, m.i.rates], 3, False).value == "#REF!"
        assert HLOOKUP("#N/A", [m.i.years, m.i.rates], 0, False).value == "#VALUE!"
        assert HLOOKUP([2028, 2026], [m.i.years, m.i.rates], 2, False).value == [0.03, 0.05]

    def test_over_a_table_of_records_the_labels_are_the_first_row_s_cells(self):
        m = _model()
        assert HLOOKUP("North", m.deals, 2, False).value == "South"
        assert HLOOKUP("Alpha", m.deals, 1, False).value == "Alpha"

    def test_type_and_unit_come_from_the_row(self):
        from modeleon.core.unit import unit_key
        m = _model()
        rate = HLOOKUP(2027, [m.i.years, m.i.rates], 2, False)
        assert rate.value_type == "float"
        assert unit_key(rate._unit) == unit_key(m.i.rates._unit)
        assert HLOOKUP(2027, [m.i.years, m.i.rates], 1, False).value_type == "int"
        assert HLOOKUP(2027, [m.i.years, m.i.rates], 1, False)._unit is None
        out = HLOOKUP(2027, [m.i.years, m.i.rates], 2, False)
        assert out._source_code == "HLOOKUP(2027, [i.years, i.rates], 2, False)"


# ══════════════════════════════════════════════════════════════════════
# The registry and the AST
# ══════════════════════════════════════════════════════════════════════


class TestTheRegistry:
    @pytest.mark.parametrize("name, positions, prefix", [
        ("MATCH", (1,), ""), ("INDEX", (0,), ""), ("XLOOKUP", (1, 2), "_xlfn."),
        ("VLOOKUP", (1,), ""), ("HLOOKUP", (1,), ""),
    ])
    def test_every_function_is_a_range_function(self, name, positions, prefix):
        s = spec(name)
        assert s is not None and s.shape == "range" and not s.bare
        assert is_range_taking(name) and is_passthrough(name)
        assert range_positions(name) == positions
        assert file_prefix(name) == prefix
        assert not admits_in_loop(name, 3)

    def test_na_is_written_as_given(self):
        s = spec("NA")
        assert s is not None and s.bare and s.shape == "scalar" and is_passthrough("NA")


class TestTheAST:
    def test_every_function_carries_a_funccall(self):
        m = _model()
        for out in (
            MATCH(m.i.pick, m.i.years, 0),
            INDEX(m.i.rates, m.i.pick),
            INDEX(m.deals, m.i.pick, 1),
            XLOOKUP(m.i.pick, m.i.years, m.i.rates),
            XLOOKUP(m.i.pick, m.i.infl),
            VLOOKUP(m.i.pick, m.deals, 2),
            HLOOKUP(m.i.pick, [m.i.years, m.i.rates], 2),
        ):
            assert isinstance(out._expr, FuncCall)
            assert {m.i.pick.id} <= {v.id for v in out._expr.iter_refs()}

    def test_the_block_s_lines_are_dependencies(self):
        m = _model()
        out = VLOOKUP("Beta", m.deals, 3, False)
        refs = {v.id for v in out._expr.iter_refs()}
        assert {m.deals.a.region.id, m.deals.c.amount.id} <= refs
        assert m.deals.a.flow.id not in refs


def _unit_lines():
    m = mo.Model("m", default_grain="year", default_start="2026", default_periods=3)
    with m:
        m.i = mo.MultiVariable("I")
        with m.i as i:
            i.amount = mo.Variable(150.0, display_name="Amount", unit="USD")
            i.amounts = mo.Variable([90.0, 150.0, 400.0], display_name="Amounts", unit="USD")
            i.thresholds = mo.Variable([0.0, 100.0, 300.0], display_name="Thresholds",
                                       unit="EUR")
            i.fees = mo.Variable([1.0, 2.0, 3.0], display_name="Fees", unit="EUR")
            i.grades = mo.Variable([1.0, 2.0, 3.0], display_name="Grades")
            i.usd = mo.Variable(10.0, display_name="Usd", unit="USD")
            i.eur = mo.Variable(20.0, display_name="Eur", unit="EUR")
            i.k = mo.Variable([1, 2, 1], display_name="K")
    return m.i


class TestUnits:
    """A key is compared with the cells searched, as ``==`` compares two
    lines; a cell is taken from the cells returned, as ``CHOOSE`` takes
    one: two units are refused, a number or an unlabelled line goes with
    any, and the result is in the returned cells' unit, never the key's."""

    def test_a_key_in_another_unit_than_the_cells_searched_is_refused(self):
        i = _unit_lines()
        calls = {
            "MATCH": lambda: MATCH(i.amount, i.thresholds, 1),
            "XLOOKUP": lambda: XLOOKUP(i.amount, i.thresholds, i.fees, match_mode=-1),
            "VLOOKUP": lambda: VLOOKUP(i.amount, [i.thresholds, i.fees], 2),
            "HLOOKUP": lambda: HLOOKUP(i.amount, [i.thresholds, i.fees], 2),
        }
        for name, call in calls.items():
            with pytest.raises(ValueError, match=f"mo.{name} compares the key 'Amount' with "
                                                 f"'Thresholds'; they are in USD and EUR"):
                call()
        with pytest.raises(ValueError, match="compares the key 'Amount' with 'Eur'"):
            XLOOKUP(i.amount, [i.usd, i.eur], [1.0, 2.0])      # a list of lines searched
        # A typed key, an unlabelled line or a key in the cells' unit goes.
        assert MATCH(150.0, i.thresholds, 1).value == 2
        assert MATCH(i.grades, i.thresholds, 1).value == [1, 1, 1]
        out = XLOOKUP(i.amount, i.amounts, i.fees)
        assert out.value == 2.0 and str(out.unit) == "EUR"           # never the key's

    def test_an_if_not_found_in_another_unit_than_the_cells_returned_is_refused(self):
        i = _unit_lines()
        with pytest.raises(ValueError, match="mo.XLOOKUP takes its values in one unit; "
                                             "'Fees' and 'Usd' are in EUR and USD"):
            XLOOKUP(i.amount, i.amounts, i.fees, i.usd)
        assert str(XLOOKUP(i.amount, i.amounts, i.fees, 0.0).unit) == "EUR"
        assert str(XLOOKUP(i.amount, i.amounts, i.fees, i.eur).unit) == "EUR"
        # The returned cells' unit, as before: a fallback cannot lend one.
        assert XLOOKUP(i.amount, i.amounts, i.grades, i.eur).unit is None

    def test_cells_returned_in_two_units_are_refused(self):
        i = _unit_lines()
        calls = [
            lambda: INDEX([i.usd, i.eur], 2),
            lambda: INDEX([i.usd, i.eur], i.k),
            lambda: XLOOKUP(i.amount, i.amounts, [i.usd, i.eur, i.usd]),
            lambda: INDEX([i.amounts, i.fees], i.k, 2),          # a row position that moves
            lambda: VLOOKUP(2.0, [i.grades, i.amounts, i.fees], 2, False),
            lambda: HLOOKUP(2.0, [i.grades, i.amounts, i.fees], i.k, False),
        ]
        for call in calls:
            with pytest.raises(ValueError, match="takes its values in one unit; 'Amounts' and "
                                                 "'Fees' are|takes its values in one unit; "
                                                 "'Usd' and 'Eur' are"):
                call()
        # One cell addressed: its own unit; a label or a line without a unit
        # among the cells leaves none, as before.
        assert str(INDEX([i.amounts, i.fees], 2, 2).unit) == "EUR"
        assert INDEX([i.grades, i.fees], i.k, 2).unit is None

    def test_a_position_the_block_lacks_is_its_error_not_a_refusal(self):
        # The call is #VALUE! / #REF! in every period: no cell is taken.
        i = _unit_lines()
        assert INDEX([i.amounts, i.fees], 0, 2).value == "#VALUE!"
        assert INDEX([i.amounts, i.fees], 5, 2).value == "#REF!"
        assert VLOOKUP(90.0, [i.amounts, i.fees], 9, False).value == "#REF!"

    def test_the_law_gives_the_unit_the_call_is_built_with(self):
        # What a loop container derives the row by again (function_unit)
        # agrees with the call: the returned cells' unit, never the key's.
        from modeleon.core.unit import unit_key
        from modeleon.functions._helpers import function_unit, operand_unit
        i = _unit_lines()
        for out in (MATCH(i.amount, i.amounts, 0), INDEX([i.usd, i.usd], 2),
                    INDEX([i.amounts, i.fees], 2, 2), XLOOKUP(i.amount, i.amounts, i.grades),
                    XLOOKUP(i.amount, i.amounts, i.fees, 0.0),
                    VLOOKUP(2.0, [i.grades, i.fees], 2, False),
                    HLOOKUP(2.0, [i.grades, i.fees], 2, False)):
            law = function_unit(out._expr.func, out._expr.args, operand_unit)
            assert (None if law is None else unit_key(law)) == (
                None if out.unit is None else unit_key(out.unit)), out.formula

    def test_a_key_read_before_its_line_is_compared_when_the_loop_closes(self):
        def loop(body):
            m = mo.Model("m", default_grain="month", default_start="2026-01", default_periods=3)
            with m:
                m.inp = mo.MultiVariable("Inputs")
                with m.inp as i:
                    i.start = mo.Variable(100.0, unit="USD")
                    i.inflow = mo.Variable([100.0] * 3, unit="USD")
                    i.thresholds = mo.Variable([100.0, 200.0, 300.0], unit="EUR",
                                               display_name="Thresholds")
                    i.fees = mo.Variable([1.0, 2.0, 3.0], unit="EUR")
                m.loop = mo.MultiVariable("Loop", loop=True)
                with m.loop as d:
                    body(d, m.inp)
            return m.loop

        def mixed(d, i):
            d.opening = mo.lag(d.closing, fill=i.start)
            d.fee = mo.XLOOKUP(d.opening, i.thresholds, i.fees, 0.0)
            d.closing = d.opening + i.inflow
        with pytest.raises(ValueError, match="fee: mo.XLOOKUP compares the key 'opening' "
                                             "with 'Thresholds'; they are in USD and EUR"):
            loop(mixed)

        def same(d, i):
            d.opening = mo.lag(d.closing, fill=i.start)
            d.fee = mo.XLOOKUP(d.opening, i.inflow * 0 + [100.0, 200.0, 300.0], i.fees, 0.0)
            d.closing = d.opening + i.inflow
        out = loop(same)
        assert out.fee.value == [1.0, 2.0, 3.0] and str(out.fee.unit) == "EUR"
