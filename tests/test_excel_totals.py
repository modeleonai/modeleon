# SPDX-License-Identifier: Apache-2.0
"""Subtotal columns — "Jan Feb Mar (Q1) … (year)" — and they are ALIVE.

``ExcelView(timeline={'totals': ['quarter', 'year']})`` interleaves a
quarter column after each quarter's last month and a year column after
the year. The parity promise extends to them: edit January in the
written file and the quarter, the year, and every formula standing
on them move. That rules out baked numbers everywhere a formula can
honestly stand:

* a rule line   → ``=SUM(months)`` / a mean over the months' days /
  last-month ref - a month not entered counts as zero, a mean skips it;
* a formula line → its OWN formula over the operands' bucket cells —
  exactly what a hand-built book writes at the quarter column;
* a line computed month by month and summed (an IF, a product of two
  flows) → ``=SUM`` of its own month cells, as the engine folds it;
* a ruleless literal → the projection's ``#VALUE!`` teaching token,
  never an invented sum.
"""

import os
import re
import tempfile
from datetime import date, timedelta

import openpyxl
import pytest
from openpyxl.utils import get_column_letter

import modeleon as mo


def _book(model):
    path = os.path.join(tempfile.mkdtemp(), "t.xlsx")
    model.to_excel(path)
    return openpyxl.load_workbook(path)


def _tab(marked):
    """A first-level container is a tab either way: marked with
    ``excel_props={'tab': True}``, or unmarked — the default, where the
    writer makes the tab itself. The totals must not tell them apart."""
    return {"excel_props": {"tab": True}} if marked else {}


# Every content expectation below runs on both kinds of tab.
both_tabs = pytest.mark.parametrize(
    "marked", [True, False], ids=["marked-tab", "unmarked-tab"]
)


def _model(marked=True, **view_kw):
    m = mo.Model(
        "m",
        display_name="M",
        default_excel_view=mo.ExcelView(
            timeline={"totals": ["quarter", "year"], **view_kw}
        ),
    )
    m.default_start, m.default_grain, m.default_periods = "2026-01", "month", 6
    with m:
        m.plan = mo.MultiVariable("Plan", **_tab(marked))
        with m.plan as p:
            p.rate = mo.Variable(0.1, display_name="Rate")
            p.revenue = mo.Variable(
                [10, 20, 30, 40, 50, 60],
                display_name="Revenue",
                regrain=mo.up("sum"),
            )
            p.tax = p.revenue * p.rate
    return m


class TestTheBookForm:
    def test_the_header_is_hierarchical(self):
        # "Year" over "quarters" over "months" — one header row per
        # grain. A bucket spans its months and the finer totals inside
        # it, and its OWN total stands beside that span, named, from
        # the bucket's own tier downwards: "Total Q1" begins where "Q1"
        # begins, not a tier lower where it read as a thirteenth month.
        ws = _book(_model())["Plan"]
        assert ws.cell(1, 3).value == "2026"
        assert ws.cell(2, 3).value == "Q1"
        assert ws.cell(2, 7).value == "Q2"
        # The quarter total names itself ON the quarter row…
        assert ws.cell(2, 6).value == "Total Q1"
        assert ws.cell(2, 10).value == "Total Q2"
        # …so the month row carries months and nothing else.
        assert [ws.cell(3, c).value for c in range(3, 12)] == [
            "Jan", "Feb", "Mar", None,
            "Apr", "May", "Jun", None, None,
        ]
        # The year total starts on the YEAR row and runs the full stack.
        assert ws.cell(1, 11).value == "Total 2026"
        merges = sorted(str(r) for r in ws.merged_cells.ranges)
        assert "K1:K3" in merges
        assert "C1:J1" in merges          # the year: months + Q-totals
        assert "C2:E2" in merges          # Q1: its months only
        assert "F2:F3" in merges          # its total, two rows tall
        assert "G2:I2" in merges

    def test_data_starts_below_the_header_stack(self):
        ws = _book(_model())["Plan"]
        assert ws.cell(4, 1).value == "Rate"
        assert ws.freeze_panes == "C4"

    @both_tabs
    def test_a_rule_line_sums_its_months_live(self, marked):
        ws = _book(_model(marked))["Plan"]
        assert ws["F5"].value == "=SUM(C5,D5,E5)"
        assert ws["J5"].value == "=SUM(G5,H5,I5)"

    @both_tabs
    def test_the_year_never_swallows_the_quarter_cells(self, marked):
        # The year's months are NOT contiguous once quarter columns
        # stand between them — a range SUM would double-count. Explicit
        # refs, months only.
        ws = _book(_model(marked))["Plan"]
        assert ws["K5"].value == "=SUM(C5,D5,E5,G5,H5,I5)"

    @both_tabs
    def test_a_formula_line_writes_its_own_formula_at_the_bucket(self, marked):
        # The hand-built book's rule: tax(Q1) = revenue(Q1) × rate,
        # standing on the QUARTER cell — not a sum of monthly taxes,
        # not a baked number.
        ws = _book(_model(marked))["Plan"]
        assert ws["F6"].value == "=F5 * B4"
        assert ws["J6"].value == "=J5 * B4"
        assert ws["K6"].value == "=K5 * B4"

    @both_tabs
    def test_native_months_still_reference_their_own_cells(self, marked):
        # The interleave moves ADDRESSES, so native formulas follow.
        ws = _book(_model(marked))["Plan"]
        assert ws["C6"].value == "=C5 * B4"
        assert ws["G6"].value == "=G5 * B4"

    @both_tabs
    def test_a_ruleless_literal_teaches_instead_of_inventing(self, marked):
        m = mo.Model(
            "m",
            display_name="M",
            default_excel_view=mo.ExcelView(
                timeline={"totals": ["quarter"]}
            ),
        )
        m.default_start, m.default_grain, m.default_periods = (
            "2026-01", "month", 3,
        )
        with m:
            m.plan = mo.MultiVariable("Plan", **_tab(marked))
            with m.plan as p:
                p.fx_rate = mo.Variable(
                    [1.10, 1.11, 1.12], display_name="FX rate"
                )
        ws = _book(m)["Plan"]
        # No rule declared: summing an FX rate would be a lie; the
        # projection's token lands instead.
        assert ws["E3"].value == "#VALUE!"

    @both_tabs
    def test_partial_trailing_quarter_still_totals(self, marked):
        m = mo.Model(
            "m",
            display_name="M",
            default_excel_view=mo.ExcelView(
                timeline={"totals": ["quarter"]}
            ),
        )
        m.default_start, m.default_grain, m.default_periods = (
            "2026-01", "month", 4,
        )
        with m:
            m.plan = mo.MultiVariable("Plan", **_tab(marked))
            with m.plan as p:
                p.flow = mo.Variable(
                    [1, 2, 3, 4], display_name="Flow",
                    regrain=mo.up("sum"),
                )
        ws = _book(m)["Plan"]
        # Q1 after March; the lone April closes a partial Q2.
        assert ws["E3"].value == "=SUM(B3,C3,D3)"
        assert ws["G3"].value == "=SUM(F3)"

    def test_undeclared_totals_change_nothing(self):
        m = _model()
        m.default_excel_view = mo.ExcelView()
        ws = _book(m)["Plan"]
        row1 = [ws.cell(1, c).value for c in range(1, 10)]
        assert "Q1 2026" not in row1
        assert ws["C3"].value == 10  # contiguous months, untouched


class TestOtherRules:
    def _one_line(self, marked, **var_kw):
        m = mo.Model(
            "m",
            display_name="M",
            default_excel_view=mo.ExcelView(
                timeline={"totals": ["quarter"]}
            ),
        )
        m.default_start, m.default_grain, m.default_periods = (
            "2026-01", "month", 3,
        )
        with m:
            m.plan = mo.MultiVariable("Plan", **_tab(marked))
            with m.plan as p:
                p.line = mo.Variable(
                    [10, 20, 30], display_name="Line", **var_kw
                )
        return _book(m)["Plan"]

    @both_tabs
    def test_last(self, marked):
        # A balance's quarter is its closing month — a ref, not a sum.
        ws = self._one_line(marked, regrain=mo.up("last"))
        assert ws["E3"].value == "=D3"

    @both_tabs
    def test_mean(self, marked):
        # The engine's mean is weighted by the months' days (31, 28, 31 in
        # Q1 2026), and so is the book's.
        ws = self._one_line(marked, regrain=mo.up("mean"))
        assert ws["E3"].value == "=(B3*31+C3*28+D3*31)/90"


class TestEveryRuleIsAFormula:
    """Every re-grain rule totals as live arithmetic over its months — the
    engine's own fold, never a typed number (``first``, ``min``, ``max``,
    ``geometric`` and a ratio used to land as values)."""

    def _ops(self):
        m = mo.Model("m", default_grain="month", default_start="2026-01",
                     default_periods=3,
                     default_excel_view=mo.ExcelView(timeline={"totals": ["quarter"]}))
        with m:
            m.ops = mo.MultiVariable("Ops")
            with m.ops as a:
                a.first = mo.Variable([5.0, 6.0, 7.0], display_name="First", regrain=mo.up("first"))
                a.low = mo.Variable([5.0, 2.0, 7.0], display_name="Low", regrain=mo.up("min"))
                a.high = mo.Variable([5.0, 9.0, 7.0], display_name="High", regrain=mo.up("max"))
                a.growth = mo.Variable([0.01, 0.02, 0.03], display_name="Growth",
                                       regrain=mo.up("geometric"))
                a.revenue = mo.Variable([100.0, 120.0, 80.0], display_name="Revenue",
                                        regrain=mo.up("sum"))
                a.profit = mo.Variable([40.0, 30.0, 20.0], display_name="Profit",
                                       regrain=mo.up("sum"))
                a.margin = (a.profit / a.revenue).set_display_name("Margin").set_regrain(
                    mo.ratio("profit", "revenue"))
        return m

    def _total(self, ws, label):
        row = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == label)
        return ws.cell(row, 5).value

    def test_the_formulas(self):
        ws = _book(self._ops())["Ops"]
        assert self._total(ws, "First") == "=B3"
        assert self._total(ws, "Low") == "=MIN(B4,C4,D4)"
        assert self._total(ws, "High") == "=MAX(B5,C5,D5)"
        assert self._total(ws, "Growth") == "=EXP(LN(1+B6)+LN(1+C6)+LN(1+D6))-1"
        assert self._total(ws, "Margin") == "=SUM(B8,C8,D8)/SUM(B7,C7,D7)"

    @pytest.mark.slow
    def test_they_calculate_what_the_engine_does(self, tmp_path):
        m = self._ops()
        ws = TestTheTotalsCalculate._recalculated(m, tmp_path)["Ops"]
        q = m.at("quarter").ops
        for label, line in (("First", q.first), ("Low", q.low), ("High", q.high),
                            ("Growth", q.growth), ("Margin", q.margin)):
            assert self._total(ws, label) == pytest.approx(line.value[0], rel=1e-12), label


class TestTheYearSheet:
    """A sheet of ``.at('year')`` copies reads the monthly lines with the
    engine's own fold - a mean weighted by the months' days, a compound
    growth - and only their month cells, never the subtotal columns
    standing between them (``SUM(B4:P4)`` swallowed the quarters)."""

    def _model(self, totals):
        m = mo.Model("m", default_grain="month", default_start="2026-01", default_periods=12)
        with m:
            m.ops = mo.MultiVariable("Ops")
            with m.ops as a:
                a.price = mo.Variable([10.0] * 6 + [12.0] * 6, display_name="Price",
                                      regrain=mo.up("mean"))
                a.cups = mo.Variable([100.0] * 12, display_name="Cups", regrain=mo.up("sum"))
                a.growth = mo.Variable([0.01] * 12, display_name="Growth",
                                       regrain=mo.up("geometric"))
            m.year = mo.MultiVariable("By year", default_grain="year", default_start="2026",
                                      default_periods=1)
            with m.year as y:
                y.price = m.ops.price.at("year")
                y.cups = m.ops.cups.at("year")
                y.growth = m.ops.growth.at("year")
            if totals:
                m.default_excel_view = mo.ExcelView(timeline={"totals": ["quarter"]})
        return m

    @staticmethod
    def _cell(ws, label):
        row = next(r for r in range(1, ws.max_row + 1)
                   if str(ws.cell(r, 1).value or "").startswith(label))
        return ws.cell(row, 2).value

    def test_the_folds_are_formulas(self):
        ws = _book(self._model(False))["By year"]
        assert self._cell(ws, "Price").startswith("=(Ops!B2*31+Ops!C2*28+")
        assert self._cell(ws, "Price").endswith(")/365")
        assert self._cell(ws, "Growth") == "=EXP(SUMPRODUCT(LN(1+Ops!B4:M4)))-1"
        assert self._cell(ws, "Cups") == "=SUM(Ops!B3:M3)"

    def test_the_year_skips_the_subtotal_columns(self):
        ws = _book(self._model(True))["By year"]
        assert self._cell(ws, "Cups") == "=SUM(Ops!B4:D4,Ops!F4:H4,Ops!J4:L4,Ops!N4:P4)"

    @pytest.mark.slow
    @pytest.mark.parametrize("totals", [False, True], ids=["plain", "with-subtotals"])
    def test_the_year_sheet_calculates_what_the_engine_does(self, tmp_path, totals):
        m = self._model(totals)
        ws = TestTheTotalsCalculate._recalculated(m, tmp_path)["By year"]
        year = m.at("year").ops
        for label, line in (("Price", year.price), ("Cups", year.cups),
                            ("Growth", year.growth)):
            assert self._cell(ws, label) == pytest.approx(line.value[0], rel=1e-12), label

    @pytest.mark.slow
    def test_a_year_read_over_blank_months_counts_them_as_zero(self, tmp_path):
        holes = [None, None, 3.0, None]
        m = mo.Model("m", default_grain="month", default_start="2026-01", default_periods=12)
        with m:
            m.ops = mo.MultiVariable("Ops")
            with m.ops as a:
                for name, rule, first in (("price", "mean", 10.0), ("cups", "sum", 100.0),
                                          ("growth", "geometric", 0.01), ("high", "max", -5.0),
                                          ("close", "last", 7.0)):
                    setattr(a, name, mo.Variable([first] * 8 + holes, display_name=name.title(),
                                                 regrain=mo.up(rule)))
            m.year = mo.MultiVariable("By year", default_grain="year", default_start="2026",
                                      default_periods=1)
            with m.year as y:
                for name in ("price", "cups", "growth", "high", "close"):
                    setattr(y, name, getattr(m.ops, name).at("year"))
        ws = TestTheTotalsCalculate._recalculated(m, tmp_path)["By year"]
        for name in ("price", "cups", "growth", "high", "close"):
            engine = getattr(m.year, name).value
            engine = engine[0] if isinstance(engine, list) else engine
            assert self._cell(ws, name.title()) == pytest.approx(engine, rel=1e-12), name
        assert m.year.close.value in (0, [0])


class TestTheTotalsCountABlankAsZero:
    """A month not entered yet (a plan/actual line before its actuals) is a
    blank cell: every subtotal stays a live fold over the bucket's months
    and counts the blank as zero, as Excel does and as the engine's bucket
    does - a quarter with one month entered sums that month; a mean, a
    minimum and a maximum skip it, as AVERAGE, MIN and MAX do."""

    @staticmethod
    def _model():
        actual_q1 = {"sum": [7.0, None, 1.0], "mean": [11.0, None, 11.0],
                   "min": [4.0, None, 2.0], "max": [-4.0, None, -2.0],
                   "first": [None, 3.0, 4.0], "last": [5.0, 6.0, None],
                   "geometric": [0.1, None, 0.1], "units": [2.0, None, 2.0]}
        m = mo.Model("m", tracks=mo.Tracks("plan", "actual"), default_grain="month",
                     default_start="2026-01", default_periods=6,
                     default_excel_view=mo.ExcelView(timeline={"totals": ["quarter"]}))
        with m:
            m.ops = mo.MultiVariable("Ops")
            with m.ops as o:
                for recipe, q1 in actual_q1.items():
                    setattr(o, recipe, mo.Variable(
                        plan=[1.0, 2, 3, 4, 5, 6], actual=q1 + [None] * 3,
                        display_name=recipe.title(),
                        regrain=mo.up("sum" if recipe == "units" else recipe)))
                o.margin = (o.sum / o.units).set_display_name("Margin").set_regrain(
                    mo.ratio("sum", "units"))
        return m

    @staticmethod
    def _actual_totals(ws):
        rows = {str(ws.cell(r, 1).value): r for r in range(1, ws.max_row + 1)
                if ws.cell(r, 1).value}
        totals = [c for c in range(1, ws.max_column + 1)
                  if any(str(ws.cell(h, c).value or "").startswith("Total") for h in (1, 2))]
        return {label.split(" · ")[0]: [ws.cell(r, c).value for c in totals]
                for label, r in rows.items() if label.endswith(" · actual")}

    def test_every_subtotal_folds_live_over_the_blank_months(self):
        totals = self._actual_totals(_book(self._model())["Ops"])
        assert len(totals) == 9
        for label, cells in totals.items():
            for cell in cells:
                assert isinstance(cell, str) and cell.startswith("=") and cell != '=""', (
                    label, cell)
        # A mean over the months entered, 0 when none is.
        mean_q1 = totals["Mean"][0]
        assert mean_q1.startswith("=IF(COUNTA(") and "ISNUMBER(" in mean_q1, mean_q1

    def test_a_bucket_entered_whole_keeps_the_plain_mean(self):
        rows = {str(c[0].value): c for c in _book(self._model())["Ops"].iter_rows()}
        plan = [c.value for c in rows["Mean"] if isinstance(c.value, str)
                and c.value.startswith("=")]
        assert plan and all(f.startswith("=(") and "COUNTA" not in f for f in plan), plan

    def test_a_line_of_dates_keeps_its_blank_subtotal(self):
        from datetime import date

        m = mo.Model("m", default_grain="month", default_start="2026-01", default_periods=6,
                     default_excel_view=mo.ExcelView(timeline={"totals": ["quarter"]}))
        with m:
            m.ops = mo.MultiVariable("Ops")
            with m.ops as o:
                o.paid = mo.Variable([date(2026, 1, 15)] * 2 + [None] * 4, display_name="Paid",
                                     regrain=mo.up("last"))
        ws = _book(m)["Ops"]
        row = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == "Paid")
        totals = [c for c in range(1, ws.max_column + 1)
                  if any(str(ws.cell(h, c).value or "").startswith("Total") for h in (1, 2))]
        assert [ws.cell(row, c).value for c in totals] == ['=""', '=""']
        assert m.at("quarter").ops.paid.value == [None, None]

    def test_a_tracked_line_of_dates_keeps_its_blank_on_an_empty_track(self):
        """The actual track has nothing entered yet - no date to see - but
        the line is a line of dates: its quarters stay blank there."""
        from datetime import date

        m = mo.Model("m", tracks=mo.Tracks("plan", "actual"), default_grain="month",
                     default_start="2026-01", default_periods=6,
                     default_excel_view=mo.ExcelView(timeline={"totals": ["quarter"]},
                                                     tracks="rows"))
        with m:
            m.ops = mo.MultiVariable("Ops")
            with m.ops as o:
                o.paid = mo.Variable(plan=[date(2026, k, 15) for k in range(1, 7)],
                                     actual=[None] * 6, display_name="Paid", regrain=mo.up("last"))
                o.grade = mo.Variable(["B", None, "A", "A", "A", "A"], display_name="Grade",
                                      regrain=mo.up("max"))
        assert m.at("quarter").ops.paid.value["actual"] == [None, None]
        assert m.at("quarter").ops.grade.value == [None, "A"]
        ws = _book(m)["Ops"]
        totals = [c for c in range(1, ws.max_column + 1)
                  if any(str(ws.cell(h, c).value or "").startswith("Total") for h in (1, 2))]
        row = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == "Paid · actual")
        assert [ws.cell(row, c).value for c in totals] == ['=""', '=""']
        row = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == "Grade")
        assert ws.cell(row, totals[0]).value == '=""'

    def test_the_engine_counts_the_blank_as_zero(self):
        q = self._model().at("quarter").ops
        assert q.sum.value["actual"] == [8.0, 0]
        assert q.max.value["actual"] == [-2.0, 0]
        assert q.first.value["actual"] == [0, 0]
        assert q.mean.value["actual"] == [11.0, 0]
        assert q.margin.value["actual"] == [2.0, "#DIV/0!"]

    @staticmethod
    def _blended(byq=True):
        m = mo.Model("m", tracks=mo.Tracks("plan", "actual",
                                           blend=mo.blend(given="actual", follow="plan")),
                     default_grain="month", default_start="2026-01", default_periods=6,
                     default_excel_view=mo.ExcelView(tracks="rows"))
        with m:
            m.ops = mo.MultiVariable("Ops")
            with m.ops as o:
                o.revenue = mo.Variable(plan=[100.0] * 6, actual=[90.0, 95.0] + [None] * 4,
                                        regrain=mo.up("sum"), display_name="Revenue")
                o.cash = mo.Variable(plan=[1000.0, 1100, 1200, 1300, 1400, 1500],
                                     actual=[990.0, 1080.0] + [None] * 4,
                                     regrain=mo.up("last"), display_name="Cash")
                o.price = mo.Variable(plan=[10.0] * 6, actual=[11.0, 12.0] + [None] * 4,
                                      regrain=mo.up("mean"), display_name="Price")
                o.units = mo.Variable(plan=[10.0] * 6, actual=[9.0, 9.5] + [None] * 4,
                                      regrain=mo.up("sum"), display_name="Units")
                o.rev2 = (o.price * o.units).set_display_name("Rev2")
            if byq:
                m.byq = mo.MultiVariable("ByQ", default_grain="quarter",
                                         default_start="2026-Q1", default_periods=2)
                with m.byq as y:
                    y.revenue = m.ops.revenue.at("quarter")
                    y.cash = m.ops.cash.at("quarter")
        return m

    @staticmethod
    def _rows(ws):
        return {str(ws.cell(r, 1).value): [ws.cell(r, c).value for c in (2, 3)]
                for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value}

    def test_a_blended_line_read_by_quarter_is_the_engines_blend(self):
        """The actual quarter is no longer blank where a month is not
        entered, so splicing the folded tracks in the book
        (``=IF(actual_Q="",plan_Q,actual_Q)``) would take a partial quarter
        of actuals, or a future quarter's 0, for the blend."""
        m = self._blended()
        rows = self._rows(_book(m)["ByQ"])
        assert rows["Revenue"] == m.byq.revenue.value["live"] == [285.0, 300.0]
        assert rows["Cash"] == m.byq.cash.value["live"] == [1200, 1500]

    def test_a_projected_blended_product_is_the_engines_blend(self):
        from modeleon.core.projection import project_model

        q = project_model(self._blended(byq=False), "quarter", on_error="absorb")
        rows = self._rows(_book(q)["Ops"])
        assert rows["Rev2"] == q.ops.rev2.value["live"] == [313.0, 300.0]
        assert rows["Revenue"] == q.ops.revenue.value["live"] == [285.0, 300.0]

    @pytest.mark.slow
    def test_they_calculate_what_the_engine_does(self, tmp_path):
        m = self._model()
        ws = TestTheTotalsCalculate._recalculated(m, tmp_path)["Ops"]
        q = m.at("quarter").ops
        for label, cells in self._actual_totals(ws).items():
            expected = getattr(q, label.lower()).value["actual"]
            for got, want in zip(cells, expected):
                if isinstance(want, str):
                    assert got == want, label
                else:
                    assert got == pytest.approx(want, rel=1e-12, abs=1e-12), label


class TestAFormulaOverABlankMonth:
    """A formula month over an un-entered input computes over zero, in the
    engine as in Excel, so a balance carried on a recurrence ends its
    quarter at its running value in both."""

    @pytest.mark.slow
    def test_a_balance_over_a_blank_month_ends_its_quarter_as_the_book_does(self, tmp_path):
        m = mo.Model("m", default_grain="month", default_start="2026-01", default_periods=6,
                     default_excel_view=mo.ExcelView(timeline={"totals": ["quarter"]}))
        with m:
            m.ops = mo.MultiVariable("Ops")
            with m.ops as o:
                o.inflow = mo.Variable([10.0, 20.0] + [None] * 4, display_name="Inflow",
                                       regrain=mo.up("sum"))
                o.bal = mo.recurrence(100.0, "{prev} + {x}", x=o.inflow).set_display_name(
                    "Bal").set_regrain(mo.frozen(values="last"))
        ws = TestTheTotalsCalculate._recalculated(m, tmp_path)["Ops"]
        row = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == "Bal")
        totals = [c for c in range(1, ws.max_column + 1)
                  if any(str(ws.cell(h, c).value or "").startswith("Total") for h in (1, 2))]
        book = [ws.cell(row, c).value for c in totals]
        assert book == m.at("quarter").ops.bal.value


class TestARatioTotalsOnItsOwnTrack:
    def test_the_actual_ratio_reads_the_actual_lines(self):
        m = mo.Model("m", tracks=mo.Tracks("plan", "actual"), default_grain="month",
                     default_start="2026-01", default_periods=3,
                     default_excel_view=mo.ExcelView(timeline={"totals": ["quarter"]},
                                                     tracks="rows"))
        with m:
            m.ops = mo.MultiVariable("Ops")
            with m.ops as o:
                o.revenue = mo.Variable(plan=[100.0] * 3, actual=[90.0] * 3,
                                        display_name="Revenue", regrain=mo.up("sum"))
                o.profit = mo.Variable(plan=[40.0] * 3, actual=[20.0] * 3,
                                       display_name="Profit", regrain=mo.up("sum"))
                o.margin = (o.profit / o.revenue).set_display_name("Margin").set_regrain(
                    mo.ratio("profit", "revenue"))
        ws = _book(m)["Ops"]
        rows = {str(ws.cell(r, 1).value): r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value}
        total = next(c for c in range(1, ws.max_column + 1)
                     if any(str(ws.cell(h, c).value or "").startswith("Total") for h in (1, 2)))
        actual = ws.cell(rows["Margin · actual"], total).value
        assert f"B{rows['Profit · actual']}" in actual and f"B{rows['Revenue · actual']}" in actual, actual


def _totals(ws):
    """Every subtotal column of a sheet, left to right."""
    return [c for c in range(1, ws.max_column + 1)
            if any(str(ws.cell(h, c).value or "").startswith("Total") for h in (1, 2))]


def _row_values(ws, cols):
    return {str(ws.cell(r, 1).value): [ws.cell(r, c).value for c in cols]
            for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value}


def _engine_row(q, label, blend):
    """The engine's coarse values for a book row ``Name`` / ``Name · track``."""
    name, _, track = label.partition(" · ")
    value = getattr(q, name.lower()).value
    if hasattr(value, "roles"):
        value = value[track or ("live" if blend else value.roles[0])]
    return value


class TestATrackedFormulaLineFoldsLive:
    """A formula line with tracks: its display row and each track's row fold
    by the line's formula over the operands' bucket cells on the same track
    - they inlined whole tracked values as text, and the track rows were
    the ruleless '#VALUE!'."""

    @staticmethod
    def _model(blend):
        tracks = (mo.Tracks("plan", "actual", blend=mo.blend(given="actual", follow="plan"))
                  if blend else mo.Tracks("plan", "actual"))
        m = mo.Model("m", tracks=tracks, default_grain="month", default_start="2026-01",
                     default_periods=6,
                     default_excel_view=mo.ExcelView(timeline={"totals": ["quarter"]},
                                                     tracks="rows"))
        with m:
            m.ops = mo.MultiVariable("Ops")
            with m.ops as o:
                o.rev = mo.Variable(plan=[100.0] * 6, actual=[90.0, 95.0] + [None] * 4,
                                    display_name="Rev", regrain=mo.up("sum"))
                o.cost = mo.Variable(plan=[60.0] * 6, actual=[50.0, 55.0] + [None] * 4,
                                     display_name="Cost", regrain=mo.up("sum"))
                o.margin = ((o.rev - o.cost) / o.rev).set_display_name("Margin")
                o.gross = (o.rev - o.cost).set_display_name("Gross")
                o.cum = mo.cumsum(o.rev).set_display_name("Cum")
                o.prod = (o.rev * o.cost).set_display_name("Prod")
        return m

    @pytest.mark.parametrize("blend", [False, True], ids=["tracks", "blend"])
    def test_every_row_folds_on_its_own_track(self, blend):
        ws = _book(self._model(blend))["Ops"]
        cols = _totals(ws)
        rows = _row_values(ws, cols)
        for label, cells in rows.items():
            for cell in cells:
                assert "TrackValues" not in str(cell) and cell != "#VALUE!", (label, cell)
        line = {str(ws.cell(r, 1).value): r for r in range(1, ws.max_row + 1)}
        actual = rows["Margin · actual"][0]
        assert f"{ws.cell(1, cols[0]).column_letter}{line['Rev · actual']}" in actual, actual
        assert rows["Cum · actual"][1].startswith("="), rows["Cum · actual"]

    @pytest.mark.parametrize("view", ["blend", "rows"])
    def test_a_track_authored_with_its_own_formula_is_refused_as_the_engine_does(self, view):
        """``plan=<expr>`` without a rule: the engine refuses the line at a
        coarser grain, so its rows' subtotals are the teaching token too -
        never a fold against the display row's cells."""
        m = mo.Model("m", tracks=mo.Tracks("plan", "actual",
                                           blend=mo.blend(given="actual", follow="plan")),
                     default_grain="month", default_start="2026-01", default_periods=6,
                     default_excel_view=mo.ExcelView(timeline={"totals": ["quarter"]},
                                                     tracks=view))
        with m:
            m.a = mo.MultiVariable("A")
            with m.a as a:
                a.rev = mo.Variable(plan=[100.0] * 6, actual=[90.0, 95.0] + [None] * 4,
                                    display_name="Rev", regrain=mo.up("sum"))
                a.bonus = mo.Variable(plan=a.rev * 0.1, display_name="Bonus")
        ws = _book(m)["A"]
        for label, cells in _row_values(ws, _totals(ws)).items():
            if label.startswith("Bonus"):
                assert cells == ["#VALUE!", "#VALUE!"], (label, cells)

    def test_a_tracked_running_total_folds_per_track(self):
        cum = self._model(False).at("quarter").ops.cum.value
        assert {r: cum[r] for r in cum.roles} == {"plan": [300.0, 600.0], "actual": [185.0, 185.0]}

    @pytest.mark.slow
    @pytest.mark.parametrize("blend", [False, True], ids=["tracks", "blend"])
    def test_they_calculate_what_the_engine_does(self, blend, tmp_path):
        m = self._model(blend)
        q = m.at("quarter").ops
        ws = TestTheTotalsCalculate._recalculated(m, tmp_path)["Ops"]
        for label, cells in _row_values(ws, _totals(ws)).items():
            for got, want in zip(cells, _engine_row(q, label, blend)):
                if isinstance(want, str):
                    assert got == want, label
                else:
                    assert got == pytest.approx(want, rel=1e-12), label


class TestALineComputedMonthByMonthTotalsLive:
    """An IF, a product of two flows: the engine computes the line month by
    month and sums each bucket, so the bucket is the SUM of the line's own
    month cells - it was the engine's number typed in (``=3``), true until
    a month was edited in the file."""

    SUMMED = ("Flag", "Tax", "Sales", "Rounded")

    @staticmethod
    def _model(tracks=None):
        m = mo.Model("m", tracks=tracks, default_grain="month", default_start="2026-01",
                     default_periods=6,
                     default_excel_view=mo.ExcelView(timeline={"totals": ["quarter", "year"]}))
        with m:
            m.ops = mo.MultiVariable("Ops")
            with m.ops as o:
                if tracks is None:
                    o.rev = mo.Variable([200.0, 300.0] * 3, display_name="Rev",
                                        regrain=mo.up("sum"))
                else:
                    o.rev = mo.Variable(plan=[200.0, 300.0] * 3,
                                        actual=[260.0, 240.0] + [None] * 4,
                                        display_name="Rev", regrain=mo.up("sum"))
                o.units = mo.Variable([10.0, 20.0, 30.0] * 2, display_name="Units",
                                      regrain=mo.up("sum"))
                o.rate = mo.Variable(0.1, display_name="Rate")
                o.flag = mo.IF(mo.time.index >= 3, 1, 0).set_display_name("Flag")
                o.tax = mo.IF(o.rev > 250, o.rev * 0.2, 0).set_display_name("Tax")
                o.sales = (o.rev * o.units).set_display_name("Sales")
                o.rounded = mo.ROUND(o.rev / 7, 0).set_display_name("Rounded")
                o.over = (o.rev > 250).set_display_name("Over")
                o.fee = (o.rev * o.rate).set_display_name("Fee")
        return m

    @staticmethod
    def _sums(ws, row, cols):
        """Each total cell of ``row`` as the month columns its SUM lists,
        or None where it is no SUM of that row's own cells."""
        out = []
        for cell in (ws.cell(row, c).value for c in cols):
            got = re.fullmatch(rf"=SUM\(((?:[A-Z]+{row},?)+)\)", str(cell))
            out.append(None if got is None else
                       [ref.rstrip("0123456789") for ref in got.group(1).split(",")])
        return out

    @staticmethod
    def _months(ws, cols):
        """The month columns under each total column: the quarter's three,
        then the year's six."""
        letters = {c: get_column_letter(c) for c in range(1, ws.max_column + 1)}
        months = [letters[c] for c in range(cols[0] - 3, cols[-1]) if c not in cols]
        return [months[:3], months[3:6], months]

    @pytest.mark.parametrize("tracks", [
        None,
        mo.Tracks("plan", "actual"),
        mo.Tracks("plan", "actual", blend=mo.blend(given="actual", follow="plan")),
    ], ids=["untracked", "tracks", "blend"])
    def test_each_bucket_is_the_sum_of_its_own_months(self, tracks):
        ws = _book(self._model(tracks))["Ops"]
        cols = _totals(ws)
        rows = {str(ws.cell(r, 1).value): r for r in range(1, ws.max_row + 1)}
        summed = [label for label in rows if label.partition(" · ")[0] in self.SUMMED]
        assert len(summed) >= len(self.SUMMED)
        for label in summed:
            assert self._sums(ws, rows[label], cols) == self._months(ws, cols), label

    def test_a_line_of_its_operands_keeps_its_own_formula(self):
        # Rev x Rate re-computes over the quarter's revenue: its own
        # formula on the operands' bucket cells, as before.
        ws = _book(self._model())["Ops"]
        cols = _totals(ws)
        rows = {ws.cell(r, 1).value: r for r in range(1, ws.max_row + 1)}
        q1 = get_column_letter(cols[0])
        assert ws.cell(rows["Fee"], cols[0]).value == f"={q1}{rows['Rev']} * B{rows['Rate']}"

    def test_a_comparison_keeps_the_engines_count(self):
        # Excel's SUM skips a TRUE in a cell; the engine counts it as 1.
        ws = _book(self._model())["Ops"]
        cols = _totals(ws)
        row = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == "Over")
        assert self._sums(ws, row, cols) == [None, None, None]

    @pytest.mark.slow
    @pytest.mark.parametrize("blend", [None, False, True], ids=["untracked", "tracks", "blend"])
    def test_they_calculate_what_the_engine_does(self, blend, tmp_path):
        tracks = None if blend is None else (
            mo.Tracks("plan", "actual", blend=mo.blend(given="actual", follow="plan"))
            if blend else mo.Tracks("plan", "actual"))
        m = self._model(tracks)
        q, y = m.at("quarter").ops, m.at("year").ops
        ws = TestTheTotalsCalculate._recalculated(m, tmp_path)["Ops"]
        for label, cells in _row_values(ws, _totals(ws)).items():
            if label.partition(" · ")[0] not in self.SUMMED:
                continue
            want = _engine_row(q, label, blend) + _engine_row(y, label, blend)
            assert cells == pytest.approx(want, rel=1e-12), label


def test_a_kind_the_caller_did_not_project_is_projected_here(monkeypatch):
    """A caller may hand in its own projections (one that keeps a
    projection per kind across sheets, say); a kind missing there is projected from the sheet
    - it raised KeyError, and that sheet's subtotals lost every formula."""
    from modeleon.compile.excel import totals_writer

    expected = _grid(_book(_model())["Plan"])
    real = totals_writer.bucket_cells
    monkeypatch.setattr(totals_writer, "bucket_cells",
                        lambda *a, **k: real(*a, **{**k, "projected": {}}))
    assert _grid(_book(_model())["Plan"]) == expected


def test_a_tracked_line_shorter_than_its_sheet_keeps_the_book_alive():
    """A tracked IF line of six months on a sheet of twelve: its later
    quarters hold no month, and the check of its subtotals raised there -
    the whole export failed. Its quarters keep the engine's numbers; its
    year, all six months, is their SUM."""
    m = mo.Model("m", tracks=mo.Tracks("plan", "actual"), default_grain="month",
                 default_start="2026-01",
                 default_excel_view=mo.ExcelView(timeline={"totals": ["quarter", "year"]}))
    with m:
        m.ops = mo.MultiVariable("Ops")
        with m.ops as o:
            o.cost = mo.Variable([5.0] * 12, regrain=mo.up("sum"), display_name="Cost")
            o.rev = mo.Variable(plan=[200.0, 300.0] * 3, actual=[260.0, 240.0] + [None] * 4,
                                regrain=mo.up("sum"), display_name="Rev")
            o.tax = mo.IF(o.rev > 250, o.rev * 0.2, 0).set_display_name("Tax")
    ws = _book(m)["Ops"]
    tax = _row_values(ws, _totals(ws))["Tax"]
    assert tax[:2] == [60, 120], tax
    assert str(tax[-1]).startswith("=SUM("), tax


class TestARegrainedOperandKeepsItsBrackets:
    """A mean over the days, ``(a*31+b*28+c*31)/90``, or a compounded rate,
    ``EXP(...)-1``, read inside a longer formula: written bare, ``K / mean``
    divided by the days a second time and ``K * rate`` subtracted 1 after
    the product - Excel computed 0.011 for 90 and 926 for 27.27."""

    @staticmethod
    def _model():
        m = mo.Model("m", default_grain="month", default_start="2026-01", default_periods=12)
        with m:
            m.ops = mo.MultiVariable("Ops")
            with m.ops as a:
                a.price = mo.Variable([10.0] * 6 + [12.0] * 6, display_name="Price",
                                      regrain=mo.up("mean"))
                a.growth = mo.Variable([0.01] * 12, display_name="Growth",
                                       regrain=mo.up("geometric"))
                a.k = mo.Variable(900.0, display_name="K")
            m.q = mo.MultiVariable("By quarter", default_grain="quarter",
                                   default_start="2026-Q1", default_periods=4)
            with m.q as q:
                q.per_price = (m.ops.k / m.ops.price.at("quarter")).set_display_name("Per price")
                q.scaled = (m.ops.k * m.ops.growth.at("quarter")).set_display_name("Scaled")
        return m

    def test_the_operand_is_bracketed(self):
        ws = _book(self._model())["By quarter"]
        rows = {ws.cell(r, 1).value: r for r in range(1, ws.max_row + 1)}
        per_price = next(c.value for c in ws[rows["Per price"]][1:] if c.value is not None)
        scaled = next(c.value for c in ws[rows["Scaled"]][1:] if c.value is not None)
        assert re.search(r"/ \(\(.*\)/90\)$", per_price), per_price
        assert re.search(r"\* \(EXP\(.*\)-1\)$", scaled), scaled

    @pytest.mark.slow
    def test_they_calculate_what_the_engine_does(self, tmp_path):
        m = self._model()
        ws = TestTheTotalsCalculate._recalculated(m, tmp_path)["By quarter"]
        got = _row_values(ws, range(2, 6))
        assert got["Per price"] == pytest.approx(m.q.per_price.value, rel=1e-9)
        assert got["Scaled"] == pytest.approx(m.q.scaled.value, rel=1e-9)


class TestAConstantStepKeepsItsBrackets:
    """``x * (1 - share)``: the constant step reaches the coarser grain as a
    copy, and its subtotal lost the brackets - ``=E3 * 1 - A!B1``."""

    @staticmethod
    def _model():
        m = mo.Model("m", default_grain="month", default_start="2026-01", default_periods=6,
                     default_excel_view=mo.ExcelView(timeline={"totals": ["quarter"]}))
        with m:
            m.a = mo.MultiVariable("A")
            with m.a as a:
                a.share = mo.Variable(0.25, display_name="Share")
            m.b = mo.MultiVariable("B")
            with m.b as b:
                b.x = mo.Variable([100.0] * 6, display_name="X", regrain=mo.up("sum"))
                b.y = (b.x * (1 - m.a.share)).set_display_name("Y")
                b.run = mo.cumsum(b.x).set_display_name("Run")
            m.c = mo.MultiVariable("C")
            with m.c as c:
                c.w = (m.b.y * 2).set_display_name("W")
        return m

    def test_the_subtotal_keeps_them(self):
        book = _book(self._model())
        ws = book["B"]
        y = _row_values(ws, _totals(ws))["Y"]
        assert y[0] == "=E3 * (1 - A!B1)", y
        w = _row_values(book["C"], _totals(book["C"]))["W"]
        assert "(1 - A!B1)" in w[0], w

    def test_a_running_total_reads_its_own_previous_bucket(self):
        ws = _book(self._model())["B"]
        run = _row_values(ws, _totals(ws))["Run"]
        assert run[1].startswith("=E5 + "), run

    @pytest.mark.slow
    @pytest.mark.parametrize("tracked", [False, True], ids=["plain", "tracks"])
    def test_running_totals_inside_a_formula_sum_their_own_cells(self, tracked, tmp_path):
        """``cumsum(in) - cumsum(out)``: each running total sums its source's
        cells through the bucket - a range from the first to the last took in
        the month columns standing between the subtotals."""
        kw = dict(tracks=mo.Tracks("plan", "actual")) if tracked else {}
        m = mo.Model("m", default_grain="month", default_start="2026-01", default_periods=6,
                     default_excel_view=mo.ExcelView(timeline={"totals": ["quarter"]},
                                                     tracks="rows"), **kw)
        with m:
            m.a = mo.MultiVariable("A")
            with m.a as a:
                a.inc = (mo.Variable(plan=[100.0] * 6, actual=[90.0] * 6, display_name="In",
                                     regrain=mo.up("sum")) if tracked else
                         mo.Variable([100.0] * 6, display_name="In", regrain=mo.up("sum")))
                a.out = (mo.Variable(plan=[60.0] * 6, actual=[50.0] * 6, display_name="Out",
                                     regrain=mo.up("sum")) if tracked else
                         mo.Variable([60.0] * 6, display_name="Out", regrain=mo.up("sum")))
                a.cash = (mo.cumsum(a.inc) - mo.cumsum(a.out)).set_display_name("Cash")
        q = m.at("quarter").a
        ws = TestTheTotalsCalculate._recalculated(m, tmp_path)["A"]
        for label, cells in _row_values(ws, _totals(ws)).items():
            if label.startswith("Cash"):
                assert cells == pytest.approx(_engine_row(q, label, False)), label

    @pytest.mark.slow
    def test_they_calculate_what_the_engine_does(self, tmp_path):
        m = self._model()
        q = m.at("quarter")
        book = TestTheTotalsCalculate._recalculated(m, tmp_path)
        for sheet, lines in (("B", {"Y": q.b.y, "Run": q.b.run}), ("C", {"W": q.c.w})):
            ws = book[sheet]
            got = _row_values(ws, _totals(ws))
            for label, line in lines.items():
                assert got[label] == pytest.approx(line.value, rel=1e-12), label


class TestATrackedRatioFoldsTrackByTrack:
    """``mo.ratio`` on a line with tracks: sum over sum of the named lines'
    same-track series. The engine recomputed the formula over each
    operand's own rule (a mean denominator gave 6, the book 2)."""

    @staticmethod
    def _model(blend=False, den_tracked=True, typed=False):
        tracks = (mo.Tracks("plan", "actual",
                            blend=mo.blend(given="actual", follow="plan", until="2026-02"))
                  if blend else mo.Tracks("plan", "actual"))
        m = mo.Model("m", tracks=tracks, default_grain="month", default_start="2026-01",
                     default_periods=6,
                     default_excel_view=mo.ExcelView(timeline={"totals": ["quarter"]},
                                                     tracks="rows"))
        with m:
            m.ops = mo.MultiVariable("Ops")
            with m.ops as o:
                o.flow = mo.Variable(plan=[8.0] * 3 + [9.0] * 3,
                                     actual=[16.0, None, 16.0] + [None] * 3,
                                     display_name="Flow", regrain=mo.up("sum"))
                o.lvl = (mo.Variable(plan=[2.0, 4, 6, 1, 2, 3], actual=[1.0, 3.0] + [None] * 4,
                                     display_name="Lvl", regrain=mo.up("mean"))
                         if den_tracked else
                         mo.Variable([2.0, 4, 6, 1, 2, 3], display_name="Lvl",
                                     regrain=mo.up("mean")))
                if typed:
                    o.r = mo.Variable(plan=[1.0] * 6, actual=[1.0] * 6, display_name="R",
                                      regrain=mo.ratio("flow", "lvl"))
                else:
                    o.r = (o.flow / o.lvl).set_display_name("R").set_regrain(
                        mo.ratio("flow", "lvl"))
        return m

    def test_a_mean_denominator_sums_on_each_track(self):
        q = self._model().at("quarter").ops.r.value
        assert q["plan"] == [24.0 / 12.0, 27.0 / 6.0]
        assert q["actual"] == [32.0 / 4.0, "#DIV/0!"]

    def test_the_blend_divides_its_own_series(self):
        q = self._model(blend=True).at("quarter").ops.r.value
        assert q["live"] == [(16.0 + 8.0 + 8.0) / (1.0 + 3.0 + 6.0), 27.0 / 6.0]

    def test_a_line_without_tracks_divides_every_track(self):
        q = self._model(den_tracked=False).at("quarter").ops.r.value
        assert q["plan"] == [2.0, 4.5] and q["actual"] == [32.0 / 12.0, 0.0]

    def test_a_typed_ratio_line_with_tracks_projects(self):
        q = self._model(typed=True).at("quarter").ops.r.value
        assert q["plan"] == [2.0, 4.5] and q["actual"][0] == 8.0

    def test_the_book_reads_a_line_without_tracks_on_every_track(self):
        ws = _book(self._model(den_tracked=False))["Ops"]
        line = {str(ws.cell(r, 1).value): r for r in range(1, ws.max_row + 1)}
        actual = _row_values(ws, _totals(ws))["R · actual"][0]
        assert f"B{line['Flow · actual']}" in actual and f"B{line['Lvl']}" in actual, actual

    def test_a_constant_track_holds_in_every_period(self):
        m = mo.Model("m", tracks=mo.Tracks("plan", "actual"), default_grain="month",
                     default_start="2026-01", default_periods=6)
        with m:
            m.p = mo.MultiVariable("P")
            with m.p as o:
                o.flow = mo.Variable(plan=[8.0] * 6, actual=[16.0] * 6, regrain=mo.up("sum"))
                o.lvl = mo.Variable(plan=4.0, actual=[1.0, 2, 3, 4, 5, 6], regrain=mo.up("mean"))
                o.r = (o.flow / o.lvl).set_regrain(mo.ratio("flow", "lvl"))
        r = m.at("quarter").p.r.value
        assert r["plan"] == [2.0, 2.0] and r["actual"] == [8.0, 48.0 / 15.0]
        m.default_excel_view = mo.ExcelView(timeline={"totals": ["quarter"]}, tracks="rows")
        ws = _book(m)["P"]
        assert all("*3" in str(c) for c in _row_values(ws, _totals(ws))["R"]), \
            _row_values(ws, _totals(ws))["R"]

    def test_the_book_refuses_a_track_the_named_line_lacks_as_the_engine_does(self):
        m = mo.Model("m", tracks=mo.Tracks("plan", "actual"), default_grain="month",
                     default_start="2026-01", default_periods=3,
                     default_excel_view=mo.ExcelView(timeline={"totals": ["quarter"]},
                                                     tracks="rows"))
        with m:
            m.ops = mo.MultiVariable("Ops")
            with m.ops as o:
                o.flow = mo.Variable(plan=[8.0] * 3, actual=[16.0] * 3, display_name="Flow",
                                     regrain=mo.up("sum"))
                o.lvl = mo.Variable(plan=[2.0, 4, 6], display_name="Lvl", regrain=mo.up("mean"))
                o.r = mo.Variable(plan=[1.0] * 3, actual=[1.0] * 3, display_name="R",
                                  regrain=mo.ratio("flow", "lvl"))
        ws = _book(m)["Ops"]
        assert _row_values(ws, _totals(ws))["R · actual"] == ["#VALUE!"]

    @pytest.mark.slow
    @pytest.mark.parametrize("kw", [{}, {"blend": True}, {"den_tracked": False}],
                             ids=["tracks", "blend", "untracked-den"])
    def test_they_calculate_what_the_engine_does(self, kw, tmp_path):
        m = self._model(**kw)
        q = m.at("quarter").ops
        ws = TestTheTotalsCalculate._recalculated(m, tmp_path)["Ops"]
        for label, cells in _row_values(ws, _totals(ws)).items():
            if label != "R" and not label.startswith("R · "):
                continue
            for got, want in zip(cells, _engine_row(q, label, kw.get("blend", False))):
                if isinstance(want, str):
                    assert got == want, label
                else:
                    assert got == pytest.approx(want, rel=1e-12), label


class TestACompareColumnOfAFormulaLine:
    """``tracks='compare'``: the Var column of a formula line with tracks
    folds over its track columns' bucket cells - it was the ruleless
    '#VALUE!', its track columns having no fold of their own."""

    @pytest.mark.slow
    def test_var_folds_like_the_engine(self, tmp_path):
        m = mo.Model("m", tracks=mo.Tracks("plan", "actual"), default_grain="month",
                     default_start="2026-01", default_periods=6,
                     default_excel_view=mo.ExcelView(timeline={"totals": ["quarter"]},
                                                     tracks="compare"))
        with m:
            m.ops = mo.MultiVariable("Ops")
            with m.ops as o:
                o.rev = mo.Variable(plan=[100.0] * 6, actual=[90.0, 95.0, 99.0] + [None] * 3,
                                    display_name="Rev", regrain=mo.up("sum"))
                o.cost = mo.Variable(plan=[60.0] * 6, actual=[50.0, 55.0, 58.0] + [None] * 3,
                                     display_name="Cost", regrain=mo.up("sum"))
                o.gross = (o.rev - o.cost).set_display_name("Gross")
        q = m.at("quarter").ops.gross.value
        want = [f - p for f, p in zip(q["actual"], q["plan"])]
        book = _book(m)["Ops"]
        row = next(r for r in range(1, book.max_row + 1) if book.cell(r, 1).value == "Gross")
        var_cols = [c for c in range(1, book.max_column + 1)
                    if book.cell(3, c).value == "Var"
                    and any(str(book.cell(1, k).value or "").startswith("Total")
                            for k in range(max(1, c - 2), c + 1))]
        assert all(str(book.cell(row, c).value).startswith("=") for c in var_cols)
        ws = TestTheTotalsCalculate._recalculated(m, tmp_path)["Ops"]
        assert [ws.cell(row, c).value for c in var_cols] == pytest.approx(want)


class TestADayWindowStaysWithinExcelsLimits:
    def test_a_year_of_days_folds_in_ranges(self):
        m = mo.Model("m", default_grain="day", default_start="2026-01-01", default_periods=365,
                     default_excel_view=mo.ExcelView(timeline={"totals": ["year"]}))
        with m:
            m.ops = mo.MultiVariable("Ops")
            with m.ops as o:
                o.high = mo.Variable([float(i % 7) for i in range(365)], display_name="High",
                                     regrain=mo.up("max"))
                o.g = mo.Variable([0.0001] * 365, display_name="G", regrain=mo.up("geometric"))
        ws = _book(m)["Ops"]
        for row in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                v = ws.cell(row, c).value
                if isinstance(v, str) and v.startswith("="):
                    assert len(v) < 8192 and v.count(",") < 255, (row, c, v[:80])


def _grid(ws):
    return [
        [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        for r in range(1, ws.max_row + 1)
    ]


class TestEveryTabTotalsAlike:
    """A tab is a tab however it came to be: marked with
    ``excel_props={'tab': True}``, made by the writer from an unmarked
    first-level container (the default), the overview tab that collects
    Variables placed straight on the model, or the tab of a container
    written on its own with ``container.to_excel()``. Each one's total
    columns hold the same cells."""

    def _lines(self, p):
        p.rate = mo.Variable(0.1, display_name="Rate")
        p.revenue = mo.Variable(
            [10, 20, 30, 40, 50, 60], display_name="Revenue",
            regrain=mo.up("sum"),
        )
        p.tax = p.revenue * p.rate
        p.fx_rate = mo.Variable([1.1] * 6, display_name="FX rate")
        p.balance = mo.Variable(
            [1, 2, 3, 4, 5, 6], display_name="Balance", regrain=mo.up("last"),
        )
        p.gated = mo.IF(p.revenue > 25, p.revenue, 0)

    @staticmethod
    def _empty(display_name="M"):
        m = mo.Model(
            "m", display_name=display_name,
            default_excel_view=mo.ExcelView(
                timeline={"totals": ["quarter", "year"]}
            ),
        )
        m.default_start, m.default_grain, m.default_periods = (
            "2026-01", "month", 6,
        )
        return m

    def _model(self, marked, section=True):
        m = self._empty()
        with m:
            m.plan = mo.MultiVariable("Plan", **_tab(marked))
            with m.plan as p:
                self._lines(p)
                if section:
                    p.details = mo.MultiVariable("Details")
                    with p.details as d:
                        d.double = p.tax * 2
        return m

    def test_an_unmarked_tab_writes_every_cell_a_marked_tab_writes(self):
        marked = _grid(_book(self._model(True))["Plan"])
        unmarked = _grid(_book(self._model(False))["Plan"])
        assert unmarked == marked
        # …and the cells are the live ones, not a shared emptiness.
        ws = _book(self._model(False))["Plan"]
        rows = {ws.cell(r, 1).value: r for r in range(1, ws.max_row + 1)}
        tax, fx_rate = rows["Tax"], rows["FX rate"]
        assert ws[f"F{tax}"].value == f"=F{rows['Revenue']} * B{rows['Rate']}"
        assert ws[f"K{fx_rate}"].value == "#VALUE!"
        assert ws[f"F{rows['Double']}"].value == f"=F{tax} * 2"

    def test_the_overview_tab_of_root_variables(self):
        # Lines placed straight on the model land on a tab named after
        # it — the same tab a marked container of those lines makes.
        m = self._empty("M")
        with m:
            self._lines(m)
        overview = _grid(_book(m)["M"])

        tabbed = self._empty("Book")
        with tabbed:
            tabbed.m = mo.MultiVariable("M", excel_props={"tab": True})
            with tabbed.m as p:
                self._lines(p)
        assert overview == _grid(_book(tabbed)["M"])
        assert overview[5][5] == "=F5 * B4"      # Tax, Total Q1
        assert overview[6][10] == "#VALUE!"      # FX rate, Total 2026

    @both_tabs
    def test_a_container_written_on_its_own(self, marked):
        # ``m.plan.to_excel()`` writes the container's own lines onto one
        # tab — the same tab the whole model's book carries. (A section
        # inside it would become a tab of its own there, so none here.)
        whole = _grid(_book(self._model(True, section=False))["Plan"])
        m2 = self._model(marked, section=False)
        path = os.path.join(tempfile.mkdtemp(), "t.xlsx")
        m2.plan.to_excel(path)
        alone = openpyxl.load_workbook(path)["Plan"]
        assert _grid(alone) == whole
        assert alone["F6"].value == "=F5 * B4"


class TestTheTotalsCalculate:
    """Recalculated by a spreadsheet program, a formula line's total
    column shows the number Python gets by re-graining the model:
    ``model.at('quarter')`` — the line's own formula over the quarter's
    operands, never the sum of its monthly results."""

    @staticmethod
    def _recalculated(model, tmp_path):
        import shutil
        import subprocess
        from pathlib import Path

        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        mac = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
        if soffice is None and mac.exists():
            soffice = str(mac)
        if soffice is None:
            pytest.skip("LibreOffice is not installed")
        src = tmp_path / "book.xlsx"
        model.to_excel(str(src))
        outdir = tmp_path / "calc"
        profile = (tmp_path / "lo-profile").as_uri()  # private per run
        subprocess.run(
            [soffice, f"-env:UserInstallation={profile}", "--headless",
             "--convert-to", "xlsx", "--outdir", str(outdir), str(src)],
            check=True, capture_output=True, timeout=120,
        )
        return openpyxl.load_workbook(str(outdir / "book.xlsx"), data_only=True)

    @both_tabs
    @pytest.mark.slow
    def test_a_ratio_totals_to_the_ratio_of_totals(self, marked, tmp_path):
        m = _model(marked)
        with m.plan as p:
            p.costs = mo.Variable(
                [9.0, 12, 15, 38, 30, 33], display_name="Costs",
                regrain=mo.up("sum"),
            )
            p.margin = (p.revenue - p.costs) / p.revenue
        ws = self._recalculated(m, tmp_path)["Plan"]
        row = next(r for r in range(1, ws.max_row + 1)
                   if ws.cell(r, 1).value == "Margin")
        quarters = m.at("quarter").plan.margin.value
        year = m.at("year").plan.margin.value
        # Q1 = (60 − 36) / 60; the months' own margins would add up to 1.0.
        assert quarters[0] == pytest.approx(0.4)
        assert ws.cell(row, 6).value == pytest.approx(quarters[0])   # F
        assert ws.cell(row, 10).value == pytest.approx(quarters[1])  # J
        assert ws.cell(row, 11).value == pytest.approx(year[0])      # K


class TestTheTotalsWord:
    def test_the_model_speaks_its_own_language(self):
        # Any word but the default "Total" proves the override; German here.
        m = mo.Model(
            "m",
            display_name="M",
            default_excel_view=mo.ExcelView(
                timeline={"totals": ["quarter"], "totals_word": "Gesamt"}
            ),
        )
        m.default_start, m.default_grain, m.default_periods = (
            "2026-01", "month", 3,
        )
        with m:
            m.plan = mo.MultiVariable("Plan", excel_props={"tab": True})
            with m.plan as p:
                p.flow = mo.Variable(
                    [1, 2, 3], display_name="Flow", regrain=mo.up("sum"),
                )
        ws = _book(m)["Plan"]
        # One tier above the months, so the quarter's own total names
        # itself on that tier and merges down through the month row.
        assert ws.cell(1, 5).value == "Gesamt Q1"
        assert "E1:E2" in {str(r) for r in ws.merged_cells.ranges}

    def test_tier_spans_read_from_the_left(self):
        # "2026" centered floats far from the row names — and the
        # month row stays centered under it.
        ws = _book(_model())["Plan"]
        assert ws.cell(1, 3).alignment.horizontal == "left"
        assert ws.cell(2, 3).alignment.horizontal == "left"
        assert ws.cell(3, 3).alignment.horizontal == "center"


class TestNativeColumnGroups:
    def test_the_file_carries_real_outline_levels(self):
        # Excel's own "collapse 2027": months sit two buckets deep,
        # a quarter's total one (inside the year), the year's total
        # outside every group — and the summary stands RIGHT of its
        # detail, where our totals are.
        ws = _book(_model())["Plan"]
        def level(c):
            return (ws.column_dimensions[c].outline_level
                    if c in ws.column_dimensions else 0)

        assert [level(c) for c in "CDEF"] == [2, 2, 2, 1]
        assert level("K") == 0  # "Total 2026" survives every fold
        assert ws.sheet_properties.outlinePr.summaryRight is True


class TestTotalsOverride:
    def test_the_override_replaces_and_strips_the_declaration(self):
        # A caller-side override: LayoutEngine(totals_override=...) beats
        # the declared timeline.totals — ('quarter',) narrows, () strips,
        # None keeps. ``to_excel`` never passes it.
        from modeleon.compile.excel.layout import LayoutEngine

        m = _model()
        roots = [m.plan]
        eng = LayoutEngine(roots, totals_override=("quarter",))
        eng.compute_addresses()
        plan = eng.totals_by_sheet.get("Plan")
        assert plan is not None and plan.kinds == ("quarter",)

        eng2 = LayoutEngine(roots, totals_override=())
        eng2.compute_addresses()
        assert eng2.totals_by_sheet.get("Plan") is None

        eng3 = LayoutEngine(roots)
        eng3.compute_addresses()
        assert eng3.totals_by_sheet.get("Plan").kinds == ("quarter", "year")


class TestRangesNeverSwallowATotal:
    """A row's cells stop being contiguous the moment totals declare a
    column between them — and ``SUM(C3:I3)`` then counts the quarter
    twice, in a file that looks perfectly valid.

    ``totals.rule_formula`` lists explicit refs for its own buckets;
    the same rule holds wherever a range is built.
    """

    def _book(self, totals, n=6, fn=None):
        import os
        import tempfile

        import openpyxl

        view = mo.ExcelView(timeline={"totals": totals}) if totals else None
        m = mo.Model("m", display_name="M",
                     **({"default_excel_view": view} if view else {}))
        m.default_start, m.default_grain, m.default_periods = (
            "2026-01", "month", n,
        )
        with m:
            m.p = mo.MultiVariable("P", excel_props={"tab": True})
            with m.p as p:
                p.flow = mo.Variable([10.0] * n, display_name="Flow",
                                     regrain=mo.up("sum"))
                p.total_flow = mo.Variable((fn or mo.SUM)(p.flow),
                                           display_name="Total flow")
        path = os.path.join(tempfile.mkdtemp(), "t.xlsx")
        m.to_excel(path)
        ws = openpyxl.load_workbook(path)["P"]
        row = next(r for r in range(1, ws.max_row + 1)
                   if ws.cell(r, 1).value == "Total flow")
        return ws.cell(row, 2).value

    def test_an_interleaved_row_lists_its_cells(self):
        # Months C,D,E then the Q1 total at F, then G,H,I. The range
        # form would add F twice over: 90 for a model that says 60.
        formula = self._book(["quarter"])
        assert ":" not in formula
        assert formula.replace(" ", "") == "=SUM(C3,D3,E3,G3,H3,I3)"

    def test_an_interleaved_row_on_another_sheet_qualifies_every_cell(self):
        # A list of cells is not a range: a sheet prefix binds to one
        # cell only. Unqualified, every cell after the first would be
        # read from the CURRENT sheet and the sum would silently be wrong.
        import os
        import tempfile

        import openpyxl

        m = mo.Model("m", display_name="M", default_excel_view=mo.ExcelView(
            timeline={"totals": ["quarter"]}))
        m.default_start, m.default_grain, m.default_periods = (
            "2026-01", "month", 6,
        )
        with m:
            m.p = mo.MultiVariable("P", excel_props={"tab": True})
            with m.p as p:
                p.flow = mo.Variable([10.0] * 6, display_name="Flow",
                                     regrain=mo.up("sum"))
            m.s = mo.MultiVariable("S", excel_props={"tab": True})
            with m.s as s:
                s.total_flow = mo.Variable(mo.SUM(m.p.flow), display_name="Total flow")
        path = os.path.join(tempfile.mkdtemp(), "t.xlsx")
        m.to_excel(path)
        ws = openpyxl.load_workbook(path)["S"]
        row = next(r for r in range(1, ws.max_row + 1)
                   if ws.cell(r, 1).value == "Total flow")
        formula = next(ws.cell(row, c).value for c in range(2, ws.max_column + 1)
                       if str(ws.cell(row, c).value or "").startswith("="))
        assert formula.startswith("=SUM(") and ":" not in formula
        cells = [a.strip() for a in formula[len("=SUM("):-1].split(",")]
        assert len(cells) == 6
        assert all(c.startswith("P!") for c in cells), formula

    def test_a_contiguous_row_keeps_its_range(self):
        # Nothing stands between the months — the compact form is both
        # correct and what a reader expects to see.
        assert self._book(None) == "=SUM(C2:H2)"

    def test_a_range_only_function_tells_the_truth_instead(self):
        # IRR's second argument is a GUESS, so a comma list would be
        # read as one. A formula that computes something else is worse
        # than the number itself: the cell falls back to its value.
        import os
        import tempfile

        import openpyxl

        m = mo.Model("m", display_name="M",
                     default_excel_view=mo.ExcelView(
                         timeline={"totals": ["quarter"]}))
        m.default_start, m.default_grain, m.default_periods = (
            "2026-01", "month", 6,
        )
        with m:
            m.p = mo.MultiVariable("P", excel_props={"tab": True})
            with m.p as p:
                p.flow = mo.Variable([-100.0, 30.0, 30.0, 30.0, 30.0, 30.0],
                                     display_name="Flow",
                                     regrain=mo.up("sum"))
                p.rate = mo.Variable(mo.IRR(p.flow), display_name="Rate")
        path = os.path.join(tempfile.mkdtemp(), "t.xlsx")
        m.to_excel(path)
        ws = openpyxl.load_workbook(path)["P"]
        row = next(r for r in range(1, ws.max_row + 1)
                   if ws.cell(r, 1).value == "Rate")
        assert not str(ws.cell(row, 2).value).startswith("=")

    # A SLICE of the row (``SUM(x[0:4])``: January to April) is the same
    # question over fewer cells: its window still steps over the Q1
    # total, and the range form would count it in.

    def _sliced(self, fn, flows, *, other_sheet=False):
        """Write ``Result = fn(flow)`` next to the flows, or on a tab of
        its own; return the book, what the "Result" cell holds, and the
        value the model computed for it."""
        m = mo.Model("m", display_name="M", default_excel_view=mo.ExcelView(
            timeline={"totals": ["quarter"]}))
        m.default_start, m.default_grain, m.default_periods = (
            "2026-01", "month", len(flows),
        )
        with m:
            m.p = mo.MultiVariable("P", excel_props={"tab": True})
            with m.p as p:
                p.flow = mo.Variable(flows, display_name="Flow",
                                     regrain=mo.up("sum"))
                if not other_sheet:
                    p.result = mo.Variable(fn(p.flow), display_name="Result")
            if other_sheet:
                m.s = mo.MultiVariable("S", excel_props={"tab": True})
                with m.s as s:
                    s.result = mo.Variable(fn(m.p.flow), display_name="Result")
        home = m.s if other_sheet else m.p
        path = os.path.join(tempfile.mkdtemp(), "t.xlsx")
        m.to_excel(path)
        wb = openpyxl.load_workbook(path)
        ws = wb["S" if other_sheet else "P"]
        row = next(r for r in range(1, ws.max_row + 1)
                   if ws.cell(r, 1).value == "Result")
        return wb, ws.cell(row, 2).value, home.result.value

    @staticmethod
    def _sum_of_listed_cells(wb, formula, sheet):
        # Read ``=SUM(a, b, …)`` the way Excel does: each argument is
        # one cell, on its own sheet when it names one.
        assert formula.startswith("=SUM(") and ":" not in formula, formula
        total = 0.0
        for ref in formula[len("=SUM("):-1].split(","):
            where, _, addr = ref.strip().rpartition("!")
            total += wb[where or sheet][addr].value
        return total

    def test_a_sliced_window_lists_its_cells(self):
        # C,D,E are Jan–Mar, F the Q1 total, G April. ``C3:G3`` would
        # read 160 for a model that says 100.
        wb, formula, value = self._sliced(
            lambda x: mo.SUM(x[0:4]), [10.0, 20.0, 30.0, 40.0, 50.0, 60.0])
        assert value == 100.0
        assert formula.replace(" ", "") == "=SUM(C3,D3,E3,G3)"
        assert self._sum_of_listed_cells(wb, formula, "P") == value

    def test_a_sliced_window_on_another_sheet_qualifies_every_cell(self):
        wb, formula, value = self._sliced(
            lambda x: mo.SUM(x[0:4]), [10.0, 20.0, 30.0, 40.0, 50.0, 60.0],
            other_sheet=True)
        cells = [a.strip() for a in formula[len("=SUM("):-1].split(",")]
        assert len(cells) == 4
        assert all(c.startswith("P!") for c in cells), formula
        assert self._sum_of_listed_cells(wb, formula, "S") == value == 100.0

    def test_a_sliced_window_inside_one_quarter_keeps_its_range(self):
        wb, formula, value = self._sliced(
            lambda x: mo.SUM(x[3:6]), [10.0, 20.0, 30.0, 40.0, 50.0, 60.0])
        assert formula == "=SUM(G3:I3)"
        assert value == 150.0

    def test_a_sliced_window_inside_one_quarter_on_another_sheet(self):
        # A range names its sheet once, in front of the first cell.
        _wb, formula, value = self._sliced(
            lambda x: mo.SUM(x[3:6]), [10.0, 20.0, 30.0, 40.0, 50.0, 60.0],
            other_sheet=True)
        assert formula == "=SUM(P!F3:H3)"
        assert value == 150.0

    def test_a_stepped_window_lists_its_cells(self):
        # Every other month is never side by side, totals or not.
        wb, formula, value = self._sliced(
            lambda x: mo.SUM(x[0:6:2]), [10.0, 20.0, 30.0, 40.0, 50.0, 60.0])
        assert value == 90.0
        assert self._sum_of_listed_cells(wb, formula, "P") == value

    def test_a_sliced_window_for_a_range_only_function_is_its_value(self):
        # ``IRR(C3:H3, 0.1)`` would take the Q1 total as a cash flow.
        _wb, cell, value = self._sliced(
            lambda x: mo.IRR(x[0:5]), [-100.0, 30.0, 30.0, 30.0, 30.0, 30.0])
        assert not str(cell).startswith("="), cell
        assert cell == pytest.approx(value)

    # A ROLLING window (``rolling_sum`` / ``rolling_mean`` build a
    # RollingAggregate) asks the same question once a month: the three
    # months ending in April are Feb, Mar and Apr, and the Q1 total
    # stands between Mar and Apr. One range from Feb to Apr added it in.

    _FLOWS = [10.0, 20.0, 30.0, 40.0, 50.0, 60.0]

    def _rolling(self, func, fold, *, other_sheet=False):
        """Write a three-month rolling ``func`` of the flows. Return the
        book, the sheet of the "Rolling" row, the flows' row and month
        columns, and month by month what the "Rolling" cell holds next to
        the model's value."""
        from modeleon.extend import RollingAggregate

        flows = self._FLOWS
        values = [0.0, 0.0] + [fold(flows[i - 2:i + 1]) for i in range(2, len(flows))]
        m = mo.Model("m", display_name="M", default_excel_view=mo.ExcelView(
            timeline={"totals": ["quarter"]}))
        m.default_start, m.default_grain, m.default_periods = (
            "2026-01", "month", len(flows),
        )
        with m:
            m.p = mo.MultiVariable("P", excel_props={"tab": True})
            with m.p as p:
                p.flow = mo.Variable(flows, display_name="Flow",
                                     regrain=mo.up("sum"))
                if not other_sheet:
                    p.rolling = mo.Variable(values, display_name="Rolling")
            if other_sheet:
                m.s = mo.MultiVariable("S", excel_props={"tab": True})
                with m.s as s:
                    s.rolling = mo.Variable(values, display_name="Rolling")
        home = m.s if other_sheet else m.p
        home.rolling._expr = RollingAggregate(
            source=m.p.flow, window=3, func=func, fill=0.0)
        path = os.path.join(tempfile.mkdtemp(), "t.xlsx")
        m.to_excel(path)
        wb = openpyxl.load_workbook(path)

        def months(ws):
            header = next(r for r in range(1, ws.max_row + 1)
                          if any(ws.cell(r, c).value == "Jan"
                                 for c in range(1, ws.max_column + 1)))
            return [c for c in range(1, ws.max_column + 1)
                    if ws.cell(header, c).value in ("Jan", "Feb", "Mar", "Apr", "May", "Jun")]

        def row(ws, label):
            return next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == label)

        sheet = "S" if other_sheet else "P"
        ws, wp = wb[sheet], wb["P"]
        cells = [ws.cell(row(ws, "Rolling"), c).value for c in months(ws)]
        return wb, sheet, row(wp, "Flow"), months(wp), list(zip(cells, values))

    @staticmethod
    def _cells_read(formula, sheet):
        """The cells ``=FUNC(a, b:c, ...)`` reads as (sheet, column, row),
        a range spelled out cell by cell."""
        from openpyxl.utils.cell import range_boundaries

        _func, _, args = formula[1:].partition("(")
        out = []
        for arg in args[:-1].split(","):
            where, _, ref = arg.strip().rpartition("!")
            min_col, min_row, max_col, max_row = range_boundaries(ref)
            out += [(where.strip("'") or sheet, c, r)
                    for r in range(min_row, max_row + 1)
                    for c in range(min_col, max_col + 1)]
        return out

    @pytest.mark.parametrize("func, fold", [
        ("SUM", sum),
        ("AVERAGE", lambda xs: sum(xs) / len(xs)),
    ], ids=["sum", "mean"])
    def test_a_rolling_window_reads_its_own_months(self, func, fold):
        wb, sheet, flow_row, months, cells = self._rolling(func, fold)
        for i, (cell, value) in enumerate(cells[2:], start=2):
            read = self._cells_read(cell, sheet)
            assert read == [(sheet, c, flow_row) for c in months[i - 2:i + 1]], (i, cell)
            assert fold([wb[s].cell(r, c).value for s, c, r in read]) == pytest.approx(value)
        # Jan..Mar sits inside Q1: the compact range. Feb..Apr steps over
        # the Q1 total: its cells, listed.
        assert ":" in cells[2][0] and ":" not in cells[3][0]

    def test_a_rolling_window_on_another_sheet_qualifies_every_cell(self):
        wb, sheet, flow_row, months, cells = self._rolling("SUM", sum, other_sheet=True)
        april, value = cells[3]
        assert april.count("P!") == 3 and ":" not in april, april
        read = self._cells_read(april, sheet)
        assert read == [("P", c, flow_row) for c in months[1:4]]
        assert sum(wb[s].cell(r, c).value for s, c, r in read) == value

    def test_a_rolling_window_a_list_cannot_feed_is_its_value(self):
        # MEDIAN is not among the functions that read a list of cells the
        # way they read a range: across the Q1 total the cell carries the
        # model's number, inside one quarter it stays a live range.
        import statistics

        _wb, _sheet, _row, _months, cells = self._rolling("MEDIAN", statistics.median)
        (mar, _), (apr, apr_value), (may, may_value), (jun, _) = cells[2:]
        assert str(mar).startswith("=MEDIAN(") and str(jun).startswith("=MEDIAN(")
        assert apr == pytest.approx(apr_value) and may == pytest.approx(may_value)


class TestADurationIsADayCount:
    """A spreadsheet stores a date as a number of days, so a duration
    added to a date is a plain number of days in the formula — never
    Python's ``5 days, 0:00:00``, which no spreadsheet can read."""

    def _formula(self, delta):
        m = mo.Model("m", display_name="M")
        with m:
            m.p = mo.MultiVariable("P", excel_props={"tab": True})
            with m.p as p:
                p.start = mo.Variable(date(2026, 1, 1), display_name="Start")
                p.due = mo.Variable(p.start + delta, display_name="Due")
        ws = _book(m)["P"]
        assert ws["A1"].value == "Start" and ws["A2"].value == "Due"
        return ws["B2"].value

    def test_whole_days(self):
        assert self._formula(timedelta(days=90)) == "=B1 + 90"

    def test_part_of_a_day(self):
        assert self._formula(timedelta(hours=12)) == "=B1 + 0.5"

    def test_a_negative_duration(self):
        assert self._formula(timedelta(days=-5)) == "=B1 + -5"


class TestNativeOutlineGroups:
    """A collapsible outline, native to the file — both axes."""

    def test_rows_group_by_section_and_the_header_is_the_summary(self):
        m = mo.Model("m", display_name="M")
        m.default_start, m.default_grain, m.default_periods = (
            "2026-01", "month", 3,
        )
        with m:
            m.p = mo.MultiVariable("P", excel_props={"tab": True})
            with m.p as p:
                p.income = mo.MultiVariable("I. Income")
                with p.income as d:
                    d.revenue = mo.Variable([1, 2, 3], display_name="Revenue")
                    d.details = mo.MultiVariable("Details")
                    with d.details as dd:
                        dd.fx_rate = mo.Variable([1.1] * 3, display_name="FX rate")
        ws = _book(m)["P"]
        rows = {ws.cell(r, 1).value: r for r in range(1, ws.max_row + 1)}
        # Section content is one level deep; nested content two.
        assert ws.row_dimensions[rows["Revenue"]].outline_level == 1
        assert ws.row_dimensions[rows["    FX rate"] if "    FX rate" in rows
                                 else rows[next(k for k in rows if k and "FX rate" in k)]
                                 ].outline_level == 2
        # The header row is the SUMMARY, outside its own group…
        assert ws.row_dimensions[rows["I. Income"]].outline_level in (0, None)
        # …and Excel puts the ± on it because the summary sits above.
        assert ws.sheet_properties.outlinePr.summaryBelow is False

    def test_column_brackets_cover_the_whole_group(self):
        # Compare mode: a month is a GROUP of columns, and the outline
        # bracket must cover all of them — one column in four reads as
        # noise. Hidden operands sit one level deeper so expanding a
        # quarter does not unhide them.
        m = mo.Model(
            "m", display_name="M", default_grain="month",
            default_start="2026-01", default_periods=3,
            tracks=mo.Tracks(
                "plan", "actual",
                blend=mo.blend(given="actual", follow="plan",
                               until="2026-01", name="live")),
            default_excel_view=mo.ExcelView(
                tracks="compare",
                timeline={"totals": ["quarter"]}),
        )
        with m:
            m.p = mo.MultiVariable("P", excel_props={"tab": True})
            with m.p as p:
                p.flow = mo.Variable(plan=[10.0] * 3,
                                     actual=[12.0, None, None],
                                     display_name="Flow",
                                     regrain=mo.up("sum"))
        ws = _book(m)["P"]
        # Group = actual | plan | live | Var (stride 4); every month
        # column of January's group carries the month level.
        lv = {c: ws.column_dimensions[get_column_letter(c)].outline_level
              for c in range(2, 10)}
        assert lv[2] == lv[3] == lv[4] == lv[5] == 1
