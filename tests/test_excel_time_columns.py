"""A period header names period columns — and nothing else.

Two rules, one idea: the time axis is a property of the LINE, not of
the model. A model declaring a window doesn't make every sheet a
time series, and it doesn't make a constant a January figure.

  1. No period header on a sheet that carries no per-period line.
     A sheet of constants (a price list, say) must not get
     "Jan 2026" over column B — its values would read as January
     figures.
  2. On a sheet that carries BOTH, constants get a column of their
     own, left of the periods, and the header starts after it.

The addresses move with the rule, so formulas follow for free — that
is what makes this a layout change rather than a cosmetic one.
"""

import os
import tempfile

import openpyxl
import pytest

import modeleon as mo


def _sheet(model, name):
    path = os.path.join(tempfile.mkdtemp(), "t.xlsx")
    model.to_excel(path)
    return openpyxl.load_workbook(path)[name]


def _grid(ws, cols=6):
    return [
        [ws.cell(r, c).value for c in range(1, cols + 1)]
        for r in range(1, ws.max_row + 1)
    ]


def _windowed():
    m = mo.Model("m", display_name="M")
    m.default_start, m.default_grain, m.default_periods = "2026-01", "month", 3
    return m


class TestTheHeaderFollowsContent:
    def test_a_register_of_constants_gets_no_period_header(self):
        m = _windowed()
        with m:
            m.roles = mo.MultiVariable("Roles", excel_props={"tab": True})
            with m.roles as d:
                d.band = mo.Variable(500, display_name="Band")
                d.grade = mo.Variable(3, display_name="Grade")
            # A second sheet DOES carry periods — the window is real,
            # so this pins content and not "the model has no window".
            m.plan = mo.MultiVariable("Plan", excel_props={"tab": True})
            with m.plan as p:
                p.revenue = mo.Variable([1, 2, 3], display_name="Revenue")
        g = _grid(_sheet(m, "Roles"))
        assert g[0][:2] == ["Band", 500]
        assert not any("2026" in str(c) for row in g for c in row if c)

    def test_the_sheet_that_has_periods_keeps_its_header(self):
        m = _windowed()
        with m:
            m.plan = mo.MultiVariable("Plan", excel_props={"tab": True})
            with m.plan as p:
                p.revenue = mo.Variable([1, 2, 3], display_name="Revenue")
        assert _grid(_sheet(m, "Plan"))[0][1:4] == [
            "Jan 2026",
            "Feb 2026",
            "Mar 2026",
        ]

    def test_two_sheets_are_judged_separately(self):
        # Content belongs to the sheet that prints it: one tab having
        # periods must not put a header over another that hasn't.
        m = _windowed()
        with m:
            m.rates = mo.MultiVariable("Rates", excel_props={"tab": True})
            with m.rates as c:
                c.usd = mo.Variable(1.08, display_name="USD")
            m.plan = mo.MultiVariable("Plan", excel_props={"tab": True})
            with m.plan as p:
                p.series = mo.Variable([1, 2, 3], display_name="Series")
        assert not any(
            "2026" in str(c)
            for row in _grid(_sheet(m, "Rates"))
            for c in row
            if c
        )
        assert _grid(_sheet(m, "Plan"))[0][1] == "Jan 2026"


class TestConstantsKeepTheirOwnColumn:
    def _mixed(self):
        m = _windowed()
        with m:
            m.plan = mo.MultiVariable("Plan", excel_props={"tab": True})
            with m.plan as p:
                p.rate = mo.Variable(0.12, display_name="Rate")
                p.revenue = mo.Variable([100, 110, 120], display_name="Revenue")
                p.commission = p.revenue * p.rate
        return m

    def test_a_constant_is_not_a_january_figure(self):
        ws = _sheet(self._mixed(), "Plan")
        g = _grid(ws)
        # Periods start at C; B is the constants column.
        assert g[0][1] is None
        assert g[0][2:5] == ["Jan 2026", "Feb 2026", "Mar 2026"]
        assert g[1][:3] == ["Rate", 0.12, None]
        assert g[2][1:5] == ["Revenue", 100, 110, 120][1:] or g[2][2:5] == [
            100,
            110,
            120,
        ]

    def test_formulas_follow_the_moved_cells(self):
        # The point of moving ADDRESSES rather than painting: a
        # reference to the constant still lands on it.
        g = _grid(_sheet(self._mixed(), "Plan"))
        commission = next(row for row in g if row[0] == "Commission")
        assert commission[2] == "=C3 * B2"

    def test_the_pane_freezes_left_of_the_periods(self):
        # The constants column must stay in view while the periods
        # scroll — it is a lead column in everything but name.
        assert _sheet(self._mixed(), "Plan").freeze_panes == "C2"

    def test_no_empty_column_when_the_sheet_is_all_periods(self):
        m = _windowed()
        with m:
            m.plan = mo.MultiVariable("Plan", excel_props={"tab": True})
            with m.plan as p:
                p.revenue = mo.Variable([1, 2, 3], display_name="Revenue")
        g = _grid(_sheet(m, "Plan"))
        assert g[0][1] == "Jan 2026"
        assert g[1][:4] == ["Revenue", 1, 2, 3]

    def test_lead_metadata_columns_still_come_first(self):
        m = _windowed()
        m.default_excel_view = mo.ExcelView(
            meta={"article": True, "unit": "Unit", "label": "Name"}
        )
        with m:
            m.plan = mo.MultiVariable("Plan", excel_props={"tab": True})
            with m.plan as p:
                p.rate = mo.Variable(0.12, display_name="Rate")
                p.revenue = mo.Variable([1, 2, 3], display_name="Revenue")
        g = _grid(_sheet(m, "Plan"), cols=8)
        # A = article no., B = name, C = unit, D = constants, E.. = periods.
        assert g[0][1] == "Name" and g[0][2] == "Unit"
        assert g[0][4] == "Jan 2026"
        rate = next(row for row in g if row[1] == "Rate")
        assert rate[3] == pytest.approx(0.12)


class TestProjectedFieldColumns:
    """Any field can ride its own lead column, by analogy with the
    unit: ``meta={'fields': [('description', 'Comment')]}``.

    A list of PAIRS rather than free-standing keys because a view is a
    real MV tree — a field named ``description`` would collide with
    ``MultiVariable.description`` the moment it became an attribute of
    the meta node.
    """

    def _model(self, fields):
        m = _windowed()
        m.default_excel_view = mo.ExcelView(
            meta={"unit": "Unit", "label": "Name", "fields": fields}
        )
        with m:
            m.plan = mo.MultiVariable("Plan", excel_props={"tab": True})
            with m.plan as p:
                p.rate = mo.Variable(
                    0.12,
                    display_name="Rate",
                    unit="%",
                    description="assumed rate",
                )
                p.revenue = mo.Variable(
                    [1, 2, 3], display_name="Revenue", unit="$"
                )
        return m

    def test_a_field_gets_a_column_and_a_caption(self):
        g = _grid(_sheet(self._model([("description", "Comment")]), "Plan"), 8)
        # A = name, B = unit, C = the projected field, D = constants,
        # E.. = periods.
        assert g[0][:5] == ["Name", "Unit", "Comment", None, "Jan 2026"]
        assert g[1][:4] == ["Rate", "%", "assumed rate", 0.12]

    def test_a_row_without_the_field_leaves_it_empty(self):
        # A projected column reports what the model holds. It never
        # invents a value, and never borrows the row above's.
        g = _grid(_sheet(self._model([("description", "Comment")]), "Plan"), 8)
        revenue = next(row for row in g if row[0] == "Revenue")
        assert revenue[2] is None

    def test_several_fields_keep_declaration_order(self):
        m = _windowed()
        m.default_excel_view = mo.ExcelView(
            meta={
                "fields": [
                    ("description", "Comment"),
                    ("number_format", "Format"),
                ]
            }
        )
        with m:
            m.plan = mo.MultiVariable("Plan", excel_props={"tab": True})
            with m.plan as p:
                p.rate = mo.Variable(
                    0.12,
                    display_name="Rate",
                    description="assumed rate",
                    excel_props={"number_format": "0.0%"},
                )
                p.series = mo.Variable([1, 2, 3], display_name="Series")
        g = _grid(_sheet(m, "Plan"), 8)
        assert g[0][1:3] == ["Comment", "Format"]
        # The second one reads from excel_props — a field lives
        # wherever the model honestly keeps it.
        assert g[1][1:3] == ["assumed rate", "0.0%"]

    def test_an_unnamed_field_column_still_carries_its_values(self):
        # ``True`` instead of a caption: the column renders headerless,
        # the same courtesy the unit column has always had.
        g = _grid(_sheet(self._model([("description", True)]), "Plan"), 8)
        assert g[0][2] is None
        assert g[1][2] == "assumed rate"


class TestConstantsCaption:
    """The constants column can be HEADED — ``meta={'constants':
    'Value'}`` — in whatever word the report uses for it. Without
    a header the column reads as the sheet having slipped against its
    own timeline.
    """

    def _model(self, **meta):
        m = _windowed()
        if meta:
            m.default_excel_view = mo.ExcelView(meta=meta)
        with m:
            m.plan = mo.MultiVariable("Plan", excel_props={"tab": True})
            with m.plan as p:
                p.rate = mo.Variable(0.12, display_name="Rate")
                p.revenue = mo.Variable([1, 2, 3], display_name="Revenue")
        return m

    def test_the_caption_heads_the_column(self):
        g = _grid(_sheet(self._model(constants="Value"), "Plan"))
        assert g[0][1:3] == ["Value", "Jan 2026"]
        assert g[1][:2] == ["Rate", 0.12]

    def test_undeclared_stays_headerless(self):
        # Same rule as the unit column always had: the flag is the
        # caption, and nothing invents wording the model didn't say.
        g = _grid(_sheet(self._model(), "Plan"))
        assert g[0][1] is None
        assert g[0][2] == "Jan 2026"

    def test_no_caption_without_the_column(self):
        # All-periodic sheet: no constants column is reserved, so the
        # declared caption must NOT claim the label column.
        m = _windowed()
        m.default_excel_view = mo.ExcelView(meta={"constants": "Value"})
        with m:
            m.plan = mo.MultiVariable("Plan", excel_props={"tab": True})
            with m.plan as p:
                p.revenue = mo.Variable([1, 2, 3], display_name="Revenue")
        g = _grid(_sheet(m, "Plan"))
        assert g[0][1] == "Jan 2026"
        assert not any("Value" in str(c) for row in g for c in row if c)


class TestALineOnItsOwnWindow:
    """A line declared with its own ``start=`` / ``grain=`` that differs
    from its sheet's window cannot be written yet. Every series is
    written from the sheet's first period column, so such a line would
    sit under the wrong dates, and formulas reading it would pair the
    wrong periods — the file would compute other numbers than Python.
    The export refuses before writing anything.
    """

    def _refused(self, model):
        path = os.path.join(tempfile.mkdtemp(), "t.xlsx")
        with pytest.raises(ValueError) as err:
            model.to_excel(path)
        assert not os.path.exists(path)
        return str(err.value)

    def _later_start(self):
        m = _windowed()
        with m:
            m.plan = mo.MultiVariable("Plan", excel_props={"tab": True})
            with m.plan as p:
                p.revenue = mo.Variable([1, 2, 3], display_name="Revenue")
                p.bonus = mo.Variable(
                    [10, 20], start="2026-02", grain="month",
                    display_name="Bonus",
                )
        return m

    def test_a_later_start_is_refused(self):
        self._refused(self._later_start())

    def test_a_different_grain_is_refused(self):
        m = _windowed()
        with m:
            m.plan = mo.MultiVariable("Plan", excel_props={"tab": True})
            with m.plan as p:
                p.revenue = mo.Variable([1, 2, 3], display_name="Revenue")
                p.dividend = mo.Variable(
                    [30, 40], start="2026-Q1", grain="quarter",
                    display_name="Dividend",
                )
        msg = self._refused(m)
        assert "'Dividend'" in msg and "quarter" in msg

    def test_the_message_names_the_line_both_windows_and_the_way_out(self):
        msg = self._refused(self._later_start())
        assert "'Bonus'" in msg and "m.plan.bonus" in msg
        assert "2026-02" in msg and "month" in msg
        assert "'Plan'" in msg and "2026-01" in msg
        assert "start=" in msg and "extend=" in msg

    def test_an_own_window_equal_to_the_sheets_exports(self):
        m = _windowed()
        with m:
            m.plan = mo.MultiVariable("Plan", excel_props={"tab": True})
            with m.plan as p:
                p.series = mo.Variable(
                    [1, 2, 3], start="2026-01", grain="month",
                    display_name="Series",
                )
        g = _grid(_sheet(m, "Plan"))
        assert g[0][1:4] == ["Jan 2026", "Feb 2026", "Mar 2026"]
        assert g[1][:4] == ["Series", 1, 2, 3]

    def test_the_same_month_spelled_without_a_leading_zero_exports(self):
        # '2026-1' and '2026-01' are the same month: the line is on
        # its sheet's window, whatever the spelling.
        m = _windowed()
        with m:
            m.plan = mo.MultiVariable("Plan", excel_props={"tab": True})
            with m.plan as p:
                p.series = mo.Variable(
                    [1, 2, 3], start="2026-1", grain="month",
                    display_name="Series",
                )
        assert _grid(_sheet(m, "Plan"))[1][:4] == ["Series", 1, 2, 3]

    def test_an_own_window_that_ends_early_is_refused(self):
        # Same start and grain, fewer periods: Python continues the
        # series by its extend rule (20 held into March), while a
        # formula in the book would read an empty or wrong cell there.
        m = _windowed()
        with m:
            m.plan = mo.MultiVariable("Plan", excel_props={"tab": True})
            with m.plan as p:
                p.revenue = mo.Variable([1, 2, 3], display_name="Revenue")
                p.rate = mo.Variable(
                    [10, 20], start="2026-01", grain="month",
                    extend=mo.hold(), display_name="Rate",
                )
                p.total = p.revenue + p.rate
        assert m.plan.total.value == [11, 22, 23]
        msg = self._refused(m)
        assert "'Rate'" in msg and "2 period(s)" in msg
        assert "3 period(s)" in msg

    def test_extend_and_schedule_on_the_models_window_still_export(self):
        m = _windowed()
        with m:
            m.plan = mo.MultiVariable("Plan", excel_props={"tab": True})
            with m.plan as p:
                p.revenue = mo.Variable([1, 2, 3], display_name="Revenue")
                p.bonus = mo.Variable(
                    [10, 20], extend=mo.zero(), display_name="Bonus"
                )
                p.rate = mo.schedule({"2026-01": 0.1, "2026-03": 0.2})
                p.total = p.revenue + p.bonus
        g = _grid(_sheet(m, "Plan"))
        assert next(r for r in g if r[0] == "Bonus")[1:4] == [10, 20, 0]
        assert next(r for r in g if r[0] == "Rate")[1:4] == [0.1, 0.1, 0.2]
        assert next(r for r in g if r[0] == "Total")[1] == "=B2 + B3"

    def test_a_model_without_a_window_is_unaffected(self):
        m = mo.Model("m", display_name="M")
        with m:
            m.plan = mo.MultiVariable("Plan", excel_props={"tab": True})
            with m.plan as p:
                p.bonus = mo.Variable(
                    [10, 20], start="2026-02", grain="month",
                    display_name="Bonus",
                )
        assert _grid(_sheet(m, "Plan"))[0][:3] == ["Bonus", 10, 20]

    def test_a_line_derived_from_it_is_not_written_either(self):
        m = _windowed()
        with m:
            m.plan = mo.MultiVariable("Plan", excel_props={"tab": True})
            with m.plan as p:
                p.revenue = mo.Variable([1, 2, 3], display_name="Revenue")
                p.bonus = mo.Variable(
                    [10, 20], start="2026-02", grain="month",
                    extend=mo.zero(), display_name="Bonus",
                )
                p.total = p.revenue + p.bonus
        # The sum itself runs on the sheet's window — it is its operand
        # that would be misplaced, and the whole book with it.
        assert m.plan.total.time.start == "2026-01"
        assert m.plan.total.value == [1, 12, 23]
        self._refused(m)

    def test_a_derived_line_reading_an_unattached_one_is_refused(self):
        # The operand lives outside the model, so it has no cells of
        # its own: its values would be inlined into the sum from the
        # first period — under the wrong dates.
        bonus = mo.Variable(
            [10, 20], start="2026-02", grain="month",
            extend=mo.zero(), display_name="Bonus",
        )
        m = _windowed()
        with m:
            m.plan = mo.MultiVariable("Plan", excel_props={"tab": True})
            with m.plan as p:
                p.revenue = mo.Variable([1, 2, 3], display_name="Revenue")
                p.total = p.revenue + bonus
        assert m.plan.total.value == [1, 12, 23]
        msg = self._refused(m)
        assert "'Total'" in msg and "'Bonus'" in msg
