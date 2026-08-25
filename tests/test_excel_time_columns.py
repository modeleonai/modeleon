"""A period header names period columns — and nothing else.

Two rules, one idea: the time axis is a property of the LINE, not of
the model. A model declaring a window doesn't make every sheet a
time series, and it doesn't make a constant a January figure.

  1. No period header on a sheet that carries no per-period line.
     «Должности» is a register of salary bands, every value a
     constant; it printed «Jan 2026» over column B and its 500 read
     as January's.
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
            m.должности = mo.MultiVariable("Должности", excel_props={"tab": True})
            with m.должности as d:
                d.вилка = mo.Variable(500, display_name="Вилка")
                d.разряд = mo.Variable(3, display_name="Разряд")
            # A second sheet DOES carry periods — the window is real,
            # so this pins content and not "the model has no window".
            m.план = mo.MultiVariable("План", excel_props={"tab": True})
            with m.план as p:
                p.выручка = mo.Variable([1, 2, 3], display_name="Выручка")
        g = _grid(_sheet(m, "Должности"))
        assert g[0][:2] == ["Вилка", 500]
        assert not any("2026" in str(c) for row in g for c in row if c)

    def test_the_sheet_that_has_periods_keeps_its_header(self):
        m = _windowed()
        with m:
            m.план = mo.MultiVariable("План", excel_props={"tab": True})
            with m.план as p:
                p.выручка = mo.Variable([1, 2, 3], display_name="Выручка")
        assert _grid(_sheet(m, "План"))[0][1:4] == [
            "Jan 2026",
            "Feb 2026",
            "Mar 2026",
        ]

    def test_two_sheets_are_judged_separately(self):
        # Content belongs to the sheet that prints it: one tab having
        # periods must not put a header over another that hasn't.
        m = _windowed()
        with m:
            m.курсы = mo.MultiVariable("Курсы", excel_props={"tab": True})
            with m.курсы as c:
                c.usd = mo.Variable(450, display_name="USD")
            m.план = mo.MultiVariable("План", excel_props={"tab": True})
            with m.план as p:
                p.ряд = mo.Variable([1, 2, 3], display_name="Ряд")
        assert not any(
            "2026" in str(c)
            for row in _grid(_sheet(m, "Курсы"))
            for c in row
            if c
        )
        assert _grid(_sheet(m, "План"))[0][1] == "Jan 2026"


class TestConstantsKeepTheirOwnColumn:
    def _mixed(self):
        m = _windowed()
        with m:
            m.план = mo.MultiVariable("План", excel_props={"tab": True})
            with m.план as p:
                p.ставка = mo.Variable(0.12, display_name="Ставка")
                p.выручка = mo.Variable([100, 110, 120], display_name="Выручка")
                p.ндс = p.выручка * p.ставка
        return m

    def test_a_constant_is_not_a_january_figure(self):
        ws = _sheet(self._mixed(), "План")
        g = _grid(ws)
        # Periods start at C; B is the constants column.
        assert g[0][1] is None
        assert g[0][2:5] == ["Jan 2026", "Feb 2026", "Mar 2026"]
        assert g[1][:3] == ["Ставка", 0.12, None]
        assert g[2][1:5] == ["Выручка", 100, 110, 120][1:] or g[2][2:5] == [
            100,
            110,
            120,
        ]

    def test_formulas_follow_the_moved_cells(self):
        # The point of moving ADDRESSES rather than painting: a
        # reference to the constant still lands on it.
        g = _grid(_sheet(self._mixed(), "План"))
        ндс = next(row for row in g if row[0] == "Ндс")
        assert ндс[2] == "=C3 * B2"

    def test_the_pane_freezes_left_of_the_periods(self):
        # The constants column must stay in view while the periods
        # scroll — it is a lead column in everything but name.
        assert _sheet(self._mixed(), "План").freeze_panes == "C2"

    def test_no_empty_column_when_the_sheet_is_all_periods(self):
        m = _windowed()
        with m:
            m.план = mo.MultiVariable("План", excel_props={"tab": True})
            with m.план as p:
                p.выручка = mo.Variable([1, 2, 3], display_name="Выручка")
        g = _grid(_sheet(m, "План"))
        assert g[0][1] == "Jan 2026"
        assert g[1][:4] == ["Выручка", 1, 2, 3]

    def test_lead_metadata_columns_still_come_first(self):
        m = _windowed()
        m.default_excel_view = mo.ExcelView(
            meta={"article": True, "unit": "Ед. изм.", "label": "Наименование"}
        )
        with m:
            m.план = mo.MultiVariable("План", excel_props={"tab": True})
            with m.план as p:
                p.ставка = mo.Variable(0.12, display_name="Ставка")
                p.выручка = mo.Variable([1, 2, 3], display_name="Выручка")
        g = _grid(_sheet(m, "План"), cols=8)
        # A = №, B = name, C = unit, D = constants, E.. = periods.
        assert g[0][1] == "Наименование" and g[0][2] == "Ед. изм."
        assert g[0][4] == "Jan 2026"
        ставка = next(row for row in g if row[1] == "Ставка")
        assert ставка[3] == pytest.approx(0.12)


class TestProjectedFieldColumns:
    """Any field can ride its own lead column, by analogy with the
    unit: ``meta={'fields': [('description', 'Комментарий')]}``.

    A list of PAIRS rather than free-standing keys because a view is a
    real MV tree — a field named ``description`` would collide with
    ``MultiVariable.description`` the moment it became an attribute of
    the meta node.
    """

    def _model(self, fields):
        m = _windowed()
        m.default_excel_view = mo.ExcelView(
            meta={"unit": "Ед. изм.", "label": "Наименование", "fields": fields}
        )
        with m:
            m.план = mo.MultiVariable("План", excel_props={"tab": True})
            with m.план as p:
                p.ставка = mo.Variable(
                    0.12,
                    display_name="Ставка",
                    unit="%",
                    description="по НК РК",
                )
                p.выручка = mo.Variable(
                    [1, 2, 3], display_name="Выручка", unit="₸"
                )
        return m

    def test_a_field_gets_a_column_and_a_caption(self):
        g = _grid(_sheet(self._model([("description", "Комментарий")]), "План"), 8)
        # A = name, B = unit, C = the projected field, D = constants,
        # E.. = periods.
        assert g[0][:5] == ["Наименование", "Ед. изм.", "Комментарий", None, "Jan 2026"]
        assert g[1][:4] == ["Ставка", "%", "по НК РК", 0.12]

    def test_a_row_without_the_field_leaves_it_empty(self):
        # A projected column reports what the model holds. It never
        # invents a value, and never borrows the row above's.
        g = _grid(_sheet(self._model([("description", "Комментарий")]), "План"), 8)
        выручка = next(row for row in g if row[0] == "Выручка")
        assert выручка[2] is None

    def test_several_fields_keep_declaration_order(self):
        m = _windowed()
        m.default_excel_view = mo.ExcelView(
            meta={
                "fields": [
                    ("description", "Комментарий"),
                    ("number_format", "Формат"),
                ]
            }
        )
        with m:
            m.план = mo.MultiVariable("План", excel_props={"tab": True})
            with m.план as p:
                p.ставка = mo.Variable(
                    0.12,
                    display_name="Ставка",
                    description="по НК РК",
                    excel_props={"number_format": "0.0%"},
                )
                p.ряд = mo.Variable([1, 2, 3], display_name="Ряд")
        g = _grid(_sheet(m, "План"), 8)
        assert g[0][1:3] == ["Комментарий", "Формат"]
        # The second one reads from excel_props — a field lives
        # wherever the model honestly keeps it.
        assert g[1][1:3] == ["по НК РК", "0.0%"]

    def test_an_unnamed_field_column_still_carries_its_values(self):
        # ``True`` instead of a caption: the column renders headerless,
        # the same courtesy the unit column has always had.
        g = _grid(_sheet(self._model([("description", True)]), "План"), 8)
        assert g[0][2] is None
        assert g[1][2] == "по НК РК"


class TestConstantsCaption:
    """The constants column can be HEADED — ``meta={'constants':
    'Значение'}`` — the book form's own word for it. Reported live:
    without a header the column read as the sheet having slipped
    against its own timeline.
    """

    def _model(self, **meta):
        m = _windowed()
        if meta:
            m.default_excel_view = mo.ExcelView(meta=meta)
        with m:
            m.план = mo.MultiVariable("План", excel_props={"tab": True})
            with m.план as p:
                p.ставка = mo.Variable(0.12, display_name="Ставка")
                p.выручка = mo.Variable([1, 2, 3], display_name="Выручка")
        return m

    def test_the_caption_heads_the_column(self):
        g = _grid(_sheet(self._model(constants="Значение"), "План"))
        assert g[0][1:3] == ["Значение", "Jan 2026"]
        assert g[1][:2] == ["Ставка", 0.12]

    def test_undeclared_stays_headerless(self):
        # Same rule as the unit column always had: the flag is the
        # caption, and nothing invents wording the model didn't say.
        g = _grid(_sheet(self._model(), "План"))
        assert g[0][1] is None
        assert g[0][2] == "Jan 2026"

    def test_no_caption_without_the_column(self):
        # All-periodic sheet: no constants column is reserved, so the
        # declared caption must NOT claim the label column.
        m = _windowed()
        m.default_excel_view = mo.ExcelView(meta={"constants": "Значение"})
        with m:
            m.план = mo.MultiVariable("План", excel_props={"tab": True})
            with m.план as p:
                p.выручка = mo.Variable([1, 2, 3], display_name="Выручка")
        g = _grid(_sheet(m, "План"))
        assert g[0][1] == "Jan 2026"
        assert not any("Значение" in str(c) for row in g for c in row if c)
