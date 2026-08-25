"""``excel_props={'hidden': True}`` — служебные строки и листы.

Скрытая строка — не пропущенная: она размечается и пишется полностью
(значения, формулы, подписи), поэтому всё, что на неё ссылается,
остаётся живым, а оси чартов в IDE выводятся из её адресов. Прячет её
НАТИВНЫЙ механизм Excel (``row_dimensions[r].hidden`` /
``sheet_state='hidden'``) — книга аудируема, глаз не мусорится.
Флаг наследуется вниз по дереву MV — аналитике карточки достаточно
одного ключа на секции.
"""

import openpyxl
import pytest

import modeleon as mo


@pytest.fixture()
def book(tmp_path):
    def _build(model):
        path = tmp_path / "out.xlsx"
        model.to_excel(str(path))
        return openpyxl.load_workbook(str(path))

    return _build


class TestHiddenRows:
    def test_flagged_row_is_written_but_natively_hidden(self, book):
        m = mo.Model("m")
        with m:
            m.лист = mo.MultiVariable("Лист", excel_props={"tab": True})
            with m.лист as л:
                л.выручка = mo.Variable([10.0, 20.0], display_name="Выручка")
                л.служебная = mo.Variable(
                    [1.0, 2.0],
                    display_name="Служебная",
                    excel_props={"hidden": True},
                )
        wb = book(m)
        ws = wb["Лист"]
        rows = {
            c.value: c.row
            for col in ws.iter_cols(max_col=2)
            for c in col
            if isinstance(c.value, str)
        }
        # Написана — и значения на месте…
        served = rows["Служебная"]
        assert ws.cell(row=served, column=2).value == 1.0
        # …но строка нативно скрыта; обычная — нет.
        assert ws.row_dimensions[served].hidden is True
        assert not ws.row_dimensions[rows["Выручка"]].hidden

    def test_flag_inherits_down_the_mv_tree(self, book):
        m = mo.Model("m")
        with m:
            m.лист = mo.MultiVariable("Лист", excel_props={"tab": True})
            with m.лист as л:
                л.видимая = mo.Variable([5.0], display_name="Видимая")
                л.аналитика = mo.MultiVariable(
                    "Аналитика (карточка)", excel_props={"hidden": True}
                )
                л.аналитика.зеркало = mo.Variable(
                    [7.0], display_name="Зеркало"
                )
        wb = book(m)
        ws = wb["Лист"]
        by_label = {
            c.value: c.row
            for col in ws.iter_cols(max_col=2)
            for c in col
            if isinstance(c.value, str)
        }
        # Секционный заголовок и строка-потомок скрыты ОДНИМ ключом.
        assert ws.row_dimensions[by_label["Аналитика (карточка)"]].hidden is True
        assert ws.row_dimensions[by_label["Зеркало"]].hidden is True
        assert not ws.row_dimensions[by_label["Видимая"]].hidden

    def test_references_to_hidden_rows_stay_live_formulas(self, book):
        # Скрытая строка сохраняет адреса — формула поверх неё остаётся
        # живой ссылкой, не запечённым литералом.
        m = mo.Model("m")
        with m:
            m.лист = mo.MultiVariable("Лист", excel_props={"tab": True})
            with m.лист as л:
                л.база = mo.Variable(
                    [3.0], display_name="База", excel_props={"hidden": True}
                )
                л.двойная = mo.Variable(
                    л.база * 2, display_name="Двойная"
                )
        wb = book(m)
        ws = wb["Лист"]
        rows = {
            c.value: c.row
            for col in ws.iter_cols(max_col=2)
            for c in col
            if isinstance(c.value, str)
        }
        formula = ws.cell(row=rows["Двойная"], column=2).value
        assert isinstance(formula, str) and formula.startswith("=")
        # Живая ссылка на ячейку скрытой строки — не запечённый литерал.
        assert f"B{rows['База']}" in formula


class TestHiddenSheet:
    def test_flagged_tab_hides_the_whole_sheet(self, book):
        m = mo.Model("m")
        with m:
            m.данные = mo.MultiVariable("Данные", excel_props={"tab": True})
            with m.данные as д:
                д.x = mo.Variable([1.0], display_name="X")
            m.служебный = mo.MultiVariable(
                "Служебный",
                excel_props={"tab": True, "hidden": True},
            )
            with m.служебный as с:
                с.y = mo.Variable([2.0], display_name="Y")
        wb = book(m)
        assert wb["Данные"].sheet_state == "visible"
        assert wb["Служебный"].sheet_state == "hidden"
        # Лист написан целиком — аудит возможен.
        assert wb["Служебный"].max_row >= 1


class TestHiddenSurvivesTabWrappers:
    def test_hidden_mv_wraps_into_a_hidden_virtual_sheet(self, book):
        # Верхнеуровневый MV без tab-флага заворачивается движком в
        # ВИРТУАЛЬНЫЙ таб (путь страниц-виджетов) — обёртка не должна
        # ТЕРЯТЬ служебность: скрытый источник → скрытый лист.
        m = mo.Model("m")
        with m:
            m.данные = mo.MultiVariable("Данные", excel_props={"tab": True})
            with m.данные as д:
                д.x = mo.Variable([1.0], display_name="X")
            m.служебный = mo.MultiVariable(
                "Служебный", excel_props={"hidden": True}
            )
            with m.служебный as с:
                с.y = mo.Variable([9.0], display_name="Y")
        wb = book(m)
        служ = next(n for n in wb.sheetnames if "Служебный" in n)
        assert wb[служ].sheet_state == "hidden"
