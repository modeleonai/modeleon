# SPDX-License-Identifier: Apache-2.0
"""Subtotal columns — «янв фев мар (1 кв) … (год)» — and they are ALIVE.

``ExcelView(timeline={'totals': ['quarter', 'year']})`` interleaves a
quarter column after each quarter's last month and a year column after
the year. The parity promise extends to them: edit January in the
downloaded file and the quarter, the year, and every formula standing
on them move. That rules out baked numbers everywhere a formula can
honestly stand:

* a rule line   → ``=SUM(месяцы)`` / ``=AVERAGE`` / last-month ref;
* a formula line → its OWN formula over the operands' bucket cells —
  exactly what a hand-built book writes at the quarter column;
* a ruleless literal → the projection's ``#VALUE!`` teaching token,
  never an invented sum.
"""

import os
import tempfile

import openpyxl
from openpyxl.utils import get_column_letter

import modeleon as mo


def _book(model):
    path = os.path.join(tempfile.mkdtemp(), "t.xlsx")
    model.to_excel(path)
    return openpyxl.load_workbook(path)


def _model(**view_kw):
    m = mo.Model(
        "m",
        display_name="M",
        default_excel_view=mo.ExcelView(
            timeline={"totals": ["quarter", "year"], **view_kw}
        ),
    )
    m.default_start, m.default_grain, m.default_periods = "2026-01", "month", 6
    with m:
        m.план = mo.MultiVariable("План", excel_props={"tab": True})
        with m.план as p:
            p.ставка = mo.Variable(0.1, display_name="Ставка")
            p.выручка = mo.Variable(
                [10, 20, 30, 40, 50, 60],
                display_name="Выручка",
                regrain=mo.up("sum"),
            )
            p.налог = p.выручка * p.ставка
    return m


class TestTheBookForm:
    def test_the_header_is_hierarchical(self):
        # «Год» over «кварталы» over «месяцы» — founder ask: «в 2
        # линии чтобы выглядело лучше». A bucket spans its months and
        # the finer totals inside it, and its OWN total stands beside
        # that span, named, from the bucket's own tier downwards:
        # «Итого Q1» begins where «Q1» begins (founder ask), not a
        # tier lower where it read as a thirteenth month.
        ws = _book(_model())["План"]
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
        ws = _book(_model())["План"]
        assert ws.cell(4, 1).value == "Ставка"
        assert ws.freeze_panes == "C4"

    def test_a_rule_line_sums_its_months_live(self):
        ws = _book(_model())["План"]
        assert ws["F5"].value == "=SUM(C5,D5,E5)"
        assert ws["J5"].value == "=SUM(G5,H5,I5)"

    def test_the_year_never_swallows_the_quarter_cells(self):
        # The year's months are NOT contiguous once quarter columns
        # stand between them — a range SUM would double-count. Explicit
        # refs, months only.
        ws = _book(_model())["План"]
        assert ws["K5"].value == "=SUM(C5,D5,E5,G5,H5,I5)"

    def test_a_formula_line_writes_its_own_formula_at_the_bucket(self):
        # The hand-built book's rule: налог(Q1) = выручка(Q1) × ставка,
        # standing on the QUARTER cell — not a sum of monthly taxes,
        # not a baked number.
        ws = _book(_model())["План"]
        assert ws["F6"].value == "=F5 * B4"
        assert ws["K6"].value == "=K5 * B4"

    def test_native_months_still_reference_their_own_cells(self):
        # The interleave moves ADDRESSES, so native formulas follow.
        ws = _book(_model())["План"]
        assert ws["C6"].value == "=C5 * B4"
        assert ws["G6"].value == "=G5 * B4"

    def test_a_ruleless_literal_teaches_instead_of_inventing(self):
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
            m.план = mo.MultiVariable("План", excel_props={"tab": True})
            with m.план as p:
                p.курс = mo.Variable(
                    [480, 481, 482], display_name="Курс"
                )
        ws = _book(m)["План"]
        # No rule declared: summing an FX rate would be a lie; the
        # projection's token lands instead.
        assert ws["E3"].value == "#VALUE!"

    def test_partial_trailing_quarter_still_totals(self):
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
            m.план = mo.MultiVariable("План", excel_props={"tab": True})
            with m.план as p:
                p.поток = mo.Variable(
                    [1, 2, 3, 4], display_name="Поток",
                    regrain=mo.up("sum"),
                )
        ws = _book(m)["План"]
        # Q1 after March; the lone April closes a partial Q2.
        assert ws["E3"].value == "=SUM(B3,C3,D3)"
        assert ws["G3"].value == "=SUM(F3)"

    def test_undeclared_totals_change_nothing(self):
        m = _model()
        m.default_excel_view = mo.ExcelView()
        ws = _book(m)["План"]
        row1 = [ws.cell(1, c).value for c in range(1, 10)]
        assert "Q1 2026" not in row1
        assert ws["C3"].value == 10  # contiguous months, untouched


class TestOtherRules:
    def _one_line(self, **var_kw):
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
            m.план = mo.MultiVariable("План", excel_props={"tab": True})
            with m.план as p:
                p.линия = mo.Variable(
                    [10, 20, 30], display_name="Линия", **var_kw
                )
        return _book(m)["План"]

    def test_last(self):
        # A balance's quarter is its closing month — a ref, not a sum.
        ws = self._one_line(regrain=mo.up("last"))
        assert ws["E3"].value == "=D3"

    def test_mean(self):
        ws = self._one_line(regrain=mo.up("mean"))
        assert ws["E3"].value == "=AVERAGE(B3,C3,D3)"


class TestTheTotalsWord:
    def test_the_model_speaks_its_own_language(self):
        m = mo.Model(
            "m",
            display_name="M",
            default_excel_view=mo.ExcelView(
                timeline={"totals": ["quarter"], "totals_word": "Итого"}
            ),
        )
        m.default_start, m.default_grain, m.default_periods = (
            "2026-01", "month", 3,
        )
        with m:
            m.план = mo.MultiVariable("План", excel_props={"tab": True})
            with m.план as p:
                p.поток = mo.Variable(
                    [1, 2, 3], display_name="Поток", regrain=mo.up("sum"),
                )
        ws = _book(m)["План"]
        # One tier above the months, so the quarter's own total names
        # itself on that tier and merges down through the month row.
        assert ws.cell(1, 5).value == "Итого Q1"
        assert "E1:E2" in {str(r) for r in ws.merged_cells.ranges}

    def test_tier_spans_read_from_the_left(self):
        # «2026» centered floats far from the row names — and the
        # month row stays centered under it.
        ws = _book(_model())["План"]
        assert ws.cell(1, 3).alignment.horizontal == "left"
        assert ws.cell(2, 3).alignment.horizontal == "left"
        assert ws.cell(3, 3).alignment.horizontal == "center"


class TestNativeColumnGroups:
    def test_the_file_carries_real_outline_levels(self):
        # Excel's own «свернуть 2027»: months sit two buckets deep,
        # a quarter's total one (inside the year), the year's total
        # outside every group — and the summary stands RIGHT of its
        # detail, where our totals are.
        ws = _book(_model())["План"]
        level = lambda c: (
            ws.column_dimensions[c].outline_level
            if c in ws.column_dimensions else 0
        )
        assert [level(c) for c in "CDEF"] == [2, 2, 2, 1]
        assert level("K") == 0  # «Итого 2026» survives every fold
        assert ws.sheet_properties.outlinePr.summaryRight is True


class TestTotalsOverrideLens:
    def test_the_lens_replaces_and_strips_the_declaration(self):
        # The viewing lens: LayoutEngine(totals_override=...) beats the
        # declared timeline.totals — ('quarter',) narrows, () strips,
        # None keeps. The export path never passes it.
        from modeleon.compile.excel.layout import LayoutEngine

        m = _model()
        roots = [m.план]
        eng = LayoutEngine(roots, totals_override=("quarter",))
        eng.compute_addresses()
        plan = eng.totals_by_sheet.get("План")
        assert plan is not None and plan.kinds == ("quarter",)

        eng2 = LayoutEngine(roots, totals_override=())
        eng2.compute_addresses()
        assert eng2.totals_by_sheet.get("План") is None

        eng3 = LayoutEngine(roots)
        eng3.compute_addresses()
        assert eng3.totals_by_sheet.get("План").kinds == ("quarter", "year")


class TestRangesNeverSwallowATotal:
    """A row's cells stop being contiguous the moment totals declare a
    column between them — and ``SUM(C3:I3)`` then counts the quarter
    twice, in a file that looks perfectly valid.

    ``totals.rule_formula`` learned this for its own buckets long ago;
    the lesson belongs wherever a range is built.
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
                p.поток = mo.Variable([10.0] * n, display_name="Поток",
                                      regrain=mo.up("sum"))
                p.всего = mo.Variable((fn or mo.SUM)(p.поток),
                                      display_name="Всего")
        path = os.path.join(tempfile.mkdtemp(), "t.xlsx")
        m.to_excel(path)
        ws = openpyxl.load_workbook(path)["P"]
        row = next(r for r in range(1, ws.max_row + 1)
                   if ws.cell(r, 1).value == "Всего")
        return ws.cell(row, 2).value

    def test_an_interleaved_row_lists_its_cells(self):
        # Months C,D,E then the Q1 total at F, then G,H,I. The range
        # form would add F twice over: 90 for a model that says 60.
        formula = self._book(["quarter"])
        assert ":" not in formula
        assert formula.replace(" ", "") == "=SUM(C3,D3,E3,G3,H3,I3)"

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
                p.поток = mo.Variable([-100.0, 30.0, 30.0, 30.0, 30.0, 30.0],
                                      display_name="Поток",
                                      regrain=mo.up("sum"))
                p.ставка = mo.Variable(mo.IRR(p.поток), display_name="Ставка")
        path = os.path.join(tempfile.mkdtemp(), "t.xlsx")
        m.to_excel(path)
        ws = openpyxl.load_workbook(path)["P"]
        row = next(r for r in range(1, ws.max_row + 1)
                   if ws.cell(r, 1).value == "Ставка")
        assert not str(ws.cell(row, 2).value).startswith("=")


class TestNativeOutlineGroups:
    """The grid's collapsible outline, in the file — both axes."""

    def test_rows_group_by_section_and_the_header_is_the_summary(self):
        m = mo.Model("m", display_name="M")
        m.default_start, m.default_grain, m.default_periods = (
            "2026-01", "month", 3,
        )
        with m:
            m.p = mo.MultiVariable("P", excel_props={"tab": True})
            with m.p as p:
                p.доходы = mo.MultiVariable("I. Доходы")
                with p.доходы as d:
                    d.выручка = mo.Variable([1, 2, 3], display_name="Выручка")
                    d.детали = mo.MultiVariable("Детали")
                    with d.детали as dd:
                        dd.курс = mo.Variable([480] * 3, display_name="Курс")
        ws = _book(m)["P"]
        rows = {ws.cell(r, 1).value: r for r in range(1, ws.max_row + 1)}
        # Section content is one level deep; nested content two.
        assert ws.row_dimensions[rows["Выручка"]].outline_level == 1
        assert ws.row_dimensions[rows["    Курс"] if "    Курс" in rows
                                 else rows[next(k for k in rows if k and "Курс" in k)]
                                 ].outline_level == 2
        # The header row is the SUMMARY, outside its own group…
        assert ws.row_dimensions[rows["I. Доходы"]].outline_level in (0, None)
        # …and Excel puts the ± on it because the summary sits above.
        assert ws.sheet_properties.outlinePr.summaryBelow is False

    def test_column_brackets_cover_the_whole_group(self):
        # Compare mode: a month is a GROUP of columns, and the outline
        # bracket must cover all of them — one column in four read as
        # noise (live report). Hidden operands sit one level deeper so
        # expanding a quarter does not unhide them.
        m = mo.Model(
            "m", display_name="M", default_grain="month",
            default_start="2026-01", default_periods=3,
            tracks=mo.Tracks(
                "план", "факт",
                blend=mo.blend(given="факт", follow="план",
                               until="2026-01", name="live")),
            default_excel_view=mo.ExcelView(
                tracks="compare",
                timeline={"totals": ["quarter"]}),
        )
        with m:
            m.p = mo.MultiVariable("P", excel_props={"tab": True})
            with m.p as p:
                p.поток = mo.Variable(план=[10.0] * 3,
                                      факт=[12.0, None, None],
                                      display_name="Поток",
                                      regrain=mo.up("sum"))
        ws = _book(m)["P"]
        # Group = факт | план | live | откл (stride 4); every month
        # column of January's group carries the month level.
        lv = {c: ws.column_dimensions[get_column_letter(c)].outline_level
              for c in range(2, 10)}
        assert lv[2] == lv[3] == lv[4] == lv[5] == 1
