"""``bg_nonzero`` — value-driven fill as a LIVE conditional rule.

A Gantt band / active-month mask paints its cells wherever the value
is nonzero; zeros and blanks stay bare. The rule rides the workbook
(not baked fills), so editing the file's values repaints the band.
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


def _windowed():
    m = mo.Model("m", display_name="M")
    m.default_start, m.default_grain, m.default_periods = "2026-01", "month", 3
    return m


class TestBgNonzero:
    def test_unknown_key_no_more(self):
        # The whitelist accepts the new key — construction must not raise.
        mo.Variable([0, 1, 0], excel_props={"bg_nonzero": "1C9499"})

    def test_a_conditional_rule_covers_the_value_cells(self):
        m = _windowed()
        with m:
            m.план = mo.MultiVariable("План", excel_props={"tab": True})
            with m.план as p:
                p.полоса = mo.Variable(
                    [0.0, 1.0, 1.0],
                    display_name="Срок исполнения",
                    excel_props={"bg_nonzero": "1C9499"},
                )
        ws = _sheet(m, "План")
        rules = [
            (str(rng), rule)
            for rng in ws.conditional_formatting
            for rule in rng.rules
        ]
        assert rules, "no conditional formatting emitted"
        sqref, rule = rules[0]
        # All three value cells are covered (row 2: B..D under the header).
        assert "B2" in sqref and "D2" in sqref
        # Numbers only: CellIs «notEqual 0» painted error tokens and
        # any text — the grid's own law is nonzero NUMBERS.
        assert rule.formula == ["AND(ISNUMBER(B2),B2<>0)"]
        assert rule.dxf.fill.bgColor.rgb.endswith("1C9499")

    def test_no_rule_without_the_prop(self):
        m = _windowed()
        with m:
            m.план = mo.MultiVariable("План", excel_props={"tab": True})
            with m.план as p:
                p.выручка = mo.Variable([1, 2, 3], display_name="Выручка")
        ws = _sheet(m, "План")
        assert len(list(ws.conditional_formatting)) == 0
