# SPDX-License-Identifier: Apache-2.0
"""Long formulas: a sum over hundreds of rows, and Excel's length limit.

``sum(rows)`` builds an unnamed step per row — the previous step plus that
row — and the workbook spells the sum through them: ``=B2 + 0 + B3 + …``.
Rendered step inside step, a few hundred rows outgrew Python's stack and
the cell fell back to its number without a word. The run renders in a loop
now, live up to what Excel keeps in one formula (8192 characters); past
that the cell holds its number and the log says why.
"""

from __future__ import annotations

import functools
import logging
import operator
import re
import shutil
import subprocess
from pathlib import Path

import openpyxl
import pytest

import modeleon as mo


def _model(n, *, rows_sheet="S", total_sheet=None, formulas=None):
    """``n`` input rows on ``rows_sheet``; ``formulas(rows, k)`` names the
    lines to write — on ``total_sheet`` when given, else beside the rows."""
    m = mo.Model("m", default_grain="month", default_start="2026-01", default_periods=2)
    with m:
        m.s = mo.MultiVariable(rows_sheet, excel_props={"tab": True})
        with m.s as s:
            rows = []
            for i in range(n):
                setattr(s, f"r{i}", mo.Variable([1.0 + i % 7, 0.5 * (i % 5)],
                                                display_name=f"R{i}"))
                rows.append(getattr(s, f"r{i}"))
            s.k = mo.Variable(1.001, display_name="K")
        lines = (formulas or (lambda rows, k: {"Total": sum(rows)}))(rows, s.k)
        if total_sheet is None:
            with m.s as s:
                for i, (label, expr) in enumerate(lines.items()):
                    setattr(s, f"line{i}", mo.Variable(expr, display_name=label))
        else:
            m.t = mo.MultiVariable(total_sheet, excel_props={"tab": True})
            with m.t as t:
                for i, (label, expr) in enumerate(lines.items()):
                    setattr(t, f"line{i}", mo.Variable(expr, display_name=label))
    return m


def _book(model, tmp_path):
    out = tmp_path / "book.xlsx"
    model.to_excel(str(out))
    return openpyxl.load_workbook(str(out))


def _row(ws, label):
    """The value cells right of ``label`` in column A."""
    return [value for _col, value in _cells(ws, label)]


def _cells(ws, label):
    """``(column letter, value)`` of the filled cells right of ``label``."""
    for row in ws.iter_rows():
        if row[0].value == label:
            return [(c.column_letter, c.value) for c in row[1:] if c.value is not None]
    raise KeyError(label)


def _label_rows(ws):
    return {row[0].value: row[0].row for row in ws.iter_rows() if row[0].value}


class TestALongSumIsALiveFormula:
    def test_four_hundred_rows_sum_in_one_formula(self, tmp_path):
        ws = _book(_model(400), tmp_path)["S"]
        at = _label_rows(ws)
        cells = _cells(ws, "Total")
        assert len(cells) == 2
        for col, formula in cells:
            assert formula.startswith("=") and "(" not in formula
            refs = re.findall(r"[A-Z]+\d+", formula)
            assert sorted(refs) == sorted(f"{col}{at[f'R{i}']}" for i in range(400))

    def test_signs_and_brackets_keep_their_places(self, tmp_path):
        model = _model(300, formulas=lambda rows, k: {
            "Net": rows[0] - (rows[1] - rows[2]) - sum(rows[3:300]),
        })
        ws = _book(model, tmp_path)["S"]
        at = _label_rows(ws)
        col, formula = _cells(ws, "Net")[0]
        head, rest = formula.split(" - (", 1)
        assert head == f"={col}{at['R0']}"
        inner, tail = rest.split(") - (", 1)
        assert inner == f"{col}{at['R1']} - {col}{at['R2']}"
        assert tail.endswith(")") and "(" not in tail[:-1]
        refs = re.findall(r"[A-Z]+\d+", tail)
        assert sorted(refs) == sorted(f"{col}{at[f'R{i}']}" for i in range(3, 300))

    def test_past_excels_length_the_cell_holds_its_number(self, tmp_path, caplog):
        # Every reference repeats the quoted sheet name: ~40 characters a
        # term, so 300 rows run past Excel's 8192.
        model = _model(300, rows_sheet="Inputs of a long portfolio", total_sheet="T")
        with caplog.at_level(logging.WARNING, logger="modeleon"):
            cells = _row(_book(model, tmp_path)["T"], "Total")
        assert cells == pytest.approx(model.t.line0.value)
        assert "past Excel's limit of 8,192" in caplog.text

    def test_a_long_sum_still_finds_an_input_on_a_window_of_its_own(self, tmp_path):
        # The book does not write ``early`` and it starts a year before the
        # sheet: the check that refuses it walks the sum's 1500 unnamed
        # steps to reach it, on a stack of its own.
        early = mo.Variable([1.0, 2.0], start="2025-03", grain="month", display_name="Early")
        loose = [mo.Variable([1.0, 2.0]) for _ in range(1500)]
        m = mo.Model("m", default_grain="month", default_start="2026-01", default_periods=2)
        with m:
            m.s = mo.MultiVariable("S", excel_props={"tab": True})
            with m.s as s:
                s.total = mo.Variable(sum([early] + loose), display_name="Total")
        with pytest.raises(ValueError, match="window of its own"):
            m.to_excel(str(tmp_path / "book.xlsx"))


def _soffice() -> str | None:
    """Path to the LibreOffice command line, or None when not installed."""
    found = shutil.which("soffice") or shutil.which("libreoffice")
    if found:
        return found
    mac = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    return str(mac) if mac.exists() else None


@pytest.mark.slow
def test_long_runs_compute_what_python_computed(tmp_path):
    """Sums, mixed signs, brackets and products of hundreds of terms,
    recalculated in LibreOffice: every cell a live formula, every number
    the engine's."""
    soffice = _soffice()
    if soffice is None:
        pytest.skip("LibreOffice is not installed")
    skip_one = None

    def lines(rows, k):
        nonlocal skip_one
        skip_one = rows[5]
        return {
            "Total": sum(rows),
            "Alt": functools.reduce(lambda a, b: a - b if b is skip_one else a + b, rows),
            "Net": rows[0] - (rows[1] - rows[2]) - sum(rows[3:40]),
            "Prod": functools.reduce(operator.mul, [k] * 250) * rows[0],
            "Mix": (rows[0] + rows[1]) * k * k / k - sum(rows[2:200]) * 2,
        }

    model = _model(300, rows_sheet="Inputs", total_sheet="Calc", formulas=lines)
    src = tmp_path / "book.xlsx"
    model.to_excel(str(src))
    written = openpyxl.load_workbook(str(src))["Calc"]
    outdir = tmp_path / "calc"
    profile = (tmp_path / "lo-profile").as_uri()
    subprocess.run(
        [soffice, f"-env:UserInstallation={profile}", "--headless",
         "--convert-to", "xlsx", "--outdir", str(outdir), str(src)],
        check=True, capture_output=True, timeout=120,
    )
    calc = openpyxl.load_workbook(str(outdir / "book.xlsx"), data_only=True)["Calc"]
    for i, label in enumerate(("Total", "Alt", "Net", "Prod", "Mix")):
        assert all(str(f).startswith("=") for f in _row(written, label)), label
        expected = getattr(model.t, f"line{i}").value
        assert _row(calc, label) == pytest.approx(expected), label
