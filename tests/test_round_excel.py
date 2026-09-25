# SPDX-License-Identifier: Apache-2.0
"""``mo.ROUND`` computes what the spreadsheet's ``ROUND`` computes.

Python's built-in ``round`` sends a half to the nearest EVEN digit
(``round(2.5) == 2``) and works on the binary float, so ``1.005`` — held
as 1.00499999… — rounds down to ``1.0``. The spreadsheet's ``ROUND``
sends a half AWAY FROM ZERO (``ROUND(2.5, 0) = 3``) and works on the
number as the cell shows it, to 15 significant digits, so
``ROUND(1.005, 2) = 1.01``. ``mo.ROUND`` is written into the workbook as
``=ROUND(...)``, so it must compute the latter or Python and the file
disagree. The last test recalculates a written workbook to show they
agree.
"""

from __future__ import annotations

import subprocess
from pathlib import Path

import openpyxl
import pytest

import modeleon as mo


class TestHalvesRoundAwayFromZero:
    @pytest.mark.parametrize("value, digits, expected", [
        (2.5, 0, 3.0),
        (-2.5, 0, -3.0),
        (0.5, 0, 1.0),
        (-0.5, 0, -1.0),
        (1.5, 0, 2.0),
        (1.45, 1, 1.5),
        (0.125, 2, 0.13),
        (-0.125, 2, -0.13),
        (1.005, 2, 1.01),
        (-1.005, 2, -1.01),
        (2.675, 2, 2.68),
        (0.285, 2, 0.29),
        (0.045, 2, 0.05),
    ])
    def test_scalar(self, value, digits, expected):
        assert mo.ROUND(value, digits)._value == expected

    def test_list_is_rounded_element_by_element(self):
        v = mo.Variable([0.5, 1.5, 2.5, -2.5, None, 3.7])
        assert mo.ROUND(v, 0)._value == [1.0, 2.0, 3.0, -3.0, None, 4.0]

    def test_list_with_decimals(self):
        v = mo.Variable([0.125, 1.005, 2.675, -1.005])
        assert mo.ROUND(v, 2)._value == [0.13, 1.01, 2.68, -1.01]

    def test_digits_given_as_a_variable(self):
        digits = mo.Variable(2)
        v = mo.Variable([1.005, 2.675, 3.14159])
        assert mo.ROUND(v, digits)._value == [1.01, 2.68, 3.14]


class TestRoundsTheNumberAsShown:
    def test_a_computed_half_rounds_like_the_half_it_shows(self):
        # 0.35 * 7 is held as 2.4499999999999997; a cell shows it as 2.45
        # and the spreadsheet rounds it to 2.5.
        amount = mo.Variable(0.35) * 7
        assert mo.ROUND(amount, 1)._value == 2.5

    def test_one_step_below_a_half_still_shows_as_the_half(self):
        assert mo.ROUND(2.4999999999999996, 0)._value == 3.0
        assert mo.ROUND(1.0049999999999997, 2)._value == 1.01

    def test_a_half_in_the_sixteenth_digit_rounds_away_from_zero(self):
        # 999999999999998.5 is held exactly; shown to 15 digits the half
        # goes away from zero (…999), not to the even digit (…998).
        assert mo.ROUND(999999999999998.5, 0)._value == 999999999999999.0
        assert mo.ROUND(-999999999999998.5, 0)._value == -999999999999999.0
        assert mo.ROUND(12345678901234.25, 1)._value == 12345678901234.3

    def test_ordinary_values_are_unchanged(self):
        assert mo.ROUND(3.14159, 2)._value == 3.14
        assert mo.ROUND(3.7, 0)._value == 4.0
        assert mo.ROUND(-3.7, 0)._value == -4.0
        assert mo.ROUND(1.23, 2)._value == 1.23
        assert mo.ROUND(0.1 + 0.2, 1)._value == 0.3


class TestNegativeDigits:
    @pytest.mark.parametrize("value, digits, expected", [
        (1234.5, -2, 1200),
        (1250, -2, 1300),
        (-1250, -2, -1300),
        (1249.99, -2, 1200),
        (5, -1, 10),
        (-5, -1, -10),
        (123.456, -1, 120),
    ])
    def test_rounds_to_tens_hundreds(self, value, digits, expected):
        assert mo.ROUND(value, digits)._value == expected


class TestSixteenWholeDigits:
    """A number of 16 or more whole digits has no decimals left to show;
    it is rounded as held, and its whole digits never move."""

    @pytest.mark.parametrize("value, digits, expected", [
        (1234567890123456.0, 0, 1234567890123456.0),
        (1234567890123456.75, 0, 1234567890123457.0),
        (1234567890123456.25, 0, 1234567890123456.0),
        (1234567890123455.5, 0, 1234567890123456.0),
        (-1234567890123456.5, 0, -1234567890123457.0),
        (1234567890123449.5, -2, 1234567890123400.0),
        (1e16 + 2, 0, 1e16 + 2),
    ])
    def test_rounded_as_held(self, value, digits, expected):
        assert mo.ROUND(value, digits)._value == expected

    def test_float_and_whole_number_agree(self):
        assert mo.ROUND(1234567890123456.0, 0)._value == \
            mo.ROUND(1234567890123456, 0)._value


class TestTypes:
    def test_float_stays_float(self):
        result = mo.ROUND(2.5, 0)._value
        assert result == 3.0 and isinstance(result, float)

    def test_whole_number_stays_whole(self):
        result = mo.ROUND(7, 2)._value
        assert result == 7 and isinstance(result, int)

    def test_whole_number_to_hundreds_stays_whole(self):
        result = mo.ROUND(1250, -2)._value
        assert result == 1300 and isinstance(result, int)

    def test_none_passes_through(self):
        assert mo.ROUND(mo.Variable([None, 2.5]), 0)._value == [None, 3.0]

    def test_formula_is_unchanged(self):
        assert mo.ROUND(2.5, 0).formula == "ROUND(2.5, 0)"


def _soffice() -> str | None:
    """Path to the LibreOffice command line, or None when not installed."""
    import shutil

    found = shutil.which("soffice") or shutil.which("libreoffice")
    if found:
        return found
    mac = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    return str(mac) if mac.exists() else None


def _recalculated(model, tmp_path) -> openpyxl.Workbook:
    """Write ``model``, let LibreOffice calculate every formula, and
    return the recalculated workbook's stored values."""
    soffice = _soffice()
    if soffice is None:
        pytest.skip("LibreOffice is not installed")
    src = tmp_path / "book.xlsx"
    model.to_excel(str(src))
    outdir = tmp_path / "calc"
    profile = (tmp_path / "lo-profile").as_uri()  # private: parallel runs don't clash
    subprocess.run(
        [soffice, f"-env:UserInstallation={profile}", "--headless",
         "--convert-to", "xlsx", "--outdir", str(outdir), str(src)],
        check=True, capture_output=True, timeout=120,
    )
    return openpyxl.load_workbook(str(outdir / "book.xlsx"), data_only=True)


def _row(ws, label):
    """The value cells right of ``label`` in column A."""
    for row in ws.iter_rows():
        if row[0].value == label:
            return [c.value for c in row[1:] if c.value is not None]
    raise KeyError(label)


#: digits -> the values rounded at that many digits. Halves on both
#: sides of zero, numbers one float step below a half, and ordinary
#: values that must not move. A typed number reaches the book with 16
#: significant digits, so a step below a half arrives as the half; the
#: Price x Qty rows below build such numbers inside the sheet instead.
_BATTERY = {
    0: [0.5, 1.5, 2.5, -2.5, -0.5, 3.7, 2.4999999999999996, 1234.5],
    1: [1.45, -1.45, 0.05, 2.25, 0.35, 9.95, 3.14159, -0.15],
    2: [0.125, 1.005, -1.005, 2.675, 0.285, 3.14159, 1.0049999999999997, 0.045],
    -2: [1234.5, 1250, -1250, 1249.99, 50, 149.5, -50, 99999],
}


#: (digits, whole parts, fractions): halves held exactly in the 16th
#: digit, and numbers of 16 whole digits.
_LARGE = [
    (0, [999999999999998, 123456789012344, -999999999999998,
         1234567890123456, 1234567890123455, 1000000000000000,
         1234567890123456, 4503599627370495],
        [0.5, 0.5, -0.5, 0.75, 0.5, 0.5, 0.25, 0.5]),
    (1, [12345678901234, 12345678901232], [0.25, 0.25]),
    (2, [1234567890123, 1234567890122], [0.125, 0.625]),
    (-2, [1234567890123449, 1234567890123449], [0.5, 0.75]),
]


def test_workbook_rounds_like_python(tmp_path):
    m = mo.Model("Rounding")
    rounded = {}
    for n, (digits, values) in enumerate(_BATTERY.items()):
        inp = mo.Variable(values, display_name=f"Input {n}")
        setattr(m, f"input_{n}", inp)
        out = mo.ROUND(inp, digits)
        setattr(m, f"rounded_{n}", out)
        out._display_name = f"Rounded {n}"
        rounded[out._display_name] = out
    # Inputs computed inside the workbook, not typed in: 0.35 * 7 is
    # 2.4499999999999997 in both Python and the sheet, shown as 2.45.
    m.price = mo.Variable([0.35, 0.15, 1.15, 2.55, 0.35, 0.95, 2.05, 4.35],
                          display_name="Price")
    m.qty = mo.Variable([7, 3, 3, 3, 3, 7, 7, 3], display_name="Qty")
    m.amount = m.price * m.qty
    m.amount._display_name = "Amount"
    m.places = mo.Variable(1, display_name="Places")
    m.amount_rounded = mo.ROUND(m.amount, m.places)
    m.amount_rounded._display_name = "Amount rounded"
    rounded["Amount rounded"] = m.amount_rounded
    # Very large amounts, built in the sheet as whole part + fraction so
    # the book holds exactly what Python holds. Read back the rounded
    # amount minus its whole part: small, so no digit is lost on the way.
    for digits, whole, fraction in _LARGE:
        n = len(rounded)
        base = mo.Variable(whole, display_name=f"Whole {n}")
        setattr(m, f"whole_{n}", base)
        part = mo.Variable(fraction, display_name=f"Fraction {n}")
        setattr(m, f"fraction_{n}", part)
        amount = base + part
        setattr(m, f"large_{n}", amount)
        moved = mo.ROUND(amount, digits) - base
        setattr(m, f"moved_{n}", moved)
        moved._display_name = f"Moved {n}"
        rounded[moved._display_name] = moved

    ws = _recalculated(m, tmp_path).active
    for label, var in rounded.items():
        assert _row(ws, label) == var._value, label
