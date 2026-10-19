# SPDX-License-Identifier: Apache-2.0
"""A number format's bare literal is quoted on the way to the book.

``0.00x`` is how a multiple is typed. Excel's own dialog stores it as
``0.00"x"``; written as typed it is a format code with a letter that
means nothing, which a stricter reader rejects or rewrites. The
resolved format quotes a run of letters that is not a format code and
leaves every real format exactly as declared.
"""

import pytest
from openpyxl import load_workbook

import modeleon as mo
from modeleon.compile.excel.writer import quote_bare_literals, resolve_number_format

# Every format the suite and the house look declare, and the ones a
# model is likely to: none of these may change.
UNCHANGED = [
    "General", "general", "0", "0.0", "0.00", "0.0000", "#,##0", "#,##0.0", "#,##0.00",
    "#,##0,", "0%", "0.0%", "0.00%", "#,##0;;", '"X";;', ";;;",
    "#,##0;[Red](#,##0)", "#,##0,;[Red](#,##0,)", '#,##0.0;(#,##0.0);"–"', "[Red]0.0",
    "[Blue][>100]0;[Red]0", '#,##0 "USD m"', '#,##0.00 "USD"', "$#,##0", "[$USD] #,##0",
    "[$€-407] #,##0.00", '0 "days"', "# ?/?", "0\\x", '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
    "@", "0.00E+00", "0.00e-00", "##0.0E+0",
    "dd mmm yyyy", "mmm yyyy", "yyyy-mm-dd", "yyyy-mm-dd hh:mm:ss", "hh:mm", "yy", "yyyy",
    "d-mmm-yy", "dd/mm/yyyy", "mm/dd/yy", "ddd", "dddd", "mmmmm", "DD MMM YYYY",
    "h:mm AM/PM", "h:mm:ss A/P", "[h]:mm:ss", "[mm]:ss", "b2yyyy", "0.00\"x\"",
    "yyyymmdd", "hhmmss", "ddmmyyyy", "mmmyy",
]

# Adjacent date and time tokens: a run drawn from ``d m y h s g`` is
# read token by token (``yyyymmdd`` is a year, a month and a day), never
# as one word to quote.
ADJACENT = [
    "yyyymmdd", "hhmm", "mmss", "ddmmyyyy", "yymmdd", "hhmmss", "mmmyy", "mmmdd",
    "yyyymmdd hhmmss", "dmy", "[$-409]yyyymmdd", "yyyymmdd;@",
]


@pytest.mark.parametrize("fmt", UNCHANGED)
def test_a_real_format_round_trips_unchanged(fmt):
    assert quote_bare_literals(fmt) == fmt


@pytest.mark.parametrize("fmt", ADJACENT)
def test_adjacent_date_tokens_stay_code(fmt):
    assert quote_bare_literals(fmt) == fmt


def test_the_book_keeps_a_compact_date_format(tmp_path):
    m = mo.Model("m", display_name="M")
    m.closing = mo.Variable(
        mo.DATE(2026, 3, 31), display_name="Closing", excel_props={"number_format": "yyyymmdd"}
    )
    assert resolve_number_format(m.closing) == "yyyymmdd"
    assert _first_value_cell(m, tmp_path).number_format == "yyyymmdd"


@pytest.mark.parametrize(
    "typed, quoted",
    [
        ("0.00x", '0.00"x"'),
        ("0.0X", '0.0"X"'),
        ("0.0x;(0.0x);-", '0.0"x";(0.0"x");-'),
        ("#,##0 bn", '#,##0 "bn"'),
        ("0 days", '0 "days"'),
        ("0.0 Stück", '0.0 "Stück"'),  # letters beyond ASCII
        ('0.0x "p.a."', '0.0"x" "p.a."'),
        ("[Red]0.0x", '[Red]0.0"x"'),
    ],
)
def test_a_bare_literal_is_quoted(typed, quoted):
    assert quote_bare_literals(typed) == quoted
    assert quote_bare_literals(quoted) == quoted


def _first_value_cell(m, tmp_path):
    path = tmp_path / "fmt.xlsx"
    m.to_excel(str(path))
    return load_workbook(str(path))["M"]["B1"]


def test_the_book_carries_the_quoted_format(tmp_path):
    m = mo.Model("m", display_name="M")
    m.multiple = mo.Variable(7.0, display_name="Multiple", excel_props={"number_format": "0.00x"})
    assert resolve_number_format(m.multiple) == '0.00"x"'
    assert _first_value_cell(m, tmp_path).number_format == '0.00"x"'


def test_the_view_format_by_unit_is_quoted_too(tmp_path):
    m = mo.Model("m", display_name="M")
    m.default_excel_view = mo.ExcelView(formats={"x": "0.00x", "number": "#,##0.0"})
    m.multiple = mo.Variable(7.0, display_name="Multiple", unit="x")
    assert _first_value_cell(m, tmp_path).number_format == '0.00"x"'


def test_a_section_format_is_quoted_for_its_lines(tmp_path):
    m = mo.Model("m", display_name="M")
    m.returns = mo.MultiVariable("Returns", excel_props={"number_format": "0.0x"})
    m.returns.moic = mo.Variable(2.5, display_name="MOIC")
    assert resolve_number_format(m.returns.moic) == '0.0"x"'


def test_no_format_stays_none():
    m = mo.Model("m", display_name="M")
    m.x = mo.Variable(1.0, display_name="X")
    assert resolve_number_format(m.x) is None
