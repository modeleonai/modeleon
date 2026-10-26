# SPDX-License-Identifier: Apache-2.0
"""``LEFT``, ``RIGHT``, ``MID``, ``TRIM``, ``FIND``, ``SEARCH``,
``SUBSTITUTE``, ``VALUE``, ``TEXT``, ``REPT`` and ``TEXTJOIN`` compute
what the spreadsheet computes, with the formula as the product.

Each of the first ten is one call per period through the shared helper:
an operand reads as the cell does (a blank is empty text, a number its
General text, a date its serial number, TRUE the word), a count is
truncated to a whole number, numeric text is a number and other text
``#VALUE!`` where a number is asked for, the first spreadsheet error
among a period's operands is that period's result while text that only
starts with ``#`` is text, scalars broadcast, tracks lift, and no result
carries a unit. ``TEXTJOIN`` joins whole rows into one text.
"""

from __future__ import annotations

import logging
from datetime import date, datetime, timedelta

import pytest

import modeleon as mo
from modeleon.core.expr import FuncCall, Literal, VarRef
from modeleon.functions.registry import (
    admits_in_loop,
    excel_name,
    file_prefix,
    is_elementwise,
    is_range_taking,
    range_positions,
    spec,
)
from modeleon.functions.text import (
    FIND,
    LEFT,
    MID,
    REPT,
    RIGHT,
    SEARCH,
    SUBSTITUTE,
    TEXT,
    TEXTJOIN,
    TRIM,
    VALUE,
)

ELEMENTWISE = {
    "LEFT": LEFT, "RIGHT": RIGHT, "MID": MID, "TRIM": TRIM, "FIND": FIND,
    "SEARCH": SEARCH, "SUBSTITUTE": SUBSTITUTE, "VALUE": VALUE, "TEXT": TEXT,
    "REPT": REPT,
}


# ══════════════════════════════════════════════════════════════════════
# Slicing text
# ══════════════════════════════════════════════════════════════════════


class TestLeft:
    @pytest.mark.parametrize("text, n, expected", [
        ("Alpha Fund", 5, "Alpha"),
        ("abc", 1, "a"),
        ("abc", 0, ""),
        ("abc", 10, "abc"),
        ("abcdef", 2.9, "ab"),
        ("abcdef", "2", "ab"),
        ("abc", -1, "#VALUE!"),
        ("abc", "x", "#VALUE!"),
        ("", 3, ""),
    ])
    def test_the_first_characters(self, text, n, expected):
        assert LEFT(text, n).value == expected

    def test_one_character_by_default(self):
        assert LEFT("abc").value == "a"
        assert LEFT("abc").formula == "LEFT('abc')"

    def test_a_blank_is_empty_text_and_a_blank_count_is_zero(self):
        assert LEFT(None, 2).value == ""
        assert LEFT("abc", mo.Variable(None)).value == ""

    def test_a_number_a_date_and_a_boolean_read_as_the_cell_shows_them(self):
        assert LEFT(12345.5, 3).value == "123"
        assert LEFT(1 / 3, 20).value == "0.333333333333333"
        assert LEFT(2022.0, 4).value == "2022"
        assert LEFT(999999999999999, 20).value == "999999999999999"
        assert LEFT(0.0001, 20).value == "0.0001"
        assert LEFT(date(2026, 1, 1), 4).value == "4602"
        assert LEFT(mo.Variable(datetime(2026, 1, 1, 12)), 7).value == "46023.5"
        assert LEFT(timedelta(hours=36), 3).value == "1.5"
        assert LEFT(True, 2).value == "TR"

    @pytest.mark.parametrize("number, text", [
        (0.00001, "0.00001"),
        (-0.00001, "-0.00001"),
        (3e-7, "0.0000003"),
        (5.6789012345e-8, "0.000000056789012345"),
        (9.999999999999999e-19, "0.000000000000000001"),
        (1e15, "1000000000000000"),
        (1234567890123456, "1234567890123460"),
        (1.2345678901234567e19, "12345678901234600000"),
        (-1.2345678901234567e19, "-12345678901234600000"),
        (999999.9999999999, "1000000"),
        (123.45678901234568, "123.456789012346"),
        (0.1 + 0.2, "0.3"),
        # E notation from 1E20, and for a fraction whose spelling would
        # pass 20 characters; the exponent has two digits at least.
        (1e20, "1E+20"),
        (1.2345678901234568e20, "1.23456789012346E+20"),
        (1.2345678901234568e-5, "1.23456789012346E-05"),
        (5.67890123456e-8, "5.67890123456E-08"),
        (1e-19, "1E-19"),
        (2e-50, "2E-50"),
    ])
    def test_a_number_reads_as_its_general_text(self, number, text):
        # Excel's General text: 15 significant digits, written out in full
        # between those bounds. LibreOffice writes a 16th digit and three-
        # digit exponents (pinned in test_excel_text.py).
        assert LEFT(number, 30).value == text

    def test_a_number_no_cell_holds_is_an_error(self):
        assert LEFT(float("inf"), 2).value == "#NUM!"
        assert LEFT("abc", float("nan")).value == "#NUM!"
        assert REPT("a", float("inf")).value == "#NUM!"

    def test_a_series_and_a_count_that_varies_by_period(self):
        names = mo.Variable(["Alpha Fund", "beta co", None])
        assert LEFT(names, 2).value == ["Al", "be", ""]
        assert LEFT(names, [1, 2, 3]).value == ["A", "be", ""]
        assert LEFT(names, mo.Variable(3)).value == ["Alp", "bet", ""]


class TestRight:
    @pytest.mark.parametrize("text, n, expected", [
        ("AB-123", 3, "123"),
        ("abc", 1, "c"),
        ("abc", 0, ""),
        ("abc", 10, "abc"),
        ("abcdef", 2.9, "ef"),
        ("abc", -1, "#VALUE!"),
        (1234.5, 2, ".5"),
    ])
    def test_the_last_characters(self, text, n, expected):
        assert RIGHT(text, n).value == expected

    def test_one_character_by_default_and_a_series(self):
        assert RIGHT("abc").value == "c"
        assert RIGHT(mo.Variable(["AB-123", "CD-45", None]), 3).value == ["123", "-45", ""]


class TestMid:
    @pytest.mark.parametrize("text, start, n, expected", [
        ("AB-123", 4, 2, "12"),
        ("abcdef", 1, 3, "abc"),
        ("abcdef", 2, 100, "bcdef"),
        ("abcdef", 7, 2, ""),
        ("abcdef", 2.7, 2.2, "bc"),
        ("abcdef", 0, 1, "#VALUE!"),
        ("abcdef", 1, -1, "#VALUE!"),
        ("abcdef", "x", 1, "#VALUE!"),
    ])
    def test_from_a_one_based_position(self, text, start, n, expected):
        assert MID(text, start, n).value == expected

    def test_a_series(self):
        codes = mo.Variable(["AB-123", "CD-45", "EF-6"])
        assert MID(codes, 4, 2).value == ["12", "45", "6"]
        assert MID(codes, [1, 4, 1], 2).value == ["AB", "45", "EF"]


class TestTrim:
    def test_collapses_runs_of_spaces(self):
        assert TRIM("  a   b  ").value == "a b"
        assert TRIM("a").value == "a"
        assert TRIM("   ").value == ""

    def test_spaces_only(self):
        assert TRIM("\ta  b c ").value == "\ta b c"

    def test_a_blank_a_number_and_a_series(self):
        assert TRIM(None).value == ""
        assert TRIM(12.5).value == "12.5"
        assert TRIM(mo.Variable([" x ", "y  z", None])).value == ["x", "y z", ""]


# ══════════════════════════════════════════════════════════════════════
# Finding text
# ══════════════════════════════════════════════════════════════════════


class TestFind:
    @pytest.mark.parametrize("find, within, expected", [
        ("-", "AB-123", 3),
        ("a", "banana", 2),
        ("A", "banana", "#VALUE!"),
        ("na", "banana", 3),
        ("x", "abc", "#VALUE!"),
        ("", "abc", 1),
        ("?", "a?b", 2),
        ("*", "a*b", 2),
        ("a", "", "#VALUE!"),
        (1, 2015, 3),
    ])
    def test_case_sensitive_without_wildcards(self, find, within, expected):
        assert FIND(find, within).value == expected

    @pytest.mark.parametrize("start, expected", [
        (1, 2), (2, 2), (3, 4), (2.9, 2), (5, 6), (6, 6), (7, "#VALUE!"),
        (0, "#VALUE!"), (-1, "#VALUE!"),
    ])
    def test_from_a_start_position(self, start, expected):
        assert FIND("a", "banana", start).value == expected

    def test_empty_text_is_found_at_the_start_position(self):
        # Excel's documented rule: empty find_text matches at start_num,
        # and a start past the end is #VALUE! whatever is looked for.
        # LibreOffice gives #VALUE! for any empty find_text (pinned in
        # test_excel_text.py).
        assert FIND("", "abc").value == 1
        assert FIND("", "abc", 3).value == 3
        assert FIND("", "abc", 4).value == "#VALUE!"
        assert FIND("", "").value == "#VALUE!"
        assert FIND("", "abc", 0).value == "#VALUE!"
        assert SEARCH("", "abc").value == 1
        assert SEARCH("", "abc", 4).value == "#VALUE!"
        assert SEARCH("*", "").value == "#VALUE!"

    def test_a_series_a_position_type_and_the_formula(self):
        names = mo.Variable(["Alpha Fund", "beta co", None])
        found = FIND("a", names, 2)
        assert found.value == [5, 4, "#VALUE!"]
        assert found.value_type == 'int' and found.var_type == 'list'
        assert FIND("-", "AB-123").formula == "FIND('-', 'AB-123')"
        assert FIND("-", "AB-123", 2).formula == "FIND('-', 'AB-123', 2)"


class TestSearch:
    @pytest.mark.parametrize("find, within, expected", [
        ("A", "banana", 2),
        ("NA", "banana", 3),
        ("b*d", "xabcdx", 3),
        ("a?c", "xabcx", 2),
        ("a~?", "xa?b", 2),
        ("~*", "a*b", 2),
        ("~~", "a~b", 2),
        ("~a", "xab", 2),
        ("a~", "a~b", 1),
        ("~", "a~b", 2),
        ("?", "ab", 1),
        ("x", "abc", "#VALUE!"),
        ("a*f", "Alpha Fund", 1),
        ("a*f", "beta co", "#VALUE!"),
    ])
    def test_case_insensitive_with_wildcards(self, find, within, expected):
        assert SEARCH(find, within).value == expected

    def test_from_a_start_position(self):
        assert SEARCH("a", "Banana", 3).value == 4
        assert SEARCH("a", "Banana", 7).value == "#VALUE!"
        assert SEARCH("a", "Banana", 0).value == "#VALUE!"

    def test_a_series_and_the_formula(self):
        names = mo.Variable(["Alpha Fund", "beta co", None])
        assert SEARCH("a*f", names).value == [1, "#VALUE!", "#VALUE!"]
        assert SEARCH("A", "banana", 3).formula == "SEARCH('A', 'banana', 3)"


# ══════════════════════════════════════════════════════════════════════
# Rewriting text
# ══════════════════════════════════════════════════════════════════════


class TestSubstitute:
    @pytest.mark.parametrize("text, old, new, expected", [
        ("AB-123", "-", "", "AB123"),
        ("a.b.c", ".", ",", "a,b,c"),
        ("banana", "a", "o", "bonono"),
        ("banana", "A", "o", "banana"),
        ("abc", "", "x", "abc"),
        ("abc", "x", "y", "abc"),
        (1234.5, ".", ",", "1234,5"),
        (None, "a", "b", ""),
    ])
    def test_every_occurrence(self, text, old, new, expected):
        assert SUBSTITUTE(text, old, new).value == expected

    @pytest.mark.parametrize("text, old, new, instance, expected", [
        ("banana", "a", "o", 1, "bonana"),
        ("banana", "a", "o", 3, "banano"),
        ("banana", "a", "o", 4, "banana"),
        ("abab", "b", "x", 1.9, "axab"),
        ("aaa", "aa", "b", 1, "ba"),
        ("aaa", "aa", "b", 2, "aaa"),
        ("a.b.c", ".", "", 0, "#VALUE!"),
        ("a.b.c", ".", "", -1, "#VALUE!"),
        ("a.b.c", ".", "", "x", "#VALUE!"),
    ])
    def test_one_instance(self, text, old, new, instance, expected):
        assert SUBSTITUTE(text, old, new, instance).value == expected

    def test_a_series_and_the_formula(self):
        names = mo.Variable(["Alpha Fund", "beta co", None])
        assert SUBSTITUTE(names, "a", "A", 1).value == ["AlphA Fund", "betA co", ""]
        assert SUBSTITUTE(names, " ", "_").value == ["Alpha_Fund", "beta_co", ""]
        assert SUBSTITUTE("a-b", "-", "").formula == "SUBSTITUTE('a-b', '-', '')"
        assert SUBSTITUTE("a-b", "-", "", 1).formula == "SUBSTITUTE('a-b', '-', '', 1)"


class TestRept:
    @pytest.mark.parametrize("text, n, expected", [
        ("ab", 3, "ababab"),
        ("ab", 2.9, "abab"),
        ("ab", 0, ""),
        ("", 5, ""),
        ("#", 3, "###"),
        (1.5, 2, "1.51.5"),
        (None, 3, ""),
        ("ab", -1, "#VALUE!"),
        ("ab", "x", "#VALUE!"),
        ("a", 32767, "a" * 32767),
        ("a", 32768, "#VALUE!"),
    ])
    def test_repeats(self, text, n, expected):
        assert REPT(text, n).value == expected

    def test_a_series_and_the_formula(self):
        assert REPT("|", mo.Variable([3, 0, 1])).value == ["|||", "", "|"]
        assert REPT("|", 3).formula == "REPT('|', 3)"


# ══════════════════════════════════════════════════════════════════════
# Numbers from text and text from numbers
# ══════════════════════════════════════════════════════════════════════


class TestValue:
    @pytest.mark.parametrize("text, expected", [
        ("1,234", 1234),
        ("1,234,567.5", 1234567.5),
        ("12%", 0.12),
        ("(1,234)", -1234),
        ("(5)", -5),
        ("1e3", 1000.0),
        ("  -12.5%  ", -0.125),
        ("+5", 5),
        (".5", 0.5),
        ("5.", 5.0),
        ("-0.75", -0.75),
        ("1234,567", 1234567),
        ("1,234.", 1234.0),
        ("1,234.5e2", 123450.0),
        ("- 5", -5),
        ("5 %", 0.05),
        ("( 5 )", -5),
        ("(1,234)%", -12.34),
        ("-5%", -0.05),
        ("1E+15", 1e15),
    ])
    def test_numeric_text(self, text, expected):
        assert VALUE(text).value == expected

    def test_an_iso_date_is_its_serial(self):
        assert VALUE("2026-01-01").value == 46023
        assert VALUE(" 2026-01-31 ").value == 46053
        assert VALUE("2026-13-01").value == "#VALUE!"

    @pytest.mark.parametrize("text, expected", [
        ("12:00", 0.5),
        ("1:30", 0.0625),
        ("01:30", 0.0625),
        ("12:30", 12.5 / 24),
        ("12:5", (12 * 60 + 5) / 1440),
        ("25:00", 25 / 24),
        ("12:30:45", (12 * 3600 + 30 * 60 + 45) / 86400),
        ("12:30:45.5", (12 * 3600 + 30 * 60 + 45.5) / 86400),
        ("0:00", 0.0),
        ("24:00", 1.0),
        ("1:30 PM", 0.5625),
        ("1:30pm", 0.5625),
        ("12:00 AM", 0.0),
        ("12:30 AM", 0.5 / 24),
        ("12 PM", 0.5),
        ("1 AM", 1 / 24),
        ("12:30:45 PM", (12 * 3600 + 30 * 60 + 45) / 86400),
        (" 3:00 PM ", 0.625),
        ("+1:30", 0.0625),
        ("+12 PM", 0.5),
        ("2026-01-31 18:00", 46053.75),
        ("2026-01-31  12:30:15", 46053 + (12 * 3600 + 30 * 60 + 15) / 86400),
        ("2026-01-31 25:00", 46053 + 25 / 24),
    ])
    def test_a_time_of_day_is_its_fraction_of_a_day(self, text, expected):
        assert VALUE(text).value == pytest.approx(expected, abs=1e-12)

    @pytest.mark.parametrize("text", [
        "12:60", "0:90", "1:30:60", "13:30 PM", "1:30 P", "1:30 a.m.", ":30", "12:", "1.5:30",
        "-1:30", "-12 PM", "+-1:30", "2026-01-31T12:00", "2026-01-31 1:30 P",
        "2026-02-30 12:00",
    ])
    def test_other_time_text_is_an_error(self, text):
        assert VALUE(text).value == "#VALUE!"

    @pytest.mark.parametrize("text", [
        "abc", "", "  ", "1,23", "12,34,567", "1,234,", ",123", "1 234", "$1,234", "(-5)", "-(5)",
        "+-5", "--5", "%", "%5", "5%%", "(5%)", "1e3%", "1e", "0x10", "1.2.3", "01/02/2026",
        "TRUE", "\u0663", True,
    ])
    def test_other_text_is_an_error(self, text):
        assert VALUE(text).value == "#VALUE!"

    def test_a_number_too_large_for_a_cell(self):
        assert VALUE("1e999").value == "#NUM!"
        assert VALUE("1" + "0" * 400).value == "#NUM!"
        assert VALUE("1" + "0" * 5000).value == "#NUM!"

    def test_an_integral_text_is_an_int_while_a_cell_holds_it_exactly(self):
        assert VALUE("123456789012345").value == 123456789012345
        assert isinstance(VALUE("000000000000001").value, int)
        long = VALUE("12345678901234567891").value
        assert isinstance(long, float) and long == 1.2345678901234567e19

    def test_a_number_a_blank_and_a_date(self):
        assert VALUE(12.5).value == 12.5
        assert VALUE(None).value == 0
        assert VALUE(date(2026, 1, 1)).value == 46023
        assert VALUE(timedelta(hours=36)).value == 1.5
        assert VALUE(float("inf")).value == "#NUM!"

    def test_a_series_its_type_and_the_formula(self):
        texts = mo.Variable(["1,234", "12%", "x"], display_name="texts")
        got = VALUE(texts)
        assert got.value == [1234, 0.12, "#VALUE!"]
        assert got.value_type == 'float' and got.var_type == 'list'
        assert got.formula == "VALUE(texts)"


class TestText:
    @pytest.mark.parametrize("value, fmt, expected", [
        (2.5, "0", "3"),
        (-2.5, "0", "-3"),
        (0.5, "0", "1"),
        (-0.4, "0", "-0"),
        (12.345, "0.0", "12.3"),
        (1.005, "0.00", "1.01"),
        (1234.5678, "0.000", "1234.568"),
        (1234.5678, "#,##0", "1,235"),
        (-1234.5, "#,##0", "-1,235"),
        (1e6, "#,##0", "1,000,000"),
        (1234.5678, "#,##0.00", "1,234.57"),
        (1234.5678, "0%", "123457%"),
        (0.1234, "0.0%", "12.3%"),
        (0.125, "0.00%", "12.50%"),
        (7, "0.00", "7.00"),
        (None, "0", "0"),
        (5, "00", "05"),
        (-5, "000", "-005"),
        (1234, "00", "1234"),
        (0.5, "00", "01"),
        (7, "000%", "700%"),
        (3, "000.0", "003.0"),
        # the padding counts digits; the thousands are grouped after it
        (5, "#,##00", "05"),
        (1234, "#,##000000", "001,234"),
        (-1234.5, "#,##000000.0", "-001,234.5"),
        (1234567, "#,##0000000000", "0,001,234,567"),
    ])
    def test_number_formats(self, value, fmt, expected):
        assert TEXT(value, fmt).value == expected

    @pytest.mark.parametrize("fmt, expected", [
        ("yyyy-mm-dd", "2026-01-01"),
        ("dd.mm.yyyy", "01.01.2026"),
        ("mmm yyyy", "Jan 2026"),
        ("yyyy", "2026"),
        ("yy", "26"),
        ("mmmm d, yyyy", "January 1, 2026"),
        ("d mmmm yyyy", "1 January 2026"),
        ("ddd dd", "Thu 01"),
        ("dddd", "Thursday"),
        ("m/d/yyyy", "1/1/2026"),
        ("YYYY-MM-DD", "2026-01-01"),
        ("mmm", "Jan"),
        ("mmmmm", "J"),
        ("mmmmmm", "J1"),
        ("ddddd", "Thursday1"),
    ])
    def test_date_formats_over_a_date_and_its_serial(self, fmt, expected):
        assert TEXT(date(2026, 1, 1), fmt).value == expected
        assert TEXT(46023, fmt).value == expected

    def test_general(self):
        assert TEXT(1234.5678, "General").value == "1234.5678"
        assert TEXT(1 / 3, "general").value == "0.333333333333333"
        assert TEXT(0.1 + 0.2, "General").value == "0.3"
        assert TEXT(2022.0, "@").value == "2022"
        assert TEXT(date(2026, 1, 1), "General").value == "46023"
        assert TEXT(0.00001, "General").value == "0.00001"
        assert TEXT(1e15, "General").value == "1000000000000000"
        assert TEXT(1e20, "General").value == "1E+20"

    @pytest.mark.parametrize("value, fmt, expected", [
        (123456789012345678.0, "0", "123456789012346000"),
        (123456789012345678.0, "#,##0.00", "123,456,789,012,346,000.00"),
        (1234567890123456, "#,##0", "1,234,567,890,123,460"),
        (1 / 3, "0.000000000000000000", "0.333333333333333000"),
        (0.1 + 0.2, "0.000000000000000000", "0.300000000000000000"),
        (0.1 + 0.2, "0.00000000000000000%", "30.00000000000000000%"),
        (12345678901234.5678, "0.0000", "12345678901234.6000"),
        (2.675, "0.00", "2.68"),
        (1e300, "0", "1" + "0" * 300),
    ])
    def test_a_format_shows_15_significant_digits(self, value, fmt, expected):
        assert TEXT(value, fmt).value == expected

    @pytest.mark.parametrize("value, expected", [
        (1.5, ""), (0, ""), (-1.5, ""), (None, ""), ("12", ""), (date(2026, 1, 1), ""),
        ("abc", "abc"), (True, "TRUE"),
    ])
    def test_an_empty_format_shows_nothing(self, value, expected):
        assert TEXT(value, "").value == expected

    def test_a_number_no_cell_holds_is_an_error(self):
        assert TEXT(float("nan"), "0").value == "#NUM!"
        assert TEXT(float("inf"), "General").value == "#NUM!"

    def test_text_and_booleans(self):
        assert TEXT("12", "0.00").value == "12.00"
        assert TEXT("12%", "0.00").value == "0.12"
        assert TEXT("2026-01-01", "mmm yyyy").value == "Jan 2026"
        assert TEXT("abc", "0").value == "abc"
        assert TEXT("abc", "yyyy").value == "abc"
        assert TEXT(True, "0").value == "TRUE"
        assert TEXT(False, "0.0").value == "FALSE"

    def test_the_format_is_text_even_when_it_starts_with_a_hash(self):
        assert TEXT(mo.Variable([1234.5, -0.4, None]), "#,##0").value == ["1,235", "-0", "0"]
        assert TEXT(1234.5, mo.Variable("#,##0.0")).value == "1,234.5"

    def test_a_format_that_varies_by_period(self):
        got = TEXT(1234.5, mo.Variable(["0", "0.0", "#,##0"]))
        assert got.value == ["1235", "1234.5", "1,235"]
        assert got.var_type == 'list'

    def test_a_series_its_type_and_the_formula(self):
        amounts = mo.Variable([1234.5678, -0.4, 2.5], display_name="amounts")
        got = TEXT(amounts, "#,##0.00")
        assert got.value == ["1,234.57", "-0.40", "2.50"]
        assert got.value_type == 'string' and got.var_type == 'list'
        assert got.formula == "TEXT(amounts, '#,##0.00')"

    def test_an_unknown_format_warns_once_and_gives_general_text(self, caplog):
        with caplog.at_level(logging.WARNING, logger="modeleon.functions.text"):
            got = TEXT(mo.Variable([1.5, 2.5]), "0.0E+00")
        assert got.value == ["1.5", "2.5"]
        warnings = [r for r in caplog.records if "0.0E+00" in r.getMessage()]
        assert len(warnings) == 1
        with caplog.at_level(logging.WARNING, logger="modeleon.functions.text"):
            caplog.clear()
            assert TEXT(46023, "hh:mm").value == "46023"
        assert any("hh:mm" in r.getMessage() for r in caplog.records)

    def test_a_known_format_does_not_warn(self, caplog):
        with caplog.at_level(logging.WARNING, logger="modeleon.functions.text"):
            TEXT(1.5, "0.0")
            TEXT(46023, "yyyy-mm-dd")
            TEXT(1.5, "General")
        assert not caplog.records


# ══════════════════════════════════════════════════════════════════════
# Joining a row
# ══════════════════════════════════════════════════════════════════════


class TestTextjoin:
    def test_joins_a_series_into_one_text(self):
        names = mo.Variable(["Alpha Fund", "beta co", None], display_name="names")
        got = TEXTJOIN(", ", True, names)
        assert got.value == "Alpha Fund, beta co"
        assert got.var_type == 'scalar' and got.value_type == 'string'
        assert got.formula == "TEXTJOIN(', ', True, names)"

    def test_ignore_empty(self):
        assert TEXTJOIN("-", True, "a", "", None, "b").value == "a-b"
        assert TEXTJOIN("-", False, "a", "", None, "b").value == "a---b"
        assert TEXTJOIN("-", 0, "a", "", "b").value == "a--b"
        assert TEXTJOIN("-", None, "a", "", "b").value == "a--b"
        assert TEXTJOIN("-", 2, "a", "", "b").value == "a-b"
        assert TEXTJOIN("-", "0", "a", "", "b").value == "a--b"
        assert TEXTJOIN("-", "TRUE", "a", "", "b").value == "a-b"
        assert TEXTJOIN("-", "false", "a", "", "b").value == "a--b"
        assert TEXTJOIN("-", "x", "a", "", "b").value == "#VALUE!"
        assert TEXTJOIN("-", mo.Variable(True), "a", "", "b").value == "a-b"
        assert TEXTJOIN("-", False, mo.Variable(["x", None, "y"])).value == "x--y"

    def test_scalars_series_and_the_delimiter_from_a_cell(self):
        codes = mo.Variable(["AB-123", "CD-45", "EF-6"])
        sep = mo.Variable(" | ")
        assert TEXTJOIN(sep, False, codes, "end").value == "AB-123 | CD-45 | EF-6 | end"
        assert TEXTJOIN("", True, "a", "b").value == "ab"
        assert TEXTJOIN(",", True, ["a", "b"], "c").value == "a,b,c"
        assert TEXTJOIN("#", True, "a", "b").value == "a#b"

    def test_operands_read_as_the_spreadsheet_reads_them(self):
        assert TEXTJOIN(", ", True, 1.5, True, date(2026, 1, 1), 2022.0).value == \
            "1.5, TRUE, 46023, 2022"

    def test_an_error_among_the_operands_is_the_result(self):
        assert TEXTJOIN(",", True, mo.Variable(["a", "#DIV/0!", "b"])).value == "#DIV/0!"
        assert TEXTJOIN(mo.Variable("#N/A"), True, "a").value == "#N/A"
        assert TEXTJOIN(",", True, "#12", "b").value == "#12,b"

    def test_a_text_longer_than_a_cell_holds_is_an_error(self):
        assert TEXTJOIN("", True, "a" * 20000, "b" * 20000).value == "#VALUE!"

    def test_a_series_delimiter_or_flag_is_refused(self):
        with pytest.raises(ValueError, match="delimiter is one value"):
            TEXTJOIN(mo.Variable([",", ";"]), True, "a")
        with pytest.raises(ValueError, match="ignore_empty is one value"):
            TEXTJOIN(",", mo.Variable([True, False]), "a")

    def test_no_text_is_refused(self):
        with pytest.raises(ValueError, match=r"Syntax: mo\.TEXTJOIN.*delimiter and ignore_empty"):
            TEXTJOIN(",", True)

    def test_a_line_with_tracks_is_refused(self):
        m = mo.Model("m", tracks=mo.Tracks("plan", "actual"),
                     default_grain="month", default_start="2026-01", default_periods=2)
        with m:
            m.x = mo.Variable(plan=["a", "b"], actual=["c", None])
            with pytest.raises(ValueError, match="one track at a time"):
                TEXTJOIN(",", True, m.x)

    def test_the_ast_and_the_registry(self):
        names = mo.Variable(["a", "b"])
        expr = TEXTJOIN(", ", True, names)._expr
        assert isinstance(expr, FuncCall) and expr.func == "TEXTJOIN"
        assert isinstance(expr.args[0], Literal) and expr.args[0].value == ", "
        assert isinstance(expr.args[1], Literal) and expr.args[1].value is True
        assert isinstance(expr.args[2], VarRef) and expr.args[2].var is names
        assert spec("TEXTJOIN").shape == 'range'
        assert is_range_taking("TEXTJOIN") and range_positions("TEXTJOIN") is None
        assert file_prefix("TEXTJOIN") == "_xlfn." and excel_name("TEXTJOIN") == "TEXTJOIN"
        assert not admits_in_loop("TEXTJOIN", 3)


# ══════════════════════════════════════════════════════════════════════
# The shared law, across the element-wise helpers
# ══════════════════════════════════════════════════════════════════════


class TestTheSharedLaw:
    def test_scalars_only_is_one_scalar(self):
        for fn, args in ((LEFT, ("abc", 2)), (FIND, ("b", "abc")), (VALUE, ("12",)),
                         (TEXT, (1.5, "0.0")), (REPT, ("a", 2)), (TRIM, (" a ",))):
            out = fn(*args)
            assert out.var_type == 'scalar' and not isinstance(out.value, list), fn

    def test_a_one_value_series_broadcasts(self):
        assert LEFT(mo.Variable(["abc"]), mo.Variable([1, 2, 3])).value == ["a", "ab", "abc"]
        assert REPT(mo.Variable(["ab"]), [1, 2]).value == ["ab", "abab"]

    def test_series_of_different_lengths_are_refused_as_the_operators_refuse_them(self):
        with pytest.raises(ValueError, match="different lengths"):
            LEFT(mo.Variable(["abc", "de", "f"]), [1, 2])

    def test_the_first_error_among_a_period_s_operands_is_the_result(self):
        assert LEFT(mo.Variable(["ab", "#NUM!", "cd"]), 1).value == ["a", "#NUM!", "c"]
        assert MID("#DIV/0!", "#VALUE!", 1).value == "#DIV/0!"
        assert FIND("a", "banana", "#N/A").value == "#N/A"
        assert TEXT(mo.Variable(["#REF!", 1.5]), "0.0").value == ["#REF!", "1.5"]
        assert VALUE("#NAME?").value == "#NAME?"
        assert SUBSTITUTE("abc", "b", "#NULL!").value == "#NULL!"

    def test_text_that_only_starts_with_a_hash_is_text(self):
        assert LEFT("#12", 1).value == "#"
        assert FIND("#", "a#b").value == 2
        assert SUBSTITUTE("#1", "#", "No. ").value == "No. 1"
        assert REPT("#", 3).value == "###"
        assert TEXT(1234, "#,##0").value == "1,234"
        assert VALUE("#12").value == "#VALUE!"

    @pytest.mark.parametrize("call, value", [
        (lambda: LEFT(None, 2), ""),
        (lambda: RIGHT(None), ""),
        (lambda: MID(None, 1, 1), ""),
        (lambda: TRIM(None), ""),
        (lambda: FIND("0", None), "#VALUE!"),
        (lambda: SEARCH(None, "abc"), 1),
        (lambda: SUBSTITUTE(None, "a", "b"), ""),
        (lambda: SUBSTITUTE("a0b", "0", None), "ab"),
        (lambda: SUBSTITUTE("abc", None, "x"), "abc"),
        (lambda: TEXT(1.5, None), ""),
        (lambda: REPT(None, 3), ""),
        (lambda: TEXTJOIN(None, False, "a", "b"), "ab"),
        (lambda: TEXTJOIN("-", False, "a", None, "b"), "a--b"),
        (lambda: mo.LEN(None), 0),
        (lambda: mo.UPPER(None), ""),
        (lambda: mo.LOWER(["AB", None]), ["ab", ""]),
        (lambda: mo.CONCAT("a", None, "b"), "ab"),
    ])
    def test_a_literal_blank_where_text_is_read_is_written_as_empty_text(self, call, value):
        # The workbook spells a literal None as 0, whose text is "0"; a
        # blank where text is read goes in as "" — what Python read.
        out = call()
        assert out.value == value
        literals = [a.value for a in out._expr.args if isinstance(a, Literal)]
        flat = [v for lit in literals for v in (lit if isinstance(lit, list) else [lit])]
        assert None not in flat and "" in flat

    def test_a_gap_in_a_literal_list_is_empty_text(self):
        out = LEFT(["ab", None, "cd"], 1)
        assert out.value == ["a", "", "c"]
        assert out._expr.args[0].value == ["ab", "", "cd"]

    def test_a_blank_count_stays_a_blank(self):
        # A count reads a blank as 0, as the workbook's 0 does.
        out = MID("abc", 2, None)
        assert out.value == ""
        assert out._expr.args[2].value is None

    def test_tracked_operands_lift_per_track(self):
        m = mo.Model("m", tracks=mo.Tracks("plan", "actual"),
                     default_grain="month", default_start="2026-01", default_periods=2)
        with m:
            m.x = mo.Variable(plan=["a1", "b2"], actual=["c3", None])
            m.l = LEFT(m.x, 1)
            m.f = FIND("1", m.x)
        assert dict(m.l._value.items()) == {"plan": ["a", "b"], "actual": ["c", ""]}
        assert dict(m.f._value.items()) == {"plan": [2, "#VALUE!"], "actual": ["#VALUE!", "#VALUE!"]}
        assert m.l.var_type == 'list'

    def test_an_axised_operand_is_refused(self):
        region = mo.Variable(["N", "S"], display_name="region")
        by_region = mo.Variable(["x", "y"], indexed_by=[region])
        with pytest.raises(ValueError, match="LEFT"):
            LEFT(by_region, 1)

    def test_value_types(self):
        text = mo.Variable(["abc", "de"])
        for fn, args in ((LEFT, (text, 1)), (RIGHT, (text, 1)), (MID, (text, 1, 1)),
                         (TRIM, (text,)), (SUBSTITUTE, (text, "a", "b")),
                         (TEXT, (mo.Variable([1.5, 2.5]), "0")), (REPT, (text, 2))):
            assert fn(*args).value_type == 'string', fn
        assert FIND("a", text).value_type == 'int'
        assert SEARCH("a", text).value_type == 'int'
        assert VALUE(mo.Variable(["1", "2"])).value_type == 'float'

    def test_no_result_carries_a_unit(self):
        labels = mo.Variable(["a1", "b2"], unit="$")
        amount = mo.Variable([1234.5, 2.5], unit="$")
        k = mo.Variable(2, unit="$")
        for out in (LEFT(labels, k), RIGHT(labels, 1), MID(labels, 1, 1), TRIM(labels),
                    FIND("1", labels), SEARCH("1", labels), SUBSTITUTE(labels, "1", "2"),
                    VALUE(amount), TEXT(amount, "#,##0"), REPT(labels, k),
                    TEXTJOIN(",", True, labels)):
            assert out._unit is None, out.formula

    def test_the_ast_is_a_call_over_the_operands(self):
        names = mo.Variable(["abc", "de"])
        expr = LEFT(names, 2)._expr
        assert isinstance(expr, FuncCall) and expr.func == "LEFT"
        assert isinstance(expr.args[0], VarRef) and expr.args[0].var is names
        assert isinstance(expr.args[1], Literal) and expr.args[1].value == 2
        for name, fn in ELEMENTWISE.items():
            assert fn(*(["abc", 1, 1][:_arity(name)]))._expr.func == name

    def test_every_helper_is_registered_one_call_per_period(self):
        for name in ELEMENTWISE:
            assert is_elementwise(name), name
            assert admits_in_loop(name, 2), name
            assert file_prefix(name) == "" and excel_name(name) == name, name
            assert not spec(name).bare, name


def _arity(name: str) -> int:
    return {"TRIM": 1, "VALUE": 1, "MID": 3, "SUBSTITUTE": 3}.get(name, 2)
