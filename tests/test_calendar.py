# SPDX-License-Identifier: Apache-2.0
"""The calendar: rulers answer four questions; a family says how they nest."""
from datetime import date

import pytest

from modeleon.core.calendar import (
    DAY, GREGORIAN, HALF, ISO_YEAR, MONTH, QUARTER, WEEK, YEAR, Calendar, Ruler, _closure,
)


class TestTheFourQuestions:
    @pytest.mark.parametrize("ruler, label, first", [
        (DAY, "2026-01-05", date(2026, 1, 5)),
        (MONTH, "2026-03", date(2026, 3, 1)),
        (QUARTER, "2026-Q3", date(2026, 7, 1)),
        (HALF, "2026-H2", date(2026, 7, 1)),
        (YEAR, "2026", date(2026, 1, 1)),
        (WEEK, "2026-W01", date(2025, 12, 29)),
        (WEEK, "2026-W53", date(2026, 12, 28)),
        (ISO_YEAR, "2026", date(2025, 12, 29)),
    ])
    def test_parse_names_a_first_day_and_label_reads_it_back(self, ruler, label, first):
        assert ruler.parse(label) == first
        assert ruler.label(first) == label

    @pytest.mark.parametrize("ruler, day, first", [
        (DAY, date(2026, 2, 14), date(2026, 2, 14)),
        (MONTH, date(2026, 2, 14), date(2026, 2, 1)),
        (QUARTER, date(2026, 2, 14), date(2026, 1, 1)),
        (QUARTER, date(2026, 11, 30), date(2026, 10, 1)),
        (HALF, date(2026, 6, 30), date(2026, 1, 1)),
        (HALF, date(2026, 7, 1), date(2026, 7, 1)),
        (YEAR, date(2026, 11, 30), date(2026, 1, 1)),
        (WEEK, date(2026, 1, 1), date(2025, 12, 29)),       # a Thursday
        (ISO_YEAR, date(2027, 1, 2), date(2025, 12, 29)),   # a Saturday of ISO week 2026-W53
    ])
    def test_start_of_is_the_period_holding_the_day(self, ruler, day, first):
        assert ruler.start_of(day) == first

    @pytest.mark.parametrize("ruler, first, nxt", [
        (DAY, date(2024, 2, 28), date(2024, 2, 29)),
        (MONTH, date(2026, 12, 1), date(2027, 1, 1)),
        (QUARTER, date(2026, 10, 1), date(2027, 1, 1)),
        (YEAR, date(2026, 1, 1), date(2027, 1, 1)),
    ])
    def test_next_start_crosses_the_year_end(self, ruler, first, nxt):
        assert ruler.next_start(first) == nxt

    def test_measure_is_the_calendars_day_count(self):
        assert DAY.measure(date(2026, 1, 1)) == 1
        assert MONTH.measure(date(2024, 2, 1)) == 29
        assert MONTH.measure(date(2026, 2, 1)) == 28
        assert QUARTER.measure(date(2026, 1, 1)) == 90
        assert HALF.measure(date(2026, 1, 1)) == 181
        assert HALF.measure(date(2026, 7, 1)) == 184
        assert YEAR.measure(date(2024, 1, 1)) == 366
        assert WEEK.measure(date(2026, 1, 5)) == 7
        assert ISO_YEAR.measure(date(2025, 12, 29)) == 7 * 53     # ISO 2026 has 53 weeks
        assert ISO_YEAR.measure(date(2027, 1, 4)) == 7 * 52

    def test_count_is_inclusive(self):
        assert MONTH.count(date(2025, 11, 1), date(2026, 2, 1)) == 4
        assert QUARTER.count(date(2026, 1, 1), date(2026, 1, 1)) == 1
        assert DAY.count(date(2026, 1, 30), date(2026, 2, 2)) == 4
        assert YEAR.count(date(2024, 1, 1), date(2026, 1, 1)) == 3

    def test_styles(self):
        assert MONTH.label(date(2026, 3, 1), "finance") == "Mar 2026"
        assert MONTH.label(date(2026, 3, 1), "compact") == "Mar-26"
        assert QUARTER.label(date(2026, 7, 1), "finance") == "Q3 2026"
        assert HALF.label(date(2026, 7, 1), "finance") == "H2 2026"
        assert HALF.label(date(2026, 7, 1), "compact") == "H2-26"
        assert HALF.label(date(2026, 1, 1)) == "2026-H1"
        assert YEAR.label(date(2026, 1, 1), "compact") == "FY26"
        assert WEEK.label(date(2025, 12, 29), "finance") == "W01 2026"
        assert WEEK.label(date(2025, 12, 29), "compact") == "W01-26"
        assert ISO_YEAR.label(date(2025, 12, 29)) == "2026"
        assert DAY.label(date(2026, 1, 5), "finance") == "05 Jan 2026"

    @pytest.mark.parametrize("ruler, label, message", [
        (DAY, "2026-W01", "Time day start must be 'YYYY-MM-DD'"),
        (MONTH, "2026-13", "Month must be 01-12"),
        (MONTH, "2026", "Time month start must be 'YYYY-MM'"),
        (QUARTER, "2026-Q5", "Quarter must be Q1-Q4"),
        (HALF, "2026-H3", "Half must be H1-H2"),
        (HALF, "2026-Q1", "Time half start must be 'YYYY-Hn'"),
        (YEAR, "FY26", "Time year start must be 'YYYY'"),
        (WEEK, "2026-01", "Time week start must be 'YYYY-Www'"),
        (WEEK, "2027-W53", "names no ISO week: the year 2027 has 52 weeks"),
        (ISO_YEAR, "W2026", "Time iso_year start must be 'YYYY'"),
    ])
    def test_a_bad_label_is_refused_with_its_grammar(self, ruler, label, message):
        with pytest.raises(ValueError, match=message):
            ruler.parse(label)


class TestTheFamily:
    def test_the_gregorian_rulers_are_the_grains_finest_first(self):
        assert GREGORIAN.grains == ("hour", "day", "week", "month", "quarter", "half", "year",
                                    "iso_year")
        assert [r.rank for r in GREGORIAN.rulers] == [0, 1, 2, 3, 4, 5, 6, 7]
        assert [(r.word, r.letter) for r in GREGORIAN.rulers] == [
            ("Hourly", "Hr"), ("Daily", "D"), ("Weekly", "W"), ("Monthly", "M"),
            ("Quarterly", "Q"), ("Half-yearly", "H"), ("Yearly", "Y"), ("Yearly (ISO)", "IY"),
        ]

    def test_nesting_is_a_relation_not_an_order(self):
        assert GREGORIAN.nests("day", "month")
        assert GREGORIAN.nests("month", "year")
        assert GREGORIAN.nests("quarter", "half") and GREGORIAN.nests("half", "year")
        assert not GREGORIAN.nests("year", "month")
        assert not GREGORIAN.nests("month", "month")
        # A week lies inside one ISO year and in no month, quarter or year.
        assert GREGORIAN.nests("day", "week") and GREGORIAN.nests("week", "iso_year")
        assert GREGORIAN.nests("day", "iso_year")
        for coarse in ("month", "quarter", "half", "year"):
            assert not GREGORIAN.nests("week", coarse)
        assert not GREGORIAN.nests("month", "iso_year")
        assert not GREGORIAN.comparable("year", "iso_year")
        assert GREGORIAN.rolls_up_to("week") == ("iso_year",)
        assert GREGORIAN.rolls_up_to("day") == ("week", "month", "quarter", "half", "year", "iso_year")
        # An hour lies inside its day, and so inside everything a day is.
        assert GREGORIAN.rolls_up_to("hour") == ("day",) + GREGORIAN.rolls_up_to("day")
        assert not GREGORIAN.nests("day", "hour") and not GREGORIAN.nests("hour", "hour")

    def test_every_claim_is_checked_on_the_line(self):
        GREGORIAN.verify()
        # A false claim names its first bad period.
        bad = Calendar("bad", (DAY, WEEK, MONTH), _closure([("day", "week"), ("week", "month")]))
        with pytest.raises(AssertionError, match="week does not nest in month"):
            bad.verify()

    def test_coarser_never_says_yes_to_a_word_it_does_not_know(self):
        assert GREGORIAN.coarser("quarter", "month")
        assert GREGORIAN.coarser("half", "quarter")
        assert not GREGORIAN.coarser("month", "quarter")
        assert not GREGORIAN.coarser("fortnight", "month")
        assert not GREGORIAN.coarser("year", "week")
        assert not GREGORIAN.coarser("month", "week")
        assert GREGORIAN.coarser("iso_year", "week")
        assert not GREGORIAN.coarser(None, "month")

    def test_an_unknown_ruler_is_refused_by_name(self):
        with pytest.raises(ValueError, match="Unknown time grain 'minute'; supported: hour, day, week, month, quarter, half, year, iso_year"):
            GREGORIAN.ruler("minute")
        assert not GREGORIAN.knows("minute")
        assert GREGORIAN.knows("week") and GREGORIAN.knows("hour")

    def test_whole_months_and_stubs(self):
        assert [r.months for r in GREGORIAN.rulers] == [None, None, None, 1, 3, 6, 12, None]
        assert [r.stub_unit for r in GREGORIAN.rulers] == [None, None, None, None, "month", "month", "month", "week"]
        assert [r.stride_days for r in GREGORIAN.rulers] == [None, 1, 7, None, None, None, None, None]
        assert [r.stride_hours for r in GREGORIAN.rulers] == [1, None, None, None, None, None, None, None]
        assert [r.name for r in GREGORIAN.rulers if r.sub_day] == ["hour"]

    def test_a_ruler_is_a_rule_not_a_table(self):
        # The base class answers measure and count from next_start alone,
        # so a new ruler needs only the four questions.
        class Fortnight(Ruler):
            name, rank = "fortnight", 0

            def parse(self, label):
                return date.fromisoformat(label)

            def label(self, first, style="iso"):
                return first.isoformat()

            def start_of(self, day):
                return day

            def next_start(self, first):
                from datetime import timedelta
                return first + timedelta(days=14)

        f = Fortnight()
        assert f.measure(date(2026, 1, 1)) == 14
        assert f.count(date(2026, 1, 1), date(2026, 1, 29)) == 3
        cal = Calendar("test", (f,))
        assert cal.grains == ("fortnight",)
        assert repr(cal) == "Calendar('test', ('fortnight',))"
