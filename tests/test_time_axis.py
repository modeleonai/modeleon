# SPDX-License-Identifier: Apache-2.0
"""Tests for ``mo.Time`` — the discrete time axis (P0a, slice 1).

``mo.Time`` builds a period-labeled axis Variable. It rides the existing
``indexed_by`` / ``Shape`` machinery (an axis is just a Variable, per
``shape.py``) — so a time-varying Variable declares ``indexed_by=[time]``
and shape inference already works. This slice covers monthly grain only;
extent is start + (end | periods). Other grains and grain-on-projection
land in later slices.
"""

from __future__ import annotations

import pytest

import modeleon as mo
from modeleon.core.shape import Shape


class TestTimeLabels:
    """``mo.Time`` generates the period labels as its value."""

    def test_month_start_end(self):
        t = mo.Time('2024-01', '2024-03')
        assert t._value == ['2024-01', '2024-02', '2024-03']

    def test_month_start_periods(self):
        t = mo.Time('2024-01', periods=3)
        assert t._value == ['2024-01', '2024-02', '2024-03']

    def test_spans_year_boundary(self):
        t = mo.Time('2024-11', '2025-02')
        assert t._value == ['2024-11', '2024-12', '2025-01', '2025-02']

    def test_length_matches_periods(self):
        assert len(mo.Time('2024-01', periods=12)._value) == 12

    def test_single_period(self):
        assert mo.Time('2024-01', '2024-01')._value == ['2024-01']


class TestTimeIsAnAxisVariable:
    """``mo.Time`` is a Variable, so it works as an axis in ``indexed_by``
    and the existing shape machinery picks it up unchanged."""

    def test_is_a_variable(self):
        assert isinstance(mo.Time('2024-01', periods=2), mo.Variable)

    def test_marked_as_time_axis(self):
        assert mo.Time('2024-01', periods=2)._axis_kind == 'time'

    def test_usable_as_indexed_by_axis(self):
        t = mo.Time('2024-01', periods=3)
        revenue = mo.Variable([100, 110, 121], indexed_by=[t])
        assert revenue._indexed_by == (t,)
        assert revenue.shape == Shape((t,))

    def test_arithmetic_preserves_time_axis(self):
        t = mo.Time('2024-01', periods=3)
        revenue = mo.Variable([100, 110, 121], indexed_by=[t])
        cogs = revenue * 0.6
        assert cogs._indexed_by == (t,)
        assert cogs._value == [60.0, 66.0, 72.6]


class TestTimeValidation:
    """Extent must be unambiguous; unsupported grains fail loudly."""

    def test_requires_end_or_periods(self):
        with pytest.raises(ValueError):
            mo.Time('2024-01')

    def test_rejects_both_end_and_periods(self):
        with pytest.raises(ValueError):
            mo.Time('2024-01', '2024-06', periods=6)

    def test_end_before_start(self):
        with pytest.raises(ValueError):
            mo.Time('2024-06', '2024-01')

    def test_nonpositive_periods(self):
        with pytest.raises(ValueError):
            mo.Time('2024-01', periods=0)

    def test_unknown_grain_raises(self):
        # 'week' is a deferred grain; 'fortnight' is nonsense — both refused.
        with pytest.raises(ValueError):
            mo.Time('2024-01', periods=4, grain='fortnight')
        with pytest.raises(ValueError):
            mo.Time('2024-01', periods=4, grain='week')


class TestMultiGrain:
    """Day / quarter / year grains, alongside the default month."""

    def test_day_start_end(self):
        t = mo.Time('2024-01-01', '2024-01-04', grain='day')
        assert t._value == ['2024-01-01', '2024-01-02', '2024-01-03', '2024-01-04']

    def test_day_crosses_month_boundary(self):
        t = mo.Time('2024-01-30', periods=3, grain='day')
        assert t._value == ['2024-01-30', '2024-01-31', '2024-02-01']

    def test_day_bad_format_raises(self):
        with pytest.raises(ValueError):
            mo.Time('2024-01', periods=3, grain='day')   # needs YYYY-MM-DD

    def test_quarter_start_end(self):
        t = mo.Time('2024-Q1', '2024-Q4', grain='quarter')
        assert t._value == ['2024-Q1', '2024-Q2', '2024-Q3', '2024-Q4']

    def test_quarter_crosses_year_boundary(self):
        t = mo.Time('2024-Q3', periods=3, grain='quarter')
        assert t._value == ['2024-Q3', '2024-Q4', '2025-Q1']

    def test_year_start_end(self):
        t = mo.Time('2024', '2026', grain='year')
        assert t._value == ['2024', '2025', '2026']

    def test_year_periods(self):
        assert mo.Time('2024', periods=3, grain='year')._value == ['2024', '2025', '2026']

    def test_grain_is_recorded(self):
        assert mo.Time('2024-Q1', periods=2, grain='quarter')._grain == 'quarter'
