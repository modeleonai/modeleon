# SPDX-License-Identifier: Apache-2.0
"""The extent law (§16.6) + ``mo.schedule`` (§16.7).

Both are adoption-time behaviors: the ambient window only exists once
the Variable joins a windowed tree, so the hook in
``_register_component`` / ``_run_adoption_hooks`` is where a schedule
materializes and where a wrong-length bare list dies loudly instead of
detonating later at an unrelated broadcast.
"""

from __future__ import annotations

import pytest

import modeleon as mo


def _model(periods: int = 24) -> "mo.Model":
    return mo.Model(
        'm', default_grain='month', default_start='2025-01',
        default_periods=periods,
    )


class TestExtentLaw:
    def test_wrong_length_bare_list_dies_at_adoption(self):
        m = _model()
        with pytest.raises(ValueError, match="24-period window"):
            with m:
                m.короткий = mo.Variable([1.0, 2.0, 3.0])

    def test_full_length_scalar_and_len1_pass(self):
        m = _model(4)
        with m:
            m.полный = mo.Variable([1.0, 2.0, 3.0, 4.0])
            m.скаляр = mo.Variable(0.2)
            m.один = mo.Variable([5.0])
        assert m.полный._value == [1.0, 2.0, 3.0, 4.0]

    def test_own_location_exempts(self):
        # Declared partial: its own start= anchors it — the law's
        # explicit escape hatch for genuinely partial series.
        m = _model()
        with m:
            m.свой = mo.Variable([1.0, 2.0], start='2025-03', grain='month')
        assert m.свой.time.start == '2025-03'

    def test_keyed_list_exempts(self):
        # A labeled list (dict sugar) is not a time series.
        m = _model()
        with m:
            m.кейсы = mo.Variable({'Bear': 1.0, 'Base': 2.0, 'Bull': 3.0})
        assert m.кейсы._keys == ['Bear', 'Base', 'Bull']

    def test_axised_exempts(self):
        # Coordinates, not periods — the rank-1 guard owns that story.
        m = _model()
        сценарии = mo.Variable(['Bear', 'Bull'])
        with m:
            m.осевой = mo.Variable([1.0, 2.0], indexed_by=[сценарии])
        assert m.осевой.time is None

    def test_windowless_model_is_untouched(self):
        with mo.Model('w') as m:
            m.x = mo.Variable([1.0, 2.0, 3.0])
        assert m.x._value == [1.0, 2.0, 3.0]

    def test_operator_results_pass(self):
        m = _model(4)
        with m:
            m.a = mo.Variable([1.0] * 4)
            m.b = m.a * 2
        assert m.b._value == [2.0] * 4


class TestSchedule:
    def test_materializes_on_adoption(self):
        m = _model()
        with m:
            m.параметры = mo.MultiVariable()
            m.параметры.ндс = mo.schedule(
                {'2025-01': 0.12, '2026-01': 0.16}
            )
        v = m.параметры.ндс._value
        assert len(v) == 24
        assert v[0] == 0.12 and v[11] == 0.12
        assert v[12] == 0.16 and v[23] == 0.16

    def test_formula_over_schedule(self):
        m = _model()
        with m:
            m.ставка = mo.schedule({'2025-01': 0.10, '2026-01': 0.20})
            m.цена = mo.Variable([100.0] * 24)
            m.итог = m.цена * (1 + m.ставка)
        assert m.итог._value[0] == pytest.approx(110.0)
        assert m.итог._value[12] == pytest.approx(120.0)

    def test_default_regrain_is_mean(self):
        # A rate re-grains by mean, never by sum.
        m = _model(12)
        with m:
            m.ставка = mo.schedule({'2025-01': 0.12})
        q = m.ставка.at('quarter')
        assert q._value == pytest.approx([0.12, 0.12, 0.12, 0.12])

    def test_step_off_boundary_dies(self):
        m = _model()
        with pytest.raises(ValueError, match="boundary"):
            with m:
                m.ндс = mo.schedule({'2025-01': 0.12, '2026-02-15': 0.16})

    def test_window_start_uncovered_dies(self):
        m = _model()
        with pytest.raises(ValueError, match="window start"):
            with m:
                m.ндс = mo.schedule({'2025-06': 0.12})

    def test_step_beyond_window_is_legal(self):
        # The law only binds steps INSIDE the window span; a future
        # change beyond it simply doesn't materialize.
        m = _model(12)
        with m:
            m.ндс = mo.schedule({'2025-01': 0.12, '2027-01': 0.20})
        assert m.ндс._value == [0.12] * 12

    def test_empty_or_bad_keys_die(self):
        with pytest.raises(TypeError, match="non-empty"):
            mo.schedule({})
        with pytest.raises(TypeError, match="ISO period labels"):
            mo.schedule({2025: 0.12})


class TestExtend:
    """``extend=`` (§16.5, pulled forward): the declared continuation
    legalizes the prefix spelling the extent law otherwise refuses."""

    def test_zero_pads_a_flow(self):
        m = _model()
        with m:
            m.капекс = mo.Variable([500.0, 500.0, 500.0], extend=mo.zero())
        v = m.капекс._value
        assert len(v) == 24
        assert v[:3] == [500.0, 500.0, 500.0]
        assert v[3:] == [0.0] * 21

    def test_hold_carries_a_rate(self):
        m = _model()
        with m:
            m.ставка = mo.Variable([0.10, 0.10, 0.12], extend=mo.hold())
        v = m.ставка._value
        assert len(v) == 24
        assert v[2] == 0.12 and v[23] == 0.12

    def test_full_length_with_extend_is_untouched(self):
        m = _model(3)
        with m:
            m.x = mo.Variable([1.0, 2.0, 3.0], extend=mo.hold())
        assert m.x._value == [1.0, 2.0, 3.0]

    def test_extended_series_computes_downstream(self):
        m = _model(6)
        with m:
            m.поток = mo.Variable([100.0, 100.0], extend=mo.zero())
            m.итог = mo.cumsum(m.поток)
        assert m.итог._value == [100.0, 200.0, 200.0, 200.0, 200.0, 200.0]

    def test_bad_extend_value_dies(self):
        with pytest.raises(TypeError, match="mo.zero"):
            mo.Variable([1.0], extend="zero")

    def test_windowless_extend_stays_partial(self):
        # No window → nothing to materialize to; the declaration is
        # inert (the track layer's date alignment will honor it).
        with mo.Model('w') as m:
            m.x = mo.Variable([1.0, 2.0], extend=mo.zero())
        assert m.x._value == [1.0, 2.0]
