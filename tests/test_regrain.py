# SPDX-License-Identifier: Apache-2.0
"""Tests for re-grain metadata: start=/grain=/regrain= on a Variable.

This slice declares the rule INERTLY — no bucketing, no projection. It
covers the RegrainSpec constructors (mo.up / mo.frozen / mo.ratio), the
time-location metadata (var.time), and the constructor guards. The actual
re-grain (calendar bucketing, two-phase eval, projection) lands later.
"""

from __future__ import annotations

import pytest

import modeleon as mo
from modeleon.core.regrain import RegrainSpec


class TestRegrainConstructors:
    """mo.up / mo.frozen / mo.ratio build a RegrainSpec; rules are stored
    inertly (a recipe name, a callable, or a per-transition dict)."""

    def test_up_named_recipe(self):
        spec = mo.up('sum')
        assert isinstance(spec, RegrainSpec)
        assert spec.default == 'sum'
        assert spec.frozen is False

    def test_up_callable(self):
        def fn(g):
            return g._self
        assert mo.up(fn).default is fn

    def test_up_per_transition_dict(self):
        spec = mo.up({('day', 'week'): 'last', ('week', 'month'): 'mean'})
        assert spec.overrides[('day', 'week')] == 'last'
        assert spec.overrides[('week', 'month')] == 'mean'
        assert spec.default is None

    def test_up_unknown_recipe_raises(self):
        with pytest.raises(ValueError):
            mo.up('median')        # not a known recipe

    def test_up_needs_a_rule(self):
        with pytest.raises(ValueError):
            mo.up()

    def test_up_down_reserved_raises(self):
        with pytest.raises(NotImplementedError):
            mo.up('sum', down='divide')

    def test_frozen(self):
        spec = mo.frozen()
        assert spec.frozen is True
        assert spec.default is None

    def test_ratio_stores_sources(self):
        spec = mo.ratio('rev', 'users')
        assert spec.default.num == 'rev'
        assert spec.default.den == 'users'


class TestTimeLocationMetadata:
    """start=/grain= locate a Variable; var.time reads it back. No
    metadata = time-agnostic."""

    def test_located_variable(self):
        rev = mo.Variable([100, 110, 121], start='2024-01', grain='month',
                          regrain=mo.up('sum'))
        assert rev.time.start == '2024-01'
        assert rev.time.grain == 'month'
        assert rev._regrain.default == 'sum'

    def test_agnostic_variable(self):
        assert mo.Variable(0.2).time is None

    def test_located_without_regrain(self):
        v = mo.Variable([1, 2, 3], start='2024-01', grain='month')
        assert v.time.grain == 'month'
        assert v._regrain is None


class TestGuards:
    """Constructor guards keep the metadata coherent."""

    def test_start_without_grain_raises(self):
        with pytest.raises(ValueError):
            mo.Variable([1, 2, 3], start='2024-01')

    def test_grain_without_start_raises(self):
        with pytest.raises(ValueError):
            mo.Variable([1, 2, 3], grain='month')

    def test_unknown_grain_raises(self):
        with pytest.raises(ValueError):
            mo.Variable([1, 2, 3], start='2024-01', grain='fortnight')

    def test_regrain_without_grain_pends_inheritance(self):
        # grain may be inherited from a parent MV's default_grain, so a rule
        # without an explicit grain is allowed at construction. Standalone, the
        # Variable is effectively agnostic (no resolvable grain) — caught later.
        v = mo.Variable([1, 2, 3], regrain=mo.up('sum'))
        assert v.time is None
        assert v._regrain.default == 'sum'

    def test_regrain_on_formula_allowed(self):
        # Re-grain applies each line's rule to its OWN value, so a formula
        # carries a rule too (a flow sums, a ratio uses mo.ratio).
        v = mo.Variable(formula='x + 1', start='2024-01', grain='month',
                        regrain=mo.up('sum'))
        assert v._regrain.default == 'sum'

    def test_frozen_on_formula_allowed(self):
        # A recurrence/roll-forward series IS a formula but opts out via frozen.
        v = mo.Variable(formula='x + 1', start='2024-01', grain='month',
                        regrain=mo.frozen())
        assert v._regrain.frozen is True

    def test_regrain_must_be_a_spec(self):
        with pytest.raises(TypeError):
            mo.Variable([1, 2, 3], start='2024-01', grain='month', regrain='sum')


class TestDefaultGrainCascade:
    """`default_start` / `default_grain` on an ancestor MV cascade to children
    that declare no grain of their own (the `_owner` -> `_parent` walk)."""

    def test_child_inherits_default(self):
        m = mo.Model('m')
        m.default_start = '2024-01'
        m.default_grain = 'month'
        m.rev = mo.Variable([100, 110, 121], regrain=mo.up('sum'))
        assert m.rev.time.grain == 'month'
        assert m.rev.time.start == '2024-01'

    def test_own_grain_overrides_default(self):
        m = mo.Model('m')
        m.default_start = '2024-01'
        m.default_grain = 'month'
        m.q = mo.Variable([1, 2], start='2024-Q1', grain='quarter')
        assert m.q.time.grain == 'quarter'

    def test_nested_mv_inherits_from_root(self):
        m = mo.Model('m')
        m.default_start = '2024-01'
        m.default_grain = 'month'
        m.pnl = mo.MultiVariable('P&L')
        m.pnl.opex = mo.Variable([10, 20, 30], regrain=mo.up('sum'))
        assert m.pnl.opex.time.grain == 'month'

    def test_no_default_is_agnostic(self):
        m = mo.Model('m')                      # no default_grain anywhere
        m.x = mo.Variable([1, 2, 3])
        assert m.x.time is None


class TestDefaultWindow:
    """`default_periods` completes the model's default projection window
    (start, grain, periods), resolved by `resolve_default_window`."""

    def test_resolves_full_window(self):
        from modeleon.core.time import resolve_default_window
        m = mo.Model('m')
        m.default_start = '2024-01'
        m.default_grain = 'month'
        m.default_periods = 12
        w = resolve_default_window(m)
        assert (w.start, w.grain, w.periods) == ('2024-01', 'month', 12)

    def test_periods_optional(self):
        from modeleon.core.time import resolve_default_window
        m = mo.Model('m')
        m.default_start = '2024-01'
        m.default_grain = 'month'
        assert resolve_default_window(m).periods is None

    def test_resolves_from_a_child_node(self):
        from modeleon.core.time import resolve_default_window
        m = mo.Model('m')
        m.default_start = '2024-01'
        m.default_grain = 'month'
        m.default_periods = 6
        m.pnl = mo.MultiVariable('P&L')
        assert resolve_default_window(m.pnl).periods == 6

    def test_no_default_window(self):
        from modeleon.core.time import resolve_default_window
        assert resolve_default_window(mo.Model('m')) is None


class TestBucketing:
    """Pure (native_grain, target_grain, start) -> bucket ranges — the
    alignment that lets several series share one grouping."""

    def test_month_to_quarter(self):
        from modeleon.core.time import bucket_ranges
        assert bucket_ranges('month', 'quarter', '2024-01', 6) == [
            ('2024-Q1', 0, 3), ('2024-Q2', 3, 6)]

    def test_month_to_year_partial(self):
        from modeleon.core.time import bucket_ranges
        # 6 months of data -> one partial-year bucket.
        assert bucket_ranges('month', 'year', '2024-01', 6) == [('2024', 0, 6)]

    def test_month_to_quarter_offset_start(self):
        from modeleon.core.time import bucket_ranges
        # Start mid-quarter: Feb,Mar -> Q1 (2 cells); Apr,May,Jun -> Q2 (3).
        assert bucket_ranges('month', 'quarter', '2024-02', 5) == [
            ('2024-Q1', 0, 2), ('2024-Q2', 2, 5)]

    def test_day_to_month(self):
        from modeleon.core.time import bucket_ranges
        # Jan30, Jan31, Feb01.
        assert bucket_ranges('day', 'month', '2024-01-30', 3) == [
            ('2024-01', 0, 2), ('2024-02', 2, 3)]

    def test_cannot_coarsen_to_finer(self):
        from modeleon.core.time import bucket_ranges
        with pytest.raises(ValueError):
            bucket_ranges('quarter', 'month', '2024-Q1', 4)


class TestRegrainSeries:
    """Roll a native series up to a coarser grain via a named recipe."""

    def test_sum_to_quarter(self):
        from modeleon.core.time import regrain_series
        labels, vals = regrain_series([1, 2, 3, 4, 5, 6], 'month', 'quarter',
                                      '2024-01', 'sum')
        assert labels == ['2024-Q1', '2024-Q2']
        assert vals == [6, 15]

    def test_last_to_quarter(self):
        from modeleon.core.time import regrain_series
        _, vals = regrain_series([100, 110, 121, 130, 142, 150], 'month',
                                 'quarter', '2024-01', 'last')
        assert vals == [121, 150]      # quarter-end (a stock)

    def test_mean_to_year_is_time_weighted(self):
        from modeleon.core.time import regrain_series
        # mean weights by calendar days (2024 is a leap year: Feb = 29), not
        # by period count — a 31-day month outweighs a 29-day one.
        values = [10, 10, 11, 11, 12, 12]
        days = [31, 29, 31, 30, 31, 30]
        expected = sum(v * d for v, d in zip(values, days)) / sum(days)
        _, vals = regrain_series(values, 'month', 'year', '2024-01', 'mean')
        assert abs(vals[0] - expected) < 1e-12

    def test_geometric_rate(self):
        from modeleon.core.time import regrain_series
        # 1% per month, 3 months -> compounded, not summed.
        _, vals = regrain_series([0.01, 0.01, 0.01], 'month', 'quarter',
                                 '2024-01', 'geometric')
        assert abs(vals[0] - (1.01**3 - 1)) < 1e-12


class TestSemantics52:
    """§5.2 corrections: quotients recompute as ratio-of-sums, mo.ratio
    resolves siblings, frozen splits formula from values."""

    def _windowed(self):
        m = mo.Model('m')
        m.default_start = '2024-01'
        m.default_grain = 'month'
        return m

    def test_quotient_recomputes_as_ratio_of_sums(self):
        # margin = gp / rev at quarter = Σgp / Σrev — never mean of ratios.
        m = self._windowed()
        m.gp = mo.Variable([50.0, 60.0, 70.0], regrain=mo.up('sum'))
        m.rev = mo.Variable([100.0, 150.0, 200.0], regrain=mo.up('sum'))
        m.margin = m.gp / m.rev
        q = m.at('quarter')
        assert abs(q.margin._value[0] - (180.0 / 450.0)) < 1e-12

    def test_mo_ratio_resolves_siblings(self):
        m = self._windowed()
        m.rev = mo.Variable([100.0, 150.0, 200.0], regrain=mo.up('sum'))
        m.users = mo.Variable([10.0, 20.0, 20.0], regrain=mo.up('last'))
        m.arpu = mo.Variable([10.0, 7.5, 10.0], regrain=mo.ratio('rev', 'users'))
        q = m.at('quarter')
        assert abs(q.arpu._value[0] - (450.0 / 50.0)) < 1e-12

    def test_mo_ratio_resolves_across_containers_by_path(self):
        """ARPU divides revenue (P&L) by users (assumptions) — ratio
        sources are routinely NOT siblings, so a dotted path from an
        enclosing container must resolve."""
        m = mo.Model('m')
        m.default_start = '2024-01'
        m.default_grain = 'month'
        m.assumptions = mo.MultiVariable('Assumptions')
        m.assumptions.users = mo.Variable([10.0, 20.0, 20.0], regrain=mo.up('last'))
        m.pnl = mo.MultiVariable('P&L')
        m.pnl.revenue = mo.Variable([100.0, 150.0, 200.0], regrain=mo.up('sum'))
        m.pnl.arpu = mo.Variable(
            [10.0, 7.5, 10.0], regrain=mo.ratio('revenue', 'assumptions.users')
        )
        q = m.at('quarter')
        assert abs(q.pnl.arpu._value[0] - (450.0 / 50.0)) < 1e-12

    def test_mo_ratio_unresolvable_source_teaches(self):
        m = self._windowed()
        m.rev = mo.Variable([1.0, 2.0, 3.0], regrain=mo.up('sum'))
        m.x = mo.Variable([1.0, 2.0, 3.0], regrain=mo.ratio('rev', 'nope'))
        with pytest.raises(ValueError, match="doesn't resolve"):
            m.at('quarter')

    def test_mo_ratio_zero_denominator_is_error_cell(self):
        m = self._windowed()
        m.rev = mo.Variable([0.0, 0.0, 0.0], regrain=mo.up('sum'))
        m.users = mo.Variable([0.0, 0.0, 0.0], regrain=mo.up('sum'))
        m.arpu = mo.Variable([0.0, 0.0, 0.0], regrain=mo.ratio('rev', 'users'))
        q = m.at('quarter')
        assert q.arpu._value[0] == '#DIV/0!'

    def test_frozen_with_value_rule_projects(self):
        # A recurrence's formula is frozen; its VALUES re-grain as a stock —
        # quarter-end cash is the last monthly close.
        cash = mo.Variable([100, 120, 150, 170, 200, 260], start='2024-01',
                           grain='month', regrain=mo.frozen(values='last'))
        assert cash.at('quarter')._value == [150, 260]

    def test_bare_frozen_still_refuses(self):
        spine = mo.Variable([1, 2, 3], start='2024-01', grain='month',
                            regrain=mo.frozen())
        with pytest.raises(ValueError, match='value rule'):
            spine.at('quarter')

    def test_frozen_values_rejects_unknown_recipe(self):
        with pytest.raises(ValueError):
            mo.frozen(values='blend')


class TestWindowConstructor:
    """The ambient window as constructor kwargs — the blessed authoring shape:
    ``mo.Model('m', default_grain='month', default_start='2026-01',
    default_periods=24)``."""

    def test_model_kwargs_set_window(self):
        m = mo.Model('m', default_grain='month', default_start='2026-01',
                     default_periods=24)
        assert (m.default_grain, m.default_start, m.default_periods) == \
            ('month', '2026-01', 24)

    def test_list_child_inherits_native_location(self):
        m = mo.Model('m', default_grain='month', default_start='2026-01',
                     default_periods=3)
        m.revenue = mo.Variable([1, 2, 3], regrain=mo.up('sum'))
        loc = m.revenue.time
        assert loc is not None and (loc.start, loc.grain) == ('2026-01', 'month')
        assert m.revenue.at('quarter')._value == [6]

    def test_nested_mv_window_kwargs(self):
        m = mo.MultiVariable('Ops', default_grain='quarter',
                             default_start='2026-Q1')
        assert (m.default_grain, m.default_start) == ('quarter', '2026-Q1')

    def test_start_without_grain_raises(self):
        with pytest.raises(ValueError, match='default_grain'):
            mo.Model('m', default_start='2026-01')

    def test_unknown_grain_raises(self):
        with pytest.raises(ValueError, match='default_grain'):
            mo.Model('m', default_grain='fortnight')

    def test_bad_start_format_raises(self):
        with pytest.raises(ValueError):
            mo.Model('m', default_grain='month', default_start='Jan 2026')

    def test_bad_periods_raise(self):
        with pytest.raises(TypeError):
            mo.Model('m', default_grain='month', default_periods=True)
        with pytest.raises(ValueError):
            mo.Model('m', default_grain='month', default_periods=0)


class TestWeightedRecipes:
    """Weight-aware apply_recipe: day-count weights, domain guards, and
    error-token absorption (one bad bucket = one error cell, never a crash)."""

    def test_period_weights_calendar_days(self):
        from modeleon.core.time import period_weights
        assert period_weights('month', '2024-01', 3) == [31, 29, 31]  # leap Feb
        assert period_weights('month', '2023-01', 2) == [31, 28]
        assert period_weights('year', '2024', 2) == [366, 365]
        assert period_weights('day', '2024-02-28', 3) == [1, 1, 1]
        assert period_weights('quarter', '2024-Q1', 2) == [91, 91]

    def test_weighted_mean_month_to_quarter(self):
        from modeleon.core.time import regrain_series
        _, vals = regrain_series([1.0, 2.0, 3.0], 'month', 'quarter',
                                 '2024-01', 'mean')
        expected = (1.0 * 31 + 2.0 * 29 + 3.0 * 31) / (31 + 29 + 31)
        assert abs(vals[0] - expected) < 1e-12

    def test_mean_without_weights_is_uniform(self):
        from modeleon.core.regrain import apply_recipe
        assert apply_recipe([1.0, 2.0, 3.0], 'mean') == 2.0

    def test_sum_ignores_weights(self):
        from modeleon.core.regrain import apply_recipe
        assert apply_recipe([1, 2, 3], 'sum', [31, 29, 31]) == 6

    def test_geometric_minus_100pct_is_num_error(self):
        from modeleon.core.regrain import apply_recipe
        assert apply_recipe([0.01, -1.0, 0.02], 'geometric') == '#NUM!'
        assert apply_recipe([0.01, -1.5], 'geometric') == '#NUM!'

    def test_error_token_absorbs_through_any_recipe(self):
        from modeleon.core.regrain import apply_recipe
        assert apply_recipe([1, '#DIV/0!', 3], 'sum') == '#DIV/0!'
        assert apply_recipe(['#VALUE!', 2], 'last', [31, 29]) == '#VALUE!'

    def test_zero_total_weight_mean_is_div0(self):
        from modeleon.core.regrain import apply_recipe
        assert apply_recipe([1.0, 2.0], 'mean', [0, 0]) == '#DIV/0!'

    def test_weights_length_mismatch_raises(self):
        from modeleon.core.regrain import apply_recipe
        with pytest.raises(ValueError):
            apply_recipe([1, 2, 3], 'mean', [31, 29])


class TestProjectionAt:
    """`.at(grain)` re-grains by applying each line's own rule to its own
    computed value — on a Variable, and across a whole model."""

    def test_variable_at_sum(self):
        rev = mo.Variable([1, 2, 3, 4, 5, 6], start='2024-01', grain='month',
                          regrain=mo.up('sum'))
        q = rev.at('quarter')
        assert q._value == [6, 15]
        assert q.time.grain == 'quarter'
        assert q.time.start == '2024-Q1'

    def test_variable_at_last(self):
        seats = mo.Variable([100, 110, 121, 130, 142, 150], start='2024-01',
                            grain='month', regrain=mo.up('last'))
        assert seats.at('quarter')._value == [121, 150]

    def test_agnostic_at_is_broadcast_copy(self):
        assert mo.Variable(0.2).at('quarter')._value == 0.2

    def test_identity_grain_is_noop(self):
        v = mo.Variable([1, 2, 3], start='2024-01', grain='month', regrain=mo.up('sum'))
        assert v.at('month')._value == [1, 2, 3]

    def test_no_rule_raises(self):
        v = mo.Variable([1, 2, 3], start='2024-01', grain='month')
        with pytest.raises(ValueError):
            v.at('quarter')

    def test_frozen_refuses(self):
        v = mo.Variable([1, 2, 3], start='2024-01', grain='month', regrain=mo.frozen())
        with pytest.raises(ValueError):
            v.at('quarter')

    def test_model_at_regrains_inputs(self):
        m = mo.Model('m')
        m.default_start = '2024-01'
        m.default_grain = 'month'
        m.seats = mo.Variable([100, 110, 121, 130, 142, 150], regrain=mo.up('last'))
        m.opex = mo.Variable([10, 20, 30, 40, 50, 60], regrain=mo.up('sum'))
        q = m.at('quarter')
        assert q.seats._value == [121, 150]
        assert q.opex._value == [60, 150]      # 10+20+30, 40+50+60

    def test_scalar_formula_copies_under_projection(self):
        # Rate conversions ((1+annual)**(1/12)-1) and horizon
        # reductions are grain-independent scalars — they copy, never
        # recompute (recompute choked on '**' and painted them red).
        m = mo.Model('m')
        m.default_start = '2024-01'
        m.default_grain = 'month'
        m.annual = mo.Variable(0.05)
        m.monthly = mo.Variable((1 + m.annual) ** (1 / 12) - 1)
        m.flows = mo.Variable([10.0, 20.0, 30.0, 40.0, 50.0, 60.0],
                              regrain=mo.up('sum'))
        m.total = mo.SUM(m.flows)
        q = m.at('quarter')
        assert q.monthly._value == m.monthly._value
        assert q.total._value == 210.0
        assert getattr(q.monthly, '_regrain_error', None) is None

    def test_absorbed_row_keeps_source_code(self):
        # A red #VALUE! cell must keep its identity — display name AND
        # the source text (an empty formula bar reads as "code
        # disappeared").
        m = mo.Model('m')
        m.default_start = '2024-01'
        m.default_grain = 'month'
        m.price = mo.recurrence(start=100.0, formula='{prev} * 1.01',
                                periods=6, display_name='Price')
        from modeleon.core.projection import project_model
        q = project_model(m, 'quarter', on_error='absorb')
        assert q.price._value == '#VALUE!'
        assert q.price._display_name == 'Price'
        assert 'recurrence' in (q.price._source_code or '')
        assert 'grain-frozen' in q.price._regrain_error

    def test_cumsum_check_row_regrains(self):
        # debt − (loan − cumsum(principal)) — the doctrine's debt
        # check. cumsum re-grains as cumsum of the summed input:
        # exact at quarter ends.
        m = mo.Model('m')
        m.default_start = '2024-01'
        m.default_grain = 'month'
        m.principal = mo.Variable([10.0] * 6, regrain=mo.up('sum'))
        m.cum = mo.cumsum(m.principal)
        q = m.at('quarter')
        assert q.cum._value == [30.0, 60.0]

    def test_if_tax_line_flow_regrains(self):
        # mo.IF is non-linear — it evaluates at NATIVE grain and sums
        # per bucket (flow default), so quarterly tax ties to the
        # monthly cash tax even when a loss month sits inside a
        # profitable quarter.
        m = mo.Model('m')
        m.default_start = '2024-01'
        m.default_grain = 'month'
        m.ebit = mo.Variable([-10.0, 20.0, 30.0, 40.0, 50.0, 60.0],
                             regrain=mo.up('sum'))
        m.tax = mo.IF(m.ebit > 0, m.ebit * 0.2, 0.0)
        q = m.at('quarter')
        assert q.tax._value == [pytest.approx(10.0), pytest.approx(30.0)]

    def test_model_at_same_grain_is_identity(self):
        # The grid's grain dropdown can ask for the NATIVE grain —
        # identity projection, not a "coarsens only" error (the window
        # must carry over verbatim, values unchanged).
        m = mo.Model('m')
        m.default_start = '2024-01'
        m.default_grain = 'month'
        m.default_periods = 6
        m.opex = mo.Variable([10, 20, 30, 40, 50, 60], regrain=mo.up('sum'))
        same = m.at('month')
        assert same.opex._value == [10, 20, 30, 40, 50, 60]
        assert same.default_grain == 'month'
        assert same.default_start == '2024-01'
        assert same.default_periods == 6

    def test_lag_row_recomputes_at_coarser_grain(self):
        # ΔNWC = nwc − mo.lag(nwc) projects by recomputing the shift in
        # the TARGET grain: delta_q = nwc_q − nwc_{q−1}.
        m = mo.Model('m')
        m.default_start = '2024-01'
        m.default_grain = 'month'
        m.nwc = mo.Variable(
            [100.0, 130.0, 120.0, 140.0, 150.0, 145.0],
            regrain=mo.up('last'),
        )
        m.delta = (m.nwc - mo.lag(m.nwc)).set_display_name('Change in NWC')
        q = m.at('quarter')
        assert q.nwc._value == [120.0, 145.0]
        assert q.delta._value == [120.0, 25.0]

    def test_lag_with_variable_fill_recomputes_at_coarser_grain(self):
        # opening = mo.lag(closing, fill=opening_0): opening_q =
        # closing_{q−1}, first quarter reads the fill assumption.
        m = mo.Model('m')
        m.default_start = '2024-01'
        m.default_grain = 'month'
        m.opening0 = mo.Variable(75.0)
        m.closing = mo.Variable(
            [90.0, 80.0, 70.0, 60.0, 50.0, 40.0], regrain=mo.up('last'),
        )
        m.opening = mo.lag(m.closing, fill=m.opening0)
        q = m.at('quarter')
        assert q.closing._value == [70.0, 40.0]
        assert q.opening._value == [75.0, 70.0]

    def test_model_at_with_formula_rule(self):
        m = mo.Model('m')
        m.default_start = '2024-01'
        m.default_grain = 'month'
        m.a = mo.Variable([1, 2, 3, 4, 5, 6], regrain=mo.up('sum'))
        m.total = (m.a + m.a).set_regrain(mo.up('sum'))    # formula carries a rule
        assert m.at('quarter').total._value == [12, 30]    # (2+4+6), (8+10+12)

    def test_additive_formula_without_rule_recomputes(self):
        # gross = revenue - cost is additive -> no rule needed; it recomputes
        # over the re-grained inputs (== sum of monthly gross, the same number).
        m = mo.Model('m')
        m.default_start = '2024-01'
        m.default_grain = 'month'
        m.rev = mo.Variable([10, 11, 12, 13, 14, 15], regrain=mo.up('sum'))
        m.cost = mo.Variable([4, 4, 5, 5, 6, 6], regrain=mo.up('sum'))
        m.gross = (m.rev - m.cost).set_display_name('Gross')   # NO rule
        q = m.at('quarter')
        assert q.rev._value == [33, 42]
        assert q.cost._value == [13, 17]
        assert q.gross._value == [20, 25]                      # rev_q - cost_q

    def test_multiplicative_flow_evaluates_then_regrains(self):
        # revenue = price * seats can't recompute from aggregates — but its
        # OUTPUT is a flow, so the §5.2 default evaluates at native grain and
        # sums the result per bucket: quarterly revenue = Σ monthly (p·s).
        # (Previously a hard error; the correct number was always available.)
        m = mo.Model('m')
        m.default_start = '2024-01'
        m.default_grain = 'month'
        m.price = mo.Variable([10, 11, 12, 13, 14, 15], regrain=mo.up('mean'))
        m.seats = mo.Variable([100, 110, 120, 130, 140, 150], regrain=mo.up('last'))
        m.rev = (m.price * m.seats)                            # mult flow, NO rule
        q = m.at('quarter')
        native = [p * s for p, s in zip(m.price._value, m.seats._value)]
        assert q.rev._value == [sum(native[:3]), sum(native[3:])]
        # NOT the wrong aggP*aggS shape:
        assert q.rev._value[0] != q.price._value[0] * q.seats._value[0]

    def test_compound_formula_intermediate_recomputes(self):
        # `a - b - c` builds a floating intermediate `(a - b)` (no owner, so its
        # .time is None); the projection must still recompute it, not leave it
        # un-re-grained as if it were agnostic.
        m = mo.Model('m')
        m.default_start = '2024-01'
        m.default_grain = 'month'
        m.a = mo.Variable([1, 2, 3, 4, 5, 6], regrain=mo.up('sum'))
        m.b = mo.Variable([1, 1, 1, 1, 1, 1], regrain=mo.up('sum'))
        m.c = mo.Variable([2, 2, 2, 2, 2, 2], regrain=mo.up('sum'))
        m.total = (m.a - m.b - m.c).set_display_name('Total')   # compound, no rule
        # a_q=[6,15], b_q=[3,3], c_q=[6,6] -> [6-3-6, 15-3-6]
        assert m.at('quarter').total._value == [-3, 6]

    def test_scalar_mult_formula_recomputes(self):
        # revenue * (1 - tax) where tax is agnostic is LINEAR -> recompute is safe.
        m = mo.Model('m')
        m.default_start = '2024-01'
        m.default_grain = 'month'
        m.rev = mo.Variable([10, 20, 30, 40, 50, 60], regrain=mo.up('sum'))
        m.tax = mo.Variable(0.1)                               # agnostic scalar
        m.net = (m.rev * (1 - m.tax)).set_display_name('Net')  # NO rule (linear)
        assert m.at('quarter').net._value == [54.0, 135.0]     # 0.9 * [60, 150]


class TestAliasProjection:
    """``report.revenue = pnl.revenue`` — surfacing one block's row in
    another. Adoption into a second parent renders as ``revenue.copy()``,
    which is a formula with nothing to compute; the grain switch must
    follow the source instead of refusing the line (Arrual P&L,
    2026-08-04)."""

    def test_alias_projects_like_its_source(self) -> None:
        from modeleon.core.projection import project_model

        m = mo.Model('m', default_grain='month', default_start='2025-01',
                     default_periods=6)
        with m:
            m.pnl = mo.MultiVariable()
            with m.pnl as pnl:
                pnl.revenue = mo.Variable(
                    [1.0, 2.0, 3.0, 4.0, 5.0, 6.0], regrain=mo.up('sum')
                )
            m.report = mo.MultiVariable()
            with m.report as report:
                report.revenue = m.pnl.revenue

        q = project_model(m, 'quarter')
        assert q.pnl.revenue._value == [6.0, 15.0]
        assert q.report.revenue._value == [6.0, 15.0]

    def test_alias_of_a_formula_row_projects_too(self) -> None:
        from modeleon.core.projection import project_model

        m = mo.Model('m', default_grain='month', default_start='2025-01',
                     default_periods=6)
        with m:
            m.a = mo.Variable([1.0] * 6, regrain=mo.up('sum'))
            m.b = mo.Variable([2.0] * 6, regrain=mo.up('sum'))
            m.total = m.a + m.b
            m.board = mo.MultiVariable()
            with m.board as board:
                board.total = m.total

        q = project_model(m, 'quarter')
        assert q.board.total._value == [9.0, 9.0]
