# SPDX-License-Identifier: Apache-2.0
"""The track value layer.

A track-axised Variable stores role → time-series tracks. One formula
broadcasts across coordinates; the mismatch law is loud; time-coupled
functions lift per track; slices are first-class AST nodes.
"""

from __future__ import annotations

import re

import pytest

import modeleon as mo
from modeleon.core.expr import Restrict
from modeleon.core.tracks import TrackValues


def _model(periods: int = 4) -> "mo.Model":
    return mo.Model('m', default_grain='month', default_start='2025-01',
                    default_periods=periods)


class TestConstruction:
    def test_tracked_variable(self):
        m = _model()
        with m:
            m.x = mo.Variable(tracks={'plan': [1.0] * 4, 'actual': [2.0] * 4})
        assert isinstance(m.x._value, TrackValues)
        assert set(m.x._value.roles) == {'plan', 'actual'}
        assert m.x.var_type == 'list'

    def test_time_resolves_from_ambient(self):
        m = _model()
        with m:
            m.x = mo.Variable(tracks={'plan': [1.0] * 4})
        assert m.x.time is not None and m.x.time.grain == 'month'

    def test_extent_law_applies_per_track(self):
        m = _model(24)
        with pytest.raises(ValueError, match="24-period window"):
            with m:
                m.x = mo.Variable(tracks={'plan': [1.0, 2.0]})

    def test_ragged_tracks_rejected(self):
        with pytest.raises(ValueError, match="one time length"):
            TrackValues({'plan': [1.0, 2.0], 'actual': [1.0]})

    def test_nested_tracks_rejected(self):
        with pytest.raises(TypeError, match="second axis"):
            TrackValues({'plan': {'nested': [1.0]}})

    def test_exclusive_with_value_and_indexed_by(self):
        with pytest.raises(ValueError, match="exclusive"):
            mo.Variable([1.0], tracks={'plan': [1.0]})
        axis = mo.Variable(['a', 'b'])
        with pytest.raises(ValueError, match="axis budget"):
            mo.Variable(tracks={'plan': [1.0, 2.0]}, indexed_by=[axis])


class TestBroadcast:
    def _rev(self, m):
        with m:
            m.revenue = mo.Variable(tracks={
                'plan': [100.0, 110.0, 120.0, 130.0],
                'actual': [95.0, 118.0, 0.0, 0.0],
            })
        return m.revenue

    def test_tracked_minus_plain_broadcasts_into_every_role(self):
        m = _model()
        rev = self._rev(m)
        with m:
            m.costs = mo.Variable([50.0] * 4)
            m.margin = rev - m.costs
        v = m.margin._value
        assert isinstance(v, TrackValues)
        assert v['plan'] == [50.0, 60.0, 70.0, 80.0]
        assert v['actual'] == [45.0, 68.0, -50.0, -50.0]

    def test_tracked_times_scalar(self):
        m = _model()
        rev = self._rev(m)
        with m:
            m.doubled = rev * 2
        assert m.doubled._value['plan'] == [200.0, 220.0, 240.0, 260.0]

    def test_same_roles_zip_per_track(self):
        m = _model()
        with m:
            m.a = mo.Variable(tracks={'plan': [1.0] * 4, 'actual': [2.0] * 4})
            m.b = mo.Variable(tracks={'plan': [10.0] * 4, 'actual': [20.0] * 4})
            m.c = m.a + m.b
        assert m.c._value['plan'] == [11.0] * 4
        assert m.c._value['actual'] == [22.0] * 4

    def test_mismatched_role_sets_die_loudly(self):
        m = _model()
        with m:
            m.a = mo.Variable(tracks={'plan': [1.0] * 4, 'actual': [2.0] * 4})
            m.b = mo.Variable(tracks={'plan': [1.0] * 4, 'base': [9.0] * 4})
        with pytest.raises(ValueError, match="never silently intersects"):
            _ = m.a + m.b

    def test_comparison_lifts(self):
        m = _model()
        rev = self._rev(m)
        flag = rev > 100
        assert flag._value['plan'] == [False, True, True, True]


class TestSlice:
    def test_at_basis_returns_the_track(self):
        m = _model()
        with m:
            m.x = mo.Variable(tracks={'plan': [1.0, 2.0, 3.0, 4.0],
                                      'actual': [9.0] * 4})
        sliced = m.x.at(track='plan')
        assert sliced._value == [1.0, 2.0, 3.0, 4.0]
        assert isinstance(sliced._expr, Restrict)
        assert sliced._expr.label == 'plan'

    def test_variance_is_one_formula(self):
        m = _model()
        with m:
            m.x = mo.Variable(tracks={'plan': [100.0] * 4,
                                      'actual': [95.0, 118.0, 0.0, 0.0]})
            m.variance = m.x.at(track='actual') - m.x.at(track='plan')
        assert m.variance._value == [-5.0, 18.0, -100.0, -100.0]

    def test_unknown_role_dies_with_the_declared_list(self):
        x = mo.Variable(tracks={'plan': [1.0, 2.0]})
        with pytest.raises(ValueError, match="plan"):
            x.at(track='actual')

    def test_at_basis_on_plain_variable_dies(self):
        with pytest.raises(ValueError, match="no track coordinates"):
            mo.Variable([1.0, 2.0]).at(track='plan')

    def test_at_needs_something(self):
        with pytest.raises(TypeError, match="grain"):
            mo.Variable(tracks={'plan': [1.0]}).at()


class TestLifting:
    def _rev(self, m):
        with m:
            m.revenue = mo.Variable(tracks={
                'plan': [100.0, 110.0, 120.0, 130.0],
                'actual': [95.0, 118.0, 0.0, 0.0],
            })
        return m.revenue

    def test_lag_shifts_each_track(self):
        m = _model()
        rev = self._rev(m)
        lagged = mo.lag(rev)
        assert lagged._value['plan'] == [0.0, 100.0, 110.0, 120.0]
        assert lagged._value['actual'] == [0.0, 95.0, 118.0, 0.0]

    def test_cumsum_accumulates_each_track(self):
        m = _model()
        rev = self._rev(m)
        acc = mo.cumsum(rev)
        assert acc._value['plan'] == [100.0, 210.0, 330.0, 460.0]
        assert acc._value['actual'] == [95.0, 213.0, 213.0, 213.0]

    def test_sum_reduces_each_track(self):
        m = _model()
        rev = self._rev(m)
        total = mo.SUM(rev)
        assert total._value['plan'] == 460.0
        assert total._value['actual'] == 213.0

    def test_if_lifts_across_roles(self):
        m = _model()
        rev = self._rev(m)
        flag = mo.IF(rev > 100, 1.0, 0.0)
        assert flag._value['plan'] == [0.0, 1.0, 1.0, 1.0]
        assert flag._value['actual'] == [0.0, 1.0, 0.0, 0.0]

    def test_recurrence_with_tracked_driver_lifts(self):
        # Rank-lifted: one chain per track — see TestRecurrenceLift.
        m = _model()
        rev = self._rev(m)
        r = mo.recurrence(start=0.0, formula="{prev} + {r}",
                          variables={"r": rev}, periods=4)
        assert r._value['plan'] == [0.0, 110.0, 230.0, 360.0]
        assert r._value['actual'] == [0.0, 118.0, 118.0, 118.0]


class TestGates:
    def _tracked(self):
        m = _model()
        with m:
            m.x = mo.Variable(tracks={'plan': [1.0] * 4, 'actual': [2.0] * 4},
                              display_name='X')
        return m.x

    def test_whole_cube_grain_projects_per_track(self):
        """at(grain) on a tracked variable re-grains EACH track by the
        line's rule — the axis passes through (was a refusal)."""
        m = _model(12)
        with m:
            m.x = mo.Variable(
                tracks={'plan': [1.0] * 12, 'actual': [2.0] * 12},
                regrain=mo.up('sum'),
            )
        q = m.x.at('quarter')
        assert q._value['plan'] == [3.0] * 4
        assert q._value['actual'] == [6.0] * 4

    def test_whole_cube_without_rule_still_teaches(self):
        with pytest.raises(ValueError, match="re-grain rule"):
            self._tracked().at('quarter')

    def test_slice_then_grain_works(self):
        m = _model(12)
        with m:
            m.x = mo.Variable(
                tracks={'plan': [1.0] * 12, 'actual': [2.0] * 12},
                regrain=mo.up('sum'),
            )
        q = m.x.at('quarter', track='plan')
        assert q._value == [3.0, 3.0, 3.0, 3.0]

    def test_emission_expands_tracks_into_rows(self):
        """The xlsx writer expands a tracked line — the display default
        keeps the line's name, other tracks follow as labeled value
        rows (was a refusal)."""
        import os
        import tempfile
        from openpyxl import load_workbook
        m = _model()
        with m:
            m.x = mo.Variable(tracks={'plan': [1.0] * 4, 'actual': [2.0] * 4},
                              display_name='Ex')
        p = os.path.join(tempfile.mkdtemp(), 't.xlsx')
        m.to_excel(p)
        labels = [r[0] for r in load_workbook(p).worksheets[0]
                  .iter_rows(min_row=2, max_row=3, max_col=1,
                             values_only=True)]
        assert labels[0] == 'Ex'             # default row keeps the name
        assert 'actual' in (labels[1] or '')  # the other track, labeled

    def test_model_projection_projects_per_track(self):
        from modeleon.core.projection import project_variable
        m = _model(12)
        with m:
            m.x = mo.Variable(
                tracks={'plan': [1.0] * 12, 'actual': [2.0] * 12},
                regrain=mo.up('sum'),
            )
        p = project_variable(m.x, 'quarter', {})
        assert p._value['plan'] == [3.0] * 4
        assert p._value['actual'] == [6.0] * 4


class TestRecurrenceLift:
    """Recurrence over tracks = N independent chains.

    The plan balance rolls over plan flows, the actual balance over
    actual flows. (The re-anchored chain of a blended model's ``live``
    track is covered by TestBlendLiveTrack, not this class.)
    """

    def test_tracked_driver_rolls_per_track(self):
        m = _model()
        with m:
            m.flow = mo.Variable(tracks={
                'plan': [100.0, 100.0, 100.0, 100.0],
                'actual': [95.0, 120.0, 0.0, 0.0],
            })
            m.balance = mo.recurrence(
                start=1000.0, formula="{prev} + {f}",
                variables={"f": m.flow},
            )
        assert m.balance._value['plan'] == [1000.0, 1100.0, 1200.0, 1300.0]
        assert m.balance._value['actual'] == [1000.0, 1120.0, 1120.0, 1120.0]

    def test_tracked_start_spawns_chains(self):
        m = _model(3)
        with m:
            m.opening = mo.Variable(tracks={'plan': 10.0, 'actual': 20.0})
            m.chain = mo.recurrence(start=m.opening, formula="{prev} * 2",
                                    periods=3)
        assert m.chain._value['plan'] == [10.0, 20.0, 40.0]
        assert m.chain._value['actual'] == [20.0, 40.0, 80.0]

    def test_mismatched_role_sets_die(self):
        m = _model()
        with m:
            m.a = mo.Variable(tracks={'plan': [1.0] * 4, 'actual': [2.0] * 4})
            m.b = mo.Variable(tracks={'plan': [1.0] * 4, 'base': [9.0] * 4})
        with pytest.raises(ValueError, match="differ"):
            mo.recurrence(start=0.0, formula="{prev} + {a} + {b}",
                          variables={"a": m.a, "b": m.b})

    def test_lambda_mode_lifts_too(self):
        m = _model(3)
        with m:
            m.opening = mo.Variable(tracks={'plan': 1.0, 'actual': 2.0})
            m.r = mo.recurrence(start=m.opening,
                                formula=lambda prev, t: prev + 1, periods=3)
        assert m.r._value['plan'] == [1.0, 2.0, 3.0]
        assert m.r._value['actual'] == [2.0, 3.0, 4.0]


class TestCyrillicPlaceholders:
    def test_cyrillic_template_names_substitute(self):
        # The placeholder regex used to be ASCII-only: a Cyrillic "{п}"
        # was not substituted and parsed as a set literal. Names in a
        # model can be in any language — the non-Latin name below is
        # deliberate.
        r = mo.recurrence(start=0.0, formula="{prev} + {доход}",
                          variables={"доход": mo.Variable([1.0, 2.0, 3.0])})
        assert r._value == [0.0, 2.0, 5.0]


class TestDateAlignment:
    """The date-aligned zip: operands with different windows
    align by CALENDAR; outside its own extent each side follows its
    declared extend rule — never a positional zip, never a guess."""

    def _m(self):
        return mo.Model('m', default_grain='month',
                        default_start='2025-01', default_periods=6)

    def test_flow_extends_with_zeros(self):
        m = self._m()
        with m:
            m.revenue = mo.Variable([100.0] * 6)
            m.bonus = mo.Variable([10.0] * 3, start='2025-04',
                                  grain='month', extend=mo.zero())
            m.total = m.revenue + m.bonus
        assert m.total._value == [100.0, 100.0, 100.0, 110.0, 110.0, 110.0]

    def test_rate_holds_forward_and_backward(self):
        m = self._m()
        with m:
            m.base = mo.Variable([100.0] * 6)
            m.rate = mo.Variable([0.10, 0.12], start='2025-05',
                                 grain='month', extend=mo.hold())
            m.fee = m.base * m.rate
        # Backward with the first value, forward with the last.
        assert m.fee._value == [10.0, 10.0, 10.0, 10.0, 10.0, 12.0]

    def test_result_is_located_at_the_union(self):
        m = self._m()
        with m:
            m.a = mo.Variable([1.0] * 6)
            m.b = mo.Variable([1.0] * 2, start='2025-05', grain='month',
                              extend=mo.zero())
            m.c = m.a + m.b
        assert m.c.time.start == '2025-01'
        assert len(m.c._value) == 6

    def test_undeclared_continuation_teaches(self):
        m = self._m()
        with pytest.raises(ValueError, match="Say what continues it"):
            with m:
                m.a = mo.Variable([1.0] * 6)
                m.b = mo.Variable([5.0, 5.0], start='2025-03',
                                  grain='month')
                m.x = m.a + m.b

    def test_same_window_stays_positional(self):
        # The fast path is untouched — byte-for-byte as before.
        m = self._m()
        with m:
            m.a = mo.Variable([1.0] * 6)
            m.b = mo.Variable([2.0] * 6)
            m.c = m.a + m.b
        assert m.c._value == [3.0] * 6

    def test_grain_mismatch_refuses(self):
        a = mo.Variable([1.0] * 6, start='2025-01', grain='month')
        b = mo.Variable([1.0, 1.0], start='2025-01', grain='quarter')
        with pytest.raises(ValueError, match="different grains"):
            _ = a + b


class TestAggregateLifts:
    def test_max_min_average_reduce_per_track(self):
        m = _model()
        with m:
            m.x = mo.Variable(tracks={'plan': [1.0, 5.0, 3.0, 2.0],
                                      'actual': [9.0, 0.0, 0.0, 0.0]})
        assert dict(mo.MAX(m.x)._value.items()) == {'plan': 5.0, 'actual': 9.0}
        assert dict(mo.MIN(m.x)._value.items()) == {'plan': 1.0, 'actual': 0.0}
        assert dict(mo.AVERAGE(m.x)._value.items()) == {
            'plan': 2.75, 'actual': 2.25,
        }


# ─── mo.Tracks declaration + role-kwarg authoring ──────────────────


def _declared(periods: int = 6, **decl_labels) -> "mo.Model":
    labels = decl_labels or {'actual': 'Actual', 'plan': 'Budget'}
    return mo.Model('m', tracks=mo.Tracks(**labels),
                    default_grain='month', default_start='2025-01',
                    default_periods=periods)


class TestTracksDeclaration:
    def test_declaration_on_model(self):
        m = _declared()
        assert m.tracks.names == ('actual', 'plan')
        assert m.tracks.label('plan') == 'Budget'

    def test_names_are_user_content(self):
        """Any words, any language, any count — the engine attaches no
        semantics to a track name."""
        # The non-Latin name is deliberate: a track name may be in any
        # script.
        d = mo.Tracks('actual', 'budget_approved', 'бюджет')
        assert d.names == ('actual', 'budget_approved', 'бюджет')
        assert d.label('actual') == 'actual'   # positional: name is the label

    def test_labeled_spelling(self):
        d = mo.Tracks('actual', budget='Budget 2026')
        assert d.names == ('actual', 'budget')
        assert d.label('budget') == 'Budget 2026'

    def test_single_track_is_legal(self):
        m = mo.Model('m', tracks=mo.Tracks('plan'),
                     default_grain='month', default_start='2025-01',
                     default_periods=3)
        with m:
            m.x = mo.Variable(plan=[1, 2, 3])
        assert m.x._value['plan'] == [1, 2, 3]

    def test_empty_declaration_refused(self):
        with pytest.raises(TypeError, match="name them"):
            mo.Tracks()

    def test_declaration_must_be_tracks(self):
        with pytest.raises(TypeError, match="mo.Tracks"):
            mo.Model('t', tracks={'actual': 'a'})

    def test_labels_are_nonempty_strings(self):
        with pytest.raises(TypeError, match="non-empty"):
            mo.Tracks(actual='')


class TestRoleKwargAuthoring:
    def test_role_lists_materialize(self):
        m = _declared()
        with m:
            m.revenue = mo.Variable(actual=[1] * 6, plan=[10] * 6)
        v = m.revenue._value
        assert isinstance(v, TrackValues)
        assert v.roles == ('actual', 'plan')
        assert v['plan'] == [10] * 6
        assert m.revenue.var_type == 'list'

    def test_arithmetic_broadcasts(self):
        m = _declared()
        with m:
            m.a = mo.Variable(actual=[1] * 6, plan=[10] * 6)
            m.b = m.a * 2
        assert m.b._value['actual'] == [2] * 6
        assert m.b._value['plan'] == [20] * 6

    def test_pin_spelling_shared_formula(self):
        """Positional shared expression + actual override."""
        m = _declared()
        with m:
            m.price = mo.Variable(10)
            m.units = mo.Variable([1, 1, 2, 2, 3, 3])
            m.income = mo.Variable(m.price * m.units,
                                   actual=[9, 11, 22, 18, 33, 27])
        d = m.income._value
        assert d['plan'] == [10, 10, 20, 20, 30, 30]
        assert d['actual'] == [9, 11, 22, 18, 33, 27]
        # the shared expression stays the canonical formula
        assert m.income._expr is not None

    def test_coordinate_aligned_pull(self):
        """A tracked operand contributes its own same-role track."""
        m = _declared()
        with m:
            m.a = mo.Variable(actual=[1.0] * 6, plan=[10.0] * 6)
            m.b = mo.Variable(plan=m.a * 1.1, actual=m.a * 1.0)
        assert abs(m.b._value['plan'][0] - 11.0) < 1e-9
        assert m.b._value['actual'][0] == 1.0

    def test_variable_operand_is_dep_edge(self):
        m = _declared()
        with m:
            m.feed = mo.Variable([5] * 6)
            m.linked = mo.Variable(plan=[1] * 6, actual=m.feed)
        assert m.feed in m.linked._dependency_refs

    def test_slice_after_materialization(self):
        m = _declared()
        with m:
            m.x = mo.Variable(actual=[1] * 6, plan=[2] * 6)
        assert m.x.at(track='plan')._value == [2] * 6


class TestRoleKwargGates:
    def test_no_declaration_teaches(self):
        m = mo.Model('bare', default_grain='month',
                     default_start='2025-01', default_periods=3)
        with pytest.raises(ValueError, match="mo.Tracks"):
            with m:
                m.x = mo.Variable(plan=[1, 2, 3], actual=[1, 2, 3])

    def test_undeclared_role_refused(self):
        m = _declared()
        with pytest.raises(ValueError, match="forecast"):
            with m:
                m.y = mo.Variable(actual=[1] * 6, plan=[1] * 6,
                                  forecast=[1] * 6)

    def test_subset_materializes_incrementally(self):
        """A track subset is legal — incremental authoring accumulates
        tracks one statement at a time; the mismatch law polices any
        cross-track arithmetic against an incomplete set."""
        m = _declared()
        with m:
            m.z = mo.Variable(actual=[1.0] * 6)
        assert m.z._value.roles == ('actual',)
        m.z.plan = [2.0] * 6
        assert m.z._value['plan'] == [2.0] * 6

    def test_roles_exclusive_with_formula(self):
        with pytest.raises(ValueError, match="positionally"):
            mo.Variable(formula='a+b', actual=[1])

    def test_roles_exclusive_with_tracks(self):
        with pytest.raises(ValueError, match="two spellings"):
            mo.Variable(tracks={'plan': [1]}, actual=[1])

    def test_roles_exclusive_with_indexed_by(self):
        axis = mo.Variable(['a', 'b'])
        with pytest.raises(ValueError, match="axis budget"):
            mo.Variable(indexed_by=[axis], actual=[1, 2])

    def test_track_window_extent_still_applies(self):
        m = _declared(periods=6)
        with pytest.raises(ValueError, match="cover the"):
            with m:
                m.short = mo.Variable(actual=[1, 2, 3], plan=[1, 2, 3])


class TestRoleKwargLifecycle:
    """Lifecycle traps in adoption-time materialization — each of these
    reproduced a silent wrong value or a crash before its fix."""

    def test_two_list_operands_no_crash_both_edges(self):
        """Membership via ``in`` routed through the overloaded __eq__ —
        list-valued operands crashed, equal scalars dropped the edge."""
        m = _declared()
        with m:
            m.budget = mo.Variable([1.0] * 6)
            m.ledger = mo.Variable([2.0] * 6)
            m.revenue = mo.Variable(plan=m.budget, actual=m.ledger)
        refs = m.revenue._dependency_refs
        assert any(r is m.budget for r in refs)
        assert any(r is m.ledger for r in refs)

    def test_equal_scalar_operands_keep_both_edges(self):
        m = _declared()
        with m:
            m.a = mo.Variable(5.0)
            m.b = mo.Variable(5.0)
            m.c = mo.Variable(plan=m.a, actual=m.b)
        refs = m.c._dependency_refs
        assert any(r is m.a for r in refs) and any(r is m.b for r in refs)

    def test_tracked_shared_expression_honors_overrides(self):
        """A tracked shared expr made _value TrackValues at construction
        — the old type-based guard skipped materialization and silently
        discarded the overrides."""
        m = _declared()
        with m:
            m.price = mo.Variable(actual=[10.0] * 6, plan=[10.0] * 6)
            m.units = mo.Variable([1.0, 1.0, 2.0, 2.0, 3.0, 3.0])
            m.income = mo.Variable(m.price * m.units,
                                   actual=[9.0, 11.0, 22.0, 18.0, 33.0, 27.0])
        d = m.income._value
        assert d['actual'] == [9.0, 11.0, 22.0, 18.0, 33.0, 27.0]
        assert d['plan'] == [10.0, 10.0, 20.0, 20.0, 30.0, 30.0]

    def test_tracked_shared_variable_does_not_alias(self):
        m = _declared()
        with m:
            m.budget = mo.Variable(actual=[1.0] * 6, plan=[2.0] * 6)
            m.revision = mo.Variable(m.budget, actual=[9.0] * 6)
        assert m.revision._value['actual'] == [9.0] * 6
        assert m.revision._value['plan'] == [2.0] * 6
        assert m.revision._value is not m.budget._value

    def test_consumer_before_operand_now_works(self):
        """Pure track kwargs materialize AT CONSTRUCTION (from their own
        names), so a floating operand already carries its coordinates —
        the class-body pattern works and adoption order stops mattering
        for this shape."""
        m = _declared()
        a = mo.Variable(actual=[1.0] * 6, plan=[2.0] * 6)
        b = mo.Variable(plan=a, actual=[3.0] * 6)
        m.b = b
        assert m.b._value['plan'] == [2.0] * 6
        assert m.b._value['actual'] == [3.0] * 6

    def test_unadopted_schedule_operand_is_loud(self):
        m = _declared(periods=3)
        with pytest.raises(ValueError, match="no value yet"):
            with m:
                m.v = mo.Variable(actual=[99.0, 98.0, 97.0],
                                  plan=mo.schedule({'2025-01': 100}))

    def test_floating_subtree_with_local_window_defers(self):
        """The no-declaration error fired mid-construction on a floating
        subtree whose own window resolved; now it defers to the mount."""
        pnl = mo.MultiVariable('pnl', default_grain='month',
                               default_start='2025-01', default_periods=6)
        pnl.rev = mo.Variable(actual=[1.0] * 6, plan=[2.0] * 6)
        m = mo.Model('m', tracks=mo.Tracks(actual='Actual', plan='Budget'))
        m.pnl = pnl
        assert isinstance(m.pnl.rev._value, TrackValues)
        assert m.pnl.rev._value['plan'] == [2.0] * 6

    def test_factory_subtree_with_local_window_defers(self):
        pnl = mo.MultiVariable(
            'pnl2', default_grain='month', default_start='2025-01',
            default_periods=6,
            rev=mo.Variable(actual=[1.0] * 6, plan=[2.0] * 6))
        m = mo.Model('m', tracks=mo.Tracks(actual='Actual', plan='Budget'))
        m.pnl2 = pnl
        assert isinstance(m.pnl2.rev._value, TrackValues)

    def test_copy_carries_role_stash(self):
        """Clone-on-ownership-change stripped _role_kwargs — the copy
        landed valueless with no error."""
        tmpl = mo.MultiVariable('tmpl')
        tmpl.x = mo.Variable(actual=[1.0, 2.0, 3.0], plan=[4.0, 5.0, 6.0])
        m = _declared(periods=3)
        m.x = tmpl.x
        assert isinstance(m.x._value, TrackValues)
        assert m.x._value['plan'] == [4.0, 5.0, 6.0]

    def test_schedule_as_shared_value(self):
        """The fall-through's stated purpose was unreachable: the wrapper
        never inherited _schedule from the positional Variable."""
        m = _declared(periods=3)
        with m:
            m.rate = mo.Variable(
                mo.schedule({'2025-01': 100.0, '2025-03': 120.0}),
                actual=[99.0, 98.0, 97.0])
        r = m.rate._value
        assert r['actual'] == [99.0, 98.0, 97.0]
        assert r['plan'] == [100.0, 100.0, 120.0]

    def test_plain_schedule_wrapper_materializes(self):
        m = _declared(periods=3)
        with m:
            m.rate = mo.Variable(mo.schedule({'2025-01': 0.12}))
        assert m.rate._value == [0.12, 0.12, 0.12]

    def test_windowless_model_scalar_roles_materialize(self):
        """The window gate swallowed role materialization; scalar tracks
        need no window (time lives inside a track)."""
        m = mo.Model('m', tracks=mo.Tracks(actual='Actual', plan='Budget'))
        with m:
            m.kpi = mo.Variable(actual=100.0, plan=120.0)
        assert isinstance(m.kpi._value, TrackValues)
        assert m.kpi.at(track='plan')._value == 120.0

    def test_eager_compute_inside_a_class_body_works(self):
        """The class-body pattern: a tracked field is used by the
        very next line, long before the instance joins a model. Pure
        kwargs materialize at construction, so the arithmetic computes
        per track instead of dying on a pending operand."""
        pnl = mo.MultiVariable('pnl3')
        pnl.rev = mo.Variable(actual=[1.0] * 6, plan=[2.0] * 6)
        pnl.margin = pnl.rev * 0.5
        assert pnl.margin._value['actual'] == [0.5] * 6
        assert pnl.margin._value['plan'] == [1.0] * 6

    def test_pin_spelling_still_defers_and_guards(self):
        """The pin spelling needs the declared vocabulary to fill the
        tracks the shared expression owes — so it still defers, and
        eager arithmetic on it still teaches."""
        pnl = mo.MultiVariable('pnl4')
        pnl.rev = mo.Variable([1.0] * 6, actual=[2.0] * 6)
        with pytest.raises(ValueError, match="materialized"):
            pnl.margin = pnl.rev * 0.5

    def test_at_track_works_on_a_floating_variable(self):
        v = mo.Variable(actual=[1.0] * 3, plan=[2.0] * 3)
        assert v.at(track='actual')._value == [1.0] * 3

    def test_at_track_on_a_pending_pin_teaches(self):
        v = mo.Variable([9.0] * 3, actual=[1.0] * 3)
        with pytest.raises(ValueError, match="materialized"):
            v.at(track='actual')

    def test_crystallization_refire_is_idempotent(self):
        m = _declared()
        sub = mo.MultiVariable('sub')
        sub.v = mo.Variable(actual=[1.0] * 6, plan=[2.0] * 6)
        m.sub = sub
        assert m.sub.v._value['actual'] == [1.0] * 6

    def test_none_track_rejected_by_value_layer(self):
        with pytest.raises(TypeError, match="None"):
            TrackValues({'actual': [1.0], 'plan': None})


class TestDottedCoordinateRead:
    """``revenue.actual`` reads a coordinate — sugar over
    ``.at(track=...)``, real attributes always win (getattr fires only
    on lookup miss)."""

    def test_dotted_read_equals_at(self):
        m = _declared()
        with m:
            m.revenue = mo.Variable(actual=[1.0] * 6, plan=[2.0] * 6)
        assert m.revenue.actual._value == m.revenue.at(track='actual')._value

    def test_dotted_read_in_formula(self):
        m = _declared()
        with m:
            m.a = mo.Variable(actual=[10.0] * 6, plan=[7.0] * 6)
            m.d = m.a.actual - m.a.plan
        assert m.d._value == [3.0] * 6

    def test_real_attributes_win(self):
        m = _declared()
        with m:
            m.x = mo.Variable(actual=[1.0] * 6, plan=[1.0] * 6)
        assert not isinstance(m.x.value, mo.Variable)   # the property, not a track
        assert m.x.formula is None or isinstance(m.x.formula, str)

    def test_miss_teaches_track_names(self):
        m = _declared()
        with m:
            m.x = mo.Variable(actual=[1.0] * 6, plan=[1.0] * 6)
        with pytest.raises(AttributeError, match="actual, plan"):
            m.x.forecast

    def test_plain_variable_unchanged(self):
        v = mo.Variable([1.0, 2.0])
        with pytest.raises(AttributeError):
            v.nonexistent


class TestDottedTrackWrite:
    """MV-style incremental write — ``revenue.actual = [...]`` as
    its own statement, after the definition, without rewriting it."""

    def test_incremental_authoring(self):
        m = _declared()
        with m:
            m.revenue = mo.Variable()
        m.revenue.actual = [1.0] * 6
        m.revenue.plan = [2.0] * 6
        assert m.revenue._value['actual'] == [1.0] * 6
        assert m.revenue._value['plan'] == [2.0] * 6
        assert m.revenue.var_type == 'list'

    def test_pin_over_shared_formula(self):
        """The line has ONE formula; the actual arrives afterwards, bound
        by dot — the formula keeps computing only the un-pinned tracks
        (given beats derived)."""
        m = _declared()
        with m:
            m.price = mo.Variable([10.0] * 6)
            m.units = mo.Variable([2.0] * 6)
            m.income = mo.Variable(m.price * m.units)
        m.income.actual = [19.0, 21.0, 20.0, 20.0, 20.0, 18.0]
        v = m.income._value
        assert v['plan'] == [20.0] * 6                    # formula-derived
        assert v['actual'] == [19.0, 21.0, 20.0, 20.0, 20.0, 18.0]
        assert m.income.formula == 'price * units'        # ONE formula, intact

    def test_undeclared_name_teaches(self):
        m = _declared()
        with m:
            m.x = mo.Variable(actual=[1.0] * 6, plan=[1.0] * 6)
        with pytest.raises(ValueError, match="declared track"):
            m.x.plam = [1.0] * 6      # typo of plan

    def test_wrong_length_refused(self):
        m = _declared()
        with m:
            m.x = mo.Variable(actual=[1.0] * 6, plan=[1.0] * 6)
        with pytest.raises(ValueError, match="6-period"):
            m.x.plan = [1.0, 2.0]

    def test_variable_operand_becomes_dep_edge(self):
        m = _declared()
        with m:
            m.ledger = mo.Variable([5.0] * 6)
            m.x = mo.Variable(plan=[1.0] * 6)
        m.x.actual = m.ledger
        assert any(r is m.ledger for r in m.x._dependency_refs)
        assert m.x._value['actual'] == [5.0] * 6

    def test_floating_binding_stashes_until_adoption(self):
        v = mo.Variable()
        v.actual = [1.0, 2.0, 3.0]
        v.plan = [4.0, 5.0, 6.0]
        m = _declared(periods=3)
        m.v = v
        assert m.v._value['actual'] == [1.0, 2.0, 3.0]
        assert m.v._value['plan'] == [4.0, 5.0, 6.0]

    def test_engine_attrs_untouched(self):
        v = mo.Variable([1.0])
        v.var_type = 'list'        # plain engine attr — normal path
        v.formula = 'a + b'        # property setter — normal path
        assert v.var_type == 'list'



class TestBlendLiveTrack:
    """The blended ``live`` track is a coordinate of its own, computed
    like any other track. Data lines splice givens; memoryless
    formulas coincide with the splice via broadcast; stateful chains
    re-anchor because they roll over live operand flows. The plan (here
    the budget) is never rewritten."""

    def _m(self, until='2025-02', periods=4):
        return mo.Model('m',
            tracks=mo.Tracks('actual', 'budget',
                             blend=mo.blend(given='actual', follow='budget',
                                            until=until)),
            default_grain='month', default_start='2025-01',
            default_periods=periods)

    def _flows(self, m):
        with m:
            m.flow = mo.Variable(actual=[8.0, 7.0, 0.0, 0.0],
                                 budget=[10.0] * 4)
        return m.flow

    def test_data_line_splices_givens(self):
        m = self._m()
        flow = self._flows(m)
        v = flow._value
        assert v['live'] == [8.0, 7.0, 10.0, 10.0]
        assert v['actual'] == [8.0, 7.0, 0.0, 0.0]    # untouched
        assert v['budget'] == [10.0] * 4              # the plan is inviolable

    def test_memoryless_formula_coincides_with_splice(self):
        m = self._m()
        flow = self._flows(m)
        with m:
            m.doubled = flow * 2
        assert m.doubled._value['live'] == [16.0, 14.0, 20.0, 20.0]

    def test_recurrence_re_anchors(self):
        """live[t>close] continues LIVE's own past, not the plan's — the
        whole reason live is a track of its own, not an output splice."""
        m = self._m()
        flow = self._flows(m)
        with m:
            m.balance = mo.recurrence(start=0.0, formula="{prev} + {f}",
                                      variables={"f": flow}, periods=4)
        o = m.balance._value
        assert o['budget'] == [0.0, 10.0, 20.0, 30.0]  # plan from plan
        assert o['actual'] == [0.0, 7.0, 7.0, 7.0]
        assert o['live'] == [0.0, 7.0, 17.0, 27.0]     # 17 ≠ the plan's 20
        assert o['live'][2] == o['live'][1] + 10.0     # re-anchored

    def test_cumsum_re_anchors(self):
        m = self._m()
        flow = self._flows(m)
        acc = mo.cumsum(flow)
        assert acc._value['live'] == [8.0, 15.0, 25.0, 35.0]

    def test_moving_the_boundary_re_splices(self):
        m = self._m(until='2025-01')
        flow = self._flows(m)
        assert flow._value['live'] == [8.0, 10.0, 10.0, 10.0]

    def test_missing_given_falls_back_to_follow(self):
        m = self._m()
        with m:
            m.rent = mo.Variable(budget=[50.0] * 4)
        assert m.rent._value['live'] == [50.0] * 4

    def test_dotted_and_at_slices_reach_live(self):
        m = self._m()
        flow = self._flows(m)
        assert flow.live._value == [8.0, 7.0, 10.0, 10.0]
        assert flow.at(track='live')._value == [8.0, 7.0, 10.0, 10.0]

    def test_dotted_binding_re_splices(self):
        m = self._m()
        with m:
            m.flow = mo.Variable(budget=[10.0] * 4)
        assert m.flow._value['live'] == [10.0] * 4
        m.flow.actual = [8.0, 7.0, 0.0, 0.0]
        assert m.flow._value['live'] == [8.0, 7.0, 10.0, 10.0]

    def test_blend_validates_declared_tracks(self):
        with pytest.raises(ValueError, match="not\\s+declared"):
            mo.Tracks('actual', blend=mo.blend(given='actual', follow='plan',
                                               until='x'))
        with pytest.raises(ValueError, match="same track"):
            mo.blend(given='actual', follow='actual', until='x')
        with pytest.raises(ValueError, match="collides"):
            mo.Tracks('actual', 'budget', 'live',
                      blend=mo.blend(given='actual', follow='budget',
                                     until='x'))

    def test_no_blend_no_live(self):
        m = mo.Model('m', tracks=mo.Tracks('actual', 'budget'),
                     default_grain='month', default_start='2025-01',
                     default_periods=4)
        with m:
            m.x = mo.Variable(actual=[1.0] * 4, budget=[2.0] * 4)
        assert 'live' not in m.x._value.roles


class TestRoleBoundCompoundLists:
    """A role kwarg bound to a LIST of Variables lowers its elements.

    The positional path lowers compound lists through the formula
    builder; the role path used to store raw Variable objects inside
    TrackValues — and from there they leaked into serialized output
    as reprs. Elements must lower to values with dependency
    edges, exactly like a whole-Variable operand.
    """

    def _m(self):
        return mo.Model('m',
            tracks=mo.Tracks('plan', 'forecast', 'actual',
                             blend=mo.blend(given='actual', follow='forecast')),
            default_grain='month', default_start='2026-01',
            default_periods=3)

    def test_variable_elements_lower_to_values(self):
        m = self._m()
        with m:
            m.inputs = mo.MultiVariable('Inputs')
            with m.inputs as inp:
                inp.amount = mo.Variable(10.0)
            row = [mo.Variable(0.0), inp.amount, mo.Variable(0.0)]
            m.x = mo.Variable(plan=row, forecast=row, actual=[None] * 3)
        v = m.x._value
        for role in ('plan', 'forecast'):
            assert v[role] == [0.0, 10.0, 0.0]
            assert all(not isinstance(c, mo.Variable) for c in v[role])
        assert v['actual'] == [None, None, None]

    def test_lowered_elements_keep_dependency_edges(self):
        m = self._m()
        with m:
            m.inputs = mo.MultiVariable('Inputs')
            with m.inputs as inp:
                inp.amount = mo.Variable(10.0)
            m.x = mo.Variable(plan=[inp.amount, 0.0, 0.0], actual=[None] * 3)
        assert any(d is m.inputs.amount for d in m.x._dependency_refs)

    def test_valueless_element_is_loud(self):
        m = self._m()
        with m:
            ghost = mo.Variable(formula='deferred')
            ghost._value = None
            with pytest.raises(ValueError):
                m.x = mo.Variable(plan=[ghost, 0.0, 0.0])


class TestBlendWithoutBoundary:
    """``mo.blend`` without ``until`` — actuals that arrive as events.

    A boundary is a promise that the given track is complete up to a
    date. Facts entered as they happen make no such promise: they
    land irregularly, out of order, and sometimes not at all.
    The boundary-less spec says the rule directly — given where it
    carries a value, follow everywhere else — with no date to keep
    current.
    """

    def _m(self, periods=6):
        return mo.Model('m',
            tracks=mo.Tracks('plan', 'forecast', 'actual',
                             blend=mo.blend(given='actual', follow='forecast')),
            default_grain='month', default_start='2026-01',
            default_periods=periods)

    def test_given_wins_wherever_it_has_a_value(self):
        m = self._m()
        with m:
            m.flow = mo.Variable(plan=[10.0] * 6, forecast=[9.0] * 6,
                                 actual=[7.0, None, 0.0, None, 33.0, None])
        v = m.flow._value
        # 0.0 is a FACT ("nothing happened"), not a hole — it must win
        # over the forecast; only None yields.
        assert v['live'] == [7.0, 9.0, 0.0, 9.0, 33.0, 9.0]
        assert v['actual'] == [7.0, None, 0.0, None, 33.0, None]
        assert v['plan'] == [10.0] * 6        # the plan is inviolable

    def test_a_late_period_actual_is_not_overruled(self):
        """The difference a boundary makes: past it, the bounded form
        prefers ``follow`` and a typed actual loses. Without a boundary
        the actual stands wherever the author put it."""
        bounded = mo.Model('b',
            tracks=mo.Tracks('forecast', 'actual',
                             blend=mo.blend(given='actual', follow='forecast',
                                            until='2026-01')),
            default_grain='month', default_start='2026-01',
            default_periods=3)
        with bounded:
            bounded.x = mo.Variable(forecast=[9.0] * 3,
                                    actual=[None, 5.0, None])
        assert bounded.x._value['live'] == [9.0, 9.0, 9.0]   # actual lost

        free = self._m(periods=3)
        with free:
            free.x = mo.Variable(plan=[9.0] * 3, forecast=[9.0] * 3,
                                 actual=[None, 5.0, None])
        assert free.x._value['live'] == [9.0, 5.0, 9.0]      # actual stands

    def test_empty_given_leaves_follow_intact(self):
        m = self._m(periods=3)
        with m:
            m.x = mo.Variable(plan=[1.0] * 3, forecast=[2.0] * 3,
                              actual=[None, None, None])
        assert m.x._value['live'] == [2.0, 2.0, 2.0]

    def test_formulas_carry_the_boundary_less_live(self):
        m = self._m(periods=3)
        with m:
            m.a = mo.Variable(plan=[10.0] * 3, forecast=[9.0] * 3,
                              actual=[7.0, None, None])
            m.b = mo.Variable(plan=[1.0] * 3, forecast=[1.0] * 3,
                              actual=[2.0, None, None])
            m.s = mo.Variable(m.a + m.b)
        assert m.s._value['live'] == [9.0, 10.0, 10.0]

    def test_repr_omits_the_absent_boundary(self):
        spec = mo.blend(given='actual', follow='forecast')
        assert 'until' not in repr(spec)
        assert repr(mo.blend(given='a', follow='f', until='2026-01')).count(
            'until') == 1

    def test_a_blank_boundary_is_still_a_mistake(self):
        # Omitting the boundary is a choice; spelling it empty is a typo.
        with pytest.raises(TypeError):
            mo.blend(given='actual', follow='forecast', until='')


class TestBlendHardening:
    """Defects in the first blend implementation, each reproduced
    before its fix."""

    def _m(self, until='2025-02', periods=4, name=None):
        kw = dict(given='actual', follow='budget', until=until)
        if name:
            kw['name'] = name
        return mo.Model('m',
            tracks=mo.Tracks('actual', 'budget', blend=mo.blend(**kw)),
            default_grain='month', default_start='2025-01',
            default_periods=periods)

    def test_until_compares_parsed_anchors_not_strings(self):
        """Lexicographic '2025-12' <= '2025-2' silently misrouted the
        whole splice. Boundaries now compare PARSED anchors: a lenient
        spelling splices correctly; a malformed one teaches."""
        m = self._m(until='2025-2')     # unpadded February — still February
        with m:
            m.x = mo.Variable(actual=[8.0, 7.0, 0.0, 0.0],
                              budget=[10.0] * 4)
        assert m.x._value['live'] == [8.0, 7.0, 10.0, 10.0]

        m2 = self._m(until='some date')
        with pytest.raises(ValueError, match="period label"):
            with m2:
                m2.x = mo.Variable(actual=[1.0] * 4, budget=[2.0] * 4)

    def test_scalar_lines_get_live(self):
        """Scalar-per-track assumptions skipped synthesis and broke
        every multiplication via the mismatch law."""
        m = self._m()
        with m:
            m.volume = mo.Variable(actual=[8.0, 7.0, 0.0, 0.0],
                                   budget=[10.0] * 4)
            m.price = mo.Variable(actual=100.0, budget=90.0)
            m.revenue = m.volume * m.price
        assert m.price._value['live'] == [100.0, 100.0, 90.0, 90.0]
        assert m.revenue._value['live'] == [800.0, 700.0, 900.0, 900.0]

    def test_length_one_track_splices_as_constant(self):
        m = self._m()
        with m:
            m.y = mo.Variable(actual=[8.0], budget=[10.0])
        assert m.y._value['live'] == [8.0, 8.0, 10.0, 10.0]

    def test_ragged_track_still_teaches(self):
        """Synthesis ran before the extent law and crashed with a raw
        IndexError; the teaching error must win."""
        m = self._m()
        with pytest.raises(ValueError, match="4-period window"):
            with m:
                m.z = mo.Variable(actual=[8.0, 7.0], budget=[10.0, 10.0])

    def test_tracks_dict_authoring_live_is_loud(self):
        m = self._m()
        with pytest.raises(ValueError, match="synthesizes"):
            with m:
                m.x = mo.Variable(tracks={'live': [1.0] * 4,
                                          'budget': [1.0] * 4})

    def test_kwarg_authoring_live_names_the_blend(self):
        m = self._m()
        with pytest.raises(ValueError, match="synthesizes"):
            with m:
                m.x = mo.Variable(live=[1.0] * 4, budget=[1.0] * 4)

    def test_dotted_write_to_live_teaches_sources(self):
        m = self._m()
        with m:
            m.x = mo.Variable(actual=[1.0] * 4, budget=[2.0] * 4)
        with pytest.raises(ValueError, match="SYNTHESIZED"):
            m.x.live = [9.0] * 4

    def test_blend_without_window_refused_at_declaration(self):
        with pytest.raises(ValueError, match="window"):
            mo.Model('m', tracks=mo.Tracks(
                'actual', 'budget',
                blend=mo.blend(given='actual', follow='budget',
                               until='2025-02')))

    def test_pin_on_derived_line_continues_inherited_live(self):
        """A binding on a formula line must NOT replace the inherited
        live tail with the follow track (an output splice — exactly
        what computing live as a track of its own avoids)."""
        m = self._m()
        with m:
            m.flow = mo.Variable(actual=[8.0, 7.0, 0.0, 0.0],
                                 budget=[10.0] * 4)
            m.doubled = mo.Variable(m.flow * 2)
        inherited_tail = m.doubled._value['live'][2:]
        m.doubled.actual = [16.0, 15.0, 0.0, 0.0]
        v = m.doubled._value
        assert v['live'][:2] == [16.0, 15.0]          # pinned data
        assert v['live'][2:] == inherited_tail        # inherited, not budget

    def test_copy_into_blendless_context_sheds_live(self):
        m = self._m()
        with m:
            m.x = mo.Variable(actual=[1.0] * 4, budget=[2.0] * 4)
        assert 'live' in m.x._value.roles
        plain = mo.Model('plain', tracks=mo.Tracks('actual', 'budget'),
                         default_grain='month', default_start='2025-01',
                         default_periods=4)
        plain.x = m.x                     # ownership change → copy
        assert 'live' not in plain.x._value.roles

    def test_none_holes_fall_to_follow(self):
        m = self._m()
        with m:
            m.x = mo.Variable(actual=[8.0, None, 0.0, 0.0],
                              budget=[10.0] * 4)
        assert m.x._value['live'][1] == 10.0

    def test_own_located_lines_skip_synthesis(self):
        m = self._m()
        with m:
            m.yr = mo.Variable(tracks={'actual': [1.0], 'budget': [2.0]},
                               start='2025-01', grain='year')
        assert 'live' not in m.yr._value.roles


class TestExcelViewTracksParam:
    """The view owns the PRINTED form (resolved through the ExcelView
    cascade): tracks='blend' prints the clean workbook (display-default
    rows only), 'rows' adds labeled per-track rows, 'compare' draws
    period-major column groups; grain= declares the grain the view
    prints at."""

    def _m(self, view=None):
        m = mo.Model('m',
            tracks=mo.Tracks('actual', 'budget',
                             blend=mo.blend(given='actual', follow='budget',
                                            until='2025-02')),
            default_grain='month', default_start='2025-01',
            default_periods=4)
        with m:
            m.flow = mo.Variable(actual=[8.0, 7.0, 0.0, 0.0],
                                 budget=[10.0] * 4)
        if view is not None:
            m.default_excel_view = view
        return m

    def _labels(self, m):
        import os
        import tempfile
        from openpyxl import load_workbook
        p = os.path.join(tempfile.mkdtemp(), 't.xlsx')
        m.to_excel(p)
        return [r[0] for r in load_workbook(p).worksheets[0]
                .iter_rows(min_row=2, max_row=4, max_col=1,
                           values_only=True)]

    def test_rows_is_the_default(self):
        assert self._labels(self._m()) == [
            'Flow', 'flow · actual', 'flow · budget',
        ]

    def test_blend_prints_default_rows_only(self):
        labels = self._labels(self._m(mo.ExcelView(tracks='blend')))
        assert labels[0] == 'Flow' and labels[1] is None

    def test_compare_draws_period_major_column_groups(self):
        # The compare layout: each period is a GROUP of one column per
        # track, the blend column a LIVE splice over its group
        # neighbours, the deviation ("Var") a live difference.
        import os
        import tempfile

        from openpyxl import load_workbook

        p = os.path.join(tempfile.mkdtemp(), 't.xlsx')
        self._m(mo.ExcelView(tracks='compare')).to_excel(p)
        ws = load_workbook(p).worksheets[0]
        # One row — the line keeps its name; no "· actual" sub-rows.
        labels = [ws.cell(r, 1).value for r in range(1, ws.max_row + 1)]
        assert 'Flow' in labels
        assert not any('·' in str(x) for x in labels if x)
        row = labels.index('Flow') + 1
        # Track words under the period labels, one group per month:
        # given, follow, blend (declared order — no selection given).
        assert [ws.cell(2, c).value for c in range(2, 6)] == [
            'actual', 'budget', 'live', 'Var',
        ]
        assert ws.cell(1, 2).value is not None      # period label
        # The blend column splices its own GROUP NEIGHBOURS, live.
        assert ws.cell(row, 4).value == '=IF(B3="",C3,B3)'
        # The deviation is a live difference, not a baked number.
        assert ws.cell(row, 5).value == '=B3 - C3'
        # Numbers land per track: actual 8, budget 10.
        assert ws.cell(row, 2).value == 8
        assert ws.cell(row, 3).value == 10
        # The next month's group starts one stride (4) over.
        assert ws.cell(row, 6).value == 7

    def test_cumsum_prints_a_live_chain_under_compare(self):
        # A tracked ``mo.cumsum`` keeps its formula SYMBOLIC (the
        # rank-lifted MethodCall) instead of delegating to recurrence,
        # and the renderer used to print it verbatim — ``=B3.cumsum()``
        # in every cell of the row. The chain law: period 0 reads the
        # input's cell, period t reads OWN previous-period cell (same
        # track — the per-track address list absorbs the compare
        # layout's column stride) plus the input's current cell.
        # Cross-sheet consumers qualify the input, never the own-row
        # prev.
        import os
        import tempfile

        from openpyxl import load_workbook

        m = mo.Model('m', tracks=mo.Tracks('plan', 'actual'),
                     default_grain='month', default_start='2025-01',
                     default_periods=3)
        m.detail = mo.MultiVariable(display_name='New Clients',
                                    excel_props={'tab': True})
        m.detail.new = mo.Variable(tracks={'plan': [10.0, 20.0, 30.0],
                                           'actual': [11.0, 19.0, 0.0]})
        m.detail.cumulative = mo.Variable(mo.cumsum(m.detail.new),
                                          display_name='Cumulative')
        m.summary = mo.MultiVariable(display_name='Summary',
                                     excel_props={'tab': True})
        m.summary.to_date = mo.Variable(mo.cumsum(m.detail.new),
                                        display_name='To Date')
        m.default_excel_view = mo.ExcelView(tracks='compare')

        p = os.path.join(tempfile.mkdtemp(), 't.xlsx')
        m.to_excel(p)
        wb = load_workbook(p)
        detail_ws, summary_ws = wb['New Clients'], wb['Summary']
        row = next(r for r in range(1, detail_ws.max_row + 1)
                   if detail_ws.cell(r, 1).value == 'Cumulative')
        # Own sheet: seed reads the input, later periods chain OWN
        # prev + input, one stride (3: plan/actual/var) apart.
        assert detail_ws.cell(row, 2).value == '=B3'
        assert detail_ws.cell(row, 5).value == '=B4 + E3'
        assert detail_ws.cell(row, 6).value == '=C4 + F3'
        # Cross-sheet: the input is qualified, the prev is not.
        srow = next(r for r in range(1, summary_ws.max_row + 1)
                    if summary_ws.cell(r, 1).value == 'To Date')
        assert summary_ws.cell(srow, 2).value == "='New Clients'!B3"
        assert summary_ws.cell(srow, 5).value == "=B3 + 'New Clients'!E3"
        # And nothing anywhere prints the Python spelling.
        for ws in wb.worksheets:
            for r in ws.iter_rows():
                for c in r:
                    assert '.cumsum()' not in str(c.value)

    def _metrics_book(self, metrics, plan='plan', total='total'):
        # A metric selection is the caller's: it goes to
        # ``expand_tracked_tree`` directly (a declared compare view
        # shows the deviation alone), and the expanded tree is written
        # as is. Returns the sheet and each emitted line's values by
        # display name.
        import os
        import tempfile

        from openpyxl import load_workbook

        from modeleon.compile.excel.writer import expand_tracked_tree

        m = mo.Model('m', tracks=mo.Tracks(plan, 'actual'),
                     default_grain='month', default_start='2025-01',
                     default_periods=2)
        with m:
            m.p = mo.MultiVariable('P', excel_props={'tab': True})
            with m.p as p:
                p.a = mo.Variable(tracks={plan: [10.0, 20.0],
                                          'actual': [12.0, 18.0]},
                                  display_name='A')
                p.b = mo.Variable(tracks={plan: [30.0, 20.0],
                                          'actual': [28.0, 22.0]},
                                  display_name='B')
                setattr(p, total, mo.Variable(p.a + p.b, display_name='Total'))
        expanded = expand_tracked_tree(m, mode='compare', metrics=metrics)
        path = os.path.join(tempfile.mkdtemp(), 't.xlsx')
        expanded.to_excel(path)
        values = {v.display_name: v.value
                  for v in expanded.p._components.values()}
        return load_workbook(path)['P'], values

    def test_older_metric_spellings_still_select_their_columns(self):
        # "откл" and "уд.вес" are the older spellings of var and share,
        # still accepted, and a column is captioned the way it was
        # spelled. The non-Latin text is deliberate: it is the
        # compatibility surface under test.
        old, old_values = self._metrics_book(['откл', 'уд.вес'])
        new, new_values = self._metrics_book(['var', 'share'])
        assert [old.cell(2, c).value for c in range(2, 6)] == [
            'plan', 'actual', 'откл', 'уд. вес',
        ]
        assert [new.cell(2, c).value for c in range(2, 6)] == [
            'plan', 'actual', 'Var', 'Share',
        ]
        # The same metrics under either caption: the same live cells.
        assert old.cell(3, 4).value == '=C3 - B3'
        assert old.cell(3, 5).value == '=B3 / B5'
        for c in (4, 5, 8, 9):
            assert old.cell(3, c).value == new.cell(3, c).value
        assert old_values['A · откл'] == new_values['A · Var'] == [2.0, -2.0]
        assert old_values['A · уд. вес'] == new_values['A · Share'] == [0.25, 0.5]

    def test_share_measures_against_total_in_either_spelling(self):
        # A share's denominator is the section's line named "total" —
        # or "итого", the older spelling (the non-Latin name is
        # deliberate). The same model gives the same numbers either way.
        for total in ('total', 'итого'):
            ws, values = self._metrics_book(['share'], total=total)
            rows = {ws.cell(r, 1).value: r for r in range(3, ws.max_row + 1)}
            assert [ws.cell(2, c).value for c in range(2, 5)] == [
                'plan', 'actual', 'Share',
            ]
            assert ws.cell(rows['A'], 4).value == f"=B{rows['A']} / B{rows['Total']}"
            assert ws.cell(rows['B'], 4).value == f"=B{rows['B']} / B{rows['Total']}"
            assert values['A · Share'] == [0.25, 0.5]
            assert values['B · Share'] == [0.75, 0.5]

    def test_deviation_is_actual_minus_plan_in_either_spelling(self):
        # The deviation subtracts the track named "plan" — or "план",
        # the older spelling (the non-Latin name is deliberate). The
        # plan is declared FIRST, so a positional rule would get the
        # sign backwards.
        for plan in ('plan', 'план'):
            ws, values = self._metrics_book(['var'], plan=plan)
            assert [ws.cell(2, c).value for c in range(2, 5)] == [
                plan, 'actual', 'Var',
            ]
            assert ws.cell(3, 4).value == '=C3 - B3'
            assert values['A · Var'] == [2.0, -2.0]
            assert values['B · Var'] == [-2.0, 2.0]

    def test_vocabularies(self):
        with pytest.raises(ValueError, match="blend"):
            mo.ExcelView(tracks='nonsense')
        with pytest.raises(ValueError, match="grain"):
            mo.ExcelView(grain='decade')

    def test_resolved_view_carries_fields(self):
        from modeleon.compile.excel.view import resolve_excel_view
        m = self._m(mo.ExcelView(tracks='blend', grain='quarter'))
        r = resolve_excel_view(m.flow)
        assert r.tracks == 'blend'
        assert r.grain == 'quarter'


class TestABlankCountsAsZero:
    """As in Excel, a month nobody entered counts as zero in a formula:
    a balance stays where it was, a profit over blank months is 0. Only
    an input stays blank. Whether an actual was entered is the blend's
    question, asked of the entries a formula is built from - a derived
    line takes its live value from its operands' live values either way."""

    def test_a_balance_stays_where_it_was(self):
        m = mo.Model('m',
            tracks=mo.Tracks('actual', 'budget',
                             blend=mo.blend(given='actual', follow='budget',
                                            until='2025-02')),
            default_grain='month', default_start='2025-01',
            default_periods=4)
        with m:
            m.flow = mo.Variable(actual=[8.0, 7.0, None, None],
                                 budget=[10.0] * 4)
            m.balance = mo.recurrence(start=0.0, formula="{prev} + {f}",
                                      variables={"f": m.flow}, periods=4)
        o = m.balance._value
        assert o['actual'] == [0.0, 7.0, 7.0, 7.0]
        assert o['budget'] == [0.0, 10.0, 20.0, 30.0]
        assert o['live'] == [0.0, 7.0, 17.0, 27.0]      # re-anchored, alive

    def test_a_formula_over_blank_months_is_zero(self):
        m = mo.Model('m',
            tracks=mo.Tracks('actual', 'budget',
                             blend=mo.blend(given='actual', follow='budget',
                                            until='2025-02')),
            default_grain='month', default_start='2025-01',
            default_periods=4)
        with m:
            m.revenue = mo.Variable(actual=[8.0, 7.0, None, None],
                                    budget=[10.0] * 4)
            m.expenses = mo.Variable(actual=[6.0, 5.0, None, None],
                                     budget=[7.0] * 4)
            m.profit = mo.Variable(m.revenue - m.expenses)
        p = m.profit._value
        assert p['actual'] == [2.0, 2.0, 0.0, 0.0]
        assert p['budget'] == [3.0] * 4
        assert p['live'] == [2.0, 2.0, 3.0, 3.0]

    def test_a_derived_actual_takes_the_forecast_where_nothing_was_entered(self):
        m = mo.Model('m',
            tracks=mo.Tracks('plan', 'forecast', 'actual',
                             blend=mo.blend(given='actual', follow='forecast',
                                            name='live')),
            default_grain='month', default_start='2025-01',
            default_periods=4)
        with m:
            m.a = mo.Variable([10.0] * 4, forecast=[11.0] * 4,
                              actual=[9.0, 8.0, 7.0, None])
            m.b = mo.Variable([20.0] * 4, forecast=[21.0] * 4,
                              actual=[19.0, None, None, None])
            m.total = mo.Variable(m.a + m.b, plan=[30.0] * 4)
        t = m.total._value
        assert t['actual'] == [28.0, 8.0, 7.0, 0.0]
        # Something entered: the actual, its blanks as zeros. Nothing
        # entered: the forecast.
        assert t['live'] == [28.0, 8.0, 7.0, 32.0]

    def test_genuine_errors_still_absorb(self):
        """Absence is quiet; failure stays loud."""
        a = mo.Variable([1.0, 2.0])
        b = mo.Variable([0.0, 2.0])
        assert (a / b)._value[0] == '#DIV/0!'


class TestTracksInsideInstances:
    """A portfolio of project instances, each with a tracked field
    authored by constructor params. Pure track kwargs
    materialize PROVISIONALLY at construction (from their own names),
    so the class body's next line computes; adoption validates the
    names against the declaration, attaches the blend, and memoryless
    derived lines get their live as the splice of their own tracks
    (splice-equal for per-period formulas)."""

    def _model(self):
        class Project(mo.MultiVariableClass):
            def compute(self, rate=100.0, hours_plan=None,
                        hours_actual=None):
                self.hours = mo.Variable(plan=hours_plan, actual=hours_actual)
                self.revenue = self.hours * rate

        m = mo.Model('m',
            tracks=mo.Tracks('actual', 'plan',
                             blend=mo.blend(given='actual', follow='plan',
                                            until='2025-02')),
            default_grain='month', default_start='2025-01',
            default_periods=4)
        with m:
            m.alpha = Project(rate=100.0, hours_plan=[10.0] * 4,
                              hours_actual=[9.0, 11.0, None, None])
            m.beta = Project(rate=150.0, hours_plan=[20.0] * 4,
                             hours_actual=[18.0, 21.0, None, None])
            m.portfolio = m.alpha.revenue + m.beta.revenue
        return m

    def test_tracked_field_in_class_body(self):
        m = self._model()
        v = m.alpha.hours._value
        assert v['plan'] == [10.0] * 4
        assert v['actual'] == [9.0, 11.0, None, None]
        assert v['live'] == [9.0, 11.0, 10.0, 10.0]

    def test_derived_line_in_class_body_gets_spliced_live(self):
        m = self._model()
        v = m.alpha.revenue._value
        assert v['live'] == [900.0, 1100.0, 1000.0, 1000.0]

    def test_portfolio_aggregation_across_instances(self):
        m = self._model()
        p = m.portfolio._value
        assert p['plan'] == [4000.0] * 4
        assert p['actual'][:2] == [3600.0, 4250.0]
        assert p['live'] == [3600.0, 4250.0, 4000.0, 4000.0]

    def test_typo_in_class_body_still_teaches_at_adoption(self):
        class Broken(mo.MultiVariableClass):
            def compute(self):
                self.x = mo.Variable(actuall=[1.0] * 4)

        m = mo.Model('m', tracks=mo.Tracks('actual', 'plan'),
                     default_grain='month', default_start='2025-01',
                     default_periods=4)
        with pytest.raises(ValueError, match="actuall"):
            with m:
                m.k = Broken()

    def test_stateful_derived_in_class_body_has_no_live(self):
        """A recurrence inside a class body rolled before its driver
        had live — splicing its output would break the live track's
        rule (a chain must roll over live flows), so live is honestly
        ABSENT and the mismatch law polices any cross-use."""
        class Accumulator(mo.MultiVariableClass):
            def compute(self, flow_plan=None, flow_actual=None):
                self.flow = mo.Variable(plan=flow_plan, actual=flow_actual)
                self.balance = mo.cumsum(self.flow)

        m = mo.Model('m',
            tracks=mo.Tracks('actual', 'plan',
                             blend=mo.blend(given='actual', follow='plan',
                                            until='2025-02')),
            default_grain='month', default_start='2025-01',
            default_periods=4)
        with m:
            m.n = Accumulator(flow_plan=[10.0] * 4,
                              flow_actual=[9.0, 11.0, None, None])
        assert 'live' not in m.n.balance._value.roles


class TestScalarBlend:
    """Non-temporal lines carry the axis too — without a timeline."""

    def _model(self):
        m = mo.Model(
            'm', display_name='M', default_grain='month',
            default_start='2026-01', default_periods=6,
            tracks=mo.Tracks(
                'plan', 'actual',
                blend=mo.blend(given='actual', follow='plan',
                               until='2026-03', name='forecast')))
        return m

    def test_a_scalar_line_blends_to_a_scalar(self):
        # A record's scalar field — an amount: plan 50,000,
        # actual 52,000.
        # The time splice used to spread the constant into a full
        # 6-period series — a record field quietly turned temporal.
        m = self._model()
        with m:
            m.p = mo.MultiVariable('P', excel_props={'tab': True})
            m.p._records_container = True  # the records-container contract
            with m.p as p:
                p.amount = mo.Variable(plan=50000, actual=52000,
                                       display_name='Amount')
        tv = m.p.amount._value
        assert tv['forecast'] == 52000
        assert not isinstance(tv['forecast'], list)

    def test_actual_absent_the_scalar_forecast_is_the_plan(self):
        m = self._model()
        with m:
            m.p = mo.MultiVariable('P', excel_props={'tab': True})
            m.p._records_container = True
            with m.p as p:
                p.term = mo.Variable(plan=12, display_name='Term')
        assert m.p.term._value['forecast'] == 12

    def test_temporal_lines_still_splice_per_period(self):
        m = self._model()
        with m:
            m.p = mo.MultiVariable('P', excel_props={'tab': True})
            with m.p as p:
                p.series = mo.Variable(plan=[1.0] * 6, display_name='Series')
        assert len(m.p.series._value['forecast']) == 6


class TestBlendRowIsAlive:
    """The blended row references its subrows — the file stays live."""

    def _book(self, model):
        import os
        import tempfile

        import openpyxl

        path = os.path.join(tempfile.mkdtemp(), "b.xlsx")
        model.to_excel(path)
        return openpyxl.load_workbook(path)

    def _model(self, **view_kw):
        m = mo.Model(
            'm', display_name='M', default_grain='month',
            default_start='2026-01', default_periods=4,
            tracks=mo.Tracks(
                'plan', 'actual',
                blend=mo.blend(given='actual', follow='plan',
                               until='2026-02')),
            **view_kw)
        with m:
            m.p = mo.MultiVariable('P', excel_props={'tab': True})
            with m.p as p:
                p.revenue = mo.Variable(
                    plan=[10.0] * 4, actual=[12.0, 11.0, None, None],
                    display_name='Revenue')
                p.tax = p.revenue * 0.1
        return m

    def test_the_splice_is_a_formula_over_its_subrows(self):
        ws = self._book(self._model())['P']
        # Blend row 2; plan row 3; actual row 4. Before the boundary the
        # ACTUAL wins with a fallback; after it the PLAN does.
        assert ws['B2'].value == '=IF(B4="",B3,B4)'
        assert ws['C2'].value == '=IF(C4="",C3,C4)'
        assert ws['D2'].value == '=IF(D3="",D4,D3)'
        assert ws['E2'].value == '=IF(E3="",E4,E3)'

    def test_a_derived_line_keeps_its_own_formula(self):
        # tax = revenue × 0.1 — it references the BLEND row (the live
        # track). Replaying the splice on an output would be wrong.
        ws = self._book(self._model())['P']
        assert ws['B5'].value == '=B2 * 0.1'

    def test_a_computed_plan_with_authored_actuals_still_splices(self):
        # The plan is a FORMULA (over another row), actual and forecast
        # are authored series. The line has both an expression and role
        # kwargs — the engine calls it a DATA line and splices it, so
        # the workbook must splice it too. Left with its base
        # expression the file would show the PLAN where the engine
        # computes the blend (here the forecast places the amount a
        # month later than the plan).
        m = mo.Model(
            'm', display_name='M', default_grain='month',
            default_start='2026-01', default_periods=4,
            tracks=mo.Tracks(
                'plan', 'forecast', 'actual',
                blend=mo.blend(given='actual', follow='forecast')),
        )
        with m:
            m.p = mo.MultiVariable('P', excel_props={'tab': True})
            with m.p as p:
                p.volume = mo.Variable([0.0, 20.0, 0.0, 0.0],
                                       display_name='Volume')
                p.revenue = mo.Variable(
                    p.volume * 1.0,
                    forecast=[0.0, 0.0, 20.0, 0.0],
                    actual=[None, None, None, None],
                    display_name='Revenue')
        ws = self._book(m)['P']
        rows = {ws.cell(r, 1).value: r for r in range(1, ws.max_row + 1)}
        blend = rows['Revenue']
        given = rows['Revenue · actual']
        follow = rows['Revenue · forecast']
        # Boundary-less: given wins wherever it carries a value.
        assert ws.cell(blend, 2).value == (
            f'=IF(B{given}="",B{follow},B{given})')
        # …and the splice, not the plan formula: February is the
        # forecast row's zero, March its 20.
        assert ws.cell(blend, 4).value == (
            f'=IF(D{given}="",D{follow},D{given})')

    def test_the_settled_ink_is_a_rule_not_a_stamp(self):
        # The blended cell goes blue WHILE an actual stands behind it —
        # as a conditional rule over the actual cell, so typing an actual
        # into the written file colours its own month. A stamped
        # font would freeze provenance at write time, the same
        # dead-values habit the splice formulas exist to kill.
        ws = self._book(self._model())['P']
        rules = ws.conditional_formatting
        by_range = {}
        for rng in rules:
            for rule in rng.rules:
                by_range[str(rng.sqref)] = rule
        # Blend row 2, actual row 4: each month keys off ITS actual cell.
        assert by_range['B2'].formula == ['B4<>""']
        assert by_range['C2'].formula == ['C4<>""']
        # Font only — a fill would cut holes in the house style's
        # section bands.
        assert by_range['B2'].dxf.font is not None
        assert by_range['B2'].dxf.fill is None

    def test_the_blend_only_book_bakes_nothing_extra(self):
        # tracks='blend' emits the default row ALONE — no subrows to
        # reference, so the row keeps its computed values.
        m = self._model(
            default_excel_view=mo.ExcelView(tracks='blend'))
        ws = self._book(m)['P']
        assert ws['B2'].value == 12.0
        labels = [ws.cell(r, 1).value for r in range(1, ws.max_row + 1)]
        assert not any('·' in str(x) for x in labels if x)  # no subrows


class TestADerivedActualAsksWhatWasEntered:
    """A line can enter its plan and DERIVE its actual: the actual of a
    sum is the sum of the actuals. Its actual row is then a formula,
    which Excel computes as 0 over the months nobody entered - where the
    engine takes the forecast. So the splice asks the rows of entries
    the formula is built from, as the engine's blend does (nothing
    entered: the forecast), not ``=""`` of the formula's own cell (the
    written book showed Q4 revenue as 0 where the engine computed its
    forecast)."""

    def _model(self):
        m = mo.Model(
            'm', display_name='M', default_grain='month',
            default_start='2026-01', default_periods=4,
            tracks=mo.Tracks('plan', 'forecast', 'actual',
                             blend=mo.blend(given='actual', follow='forecast',
                                            name='live')))
        with m:
            m.p = mo.MultiVariable('Projects', excel_props={'tab': True})
            with m.p as p:
                p.a = mo.Variable([10.0] * 4, forecast=[11.0] * 4,
                                  actual=[9.0, 8.0, None, None], display_name='A')
                p.b = mo.Variable([20.0] * 4, forecast=[21.0] * 4,
                                  actual=[19.0, 18.0, None, None], display_name='B')
                p.rate = mo.Variable(0.5, display_name='Rate')
            m.s = mo.MultiVariable('P and L', excel_props={'tab': True})
            with m.s as s:
                s.revenue = mo.Variable(m.p.a + m.p.b, plan=[25.0] * 4,
                                        display_name='Revenue')
                s.cost = mo.Variable(m.p.a * m.p.rate, plan=[5.0] * 4,
                                     display_name='Cost')
                s.gross = mo.Variable(m.p.a + m.p.b, display_name='Gross')
                s.margin = mo.Variable(s.gross - s.cost, plan=[20.0] * 4,
                                       display_name='Margin')
                s.running = mo.Variable(mo.cumsum(m.p.a), plan=[10.0] * 4,
                                        display_name='Running')
        return m

    def _sheet(self, m, tmp_path, name='P and L'):
        import openpyxl

        path = tmp_path / 'b.xlsx'
        m.to_excel(str(path))
        ws = openpyxl.load_workbook(str(path))[name]
        rows = {ws.cell(r, 1).value: r for r in range(1, ws.max_row + 1)}
        return ws, rows

    def test_the_splice_asks_the_entries_of_a_sum(self, tmp_path):
        ws, rows = self._sheet(self._model(), tmp_path)
        r = rows['Revenue']
        # March: Projects' March is column E (column B holds its Rate).
        # COUNTA, not COUNT: an entered error or TRUE/FALSE is an entry.
        assert ws.cell(r, 4).value == (
            f"=IF(COUNTA(Projects!E5,Projects!E9)=0,"
            f"D{rows['Revenue · forecast']},D{rows['Revenue · actual']})")

    def test_a_constant_is_no_months_entry(self, tmp_path):
        # cost = a × rate: the rate is in every month, so it is no
        # month's entry - one entry is asked directly.
        ws, rows = self._sheet(self._model(), tmp_path)
        assert ws.cell(rows['Cost'], 2).value.startswith(
            '=IF(Projects!C5="",')

    def test_the_count_reads_through_a_derived_line(self, tmp_path):
        # margin = gross - cost, both formulas: their own rows are never
        # blank in Excel, so the count reaches what they are built from.
        ws, rows = self._sheet(self._model(), tmp_path)
        assert ws.cell(rows['Margin'], 2).value.startswith(
            '=IF(COUNTA(Projects!C5,Projects!C9)=0,')

    def test_a_formula_it_cannot_read_keeps_the_blank_test(self, tmp_path):
        ws, rows = self._sheet(self._model(), tmp_path)
        r = rows['Running']
        assert ws.cell(r, 2).value == (
            f"=IF(B{rows['Running · actual']}=\"\",B{rows['Running · forecast']},"
            f"B{rows['Running · actual']})")

    def test_the_settled_ink_counts_the_same_entries(self, tmp_path):
        ws, rows = self._sheet(self._model(), tmp_path)
        ref = f"B{rows['Revenue']}"
        formulas = [rule.formula for rng in ws.conditional_formatting
                    for rule in rng.rules if str(rng.sqref) == ref]
        assert formulas == [['COUNTA(Projects!C5,Projects!C9)>0']]

    def test_a_sheet_name_with_a_space_is_quoted(self):
        from modeleon.compile.excel.blend_writer import _qualified

        assert _qualified('C5', 'P and L', 'Totals') == "'P and L'!C5"
        assert _qualified('C5', "Bob's", 'Totals') == "'Bob''s'!C5"
        assert _qualified('C5', 'Plan\n', 'Totals') == "'Plan\n'!C5"
        assert _qualified('C5', 'Projects', 'Totals') == 'Projects!C5'
        assert _qualified('C5', 'Totals', 'Totals') == 'C5'

    def test_more_entries_than_one_call_takes(self):
        from modeleon.compile.excel.addresses import VariableAddresses
        from modeleon.compile.excel.blend_writer import _entered

        rows = [f'r{k}' for k in range(300)]
        addresses = {r: VariableAddresses(name='A1', formula='B1',
                                          values=[f'B{k + 1}', f'C{k + 1}'])
                     for k, r in enumerate(rows)}
        blank, entered = _entered(rows, 1, 2, addresses, {r: 'S' for r in rows}, 'S')
        assert blank.count('COUNTA(') == 2
        assert blank.startswith('COUNTA(C1,') and '+COUNTA(C256,' in blank
        assert blank.endswith(')=0') and entered.endswith(')>0')
        # A row laid out on other months makes the test unknowable.
        addresses['r7'] = VariableAddresses(name='A1', formula='B1', values=['B8'])
        assert _entered(rows, 1, 2, addresses, {r: 'S' for r in rows}, 'S') is None

    def _partial_model(self):
        m = mo.Model(
            'm', display_name='M', default_grain='month',
            default_start='2026-01', default_periods=4,
            tracks=mo.Tracks('plan', 'forecast', 'actual',
                             blend=mo.blend(given='actual', follow='forecast',
                                            name='live')))
        with m:
            m.p = mo.MultiVariable('Projects', excel_props={'tab': True})
            with m.p as p:
                # B's actual is in for January only: from February the
                # sum is A's actual alone, until neither is entered.
                p.a = mo.Variable([10.0] * 4, forecast=[11.0] * 4,
                                  actual=[9.0, 8.0, 7.0, None], display_name='A')
                p.b = mo.Variable([20.0] * 4, forecast=[21.0] * 4,
                                  actual=[19.0, None, None, None], display_name='B')
                # An entered error is an entry: it carries on, it does
                # not give way to the forecast.
                p.e = mo.Variable([10.0] * 4, forecast=[11.0] * 4,
                                  actual=['#N/A', 8.0, None, None], display_name='E')
            m.s = mo.MultiVariable('P and L', excel_props={'tab': True})
            with m.s as s:
                s.total = mo.Variable(m.p.a + m.p.b, plan=[30.0] * 4,
                                      display_name='Total')
                s.broken = mo.Variable(m.p.e * 2, plan=[20.0] * 4,
                                       display_name='Broken')
        return m

    @pytest.mark.slow
    @pytest.mark.parametrize('which', ['entered', 'partial'])
    def test_the_book_takes_the_forecast_where_the_engine_does(self, tmp_path, which):
        import shutil
        import subprocess
        from pathlib import Path

        import openpyxl

        soffice = shutil.which('soffice') or shutil.which('libreoffice')
        mac = Path('/Applications/LibreOffice.app/Contents/MacOS/soffice')
        if soffice is None and mac.exists():
            soffice = str(mac)
        if soffice is None:
            pytest.skip('LibreOffice is not installed')
        m = self._model() if which == 'entered' else self._partial_model()
        lines = ((('Revenue', m.s.revenue), ('Cost', m.s.cost),
                  ('Margin', m.s.margin)) if which == 'entered' else
                 (('Total', m.s.total), ('Broken', m.s.broken)))
        src = tmp_path / 'book.xlsx'
        m.to_excel(str(src))
        subprocess.run(
            [soffice, f"-env:UserInstallation={(tmp_path / 'lo').as_uri()}",
             '--headless', '--convert-to', 'xlsx', '--outdir',
             str(tmp_path / 'calc'), str(src)],
            check=True, capture_output=True, timeout=120)
        ws = openpyxl.load_workbook(str(tmp_path / 'calc' / 'book.xlsx'),
                                    data_only=True)['P and L']
        rows = {ws.cell(r, 1).value: r for r in range(1, ws.max_row + 1)}
        for label, line in lines:
            book = [ws.cell(rows[label], c).value for c in range(2, 6)]
            want = line.value['live']
            assert [x for x in book if isinstance(x, str)] == [
                x for x in want if isinstance(x, str)], label
            assert [x for x in book if not isinstance(x, str)] == pytest.approx(
                [x for x in want if not isinstance(x, str)]), label


class TestAWrapperReferencesAHandTrackedLine:
    """``mo.Variable(row, plan=[...])`` where ``row`` enters its actual by
    hand: the row's expression describes its other tracks only, so a copy
    of it wrote the plan's formula into the wrapper's actual row (a book
    showed interest as actual for months nobody entered). The wrapper
    references the row instead."""

    def _book(self, tmp_path):
        import openpyxl

        m = mo.Model('m', default_grain='month', default_start='2026-01',
                     default_periods=4,
                     tracks=mo.Tracks('plan', 'forecast', 'actual',
                                      blend=mo.blend(given='actual', follow='forecast',
                                                     name='live')))
        with m:
            m.s = mo.MultiVariable('Savings', excel_props={'tab': True})
            with m.s as s:
                s.principal = mo.Variable(100000.0, display_name='Principal')
                s.rate = mo.Variable(0.06, display_name='Rate')
                s.rev = mo.Variable([s.principal * s.rate / 12 for _ in range(4)],
                                    actual=[s.principal * s.rate / 12,
                                            s.principal * s.rate / 12, None, None],
                                    display_name='Rev')
            m.p = mo.MultiVariable('PnL', excel_props={'tab': True})
            with m.p as p:
                p.x = mo.Variable(m.s.rev, plan=[450.0] * 4, display_name='X')
        path = tmp_path / 'b.xlsx'
        m.to_excel(str(path))
        ws = openpyxl.load_workbook(str(path))['PnL']
        return m, {ws.cell(r, 1).value: [ws.cell(r, c).value for c in range(2, 6)]
                   for r in range(1, ws.max_row + 1)}

    def test_the_actual_row_reads_the_lines_actual(self, tmp_path):
        m, rows = self._book(tmp_path)
        # Savings: Jan is column C (column B holds the constants); Rev's
        # actual is its row 7.
        assert rows['X · actual'] == ['=Savings!C7', '=Savings!D7', '=Savings!E7',
                                      '=Savings!F7']
        assert rows['X · forecast'] == ['=Savings!C6', '=Savings!D6', '=Savings!E6',
                                        '=Savings!F6']

    def test_the_formula_names_the_line(self, tmp_path):
        m, _ = self._book(tmp_path)
        assert m.p.x._expr.to_string() == 'rev'
        # A reference shows a blank as 0, as Excel's =A1 does; nothing
        # was entered there, so the live row takes the forecast.
        assert m.p.x.value['actual'][2] == 0
        assert m.p.x.value['live'][2] == 500.0


class TestTrackSubrowsCarryTheirOwnFormula:
    """A DERIVED track explains itself in ITS OWN coordinates.

    The blend row has always carried the line's formula over its
    operands' blend rows. The other subrows got numbers, so "Total ·
    plan" sat frozen under a row that recomputed — change January's
    plan in the written file and the plan total did not move. The
    engine already knows the answer (a derived track's value IS the
    line's expression over the operands' same-track values), so the
    same expression through a per-track address book is the honest
    formula.
    """

    def _book(self, tmp_path, **kw):
        import openpyxl

        m = mo.Model(
            "m", display_name="M",
            tracks=mo.Tracks("plan", "actual",
                             blend=mo.blend(given="actual", follow="plan")),
            default_grain="month", default_start="2026-01",
            default_periods=2,
            default_excel_view=mo.ExcelView(tracks="rows"),
        )
        with m:
            m.s = mo.MultiVariable("S", excel_props={"tab": True})
            with m.s as s:
                s.a = mo.Variable([10.0, 20.0], display_name="A",
                                  actual=[11.0, None])
                s.rate = mo.Variable(0.5, display_name="Rate")
                s.derived = mo.Variable(s.a * s.rate, display_name="Derived")
                s.mixed = mo.Variable(s.a + s.a, display_name="Mixed",
                                      actual=[99.0, 98.0])
        path = tmp_path / "b.xlsx"
        m.to_excel(str(path))
        ws = openpyxl.load_workbook(str(path))["S"]
        return {
            (ws.cell(row=r, column=1).value
             or ws.cell(row=r, column=2).value): [
                ws.cell(row=r, column=c).value for c in (3, 4)
            ]
            for r in range(1, 40)
            if (ws.cell(row=r, column=1).value
                or ws.cell(row=r, column=2).value)
        }

    def test_a_derived_track_reads_its_own_track(self, tmp_path):
        rows = self._book(tmp_path)
        plan = rows["Derived · plan"][0]
        actual = rows["Derived · actual"][0]
        assert isinstance(plan, str) and plan.startswith("=")
        assert isinstance(actual, str) and actual.startswith("=")
        # …and they point at DIFFERENT rows — each its own track's.
        assert plan != actual

    def test_an_untracked_operand_keeps_its_single_cell(self, tmp_path):
        rows = self._book(tmp_path)
        # "Rate" has no tracks, so every track's formula multiplies by
        # the same cell — a constant reads the same from anywhere.
        plan, actual = rows["Derived · plan"][0], rows["Derived · actual"][0]
        shared = set(re.findall(r"B\d+", plan)) & set(re.findall(r"B\d+", actual))
        assert shared, (plan, actual)

    def test_an_authored_track_stays_data(self, tmp_path):
        rows = self._book(tmp_path)
        # "A" typed both tracks; "Mixed" typed only actual. Typed series
        # are data — a formula there would overwrite what the user
        # entered.
        assert rows["A · plan"][0] == 10.0
        assert rows["A · actual"][0] == 11.0
        assert rows["Mixed · actual"][0] == 99.0
        # …while Mixed's plan is derived and does get one.
        assert str(rows["Mixed · plan"][0]).startswith("=")

    def test_the_blend_row_is_left_to_its_own_writer(self, tmp_path):
        rows = self._book(tmp_path)
        # A spliced line's blend still asks its subrows; a derived
        # line's blend still carries the shared formula.
        assert str(rows["A"][0]).startswith("=IF(")
        assert str(rows["Derived"][0]).startswith("=")


class TestLoopBuiltIntermediatesReferenceTheirRows:
    """A loop-built model computes intermediates as per-month
    expressions in local Python lists, then materializes the same
    lists as rows. The row's wrapper shares the ORIGINAL expression
    tree with every downstream consumer — identity is the only join —
    and without it each consumer unrolled the intermediates wholesale:
    a consumer cell repeated an intermediate's expression instead of
    referencing the row that already holds it, growing longer with
    every hop.
    """

    def _book(self, tmp_path):
        import openpyxl

        m = mo.Model(
            "m", display_name="M",
            tracks=mo.Tracks("plan", "actual",
                             blend=mo.blend(given="actual", follow="plan")),
            default_grain="month", default_start="2026-01",
            default_periods=2,
        )
        with m:
            m.s = mo.MultiVariable("S", excel_props={"tab": True})
            with m.s as s:
                s.rate = mo.Variable(0.1, display_name="Rate")
                gross = [mo.Variable(100.0), mo.Variable(200.0)]
                fee = [x * s.rate for x in gross]
                net = [x - f for x, f in zip(gross, fee)]
                s.gross = mo.Variable(gross, display_name="Gross",
                                      actual=[None, None])
                s.fee = mo.Variable(fee, display_name="Fee",
                                    actual=[None, None])
                s.net = mo.Variable(net, display_name="Net",
                                    actual=[None, None])
        path = tmp_path / "b.xlsx"
        m.to_excel(str(path))
        ws = openpyxl.load_workbook(str(path))["S"]
        return {
            (ws.cell(row=r, column=1).value
             or ws.cell(row=r, column=2).value): [
                ws.cell(row=r, column=c).value for c in (3, 4)
            ]
            for r in range(1, 30)
            if (ws.cell(row=r, column=1).value
                or ws.cell(row=r, column=2).value)
        }

    def test_the_consumer_references_the_intermediate_rows(self, tmp_path):
        rows = self._book(tmp_path)
        fee = str(rows["Fee · plan"][0])
        net = str(rows["Net · plan"][0])
        # Fee references the gross subrow's cell, not an unrolled
        # copy of its expression…
        assert fee.startswith("=") and "*" in fee
        # …and Net references BOTH rows by cell: it must contain
        # cell refs and no second copy of the multiplication.
        assert net.startswith("=")
        assert "*" not in net, net
    def test_lengths_stay_flat(self, tmp_path):
        rows = self._book(tmp_path)
        # The unrolled form grew with every hop; the referenced form
        # stays a handful of cells.
        assert len(str(rows["Net · plan"][0])) < 30



def test_a_grain_mismatch_names_both_lines():
    m = mo.Model('Deal', default_grain='year', default_start='2026', default_periods=2)
    with m:
        m.ops = mo.MultiVariable('Ops')
        with m.ops as o:
            o.cash = mo.Variable([40.0, 40.0], display_name='Cash')
        m.debt = mo.MultiVariable('Debt', default_grain='quarter', default_start='2026-Q1',
                                  default_periods=8)
        with m.debt as d:
            d.interest = mo.Variable([1.0] * 8, display_name='Interest')
            with pytest.raises(ValueError) as err:
                d.left = m.ops.cash - d.interest
    msg = str(err.value)
    assert "'Cash' (years)" in msg and "'Interest' (quarters)" in msg
    assert "interest.at('year')" in msg
