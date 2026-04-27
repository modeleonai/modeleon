# SPDX-License-Identifier: Apache-2.0
"""Tests for ``QPath`` — the qualified-path identity type (ADR-008)."""

from __future__ import annotations

import pytest

from modeleon.core.qpath import QPath


class TestConstruction:
    def test_empty_is_root(self):
        assert QPath().is_root
        assert QPath(()) == QPath()

    def test_from_segments(self):
        p = QPath(("a", "b", "c"))
        assert p.segments == ("a", "b", "c")


class TestRoot:
    def test_root_is_distinct(self):
        assert QPath.ROOT.is_root
        assert not QPath.ROOT.is_floating

    def test_root_is_reusable_singleton(self):
        assert QPath.ROOT == QPath()


class TestFloating:
    def test_floating_namespace(self):
        p = QPath.floating(3, kind="v")
        assert p.is_floating
        assert p.segments == ("__floating__", "v3")

    def test_floating_mv_kind(self):
        p = QPath.floating(7, kind="m")
        assert p.is_floating
        assert p.segments == ("__floating__", "m7")

    def test_rooted_path_is_not_floating(self):
        assert not QPath(("acme", "pnl", "revenue")).is_floating


class TestStr:
    def test_dotted_rendering(self):
        assert str(QPath(("acme", "pnl", "revenue"))) == "acme.pnl.revenue"

    def test_root_rendering(self):
        assert str(QPath.ROOT) == "<root>"

    def test_floating_rendering(self):
        assert str(QPath.floating(3)) == "__floating__.v3"


class TestChild:
    def test_child_extends(self):
        p = QPath(("a",))
        assert p.child("b").segments == ("a", "b")

    def test_child_from_root(self):
        assert QPath.ROOT.child("acme").segments == ("acme",)

    def test_child_with_empty_seg_raises(self):
        with pytest.raises(ValueError):
            QPath(("a",)).child("")


class TestParent:
    def test_parent_walks_up(self):
        p = QPath(("a", "b", "c"))
        assert p.parent().segments == ("a", "b")

    def test_root_has_no_parent(self):
        with pytest.raises(ValueError):
            QPath.ROOT.parent()


class TestLeaf:
    def test_leaf_last_segment(self):
        assert QPath(("a", "b", "c")).leaf == "c"

    def test_root_has_no_leaf(self):
        with pytest.raises(ValueError):
            QPath.ROOT.leaf


class TestAnonymous:
    def test_anon_leaf(self):
        assert QPath(("s", "_anon_abc123")).is_anonymous

    def test_named_leaf(self):
        assert not QPath(("s", "revenue")).is_anonymous

    def test_root_not_anonymous(self):
        assert not QPath.ROOT.is_anonymous


class TestImmutability:
    def test_frozen(self):
        p = QPath(("a",))
        with pytest.raises(Exception):
            p.segments = ("b",)  # type: ignore[misc]

    def test_hashable(self):
        d = {QPath(("a", "b")): 1, QPath(("c",)): 2}
        assert d[QPath(("a", "b"))] == 1
