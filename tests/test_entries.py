# SPDX-License-Identifier: Apache-2.0
"""Which entered rows a derived track is computed from, read once per
model through the shared readings - and the same answer a fresh walk
gives, a chain deeper than the walk reads included."""

from __future__ import annotations

import pytest

import modeleon as mo
from modeleon.core import entries

P = 12


def _chain(depth: int):
    m = mo.Model("m", default_grain="month", default_start="2026-01", default_periods=P,
                 tracks=mo.Tracks("plan", "actual", blend=mo.blend(
                     given="actual", follow="plan", until="2026-06")))
    with m:
        m.inp = mo.MultiVariable("Inputs")
        m.inp.rev = mo.Variable(plan=[100.0] * P, actual=[110.0] * 6 + [0.0] * (P - 6))
        m.s = mo.MultiVariable("Section")
        prev = m.inp.rev
        for j in range(depth):
            setattr(m.s, f"l{j}", prev * 0.5 + m.inp.rev)
            prev = getattr(m.s, f"l{j}")
    return [getattr(m.s, f"l{j}")._value.as_dict() for j in range(depth)]


@pytest.mark.parametrize("depth", [entries.MAX_DEPTH - 5, entries.MAX_DEPTH + 5, 90])
def test_the_shared_readings_answer_as_a_fresh_walk(depth, monkeypatch):
    shared = _chain(depth)
    monkeypatch.setattr(entries, "readings_under", lambda node: None)
    assert _chain(depth) == shared


def test_a_chain_deeper_than_the_walk_is_read_once_per_line(monkeypatch):
    # Past MAX_DEPTH the shared walk gives up and the line is read afresh;
    # what it learnt on the way down tells the next line to give up at
    # once, so the shared walk reads a line or two per line joined, not
    # MAX_DEPTH of them.
    calls = []
    read = entries._Walk._line

    def counting(self, var, depth, own):
        if self.kept is not None:
            calls.append(depth)
        return read(self, var, depth, own)

    monkeypatch.setattr(entries._Walk, "_line", counting)
    n = 3 * entries.MAX_DEPTH
    _chain(n)
    assert len(calls) < 5 * n
