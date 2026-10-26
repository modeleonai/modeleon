# SPDX-License-Identifier: Apache-2.0
"""Passthrough hygiene: the Excel renderer writes ``=NAME(args)`` for every
name the function registry passes through, so every such name is a
function an author can call — exported from ``modeleon`` and
``modeleon.functions``, the same object in both — unless it is registered
``bare``: written as given, with no helper of its own (``SUMPRODUCT``,
what ``SUM(a * b)`` over two rows is written as; ``NA()``, the not-found
value an ``XLOOKUP`` is written with). A bare name that gained a helper
is a stale flag, and a registered name without a helper is a formula the
renderer can write but an author cannot reach.
"""

from __future__ import annotations

import pytest

import modeleon as mo
from modeleon import functions
from modeleon.functions.registry import names, passthrough_names, spec


def _excel_cased(exported):
    return {n for n in exported if n.isupper()}


@pytest.mark.parametrize("name", sorted(passthrough_names()))
def test_every_passthrough_name_has_an_exported_helper_unless_bare(name):
    if spec(name).bare:
        assert not hasattr(functions, name) and not hasattr(mo, name), (
            f"{name} has a helper now: drop bare=True from its register() line")
        return
    assert name in functions.__all__ and name in mo.__all__, name
    assert callable(getattr(functions, name)), name
    assert getattr(mo, name) is getattr(functions, name), name


def test_a_name_written_some_other_way_has_a_helper_too():
    # CONCAT is written as the ``&`` operator.
    others = names() - passthrough_names()
    assert others == {"CONCAT"}
    for name in others:
        assert not spec(name).bare and name in mo.__all__, name


def test_the_package_exports_exactly_the_registered_helpers():
    registered = {n for n in names() if not spec(n).bare}
    assert _excel_cased(functions.__all__) == registered
    assert _excel_cased(mo.__all__) == registered


def test_the_bare_names_are_the_ones_without_a_helper():
    assert {n for n in names() if spec(n).bare} == {"SUMPRODUCT", "NA"}
