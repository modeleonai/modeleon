# SPDX-License-Identifier: Apache-2.0
"""Architectural guardrails.

These tests enforce invariants that the codebase already follows but
would be easy to break accidentally in a future PR. Each test here
documents a design decision in ADR-010.
"""

from __future__ import annotations

import ast
from pathlib import Path


CORE_DIR = Path(__file__).parent.parent / "src" / "modeleon" / "core"


class TestCoreIsRendererNeutral:
    """``core/*.py`` must not import from ``compile/`` at module scope.

    The engine splits into ``core/`` (renderer-neutral primitives: Variable,
    MultiVariable, QPath, Expr, Unit) and ``compile/`` (renderers: Excel,
    JSON, Walker). Core must stay pure so any future renderer can live
    alongside Excel/JSON without touching Variable or Expr.

    Runtime-dispatched imports inside methods are fine (and used by
    ``MultiVariableBase.to_excel``). Module-scope imports are not.
    """

    def test_no_compile_imports_at_module_scope(self):
        offenders = []
        for py in sorted(CORE_DIR.rglob("*.py")):
            tree = ast.parse(py.read_text(encoding="utf-8"), filename=str(py))
            for node in tree.body:  # only module-level nodes
                if isinstance(node, ast.ImportFrom):
                    mod = node.module or ""
                    if mod.startswith("modeleon.compile") or mod.startswith("..compile") or mod == "compile":
                        offenders.append(f"{py.relative_to(CORE_DIR)}: from {mod}")
                elif isinstance(node, ast.Import):
                    for alias in node.names:
                        if alias.name.startswith("modeleon.compile"):
                            offenders.append(f"{py.relative_to(CORE_DIR)}: import {alias.name}")
        assert not offenders, (
            "core/ modules must not import from compile/ at module scope "
            "(see ADR-010). Move the import inside the method that needs "
            "it for runtime dispatch. Offenders:\n  " + "\n  ".join(offenders)
        )
