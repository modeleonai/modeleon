# SPDX-License-Identifier: Apache-2.0
"""Modeleon-specific exception + warning types.

Named categories let callers distinguish model-shape errors from
unrelated ``ValueError``s / ``UserWarning``s, and give users a clean
search term when a traceback or a deprecation-style message surfaces.
"""

import warnings
from typing import List


class CircularDependencyError(ValueError):
    """Raised when Variable dependencies form a cycle.

    Cycles make it impossible to compute values or generate Excel formulas
    in a deterministic order. The exception's ``cycle`` attribute lists the
    variable IDs forming the cycle (in traversal order, with
    ``cycle[0] == cycle[-1]``).
    """

    def __init__(self, cycle: List[str]):
        self.cycle = cycle
        path = " -> ".join(cycle)
        super().__init__(f"Circular dependency detected: {path}")


class CrossScopeReferenceWarning(UserWarning):
    """Emitted when a formula references a Variable outside the current
    emission's layout.

    Two scenarios trigger this:

    1. **Cross-model reference** — a formula in model ``M2`` references
       a Variable that lives in a different model ``M1``. The Excel
       renderer has no cell address for ``M1``'s Variable in
       ``M2``'s workbook, so it inlines the *value* as a literal.
       The resulting cell is a dead number, not a live link.
    2. **Floating Variable** — a Variable was created outside any
       ``with MultiVariable(...):`` block and then used in a formula
       that is emitted. Same outcome: value inlined, link lost.

    Both are almost always a user mistake — the intent was a live
    reference. The warning fires at emission time with the name of
    the orphan Variable so the fix is obvious (put it inside a
    ``with:`` block, or move the formula to the model that owns it).
    """


warnings.filterwarnings('always', category=CrossScopeReferenceWarning)
