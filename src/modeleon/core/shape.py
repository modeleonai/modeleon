# SPDX-License-Identifier: Apache-2.0
"""Shape — what axes a Variable is laid out along.

A :class:`Shape` is a tuple of axes. An axis is just another
:class:`~modeleon.core.variable.Variable` that another Variable
declared via ``indexed_by=[…]``. There is no separate ``Axis`` /
``AxisProvider`` registry — the dep graph already records who
references whom; shape comes along for the ride.

A Variable's shape is ``Shape(var._indexed_by)``. Operator-built
Variables have ``_indexed_by`` set automatically by the broadcast
in :class:`~modeleon.core.variable_ops._VariableArithmetic`. Pure
inputs declare ``indexed_by=[axis_var]`` at construction.

Scalars have ``Shape(())``. Multi-axis Variables list their axes in
broadcast order.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any, Tuple

if TYPE_CHECKING:
    from .variable import Variable


@dataclass(frozen=True, slots=True)
class Shape:
    """Tuple of axes a Variable is laid out along.

    The axes themselves are :class:`Variable` instances — whatever the
    user passed to ``indexed_by=[…]``. Axis identity is by Python
    object identity, not by axis-label equality.
    """

    axes: Tuple[Any, ...] = field(default_factory=tuple)

    @property
    def is_scalar(self) -> bool:
        """``True`` when there are no axes — a single value."""
        return not self.axes

    @property
    def length(self) -> int:
        """Total cell count if every axis dimension is multiplied out.

        Each axis's length is the length of its ``_value`` (the list
        of axis labels). Scalar shapes return 1.
        """
        if not self.axes:
            return 1
        n = 1
        for ax in self.axes:
            v = getattr(ax, '_value', None)
            n *= len(v) if isinstance(v, list) else 1
        return n

    def __repr__(self) -> str:
        if not self.axes:
            return 'Shape(scalar)'
        names = []
        for ax in self.axes:
            # Prefer the python-binding name (set by adoption); fall
            # back to an explicit display_name; last resort, the
            # short class+hex form so the repr stays readable.
            n = (
                getattr(ax, '_python_name', None)
                or getattr(ax, '_display_name', None)
                or f"<{type(ax).__name__}#{id(ax) & 0xFFFF:04x}>"
            )
            names.append(str(n))
        return f"Shape({', '.join(names)})"


def shape_of(var: "Variable") -> Shape:
    """Read a Variable's shape off the dep graph.

    Returns ``Shape(var._indexed_by)``. Scalar Variables (no
    ``indexed_by`` declared, no operator broadcast) return
    ``Shape(())``.
    """
    return Shape(getattr(var, '_indexed_by', ()))
