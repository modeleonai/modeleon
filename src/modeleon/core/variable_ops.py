# SPDX-License-Identifier: Apache-2.0
"""Arithmetic and comparison mixin for :class:`Variable`.

Split out from ``variable.py`` to keep the main file focused on
identity, construction, and representation. This module owns:

- Broadcasting (scalar/list, list/list with length-1 fallback)
- Safe-op semantics (Excel-style ``#DIV/0!`` and error propagation)
- Key alignment and propagation for keyed (dict-sourced) Variables
- Unit propagation through arithmetic
- The shared ``_binary_op`` / ``_comparison_op`` scaffolds
- All binary, unary, and comparison dunders

``_VariableArithmetic`` is a **mixin** — it assumes ``self`` is a
:class:`Variable` and relies on ``self._value``, ``self._keys``,
``self._as_operand``, plus the ``expr`` property setter (which updates
``_expr`` and ``_dependency_refs`` together). Variable inherits from
this class so the operator surface is available everywhere Variable
is, while the file on disk stays coherent.

Lazy imports for :class:`Variable` inside each method avoid the
circular import that would arise if we imported at module scope
(``variable.py`` imports this module; this module needs Variable to
construct results).
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from .expr import BinOp, Compare, Literal, UnaryOp, VarRef

if TYPE_CHECKING:
    from .variable import Variable


class _VariableArithmetic:
    """Operator-overloading mixin for :class:`Variable`."""

    # ─── Broadcasting ───────────────────────────────────────────
    #
    # COMPUTE CHOKEPOINT. Every arithmetic value produced by the engine
    # flows through this function (operator overloads in
    # ``_VariableArithmetic`` call it). Do not add parallel compute
    # paths — keep all value arithmetic routed through here.

    @staticmethod
    def _broadcast_operation(left, right, op, safe_divide: bool = False):
        """Element-wise op with NumPy-style broadcasting.

        Scalar/scalar → scalar. List/scalar → element-wise list.
        Scalar/list → element-wise list. List/list → element-wise;
        length-1 operands broadcast over the longer list.

        Excel-style error markers (``'#DIV/0!'``, ``'#NAME?'``, …)
        propagate through operations. With ``safe_divide=True``, division
        by zero yields ``'#DIV/0!'`` instead of raising.
        """
        def _is_excel_error(v) -> bool:
            return isinstance(v, str) and v.startswith('#')

        def safe_op(a, b):
            if _is_excel_error(a):
                return a
            if _is_excel_error(b):
                return b
            if safe_divide and b == 0:
                return '#DIV/0!'
            try:
                return op(a, b)
            except Exception:
                return '#VALUE!'

        left_is_list = isinstance(left, list)
        right_is_list = isinstance(right, list)

        if not left_is_list and not right_is_list:
            return safe_op(left, right)
        elif left_is_list and not right_is_list:
            return [safe_op(x, right) for x in left]
        elif not left_is_list and right_is_list:
            return [safe_op(left, x) for x in right]
        else:
            if len(left) != len(right):
                if len(left) == 1:
                    return [safe_op(left[0], x) for x in right]
                elif len(right) == 1:
                    return [safe_op(x, right[0]) for x in left]
                raise ValueError(
                    f"Cannot broadcast lists of different lengths: "
                    f"{len(left)} vs {len(right)}. "
                    f"Element-wise ops need equal lengths, or one side of "
                    f"length 1 which broadcasts over the other. "
                    f"Common fix: align the two list-valued inputs to the "
                    f"same period count."
                )
            return [safe_op(a, b) for a, b in zip(left, right)]

    @staticmethod
    def _align_keyed_values(left, right):
        """Reorder right's list to match left's key order when both are keyed.

        If only one side is keyed, or keys match already, pass through.
        Raises if key sets disagree — silently reordering a mismatch would
        produce a silently wrong result.
        """
        from .variable import Variable
        left_val = left._value if isinstance(left, Variable) else left
        right_val = right._value if isinstance(right, Variable) else right

        left_keys = getattr(left, '_keys', None) if isinstance(left, Variable) else None
        right_keys = getattr(right, '_keys', None) if isinstance(right, Variable) else None

        if (left_keys is not None and right_keys is not None
                and isinstance(right_val, list)):
            if set(left_keys) != set(right_keys):
                raise ValueError(
                    f"Cannot align Variables with different key sets: "
                    f"{left_keys} vs {right_keys}. "
                    f"Element-wise ops on keyed Variables need matching keys "
                    f"so each value has a partner. Either pick the same keys "
                    f"for both sides, or drop keys on one side (pass through "
                    f"``.value``) and use positional broadcasting instead."
                )
            if left_keys != right_keys:
                key_to_idx = {k: i for i, k in enumerate(right_keys)}
                right_val = [right_val[key_to_idx[k]] for k in left_keys]

        return left_val, right_val

    @staticmethod
    def _propagate_keys(left, right):
        """Result keys come from whichever operand has them; left wins if both."""
        from .variable import Variable
        left_keys = getattr(left, '_keys', None) if isinstance(left, Variable) else None
        right_keys = getattr(right, '_keys', None) if isinstance(right, Variable) else None
        if left_keys is not None:
            return list(left_keys)
        if right_keys is not None:
            return list(right_keys)
        return None

    # ─── Units ──────────────────────────────────────────────────

    @staticmethod
    def _unit_mul(left_unit, right_unit):
        """Resulting unit for multiplication."""
        if left_unit is None and right_unit is None:
            return None
        from .unit import Unit
        lu = left_unit if left_unit is not None else Unit.dimensionless()
        ru = right_unit if right_unit is not None else Unit.dimensionless()
        result = lu * ru
        return result if result._unit_components else None

    @staticmethod
    def _unit_div(left_unit, right_unit):
        """Resulting unit for division."""
        if left_unit is None and right_unit is None:
            return None
        from .unit import Unit
        lu = left_unit if left_unit is not None else Unit.dimensionless()
        ru = right_unit if right_unit is not None else Unit.dimensionless()
        result = lu / ru
        return result if result._unit_components else None

    @staticmethod
    def _unit_add(left_unit, right_unit):
        """Resulting unit for addition/subtraction.

        Both-None → None. One-sided (scalar + unitful) inherits the
        unit. Both-sides-present must agree — otherwise raise, because
        silently returning one side would mask a real modeling bug.
        """
        if left_unit is None and right_unit is None:
            return None
        if left_unit is not None and right_unit is not None:
            if left_unit != right_unit:
                raise ValueError(
                    f"Cannot add/subtract values with incompatible units: "
                    f"{left_unit} and {right_unit}. Either convert one side "
                    f"to the other's unit before the operation, or drop the "
                    f"unit on one side if the arithmetic is intentional."
                )
            return left_unit
        return left_unit if left_unit is not None else right_unit

    @staticmethod
    def _unit_pow(unit, exponent):
        """Resulting unit for ``x ** n``.

        Only defined for integer exponents (``hr ** 2 → hr^2``). Non-int
        exponents drop the unit — Modeleon doesn't track fractional-power
        units because they rarely appear in financial modeling and the
        rendering gets ugly.
        """
        if unit is None or not isinstance(exponent, int):
            return None
        result = unit ** exponent
        # ``hr ** 0`` → dimensionless; collapse the empty-components
        # Unit to ``None`` to match the ``mul`` / ``div`` convention.
        return result if result._unit_components else None

    def _get_unit(self):
        """Get unit from self, or None for scalars."""
        return getattr(self, '_unit', None)

    # ─── Binary operators ──────────────────────────────────────

    def _binary_op(self, other, op_symbol: str, op_func, unit_fn,
                   safe_divide: bool = False, parenthesize: bool = False,
                   reversed: bool = False) -> "Variable":
        """Shared implementation for all binary arithmetic operators.

        Args:
            other: Right operand — another Variable, a scalar, or a list.
            op_symbol: Text inserted between operands in the formula string.
            op_func: Element-wise function applied for value computation.
            unit_fn: ``(left_unit, right_unit) -> result_unit`` or ``None``.
                     Pass ``None`` for operators that don't track units
                     (``**``, ``%``, and ``//`` when scalar is on the right).
            safe_divide: Route through the safe-division path so 1/0 becomes
                         ``"#DIV/0!"`` instead of raising.
            parenthesize: Wrap compound-unnamed-formula operands in parens
                          to preserve precedence (``(a+b) / (c+d)``).
            reversed: ``True`` for reflected ops (``__rsub__``, etc.).
                      In the formula string, operand order flips; dependency
                      refs still record ``[self, other]`` in that order.
        """
        from .variable import Variable

        self_u = self._get_unit()

        if isinstance(other, Variable):
            left, right = (other, self) if reversed else (self, other)
            expr_node = BinOp(op_symbol, left._as_operand(parenthesize),
                              right._as_operand(parenthesize))
            result = Variable()
            result._set_expr(expr_node)
            lv, rv = self._align_keyed_values(left, right)
            result._value = self._broadcast_operation(lv, rv, op_func, safe_divide=safe_divide)
            if unit_fn is not None:
                result._unit = unit_fn(left._get_unit(), right._get_unit())
        else:
            if reversed:
                expr_node = BinOp(op_symbol, Literal(other), self._as_operand(parenthesize))
                val = self._broadcast_operation(other, self._value, op_func, safe_divide=safe_divide)
            else:
                expr_node = BinOp(op_symbol, self._as_operand(parenthesize), Literal(other))
                val = self._broadcast_operation(self._value, other, op_func, safe_divide=safe_divide)
            result = Variable()
            result._set_expr(expr_node)
            result._value = val
            if unit_fn is not None:
                # Scalars are treated as dimensionless when computing the output unit
                result._unit = unit_fn(None, self_u) if reversed else unit_fn(self_u, None)

        if isinstance(result._value, list) and len(result._value) > 1:
            result.var_type = "list"
        k_left, k_right = (other, self) if reversed else (self, other)
        result._keys = self._propagate_keys(k_left, k_right)
        return result

    # ─── Binary dunders ────────────────────────────────────────

    def __add__(self, other) -> "Variable":
        return self._binary_op(other, "+", lambda a, b: a + b, self._unit_add)

    def __radd__(self, other) -> "Variable":
        return self.__add__(other)

    def __sub__(self, other) -> "Variable":
        return self._binary_op(other, "-", lambda a, b: a - b, self._unit_add)

    def __rsub__(self, other) -> "Variable":
        return self._binary_op(other, "-", lambda a, b: a - b, self._unit_add, reversed=True)

    def __mul__(self, other) -> "Variable":
        return self._binary_op(other, "*", lambda a, b: a * b, self._unit_mul)

    def __rmul__(self, other) -> "Variable":
        return self.__mul__(other)

    def __truediv__(self, other) -> "Variable":
        return self._binary_op(other, "/", lambda a, b: a / b, self._unit_div,
                               safe_divide=True, parenthesize=True)

    def __rtruediv__(self, other) -> "Variable":
        return self._binary_op(other, "/", lambda a, b: a / b, self._unit_div,
                               safe_divide=True, parenthesize=True, reversed=True)

    def __pow__(self, other) -> "Variable":
        result = self._binary_op(other, "**", lambda a, b: a ** b, None)
        # Scalar integer exponent → raise the unit to that power.
        # Variable-on-the-right exponents leave the unit as None.
        from .variable import Variable
        if not isinstance(other, Variable):
            result._unit = self._unit_pow(self._get_unit(), other)
        return result

    def __rpow__(self, other) -> "Variable":
        return self._binary_op(other, "**", lambda a, b: a ** b, None, reversed=True)

    def __mod__(self, other) -> "Variable":
        return self._binary_op(other, "%", lambda a, b: a % b, None)

    def __rmod__(self, other) -> "Variable":
        return self._binary_op(other, "%", lambda a, b: a % b, None, reversed=True)

    def __floordiv__(self, other) -> "Variable":
        return self._binary_op(other, "//", lambda a, b: a // b, self._unit_div)

    def __rfloordiv__(self, other) -> "Variable":
        return self._binary_op(other, "//", lambda a, b: a // b, None, reversed=True)

    # ─── Unary dunders ─────────────────────────────────────────

    def __neg__(self) -> "Variable":
        """-a (unary negation)"""
        from .variable import Variable
        result = Variable()
        result._set_expr(UnaryOp("-", self._as_operand(parenthesize=True)))
        if isinstance(self._value, list):
            result._value = [-x for x in self._value]
        else:
            result._value = -self._value
        result._unit = self._get_unit()
        result._keys = list(self._keys) if self._keys is not None else None
        if isinstance(result._value, list) and len(result._value) > 1:
            result.var_type = 'list'
        return result

    def __pos__(self) -> "Variable":
        """+a is a no-op — return self."""
        return self

    # ─── Comparisons ───────────────────────────────────────────
    #
    # COMPUTE CHOKEPOINT for comparisons. Parallel to
    # ``_broadcast_operation`` for arithmetic — do not add parallel
    # compute paths.

    def _comparison_op(self, other, op_symbol: str, op_func) -> "Variable":
        """Shared implementation for all comparison operators."""
        from .variable import Variable

        if isinstance(other, Variable):
            result = Variable()
            result._set_expr(Compare(op_symbol, VarRef(self), VarRef(other)))
            lv, rv = self._align_keyed_values(self, other)
            result._value = self._broadcast_operation(lv, rv, op_func)
        else:
            result = Variable()
            result._set_expr(Compare(op_symbol, VarRef(self), Literal(other)))
            result._value = self._broadcast_operation(self._value, other, op_func)

        if isinstance(result._value, list) and len(result._value) > 1:
            result.var_type = 'list'
        result._keys = self._propagate_keys(self, other)
        return result

    # ─── Comparison dunders ────────────────────────────────────

    def __eq__(self, other) -> "Variable":
        return self._comparison_op(other, '==', lambda a, b: a == b)

    def __ne__(self, other) -> "Variable":
        return self._comparison_op(other, '!=', lambda a, b: a != b)

    def __lt__(self, other) -> "Variable":
        return self._comparison_op(other, '<', lambda a, b: a < b)

    def __le__(self, other) -> "Variable":
        return self._comparison_op(other, '<=', lambda a, b: a <= b)

    def __gt__(self, other) -> "Variable":
        return self._comparison_op(other, '>', lambda a, b: a > b)

    def __ge__(self, other) -> "Variable":
        return self._comparison_op(other, '>=', lambda a, b: a >= b)
