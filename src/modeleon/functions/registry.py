# SPDX-License-Identifier: Apache-2.0
"""The function registry — what the engine knows about each function name.

One entry per name a :class:`~modeleon.core.expr.FuncCall` can carry: how
the call is shaped, what the workbook stores it as, and whether a loop
container may run it one period at a time. The Excel renderer and the
loop container read the registry; a function registers itself from its
own module, next to its definition::

    from .registry import register

    def IFERROR(value, fallback):
        ...

    register("IFERROR", shape="elementwise")

Shapes:

- ``'elementwise'`` — one call per period: a series operand gives a
  series, the workbook writes this period's cells, a loop container runs
  it period by period (``IF``, ``ABS``, ``ROUND``, …).
- ``'reduction'`` — reads whole series and gives one number; the workbook
  writes a list operand as a range (``SUM(row)`` is ``=SUM(B2:F2)``). A
  ``per_period_arity`` makes a call over that many values or more, a
  series among them, one call per period instead: ``SUM(a, b)`` is a
  subtotal, ``MIN`` / ``MAX`` over two operands are a cap or a floor,
  ``MEDIAN(a, b)`` the median of each period's two cells.
- ``'range'`` — fixed arguments, some of them ranges. ``range_args`` names
  their positions; None means every list argument is a range and every
  other one a cell (``NPV(rate, flows)``, ``SUMPRODUCT``). A
  ``per_period_arity`` makes a call whose ranges list that many lines or
  more, a series among them, one call per period over this period's
  cells: ``PERCENTILE_INC([a, b, c], k)``.
- ``'scalar'`` — a plain call over its arguments, written like an
  element-wise one, that a loop container does not run period by period
  (``YEARFRAC``, ``DAYS``).

A name the workbook stores with a prefix (``_xlfn.`` for the functions
Excel added after 2007) registers it as ``file_prefix``; a name Excel
spells differently registers ``excel``. A name without a helper of its
own, which the renderer only writes as given, is ``bare``.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal, Sequence

Shape = Literal['elementwise', 'reduction', 'range', 'scalar']

SHAPES: tuple[str, ...] = ('elementwise', 'reduction', 'range', 'scalar')


@dataclass(frozen=True)
class FunctionSpec:
    """How one function name is shaped; see the module docstring."""

    name: str
    shape: Shape
    excel: str
    file_prefix: str = ''
    range_args: tuple[int, ...] | None = None
    per_period_arity: int | None = None
    passthrough: bool = True
    bare: bool = False


_SPECS: dict[str, FunctionSpec] = {}


def register(
    name: str,
    *,
    shape: Shape,
    excel: str | None = None,
    file_prefix: str = '',
    range_args: Sequence[int] | None = None,
    per_period_arity: int | None = None,
    passthrough: bool = True,
    bare: bool = False,
) -> FunctionSpec:
    """Register ``name`` (upper-case, as the call is written) with its shape.

    ``excel`` is the name the workbook writes when it differs;
    ``file_prefix`` what the file stores before it (``'_xlfn.'``);
    ``range_args`` the positions written as ranges (``'range'`` shape only;
    None = every list argument); ``per_period_arity`` the count of values
    — a ``'reduction'``'s operands, the lines a ``'range'`` function's
    ranges list — from which a call with a series among them is one call
    per period; ``passthrough``
    False when the renderer writes the call some other way (``CONCAT`` is
    the ``&`` operator); ``bare`` when the name has no helper of its own. Registering
    the same spec twice is one entry; a different spec for a registered
    name is refused.
    """
    if not name or name != name.upper() or not name.isidentifier():
        raise ValueError(
            f"a function registers under its upper-case identifier, as the call "
            f"is written (SUM, IFERROR): got {name!r}."
        )
    if shape not in SHAPES:
        raise ValueError(
            f"{name}: shape must be one of {', '.join(SHAPES)}; got {shape!r}."
        )
    positions: tuple[int, ...] | None = None
    if range_args is not None:
        if shape != 'range':
            raise ValueError(
                f"{name}: range_args names the range positions of a 'range' "
                f"function; a '{shape}' function has none."
            )
        positions = tuple(range_args)
        if not positions or any(not isinstance(p, int) or p < 0 for p in positions):
            raise ValueError(
                f"{name}: range_args are argument positions, 0-based: got {range_args!r}."
            )
    if per_period_arity is not None:
        if shape not in ('reduction', 'range'):
            raise ValueError(
                f"{name}: per_period_arity is how many values make a 'reduction' or a "
                f"'range' function one call per period; a '{shape}' function has none."
            )
        if per_period_arity < 2:
            raise ValueError(
                f"{name}: per_period_arity is at least 2 (one value is the "
                f"reduction itself); got {per_period_arity}."
            )
    new = FunctionSpec(
        name=name, shape=shape, excel=excel or name, file_prefix=file_prefix,
        range_args=positions, per_period_arity=per_period_arity,
        passthrough=passthrough, bare=bare,
    )
    old = _SPECS.get(name)
    if old is not None and old != new:
        raise ValueError(
            f"{name} is already registered as {old}; a second, different "
            f"registration ({new}) would change how every {name} call is written."
        )
    _SPECS[name] = new
    return new


# ─── Queries ────────────────────────────────────────────────────────


def spec(name: str) -> FunctionSpec | None:
    """The spec registered under ``name``, or None."""
    return _SPECS.get(name)


def names() -> frozenset[str]:
    """Every registered name."""
    return frozenset(_SPECS)


def is_passthrough(name: str) -> bool:
    """Does the Excel renderer write ``=NAME(args)`` for this name?"""
    s = _SPECS.get(name)
    return s is not None and s.passthrough


def is_range_taking(name: str) -> bool:
    """Is a list argument of this name written as a range (a reduction or
    a range function)?"""
    s = _SPECS.get(name)
    return s is not None and s.shape in ('reduction', 'range')


def range_positions(name: str) -> tuple[int, ...] | None:
    """The argument positions written as ranges, or None for every list
    argument (and for a name that takes no range)."""
    s = _SPECS.get(name)
    return s.range_args if s is not None else None


def is_elementwise(name: str) -> bool:
    """Is this name one call per period?"""
    s = _SPECS.get(name)
    return s is not None and s.shape == 'elementwise'


def per_period_arity(name: str) -> int | None:
    """The count of values from which a call with a series among them is
    one call per period (``MIN`` / ``MAX``, ``MEDIAN``, ``PERCENTILE_INC``:
    2), or None."""
    s = _SPECS.get(name)
    return s.per_period_arity if s is not None else None


def admits_in_loop(name: str, n_args: int) -> bool:
    """May a loop container run a call of ``name`` over ``n_args`` values
    one period at a time? The element-wise names, and a reduction or a
    range function over at least its ``per_period_arity`` values — a
    reduction's operands, the lines its ranges list."""
    s = _SPECS.get(name)
    if s is None:
        return False
    if s.shape == 'elementwise':
        return True
    return s.per_period_arity is not None and n_args >= s.per_period_arity


def file_prefix(name: str) -> str:
    """What the workbook file stores before the name (``''`` or ``'_xlfn.'``)."""
    s = _SPECS.get(name)
    return s.file_prefix if s is not None else ''


def excel_name(name: str) -> str:
    """The name the workbook writes — Excel's spelling when it differs."""
    s = _SPECS.get(name)
    return s.excel if s is not None else name


# ─── Views, for the tests that pin them ─────────────────────────────


def passthrough_names() -> frozenset[str]:
    """Every name the Excel renderer writes as ``=NAME(args)``."""
    return frozenset(n for n, s in _SPECS.items() if s.passthrough)


def range_taking_names() -> frozenset[str]:
    """Every name whose list arguments are written as ranges."""
    return frozenset(n for n, s in _SPECS.items() if s.shape in ('reduction', 'range'))


def elementwise_names() -> frozenset[str]:
    """Every name that is one call per period."""
    return frozenset(n for n, s in _SPECS.items() if s.shape == 'elementwise')


def file_prefixes() -> dict[str, str]:
    """The names the file stores with a prefix, and the prefix of each."""
    return {n: s.file_prefix for n, s in _SPECS.items() if s.file_prefix}
