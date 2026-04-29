# SPDX-License-Identifier: Apache-2.0
"""Axis — the dimension a list-shaped Variable extends along.

An axis names "what makes one cell different from the next" for a
list-valued Variable. The simplest case is positional —
``Variable([100, 110, 121])`` extends along an anonymous
:class:`PositionalAxis` whose only structure is its length.

Calendar-aware axes (with dates, fiscal-year offsets, period labels)
extend :class:`Axis` with richer semantics — see future modules.

Axes are paired with their owner via :class:`~modeleon.core.shape.Shape`.
A Variable's ``shape.axis`` returns the axis the Variable currently
binds to; mutating the owner replaces the axis atomically.
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class Axis:
    """Base axis — a single dimension a Variable extends along.

    Concrete axes carry at minimum a ``length`` (number of cells along
    the axis) and a ``name`` (used in error messages and for axis
    identity comparisons).
    """

    length: int
    name: str = "period"

    def __post_init__(self) -> None:
        if self.length < 0:
            raise ValueError(f"Axis length must be non-negative, got {self.length}.")


@dataclass(frozen=True, slots=True)
class PositionalAxis(Axis):
    """Anonymous positional axis — the default for plain list values.

    A Variable constructed with a list and no explicit axis (
    ``Variable([100, 200, 300])``) gets a ``PositionalAxis`` of the
    matching length. Two ``PositionalAxis`` instances of the same
    length are equal by structure but not necessarily by identity —
    Variables on separate ``PositionalAxis`` objects are not
    automatically alignable for cross-axis arithmetic.
    """

    name: str = "position"
