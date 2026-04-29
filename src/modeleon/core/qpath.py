# SPDX-License-Identifier: Apache-2.0
"""Qualified path identity.

A ``QPath`` is a tuple of name segments from the model root to a leaf,
e.g. ``("acme", "pnl", "revenue")``. The empty tuple is the model root.
Nodes not yet adopted into a rooted tree live in the ``__floating__``
namespace, e.g. ``("__floating__", "v3")`` — still a valid ``QPath``,
still the single source of identity.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import ClassVar


_FLOATING_SEGMENT = "__floating__"


@dataclass(frozen=True, slots=True)
class QPath:
    """Qualified path identity: a tuple of name segments.

    ``QPath.ROOT`` is the empty path (the model root itself).
    Unrooted nodes carry a path under ``("__floating__", …)`` — distinct
    per object via a monotonic sequence number. Every MV/Variable has a
    path from construction; adoption re-roots the node into a Model
    subtree and replaces the floating path with the qualified one.
    """

    segments: tuple[str, ...] = ()

    ROOT: ClassVar["QPath"]

    def __str__(self) -> str:
        if not self.segments:
            return "<root>"
        return ".".join(self.segments)

    def __repr__(self) -> str:
        return f"QPath({self!s})"

    @property
    def is_floating(self) -> bool:
        return bool(self.segments) and self.segments[0] == _FLOATING_SEGMENT

    @property
    def is_root(self) -> bool:
        return not self.segments

    @property
    def is_anonymous(self) -> bool:
        return bool(self.segments) and self.segments[-1].startswith("_anon_")

    @property
    def leaf(self) -> str:
        if not self.segments:
            raise ValueError("root has no leaf")
        return self.segments[-1]

    def child(self, seg: str) -> "QPath":
        if not seg:
            raise ValueError("segment must be non-empty")
        return QPath(self.segments + (seg,))

    def parent(self) -> "QPath":
        if not self.segments:
            raise ValueError("root has no parent")
        return QPath(self.segments[:-1])

    @staticmethod
    def floating(token: int, kind: str = "v") -> "QPath":
        """Synthesize a floating-namespace path for an unrooted node.

        ``kind`` is a short discriminator: ``"v"`` for Variables, ``"m"``
        for MultiVariables. ``token`` is typically ``id(obj)`` — the
        Python object address, unique during the object's lifetime.
        Rendered in hex for compactness.
        """
        return QPath((_FLOATING_SEGMENT, f"{kind}{token:x}"))


QPath.ROOT = QPath(())
