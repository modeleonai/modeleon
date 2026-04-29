# SPDX-License-Identifier: Apache-2.0
"""Base — minimal identity surface shared by every named node in a model.

Carries the bare minimum that any named thing in a model needs:
qualified-path identity, a Python-binding name, and an optional
user-facing display label. Variable and MultiVariable both inherit
from Base via :class:`~modeleon.core.component.Component`.

Concrete subclasses must implement :attr:`path` — its resolution rules
differ between leaves (Variable) and containers (MultiVariable).
"""

from __future__ import annotations

from typing import Optional

from .qpath import QPath


class Base:
    """Common minimum: identity (``_qualified_id``), python_name, display label.

    Strictly minimal — anything composition-related (owner / parent
    references, adoption machinery) lives on
    :class:`~modeleon.core.component.Component`. Subclass directly when
    you need identity + a display label without the model-tree
    composition surface.
    """

    def __init__(self, display_name: Optional[str] = None):
        # Qualified-path identity. Starts as ``None``; the ``.path``
        # property on the concrete subclass synthesizes a
        # floating-namespace path from ``id(self)`` until adoption
        # crystallizes a rooted path.
        self._qualified_id: Optional[QPath] = None
        # The Python identifier this node is bound to as a component
        # of its parent. Set by adoption (``parent.x = node`` writes
        # ``node._python_name = "x"`` directly) — there is no public
        # constructor kwarg for this field. Top-level naming goes
        # through :class:`~modeleon.core.model.Model` instead.
        self._python_name: Optional[str] = None
        # Explicit user-facing label, set via constructor. The
        # ``display_name`` property on each subclass falls back to a
        # humanized identifier when this is None.
        self._display_name: Optional[str] = (
            None if display_name is None else str(display_name)
        )

    @property
    def python_name(self) -> Optional[str]:
        """The Python identifier this node is bound to as a component.

        Read-only. Set internally by adoption (``parent.x = node``
        triggers the parent's ``_register_component`` which writes
        ``node._python_name = "x"`` directly) or by
        :class:`~modeleon.core.model.Model`'s constructor for the
        root. Returns ``None`` for unattached nodes — there is no
        frame-introspection auto-discovery.
        """
        return self._python_name

    @property
    def path(self) -> QPath:
        """Qualified-path identity. Concrete subclasses implement this.

        An un-crystallized node returns a floating-namespace path keyed
        by ``id(self)``; once adoption attaches the node to a parent
        with a rooted path, the subclass crystallizes a permanent
        :class:`QPath` and caches it in ``_qualified_id``.
        """
        raise NotImplementedError(
            f"{type(self).__name__} must implement the .path property."
        )

    @property
    def id(self) -> str:
        """Dotted-string form of :attr:`path` — the canonical string id."""
        return str(self.path)
