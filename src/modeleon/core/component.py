# SPDX-License-Identifier: Apache-2.0
"""Component — typing layer for nodes that live in the model tree.

Variable and MultiVariable both inherit from Component. Subclasses of
:class:`~modeleon.core.base.Base` that exist outside the model tree
(metadata-only objects with identity but no owner/parent linkage)
inherit from Base directly.

Component owns the small, identical surface that every model-tree node
shares: ``excel_props`` storage and validation, the
``display_name`` setter, and ``set_*`` chaining helpers. Per-class
machinery (path resolution, ``__repr__``, the ``display_name``
fallback chain) stays on Variable and MultiVariable.
"""

from __future__ import annotations

from typing import Any, Dict, FrozenSet, Optional, Self

from .base import Base


class Component(Base):
    """Marker base for nodes that participate in the model tree.

    Variable and MultiVariable inherit from Component. Anything that is
    addressable inside a model-tree walk should be a Component;
    metadata-only objects with identity (but no model-tree presence)
    should subclass :class:`Base` directly.
    """

    # Subclasses override with their own valid Excel-renderer keys.
    # Empty by default so a bare Component carries no styling surface.
    _EXCEL_PROP_KEYS: FrozenSet[str] = frozenset()

    def __init__(
        self,
        display_name: Optional[str] = None,
        excel_props: Optional[Dict[str, Any]] = None,
    ):
        super().__init__(display_name=display_name)
        self._excel_props: Dict[str, Any] = self._validate_excel_props(excel_props)

    @classmethod
    def _validate_excel_props(
        cls, excel_props: Optional[Dict[str, Any]]
    ) -> Dict[str, Any]:
        """Copy and validate ``excel_props`` against ``cls._EXCEL_PROP_KEYS``.

        Unknown keys raise :class:`TypeError` so typos surface at
        construction time rather than vanishing into a styling no-op.
        """
        props: Dict[str, Any] = dict(excel_props) if excel_props else {}
        unknown = set(props) - cls._EXCEL_PROP_KEYS
        if unknown:
            raise TypeError(
                f"{cls.__name__}(excel_props=...) got unknown key(s): "
                f"{sorted(unknown)}. "
                f"Known keys: {sorted(cls._EXCEL_PROP_KEYS)}."
            )
        return props

    @property
    def excel_props(self) -> Dict[str, Any]:
        """Excel-renderer hints attached to this node.

        Returns the live internal dict — prefer :meth:`set_style` (or
        the ``excel_props=`` constructor kwarg) for additions; direct
        mutation works but bypasses validation.
        """
        return self._excel_props

    def set_style(self, **kwargs: Any) -> Self:
        """Update Excel-renderer styling hints (bold, bg, font_color,
        number_format, …). Unknown keys are silently ignored — use the
        ``excel_props=`` constructor kwarg to get strict validation.
        """
        for k, v in kwargs.items():
            if k in self._EXCEL_PROP_KEYS:
                self._excel_props[k] = v
        return self

    def set_display_name(self, display_name: str) -> Self:
        """Set the display label and return ``self`` for chaining."""
        self._display_name = display_name
        return self

    @property
    def code(self) -> str:
        """Render this node as a runnable Modeleon DSL snippet.

        Sibling to ``_repr_html_`` and ``to_excel`` — same idea, target
        is Python source. For a :class:`~modeleon.Variable` this is one
        line (``name = mo.Variable(...)``). For a
        :class:`~modeleon.MultiVariable` or :class:`~modeleon.Model`,
        the header line plus every descendant qualified under this
        node's ``python_name``.

        Variable formulas keep their original form when every
        reference resolves inside the rendered subtree; otherwise
        they fall back to the computed value so the snippet runs
        standalone (Excel-style values fallback).

        Lossy w.r.t. user-written ``.py`` — only the DSL subset
        round-trips. Code between Variable declarations (imports,
        helpers, comments) lives in source bytes and is invisible
        to runtime.
        """
        from ..display.code import render_code
        return render_code(self)
