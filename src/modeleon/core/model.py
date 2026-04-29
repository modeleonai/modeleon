# SPDX-License-Identifier: Apache-2.0
"""Model — the top-level container for a model tree.

A :class:`Model` is the canonical entry point: every model has a
single root, and its name is the only explicit naming gesture you
need. Children adopted under it (``acme.pnl = mo.MultiVariable(...)``)
crystallize their paths automatically — nothing below the root needs
to be named by hand.

::

    acme = mo.Model('acme')
    acme.pnl = mo.MultiVariable(excel_props={'tab': True})
    acme.pnl.revenue = mo.Variable(1_000_000)
    acme.pnl.cogs = acme.pnl.revenue * 0.6

    acme.path                   # acme
    acme.pnl.revenue.path       # acme.pnl.revenue

The factory pattern still works:

::

    quick = mo.Model('quick',
        revenue=mo.Variable(1_000_000),
        cogs=mo.Variable(0.6),
    )
"""

from __future__ import annotations

from typing import Any, Dict, Optional

from .humanize import humanize_identifier
from .multi_variable import MultiVariable


class Model(MultiVariable):
    """Top-level model container — always rooted at its name.

    Inherits everything from :class:`MultiVariable` (component storage,
    adoption, ``with`` block context, ``to_excel``, HTML repr) and adds
    a single semantic difference: the constructor takes a required
    ``name`` positional, and that name crystallizes the model's path
    at construction time. No further explicit naming is needed for
    any descendant.
    """

    def __init__(
        self,
        name: str,
        display_name: Optional[str] = None,
        excel_props: Optional[Dict[str, Any]] = None,
        **components: Any,
    ):
        if not isinstance(name, str) or not name:
            raise TypeError(
                "Model(name) requires a non-empty string. "
                f"Got {type(name).__name__}: {name!r}."
            )
        super().__init__(
            display_name=display_name or humanize_identifier(name),
            excel_props=excel_props,
            **components,
        )
        # Crystallize the root path. Adoption then derives every
        # descendant's path from this rooted prefix.
        self._python_name = name
