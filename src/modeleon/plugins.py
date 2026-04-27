# SPDX-License-Identifier: Apache-2.0
"""Plugin extension registry for modeleon.

Extensions register capabilities here via Python entry points. Users
always write ``import modeleon as mo`` — extension features appear
automatically when a package exposing the ``modeleon.plugins`` entry
point is installed.
"""

from __future__ import annotations

import importlib.metadata
import logging
from typing import Any

logger = logging.getLogger(__name__)

_registry: dict[str, Any] = {
    # Extra kwargs accepted by Variable (e.g., control=, access=, style=)
    "variable_kwargs": {},
    # Additional compiler passes (run after core compilation)
    "compiler_passes": [],
    # Additional node validators (run on graph validation)
    "node_validators": [],
    # Track which plugins have loaded
    "_loaded": [],
}

_plugins_loaded = False


def get_registry() -> dict[str, Any]:
    """Return the plugin registry. Read-only access for inspection and tests."""
    return _registry


def load_plugins() -> None:
    """Discover and load all installed plugins from the 'modeleon.plugins' entry point group.

    Called once on engine import. Safe to call multiple times (idempotent).
    """
    global _plugins_loaded
    if _plugins_loaded:
        return
    _plugins_loaded = True

    eps = importlib.metadata.entry_points(group="modeleon.plugins")
    for ep in eps:
        try:
            plugin_module = ep.load()
            plugin_module.register(_registry)
            _registry["_loaded"].append(ep.name)
            logger.debug("Loaded plugin: %s", ep.name)
        except Exception:
            logger.warning("Failed to load plugin: %s", ep.name, exc_info=True)


def reset_registry() -> None:
    """Reset the registry to its initial state. For testing only."""
    global _plugins_loaded
    _registry["variable_kwargs"].clear()
    _registry["compiler_passes"].clear()
    _registry["node_validators"].clear()
    _registry["_loaded"].clear()
    _plugins_loaded = False
