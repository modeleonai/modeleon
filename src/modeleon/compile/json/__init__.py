# SPDX-License-Identifier: Apache-2.0
"""JSON renderer subpackage.

Renders ``Expr`` trees to renderer-neutral JSON-serializable dicts —
useful for diff views, inspection tools, LLM formula analysis, or any
consumer that needs the formula structure without target-specific
formatting. See :mod:`~modeleon.compile.json.renderer`.
"""

from .renderer import JsonRenderer, to_json

__all__ = ["JsonRenderer", "to_json"]
