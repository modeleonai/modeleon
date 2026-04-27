# SPDX-License-Identifier: Apache-2.0
"""Model → output pipeline.

Layered so the AST stays renderer-neutral and new renderers slot in
alongside the Excel one without touching core:

- :mod:`.renderer`      — ``Renderer`` Protocol + ``RenderCtx``. The
                          contract every renderer satisfies.
- :mod:`.walker`        — ``Walker``: generic AST dispatcher. Takes a
                          renderer, routes each ``Expr`` node type to
                          ``renderer.render_*``.
- :mod:`.excel`         — Excel renderer subpackage:
                          ``renderer.py`` / ``translator.py`` /
                          ``layout.py`` / ``writer.py``.
- :mod:`.json`          — JSON renderer subpackage (proves the seam).

Renderers produce output strings / dicts from the AST — they do not
compute values. ``Variable._value`` is computed eagerly in Python at
operator time; renderers consume the formula structure to emit Excel
formulas or JSON dicts.

The :class:`~modeleon.compile.excel.addresses.VariableAddresses`
dataclass — the contract between layout and the Excel renderer — lives
in :mod:`modeleon.compile.excel.addresses` now that ``core/`` is
free of Excel-specific types.

Called via ``root.to_excel(path)`` on any MultiVariable. For the JSON
form, use ``mo.to_json(expr)`` or :func:`.json.to_json`.

The names in ``__all__`` are the extension surface that downstream
packages may instantiate directly. User code should go through
``root.to_excel(path)`` or ``to_json(var._expr)`` instead of these
low-level entry points.
"""

from .compat import CompatIssue, check_compat
from .renderer import Renderer, RenderCtx
from .excel import ExcelRenderer, ExcelTranslator, LayoutEngine, to_excel
from .json import JsonRenderer, to_json
from .walker import Walker

__all__ = [
    "CompatIssue",
    "Renderer",
    "ExcelRenderer",
    "ExcelTranslator",
    "JsonRenderer",
    "LayoutEngine",
    "RenderCtx",
    "Walker",
    "check_compat",
    "to_excel",
    "to_json",
]
