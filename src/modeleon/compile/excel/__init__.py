# SPDX-License-Identifier: Apache-2.0
"""Excel renderer subpackage.

Every Excel-specific compile-stage concern lives here:

- :mod:`~modeleon.compile.excel.renderer`     — ``ExcelRenderer``: per-node renderer.
- :mod:`~modeleon.compile.excel.layout`      — ``LayoutEngine``: cell-address assignment.
- :mod:`~modeleon.compile.excel.translator`  — ``ExcelTranslator``: facade over the renderer.
- :mod:`~modeleon.compile.excel.writer`      — ``to_excel``: drives layout + translator + openpyxl.

Generic infrastructure (``Renderer`` Protocol, ``Walker``) lives one
level up at :mod:`modeleon.compile`. Adding a parallel
``compile/google_sheets/`` or ``compile/pandas/`` subpackage is the
natural extension point.
"""

from .renderer import ExcelRenderer
from .layout import LayoutEngine
from .translator import ExcelTranslator
from .writer import to_excel

__all__ = ["ExcelRenderer", "ExcelTranslator", "LayoutEngine", "to_excel"]
