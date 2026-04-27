# SPDX-License-Identifier: Apache-2.0
"""Excel cell addresses for a Variable after layout.

The :class:`VariableAddresses` dataclass is the contract between the
layout pass (:class:`~modeleon.compile.excel.layout.LayoutEngine`,
which produces instances) and its consumers
(:class:`~modeleon.compile.excel.writer`,
:class:`~modeleon.compile.excel.renderer.ExcelRenderer`,
:class:`~modeleon.compile.excel.translator.ExcelTranslator`).

Lives in ``compile/excel/`` because it's Excel-specific — other
backends (JSON, Pandas) have completely different addressing or none
at all. Variables themselves never carry a reference to this type;
the addresses dict is keyed by ``var.id`` and consumed at emission
time.
"""

from dataclasses import dataclass
from typing import List


@dataclass
class VariableAddresses:
    """Excel addresses for a Variable.

    Assigned by :class:`~modeleon.compile.excel.layout.LayoutEngine`
    after all Variables have been created and before the writer runs.

    Fields:
        name:    Label-cell address (``"A4"``) — where the Variable's
                 display name is written in column A.
        formula: Formula-cell address (``"B4"``) — the cell used for
                 single-column rendering of a formula (unused for list
                 Variables; retained as a separate field for consumers
                 that distinguish the formula cell from the value run).
        values:  Per-period value cell addresses
                 (``["C4", "D4", "E4", ...]``). For a scalar Variable
                 this is a single-element list; for list Variables, one
                 entry per period.
    """

    name: str
    formula: str
    values: List[str]
