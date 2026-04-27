# SPDX-License-Identifier: Apache-2.0
"""Identifier → display-label conversion.

Shared helper that turns a Python identifier (``revenue_cogs``, ``XIRR``,
``Smth_smth2``) into a human-readable spreadsheet label. Used as the
fallback in ``Variable.display_name`` and ``MultiVariableBase.display_name``
so a Python binding name produces a sensible label without the user
having to pass ``display_name=`` for every Variable/MV.
"""

from __future__ import annotations


def identifier_from_label(label: str) -> str:
    """Inverse of ``humanize_identifier``: ``"Revenue NPV"`` → ``"revenue_npv"``.

    Used as a formula-rendering fallback when a Variable has a
    ``display_name`` but no ``python_name`` (e.g. a Variable created at the
    top level with ``mo.Variable(value, display_name="Revenue")`` — no
    attribute assignment to a parent). Non-alphanumeric characters become
    underscores; runs of underscores collapse.
    """
    out_chars: list[str] = []
    for ch in label:
        if ch.isalnum():
            out_chars.append(ch.lower())
        elif out_chars and out_chars[-1] != '_':
            out_chars.append('_')
    return ''.join(out_chars).strip('_') or label


def humanize_identifier(identifier: str) -> str:
    """Split on ``_``, title-case each token, preserve ALL-CAPS tokens.

    Examples::

        revenue            → "Revenue"
        total_revenue      → "Total Revenue"
        XIRR               → "XIRR"
        revenue_NPV        → "Revenue NPV"
        Smth_smth2         → "Smth Smth2"
        _var_5             → "Var 5"   (leading ``_`` dropped)
    """
    tokens = [t for t in identifier.split('_') if t]
    return ' '.join(t if t.isupper() else t.title() for t in tokens)
