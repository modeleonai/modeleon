# SPDX-License-Identifier: Apache-2.0
"""Core DSL primitives.

Layer cake (one coherent job per module):

    base / component   — identity (QPath, name, display label) + tree typing
    qpath              — qualified-path identity
    humanize           — identifier → display-label conversion

    Variable           — the atomic cell primitive (value or formula)
    ├── variable_init  — construction dispatch (value/formula/pyformula → state)
    └── variable_ops   — arithmetic + comparison operators + broadcasting

    MultiVariableBase  — the tree primitive
    ├── mv_context     — with-block lifecycle + component adoption
    ├── mv_inspect     — tree navigation, cycle detection, metadata export
    ├── MultiVariable / MultiVariableClass — subclasses (multi_variable)
    └── Model          — the top-level container (model)

    shape / time       — axes (``indexed_by``) and the period axis + grains
    projection         — grain projection of a Variable onto another window
    regrain            — ``mo.up`` / ``mo.frozen`` / ``mo.ratio`` re-grain specs
    extend             — out-of-extent behavior (``extend=``)
    tracks / tracks_decl / blend — the tracks axis and its declaration
    unit               — measurement units with algebra
    expr               — AST node types (BinOp, VarRef, Literal, SelfRef, …)
    errors             — CircularDependencyError, CrossScopeReferenceWarning
"""
