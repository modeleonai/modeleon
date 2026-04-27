# SPDX-License-Identifier: Apache-2.0
"""Core DSL primitives.

Layer cake (each file ~250-620 lines, one coherent job):

    Variable           — the atomic cell primitive (value or formula)
    ├── variable_init  — construction dispatch (value/formula/pyformula → state)
    ├── variable_ops   — arithmetic + comparison operators + broadcasting

    MultiVariableBase  — the tree primitive
    ├── mv_context     — with-block lifecycle + component adoption
    └── MultiVariable / Sheet / Model — subclass cascade (one file)

    expr               — AST node types (BinOp, VarRef, Literal, SelfRef, …)
    ids                — module-level id counters (_var_N, _mv_N, creation_seq)
    addresses          — VariableAddresses dataclass (layout output)
    errors             — CircularDependencyError
"""
