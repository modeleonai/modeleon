# SPDX-License-Identifier: Apache-2.0
"""Tree navigation, cycle detection, and metadata export for MVs.

Pulled out of ``multi_variable.py`` so the core primitive stays focused
on component storage + identity + attribute protocol. Everything in
here is stateless over ``MultiVariableBase`` — it just reads state.

The mixin :class:`_MVInspect` is inherited by
:class:`MultiVariableBase`. Methods read ``self._components``,
``self._component_order``, ``self._parent``, ``self._is_sheet``,
``self.id``, ``self.display_name``, etc. — the public surface of the
MV primitive.

Lazy imports for :class:`Variable` and :class:`MultiVariableBase`
inside methods avoid the circular import that would arise if we
resolved them at module scope.
"""

from __future__ import annotations

from typing import Any, Dict, List, Optional, TYPE_CHECKING

if TYPE_CHECKING:
    pass


class _MVInspect:
    """Tree navigation + cycle detection + serialization mixin."""

    # ─── Component introspection ────────────────────────────────

    @property
    def component_names(self) -> List[str]:
        """List of all component names in creation order."""
        return self._component_order.copy()

    @property
    def sheet_name(self) -> Optional[str]:
        """Resolved sheet (Excel tab) name for sheet-role MVs.

        Returns ``display_name`` (which itself falls back to the
        humanized python_name). Returns ``None`` for non-sheet MVs.
        """
        if self._is_sheet:
            return self.display_name
        return None

    def walk(self):
        """Yield all components in depth-first order as ``(name, component)``
        tuples. Components can be ``Variable`` or ``MultiVariableBase``
        instances."""
        from .multi_variable import MultiVariableBase
        for name in self._component_order:
            component = self._components[name]
            yield name, component
            if isinstance(component, MultiVariableBase):
                yield from component.walk()

    # ─── Cycle detection ────────────────────────────────────────

    def _collect_variables(self) -> Dict[str, Any]:
        """Return ``{var_id: Variable}`` for every Variable reachable from
        this MV's subtree. Used by cycle detection and emission.

        Keys are qualified-path ids (``.id``) — matches what
        :attr:`Variable.dependencies` emits, so cycle detection joins
        on the same keyspace.
        """
        from .variable import Variable
        seen: Dict[str, Any] = {}
        for _, comp in self.walk():
            if isinstance(comp, Variable):
                seen[comp.id] = comp
        return seen

    def detect_cycles(self) -> Optional[List[str]]:
        """Return the first dependency cycle in this MV's subtree, or None.

        Walks the subtree, collects Variables, follows ``_dependency_refs``
        to build a DAG, runs iterative DFS. Returned list closes the
        cycle (``cycle[0] == cycle[-1]``).
        """
        variables = self._collect_variables()
        graph: Dict[str, List[str]] = {
            vid: [ref.id for ref in var._dependency_refs]
            for vid, var in variables.items()
        }

        visited: set = set()
        in_stack: set = set()
        stack_path: List[str] = []

        def dfs(start: str) -> Optional[List[str]]:
            work = [(start, iter(graph.get(start, [])))]
            in_stack.add(start)
            stack_path.append(start)
            while work:
                node, deps = work[-1]
                found_next = False
                for dep in deps:
                    if dep in in_stack:
                        idx = stack_path.index(dep)
                        return stack_path[idx:] + [dep]
                    if dep not in visited and dep in graph:
                        in_stack.add(dep)
                        stack_path.append(dep)
                        work.append((dep, iter(graph.get(dep, []))))
                        found_next = True
                        break
                if not found_next:
                    work.pop()
                    in_stack.discard(node)
                    if stack_path and stack_path[-1] == node:
                        stack_path.pop()
                    visited.add(node)
            return None

        for node in list(graph.keys()):
            if node not in visited:
                cycle = dfs(node)
                if cycle:
                    return cycle
        return None

    def assert_acyclic(self) -> None:
        """Raise :class:`CircularDependencyError` if this MV's subtree
        has a dependency cycle among its Variables."""
        from .errors import CircularDependencyError
        cycle = self.detect_cycles()
        if cycle:
            raise CircularDependencyError(cycle)

    # ─── Serialization ──────────────────────────────────────────

    def to_dict(self) -> Dict[str, Any]:
        """Export MultiVariable metadata and component data as a
        JSON-serializable dict.

        Every ``id`` / ``parent`` field is the qualified-path dotted
        string — the single source of identity. Unrooted nodes carry a
        ``__floating__.mN``-shaped id.
        """
        from .multi_variable import MultiVariableBase
        from .variable import Variable

        components_data: Dict[str, Any] = {}
        for name, comp in self._components.items():
            if isinstance(comp, MultiVariableBase):
                components_data[name] = comp.to_dict()
            elif isinstance(comp, Variable):
                components_data[name] = {
                    'id': comp.id,
                    'values': comp._value,
                    'formula': comp.formula,
                }

        return {
            'id': self.id,
            'type': self.__class__.__name__,
            'python_name': self.python_name,
            'name': self.display_name,
            'components': self._component_order,
            'params': {
                k: repr(v) if not isinstance(v, (str, int, float, bool, list, dict, type(None))) else v
                for k, v in self._creation_params.items()
            },
            'sheet': self.sheet_name,
            'is_sheet': self._is_sheet,
            'parent': self._parent.id if self._parent else None,
            'components_data': components_data,
        }
