# SPDX-License-Identifier: Apache-2.0
"""Notebook smoke tests.

Runs every `.ipynb` under ``packages/engine/notebooks/`` via nbclient and
asserts no cell raises. Skipped when nbclient isn't installed (nbclient
is not a core test dep).

Part of the CP0 safety net for ADR-008.
"""

from __future__ import annotations

from pathlib import Path

import pytest

NOTEBOOKS_DIR = Path(__file__).resolve().parent.parent / "notebooks"
NOTEBOOKS = sorted(NOTEBOOKS_DIR.glob("*.ipynb"))


@pytest.mark.parametrize("nb_path", NOTEBOOKS, ids=[p.name for p in NOTEBOOKS])
def test_notebook_runs_to_completion(nb_path: Path):
    nbformat = pytest.importorskip("nbformat")
    nbclient = pytest.importorskip("nbclient")

    nb = nbformat.read(nb_path, as_version=4)
    client = nbclient.NotebookClient(
        nb,
        timeout=120,
        resources={"metadata": {"path": str(nb_path.parent)}},
    )
    client.execute()
