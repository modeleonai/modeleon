"""Shared test configuration.

``slow`` marks the tests that open a written workbook in LibreOffice and
read back what it recalculated — a few seconds each, most of the suite's
time. A quick run leaves them out; a full run keeps them:

    pytest -m "not slow"      # seconds
    pytest                    # everything
"""

from __future__ import annotations

import pytest

# Imported in the controlling process before any worker runs: under
# pytest-xdist a worker's warning (``CrossScopeReferenceWarning``) is
# rebuilt here by ``importlib.import_module`` in a receiver thread, and
# two receiver threads importing ``modeleon`` at once deadlocked on its
# import locks (``_DeadlockError``, then ``KeyError: <WorkerController>``
# in the scheduler). A module already in ``sys.modules`` takes no lock.
import modeleon  # noqa: E402,F401


def pytest_configure(config: pytest.Config) -> None:
    config.addinivalue_line(
        "markers",
        "slow: recalculates a written workbook in LibreOffice (seconds each); "
        "deselect with -m 'not slow' for a quick run",
    )
