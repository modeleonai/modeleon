"""Shared test configuration.

``slow`` marks the tests that open a written workbook in LibreOffice and
read back what it recalculated — a few seconds each, most of the suite's
time. A quick run leaves them out; a full run keeps them:

    pytest -m "not slow"      # seconds
    pytest                    # everything
"""

from __future__ import annotations

import subprocess
from pathlib import Path
from typing import Any
from urllib.parse import unquote, urlparse

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


# LibreOffice keeps an .xlsx formula's cached value unless its profile says
# to recalculate on load: its default for Excel 2007+ files is "never", and
# it ignores fullCalcOnLoad. The writer caches the engine's own values, so a
# default profile would hand the engine's numbers back, and every parity
# check against LibreOffice would compare the engine with itself. Each test
# launches LibreOffice with its own profile (-env:UserInstallation=...);
# this seeds that profile to recalculate everything on load.
_RECALC_ALWAYS_XCU = (
    '<?xml version="1.0" encoding="UTF-8"?>\n'
    '<oor:items xmlns:oor="http://openoffice.org/2001/registry" '
    'xmlns:xs="http://www.w3.org/2001/XMLSchema" '
    'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n'
    '<item oor:path="/org.openoffice.Office.Calc/Formula/Load">'
    '<prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>\n'
    '</oor:items>\n'
)


@pytest.fixture(autouse=True)
def _libreoffice_recalculates_on_load(monkeypatch: pytest.MonkeyPatch) -> None:
    real_run = subprocess.run

    def run(args: Any, *rest: Any, **kwargs: Any) -> Any:
        for arg in args if isinstance(args, (list, tuple)) else ():
            if isinstance(arg, str) and arg.startswith("-env:UserInstallation="):
                user = Path(unquote(urlparse(arg.partition("=")[2]).path)) / "user"
                user.mkdir(parents=True, exist_ok=True)
                xcu = user / "registrymodifications.xcu"
                if not xcu.exists():
                    xcu.write_text(_RECALC_ALWAYS_XCU)
        return real_run(args, *rest, **kwargs)

    monkeypatch.setattr(subprocess, "run", run)
