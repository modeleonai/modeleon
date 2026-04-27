# SPDX-License-Identifier: Apache-2.0
"""Cross-platform path handling for ``.to_excel(...)``.

The writer accepts ``str | Path`` and wraps with :class:`pathlib.Path`,
but "works on Linux" doesn't prove "works on Windows." These cases
exercise the path surface that historically breaks on Windows — spaces
in directory names, unicode filenames, mixed separators, nested
output directories that must already exist or be created.

Runs on every platform in CI; the Windows job is where it earns its
keep.
"""

from __future__ import annotations

import modeleon as mo


def _tiny_model() -> "mo.MultiVariable":
    """One tab with one Variable — minimal input for a workbook write."""
    model = mo.MultiVariable("M")
    model.s = mo.MultiVariable("S", excel_props={'tab': True})
    model.s.x = mo.Variable(1.0)
    return model


class TestPathTypes:
    """Accepts both ``str`` and :class:`pathlib.Path`."""

    def test_string_path(self, tmp_path):
        _tiny_model().to_excel(str(tmp_path / "out.xlsx"))
        assert (tmp_path / "out.xlsx").exists()

    def test_path_object(self, tmp_path):
        _tiny_model().to_excel(tmp_path / "out.xlsx")
        out = tmp_path / "out.xlsx"
        assert out.exists()


class TestTrickyFilenames:
    """Spaces, unicode, and dots in the filename — all valid on every OS."""

    def test_spaces_in_filename(self, tmp_path):
        _tiny_model().to_excel(tmp_path / "my forecast.xlsx")
        out = tmp_path / "my forecast.xlsx"
        assert out.exists()

    def test_unicode_filename(self, tmp_path):
        # Finance teams globally — Cyrillic, Chinese, emoji surrogates
        # all get hit in the wild. Bare minimum: non-ASCII basename.
        _tiny_model().to_excel(tmp_path / "модель.xlsx")
        out = tmp_path / "модель.xlsx"
        assert out.exists()

    def test_multiple_dots(self, tmp_path):
        _tiny_model().to_excel(tmp_path / "forecast.v2.final.xlsx")
        out = tmp_path / "forecast.v2.final.xlsx"
        assert out.exists()


class TestNestedDirectory:
    """Target directory must exist — we don't auto-create (surprising)."""

    def test_existing_subdir(self, tmp_path):
        subdir = tmp_path / "outputs"
        subdir.mkdir()
        _tiny_model().to_excel(subdir / "m.xlsx")
        out = subdir / "m.xlsx"
        assert out.exists()

    def test_spaces_in_subdir(self, tmp_path):
        subdir = tmp_path / "my outputs"
        subdir.mkdir()
        _tiny_model().to_excel(subdir / "m.xlsx")
        out = subdir / "m.xlsx"
        assert out.exists()
