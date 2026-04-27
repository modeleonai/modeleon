# SPDX-License-Identifier: Apache-2.0
"""Presentation-layer code for the engine.

HTML rendering for Jupyter `_repr_html_`, and anything else that turns
model state into a human-facing view. Kept separate from `core/` so
the model primitives stay free of presentation concerns.
"""

from .html import model_html, multi_variable_html, variable_html

__all__ = ["model_html", "multi_variable_html", "variable_html"]
