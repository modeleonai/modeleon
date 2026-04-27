# SPDX-License-Identifier: Apache-2.0
from modeleon.plugins import get_registry, load_plugins, reset_registry


def test_registry_has_expected_categories():
    registry = get_registry()
    assert "variable_kwargs" in registry
    assert "compiler_passes" in registry
    assert "node_validators" in registry
    assert "_loaded" in registry


def test_load_plugins_is_idempotent():
    reset_registry()
    load_plugins()
    loaded_first = list(get_registry()["_loaded"])
    load_plugins()
    loaded_second = list(get_registry()["_loaded"])
    assert loaded_first == loaded_second


def test_mock_plugin_registers():
    reset_registry()
    registry = get_registry()

    def mock_register(reg: dict) -> None:
        reg["variable_kwargs"]["test_kwarg"] = "test_handler"

    mock_register(registry)
    assert registry["variable_kwargs"]["test_kwarg"] == "test_handler"

    # Clean up
    reset_registry()
