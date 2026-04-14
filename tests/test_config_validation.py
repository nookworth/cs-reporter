"""Unit tests for validate_fields() in src/config_v2.py."""

import pytest

from src.config_v2 import validate_fields


class TestValidateFields:
    def test_missing_operation_key_raises_value_error(self):
        """Missing 'operation' key should raise ValueError."""
        config = {
            "fields": {
                "field1": {
                    "sheet": "Sheet1",
                    "column": "A",
                }
            }
        }
        with pytest.raises(ValueError) as exc_info:
            validate_fields(config)
        assert "missing required key 'operation'" in str(exc_info.value)

    def test_unknown_operation_raises_value_error_with_valid_ops(self):
        """Unknown operation should raise ValueError with valid operation names."""
        config = {
            "fields": {
                "field1": {
                    "operation": "invalid_operation",
                    "sheet": "Sheet1",
                }
            }
        }
        with pytest.raises(ValueError) as exc_info:
            validate_fields(config)
        assert "unknown operation 'invalid_operation'" in str(exc_info.value)
        assert "Valid operations are:" in str(exc_info.value)
        assert "count_rows" in str(exc_info.value)
        assert "count_value" in str(exc_info.value)
        assert "sum" in str(exc_info.value)

    def test_missing_required_parameter_raises_value_error(self):
        """Missing required parameter should raise ValueError."""
        config = {
            "fields": {
                "field1": {
                    "operation": "count_value",
                    "sheet": "Sheet1",
                }
            }
        }
        with pytest.raises(ValueError) as exc_info:
            validate_fields(config)
        assert "missing required parameter" in str(exc_info.value)
        assert "column" in str(exc_info.value)
        assert "value" in str(exc_info.value)

    def test_valid_config_passes_without_raising(self):
        """Valid config should pass without raising."""
        config = {
            "fields": {
                "field1": {
                    "operation": "count_rows",
                    "sheet": "Sheet1",
                }
            }
        }
        validate_fields(config)

    def test_valid_config_with_optional_params_passes(self):
        """Valid config with optional params should pass."""
        config = {
            "fields": {
                "field1": {
                    "operation": "sum",
                    "sheet": "Sheet1",
                    "column": "A",
                    "filters": [],
                }
            }
        }
        validate_fields(config)

    def test_empty_fields_passes(self):
        """Empty fields dict should pass."""
        config = {"fields": {}}
        validate_fields(config)

    def test_no_fields_key_passes(self):
        """Config without fields key should pass."""
        config = {"template_path": "test.pptx"}
        validate_fields(config)

    def test_non_dict_field_config_skipped(self):
        """Non-dict field config should be skipped (not raise)."""
        config = {
            "fields": {
                "field1": "not a dict",
            }
        }
        validate_fields(config)
