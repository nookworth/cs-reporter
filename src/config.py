"""
Configuration loader for the CS Reporter.
Loads and validates YAML configuration files.
"""

from pathlib import Path

import yaml


# ==============================================================================
# OPERATIONS REGISTRY
# ==============================================================================
# This registry defines all valid operations that can be used in field configs.
# Each operation specifies its required and optional parameters.
# This enables validation and provides a reference for config authors.
# ==============================================================================

OPERATIONS = {
    "count_rows": {
        "required": ["sheet"],
        "optional": ["filters"],
    },
    
    "count_value": {
        "required": ["sheet", "column", "value"],
        "optional": ["filters"],
    },
    
    "sum": {
        "required": ["sheet", "column"],
        "optional": ["filters"],
    },
    
    "avg_date_diff": {
        "required": ["sheet", "start_column", "end_column"],
        "optional": ["filters", "unit"],
    },
    
    "parse_month": {
        "required": ["sheet", "column"],
        "optional": [],
    },
    
    "parse_previous_month": {
        "required": ["sheet", "column"],
        "optional": [],
    },
    
    "read_value": {
        "required": ["sheet", "column"],
        "optional": [],
    },
}


# ==============================================================================
# VALIDATION FUNCTION
# ==============================================================================

def validate_fields(config):

    fields = config.get("fields", {})
    
    if not fields:
        return
    
    for field_name, field_config in fields.items():
        if not isinstance(field_config, dict):
            continue
        
        # VALIDATION 1: Check that 'operation' key exists
        if "operation" not in field_config:
            raise ValueError(
                f"Field '{field_name}' is missing required key 'operation'\n"
                f"Every field must specify an operation. Available operations:\n"
                f"{', '.join(sorted(OPERATIONS.keys()))}"
            )
        
        operation = field_config["operation"]
        
        # VALIDATION 2: Check that operation is valid
        if operation not in OPERATIONS:
            raise ValueError(
                f"Field '{field_name}' has unknown operation '{operation}'\n"
                f"Valid operations are: {', '.join(sorted(OPERATIONS.keys()))}"
            )
        
        # VALIDATION 3: Check that all required parameters are present
        required_params = OPERATIONS[operation]["required"]
        for param in required_params:
            if param not in field_config:
                raise ValueError(
                    f"Field '{field_name}' with operation '{operation}' is missing required parameter '{param}'\n"
                    f"Required parameters for '{operation}': {', '.join(required_params)}"
                )
    
    # Validation passed!
    return


def load_config(config_path=None):
    """
    Load the configuration from a YAML file.

    Args:
        config_path: Optional path to config file.
                    Defaults to 'config/mapping.yaml'

    Returns:
        Dictionary containing the configuration

    Raises:
        FileNotFoundError: If config file doesn't exist
        ValueError: If config is invalid
    """
    if config_path is None:
        project_root = Path(__file__).parent.parent
        config_path = project_root / "config" / "mapping.yaml"
    else:
        config_path = Path(config_path)

    if not config_path.exists():
        raise FileNotFoundError(
            f"Configuration file not found: {config_path}\n"
            f"Please create a mapping.yaml file in the config directory."
        )

    with open(config_path, encoding="utf-8") as f:
        config = yaml.safe_load(f)

    if "template_path" not in config:
        raise ValueError("Config must contain 'template_path'")

    project_root = Path(__file__).parent.parent
    template_path = Path(config["template_path"])

    if not template_path.is_absolute():
        template_path = project_root / template_path

    if not template_path.exists():
        raise FileNotFoundError(
            f"Template file not found: {template_path}\n"
            f"Please ensure the PowerPoint template exists at this location."
        )

    config["template_path"] = str(template_path)
    config["config_path"] = str(config_path)

    if "output_dir" not in config:
        config["output_dir"] = str(project_root / "output")

    validate_fields(config)

    return config
