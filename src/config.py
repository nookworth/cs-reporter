"""
Configuration loader for the CS Reporter.
Loads and validates YAML configuration files.
"""

from pathlib import Path

import yaml


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
        # Default to config/mapping.yaml in the project root
        project_root = Path(__file__).parent.parent
        config_path = project_root / "config" / "mapping.yaml"
    else:
        config_path = Path(config_path)

    if not config_path.exists():
        raise FileNotFoundError(
            f"Configuration file not found: {config_path}\n"
            f"Please create a mapping.yaml file in the config directory."
        )

    # Load YAML (with UTF-8 encoding for Windows compatibility)
    with open(config_path, encoding="utf-8") as f:
        config = yaml.safe_load(f)

    if "template_path" not in config:
        raise ValueError("Config must contain 'template_path'")

    # Resolve template path relative to project root
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

    # Set default output directory if not specified
    if "output_dir" not in config:
        config["output_dir"] = str(project_root / "output")

    return config
