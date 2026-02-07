"""
Excel file reader using pandas.
Extracts data from specified cells according to the configuration.
"""

from pathlib import Path

from . import excel_utils


class ExcelReader:
    """Reads data from Excel files based on configuration mapping."""

    def __init__(self, excel_path, config, previous_excel_path=None):
        """
        Initialize the Excel reader.

        Args:
            excel_path: Path to the current month Excel file
            config: Configuration dictionary with field mappings
            previous_excel_path: Path to the previous month Excel file
        """
        self.excel_path = Path(excel_path)
        self.previous_excel_path = (
            Path(previous_excel_path) if previous_excel_path else None
        )
        self.config = config
        self.table_fields = config.get("table_fields", {})
        self.standard_excel_fields = config.get("standard_excel_fields", {})
        self.retail_excel_fields = config.get("retail_excel_fields", {})
        self.supplier_excel_fields = config.get("supplier_excel_fields", {})

    def extract_data(self):
        """
        Extract all configured fields and tables from the Excel file.

        Returns:
            Dictionary mapping field names to their values (includes tables)
        """
        data = {}

        # Extract standard excel fields
        print("\nExtracting standard fields...")
        data.update(self._extract_field_group(self.standard_excel_fields, "standard"))

        # Extract retail excel fields
        print("\nExtracting retail fields...")
        data.update(self._extract_field_group(self.retail_excel_fields, "retail"))

        # Extract supplier excel fields
        print("\nExtracting supplier fields...")
        data.update(self._extract_field_group(self.supplier_excel_fields, "supplier"))

        # Extract table data (simplified)
        print(
            f"\nExtracting table data... (found {len(self.table_fields)} tables configured)"
        )
        for table_name, table_config in self.table_fields.items():
            try:
                sheet = table_config["sheet"]
                column_name = table_config["columns"][0][
                    "col"
                ]  # e.g., "Support Category"
                field_name = table_config["columns"][0]["name"]  # e.g., "re_cat"
                count_column = table_config["count_column"]  # e.g., "re_cat_count"
                limit = table_config.get("limit", None)  # Optional limit for top N
                uncategorized_label = table_config.get(
                    "uncategorized_label", None
                )  # e.g., "Uncategorized"

                print(f"  Reading table '{table_name}' from column '{column_name}'")

                # Get unique value counts
                value_counts = excel_utils.count_unique_values(
                    self.excel_path,
                    sheet,
                    column_name,
                    uncategorized_label=uncategorized_label,
                )

                # Sort by count descending (highest first)
                value_counts.sort(key=lambda x: x["count"], reverse=True)

                # If uncategorized_label is set, move it to the bottom
                if uncategorized_label:
                    uncategorized_items = [
                        item
                        for item in value_counts
                        if item["value"] == uncategorized_label
                    ]
                    categorized_items = [
                        item
                        for item in value_counts
                        if item["value"] != uncategorized_label
                    ]
                    value_counts = categorized_items + uncategorized_items

                # Apply limit if specified (after moving uncategorized to bottom)
                if limit:
                    value_counts = value_counts[:limit]
                    print(f"  → Limiting to top {limit} results")

                # Transform to use configured field names
                table_data = []
                for item in value_counts:
                    row = {field_name: item["value"], count_column: item["count"]}
                    table_data.append(row)

                data[table_name] = table_data
                print(f"  → Found {len(table_data)} unique categories")

            except Exception as e:
                raise ValueError(f"Error reading table '{table_name}': {e}") from e

        # Debug: Show all extracted field names
        print(
            f"\nExtracted fields: {', '.join([k for k in data if not isinstance(data[k], list)])}"
        )

        return data

    def _extract_field_group(self, field_group, group_name):
        """
        Extract fields from a field group (standard/retail/supplier).

        Args:
            field_group: Dictionary containing the field group configuration
            group_name: Name of the group (for error messages)

        Returns:
            Dictionary of extracted field values
        """
        if not field_group:
            return {}

        data = {}

        # Get the default sheet for this group
        default_sheet = field_group.get("sheet")

        # Process each field in the group
        for field_name, field_config in field_group.items():
            # Skip the 'sheet' key itself
            if field_name == "sheet":
                continue

            if not isinstance(field_config, dict):
                continue

            try:
                # Get the sheet name (use field-specific or fall back to group default)
                sheet = field_config.get("sheet", default_sheet)
                cell = field_config.get("cell")

                # Determine which Excel file to use
                # Fields with 'prev' in the name use previous month file, others use current month
                is_prev_field = "_prev_" in field_name
                excel_path = (
                    self.previous_excel_path
                    if is_prev_field and self.previous_excel_path
                    else self.excel_path
                )

                if is_prev_field and not self.previous_excel_path:
                    print(
                        f"  Warning: '{field_name}' is a previous month field but no previous Excel file provided"
                    )
                    data[field_name] = None
                    continue

                # Check if this is a row count field (ends with _req or _prev_req)
                if field_name.endswith("_req"):
                    # Count rows in the sheet
                    value = excel_utils.count_rows(excel_path, sheet)
                    data[field_name] = value
                    continue

                # Check if this is a satisfaction field
                if field_name.endswith("_sat_c"):
                    # Count "good with comment" values
                    if cell is None or cell == "":
                        print(f"  Skipping '{field_name}': no cell configuration")
                        data[field_name] = None
                        continue
                    print(
                        f"  Reading '{field_name}' from {'previous' if is_prev_field else 'current'} month file, sheet '{sheet}', column '{cell}'"
                    )
                    value = excel_utils.count_column_value(
                        excel_path, sheet, cell, "good with comment"
                    )
                    print(f"  → Found {value} occurrences")
                    data[field_name] = value
                    continue
                elif field_name.endswith("_sat"):
                    # Count "good" values (but not "good with comment")
                    if cell is None or cell == "":
                        print(f"  Skipping '{field_name}': no cell configuration")
                        data[field_name] = None
                        continue
                    print(
                        f"  Reading '{field_name}' from {'previous' if is_prev_field else 'current'} month file, sheet '{sheet}', column '{cell}'"
                    )
                    value = excel_utils.count_column_value(
                        excel_path, sheet, cell, "good"
                    )
                    print(f"  → Found {value} occurrences")
                    data[field_name] = value
                    continue

                if field_name.endswith("_reso"):
                    print(
                        f"  Reading '{field_name}' from {'previous' if is_prev_field else 'current'} month file, sheet '{sheet}'"
                    )
                    value = excel_utils.calculate_average_resolution_time(
                        excel_path, sheet
                    )
                    print(f"  → Found {value} days average resolution time")
                    data[field_name] = value
                    continue

                # Skip fields with no cell configuration yet
                if cell is None or cell == "":
                    print(f"  Skipping '{field_name}': no cell configuration")
                    data[field_name] = None
                    continue

                # Read the value from the column
                value = excel_utils.read_column_value(excel_path, sheet, cell)

                # Handle special field types
                if field_name == "month":
                    # Parse date and format as "March"
                    value = excel_utils.parse_month_from_date(value)
                elif field_name == "prev_month":
                    # Parse date, subtract one month, and format as "February"
                    value = excel_utils.parse_previous_month_from_date(value)

                data[field_name] = value

            except Exception as e:
                print(f"  Warning: Error reading '{field_name}' from {group_name}: {e}")
                data[field_name] = None

        return data
