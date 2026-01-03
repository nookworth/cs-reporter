"""
Excel file reader using pandas.
Extracts data from specified cells according to the configuration.
"""

import pandas as pd
from pathlib import Path
from datetime import datetime

from . import excel_utils


class ExcelReader:
    """Reads data from Excel files based on configuration mapping."""

    def __init__(self, excel_path, config, previous_month_data=None):
        """
        Initialize the Excel reader.

        Args:
            excel_path: Path to the Excel file
            config: Configuration dictionary with field mappings
            previous_month_data: Optional dictionary with previous month's data (currently unused)
        """
        self.excel_path = Path(excel_path)
        self.config = config
        self.table_fields = config.get('table_fields', {})
        self.standard_excel_fields = config.get('standard_excel_fields', {})
        self.retail_excel_fields = config.get('retail_excel_fields', {})
        self.supplier_excel_fields = config.get('supplier_excel_fields', {})

    def extract_data(self):
        """
        Extract all configured fields and tables from the Excel file.

        Returns:
            Dictionary mapping field names to their values (includes tables)
        """
        data = {}

        # Extract standard excel fields
        data.update(self._extract_field_group(self.standard_excel_fields, 'standard'))

        # Extract retail excel fields
        data.update(self._extract_field_group(self.retail_excel_fields, 'retail'))

        # Extract supplier excel fields
        data.update(self._extract_field_group(self.supplier_excel_fields, 'supplier'))

        # Extract table data
        for table_name, table_config in self.table_fields.items():
            try:
                table_data = self._read_table(table_name, table_config)
                data[table_name] = table_data
            except Exception as e:
                raise ValueError(
                    f"Error reading table '{table_name}': {e}"
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
        default_sheet = field_group.get('sheet')

        # Process each field in the group
        for field_name, field_config in field_group.items():
            # Skip the 'sheet' key itself
            if field_name == 'sheet':
                continue

            if not isinstance(field_config, dict):
                continue

            try:
                # Get the sheet name (use field-specific or fall back to group default)
                sheet = field_config.get('sheet', default_sheet)
                cell = field_config.get('cell')

                # Check if this is a row count field (ends with _req or _prev_req)
                if field_name.endswith('_req'):
                    # Count rows in the sheet
                    value = excel_utils.count_rows(self.excel_path, sheet)
                    data[field_name] = value
                    continue

                # Check if this is a satisfaction field
                if field_name.endswith('_sat_c'):
                    # Count "good with comment" values
                    if cell is None or cell == '':
                        print(f"  Skipping '{field_name}': no cell configuration")
                        data[field_name] = None
                        continue
                    value = excel_utils.count_column_value(
                        self.excel_path, sheet, cell, "good with comment"
                    )
                    data[field_name] = value
                    continue
                elif field_name.endswith('_sat'):
                    # Count "good" values (but not "good with comment")
                    if cell is None or cell == '':
                        print(f"  Skipping '{field_name}': no cell configuration")
                        data[field_name] = None
                        continue
                    value = excel_utils.count_column_value(
                        self.excel_path, sheet, cell, "good"
                    )
                    data[field_name] = value
                    continue

                # Skip fields with no cell configuration yet
                if cell is None or cell == '':
                    print(f"  Skipping '{field_name}': no cell configuration")
                    data[field_name] = None
                    continue

                # Read the value from the column
                value = excel_utils.read_column_value(self.excel_path, sheet, cell)

                # Handle special field types
                if field_name == 'month':
                    # Parse date and format as "March"
                    value = excel_utils.parse_month_from_date(value)
                elif field_name == 'prev_month':
                    # Parse date, subtract one month, and format as "February"
                    value = excel_utils.parse_previous_month_from_date(value)

                data[field_name] = value

            except Exception as e:
                print(f"  Warning: Error reading '{field_name}' from {group_name}: {e}")
                data[field_name] = None

        return data

    def _read_table(self, table_name, table_config):
        """
        Read a table from Excel and aggregate by category.

        Args:
            table_name: Name of the table
            table_config: Configuration with sheet, start_row, and columns

        Returns:
            List of dictionaries, each representing an aggregated row
        """
        sheet_name = table_config['sheet']
        start_row = table_config['start_row']
        columns = table_config['columns']

        # Read the sheet
        df = pd.read_excel(
            self.excel_path,
            sheet_name=sheet_name,
            header=None
        )

        # Find column indices by searching for header names
        # Header row is assumed to be start_row - 1
        header_row_index = start_row - 2  # Convert to 0-based index
        header_row = df.iloc[header_row_index]

        col_indices = {}
        for col_config in columns:
            col_name = col_config['name']
            col_heading = col_config['col']  # Now this is a heading name, not a letter

            # Search for the column heading in the header row
            col_idx = None
            for idx, header_value in enumerate(header_row):
                if str(header_value).strip() == col_heading:
                    col_idx = idx
                    break

            if col_idx is None:
                raise ValueError(
                    f"Column heading '{col_heading}' not found in sheet '{sheet_name}' "
                    f"at row {start_row - 1}. Available headers: {list(header_row.dropna())}"
                )

            col_indices[col_name] = col_idx

        # Read rows starting from start_row (convert to 0-based index)
        row_index = start_row - 1
        rows = []

        while row_index < len(df):
            # Read values from each column
            row_data = {}
            is_empty = True

            for col_name, col_idx in col_indices.items():
                try:
                    value = df.iloc[row_index, col_idx]
                    if not pd.isna(value):
                        is_empty = False
                    row_data[col_name] = value
                except (IndexError, KeyError):
                    row_data[col_name] = None

            # Stop if we hit an empty row (all values are NaN/None)
            if is_empty:
                break

            rows.append(row_data)
            row_index += 1

        # Aggregate if configured
        if table_config.get('aggregate', False):
            rows = self._aggregate_table_rows(rows, table_config)

        return rows

    def _aggregate_table_rows(self, rows, table_config):
        """
        Aggregate table rows by grouping column and either counting or summing.

        Args:
            rows: List of row dictionaries
            table_config: Configuration with group_by and aggregation settings

        Returns:
            List of aggregated row dictionaries
        """
        if not rows:
            return []

        group_by = table_config.get('group_by')
        if not group_by:
            return rows

        # Convert to DataFrame for easy aggregation
        df = pd.DataFrame(rows)

        # Drop rows where the group_by column is NaN
        df = df.dropna(subset=[group_by])

        # Check aggregation type: 'count' or 'sum'
        agg_type = table_config.get('aggregation_type', 'sum')
        count_column = table_config.get('count_column')
        sum_column = table_config.get('sum_column')

        if agg_type == 'count':
            # Count occurrences of each category
            if not count_column:
                raise ValueError(
                    f"Table config missing 'count_column' for aggregation_type='count'"
                )

            # Group and count
            aggregated = df.groupby(group_by).size().reset_index(name=count_column)

        elif agg_type == 'sum':
            # Sum a numeric column
            if not sum_column:
                raise ValueError(
                    f"Table config missing 'sum_column' for aggregation_type='sum'"
                )

            # Convert sum_column to numeric, replacing errors with 0
            df[sum_column] = pd.to_numeric(df[sum_column], errors='coerce').fillna(0)

            # Group and sum
            aggregated = df.groupby(group_by)[sum_column].sum().reset_index()
        else:
            raise ValueError(f"Unknown aggregation_type: '{agg_type}'. Use 'count' or 'sum'.")

        # Convert back to list of dictionaries
        result = aggregated.to_dict('records')

        # Format values if formatting is specified
        if table_config.get('format_values'):
            value_col = count_column if agg_type == 'count' else sum_column
            for row in result:
                row[value_col] = excel_utils.format_table_value(
                    row[value_col],
                    table_config.get('value_format', {})
                )

        return result

