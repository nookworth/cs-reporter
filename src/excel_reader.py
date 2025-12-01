"""
Excel file reader using pandas.
Extracts data from specified cells according to the configuration.
"""

import pandas as pd
from pathlib import Path


class ExcelReader:
    """Reads data from Excel files based on configuration mapping."""

    def __init__(self, excel_path, config):
        """
        Initialize the Excel reader.

        Args:
            excel_path: Path to the Excel file
            config: Configuration dictionary with field mappings
        """
        self.excel_path = Path(excel_path)
        self.config = config
        self.excel_fields = config.get('excel_fields', {})
        self.table_fields = config.get('table_fields', {})

    def extract_data(self):
        """
        Extract all configured fields and tables from the Excel file.

        Returns:
            Dictionary mapping field names to their values (includes tables)
        """
        data = {}

        # Extract single-cell fields
        for field_name, field_config in self.excel_fields.items():
            try:
                value = self._read_cell(
                    sheet_name=field_config['sheet'],
                    cell=field_config['cell']
                )
                data[field_name] = self._format_value(field_name, value)
            except Exception as e:
                raise ValueError(
                    f"Error reading field '{field_name}': {e}"
                )

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

    def _read_cell(self, sheet_name, cell):
        """
        Read a specific cell from a sheet using pandas.

        Args:
            sheet_name: Name of the Excel sheet
            cell: Cell reference (e.g., 'B5', 'C10')

        Returns:
            The value in the specified cell
        """
        # Read the specific sheet
        df = pd.read_excel(
            self.excel_path,
            sheet_name=sheet_name,
            header=None  # Don't treat first row as header
        )

        # Parse cell reference (e.g., 'B5' -> column 1, row 4)
        col_letter = ''.join(c for c in cell if c.isalpha()).upper()
        row_number = int(''.join(c for c in cell if c.isdigit()))

        # Convert column letter to index (A=0, B=1, etc.)
        col_index = self._column_letter_to_index(col_letter)
        row_index = row_number - 1  # pandas uses 0-based indexing

        # Get the value
        value = df.iloc[row_index, col_index]

        return value

    def _column_letter_to_index(self, letter):
        """
        Convert Excel column letter to 0-based index.
        A -> 0, B -> 1, Z -> 25, AA -> 26, etc.
        """
        result = 0
        for char in letter:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result - 1

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
                row[value_col] = self._format_table_value(
                    row[value_col],
                    table_config.get('value_format', {})
                )

        return result

    def _format_table_value(self, value, format_config):
        """
        Format a table value according to its configuration.

        Args:
            value: Raw value
            format_config: Formatting configuration

        Returns:
            Formatted value
        """
        if pd.isna(value):
            return ""

        format_type = format_config.get('type', 'number')
        decimals = format_config.get('decimals', 0)

        if format_type == 'currency':
            return f"${float(value):,.{decimals}f}"
        elif format_type == 'percentage':
            return f"{float(value):.{decimals}f}%"
        elif format_type == 'number':
            return f"{float(value):,.{decimals}f}"
        else:
            return str(value)

    def _format_value(self, field_name, value):
        """
        Format a value according to its configuration.

        Args:
            field_name: Name of the field
            value: Raw value from Excel

        Returns:
            Formatted value
        """
        formatting = self.config.get('formatting', {}).get(field_name, {})

        if pd.isna(value):
            return ""

        format_type = formatting.get('type', 'text')
        decimals = formatting.get('decimals', 2)

        if format_type == 'currency':
            return f"${float(value):,.{decimals}f}"
        elif format_type == 'percentage':
            return f"{float(value):.{decimals}f}%"
        elif format_type == 'number':
            return f"{float(value):,.{decimals}f}"
        else:
            return str(value)
