"""
Utility functions for Excel data processing.
"""

import pandas as pd


def count_rows(excel_path, sheet_name):
    """
    Count the total number of data rows in a sheet (excluding header).

    Args:
        excel_path: Path to the Excel file
        sheet_name: Name of the Excel sheet

    Returns:
        Number of rows (excluding header and empty rows)
    """
    # Read the sheet with first row as header
    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=0)

    # Count non-empty rows (drop rows where all values are NaN)
    return len(df.dropna(how='all'))


def read_column_value(excel_path, sheet_name, column_name):
    """
    Read the first non-null value from a column (identified by header name).

    Args:
        excel_path: Path to the Excel file
        sheet_name: Name of the Excel sheet
        column_name: Column header name

    Returns:
        The first non-null value in the column
    """
    # Read the sheet with first row as header
    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=0)

    # Find the column by name
    if column_name not in df.columns:
        raise ValueError(f"Column '{column_name}' not found in sheet '{sheet_name}'")

    # Get first non-null value
    column_data = df[column_name].dropna()
    if len(column_data) > 0:
        return column_data.iloc[0]

    return None


def parse_month_from_date(date_value):
    """
    Parse a date value and return the month name.

    Args:
        date_value: Date value (string, datetime, or pandas Timestamp)

    Returns:
        Month name (e.g., "March")
    """
    if pd.isna(date_value):
        return "Unknown"

    try:
        # Parse the date using pandas
        date = pd.to_datetime(date_value)
        # Return month name
        return date.strftime("%B")
    except:
        return "Unknown"


def parse_previous_month_from_date(date_value):
    """
    Parse a date value, subtract one month, and return the month name.

    Args:
        date_value: Date value (string, datetime, or pandas Timestamp)

    Returns:
        Previous month name (e.g., "February")
    """
    if pd.isna(date_value):
        return "Unknown"

    try:
        # Parse the date using pandas
        date = pd.to_datetime(date_value)

        # Calculate previous month
        if date.month == 1:
            # January -> December
            prev_month_num = 12
        else:
            prev_month_num = date.month - 1

        # Create a date with the previous month number
        # Use day=1 to avoid issues with different month lengths
        prev_date = date.replace(month=prev_month_num, day=1)

        # Return month name
        return prev_date.strftime("%B")
    except:
        return "Unknown"


def count_column_value(excel_path, sheet_name, column_name, value_to_count):
    """
    Count occurrences of a specific value in a column.

    Args:
        excel_path: Path to the Excel file
        sheet_name: Name of the Excel sheet
        column_name: Column header name
        value_to_count: The value to count (case-insensitive string comparison)

    Returns:
        Number of occurrences of the value
    """
    # Read the sheet with first row as header
    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=0)

    # Find the column by name
    if column_name not in df.columns:
        raise ValueError(f"Column '{column_name}' not found in sheet '{sheet_name}'")

    # Get the column data and drop NaN values
    column_data = df[column_name].dropna()

    # Count occurrences (case-insensitive)
    count = 0
    for val in column_data:
        if isinstance(val, str) and val.lower() == value_to_count.lower():
            count += 1

    return count


def format_table_value(value, format_config):
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
