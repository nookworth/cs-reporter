"""
Utility functions for Excel data processing.
"""

import datetime
import re
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
    return len(df.dropna(how="all"))


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


def count_unique_values(excel_path, sheet_name, column_name, uncategorized_label=None):
    """
    Count all unique values in a column.

    Args:
        excel_path: Path to the Excel file
        sheet_name: Name of the Excel sheet
        column_name: Column header name
        uncategorized_label: If provided, null/empty values will be grouped under this label

    Returns:
        List of dicts with value and count, e.g.:
        [{"value": "Technical", "count": 10}, {"value": "Billing", "count": 5}, ...]
    """
    # Read the sheet with first row as header
    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=0)

    # Find the column by name
    if column_name not in df.columns:
        raise ValueError(f"Column '{column_name}' not found in sheet '{sheet_name}'")

    # Handle uncategorized values
    if uncategorized_label:
        # Replace NaN with the uncategorized label
        df[column_name] = df[column_name].fillna(uncategorized_label)

        # Replace empty strings and whitespace-only strings with the uncategorized label
        df[column_name] = df[column_name].apply(
            lambda x: uncategorized_label
            if isinstance(x, str) and (x == '' or re.fullmatch(r'[\s]+', x))
            else x
        )

        # Count all values including the uncategorized ones
        value_counts = df[column_name].value_counts()
    else:
        # Drop NaN values as before
        value_counts = df[column_name].dropna().value_counts()

    # Convert to list of dicts
    result = []
    for value, count in value_counts.items():
        result.append({"value": str(value), "count": int(count)})

    return result


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

    format_type = format_config.get("type", "number")
    decimals = format_config.get("decimals", 0)

    if format_type == "currency":
        return f"${float(value):,.{decimals}f}"
    elif format_type == "percentage":
        return f"{float(value):.{decimals}f}%"
    elif format_type == "number":
        return f"{float(value):,.{decimals}f}"
    else:
        return str(value)


def calculate_average_resolution_time(
    excel_path,
    sheet_name,
    columns=["Ticket created - Date", "Ticket solved - Date"],
    assignee_column="Assignee name"
):
    """
    Calculates the average resolution time with assignee-based filtering.

    Formula: "Ticket solved - Date" - "Ticket created - Date"
    Calculated in hours for precision, then converted to days for output.

    Filtering logic:
    - Drops rows with no "Ticket solved - Date"
    - Excludes tickets where resolution > 72 hours AND assignee is "Leo Brown"
    - Includes all other tickets regardless of resolution time

    Args:
        excel_path: Path to the Excel file
        sheet_name: Name of the Excel sheet
        columns: List of [start_date_column, end_date_column]
        assignee_column: Column name containing assignee information

    Returns:
        Average resolution time in days (rounded to 2 decimals), or 0 if no valid data
    """

    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=0)

    # Validate that assignee column exists
    if assignee_column not in df.columns:
        raise ValueError(
            f"Assignee column '{assignee_column}' not found in sheet '{sheet_name}'"
        )

    # Create a working dataframe with the necessary columns
    work_df = df[[columns[0], columns[1], assignee_column]].copy()

    # Drop rows with no "Ticket solved - Date" (requirement #1)
    work_df = work_df.dropna(subset=[columns[1]])

    # Initialize counters
    total_tickets_included = 0
    total_resolution_time_hours = 0

    # Iterate through rows
    for idx, row in work_df.iterrows():
        start_date_str = row[columns[0]]
        end_date_str = row[columns[1]]
        assignee = row[assignee_column]

        # Skip if start date is NaN or whitespace
        if pd.isna(start_date_str):
            continue
        if isinstance(start_date_str, str) and re.fullmatch(r'[\s]+', start_date_str):
            continue

        # Skip if end date is whitespace (NaN already filtered by dropna)
        if isinstance(end_date_str, str) and re.fullmatch(r'[\s]+', end_date_str):
            continue

        # Calculate resolution time in hours
        try:
            diff = datetime.date.fromisoformat(end_date_str) - datetime.date.fromisoformat(
                start_date_str
            )
            # Convert to hours (timedelta.days * 24 hours)
            resolution_time_hours = diff.days * 24
        except (ValueError, TypeError):
            continue  # Skip invalid dates

        # NEW FILTERING LOGIC (requirement #3)
        # Exclude tickets where: resolution > 72 hours AND assignee is Leo Brown
        if resolution_time_hours > 72 and assignee == "Leo Brown":
            continue  # Skip this ticket

        # Include all other tickets (requirement #4)
        total_tickets_included += 1
        total_resolution_time_hours += resolution_time_hours

    # Avoid division by zero if no valid data
    if total_tickets_included == 0:
        return 0

    # Calculate average in hours, then convert to days
    average_hours = total_resolution_time_hours / total_tickets_included
    average_days = average_hours / 24

    return round(average_days, 2)


# for testing from command line
if __name__ == "__main__":
    calculate_average_resolution_time(
        excel_path="~/Downloads/ADUS Monthly review (21).xlsx",
        sheet_name="Tickets ADUS Tickets crea... 1",
    )
