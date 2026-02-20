# CS Reporter Configuration Schema
# Utility functions for Excel data processing.
# This module contains the low-level Excel operations that actually read data.
# These are the "worker functions" that excel_reader.py delegates to.
# **Date:** 2026-02-20
# **Status:** ✅ Completed
# **Component:** src

"""
KEY FUNCTIONS:
1. count_rows() - Count total tickets/rows in a sheet
2. read_column_value() - Read the first value from a column
3. parse_month_from_date() - Extract month name from a date
4. count_column_value() - Count occurrences of a specific value
5. count_unique_values() - Get all unique values and their counts
6. calculate_average_resolution_time() - Complex calculation with filtering

All functions use pandas for Excel file reading and data manipulation.
"""

import datetime
import re

import pandas as pd  # Pandas handles Excel file reading


def apply_filters(df, filters):

    if not filters:
        return df
    
    for f in filters:
        col = f["column"]
        op = f["operator"]
        val = f.get("value")
        
        if op == "equals":
            df = df[df[col].astype(str).str.lower() == str(val).lower()]
        elif op == "not_equals":
            df = df[df[col].astype(str).str.lower() != str(val).lower()]
        elif op == "greater_than":
            df = df[df[col] > val]
        elif op == "greater_than_or_equal":
            df = df[df[col] >= val]
        elif op == "less_than":
            df = df[df[col] < val]
        elif op == "less_than_or_equal":
            df = df[df[col] <= val]
        elif op == "contains":
            df = df[df[col].astype(str).str.contains(str(val), case=False, na=False)]
        elif op == "not_contains":
            df = df[~df[col].astype(str).str.contains(str(val), case=False, na=False)]
        elif op == "starts_with":
            df = df[df[col].astype(str).str.startswith(str(val), na=False)]
        elif op == "ends_with":
            df = df[df[col].astype(str).str.endswith(str(val), na=False)]
        elif op == "is_null":
            df = df[df[col].isna()]
        elif op == "is_not_null":
            df = df[df[col].notna()]
    
    return df


def count_rows(excel_path, sheet_name, filters=None):

    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=0)
    df = df.dropna(how="all")
    df = apply_filters(df, filters)
    return len(df)


def read_column_value(excel_path, sheet_name, column_name):

    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=0)

    if column_name not in df.columns:
        raise ValueError(f"Column '{column_name}' not found in sheet '{sheet_name}'")

    column_data = df[column_name].dropna()
    if len(column_data) > 0:
        return column_data.iloc[0]  # iloc[0] gets the first row

    return None


def parse_month_from_date(date_value):

    # Handle NaN/null values
    if pd.isna(date_value):
        return "Unknown"

    try:
        date = pd.to_datetime(date_value)
        return date.strftime("%B")  # %B = full month name
    except (ValueError, TypeError):
        return "Unknown"


def parse_previous_month_from_date(date_value):

    if pd.isna(date_value):
        return "Unknown"

    try:

        date = pd.to_datetime(date_value)

        if date.month == 1:  # noqa: SIM108
            prev_month_num = 12
        else:
            prev_month_num = date.month - 1

        prev_date = date.replace(month=prev_month_num, day=1)

        return prev_date.strftime("%B")
    except (ValueError, TypeError):
        return "Unknown"


def count_column_value(excel_path, sheet_name, column_name, value_to_count, filters=None):

    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=0)
    
    if column_name not in df.columns:
        raise ValueError(f"Column '{column_name}' not found in sheet '{sheet_name}'")
    
    df = apply_filters(df, filters)
    column_data = df[column_name].dropna()
    
    count = 0
    for val in column_data:
        if isinstance(val, str) and val.lower() == value_to_count.lower():
            count += 1
    
    return count


def count_unique_values(excel_path, sheet_name, column_name, uncategorized_label=None):

    # Read the sheet with first row as header
    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=0)

    # Validate column exists
    if column_name not in df.columns:
        raise ValueError(f"Column '{column_name}' not found in sheet '{sheet_name}'")

    # Handle uncategorized values (null/empty cells)
    if uncategorized_label:
        # Replace NaN (missing values) with the uncategorized label
        df[column_name] = df[column_name].fillna(uncategorized_label)

        df[column_name] = df[column_name].apply(
            lambda x: (
                uncategorized_label
                if isinstance(x, str) and (x == "" or re.fullmatch(r"[\s]+", x))
                else x
            )
        )

        # Count all values including the uncategorized ones
        value_counts = df[column_name].value_counts()
    else:
        # Drop NaN values if no uncategorized label specified
        value_counts = df[column_name].dropna().value_counts()

    result = []
    for value, count in value_counts.items():
        result.append({"value": str(value), "count": int(count)})

    return result


def sum_column(excel_path, sheet_name, column_name, filters=None):

    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=0)
    
    if column_name not in df.columns:
        raise ValueError(f"Column '{column_name}' not found in sheet '{sheet_name}'")
    
    df = apply_filters(df, filters)
    return df[column_name].sum()


def format_table_value(value, format_config):
    # Handle null/empty values
    if pd.isna(value):
        return ""

    # Get formatting parameters
    format_type = format_config.get("type", "number")
    decimals = format_config.get("decimals", 0)

    # Apply formatting based on type
    if format_type == "currency":
        return f"${float(value):,.{decimals}f}"
    elif format_type == "percentage":
        return f"{float(value):.{decimals}f}%"
    elif format_type == "number":
        return f"{float(value):,.{decimals}f}"
    else:
        # Unknown type - just convert to string
        return str(value)


def calculate_average_resolution_time(
    excel_path,
    sheet_name,
    start_column="Ticket created - Date",
    end_column="Ticket solved - Date",
    filters=None,
):

    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=0)
    df = df.dropna(subset=[end_column])
    df = apply_filters(df, filters)
    
    total_days = 0
    count = 0
    
    for _, row in df.iterrows():
        start = row[start_column]
        end = row[end_column]
        
        if pd.isna(start):
            continue
        if isinstance(start, str) and re.fullmatch(r"[\s]+", start):
            continue
        if isinstance(end, str) and re.fullmatch(r"[\s]+", end):
            continue
        
        try:
            diff = datetime.date.fromisoformat(end) - datetime.date.fromisoformat(start)
            total_days += diff.days
            count += 1
        except (ValueError, TypeError):
            continue
    
    return round(total_days / count, 2) if count > 0 else 0


# For testing from command line
if __name__ == "__main__":
    # Example test case (update path as needed)
    calculate_average_resolution_time(
        excel_path="~/Downloads/ADUS Monthly review (21).xlsx",
        sheet_name="Tickets ADUS Tickets crea... 1",
    )
