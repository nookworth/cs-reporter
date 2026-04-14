# CS Reporter
# Excel Utilities V2
# Utility functions for Excel data processing. Worker functions for excel_reader.py.
# **Date:** 2026-02-20
# **Status:** ✅ Completed
# **Component:** src

import datetime
import re
from pathlib import Path
from typing import Any

import pandas as pd


def apply_filters(df: pd.DataFrame, filters: list[dict]) -> pd.DataFrame:
    """Apply filter conditions to a DataFrame based on the filter list."""
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


def count_rows(
    excel_path: Path | str, sheet_name: str, filters: list[dict] | None = None
) -> int:
    """Count rows in a sheet, optionally applying filters."""
    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=0)
    df = df.dropna(how="all")
    df = apply_filters(df, filters)
    return len(df)


def read_column_value(excel_path: Path | str, sheet_name: str, column_name: str) -> Any:
    """Read the first non-null value from a column."""
    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=0)

    if column_name not in df.columns:
        raise ValueError(f"Column '{column_name}' not found in sheet '{sheet_name}'")

    column_data = df[column_name].dropna()
    if len(column_data) > 0:
        return column_data.iloc[0]

    return None


def parse_month_from_date(date_value: Any) -> str:
    """Extract the month name from a date value."""
    if pd.isna(date_value):
        return "Unknown"

    try:
        date = pd.to_datetime(date_value)
        return date.strftime("%B")
    except (ValueError, TypeError):
        return "Unknown"


def parse_previous_month_from_date(date_value: Any) -> str:
    """Extract the previous month name from a date value."""
    if pd.isna(date_value):
        return "Unknown"

    try:
        date = pd.to_datetime(date_value)

        prev_month_num = 12 if date.month == 1 else date.month - 1

        prev_date = date.replace(month=prev_month_num, day=1)

        return prev_date.strftime("%B")
    except (ValueError, TypeError):
        return "Unknown"


def count_column_value(
    excel_path: Path | str,
    sheet_name: str,
    column_name: str,
    value_to_count: str,
    filters: list[dict] | None = None,
) -> int:
    """Count occurrences of a specific value in a column."""
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


def count_unique_values(
    excel_path: Path | str,
    sheet_name: str,
    column_name: str,
    uncategorized_label: str | None = None,
) -> list[dict]:
    """Count unique values in a column and return as list of dicts."""
    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=0)

    if column_name not in df.columns:
        raise ValueError(f"Column '{column_name}' not found in sheet '{sheet_name}'")

    if uncategorized_label:
        df[column_name] = df[column_name].fillna(uncategorized_label)

        df[column_name] = df[column_name].apply(
            lambda x: (
                uncategorized_label
                if isinstance(x, str) and (x == "" or re.fullmatch(r"[\s]+", x))
                else x
            )
        )

        value_counts = df[column_name].value_counts()
    else:
        value_counts = df[column_name].dropna().value_counts()

    result = []
    for value, count in value_counts.items():
        result.append({"value": str(value), "count": int(count)})

    return result


def sum_column(
    excel_path: Path | str,
    sheet_name: str,
    column_name: str,
    filters: list[dict] | None = None,
) -> float:
    """Sum values in a column, optionally applying filters."""
    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=0)

    if column_name not in df.columns:
        raise ValueError(f"Column '{column_name}' not found in sheet '{sheet_name}'")

    df = apply_filters(df, filters)
    return df[column_name].sum()


def format_table_value(value: Any, format_config: dict[str, Any]) -> str:
    """Format a value according to a format configuration (currency, percentage, number)."""
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
    excel_path: Path | str,
    sheet_name: str,
    start_column: str = "Ticket created - Date",
    end_column: str = "Ticket solved - Date",
    filters: list[dict] | None = None,
) -> float:
    """Calculate average days between start and end columns."""
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
