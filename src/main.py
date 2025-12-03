#!/usr/bin/env python3
"""
Main entry point for the CS Reporter utility.
Opens a file dialog to select an Excel file, then generates a PowerPoint report.
"""

import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from datetime import datetime

from .excel_reader import ExcelReader
from .ppt_writer import PowerPointWriter
from .config import load_config
from .history import HistoryManager


def select_excel_file(title="Select Excel Input File"):
    """
    Open a file dialog to let the user select an Excel file.
    Returns the selected file path or None if canceled.
    """
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    file_path = filedialog.askopenfilename(
        title=title,
        filetypes=[
            ("Excel files", "*.xlsx *.xls"),
            ("All files", "*.*")
        ]
    )

    root.destroy()
    return file_path if file_path else None


def main():
    """Main execution function."""
    print("CS Reporter - Excel to PowerPoint Report Generator")
    print("=" * 60)

    # Load configuration
    try:
        config = load_config()
        print(f"✓ Configuration loaded from: {config['config_path']}")
    except Exception as e:
        print(f"✗ Error loading configuration: {e}")
        return 1

    # Initialize history manager
    history = HistoryManager()

    # Select current month Excel file
    print("\nPlease select CURRENT MONTH Excel file...")
    current_excel_path = select_excel_file("Select Current Month Excel File")

    if not current_excel_path:
        print("No file selected. Exiting.")
        return 0

    print(f"✓ Selected: {current_excel_path}")

    # Try to load previous month data from history
    print("\nLooking for previous month data...")
    previous_month_data = history.load_previous_month()

    if previous_month_data:
        prev_month = previous_month_data.get('month', 'Unknown')
        print(f"✓ Loaded previous month data: {prev_month}")
    else:
        print("✗ No history found")

        # Check if config requires previous month data
        has_prev_month_fields = bool(config.get('previous_month_fields'))

        if has_prev_month_fields:
            print("\nPlease select PREVIOUS MONTH Excel file...")
            print("(Press Cancel to skip previous month data)")

            prev_excel_path = select_excel_file("Select Previous Month Excel File (Optional)")

            if prev_excel_path:
                print(f"✓ Selected: {prev_excel_path}")

                # Extract data from previous month Excel
                try:
                    prev_reader = ExcelReader(prev_excel_path, config)
                    previous_month_data = prev_reader.extract_data()
                    print(f"✓ Extracted previous month data")
                except Exception as e:
                    print(f"✗ Error reading previous month Excel: {e}")
                    print("  Continuing without previous month data...")
                    previous_month_data = {}
            else:
                print("  Skipping previous month data")
                previous_month_data = {}

    # Read data from current month Excel
    try:
        reader = ExcelReader(current_excel_path, config, previous_month_data)
        data = reader.extract_data()
        print(f"✓ Extracted {len(data)} fields from current month Excel")
    except Exception as e:
        print(f"✗ Error reading Excel file: {e}")
        return 1

    # Save current month data to history
    try:
        month_str = history.get_month_from_data(data)
        if month_str:
            history_path = history.save_data(data, month_str)
            print(f"✓ Saved history: {history_path}")
        else:
            # Fall back to current date
            history_path = history.save_data(data)
            print(f"✓ Saved history: {history_path}")
    except Exception as e:
        print(f"⚠ Warning: Could not save history: {e}")
        print("  (Report will still be generated)")

    # Generate PowerPoint
    try:
        writer = PowerPointWriter(config)
        output_path = writer.generate_report(data)
        print(f"✓ Report generated: {output_path}")
        print("\nDone!")
        return 0
    except Exception as e:
        print(f"✗ Error generating PowerPoint: {e}")
        return 1


def cli_entry_point():
    """Entry point for the 'reporter' command."""
    sys.exit(main())


if __name__ == "__main__":
    cli_entry_point()
