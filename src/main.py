#!/usr/bin/env python3
"""
Main entry point for the CS Reporter utility.
Opens a file dialog to select an Excel file, then generates a PowerPoint report.
"""

import sys
import tkinter as tk
from tkinter import filedialog

from .excel_reader import ExcelReader
from .ppt_writer import PowerPointWriter
from .config import load_config


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

    # Select current month Excel file
    print("\nPlease select CURRENT MONTH Excel file...")
    current_excel_path = select_excel_file("Select Current Month Excel File")

    if not current_excel_path:
        print("No file selected. Exiting.")
        return 0

    print(f"✓ Selected: {current_excel_path}")

    # Select previous month Excel file
    print("\nPlease select PREVIOUS MONTH Excel file...")
    previous_excel_path = select_excel_file("Select Previous Month Excel File")

    if not previous_excel_path:
        print("No file selected. Exiting.")
        return 0

    print(f"✓ Selected: {previous_excel_path}")

    # Read data from both Excel files
    try:
        reader = ExcelReader(current_excel_path, config, previous_excel_path=previous_excel_path)
        data = reader.extract_data()
        print(f"✓ Extracted {len(data)} fields from Excel files")
    except Exception as e:
        print(f"✗ Error reading Excel files: {e}")
        return 1

    # WIP: History management (disabled for now)
    # This code is kept for future use but not currently active
    # history = HistoryManager()
    # try:
    #     month_str = history.get_month_from_data(data)
    #     if month_str:
    #         history_path = history.save_data(data, month_str)
    #         print(f"✓ Saved history: {history_path}")
    # except Exception as e:
    #     print(f"⚠ Warning: Could not save history: {e}")

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
