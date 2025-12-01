#!/usr/bin/env python3
"""
Main entry point for the CS Reporter utility.
Opens a file dialog to select an Excel file, then generates a PowerPoint report.
"""

import sys
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
from datetime import datetime

from .excel_reader import ExcelReader
from .ppt_writer import PowerPointWriter
from .config import load_config


def select_excel_file():
    """
    Open a file dialog to let the user select an Excel file.
    Returns the selected file path or None if canceled.
    """
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    file_path = filedialog.askopenfilename(
        title="Select Excel Input File",
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

    # Select Excel file
    print("\nPlease select an Excel file...")
    excel_path = select_excel_file()

    if not excel_path:
        print("No file selected. Exiting.")
        return 0

    print(f"✓ Selected: {excel_path}")

    # Read data from Excel
    try:
        reader = ExcelReader(excel_path, config)
        data = reader.extract_data()
        print(f"✓ Extracted {len(data)} fields from Excel")
    except Exception as e:
        print(f"✗ Error reading Excel file: {e}")
        return 1

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
