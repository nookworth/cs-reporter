#!/usr/bin/env python3
"""
Main entry point for the CS Reporter utility.

HIGH-LEVEL OVERVIEW:
This application automates the creation of monthly CS (Customer Support) reports by:
1. Loading a configuration file that defines what data to extract and where
2. Prompting the user to select TWO Excel files (current month + previous month)
3. Extracting specific data fields from both Excel files
4. Generating a PowerPoint presentation by filling in a template with the extracted data

The workflow is: Config → Excel Selection → Data Extraction → PowerPoint Generation
"""

import argparse
import sys
import tkinter as tk
from tkinter import filedialog

from .config import load_config  # Loads YAML configuration
from .excel_reader import ExcelReader  # Extracts data from Excel files
from .ppt_writer import PowerPointWriter  # Generates PowerPoint from template


def parse_args():
    """
    Parse command-line arguments.
    
    Allows the user to optionally specify a custom configuration file.
    If not specified, defaults to config/mapping.yaml
    
    Example usage:
        python -m src.main --config config/demo_mapping.yaml
    """
    parser = argparse.ArgumentParser(
        description="CS Reporter - Excel to PowerPoint Report Generator"
    )
    parser.add_argument(
        "--config",
        "-c",
        help="Path to config file (default: config/mapping.yaml)",
        default=None,
    )
    return parser.parse_args()



def select_excel_file(title="Select Excel Input File"):
    """
    Open a file dialog to let the user select an Excel file.
    
    Uses Tkinter's file dialog to provide a graphical file picker.
    The main window is hidden to avoid showing an empty Tk window.
    
    Args:
        title: Window title for the file dialog
        
    Returns:
        The selected file path (string) or None if the user canceled
    """
    root = tk.Tk()
    root.withdraw()  # Hide the main Tkinter window (we only want the dialog)

    file_path = filedialog.askopenfilename(
        title=title, filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )

    root.destroy()  # Clean up the Tkinter instance
    return file_path if file_path else None

# Main function
def main():
    """
    Main execution function.
    
    EXECUTION FLOW:
    1. Parse command-line arguments (for custom config path)
    2. Load configuration (defines what data to extract and where to put it)
    3. Prompt user to select CURRENT month Excel file
    4. Prompt user to select PREVIOUS month Excel file
    5. Extract data from both files using the configuration mappings
    6. Generate PowerPoint report by filling in template placeholders
    
    Returns:
        0 on success, 1 on error
    """
    args = parse_args()

    print("CS Reporter - Excel to PowerPoint Report Generator")
    print("=" * 60)

    # STEP 1: Load configuration
    # The config file (YAML) defines:
    # - Which Excel sheets to read from
    # - Which cells/columns contain the data
    # - Where to place that data in the PowerPoint template
    try:
        config = load_config(config_path=args.config)
        print(f"✓ Configuration loaded from: {config['config_path']}")
    except Exception as e:
        print(f"✗ Error loading configuration: {e}")
        return 1

    # STEP 2: Select current month Excel file
    # This file contains the latest month's customer support data
    print("\nPlease select CURRENT MONTH Excel file...")
    current_excel_path = select_excel_file("Select Current Month Excel File")

    if not current_excel_path:
        print("No file selected. Exiting.")
        return 0

    print(f"✓ Selected: {current_excel_path}")

    # STEP 3: Select previous month Excel file
    # This file is needed for month-to-month comparisons
    # (e.g., "ticket count increased from 45 last month to 62 this month")
    print("\nPlease select PREVIOUS MONTH Excel file...")
    previous_excel_path = select_excel_file("Select Previous Month Excel File")

    if not previous_excel_path:
        print("No file selected. Exiting.")
        return 0

    print(f"✓ Selected: {previous_excel_path}")

    # STEP 4: Read data from both Excel files
    # The ExcelReader extracts specific values based on the configuration
    # It knows which file to use for each field (current vs previous month)
    try:
        reader = ExcelReader(
            current_excel_path, config, previous_excel_path=previous_excel_path
        )
        data = reader.extract_data()
        print(f"✓ Extracted {len(data)} fields from Excel files")
    except Exception as e:
        print(f"✗ Error reading Excel files: {e}")
        return 1

    # WIP: History management (disabled for now)
    # This code is kept for future use but not currently active
    # The idea is to automatically save historical data for trend analysis
    # history = HistoryManager()
    # try:
    #     month_str = history.get_month_from_data(data)
    #     if month_str:
    #         history_path = history.save_data(data, month_str)
    #         print(f"✓ Saved history: {history_path}")
    # except Exception as e:
    #     print(f"⚠ Warning: Could not save history: {e}")

    # STEP 5: Generate PowerPoint
    # The PowerPointWriter takes the extracted data and fills in
    # a PowerPoint template by replacing {{placeholder}} tags
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
