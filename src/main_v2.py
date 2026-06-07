# CS Reporter
# Main V2
# Main entry point for the CS Reporter utility. Automates the creation of monthly CS reports.
# **Date:** 2026-02-20
# **Status:** ✅ Completed
# **Component:** src

import argparse
import sys
import tkinter as tk
from tkinter import filedialog

from .config_v2 import load_config
from .excel_reader_v2 import ExcelReader
from .ppt_writer import PowerPointWriter


def parse_args():
    parser = argparse.ArgumentParser(
        description="CS Reporter - Excel to PowerPoint Report Generator"
    )
    parser.add_argument(
        "--config",
        "-c",
        help="Path to config file (default: config/mapping_v2.yaml)",
        default=None,
    )
    return parser.parse_args()


def select_excel_file(title="Select Excel Input File"):
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        title=title, filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )

    root.destroy()
    return file_path if file_path else None


def main():
    args = parse_args()

    print("CS Reporter - Excel to PowerPoint Report Generator")
    print("=" * 60)

    try:
        config = load_config(config_path=args.config)
        print(f"[OK] Configuration loaded from: {config['config_path']}")
    except Exception as e:
        print(f"[ERROR] Error loading configuration: {e}")
        return 1

    print("\nPlease select CURRENT MONTH Excel file...")
    current_excel_path = select_excel_file("Select Current Month Excel File")

    if not current_excel_path:
        print("No file selected. Exiting.")
        return 0

    print(f"[OK] Selected: {current_excel_path}")

    print("\nPlease select PREVIOUS MONTH Excel file...")
    previous_excel_path = select_excel_file("Select Previous Month Excel File")

    if not previous_excel_path:
        print("No file selected. Exiting.")
        return 0

    print(f"[OK] Selected: {previous_excel_path}")

    try:
        reader = ExcelReader(
            current_excel_path, config, previous_excel_path=previous_excel_path
        )
        data = reader.extract_data()
        print(f"[OK] Extracted {len(data)} fields from Excel files")
    except Exception as e:
        print(f"[ERROR] Error reading Excel files: {e}")
        return 1

    try:
        writer = PowerPointWriter(config)
        output_path = writer.generate_report(data)
        print(f"[OK] Report generated: {output_path}")
        print("\nDone!")
        return 0
    except Exception as e:
        print(f"[ERROR] Error generating PowerPoint: {e}")
        return 1


def cli_entry_point():
    sys.exit(main())


if __name__ == "__main__":
    cli_entry_point()
