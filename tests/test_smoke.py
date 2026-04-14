"""
End-to-end smoke test for the V2 pipeline using scrubbed demo data.

Runs the full ExcelReader -> PowerPointWriter pipeline against
demo_data/demo_december_2025.xlsx (current) and
demo_data/demo_january_2026.xlsx (previous), then asserts extracted
field values match the known baseline.

Usage:
    .venv/bin/python tests/test_smoke.py       # standalone
    .venv/bin/python -m pytest tests/test_smoke.py -v
"""

import sys
from pathlib import Path

# Allow running as a standalone script from the project root
sys.path.insert(0, str(Path(__file__).parent.parent))

from src.config_v2 import load_config
from src.excel_reader_v2 import ExcelReader
from src.ppt_writer import PowerPointWriter

CURRENT_EXCEL = Path("demo_data/demo_january_2026.xlsx")
PREVIOUS_EXCEL = Path("demo_data/demo_december_2025.xlsx")
CONFIG_PATH = "config/demo_mapping_v2.yaml"

# Baseline values derived from the demo data files on 2026-04-12.
# If these change, the demo data or extraction logic has changed.
EXPECTED_SCALARS = {
    "month": "January",
    "prev_month": "December",
    "re_req": 92,
    "re_sat": 36,
    "re_sat_c": 26,
    "re_reso": 2.18,
    "re_prev_req": 78,
    "re_prev_sat": 26,
    "re_prev_sat_c": 22,
    "re_prev_reso": 1.88,
    "su_req": 48,
    "su_sat": 8,
    "su_sat_c": 12,
    "su_reso": 2.0,
    "su_prev_req": 41,
    "su_prev_sat": 11,
    "su_prev_sat_c": 14,
    "su_prev_reso": 1.89,
}

EXPECTED_TABLE_LENGTHS = {
    "re_sup_cat": 8,
    "su_sup_cat": 8,
    "top_orgs": 5,
}


def test_extraction():
    config = load_config(CONFIG_PATH)
    reader = ExcelReader(str(CURRENT_EXCEL), config, str(PREVIOUS_EXCEL))
    data = reader.extract_data()

    failures = []

    for field, expected in EXPECTED_SCALARS.items():
        actual = data.get(field)
        if actual != expected:
            failures.append(f"  {field}: expected {expected!r}, got {actual!r}")

    for table_name, expected_len in EXPECTED_TABLE_LENGTHS.items():
        actual = data.get(table_name, [])
        if not isinstance(actual, list) or len(actual) != expected_len:
            failures.append(
                f"  {table_name}: expected {expected_len} rows, got {len(actual) if isinstance(actual, list) else actual!r}"
            )

    if failures:
        raise AssertionError("Extraction failures:\n" + "\n".join(failures))


def test_report_generation():
    config = load_config(CONFIG_PATH)
    reader = ExcelReader(str(CURRENT_EXCEL), config, str(PREVIOUS_EXCEL))
    data = reader.extract_data()

    writer = PowerPointWriter(config)
    output_path = Path(writer.generate_report(data))

    assert output_path.exists(), f"Output file not created: {output_path}"
    assert output_path.stat().st_size > 0, f"Output file is empty: {output_path}"

    # Clean up generated file after test
    output_path.unlink()


if __name__ == "__main__":
    errors = []

    print("Running smoke tests against demo data...")

    try:
        test_extraction()
        print("  [OK] Extraction — all scalar fields and table lengths match baseline")
    except AssertionError as e:
        print(f"  [FAIL] Extraction\n{e}")
        errors.append(e)

    try:
        test_report_generation()
        print("  [OK] Report generation — output .pptx created and non-empty")
    except Exception as e:
        print(f"  [FAIL] Report generation: {e}")
        errors.append(e)

    if errors:
        print("\nSmoke tests FAILED.")
        sys.exit(1)

    print("\nAll smoke tests passed.")
    sys.exit(0)
