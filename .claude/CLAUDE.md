# CS Reporter - Project Status

**Last Updated:** January 3, 2026
**Status:** Dual-file architecture implemented, satisfaction fields working

---

## Project Overview

CS Reporter is a Python utility that:
- Reads data from TWO Excel files (current month and previous month)
- Generates PowerPoint presentations using a template
- Supports dynamic tables with aggregation
- Provides month-over-month comparison by reading both files

**Primary Use Case:** Monthly support ticket reporting with automatic comparison

**Architecture:** Always requires both current and previous month Excel files. No history management - simpler for development and debugging.

---

## Recent Changes

### January 3, 2026 - Dual-File Architecture

**Major architectural change:** Simplified to always require both Excel files instead of history management.

1. **Always Require Both Files**
   - User now selects BOTH current and previous month Excel files every time
   - No more history loading/saving (kept as WIP code for future)
   - Much simpler for development and debugging

2. **Smart Field Routing**
   - Fields with `_prev_` in the name automatically read from previous month file
   - Regular fields read from current month file
   - Example: `re_req` → current file, `re_prev_req` → previous file

3. **Implemented Satisfaction Fields**
   - All `*_sat` and `*_sat_c` fields now working
   - Count "good" and "good with comment" from Excel columns
   - Case-insensitive matching

### December 29, 2025 - Initial Restructuring

### 🔄 Major Restructuring

1. **Configuration Structure Changed**
   - **OLD:** Used `excel_fields`, `current_month_fields`, `previous_month_fields` with operation-based approach
   - **NEW:** Organized into `standard_excel_fields`, `retail_excel_fields`, `supplier_excel_fields`
   - Simpler, more intuitive grouping by category
   - Field-specific logic based on naming conventions

2. **Removed Legacy Systems**
   - Removed operation-based field extraction (`parse_month`, `count_rows`, `avg_date_diff`, `count_value` operations)
   - Removed previous month history management (will re-implement if needed)
   - Simplified `excel_reader.py` by removing unused methods

3. **New Utility Module**
   - Created `src/excel_utils.py` for all Excel helper functions
   - Separates utility functions from main reader logic
   - Easier to extend with new operations

### ✅ Currently Implemented Features

1. **Field Extraction by Category**
   - `standard_excel_fields`: Common fields like month, prev_month
   - `retail_excel_fields`: Retail-specific metrics (re_req, re_sat, etc.)
   - `supplier_excel_fields`: Supplier-specific metrics (su_req, su_sat, etc.)
   - Each category has a default sheet that fields inherit

2. **Smart Field Processing**
   - **Month fields**: Automatically parsed from date columns (e.g., "December")
   - **Previous month**: Automatically calculated (e.g., "November")
   - **Row count fields**: Fields ending in `_req` auto-count rows
   - **Column reading**: Reads first non-null value from named columns

3. **Dynamic Tables with Aggregation**
   - Auto-populate PowerPoint tables with variable rows
   - Count aggregation: Group and count occurrences
   - Sum aggregation: Group and sum numeric values
   - Column header search (no hardcoded column letters)

4. **Template-Based PowerPoint Generation**
   - Placeholder replacement: `{{field_name}}` → actual values
   - Dynamic table population: `{{table:table_name}}` → multiple rows
   - Formatting support: currency, percentages, numbers

5. **CLI with File Dialogs**
   - Simple `reporter` command
   - Automatic file dialogs for Excel selection
   - Clear progress indicators

### 🔧 Technical Improvements

1. **Fixed Installation Issues**
   - Resolved tkinter dependency on macOS (Homebrew Python)
   - Recreated virtual environment with proper Python version

2. **Code Organization**
   - Utility functions separated into `excel_utils.py`
   - Cleaner `excel_reader.py` focused on orchestration
   - Better separation of concerns

---

## Project Structure

```
cs-reporter/
├── src/
│   ├── __init__.py
│   ├── main.py              # CLI entry point with file dialogs
│   ├── excel_reader.py      # Excel field extraction orchestration
│   ├── excel_utils.py       # Excel utility functions (NEW)
│   ├── ppt_writer.py        # PowerPoint generation
│   ├── config.py            # Configuration loader
│   └── history.py           # History management
├── templates/
│   └── report_template.pptx # PowerPoint template
├── config/
│   ├── mapping.yaml         # Current configuration
│   └── mapping.yaml.example # Reference example (outdated)
├── output/
│   ├── .gitkeep
│   └── history/             # Auto-saved monthly data
│       └── .gitkeep
├── requirements.txt
├── setup.py
├── README.md
└── PROJECT_STATUS.md        # This file
```

---

## Configuration Format

Your `config/mapping.yaml` now has four main sections:

### 1. Standard Excel Fields

Common fields shared across sheets:

```yaml
standard_excel_fields:
  sheet: "Tickets ADUS Tickets crea... 1"  # Default sheet for this group

  month:
    cell: "Ticket created - Date"  # Column header name

  prev_month:
    cell: "Ticket created - Date"  # Automatically calculates previous month
```

### 2. Retail Excel Fields

Retail-specific metrics:

```yaml
retail_excel_fields:
  sheet: "Tickets ADUS Tickets crea... 1"  # Default sheet

  re_req:
    # No cell needed - auto-counts rows (field ends with _req)

  re_sat:
    cell: "Ticket satisfaction rating"  # Column to read from

  re_sat_c:
    cell: "Ticket satisfaction rating"
```

### 3. Supplier Excel Fields

Supplier-specific metrics:

```yaml
supplier_excel_fields:
  sheet: "Tickets ADUS Tickets crea... 2"  # Default sheet

  su_req:
    # No cell needed - auto-counts rows

  su_sat:
    cell: "Ticket satisfaction rating"
```

### 4. Dynamic Tables

```yaml
table_fields:
  re_sup_cat:
    sheet: "Tickets ADUS Tickets crea...1"
    start_row: 2

    columns:
      - name: "re_cat"
        col: "W"  # Column letter OR heading name

    aggregate: true
    aggregation_type: "count"  # or "sum"
    group_by: "re_cat"
    count_column: "re_cat_count"
```

---

## What's Working Now

### ✅ Implemented
- `{{month}}` - Extracts month name from "Ticket created - Date" column (e.g., "December")
- `{{prev_month}}` - Calculates previous month (e.g., "November")
- `{{re_req}}` - Counts rows in retail sheet
- `{{su_req}}` - Counts rows in supplier sheet
- `{{re_prev_req}}` - Counts rows in retail sheet FROM PREVIOUS MONTH FILE
- `{{su_prev_req}}` - Counts rows in supplier sheet FROM PREVIOUS MONTH FILE
- `{{re_sat}}`, `{{su_sat}}` - Count "good" satisfaction ratings (case-insensitive, exact match)
- `{{re_sat_c}}`, `{{su_sat_c}}` - Count "good with comment" ratings (case-insensitive)
- `{{re_prev_sat}}`, `{{su_prev_sat}}` - Count "good" FROM PREVIOUS MONTH FILE
- `{{re_prev_sat_c}}`, `{{su_prev_sat_c}}` - Count "good with comment" FROM PREVIOUS MONTH FILE
- `{{re_reso}}`, `{{su_reso}}` - Average resolution time (filters <= 3 days)
- `{{re_prev_reso}}`, `{{su_prev_reso}}` - Average resolution time FROM PREVIOUS MONTH FILE
- Dynamic tables (`re_sup_cat`, `su_sup_cat`) with aggregation

### 🚧 Not Yet Implemented
- Top organizations table

### 📦 WIP / Not Currently Used
- History management system (code exists in `src/history.py` but is commented out in `src/main.py`)
- Can be re-enabled in the future if needed

## Next Steps

### Critical Bugs to Fix

1. **Fix dynamic table population**
   - Debug `re_sup_cat` and `su_sup_cat` table generation
   - Ensure placeholders are being replaced correctly
   - Verify new rows are being created

### Immediate Tasks

1. **~~Implement remaining field types~~** ✅ DONE
   - ~~Date difference calculations~~ Implemented as `calculate_average_resolution_time()`

2. **Test with real data:**
   ```bash
   source .venv/bin/activate
   reporter
   # Select current month Excel file
   # Select previous month Excel file
   # Verify output
   ```

3. **Verify template:**
   - Ensure `report_template.pptx` has all placeholders
   - Test PowerPoint generation

### Future Enhancements

- Add error handling for missing columns
- Add validation for sheet names
- Support for more aggregation types
- Re-enable history management if needed (code exists in `src/history.py`)

---

## Legacy Documentation (Pre-Dec 29, 2025)

### Step 1: Verify Installation

```bash
cd /Users/c.morr/Repos/cs-reporter
pip install -e .
```

This creates the `reporter` command.

### Step 2: Prepare PowerPoint Template

1. Open one of your existing PowerPoint reports
2. Remove month-specific data
3. Add placeholders:
   - For text: `{{month}}`, `{{re_req}}`, etc.
   - For tables: Create a table with ONE row:
     - First cell: `{{table:re_sup_cat}} {{re_cat}}`
     - Second cell: `{{re_cat_count}}`
4. Save as `templates/report_template.pptx`

**Current file:** `templates/report_template.pptx.key` (needs to be converted)

### Step 3: Update Configuration

Open `config/mapping.yaml` and update:

1. **Sheet names**: Replace `"Tickets ADUS Tickets crea...1"` with actual names
2. **Column letters/headings**: Update to match your actual Excel structure
3. **Operations**: Verify each operation matches what you need

**Reference:** Use `config/mapping.yaml.example` as a guide

### Step 4: Test with Real Data

```bash
reporter
```

**Expected behavior:**
- Prompts for current month Excel file
- Prompts for previous month Excel file (first run only)
- Extracts data and shows progress
- Generates PowerPoint in `output/`
- Saves history to `output/history/YYYY-MM.json`

### Step 5: Verify Output

Check:
- [ ] All placeholders replaced with correct values
- [ ] Tables populated with correct row counts
- [ ] Previous month data appears correctly
- [ ] Month name is correct format

---

## Configuration Checklist

Before running, ensure your `config/mapping.yaml` has:

- [ ] Correct Excel sheet names
- [ ] Correct column letters for all operations
- [ ] Table `start_row` points to first data row (header is row before)
- [ ] Table `col` values match actual column headings
- [ ] All field names match PowerPoint placeholders
- [ ] `template_path` points to your actual template file

---

## PowerPoint Template Checklist

Your template needs:

### Simple Placeholders
- [ ] `{{month}}` - Current month name
- [ ] `{{prev_month}}` - Previous month name
- [ ] `{{re_req}}` - Retail requests count
- [ ] `{{re_prev_req}}` - Previous retail requests
- [ ] `{{re_reso}}` - Retail average resolution time
- [ ] `{{re_prev_reso}}` - Previous retail resolution time
- [ ] `{{re_sat}}` - Retail satisfaction count
- [ ] `{{re_sat_c}}` - Retail satisfaction with comment
- [ ] `{{su_req}}` - Supplier requests count
- [ ] `{{su_reso}}` - Supplier average resolution time
- [ ] `{{su_sat}}` - Supplier satisfaction count
- [ ] `{{su_sat_c}}` - Supplier satisfaction with comment

### Dynamic Tables
- [ ] **Retail Support Categories Table**
  - Row 1, Cell 1: `{{table:re_sup_cat}} {{re_cat}}`
  - Row 1, Cell 2: `{{re_cat_count}}`

- [ ] **Supplier Support Categories Table**
  - Row 1, Cell 1: `{{table:su_sup_cat}} {{su_cat}}`
  - Row 1, Cell 2: `{{su_cat_count}}`

- [ ] **Top Organizations Table**
  - Row 1, Cell 1: `{{table:top_orgs}} {{org_name}}`
  - Row 1, Cell 2: `{{org_count}}`

---

## Excel File Requirements

Your Excel files should have:

### Retail Sheet ("Tickets ADUS Tickets crea...1")
- Column B: "Ticket Created - Date" (dates)
- Column N: Resolution date column
- Column P: Satisfaction ratings ("good", "good with comment")
- Column W: Support Category (for table aggregation)

### Supplier Sheet ("Tickets ADUS Tickets crea...2")
- Column B: "Ticket Created - Date" (dates)
- Column N: Resolution date column
- Column P: Satisfaction ratings
- Column W: Support Category

### Organizations Sheet (if applicable)
- Column M: Organization names (for top orgs table)

**Note:** Sheet names and column letters need to match your actual Excel structure!

---

## Known Issues & Considerations

### Current Bugs (Jan 3, 2026)

1. **Dynamic tables not populating correctly**
   - `re_sup_cat` and `su_sup_cat` tables are implemented but:
     - Placeholder values not being filled in correctly
     - Not generating new rows as expected
   - Need to debug the table population logic in `ppt_writer.py`

### Fixed Bugs (Jan 3, 2026)

1. **~~Month field not working on title page~~** ✅ FIXED
   - Was caused by PowerPoint splitting `{{month}}` across multiple text runs
   - Fixed by the split placeholder handler in `ppt_writer.py`

2. **~~Previous month fields read current month data~~** ✅ FIXED
   - All `*_prev_*` fields now correctly read from the previous month Excel file
   - Changed architecture to always require both files instead of history management

3. **~~Some placeholders not being replaced~~** ✅ FIXED
   - PowerPoint splits placeholders across multiple "runs" (e.g., `"{{re_"` + `"sat}}"`)
   - Fixed by reconstructing full paragraph text before searching for placeholders

### Common Template Mistakes

1. **Typos in Placeholder Names**
   - ❌ Wrong: `{{re_prev_sat__c}}` (extra underscore)
   - ✅ Correct: `{{re_prev_sat_c}}`
   - The script now shows "Unreplaced fields" warnings to help catch these
   - Check the console output after running to see if any placeholders weren't replaced

2. **PowerPoint Split Placeholder Issue**
   - PowerPoint often splits placeholders across multiple "text runs" internally
   - This happens when you edit, copy/paste, or change formatting
   - Example: `{{re_sat}}` might be stored as `"{{re_"` + `"sat}}"`
   - **Fixed in Jan 3 update**: Script now handles split placeholders automatically
   - Old versions would fail to find split placeholders

3. **Case Sensitivity**
   - Placeholder names are case-sensitive
   - ❌ Wrong: `{{Re_Req}}` or `{{RE_REQ}}`
   - ✅ Correct: `{{re_req}}`
   - Must match exactly with config field names

4. **Using Debug Output**
   - After running, check console for:
     ```
     Replaced fields: month, re_req, re_sat, ...
     ⚠ Unreplaced fields: some_typo_field
     ```
   - Unreplaced fields indicate typos or missing config

### General Considerations

1. **Sheet Name Truncation**: Excel sometimes shows truncated sheet names like "Tickets ADUS Tickets crea...1"
   - Find the full name by right-clicking the sheet tab
   - Or open Excel and read the full name from the status bar

2. **Column Headers vs Letters**:
   - For tables, `col` can be either a letter ("W") or a heading name ("Support Category")
   - Column header search looks at the row immediately before `start_row`

3. **Date Calculations**:
   - `avg_date_diff` adds 1 to the difference (as requested)
   - Values > `max_days` are excluded from the average
   - Empty dates are skipped

4. **Dual-File Requirement**:
   - ALWAYS requires BOTH current and previous month Excel files
   - No history management - simpler and more predictable
   - History code still exists in `src/history.py` but is not currently used

5. **Template File Extension**:
   - Template has been converted to `.pptx` format
   - LibreOffice lock files (`.~lock.*#`) are now in `.gitignore`

---

## Testing Workflow

### First Month Run
```bash
reporter
# 1. Select November 2025 Excel (current)
# 2. Select October 2025 Excel (previous)
# 3. Check output/report_*.pptx
# 4. Verify output/history/2025-11.json was created
```

### Second Month Run
```bash
reporter
# 1. Select December 2025 Excel (current)
# 2. Script automatically loads November data from history
# 3. Check output/report_*.pptx
# 4. Verify previous month values match November's current values
```

---

## Troubleshooting Commands

```bash
# Verify installation
which reporter

# Run in debug mode
python -m src.main

# Check configuration
cat config/mapping.yaml

# View history
ls -la output/history/
cat output/history/2025-11.json

# Check dependencies
pip list | grep -E "pandas|python-pptx|PyYAML"
```

---

## Quick Reference: Field Types

| Field Pattern | Behavior | Example |
|---------------|----------|---------|
| `month` | Parses first date in column, returns month name | "December" |
| `prev_month` | Parses first date, returns previous month name | "November" |
| `*_req` | Counts rows in sheet | `re_req` → 150 |
| `*_sat_c` | Counts "good with comment" values in column | `re_sat_c` → 25 |
| `*_sat` | Counts "good" values in column (exact match) | `re_sat` → 42 |
| `*_reso` | Average resolution time (filters <= 3 days) | `re_reso` → 1.85 |
| Other fields | Reads first non-null value from column | Uses `cell` config |

## Utility Functions (excel_utils.py)

| Function | Purpose |
|----------|---------|
| `count_rows()` | Count data rows in sheet (excluding header) |
| `read_column_value()` | Read first non-null value from column |
| `parse_month_from_date()` | Extract month name from date |
| `parse_previous_month_from_date()` | Calculate and return previous month name |
| `count_column_value()` | Count occurrences of specific value in column (case-insensitive) |
| `calculate_average_resolution_time()` | Calculate average (solved - created) + 1, filtered to <= 3 days |
| `format_table_value()` | Format values (currency, percentage, number) |

---

## Files Modified

### Session Dec 29, 2025
**New Files:**
- `src/excel_utils.py` - Utility functions for Excel operations

**Modified Files:**
- `src/excel_reader.py` - Complete restructure, removed legacy operations
- `src/config.py` - Updated validation for new field structure
- `config/mapping.yaml` - Reorganized into category-based structure
- `templates/report_template.pptx` - Converted from Keynote format

**Environment Changes:**
- Installed `python-tk@3.13` via Homebrew
- Recreated virtual environment with tkinter support

### Session Jan 3, 2026
**Architectural Changes:**
- `src/main.py` - Changed to always prompt for both Excel files, disabled history management (kept as WIP)
- `src/excel_reader.py` - Added dual-file support, smart field routing based on `_prev_` in field name

**Feature Implementation:**
- `src/excel_utils.py` - Added `count_column_value()` for satisfaction ratings
- `src/excel_utils.py` - Added `calculate_average_resolution_time()` for resolution time metrics
- `src/excel_reader.py` - Added detection and handling of `_sat`, `_sat_c`, and `_reso` fields

**Bug Fixes:**
- `src/ppt_writer.py` - Fixed placeholder replacement to handle PowerPoint's split text runs
- `src/ppt_writer.py` - Added debug output showing which fields were/weren't replaced
- `src/excel_utils.py` - Added division by zero protection in `calculate_average_resolution_time()`

**Documentation:**
- `config/mapping.yaml` - Updated comments to reflect dual-file architecture and implemented fields
- `.gitignore` - Added LibreOffice lock files (`.~lock.*#`)
- `.claude/CLAUDE.md` - Updated for new architecture, features, and common template mistakes

---

## Session Progress

1. ✅ Fixed tkinter installation issue (Dec 29)
2. ✅ Converted template to `.pptx` format (Dec 29)
3. ✅ Restructured configuration to category-based approach (Dec 29)
4. ✅ Removed legacy operation-based system (Dec 29)
5. ✅ Created utility module for Excel functions (Dec 29)
6. ✅ Implemented month and prev_month parsing (Dec 29)
7. ✅ Implemented row counting for request fields (Dec 29)
8. ✅ Implemented value counting for satisfaction fields (Jan 3)
9. ✅ Changed architecture to dual-file (always require both Excel files) (Jan 3)
10. ✅ Previous month fields now correctly read from previous Excel file (Jan 3)
11. ✅ Fixed PowerPoint placeholder replacement to handle split text runs (Jan 3)
12. ✅ Added debug output for replaced/unreplaced fields (Jan 3)
13. ✅ Implemented resolution time calculation (filters <= 3 days) (Jan 3)
14. ✅ Month field on title page now working (fixed by split placeholder handler) (Jan 3)
15. ⬜ Fix dynamic table population
16. ⬜ Test full workflow with real data

---

## Questions to Answer When You Return

Before testing, you'll need to confirm:

1. **Exact sheet names** in your Excel files (not truncated)
2. **Actual column letters or headings** for:
   - Ticket Created Date
   - Resolution Date
   - Satisfaction column
   - Support Category column
   - Organization column
3. **Exact text values** for satisfaction (is it "good" or "Good" or "Satisfied"?)
4. **Which sheet has organization data** (RE or SU or separate?)

---

## Support

- **Documentation**: `README.md` has full usage guide
- **Configuration examples**: `config/mapping.yaml.example`
- **Your config**: `config/mapping.yaml` (needs customization)

Good luck! The foundation is solid, you just need to customize the config and template for your specific data structure.
