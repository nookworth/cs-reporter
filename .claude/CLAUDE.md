# CS Reporter - Project Status

**Last Updated:** December 29, 2025
**Status:** Restructured implementation, simplified field extraction approach

---

## Project Overview

CS Reporter is a Python utility that:
- Reads data from Excel files using custom operations
- Generates PowerPoint presentations using a template
- Automatically tracks month-over-month data with history management
- Supports dynamic tables with aggregation

**Primary Use Case:** Monthly support ticket reporting with automatic historical comparison

---

## Recent Changes (Dec 29, 2025)

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
- `{{re_prev_req}}` - Counts rows in retail sheet (same as re_req for now)
- `{{su_prev_req}}` - Counts rows in supplier sheet (same as su_req for now)
- Dynamic tables (`re_sup_cat`, `su_sup_cat`) with aggregation

### 🚧 Not Yet Implemented
- `{{re_reso}}`, `{{su_reso}}` - Average date difference calculations
- `{{re_sat}}`, `{{su_sat}}` - Count "good" satisfaction ratings
- `{{re_sat_c}}`, `{{su_sat_c}}` - Count "good with comment" ratings
- `{{re_prev_reso}}`, `{{su_prev_reso}}` - Previous month resolution times
- `{{re_prev_sat}}`, `{{su_prev_sat}}` - Previous month satisfaction counts
- Previous month history tracking (currently all prev_ fields read from current month)
- Top organizations table

## Next Steps

### Immediate Tasks

1. **Implement remaining field types:**
   - Date difference calculations (for `re_reso`, `su_reso`)
   - Value counting (for satisfaction fields)
   - These will need new utility functions in `excel_utils.py`

2. **Test with real data:**
   ```bash
   source .venv/bin/activate
   reporter
   ```

3. **Verify template:**
   - Ensure `report_template.pptx` has all placeholders
   - Test PowerPoint generation

### Future Enhancements

- Re-implement history management for true previous month data
- Add error handling for missing columns
- Add validation for sheet names
- Support for more aggregation types

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

4. **First Month Setup**:
   - First run requires BOTH current and previous month Excel files
   - After that, only current month is needed
   - History files are human-readable JSON in `output/history/`

5. **Template File Extension**:
   - Current template is `.pptx.key` (Keynote file?)
   - Needs to be `.pptx` (PowerPoint format)
   - Save as PowerPoint from Keynote if needed

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
| Other fields | Reads first non-null value from column | Uses `cell` config |

## Utility Functions (excel_utils.py)

| Function | Purpose |
|----------|---------|
| `count_rows()` | Count data rows in sheet (excluding header) |
| `read_column_value()` | Read first non-null value from column |
| `parse_month_from_date()` | Extract month name from date |
| `parse_previous_month_from_date()` | Calculate and return previous month name |
| `format_table_value()` | Format values (currency, percentage, number) |

---

## Files Modified This Session (Dec 29, 2025)

### New Files
- `src/excel_utils.py` - Utility functions for Excel operations

### Modified Files
- `src/excel_reader.py` - Complete restructure, removed legacy operations
- `src/config.py` - Updated validation for new field structure
- `config/mapping.yaml` - Reorganized into category-based structure
- `templates/report_template.pptx` - Converted from Keynote format
- `PROJECT_STATUS.md` - This update

### Environment Changes
- Installed `python-tk@3.13` via Homebrew
- Recreated virtual environment with tkinter support

---

## Session Progress

1. ✅ Fixed tkinter installation issue
2. ✅ Converted template to `.pptx` format
3. ✅ Restructured configuration to category-based approach
4. ✅ Removed legacy operation-based system
5. ✅ Created utility module for Excel functions
6. ✅ Implemented month and prev_month parsing
7. ✅ Implemented row counting for request fields
8. ⬜ Implement date difference calculations
9. ⬜ Implement value counting for satisfaction
10. ⬜ Test full workflow with real data

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
