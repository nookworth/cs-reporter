# CS Reporter - Project Status

**Last Updated:** December 3, 2025
**Status:** Core implementation complete, ready for testing with real data

---

## Project Overview

CS Reporter is a Python utility that:
- Reads data from Excel files using custom operations
- Generates PowerPoint presentations using a template
- Automatically tracks month-over-month data with history management
- Supports dynamic tables with aggregation

**Primary Use Case:** Monthly support ticket reporting with automatic historical comparison

---

## What's Been Implemented

### ✅ Core Features Complete

1. **Custom Excel Operations**
   - `parse_month`: Extract month from date columns
   - `count_rows`: Count total rows in a sheet
   - `avg_date_diff`: Calculate average date differences (with max threshold)
   - `count_value`: Count occurrences of specific text values

2. **Dynamic Tables with Aggregation**
   - Auto-populate PowerPoint tables with variable rows
   - Count aggregation: Group and count occurrences
   - Sum aggregation: Group and sum numeric values
   - Column header search (no hardcoded column letters)

3. **History Management**
   - Automatic saving of extracted data to `output/history/YYYY-MM.json`
   - Automatic loading of previous month data
   - Graceful fallback to manual file selection on first run

4. **Template-Based PowerPoint Generation**
   - Placeholder replacement: `{{field_name}}` → actual values
   - Dynamic table population: `{{table:table_name}}` → multiple rows
   - Formatting support: currency, percentages, numbers

5. **CLI with File Dialogs**
   - Simple `reporter` command
   - Automatic file dialogs for Excel selection
   - Clear progress indicators

---

## Project Structure

```
cs-reporter/
├── src/
│   ├── __init__.py
│   ├── main.py              # CLI entry point with dual file handling
│   ├── excel_reader.py      # Excel operations and table reading
│   ├── ppt_writer.py        # PowerPoint generation
│   ├── config.py            # Configuration loader
│   └── history.py           # History management (NEW)
├── templates/
│   └── report_template.pptx.key  # YOUR TEMPLATE (needs setup)
├── config/
│   ├── mapping.yaml         # YOUR CONFIG (needs customization)
│   └── mapping.yaml.example # Reference example
├── output/
│   ├── .gitkeep
│   └── history/             # Auto-saved monthly data (NEW)
│       └── .gitkeep
├── requirements.txt
├── setup.py
├── README.md
└── PROJECT_STATUS.md        # This file
```

---

## Configuration Format

Your `config/mapping.yaml` should have three main sections:

### 1. Current Month Fields

Operations to extract from the current month Excel file:

```yaml
current_month_fields:
  month:
    sheet: "Tickets ADUS Tickets crea...1"
    operation: "parse_month"
    column: "B"

  re_req:
    sheet: "Tickets ADUS Tickets crea...1"
    operation: "count_rows"

  re_reso:
    sheet: "Tickets ADUS Tickets crea...1"
    operation: "avg_date_diff"
    date_col_start: "B"
    date_col_end: "N"
    max_days: 3

  re_sat:
    sheet: "Tickets ADUS Tickets crea...1"
    operation: "count_value"
    column: "P"
    value: "good"
```

### 2. Previous Month Fields

Maps to historical data (no Excel operations needed):

```yaml
previous_month_fields:
  prev_month:
    source: "month"

  re_prev_req:
    source: "re_req"

  re_prev_reso:
    source: "re_reso"
```

### 3. Dynamic Tables

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

## Next Steps (When You Return)

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

## Quick Reference: Operations

| Operation | Purpose | Required Config |
|-----------|---------|----------------|
| `parse_month` | Extract month from date column | `sheet`, `column` |
| `count_rows` | Count total rows | `sheet` |
| `avg_date_diff` | Average date difference | `sheet`, `date_col_start`, `date_col_end`, `max_days` |
| `count_value` | Count text occurrences | `sheet`, `column`, `value` |

---

## Files Modified/Created Since Start

### New Files
- `src/history.py` - History management
- `config/mapping.yaml.example` - Configuration reference
- `output/history/.gitkeep` - History directory marker
- `PROJECT_STATUS.md` - This file

### Modified Files
- `src/excel_reader.py` - Added operations, previous month support
- `src/main.py` - Dual file handling, history integration
- `src/config.py` - Support for new config structure
- `config/mapping.yaml` - Updated by you (needs more updates)
- `.gitignore` - Added history files

### Files Provided by You
- `templates/report_template.pptx.key` - Template (needs conversion)

---

## When You're Ready to Proceed

1. ✅ Read this document
2. ⬜ Convert template to `.pptx` format
3. ⬜ Open one of your Excel files
4. ⬜ Document the actual sheet names and column letters
5. ⬜ Update `config/mapping.yaml` with real values
6. ⬜ Add placeholders to PowerPoint template
7. ⬜ Run `pip install -e .`
8. ⬜ Test with: `reporter`
9. ⬜ Check generated report in `output/`
10. ⬜ Verify history saved to `output/history/`

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
