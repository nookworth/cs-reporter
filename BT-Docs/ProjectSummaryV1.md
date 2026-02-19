# CS Reporter - Code Architecture Guide

---

**Title:** Code Architecture \u0026 Implementation Guide  
**Date:** 2026-02-18  
**Contact:** bentran.phoenix@gmail.com

---


## 📊 Data Flow Diagram

```
┌─────────────────────────────────────────────────────────────────┐
│                         USER ACTION                              │
│              Run: reporter --config demo_mapping.yaml            │
└───────────────────────────────┬─────────────────────────────────┘
                                │
                                ▼
┌─────────────────────────────────────────────────────────────────┐
│                    1. MAIN.PY (Entry Point)                      │
│  ───────────────────────────────────────────────────────────     │
│  • Parse command-line arguments                                  │
│  • Load configuration file (config.py)                           │
│  • Open file dialogs to select Excel files                       │
└───────────────────────────────┬─────────────────────────────────┘
                                │
                                ▼
┌─────────────────────────────────────────────────────────────────┐
│              2. EXCEL_READER.PY (Data Extraction)                │
│  ───────────────────────────────────────────────────────────     │
│  • Opens both Excel files (current + previous month)             │
│  • Reads field groups:                                           │
│    - Standard fields (month, prev_month)                         │
│    - Retail fields (re_req, re_sat, re_reso, etc.)              │
│    - Supplier fields (su_req, su_sat, su_reso, etc.)            │
│  • Extracts table data (support categories, top orgs)            │
│  • Uses excel_utils.py for operations:                           │
│    - count_rows()           → _req fields                        │
│    - count_column_value()   → _sat fields                        │
│    - calculate_avg_reso()   → _reso fields                       │
│    - count_unique_values()  → tables                             │
└───────────────────────────────┬─────────────────────────────────┘
                                │
                                ▼
┌─────────────────────────────────────────────────────────────────┐
│                    3. DATA DICTIONARY                            │
│  ───────────────────────────────────────────────────────────     │
│  {                                                                │
│    "month": "January",                                           │
│    "prev_month": "December",                                     │
│    "re_req": 92,                                                 │
│    "re_prev_req": 78,                                            │
│    "re_sat": 36,                                                 │
│    "re_sup_cat": [                                               │
│      {"re_cat": "Billing", "re_cat_count": 23},                 │
│      {"re_cat": "Technical", "re_cat_count": 18},               │
│      ...                                                         │
│    ],                                                            │
│    ...                                                           │
│  }                                                               │
└───────────────────────────────┬─────────────────────────────────┘
                                │
                                ▼
┌─────────────────────────────────────────────────────────────────┐
│            4. PPT_WRITER.PY (Report Generation)                  │
│  ───────────────────────────────────────────────────────────     │
│  • Loads PowerPoint template                                     │
│  • Replaces scalar placeholders:                                 │
│    {{month}} → "January"                                         │
│    {{re_req}} → "92"                                             │
│  • Populates dynamic tables:                                     │
│    {{table:re_sup_cat}} → Creates rows from table data          │
│  • Saves to output/report_YYYYMMDD_HHMMSS.pptx                  │
└───────────────────────────────┬─────────────────────────────────┘
                                │
                                ▼
┌─────────────────────────────────────────────────────────────────┐
│                     5. GENERATED REPORT                          │
│              output/demo_report_20260217_003919.pptx             │
└─────────────────────────────────────────────────────────────────┘
```

---

## 🔧 Key Components

### **1. main.py** - Entry Point
**Purpose:** Orchestrates the entire workflow

**Key functions:**
- `parse_args()` - Handles `--config` flag
- `select_excel_file()` - Opens file dialog
- `main()` - Main execution flow:
  1. Load config
  2. Select files (current + previous month)
  3. Extract data using ExcelReader
  4. Generate report using PowerPointWriter

---

### **2. excel_reader.py** - Data Extraction Engine
**Purpose:** Extracts data from Excel files based on configuration

**Main class:** `ExcelReader`

**Key methods:**
- `extract_data()` - Orchestrates all extraction:
  ```python
  def extract_data(self):
      data = {}
      data.update(self._extract_field_group(self.standard_excel_fields, "standard"))
      data.update(self._extract_field_group(self.retail_excel_fields, "retail"))
      data.update(self._extract_field_group(self.supplier_excel_fields, "supplier"))
      # Extract tables...
      return data
  ```

- `_extract_field_group()` - Processes each field:
  ```python
  # Field name suffix determines operation:
  if field_name.endswith("_req"):       # Count rows
      value = excel_utils.count_rows(excel_path, sheet)

  elif field_name.endswith("_sat"):     # Count "good" ratings
      value = excel_utils.count_column_value(excel_path, sheet, cell, "good")

  elif field_name.endswith("_sat_c"):   # Count "good with comment"
      value = excel_utils.count_column_value(excel_path, sheet, cell, "good with comment")

  elif field_name.endswith("_reso"):    # Average resolution time
      value = excel_utils.calculate_average_resolution_time(excel_path, sheet)
  ```

**Smart file selection:**
```python
# Fields with "_prev_" use previous month file
is_prev_field = "_prev_" in field_name
excel_path = self.previous_excel_path if is_prev_field else self.excel_path
```

---

### **3. excel_utils.py** - Excel Operations Library
**Purpose:** Low-level Excel operations using pandas

**Key functions:**
- `count_rows(excel_path, sheet)` - Counts non-header rows
- `count_column_value(excel_path, sheet, column, value)` - Counts matching cells
- `calculate_average_resolution_time(excel_path, sheet)` - Computes avg days
- `count_unique_values(excel_path, sheet, column)` - Groups and counts for tables
- `read_column_value(excel_path, sheet, column)` - Reads first value
- `parse_month_from_date(value)` - Formats dates

---

### **4. ppt_writer.py** - PowerPoint Generator
**Purpose:** Creates reports by replacing placeholders in template

**Main class:** `PowerPointWriter`

**Key methods:**
- `generate_report(data)` - Main generation:
  ```python
  def generate_report(self, data):
      prs = Presentation(self.template_path)
      self._replace_placeholders(prs, data)
      output_path = f"output/report_{timestamp}.pptx"
      prs.save(output_path)
      return output_path
  ```

- `_replace_placeholders()` - Finds and replaces:
  - Scalar placeholders: `{{field_name}}`
  - Table placeholders: `{{table:table_name}}`

- `_populate_table()` - Dynamic table filling:
  ```python
  # Table structure:
  # Row 0: Header (preserved)
  # Row 1: Template row with placeholders ({{re_cat}}, {{re_cat_count}})
  # Rows 2+: Cloned from template, filled with data
  ```

- `_replace_in_text_frame()` - Handles PowerPoint text quirks
  (PowerPoint often splits `{{placeholder}}` across multiple text runs)

---

### **5. config.py** - Configuration Loader
**Purpose:** Loads and validates YAML configuration

**Key function:**
- `load_config(config_path)` - Loads YAML, adds metadata

---

## 🎯 Field Naming Conventions (Current System)

The **field name suffix** determines the operation:

| Suffix | Operation | Example | Result |
|--------|-----------|---------|--------|
| `_req` | Count rows | `re_req` | 92 tickets |
| `_reso` | Avg resolution time (days) | `re_reso` | 2.18 days |
| `_sat` | Count "good" ratings | `re_sat` | 36 satisfied |
| `_sat_c` | Count "good with comment" | `re_sat_c` | 26 with comment |

**Prefix conventions:**
- `re_` = Retail
- `su_` = Supplier
- `_prev_` = Previous month data

**Examples:**
- `re_req` = Retail requests (current month)
- `re_prev_req` = Retail requests (previous month)
- `su_sat` = Supplier satisfaction (current month)
- `su_prev_reso` = Supplier resolution time (previous month)

---

## 📝 Template Placeholder Syntax

In the PowerPoint template:

**Scalar fields:**
```
{{month}}           → "January"
{{re_req}}          → "92"
{{re_sat}}          → "36"
{{su_prev_reso}}    → "1.89"
```

**Dynamic tables:**
```
Header row:  Category        |  Count
Template:    {{re_cat}}      |  {{re_cat_count}}
```

The code will:
1. Keep the header row
2. Use template row to create data rows
3. Delete the template row
4. Fill each data row with values from the table data

---

## 🔄 Execution Flow Example

```bash
reporter --config config/demo_mapping.yaml
```

**Step-by-step:**

1. **main.py** loads `demo_mapping.yaml`
2. **main.py** prompts for Excel files:
   - Current: `demo_january_2026.xlsx`
   - Previous: `demo_december_2025.xlsx`
3. **ExcelReader** extracts data:
   - `re_req` → Count rows in "Retail Tickets" (current) → 92
   - `re_prev_req` → Count rows in "Retail Tickets" (previous) → 78
   - `re_sat` → Count "good" in satisfaction column → 36
   - `re_sup_cat` → Group by "Support Category" → [{"re_cat": "Billing", "re_cat_count": 23}, ...]
4. **PowerPointWriter** generates report:
   - Opens `templates/demo_template.pptx`
   - Replaces `{{re_req}}` with `92`
   - Populates `{{table:re_sup_cat}}` with category rows
   - Saves to `output/demo_report_20260217_003919.pptx`
5. **Done!** Report ready to open

---

## 🚀 Future: Config-Driven Operations (Version 2)

**Current problem:** Operations are hardcoded based on field names

**Future solution:** Define operations in config:

```yaml
re_sat:
  column: "Ticket satisfaction rating"
  operation: count_matching        # ← Explicit operation
  match_value: "good"
  case_sensitive: false

re_reso:
  operation: average_date_difference
  start_column: "Ticket created - Date"
  end_column: "Ticket resolved - Date"
  filter:
    max_days: 3
```
This makes the tool **truly general-purpose** - no code changes needed for different metrics!
See `plans/generalize-to-config-driven-operations.md` for the full roadmap.

---

## 📚 File Reference

```
src/
├── main.py              → Entry point, CLI, file selection
├── excel_reader.py      → Data extraction from Excel
├── excel_utils.py       → Low-level Excel operations (pandas)
├── ppt_writer.py        → PowerPoint generation (python-pptx)
├── config.py            → YAML config loader
└── history.py           → (WIP) History tracking (disabled)

config/
├── mapping.yaml         → Production config (requires proprietary template)
└── demo_mapping.yaml    → Demo config (for demo data)

templates/
└── demo_template.pptx   → Demo PowerPoint template

scripts/
└── generate_demo_data.py → Creates demo Excel files
```
---
## 🎓 Learning Path
**To understand the codebase:**
1. **Start with:** `src/main.py` - See the high-level flow
2. **Then read:** `config/demo_mapping.yaml` - Understand the configuration
3. **Explore:** `src/excel_reader.py` - See how data is extracted
4. **Check:** `src/excel_utils.py` - See the actual Excel operations
5. **Finally:** `src/ppt_writer.py` - See how PowerPoint is generated
**To modify the code:**
1. **Add a new field?** → Update `config/demo_mapping.yaml`
2. **Add a new operation?** → Modify `excel_reader.py` (add new suffix handling)
3. **Change template layout?** → Edit `templates/demo_template.pptx`
4. **Add new table?** → Update config `table_fields` section
---
## 💡 Tips
**Debugging:**
- Watch the console output - it shows each field being processed
- Check `output/` for generated reports
- Use `--config` flag to switch between configs
**Testing:**
- Always test with demo data first: `reporter --config config/demo_mapping.yaml`
- Demo data is reproducible: `python scripts/generate_demo_data.py`
**Contributing:**
- See `BensPlan.md` for the full development guide
- See `plans/generalize-to-config-driven-operations.md` for the roadmap