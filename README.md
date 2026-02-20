# CS Reporter

**Automated monthly support ticket report generator**

Transforms Excel ticket exports into formatted PowerPoint presentations with month-over-month comparisons.

> **Note:** This project has two versions:
> - **V1 (Original)** - Created by [Your Friend's Name] - Suffix-based configuration
> - **V2 (Refactored)** - Enhanced by [Your Name] - Operation-based configuration with advanced features

---

## 🚀 Quick Start

### Choose Your Version

**V1 (Original - Recommended for existing users):**
```bash
python -m src.main
```

**V2 (Refactored - New features):**
```bash
python run_v2.py
```

### First Time Setup

**Windows:**
```cmd
cd C:\path\to\cs-reporter
setup.bat
```

**macOS/Linux:**
```bash
cd /path/to/cs-reporter
chmod +x setup.sh
./setup.sh
```

---

## 📊 What It Does

- ✅ Extracts ticket metrics (counts, satisfaction ratings, resolution times)
- ✅ Generates support category breakdowns
- ✅ Creates top organizations table
- ✅ Compares current vs. previous month automatically
- ✅ Outputs formatted PowerPoint presentation

---

## 📋 Requirements

- **Python 3.8+** ([Download](https://www.python.org/downloads/))
  - Windows users: Check "Add Python to PATH" during installation!
- **Two Excel files** (current month + previous month)
- **PowerPoint template** (`templates/report_template.pptx` or `templates/demo_template.pptx`)
- **Works on:** Windows, macOS, or Linux

---

## 📁 Files Overview

```
cs-reporter/
├── setup.sh              # One-time setup script
├── run_v2.py             # V2 launcher
├── run_v2.bat            # V2 launcher (Windows)
├── USER_GUIDE.md         # Complete documentation
├── config/
│   ├── mapping.yaml      # V1 config (suffix-based)
│   ├── mapping_v2.yaml   # V2 config (operation-based)
│   ├── demo_mapping_v2.yaml  # V2 demo config
│   └── schema_v2.yaml    # V2 operation documentation
├── templates/
│   ├── report_template.pptx  # Production template
│   └── demo_template.pptx    # Demo template
├── output/               # Generated reports go here
├── src/                  # Source code
│   ├── main.py          # V1 entry point
│   ├── main_v2.py       # V2 entry point
│   ├── config.py        # V1 config loader
│   ├── config_v2.py     # V2 config loader (with validation)
│   └── ...              # Other modules
└── BT-Docs/             # Technical documentation
```

---

## 🆕 Version 2 (V2) - New Features

**V2 is a complete refactoring with powerful new capabilities:**

### Key Improvements

#### 1. **Operation-Based Configuration**
- Explicit operation specification (no more suffix guessing)
- Clear, self-documenting config files
- Easy to understand and modify

**V1 (Suffix-based):**
```yaml
retail_excel_fields:
  re_req:
    cell: null  # Suffix _req means "count rows"
```

**V2 (Operation-based):**
```yaml
fields:
  re_req:
    operation: count_rows
    sheet: "Tickets"
    source: current
```

#### 2. **Configurable Filters**
- 12 filter operators (equals, contains, greater_than, etc.)
- Remove hardcoded business logic
- Flexible data filtering

```yaml
re_reso:
  operation: avg_date_diff
  sheet: "Tickets"
  start_column: "Ticket created - Date"
  end_column: "Ticket solved - Date"
  filters:
    - column: "Assignee"
      operator: "not_equals"
      value: "Leo Brown"
    - column: "Resolution Days"
      operator: "less_than_or_equal"
      value: 3
```

#### 3. **Config Validation**
- Validates configs on load
- Clear error messages
- Prevents runtime errors

#### 4. **7 Built-in Operations**
- `count_rows` - Count total rows
- `count_value` - Count specific values
- `sum` - Sum numeric columns
- `avg_date_diff` - Calculate average time differences
- `parse_month` - Extract month names
- `parse_previous_month` - Calculate previous month
- `read_value` - Read single values

#### 5. **Complete Documentation**
- `config/schema_v2.yaml` - Full operation reference
- `BT-Docs/` - Technical documentation
- Clear examples and guides

### How to Use V2

**Quick Test (Demo Data):**
```bash
python test_v2_config.py
```

**Production Use:**
```bash
python run_v2.py
# or
run_v2.bat  # Windows
```

**With Custom Config:**
```bash
python -m src.main_v2 --config config/mapping_v2.yaml
```

### V2 Documentation

- 📖 **Schema Reference:** `config/schema_v2.yaml`
- 🚀 **How to Run:** `BT-Docs/HOW_TO_RUN_V1_V2.md`

### Migration from V1 to V2

Both versions work independently. To migrate:

1. **Test V2 with demo data:**
   ```bash
   python test_v2_config.py
   ```

2. **Convert your config:**
   - Use `config/mapping_v2.yaml` as reference
   - See `config/schema_v2.yaml` for operation details

3. **Run side-by-side:**
   - Keep using V1: `python -m src.main`
   - Test V2: `python run_v2.py`

4. **Switch when ready:**
   - V2 is production-ready
   - Fully tested and validated

---

## 📝 Customization

### Excel Structure
**V1:** Edit `config/mapping.yaml` to match your Excel file's sheet names and column headers.

**V2:** Edit `config/mapping_v2.yaml` using operation-based format. See `config/schema_v2.yaml` for all available operations.

### Report Design
Edit `templates/report_template.pptx` to customize the look and feel of your reports.

See **USER_GUIDE.md** for detailed customization instructions.

---

## 🔧 Troubleshooting

### "Command not found"
Make sure you're using the correct launcher:

**V1:** `python -m src.main`

**V2:** `python run_v2.py` or `run_v2.bat`

### "Sheet not found" or "Column not found"
Update your config file with actual Excel sheet names and column headers.

### Placeholders Not Replaced
Check the console output for unreplaced fields and fix typos in your PowerPoint template.

### More Help
See **USER_GUIDE.md** for comprehensive troubleshooting.

---

## 📈 What Gets Generated

### Output File
- **Location:** `output/` folder
- **Format:** PowerPoint (.pptx)
- **Naming:** `report_YYYYMMDD_HHMMSS.pptx` (V1) or `demo_report_YYYYMMDD_HHMMSS.pptx` (V2)
- **Example:** `output/report_20260220_162606.pptx`

### Metrics (Current vs. Previous)
- Total tickets
- Satisfaction ratings (good, good with comment)
- Average resolution time

### Tables
- Retail support categories with counts
- Supplier support categories with counts
- Top 5 organizations by ticket volume

### Sample Output
```
CS Reporter - Excel to PowerPoint Report Generator
============================================================
[OK] Configuration loaded from: config/demo_mapping_v2.yaml
[OK] Selected: current_month.xlsx
[OK] Selected: previous_month.xlsx
[OK] Extracted 21 fields from Excel files
  Populated table 're_sup_cat' with 3 rows (6 replacements)
  Populated table 'su_sup_cat' with 3 rows (6 replacements)
  Populated table 'top_orgs' with 5 rows (10 replacements)
  Made 52 replacements
[OK] Report generated: output/demo_report_20260220_162606.pptx

Done!
```

---

## 💡 Support & Documentation

📖 **Full Documentation:** See `USER_GUIDE.md`
⚙️ **V1 Configuration:** Edit `config/mapping.yaml`
⚙️ **V2 Configuration:** Edit `config/mapping_v2.yaml`
📚 **V2 Schema:** See `config/schema_v2.yaml`
🎨 **Template Design:** Edit `templates/report_template.pptx`
📁 **Technical Docs:** See `BT-Docs/` folder

---

## 🎯 Example Workflow

### V1 (Original):

**Windows:**
```cmd
# First time
setup.bat

# Generate December report
python -m src.main
# Select: December_Tickets.xlsx (current)
# Select: November_Tickets.xlsx (previous)
# Output: output/report_20260106_143022.pptx
```

**macOS/Linux:**
```bash
# First time
./setup.sh

# Generate December report
python -m src.main
# Select: December_Tickets.xlsx (current)
# Select: November_Tickets.xlsx (previous)
# Output: output/report_20260106_143022.pptx
```

### V2 (Refactored):

**Windows:**
```cmd
# First time
setup.bat

# Generate December report with V2
run_v2.bat
# Select: December_Tickets.xlsx (current)
# Select: November_Tickets.xlsx (previous)
# Output: output/demo_report_20260220_162606.pptx
```

**macOS/Linux:**
```bash
# First time
./setup.sh

# Generate December report with V2
python run_v2.py
# Select: December_Tickets.xlsx (current)
# Select: November_Tickets.xlsx (previous)
# Output: output/demo_report_20260220_162606.pptx
```

---

## 🛠️ Technical Details

### V1 (Original)
- **Language:** Python 3
- **Key Libraries:** pandas, python-pptx, PyYAML
- **Architecture:** Suffix-based operation detection
- **Config:** `config/mapping.yaml`

### V2 (Refactored)
- **Language:** Python 3
- **Key Libraries:** pandas, python-pptx, PyYAML
- **Architecture:** Operation-based with validation
- **Config:** `config/mapping_v2.yaml`
- **Features:** Configurable filters, 12 operators, 7 operations
- **Validation:** Config validation on load

---

## 👥 Contributors

- **V1 (Original):** [Your Friend's Name] - Initial implementation
- **V2 (Refactored):** [Your Name] - Operation-based architecture, filters, validation

---

**Ready to generate your first report?**
- **V1 Users:** Run `python -m src.main`
- **V2 Users:** Run `python run_v2.py`
- **New Users:** Try V2 demo with `python test_v2_config.py`
