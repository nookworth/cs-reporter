# CS Reporter

**Automated monthly support ticket report generator**

Transforms Excel ticket exports into formatted PowerPoint presentations with month-over-month comparisons.

---

## Quick Start

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

### Generate a Report

**Windows:**
```cmd
run-reporter.bat
```

**macOS/Linux:**
```bash
./run-reporter.sh
```

Then select your current and previous month Excel files when prompted.

Your report will be saved in the `output/` folder.

---

## What It Does

- ✅ Extracts ticket metrics (counts, satisfaction ratings, resolution times)
- ✅ Generates support category breakdowns
- ✅ Creates top organizations table
- ✅ Compares current vs. previous month automatically
- ✅ Outputs formatted PowerPoint presentation

---

## Requirements

- **Python 3.8+** ([Download](https://www.python.org/downloads/))
  - Windows users: Check "Add Python to PATH" during installation!
- **Two Excel files** (current month + previous month)
- **PowerPoint template** (`templates/report_template.pptx`)
- **Works on:** Windows, macOS, or Linux

---

## Files Overview

```
cs-reporter/
├── setup.sh              # One-time setup script
├── run-reporter.sh       # Daily launcher (created by setup)
├── USER_GUIDE.md         # Complete documentation
├── config/
│   └── mapping.yaml      # Field mappings (customize for your Excel structure)
├── templates/
│   └── report_template.pptx  # PowerPoint template (customize design)
├── output/               # Generated reports go here
└── src/                  # Source code
```

---

## Customization

### Excel Structure
Edit `config/mapping.yaml` to match your Excel file's sheet names and column headers.

### Report Design
Edit `templates/report_template.pptx` to customize the look and feel of your reports.

See **USER_GUIDE.md** for detailed customization instructions.

---

## Troubleshooting

### "Command not found"
Make sure you're using the launcher:

**Windows:** `run-reporter.bat`

**macOS/Linux:** `./run-reporter.sh`

### "Sheet not found" or "Column not found"
Update `config/mapping.yaml` with your actual Excel sheet names and column headers.

### Placeholders Not Replaced
Check the console output for "⚠ Unreplaced fields" and fix typos in your PowerPoint template.

### More Help
See **USER_GUIDE.md** for comprehensive troubleshooting.

---

## What Gets Generated

### Metrics (Current vs. Previous)
- Total tickets
- Satisfaction ratings (good, good with comment)
- Average resolution time

### Tables
- Retail support categories with counts
- Supplier support categories with counts
- Top 5 organizations by ticket volume

---

## Support

📖 **Full Documentation:** See `USER_GUIDE.md`
⚙️ **Configuration:** Edit `config/mapping.yaml`
🎨 **Template Design:** Edit `templates/report_template.pptx`

---

## Example Workflow

**Windows:**
```cmd
# First time
setup.bat

# Generate December report
run-reporter.bat
# Select: December_Tickets.xlsx (current)
# Select: November_Tickets.xlsx (previous)
# Output: output/report_20260106_143022.pptx
```

**macOS/Linux:**
```bash
# First time
./setup.sh

# Generate December report
./run-reporter.sh
# Select: December_Tickets.xlsx (current)
# Select: November_Tickets.xlsx (previous)
# Output: output/report_20260106_143022.pptx
```

---

## Technical Details

- **Language:** Python 3
- **Key Libraries:** pandas, python-pptx, PyYAML
- **Architecture:** Dual-file (always requires current + previous month)
- **Template Engine:** Placeholder-based (`{{field_name}}`)

---

**Ready to generate your first report?** Run `./setup.sh` to get started!
