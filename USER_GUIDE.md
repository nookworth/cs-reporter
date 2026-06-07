# CS Reporter - User Guide

**Simple Monthly Support Ticket Report Generator**

This tool automatically generates PowerPoint reports from your monthly Excel support ticket exports.

---

## First-Time Setup

### Prerequisites
- Python 3.8 or newer ([Download here](https://www.python.org/downloads/))
  - **Windows users:** Check "Add Python to PATH" during installation!
- Works on: Windows, macOS, or Linux

### Installation Steps

**Choose your operating system:**

<details>
<summary><b>Windows</b></summary>

1. **Download the cs-reporter folder** to your computer

2. **Open Command Prompt** and navigate to the folder:
   ```cmd
   cd C:\path\to\cs-reporter
   ```

3. **Run the setup script**:
   ```cmd
   setup.bat
   ```

   This will:
   - Create a virtual environment
   - Install all required dependencies
   - Set up the `reporter` command

4. **Done!** You're ready to generate reports.

</details>

<details>
<summary><b>macOS / Linux</b></summary>

1. **Download the cs-reporter folder** to your computer

2. **Open Terminal** and navigate to the folder:
   ```bash
   cd /path/to/cs-reporter
   ```

3. **Run the setup script**:
   ```bash
   chmod +x setup.sh
   ./setup.sh
   ```

   This will:
   - Create a virtual environment
   - Install all required dependencies
   - Set up the `reporter` command

4. **Done!** You're ready to generate reports.

</details>

---

## Generating a Report

### Quick Start

**Choose your operating system:**

<details>
<summary><b>Windows</b></summary>

1. **Open Command Prompt** and go to the cs-reporter folder:
   ```cmd
   cd C:\path\to\cs-reporter
   ```

2. **Run the reporter**:
   ```cmd
   run-reporter.bat
   ```

3. **Select your Excel files** when prompted:
   - First dialog: Select **current month** Excel file
   - Second dialog: Select **previous month** Excel file

4. **Wait for processing** - you'll see progress messages

5. **Find your report** in the `output` folder:
   - File name format: `report_YYYYMMDD_HHMMSS.pptx`

</details>

<details>
<summary><b>macOS / Linux</b></summary>

1. **Open Terminal** and go to the cs-reporter folder:
   ```bash
   cd /path/to/cs-reporter
   ```

2. **Run the reporter**:
   ```bash
   ./run-reporter.sh
   ```

3. **Select your Excel files** when prompted:
   - First dialog: Select **current month** Excel file
   - Second dialog: Select **previous month** Excel file

4. **Wait for processing** - you'll see progress messages

5. **Find your report** in the `output` folder:
   - File name format: `report_YYYYMMDD_HHMMSS.pptx`

</details>

---

## Understanding the Excel File Requirements

Your Excel files must have:

### Retail Sheet
- Sheet name: `Tickets ADUS Tickets crea... 1` (or exact name from your system)
- Required columns:
  - `Ticket created - Date`
  - `Ticket solved - Date`
  - `Ticket satisfaction rating` (values: "good" or "good with comment")
  - `Support Category`

### Supplier Sheet
- Sheet name: `Tickets ADUS Tickets crea... 2` (or exact name from your system)
- Required columns: (same as Retail)
  - `Ticket created - Date`
  - `Ticket solved - Date`
  - `Ticket satisfaction rating`
  - `Support Category`
  - `Ticket organisation name` (for top organizations table)

---

## What Gets Generated

The tool extracts and calculates:

### Metrics
- **Month names** (current and previous)
- **Request counts** (total tickets)
- **Satisfaction ratings** (good ratings with/without comments)
- **Resolution time** (average days to resolve, filtered to ≤3 days)

### Tables
- **Support Categories** (retail and supplier) - with counts and "Uncategorized" for empty values
- **Top 5 Organizations** - by ticket count

### Comparison
All metrics include both current and previous month data for month-over-month comparison.

---

## Troubleshooting

### "Command not found: reporter" (Windows) or "command not found" (macOS/Linux)
**Solution:** Make sure you're using the launcher script:

**Windows:**
```cmd
run-reporter.bat
```

**macOS/Linux:**
```bash
./run-reporter.sh
```

Or activate the virtual environment first:

**Windows:**
```cmd
.venv\Scripts\activate.bat
reporter
```

**macOS/Linux:**
```bash
source .venv/bin/activate
reporter
```

### "Sheet not found" or "Column not found"
**Problem:** Your Excel file structure doesn't match the configuration.

**Solution:**
1. Check your actual sheet names in Excel
2. Open `config/mapping_v2.yaml`
3. Update the `sheet` names to match your Excel file exactly
4. Update the `column` names to match your column headers

Example:
```yaml
fields:
  re_req:
    operation: count_rows
    sheet: "Your Actual Sheet Name Here"  # Update this
```

### "Template file not found"
**Problem:** The PowerPoint template is missing.

**Solution:** Make sure `templates/report_template.pptx` exists and has all the required placeholders.

### Placeholders Still Showing in Output
**Problem:** The PowerPoint template has typos in placeholder names.

**Solution:** Check the console output after running - it will show:
```
Replaced fields: month, re_req, ...
⚠ Unreplaced fields: some_typo_field
```

Fix any typos in your template to match the field names exactly.

### Wrong Data in Tables
**Problem:** Table data looks incorrect or missing.

**Solution:**
1. Verify your Excel files have data in the expected columns
2. Check that column names in `config/mapping_v2.yaml` match your Excel exactly
3. Look for empty/null values - they'll be grouped as "Uncategorized"

---

## Customizing the Template

The PowerPoint template (`templates/report_template.pptx`) can be customized:

### Placeholders for Scalar Fields
Format: `{{field_name}}`

Examples:
- `{{month}}` - Current month name
- `{{prev_month}}` - Previous month name
- `{{re_req}}` - Retail requests (current)
- `{{re_prev_req}}` - Retail requests (previous)
- `{{re_sat}}` - Retail satisfaction count
- `{{re_reso}}` - Retail average resolution time

See `config/mapping_v2.yaml` for the complete list of available fields.

### Placeholders for Dynamic Tables
Tables need **two rows**:
- **Row 0:** Header row (e.g., "Support Category | Count")
- **Row 1:** Template row with placeholders

Template row format: `{{table:table_name}} {{field1}} | {{field2}}`

Example for retail support categories:
- Row 0: `Support Category | Count`
- Row 1: `{{table:re_sup_cat}} {{re_cat}} | {{re_cat_count}}`

### Formatting Tips
- **Header row:** Format however you want (bold, larger font, colored background)
- **Template row:** Format once - all data rows will inherit this formatting
- **Placeholders:** Don't worry about font/color - they'll be replaced with the template formatting

---

## Updating to a New Version

If you receive an updated version of cs-reporter:

1. **Back up your files:**

   **Windows:**
   ```cmd
   copy config\mapping_v2.yaml config\mapping_v2.yaml.backup
   copy templates\report_template.pptx templates\report_template.pptx.backup
   ```

   **macOS/Linux:**
   ```bash
   cp config/mapping_v2.yaml config/mapping_v2.yaml.backup
   cp templates/report_template.pptx templates/report_template.pptx.backup
   ```

2. **Replace the cs-reporter folder** with the new version

3. **Restore your custom files:**

   **Windows:**
   ```cmd
   copy config\mapping_v2.yaml.backup config\mapping_v2.yaml
   copy templates\report_template.pptx.backup templates\report_template.pptx
   ```

   **macOS/Linux:**
   ```bash
   cp config/mapping_v2.yaml.backup config/mapping_v2.yaml
   cp templates/report_template.pptx.backup templates/report_template.pptx
   ```

4. **Re-run setup:**

   **Windows:**
   ```cmd
   setup.bat
   ```

   **macOS/Linux:**
   ```bash
   ./setup.sh
   ```

---

## Getting Help

If you encounter issues:

1. **Check the console output** - it shows what went wrong
2. **Review this guide** - especially the Troubleshooting section
3. **Check your configuration** - ensure `config/mapping_v2.yaml` matches your Excel structure
4. **Verify your template** - ensure all placeholders are spelled correctly

---

## Quick Reference

### Common Commands

**Windows:**
```cmd
# First time setup
setup.bat

# Generate a report
run-reporter.bat

# Or manually
.venv\Scripts\activate.bat
reporter
```

**macOS/Linux:**
```bash
# First time setup
./setup.sh

# Generate a report
./run-reporter.sh

# Or manually
source .venv/bin/activate
reporter
```

### File Locations
- **Reports:** `output/report_*.pptx`
- **Template:** `templates/report_template.pptx`
- **Configuration:** `config/mapping_v2.yaml`

### Excel Requirements
- Two files needed: current month + previous month
- Must have matching sheet names and column headers
- All data should be in the expected format

---

**That's it!** You should now be able to generate monthly reports automatically. 🎉
