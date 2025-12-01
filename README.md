# CS Reporter

Excel to PowerPoint report generator. Automatically creates PowerPoint presentations from Excel data using a template-based approach.

## Features

- Simple CLI with file dialog - just run `reporter`
- Template-based PowerPoint generation
- Excel data extraction using pandas
- YAML configuration for field mapping
- Automatic formatting (currency, percentages, numbers)
- **Dynamic tables with auto-aggregation** - populate tables with variable-length data

## Project Structure

```
cs-reporter/
├── src/
│   ├── __init__.py
│   ├── main.py              # CLI entry point with file dialog
│   ├── excel_reader.py      # Extract data from Excel using pandas
│   ├── ppt_writer.py        # Generate PowerPoint from template
│   └── config.py            # Load and validate YAML config
├── templates/
│   └── report_template.pptx # PowerPoint template with placeholders
├── config/
│   └── mapping.yaml         # Excel → PowerPoint field mapping
├── output/                  # Generated reports go here
├── requirements.txt
├── setup.py
└── README.md
```

## Installation

1. Clone this repository
2. Install the package in development mode:

```bash
pip install -e .
```

This will install all dependencies and create the `reporter` command.

## Setup

### 1. Create Your PowerPoint Template

1. Take one of your existing PowerPoint files
2. Remove month-specific data
3. Replace data with placeholders using `{{field_name}}` syntax

Example:
```
Revenue: {{revenue}}
Customers: {{customer_count}}
Growth: {{growth_rate}}
Report for {{report_month}}
```

Save this as `templates/report_template.pptx`

### 2. Configure Field Mappings

Edit `config/mapping.yaml` to specify:
- Which Excel sheet and cell each field comes from
- How to format each field (currency, percentage, number, text)

Example:
```yaml
excel_fields:
  revenue:
    sheet: "Summary"
    cell: "B5"

  customer_count:
    sheet: "Summary"
    cell: "D10"

formatting:
  revenue:
    type: "currency"
    decimals: 2
```

### 3. Prepare Your Excel File

Make sure your Excel file has data at the cells specified in `mapping.yaml`.

## Usage

Simply run:

```bash
reporter
```

A file dialog will open. Select your Excel file, and the tool will:
1. Read data from the configured cells
2. Apply formatting rules
3. Replace placeholders in the PowerPoint template
4. Save the result to the `output/` directory

The output file will be named with a timestamp: `report_YYYYMMDD_HHMMSS.pptx`

## Configuration Reference

### Field Mapping

Each field in `excel_fields` needs:
- `sheet`: Excel sheet name
- `cell`: Cell reference (e.g., "B5", "D10")

### Formatting Types

- `currency`: Formats as `$1,234.56`
- `percentage`: Formats as `12.5%`
- `number`: Formats as `1,234.56`
- `text`: No formatting (default)

### Placeholder Syntax

In your PowerPoint template, use `{{field_name}}` where `field_name` matches a key in `excel_fields`.

### Dynamic Tables

Dynamic tables allow you to populate PowerPoint tables with a variable number of rows from Excel data. This is perfect for lists of items, support categories, metrics, etc.

#### How It Works

1. **In Excel**: Data is in consecutive rows (e.g., rows 5-15)
2. **In Config**: Specify the sheet, starting row, columns, and aggregation rules
3. **In PowerPoint**: Create a table with ONE template row containing placeholders
4. **Script**: Automatically reads data, aggregates if needed, and populates the table

#### Configuration Example

```yaml
table_fields:
  support_categories:
    sheet: "Support Data"
    start_row: 5              # First row of data (1-indexed)
                              # Header row is assumed to be start_row - 1

    # Columns to read from Excel
    columns:
      - name: "category"         # Field name for placeholders
        col: "Support Category"  # Column heading name in Excel

    # Aggregation - COUNT occurrences
    aggregate: true              # Enable aggregation
    aggregation_type: "count"    # Count how many times each category appears
    group_by: "category"         # Group by this column
    count_column: "count"        # Name for the count field (used in placeholders)

    # Format the aggregated values
    format_values: true
    value_format:
      type: "number"
      decimals: 0
```

**Aggregation Types:**

1. **`aggregation_type: "count"`** - Counts how many times each category appears
   - Only reads the grouping column from Excel
   - Perfect for: "How many support tickets in each category?"
   - Uses `count_column` to name the output field

2. **`aggregation_type: "sum"`** - Sums values from a numeric column
   - Reads both the grouping column and a numeric column
   - Perfect for: "Total revenue by region"
   - Uses `sum_column` to specify which column to sum
   - Example:
     ```yaml
     aggregation_type: "sum"
     group_by: "region"
     sum_column: "revenue"  # Column with numbers to sum
     ```

**Important:** The `col` field specifies the **column heading name** in your Excel file. The script automatically searches for this heading in the row immediately before `start_row` and reads data from that column. This is more robust than hardcoding column letters!

#### PowerPoint Template Setup

Create a table in your PowerPoint with **one row** as the template:

```
┌───────────────────────────┬───────────┐
│ {{table:support_categories}} {{category}} │ {{count}} │
└───────────────────────────┴───────────┘
```

**Important:**
- First cell must contain `{{table:table_name}}` to mark this as a dynamic table
- Other cells contain placeholders matching your column names from the config

#### Result Examples

**Example 1: Count Aggregation**

If your Excel has (only one column - "Support Category"):
```
Row 1: [Header: "Support Category"]
Row 2: Technical
Row 3: Billing
Row 4: Technical
Row 5: General
Row 6: Billing
Row 7: Technical
```

With `aggregation_type: "count"`, the script will:
1. Read all category values
2. Group by category
3. Count occurrences of each
4. Populate the table:

```
┌──────────┬───────┐
│ Technical│ 3     │
│ Billing  │ 2     │
│ General  │ 1     │
└──────────┴───────┘
```

**Example 2: Sum Aggregation**

If your Excel has (two columns - "Region" and "Revenue"):
```
Row 1: [Header: "Region" | "Revenue"]
Row 2: East           | 1000
Row 3: West           | 500
Row 4: East           | 1500
```

With `aggregation_type: "sum"` and `sum_column: "revenue"`, the script will:
1. Read region and revenue values
2. Group by region
3. Sum the revenue for each region
4. Populate the table:

```
┌──────┬──────────┐
│ East │ 2,500    │
│ West │ 500      │
└──────┴──────────┘
```

#### Without Aggregation

Set `aggregate: false` to use rows as-is without grouping/counting/summing.

#### Auto-Detection

- **Column headers**: The script looks for column headings in the row immediately before `start_row` (i.e., at row `start_row - 1`)
- **End of data**: The script automatically stops reading when it hits an empty row, so you don't need to specify the end row

**Example Excel layout:**
```
Row 1: [Header: "Support Category" | "Count"]  ← Headers at start_row - 1
Row 2: Technical                   | 5         ← Data starts at start_row
Row 3: Billing                     | 10
Row 4: Technical                   | 15
Row 5: [Empty]                                 ← Auto-stops here
```

## Example Workflow

1. Run `reporter`
2. Select `monthly_data_november_2025.xlsx`
3. Tool reads cells from Excel based on `mapping.yaml`
4. Tool replaces `{{placeholders}}` in the template
5. Saves result to `output/report_20251130_143022.pptx`

## Development

To run without installing:

```bash
python -m src.main
```

To modify field mappings, edit `config/mapping.yaml` and rerun.

## Troubleshooting

**Config file not found:**
- Ensure `config/mapping.yaml` exists

**Template not found:**
- Ensure `templates/report_template.pptx` exists
- Check that `template_path` in `mapping.yaml` is correct

**No replacements made:**
- Check that placeholder names in PowerPoint match field names in `mapping.yaml`
- Placeholders must use double curly braces: `{{field_name}}`

**Excel read error:**
- Verify sheet names and cell references in `mapping.yaml`
- Ensure Excel file has data at the specified cells

**Table not populating:**
- Check that the first cell of your table row contains `{{table:table_name}}`
- Verify the table name matches the key in `table_fields` in `mapping.yaml`
- Check that column names in the config match the placeholders in the table

**Aggregation not working:**
- Verify `aggregate: true` is set in the table config
- Check that `group_by` and `sum_column` match your column names
- Ensure the sum column contains numeric values in Excel
