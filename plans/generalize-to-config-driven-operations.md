# Plan: Generalize cs-reporter (Version 1)

> **Note:** This plan is designed for any developer or coding assistant. Code locations are referenced by function/class name rather than line numbers.

## Goal
Transform cs-reporter from suffix-based operation detection to explicit config-driven operations, making it usable with any Excel/PowerPoint combination.

## Future Direction
This refactor is a prerequisite for LLM-powered config generation. The config format should be:
- **Regular**: Consistent structure across all operations (easier for LLM to generate)
- **Explicit**: No implicit behavior based on naming conventions
- **Schema-documented**: A formal schema that an LLM can reference when generating configs

---

## New Config Format

```yaml
template_path: templates/report_template.pptx
output_dir: output

sources:
  current:
    required: true
  previous:
    required: false  # Makes dual-file optional

fields:
  # Row counting
  total_tickets:
    source: current
    sheet: "Tickets"
    operation: count_rows

  # Value counting (replaces hardcoded "good"/"good with comment")
  satisfied_count:
    source: current
    sheet: "Tickets"
    operation: count_value
    column: "Satisfaction"
    value: "good"
    case_sensitive: false

  # Date diff (replaces hardcoded columns and filters)
  avg_resolution:
    source: current
    sheet: "Tickets"
    operation: avg_date_diff
    start_column: "Created Date"
    end_column: "Resolved Date"
    filters:  # Optional, atomic filter syntax for LLM compatibility
      - column: "Assignee"
        operator: equals
        value: "Leo Brown"
      - column: "_result"  # Special: references the computed value
        operator: greater_than
        value: 72
        action: exclude

  # Sum a numeric column
  total_revenue:
    source: current
    sheet: "Sales"
    operation: sum
    column: "Amount"

  # Month parsing
  report_month:
    source: current
    sheet: "Tickets"
    operation: parse_month
    column: "Created Date"

  # Simple value read
  company_name:
    source: current
    sheet: "Summary"
    operation: read_value
    column: "Company"

table_fields:
  # Keep existing format (already fairly generic)
```

---

## Files to Modify

### 1. `src/config.py` (add validation)
- Add `OPERATIONS` registry with required/optional params
- Add `validate_fields()` function
- Validate operation exists and has required params

### 2. `config/schema.yaml` (new file)
- Document all valid operations and their parameters
- Define valid enum values (e.g., `source: current | previous`, `operator: equals | greater_than | less_than`)
- This becomes the reference for LLM config generation

### 3. `src/excel_reader.py` (main refactor)
- Replace `_extract_field_group()` with `_dispatch_operation()`
- Remove suffix detection (`_req`, `_sat`, `_reso`, `_prev_`)
- Use `field_config['operation']` and `field_config['source']` instead
- Keep table extraction mostly unchanged

### 4. `src/excel_utils.py` (parameterize)
- `count_column_value()`: already has `value_to_count` param (no change needed)
- `calculate_average_resolution_time()`: add `filters` param, remove hardcoded `72` and `"Leo Brown"`
- Add `sum_column()` function for the new `sum` operation

### 5. `config/mapping.yaml` (convert format)
- Convert all fields to explicit operation format
- Replace `_prev_` naming with `source: previous`

---

## Implementation Steps

### Step 1: Add Operation Validation (`src/config.py`)

Add at module level (after imports, before `load_config()`):

```python
OPERATIONS = {
    'count_rows': {'required': [], 'optional': ['sheet', 'filters']},
    'count_value': {'required': ['column', 'value'], 'optional': ['case_sensitive', 'filters']},
    'sum': {'required': ['column'], 'optional': ['filters']},
    'avg_date_diff': {'required': ['start_column', 'end_column'], 'optional': ['filters']},
    'parse_month': {'required': ['column'], 'optional': []},
    'parse_previous_month': {'required': ['column'], 'optional': []},
    'read_value': {'required': ['column'], 'optional': []},
}

def validate_fields(config):
    """Validate all fields have valid operations and required params."""
    fields = config.get('fields', {})
    for name, field_config in fields.items():
        if 'operation' not in field_config:
            raise ValueError(f"Field '{name}' missing 'operation'")
        op = field_config['operation']
        if op not in OPERATIONS:
            raise ValueError(f"Field '{name}' has unknown operation '{op}'")
        for param in OPERATIONS[op]['required']:
            if param not in field_config:
                raise ValueError(f"Field '{name}' missing required param '{param}'")
```

Call `validate_fields(config)` in `load_config()` before returning.

---

### Step 1b: Create Config Schema (`config/schema.yaml`)

This schema documents the config format for future LLM-powered generation:

```yaml
# Config Schema for cs-reporter
# Used by LLM to generate valid configurations

config:
  template_path: string  # Path to PowerPoint template
  output_dir: string     # Output directory for generated reports

  sources:
    current:
      required: boolean  # Always true
    previous:
      required: boolean  # Whether previous month file is needed

  fields:
    <field_name>:        # Arbitrary name, becomes {{field_name}} placeholder
      source: enum [current, previous]
      sheet: string      # Excel sheet name
      operation: enum [count_rows, count_value, sum, avg_date_diff, parse_month, parse_previous_month, read_value]
      # Additional params depend on operation (see below)

operations:
  count_rows:
    description: "Count total rows in a sheet"
    required_params: []
    optional_params: [filters]

  count_value:
    description: "Count rows where column matches a value"
    required_params:
      - column: string   # Column name to check
      - value: string    # Value to match
    optional_params:
      - case_sensitive: boolean  # Default: false
      - filters: list[filter]

  sum:
    description: "Sum values in a numeric column"
    required_params:
      - column: string   # Column name to sum
    optional_params:
      - filters: list[filter]

  avg_date_diff:
    description: "Average difference between two date columns"
    required_params:
      - start_column: string  # Start date column
      - end_column: string    # End date column
    optional_params:
      - filters: list[filter]

  parse_month:
    description: "Extract month name from first date in column"
    required_params:
      - column: string
    optional_params: []

  parse_previous_month:
    description: "Extract previous month name from first date in column"
    required_params:
      - column: string
    optional_params: []

  read_value:
    description: "Read first non-null value from column"
    required_params:
      - column: string
    optional_params: []

filter:
  description: "Atomic filter condition"
  params:
    - column: string     # Column name, or "_result" for computed value
    - operator: enum [equals, not_equals, greater_than, less_than]
    - value: any         # Value to compare against
    - action: enum [exclude, include_only]  # Default: exclude
```

---

### Step 2: Make Filters Configurable (`src/excel_utils.py`)

Update `calculate_average_resolution_time()`:

**Current** (hardcoded filter logic inside the function):
```python
if resolution_time_hours > 72 and assignee == "Leo Brown":
    continue
```

**New** (configurable with atomic filter syntax):
```python
def apply_filters(row, filters, computed_value=None):
    """
    Apply atomic filters to a row. Returns True if row should be excluded.

    Filter format:
      - column: "Assignee" or "_result" (for computed value)
        operator: equals | not_equals | greater_than | less_than
        value: <comparison value>
        action: exclude (default) | include_only
    """
    for f in filters:
        col = f.get('column')
        operator = f.get('operator')
        target = f.get('value')
        action = f.get('action', 'exclude')

        # Get the value to compare
        if col == '_result':
            actual = computed_value
        else:
            actual = row.get(col, '')

        # Evaluate the condition
        match = False
        if operator == 'equals':
            match = actual == target
        elif operator == 'not_equals':
            match = actual != target
        elif operator == 'greater_than':
            match = actual > target
        elif operator == 'less_than':
            match = actual < target

        # Apply action
        if action == 'exclude' and match:
            return True  # Exclude this row
        elif action == 'include_only' and not match:
            return True  # Exclude rows that don't match

    return False  # Keep this row

def calculate_average_resolution_time(
    excel_path,
    sheet_name,
    columns=None,
    filters=None  # NEW PARAM
):
    # ... existing code ...

    # Replace hardcoded filter with config-driven filtering
    if filters and apply_filters(row, filters, computed_value=resolution_time_hours):
        continue
```

---

### Step 3: Refactor ExcelReader (`src/excel_reader.py`)

**Replace the `__init__` constructor:**
```python
def __init__(self, excel_path, config, previous_excel_path=None):
    self.excel_path = Path(excel_path)
    self.previous_excel_path = Path(previous_excel_path) if previous_excel_path else None
    self.config = config
    self.fields = config.get("fields", {})
    self.table_fields = config.get("table_fields", {})
```

**Replace the entire `_extract_field_group()` method with this new method:**
```python
def _dispatch_operation(self, field_name, field_config):
    """Execute the configured operation for a field."""
    operation = field_config.get('operation')
    source = field_config.get('source', 'current')
    sheet = field_config.get('sheet')

    # Select Excel file based on source
    if source == 'previous':
        if not self.previous_excel_path:
            return None
        excel_path = self.previous_excel_path
    else:
        excel_path = self.excel_path

    # Dispatch to operation
    if operation == 'count_rows':
        return excel_utils.count_rows(excel_path, sheet)

    elif operation == 'count_value':
        return excel_utils.count_column_value(
            excel_path, sheet,
            field_config['column'],
            field_config['value']
        )

    elif operation == 'sum':
        return excel_utils.sum_column(
            excel_path, sheet,
            field_config['column'],
            filters=field_config.get('filters')
        )

    elif operation == 'avg_date_diff':
        return excel_utils.calculate_average_resolution_time(
            excel_path, sheet,
            columns=[field_config['start_column'], field_config['end_column']],
            filters=field_config.get('filters')
        )

    elif operation == 'parse_month':
        val = excel_utils.read_column_value(excel_path, sheet, field_config['column'])
        return excel_utils.parse_month_from_date(val)

    elif operation == 'parse_previous_month':
        val = excel_utils.read_column_value(excel_path, sheet, field_config['column'])
        return excel_utils.parse_previous_month_from_date(val)

    elif operation == 'read_value':
        return excel_utils.read_column_value(excel_path, sheet, field_config['column'])

    else:
        raise ValueError(f"Unknown operation: {operation}")
```

**Simplify `extract_data()` to use the new dispatch method:**
```python
def extract_data(self):
    data = {}

    # Extract scalar fields
    print("\nExtracting fields...")
    for field_name, field_config in self.fields.items():
        try:
            value = self._dispatch_operation(field_name, field_config)
            data[field_name] = value
            print(f"  {field_name}: {value}")
        except Exception as e:
            print(f"  Warning: {field_name} failed: {e}")
            data[field_name] = None

    # Extract tables (keep existing table extraction logic unchanged)
    # ... existing table_fields extraction code stays the same ...

    return data
```

---

### Step 4: Convert Config (`config/mapping.yaml`)

Convert from:
```yaml
retail_excel_fields:
  sheet: "Tickets ADUS Tickets crea... 1"
  re_req:
    cell:
  re_sat:
    cell: "Ticket satisfaction rating"
```

To:
```yaml
fields:
  re_req:
    source: current
    sheet: "Tickets ADUS Tickets crea... 1"
    operation: count_rows

  re_prev_req:
    source: previous
    sheet: "Tickets ADUS Tickets crea... 1"
    operation: count_rows

  re_sat:
    source: current
    sheet: "Tickets ADUS Tickets crea... 1"
    operation: count_value
    column: "Ticket satisfaction rating"
    value: "good"

  re_sat_c:
    source: current
    sheet: "Tickets ADUS Tickets crea... 1"
    operation: count_value
    column: "Ticket satisfaction rating"
    value: "good with comment"

  re_reso:
    source: current
    sheet: "Tickets ADUS Tickets crea... 1"
    operation: avg_date_diff
    start_column: "Ticket created - Date"
    end_column: "Ticket solved - Date"
    filters:
      - column: "Assignee name"
        operator: equals
        value: "Leo Brown"
      - column: "_result"
        operator: greater_than
        value: 72
        action: exclude

  month:
    source: current
    sheet: "Tickets ADUS Tickets crea... 1"
    operation: parse_month
    column: "Ticket created - Date"

  prev_month:
    source: current
    sheet: "Tickets ADUS Tickets crea... 1"
    operation: parse_previous_month
    column: "Ticket created - Date"
```

---

## Testing & Verification

1. **Before changes**: Run `reporter` with current Excel files, save output
2. **After changes**: Run `reporter` with same files, compare output
3. **Verify**: All placeholder values match between old and new runs
4. **Test validation**: Try invalid configs (missing operation, unknown operation, missing required param)

---

## Summary of Changes

| File | Lines Changed | What Changes |
|------|---------------|--------------|
| `src/config.py` | +35 lines | Add OPERATIONS registry and validate_fields() |
| `config/schema.yaml` | +70 lines (new) | Document config format for LLM generation |
| `src/excel_reader.py` | ~120 lines rewritten | Replace suffix detection with operation dispatch |
| `src/excel_utils.py` | ~50 lines | Add apply_filters(), sum_column(), update avg function |
| `config/mapping.yaml` | Full rewrite | Convert to explicit operation format |

**No changes needed:**
- `src/ppt_writer.py` - Already generic
- `src/main.py` - Just calls ExcelReader

---

## Why These Changes Support LLM Inference

When we eventually build the LLM-powered config generator:

1. **Schema file** → The LLM prompt includes this as the "target format"
2. **Atomic filters** → Simpler pattern for LLM to generate (each filter is independent)
3. **Explicit operations** → No implicit behavior to infer; LLM just matches patterns to operations
4. **Validation** → Clear error messages help LLM self-correct if it generates invalid config
