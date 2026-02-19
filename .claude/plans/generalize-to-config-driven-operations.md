# Plan: Generalize cs-reporter (Version 1)

## Goal
Transform cs-reporter from suffix-based operation detection to explicit config-driven operations, making it usable with any Excel/PowerPoint combination.

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
    filters:  # Optional, replaces hardcoded Leo Brown/72hr filter
      - column: "Assignee"
        equals: "Leo Brown"
        max_hours: 72

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

### 2. `src/excel_reader.py` (main refactor)
- Replace `_extract_field_group()` with `_dispatch_operation()`
- Remove suffix detection (`_req`, `_sat`, `_reso`, `_prev_`)
- Use `field_config['operation']` and `field_config['source']` instead
- Keep table extraction mostly unchanged

### 3. `src/excel_utils.py` (parameterize)
- `count_column_value()`: already has `value_to_count` param (no change needed)
- `calculate_average_resolution_time()`: add `filters` param, remove hardcoded `72` and `"Leo Brown"`

### 4. `config/mapping.yaml` (convert format)
- Convert all fields to explicit operation format
- Replace `_prev_` naming with `source: previous`

---

## Implementation Steps

### Step 1: Add Operation Validation (`src/config.py`)

Add after line 65:

```python
OPERATIONS = {
    'count_rows': {'required': [], 'optional': ['sheet']},
    'count_value': {'required': ['column', 'value'], 'optional': ['case_sensitive']},
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

### Step 2: Make Filters Configurable (`src/excel_utils.py`)

Update `calculate_average_resolution_time()` (lines 217-305):

**Current** (hardcoded at line 290):
```python
if resolution_time_hours > 72 and assignee == "Leo Brown":
    continue
```

**New** (configurable):
```python
def calculate_average_resolution_time(
    excel_path,
    sheet_name,
    columns=["Ticket created - Date", "Ticket solved - Date"],
    assignee_column="Assignee name",
    filters=None  # NEW PARAM
):
    # ... existing code ...

    # Replace hardcoded filter with config-driven filtering
    if filters:
        skip = False
        for f in filters:
            col_val = row.get(f.get('column', assignee_column), '')
            if f.get('equals') and col_val == f['equals']:
                if f.get('max_hours') and resolution_time_hours > f['max_hours']:
                    skip = True
                    break
        if skip:
            continue
```

---

### Step 3: Refactor ExcelReader (`src/excel_reader.py`)

**Replace constructor** (lines 15-32):
```python
def __init__(self, excel_path, config, previous_excel_path=None):
    self.excel_path = Path(excel_path)
    self.previous_excel_path = Path(previous_excel_path) if previous_excel_path else None
    self.config = config
    self.fields = config.get("fields", {})
    self.table_fields = config.get("table_fields", {})
```

**Replace `_extract_field_group()` (lines 109-231) with:**
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

**Update `extract_data()` (lines 34-107):**
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

    # Extract tables (existing code unchanged)
    # ... lines 56-100 stay the same ...

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
        equals: "Leo Brown"
        max_hours: 72

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
| `src/config.py` | +30 lines | Add OPERATIONS registry and validate_fields() |
| `src/excel_reader.py` | ~120 lines rewritten | Replace suffix detection with operation dispatch |
| `src/excel_utils.py` | ~15 lines | Add filters param to calculate_average_resolution_time() |
| `config/mapping.yaml` | Full rewrite | Convert to explicit operation format |

**No changes needed:**
- `src/ppt_writer.py` - Already generic
- `src/main.py` - Just calls ExcelReader
