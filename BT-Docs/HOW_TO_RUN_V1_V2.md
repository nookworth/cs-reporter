# How to Run V1 vs V2

## Quick Reference

| Version | Command | Config Used |
|---------|---------|-------------|
| **V1** (Original) | `python -m src.main` | `config/mapping.yaml` |
| **V2** (Refactored) | `python run_v2.py` | `config/mapping_v2.yaml` |

---

## V1 Files (Original System)

### Source Code:
- `src/main.py` - Entry point
- `src/config.py`
- `src/excel_utils.py`
- `src/excel_reader.py`

### Configuration:
- `config/mapping.yaml` - Suffix-based format

### How to Run:
```bash
# Windows
python -m src.main

# Or with custom config
python -m src.main --config config/mapping.yaml
```

---

## V2 Files (Refactored System)

### Source Code:
- `src/main_v2.py` - Entry point (uses V2)
- `src/config_v2.py`
- `src/excel_utils_v2.py`
- `src/excel_reader_v2.py`

### Configuration:
- `config/mapping_v2.yaml` - Operation-based format
- `config/schema_v2.yaml` - Documentation

### How to Run:

**Option 1: Using Python script**
```bash
python run_v2.py
```

**Option 2: Using batch file (Windows)**
```bash
run_v2.bat
```

**Option 3: Direct module call**
```bash
python -m src.main_v2 --config config/mapping_v2.yaml
```

---

## All V2 Files

Files that need to be edited/used to run V2:

### Core V2 Files:
1. ✅ `src/main_v2.py` - Main entry point for V2
2. ✅ `src/config_v2.py` - Config loader with validation
3. ✅ `src/excel_utils_v2.py` - Excel utilities with filters
4. ✅ `src/excel_reader_v2.py` - Excel reader with operations
5. ✅ `config/mapping_v2.yaml` - Production config
6. ✅ `config/demo_mapping_v2.yaml` - Demo config
7. ✅ `config/schema_v2.yaml` - Schema documentation

### V2 Launchers:
8. ✅ `run_v2.py` - Python launcher script
9. ✅ `run_v2.bat` - Windows batch launcher

### V2 Tests:
10. ✅ `test_v2_config.py` - Full integration test
11. ✅ `test_filters.py` - Filter tests
12. ✅ `test_validation.py` - Validation tests

---

## Comparison

### V1 (Original)
```
User runs: python -m src.main
    ↓
src/main.py
    ↓
Imports: config, excel_reader
    ↓
Uses: config/mapping.yaml
    ↓
Suffix-based operations (_req, _sat, _prev_)
```

### V2 (Refactored)
```
User runs: python run_v2.py
    ↓
run_v2.py → src/main_v2.py
    ↓
Imports: config_v2, excel_reader_v2
    ↓
Uses: config/mapping_v2.yaml
    ↓
Operation-based (explicit operations, filters)
```

---

## Testing V2

### Quick Test (Demo Data):
```bash
python test_v2_config.py
```
This generates demo Excel files and creates a PowerPoint report.

### Full Test (Your Data):
```bash
python run_v2.py
```
Then select your actual Excel files when prompted.

---

## File Naming Convention

**All V2 files end with `_v2`:**
- Source: `*_v2.py`
- Config: `*_v2.yaml`
- Launchers: `*_v2.py`, `*_v2.bat`
- Tests: `test_*` (all tests are for V2)

**V1 files (no suffix):**
- Source: `config.py`, `excel_utils.py`, `excel_reader.py`
- Config: `mapping.yaml`, `demo_mapping.yaml`
- Main: `main.py`

---

## Summary

✅ **To run V1:** `python -m src.main`
✅ **To run V2:** `python run_v2.py` or `run_v2.bat`

Both systems are completely independent and can run side-by-side.
