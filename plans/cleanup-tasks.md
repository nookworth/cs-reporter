# Ralph Tasks — CS Reporter Cleanup

> **NOTE (2026-06-07):** The `plans/migrate-v2-to-default.md` migration has
> superseded this plan's "V1 is still the default" assumption. V1 files are
> now deleted. These cleanup tasks are still relevant for V2/V3 modules.

- [x] Create tests/test_smoke.py — headless end-to-end smoke test using
  demo_data/ Excel fixtures with hardcoded baseline assertions. Verify:
  .venv/bin/python tests/test_smoke.py exits 0.

- [ ] Delete templates/src/ — remove the entire directory and its
  __init__.py. Nothing in the project imports from it. Verify:
  .venv/bin/python tests/test_smoke.py

- [ ] Remove __main__ blocks — delete the `if __name__ == "__main__":`
  blocks (and any imports only used by them) from src/excel_utils_v2.py
  and src/excel_utils.py. Both contain a hardcoded personal file path.
  Verify: .venv/bin/python tests/test_smoke.py

- [ ] Fix mapping_v2.yaml template path — change line 8 of
  config/mapping_v2.yaml from `templates/demo_template.pptx` to
  `templates/report_template.pptx`. Verify:
  .venv/bin/python tests/test_smoke.py

- [ ] Fix app_v3.py unsafe_allow_html — replace the footer
  st.markdown block that uses unsafe_allow_html=True with
  st.caption("CS Reporter v2.0 | Built with Streamlit"). Verify:
  .venv/bin/python tests/test_smoke.py

- [ ] Improve error handling in ppt_writer_v3.py — add `import logging`
  and replace all three bare `except Exception as e: print(...)` blocks
  in _add_charts_to_report() with logging.warning(...) calls. Do not
  change whether errors are swallowed. Verify:
  .venv/bin/python tests/test_smoke.py

- [ ] Add type annotations to excel_utils_v2.py — annotate every
  function signature. Use Path | str for file paths, list[dict] for
  filter lists. Do not change any logic. Verify:
  .venv/bin/python tests/test_smoke.py

- [ ] Add type annotations to excel_reader_v2.py — annotate __init__,
  extract_data, _dispatch_operation, _extract_field_group. Do not
  change any logic. Verify: .venv/bin/python tests/test_smoke.py

- [ ] Add type annotations to config_v2.py, chart_utils_v3.py, and
  ppt_writer_v3.py — annotate all public functions and methods. Do not
  change any logic. Verify: .venv/bin/python tests/test_smoke.py

- [ ] Add docstrings to V2 functions missing them — add a concise
  one-line docstring to every public function and method across
  src/excel_utils_v2.py, src/excel_reader_v2.py, src/config_v2.py,
  src/chart_utils_v3.py, and src/ppt_writer_v3.py that currently has
  none. Do not modify existing docstrings. Verify:
  .venv/bin/python tests/test_smoke.py

- [ ] Create tests/test_filters.py — unit tests for apply_filters() in
  src/excel_utils_v2.py using in-memory pandas DataFrames (no Excel
  files). Cover all 12 operators: equals, not_equals, greater_than,
  greater_than_or_equal, less_than, less_than_or_equal, contains,
  not_contains, starts_with, ends_with, is_null, is_not_null. Include
  NaN edge cases, case insensitivity for string ops, and chaining
  multiple filters. Verify:
  .venv/bin/python -m pytest tests/test_filters.py -v

- [ ] Create tests/test_config_validation.py — unit tests for
  validate_fields() in src/config_v2.py. Cover: missing 'operation' key
  raises ValueError, unknown operation raises ValueError with valid op
  names in the message, missing required parameter raises ValueError,
  valid config passes without raising. Construct minimal dicts inline —
  no files needed. Verify:
  .venv/bin/python -m pytest tests/test_config_validation.py -v
