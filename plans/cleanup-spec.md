# PRD: Code Cleanup & Quality Pass (Post-V2)

> **NOTE (2026-06-07):** The `plans/migrate-v2-to-default.md` migration has been
> completed — V2 is now the default CLI entry point and V1 files have been
> deleted. The "V1 is still the default" context in this spec has been
> superseded. The remaining cleanup tasks (type annotations, docstrings, etc.)
> are still relevant for V2/V3 modules.

## Goal

Clean up dead code, fix known bugs, add type annotations and docstrings,
and establish a test suite for the V2 modules introduced by collaborator
Ben Tran, without breaking existing functionality.

## Scope

- **In:** V2 source files (`src/*_v2.py`, `src/chart_utils_v3.py`,
  `src/ppt_writer_v3.py`, `app_v3.py`), dead code removal, new tests
- **Out:** V1 files (`src/excel_reader.py`, `src/excel_utils.py`,
  `src/config.py`, `src/main.py`) — do not modify these except to
  remove the `__main__` block from `src/excel_utils.py`

## Context

CS Reporter has two parallel systems:

- **V1** (suffix-based field detection) — still the default `reporter`
  CLI entry point. Leave it alone for now.
- **V2** (explicit operation-based config, generic filter system) — the
  target architecture, introduced by Ben. Used by `run_v2.py` and the
  Streamlit web UI (`app_v3.py`).

A smoke test already exists at `tests/test_smoke.py` and passes. It
runs the full V2 pipeline against scrubbed demo Excel files in
`demo_data/` and asserts known baseline values. Use it as your
regression gate after every change.

## Requirements

1. Delete `templates/src/` — the entire directory is confirmed unused
   dead code (nothing imports from it).

2. Remove `__main__` blocks with hardcoded personal file paths from
   `src/excel_utils_v2.py` and `src/excel_utils.py`. Both reference
   `~/Downloads/ADUS Monthly review (21).xlsx`.

3. Fix `config/mapping_v2.yaml` line 8: change `template_path` from
   `templates/demo_template.pptx` to `templates/report_template.pptx`.

4. Fix `app_v3.py` footer: replace `st.markdown(..., unsafe_allow_html=True)`
   with `st.caption("CS Reporter v2.0 | Built with Streamlit")`.

5. Improve error reporting in `src/ppt_writer_v3.py`: add
   `import logging` and replace the three bare
   `except Exception as e: print(...)` blocks in
   `_add_charts_to_report()` with `logging.warning(...)` calls.

6. Add type annotations to all public functions and methods in:
   `src/excel_utils_v2.py`, `src/excel_reader_v2.py`,
   `src/config_v2.py`, `src/chart_utils_v3.py`, `src/ppt_writer_v3.py`.
   Do not change any logic.

7. Add a concise docstring to every public function and method in the
   V2 modules listed above that currently lacks one. Do not modify
   existing docstrings.

8. Create `tests/test_filters.py` — unit tests for `apply_filters()` in
   `src/excel_utils_v2.py`. Use in-memory pandas DataFrames (no Excel
   files). Cover all 12 operators, NaN edge cases, case insensitivity,
   and chaining multiple filters.

9. Create `tests/test_config_validation.py` — unit tests for
   `validate_fields()` in `src/config_v2.py`. Test: missing `operation`
   key, unknown operation, missing required parameter, and a fully valid
   config. No files needed — construct minimal dicts inline.

## Constraints

- Do not rename any V2 files or change their public APIs.
- Do not modify `config/mapping.yaml` or `config/demo_mapping.yaml`.
- Python version: 3.12+. Use `X | Y` union syntax for type hints, not
  `Optional[X]` or `Union[X, Y]`.
- Do not add new dependencies.

## Acceptance Criteria

- `.venv/bin/python tests/test_smoke.py` exits 0 and prints
  "All smoke tests passed." after every individual change.
- `.venv/bin/python -m pytest tests/ -v` exits 0 with all tests passing
  after requirements 8 and 9 are complete.
- `templates/src/` no longer exists.
- No `__main__` block remains in either `excel_utils` file.
- All public V2 functions have both type annotations and docstrings.

## Loop Instructions

Work through one task at a time using the task list in
`.ralph/ralph-tasks.md`. After completing a task and confirming its
verify command exits 0, output:
<promise>READY_FOR_NEXT_TASK</promise>

When all tasks are complete and all acceptance criteria above are met,
output:
<promise>COMPLETE</promise>
