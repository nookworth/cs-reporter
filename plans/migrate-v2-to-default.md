# Plan: Make V2 the Default and Retire V1 (Ralph-loop format)

## Goal

Flip the installed `reporter` command from the V1 stack
(`src/main.py` → `config.py` → `excel_reader.py` → `excel_utils.py`) to the
V2 stack (`src/main_v2.py` → `config_v2.py` → `excel_reader_v2.py` →
`excel_utils_v2.py`), update everything that references V1, then delete the
V1-only files — **without ever breaking `import` for the V2/V3 code or the
Streamlit web UI**.

## Why this is non-trivial (read before touching anything)

- The installed CLI entry point (`setup.py` → `src.main:cli_entry_point`)
  currently runs **V1**. Many docs/scripts assume this.
- `src/ppt_writer.py` has **no `_v2` sibling**. It is the base class of
  `PowerPointWriterV3` *and* is imported directly by `main_v2.py`. **It is
  shared, load-bearing code — never delete it.**
- `main_v2.py` uses the plain `PowerPointWriter` (no charts). Charts only
  exist on the V3 path (`ppt_writer_v3.PowerPointWriterV3`), today reachable
  only via the web UI (`app_v3.py`). Making V2 the CLI default does **not**
  give CLI users charts unless we explicitly wire V3 (see Step 4).
- `main_v2.py`'s `--config` default is `None`, and its help text still says
  `config/mapping.yaml` — a **V1-format** file. V2 needs a **V2-format**
  config. The default must be corrected or the default `reporter` invocation
  will load an incompatible config.

---

## Ralph-loop protocol (do this every iteration)

1. Read this file top to bottom. Find the **first unchecked `[ ]` task** in
   the Steps list. Do **only that one task**. Do not skip ahead.
2. Make the change described. Touch only the files that task names.
3. Run the **regression gate** (below). It must pass before you check the
   task off.
4. Check the box (`[x]`) and append a one-line `RESULT:` note under the task
   (what you did + gate status).
5. Commit with message `migrate-v2: <task title>`.
6. Stop. The loop will re-invoke you for the next task.

If the gate fails and you cannot fix it within the task's scope, **revert
your change**, leave the box unchecked, and append a `BLOCKED:` note
explaining why. Then stop.

### Regression gate (must stay green after every task)

```bash
# 1. V2/V3 + web UI imports must not break
python -c "import src.main_v2, src.excel_reader_v2, src.excel_utils_v2, src.config_v2, src.ppt_writer, src.ppt_writer_v3, src.chart_utils_v3, app_v3"
# 2. Full V2 pipeline regression against scrubbed demo data
python -m pytest tests/ -q
# 3. Lint/format clean (pre-commit hooks enforce this anyway)
ruff check . && ruff format --check .
```

A task that deletes a V1 file additionally requires:
`python -c "import src.main_v2"` to still succeed (proves the deletion did
not orphan a V2 import).

### Global guardrails

- **Never** delete or rename `src/ppt_writer.py` without first repointing the
  V2/V3 imports — it is the shared base writer.
- **Never** edit V2/V3 logic to make V1 deletion easier; V2 behavior is the
  source of truth.
- Keep every commit independently green. No "fix in the next commit."

---

## Steps

### Phase 0 — Safety net

- [x] **0.1 Branch + baseline gate.** Create branch `migrate/v2-default`.
  Run the full regression gate and record the current pass/fail state.
  Confirm `tests/test_smoke.py` passes on `main` before any change.
  RESULT: Branch created from `refactor/retire-v1`. Gate: V2/V3 imports OK, test_smoke passes. Pre-existing: 1 test_filters failure (case sensitivity), 36 ruff formatting issues. Baseline recorded.

- [x] **0.2 Capture a V1 reference report.** Run the existing V1 CLI on the
  demo data (`python -m src.main --config config/demo_mapping.yaml` or via
  the smoke fixtures) and save the generated `.pptx` to
  `plans/_migration_artifacts/v1_reference.pptx` (gitignored). This is the
  human comparison target for parity in Step 1.
  RESULT: V1 reference report generated (72 replacements, 3 tables of 8/8/5 rows) -> plans/_migration_artifacts/v1_reference.pptx. Added artifacts dir to .gitignore.

### Phase 1 — Close the V1↔V2 gap before flipping anything

- [x] **1.1 Pin the canonical V2 production config.** Decide which V2-format
  file is the prod default (recommended: `config/mapping_v2.yaml`). Confirm
  it loads via `config_v2.load_config` and that `template_path` resolves to a
  real template (`templates/report_template.pptx` per cleanup-spec item 3).
  Record the chosen path here; later steps reference it as `<PROD_CONFIG>`.
  RESULT: `<PROD_CONFIG>` = `config/mapping_v2.yaml`. Loads fine: template_path→templates/report_template.pptx (exists, 233KB), 18 fields, 3 table_fields.

- [ ] **1.2 Verify `config_v2.load_config(None)` default.** Read
  `src/config_v2.py`. Confirm what path it loads when `config_path=None`. If
  it is not `<PROD_CONFIG>` (V2 format), note the discrepancy — Step 3.2
  fixes it. Do not change logic here; this task only documents the behavior.
  RESULT:

- [ ] **1.3 Output parity check, V1 vs V2.** Run V2 on the same demo data V1
  used in 0.2 (`python -m src.main_v2 --config <V2 demo config>`). Diff the
  two reports: every placeholder V1 fills and every dynamic table V1 builds
  (`re_sup_cat`, `su_sup_cat`, `top_orgs`, all `*_req/_sat/_sat_c/_reso`
  scalars) must be present and equal in the V2 output. List any gaps as
  sub-bullets under this task.
  RESULT:

- [ ] **1.4 Fix any parity gaps found in 1.3.** For each gap, fix it in the
  **V2** modules (`excel_reader_v2.py` / `excel_utils_v2.py` /
  `config_v2.py` / the V2 config). Re-run 1.3's diff until V2 ≥ V1 in
  coverage. If 1.3 found no gaps, check this off with `RESULT: no gaps`.
  RESULT:

### Phase 2 — Decide the writer (charts or not)

- [ ] **2.1 Chart decision for the CLI.** The CLI currently produces no
  charts. Choose one and record it:
  - **(A, recommended)** Wire the CLI to `PowerPointWriterV3` so CLI reports
    gain the same charts as the web UI. Requires the V2 config to carry the
    chart config the V3 writer expects.
  - **(B)** Keep the base `PowerPointWriter` (no charts) for the CLI; charts
    stay a web-UI-only feature.
  Write the choice and rationale here. Steps 4.x implement it.
  RESULT:

### Phase 3 — Flip the entry point

- [ ] **3.1 Repoint `setup.py`.** Change `console_scripts` from
  `reporter=src.main:cli_entry_point` to
  `reporter=src.main_v2:cli_entry_point`. Do not change anything else in
  `setup.py`.
  RESULT:

- [ ] **3.2 Correct `main_v2.py` defaults + hygiene.** Set the `--config`
  default to `<PROD_CONFIG>` (or make `load_config(None)` resolve to it —
  whichever 1.2 indicated). Fix the stale help text (`default:
  config/mapping.yaml` → `<PROD_CONFIG>`). Strip the trailing-whitespace on
  the import lines (`config_v2 `, `excel_reader_v2 `, `ppt_writer `). Do not
  change the pipeline logic.
  RESULT:

- [ ] **3.3 Implement the Step 2.1 writer choice in `main_v2.py`.** If (A):
  import and instantiate `PowerPointWriterV3` instead of `PowerPointWriter`.
  If (B): no code change — check off with `RESULT: option B, no change`.
  RESULT:

- [ ] **3.4 Reinstall and smoke the real command.** Run `pip install -e .`,
  then `reporter --config <PROD_CONFIG>` (or pipe a demo config) and confirm
  it runs the V2 pipeline end to end and writes a `.pptx`. The bare
  `reporter` invocation must load a V2-format config, not a V1 one.
  RESULT:

### Phase 4 — Reconcile launchers and scripts

- [ ] **4.1 `setup.sh` launcher.** The heredoc that generates
  `run-reporter.sh` calls `reporter` — now V2, so it is correct, but verify
  it passes no V1-only `--config`. Update if it hardcodes a V1 config.
  RESULT:

- [ ] **4.2 `run-reporter.bat` / `run_v2.bat` / `run_web_v3.*`.** Audit each.
  `run_v2.py` (which injects `config/demo_mapping_v2.yaml`) is now redundant
  with the default `reporter` — either delete it or convert it to a thin
  alias and note that here. Keep the web-UI launchers (`run_web_v3.*`) since
  the Streamlit app is unaffected.
  RESULT:

### Phase 5 — Update documentation

- [ ] **5.1 `README.md`.** Replace `python -m src.main` / V1-default language
  with the V2 command. Remove the "V1 vs V2, run this for V1" instructions or
  recast them as "legacy / removed."
  RESULT:

- [ ] **5.2 `BEN_READ_THIS.md`.** The onboarding `reporter --config
  config/demo_mapping.yaml` line must point at a **V2** demo config. Update
  any V1 architecture description in the roadmap section.
  RESULT:

- [ ] **5.3 `USER_GUIDE.md` + `QUICKSTART.txt`.** Confirm the documented
  `reporter` flow matches V2 behavior (config format, prompts). Update file
  names/paths as needed.
  RESULT:

- [ ] **5.4 `.claude/CLAUDE.md`.** Update "Status" and "Project Structure" to
  state V2 is the default and V1 has been retired. Add a dated "Recent
  Changes" entry for this migration. Note that `ppt_writer.py` is the shared
  base writer (kept), not a V1 file.
  RESULT:

### Phase 6 — Retire V1 (one file per task; gate after each)

- [ ] **6.1 Delete `src/history.py`.** Confirmed imported by nothing. Delete,
  run gate.
  RESULT:

- [ ] **6.2 Delete `src/main.py`.** First `grep -rn "src.main\b\|src\.main:"
  --include=*.py --include=*.toml --include=*.sh --include=*.bat .` and
  confirm only docs (already updated in Phase 5) reference it. Delete, run
  gate.
  RESULT:

- [ ] **6.3 Delete `src/excel_reader.py`.** Confirm nothing but the
  now-deleted `main.py` imported it (`grep -rn "excel_reader\b"` excluding
  `_v2`). Delete, run gate.
  RESULT:

- [ ] **6.4 Delete `src/config.py`.** Confirm no non-deleted module imports
  `from .config import` / `src.config` (excluding `config_v2`). Delete, run
  gate.
  RESULT:

- [ ] **6.5 Delete `src/excel_utils.py`.** Confirm nothing but the deleted
  `excel_reader.py` imported it (excluding `_v2`). Delete, run gate.
  RESULT:

- [ ] **6.6 Confirm `src/ppt_writer.py` is still present and imported.** This
  is a verification task, not a deletion. `import src.ppt_writer` and
  `import src.ppt_writer_v3` must both succeed. If you want to drop the `_v2`
  naming churn, this is where you'd rename `ppt_writer.py` and update the two
  importers (`main_v2.py`, `ppt_writer_v3.py`) — optional, behind the gate.
  RESULT:

### Phase 7 — Final reconciliation

- [ ] **7.1 Update stale plan/test references.** `plans/cleanup-spec.md` and
  `plans/cleanup-tasks.md` describe V1 as "still the default." Add a note
  that this migration superseded that. Confirm `tests/` only reference V2
  modules (they already do) and still pass.
  RESULT:

- [ ] **7.2 Drop the `_v2`/`_v3` suffixes (optional, gated).** If desired now
  that V1 is gone, rename V2/V3 modules to unsuffixed names and update all
  importers + `setup.py` + tests + web UI in a single commit. Skip if you
  prefer to keep the suffixes; check off with `RESULT: deferred`.
  RESULT:

- [ ] **7.3 Final full-gate + manual report.** Run the regression gate one
  last time, generate a report via the real `reporter` command on demo data,
  and eyeball it against the V1 reference from 0.2. Open the PR.
  RESULT:

---

## Done when

- `reporter` (installed entry point) runs the V2 pipeline by default.
- `src/main.py`, `src/config.py`, `src/excel_reader.py`,
  `src/excel_utils.py`, `src/history.py` are deleted.
- `src/ppt_writer.py` remains and `import src.ppt_writer_v3` works.
- `python -m pytest tests/ -q` and `app_v3` import are green.
- No doc or script tells a user to run `python -m src.main`.
