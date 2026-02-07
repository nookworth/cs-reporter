# CS Reporter - Collaborator Guide

**Getting started as a developer on this project**

This guide walks you through running CS Reporter using demo data. The production template and Excel files are proprietary and not included in the repo.

---

## Prerequisites

- **Git**
- **Python 3.8+** ([Download](https://www.python.org/downloads/))
  - **Windows users:** Check "Add Python to PATH" during installation!

---

## Setup

### 1. Clone the repo

```cmd
git clone git@github.com:nookworth/cs-reporter.git
cd cs-reporter
```

### 2. Run the setup script

**Windows:**
```cmd
setup.bat
```

**macOS / Linux:**
```bash
chmod +x setup.sh
./setup.sh
```

This creates a virtual environment, installs dependencies, and sets up the `reporter` command.

### 3. Generate demo data

The demo Excel files are not checked into the repo -- you generate them locally.

**Windows:**
```cmd
.venv\Scripts\activate.bat
python scripts/generate_demo_data.py
```

**macOS / Linux:**
```bash
source .venv/bin/activate
python scripts/generate_demo_data.py
```

This creates two files in `demo_data/`:
- `demo_january_2026.xlsx` (current month)
- `demo_december_2025.xlsx` (previous month)

### 4. Run the reporter with demo config

**Windows:**
```cmd
.venv\Scripts\activate.bat
reporter --config config/demo_mapping.yaml
```

**macOS / Linux:**
```bash
source .venv/bin/activate
reporter --config config/demo_mapping.yaml
```

Select the two demo Excel files when prompted:
1. **Current month:** `demo_data/demo_january_2026.xlsx`
2. **Previous month:** `demo_data/demo_december_2025.xlsx`

Your report will appear in `output/`.

---

## What's in the repo vs. what's not

| File | In repo? | Notes |
|------|----------|-------|
| `templates/demo_template.pptx` | Yes | Generic demo template |
| `config/demo_mapping.yaml` | Yes | Config for demo data |
| `scripts/generate_demo_data.py` | Yes | Generates demo Excel files |
| `demo_data/*.xlsx` | No | Generated locally (gitignored) |
| `templates/report_template.pptx` | No | Proprietary (gitignored) |
| `config/mapping.yaml` | Yes | Production config -- won't work without the production template and matching Excel files |

---

## Project structure

```
cs-reporter/
├── src/
│   ├── main.py              # CLI entry point, file dialogs
│   ├── excel_reader.py      # Excel data extraction
│   ├── excel_utils.py       # Excel helper functions
│   ├── ppt_writer.py        # PowerPoint generation
│   ├── config.py            # Config loader
│   └── history.py           # History management (WIP, not active)
├── config/
│   ├── mapping.yaml         # Production config
│   └── demo_mapping.yaml    # Demo config
├── templates/
│   └── demo_template.pptx   # Demo PowerPoint template
├── scripts/
│   └── generate_demo_data.py
├── demo_data/               # Generated demo Excel files (gitignored)
├── output/                  # Generated reports (gitignored)
├── setup.bat / setup.sh     # First-time setup
├── run-reporter.bat         # Production launcher (no --config flag)
├── USER_GUIDE.md            # End-user documentation (production workflow)
└── COLLABORATORS.md          # This file
```

---

## How it works

1. User selects two Excel files (current + previous month)
2. `excel_reader.py` extracts scalar fields and table data from both files
3. Fields with `_prev_` in the name are read from the previous month file
4. `ppt_writer.py` replaces `{{placeholders}}` in the template with extracted values
5. Dynamic tables (`{{table:name}}`) are populated with aggregated data
6. Output is saved to `output/report_YYYYMMDD_HHMMSS.pptx`

### Key conventions

- **Field suffixes drive behavior:** `_req` counts rows, `_sat` counts "good" ratings, `_sat_c` counts "good with comment", `_reso` calculates average resolution time
- **`--config` flag:** Use `reporter --config path/to/config.yaml` to specify a config file. Defaults to `config/mapping.yaml` if omitted.
- **Template placeholders:** `{{field_name}}` for scalars, `{{table:table_name}}` for dynamic tables

---

## Production vs. demo

The demo setup mirrors the production workflow exactly -- same code paths, same config structure, same template format. The only differences are:

- **Template:** `demo_template.pptx` (generic) vs. `report_template.pptx` (proprietary)
- **Config:** `demo_mapping.yaml` (demo sheet names) vs. `mapping.yaml` (production sheet names)
- **Data:** Generated dummy data vs. real Excel exports

If you need to test a code change, the demo path exercises all the same logic.

---

## Roadmap: Making the reporter general-purpose

### The problem

Right now, cs-reporter is **one specific report** disguised as a general-purpose tool. The field-processing logic is hardcoded into the Python code based on naming conventions:

- Field name ends in `_req`? Count rows.
- Ends in `_sat`? Count cells matching `"good"`.
- Ends in `_reso`? Calculate average date difference, filter to 3 days.

The config file controls *which* columns and sheets to read, but *what to do with them* is baked into `excel_reader.py`. A second customer with different metrics -- say, they want to sum a dollar column, or count cells matching `"resolved"` instead of `"good"` -- would need code changes, not just config changes.

Think of it like a form vs. a form builder. We built **a form**. The next step is building **the form builder** -- a tool where users define their own fields, operations, and mappings entirely through configuration.

### What "general-purpose" looks like

Today's config says *what* to read:

```yaml
re_sat:
  cell: "Ticket satisfaction rating"
```

...and the code *infers* the operation from the field name (`_sat` → count "good").

A general-purpose config would say *what to read* **and** *what to do with it*:

```yaml
re_sat:
  column: "Ticket satisfaction rating"
  operation: count_matching
  match_value: "good"
  case_sensitive: false
```

The code wouldn't need to know anything about satisfaction ratings or support tickets. It would just execute whatever operation the config describes.

### What needs to change

1. **Define a set of config-driven operations** (count rows, count matching values, average date difference, read single value, etc.)
2. **Move the business logic out of `excel_reader.py`** and into the config -- field names become arbitrary labels, not behavioral hints
3. **Same for tables** -- aggregation type, grouping, filtering, and limits are already mostly config-driven, but the column-reading logic still has hardcoded assumptions
4. **Validation** -- if users are writing their own configs, the tool needs clear error messages when a config is malformed
