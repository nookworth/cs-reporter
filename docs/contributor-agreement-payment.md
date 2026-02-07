# CS Reporter - Contributor Agreement (Payment)

**Between:** Christopher Morrison and Ben Tran

**Date:** ___________

## Project

CS Reporter - an automated Excel-to-PowerPoint report generator

## What Ben Will Do

Complete Version 1: generalize cs-reporter from hardcoded suffix-based operations to explicit config-driven operations, as defined in `plans/generalize-to-config-driven-operations.md`.

This includes:

1. Add operation validation to `src/config.py` (OPERATIONS registry, validate_fields)
2. Make filters configurable in `src/excel_utils.py` (remove hardcoded values)
3. Refactor `src/excel_reader.py` (replace suffix detection with operation dispatch)
4. Convert `config/mapping.yaml` and `config/demo_mapping.yaml` to the new explicit operation format

## Payment

$50 USD, paid upon completion of Version 1.

Version 1 is "complete" when:

- The code works as described in the plan
- Chris has tested it and confirmed it works
- The code is merged to the main branch

Payment will be sent via [PayPal / Wise / Venmo / etc.] within 7 days of completion.

## Ownership

All code written for this project belongs to Chris. Ben can reference the work in a portfolio or resume.

## Timeline

No strict deadline. Work at your own pace. If either of us wants to stop, just say so - no hard feelings.

---

| Signature | Date |
|-----------|------|
| Christopher Morrison: ___________________ | _______ |
| Ben Tran: ___________________ | _______ |
