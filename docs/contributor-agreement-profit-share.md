# CS Reporter - Contributor Agreement (Profit Share)

**Between:** Christopher Morrison and Ben Tran

**Date:** ___________

## Project

CS Reporter - an automated Excel-to-PowerPoint report generator

## What Ben Will Do

Complete Version 1: generalize cs-reporter from hardcoded suffix-based operations to explicit config-driven operations, as defined in `.claude/plans/generalize-to-config-driven-operations.md`.

This includes:

1. Add operation validation to `src/config.py` (OPERATIONS registry, validate_fields)
2. Make filters configurable in `src/excel_utils.py` (remove hardcoded values)
3. Refactor `src/excel_reader.py` (replace suffix detection with operation dispatch)
4. Convert `config/mapping.yaml` and `config/demo_mapping.yaml` to the new explicit operation format

## Profit Share

If CS Reporter generates revenue, Ben will receive 15% of net profits.

"Net profits" = Revenue minus operating costs (hosting, API fees, transaction fees, distribution costs)

This applies to:

- The CS Reporter application as currently designed
- Any direct commercial version of it (e.g., a paid desktop app, SaaS product)

This does NOT apply to:

- Unrelated projects Chris builds
- Projects that merely use similar technology

## To Qualify

Ben must complete Version 1 as described above, with code that remains in use (not rewritten by Chris) for 30 days.

## Payment Timing

If/when revenue exists, Chris will share profit numbers quarterly and pay Ben's share via [PayPal / Wise / etc.].

## Duration

This profit share lasts for 2 years from the date the app first generates revenue, or until $50,000 total has been paid out, whichever comes first. After that, Ben's share ends.

## Ownership

All code written for this project belongs to Chris. Ben can reference the work in a portfolio or resume.

## Timeline

No strict deadline. Work at your own pace. If either of us wants to stop, just say so - no hard feelings.

---

| Signature | Date |
|-----------|------|
| Christopher Morrison: ___________________ | _______ |
| Ben Tran: ___________________ | _______ |
