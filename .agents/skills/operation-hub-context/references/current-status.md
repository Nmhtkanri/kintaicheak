# Latest recorded project state

Status below was migrated from the curated Claude project memory on 2026-08-17. Verify it against the active repository before relying on it.

## Social-insurance check — active stream

- The standard-remuneration check's first half was committed on `feature/shaho-check` at `992aeb02`.
- `services/shaho_master.py`, `services/shaho_engine.py`, the classification master, configuration, and 44 tests were completed. The 2026-04 to 06 reconciliation gate matched all 713 person-months.
- Premium calculation, deduction reconciliation, decision engine, Excel/JSON output, and CLI remained unimplemented in the latest memory.
- Important classification rules: April 2026 is the final old pay-system month; hourly `a2` changed meaning over time; `a54` is paid income; and `a53` is in-kind remuneration through April but an excluded Amazon gift card from May onward.

## Completed or follow-up streams

- Expense-check decisions and deferred items were implemented; the remaining delivery action in the older note was rebuilding the executable.
- Health-check HPM CSV conversion and direct PDF reading were merged and packaged. Do not average blood-pressure values; for direct PDF reading, retain upper and lower page bands to prevent value swaps.
- Sharo CSV specifications were fixed and a 231-row check matched. Advances reduce gross pay and account 1, not net pay.
- Accounting rules include current-month employer social-insurance charges versus prior-month employee deductions. Child-support contribution 506 yen was still an open follow-up.
- Invoice-support expansion is planned in phases 0–4; open questions Q1–Q3 gate Phase 1.
- Commuting-cost monthly cap is 30,000 yen with seven exempt people maintained in a shared CSV. Confirm whether the cap applies to actual-cost reimbursements before changing it.
