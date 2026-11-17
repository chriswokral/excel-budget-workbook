# Excel Budget Workbook Generator

**Author:** Chris Wokral · **License:** MIT

Python script (`openpyxl`) that builds a personal budget workbook: transaction register, account balances, budget buckets, and a monthly summary driven by `SUMIFS` formulas. Transfers and card payments are labeled so they do not inflate spending.

All transactions are synthetic example data (fictional employers and merchants).

## What it does

| Sheet | Purpose |
|-------|---------|
| Transactions | Dated register with income, expense, transfer, and card-payment kinds |
| Accounts | Opening balances and current balances linked to the register |
| Budget | Monthly spend by category (expenses only), plus a quarter budget-vs-actual report |
| Summary | Per-month income and spending; net-worth reconciliation |

Formula cells on protected sheets are locked; editable inputs (amounts, opening balances, targets) stay unlocked.

## Quick start

```bash
python3 -m pip install -r requirements.txt
python3 generate_workbook.py
python3 check_workbook.py
python3 -m unittest discover -s tests -v
```

Optional: if LibreOffice is installed, `check_workbook.py` runs a headless conversion to confirm the file opens for recalculation, and `tests/test_budget_vs_actual.py` recalculates the real formulas and compares them with the Python mirror (those two tests are skipped without LibreOffice).

## Budget vs actual (quarter)

The Budget sheet compares each category's spend with its target for the quarter:

| Column | Formula idea |
|--------|--------------|
| Quarter spend | Jan + Feb + Mar spend |
| Quarter target | Monthly target x 3 |
| Variance | Target - spend (negative = overspent) |
| % of target used | Spend / target |
| Status | Over, On target, or Under |

Month cells above their monthly target turn red, and the Status column is red for Over and green for Under. A Total row sums every column, and an **Unbudgeted spend** check (total expenses minus the Total row) flags any expense whose category has no budget line. Targets are editable, so a what-if is just typing a new number.

Sample results, hand-checked in `tests/test_budget_vs_actual.py`:

| Category | Monthly target | Jan | Feb | Mar | Quarter spend | Quarter target | Variance | Status |
|----------|---------------:|----:|----:|----:|--------------:|---------------:|---------:|--------|
| Rent | 1100 | 1100.00 | 1100.00 | 1100.00 | 3300.00 | 3300.00 | 0.00 | On target |
| Groceries | 350 | 124.35 | 229.66 | 0.00 | 354.01 | 1050.00 | 695.99 | Under |
| Dining | 55 | **58.20** | **72.15** | 41.30 | 171.65 | 165.00 | -6.65 | Over |
| Transport | 125 | 62.50 | 0.00 | 62.50 | 125.00 | 375.00 | 250.00 | Under |
| Utilities | 130 | **142.88** | 0.00 | 118.60 | 261.48 | 390.00 | 128.52 | Under |
| Healthcare | 75 | 0.00 | 45.00 | 0.00 | 45.00 | 225.00 | 180.00 | Under |
| Shopping | 100 | 0.00 | 0.00 | 89.99 | 89.99 | 300.00 | 210.01 | Under |
| **Total** | 1935 | | | | **4347.13** | **5805.00** | **1457.87** | Under |

Bold month cells are the ones highlighted red. Utilities shows why both views matter: January ran over its monthly target, but the quarter is still under. Quarter spend ties to the Summary sheet's total spending (4347.13), so the Unbudgeted spend check is 0.00.

## Check results (Python reconciliation)

| Check | Result |
|-------|--------|
| Transfer net (in + out + card payments on both sides) | 0.00 |
| Total income (3 synthetic months) | 8992.85 |
| Total spending (expenses only, card + checking) | 4347.13 |
| Income − spending = change in net worth | 4645.72 |
| Closing net worth (checking + savings + card) | 11865.72 |

Card payments appear on Checking and Credit Card but are excluded from the `CountsAsSpending` flag, so paying the card does not double-count as spending.

## How correctness is checked

- `tests/test_reconcile.py` regenerates the workbook and asserts transfer net zero and the income/spending/net-worth identity using the same transaction list as the generator.
- `check_workbook.py` prints JSON with expected totals and, when available, cached formula results after LibreOffice headless conversion.
- `tests/test_budget_vs_actual.py` checks quarter spend, variance, status and month-over flags against hand-worked sums, runs a what-if with a lower Groceries target, and (with LibreOffice) recalculates the workbook to confirm the sheet formulas give the same numbers.
- Budget bucket formulas only sum rows where `CountsAsSpending` is `Y` (expense kind only).
