# Excel Budget Workbook Generator

**Author:** Chris Wokral · **License:** MIT

Python script (`openpyxl`) that builds a personal budget workbook: transaction register, account balances, budget buckets, and a monthly summary driven by `SUMIFS` formulas. Transfers and card payments are labeled so they do not inflate spending.

All transactions are synthetic example data (fictional employers and merchants).

## What it does

| Sheet | Purpose |
|-------|---------|
| Transactions | Dated register with income, expense, transfer, and card-payment kinds |
| Accounts | Opening balances and current balances linked to the register |
| Budget | Monthly spend by category (expenses only) |
| Summary | Per-month income and spending; net-worth reconciliation |

Formula cells on protected sheets are locked; editable inputs (amounts, opening balances, targets) stay unlocked.

## Quick start

```bash
python3 -m pip install -r requirements.txt
python3 generate_workbook.py
python3 check_workbook.py
python3 -m unittest discover -s tests -v
```

Optional: if LibreOffice is installed, `check_workbook.py` runs a headless conversion to confirm the file opens for recalculation.

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
- Budget bucket formulas only sum rows where `CountsAsSpending` is `Y` (expense kind only).
