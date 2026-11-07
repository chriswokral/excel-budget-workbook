#!/usr/bin/env python3
"""Generate a personal budget workbook with formulas and sample transactions."""

from __future__ import annotations

import argparse
from datetime import date
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Protection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.protection import SheetProtection

OUTPUT_DEFAULT = Path(__file__).resolve().parent / "output" / "personal_budget.xlsx"

# Synthetic example data only — fictional payers and merchants.
TRANSACTIONS: list[dict[str, object]] = [
    # January
    {"d": date(2025, 1, 3), "acct": "Checking", "kind": "Income", "cat": "Salary", "amt": 2850.00, "memo": "Acme Studio payroll"},
    {"d": date(2025, 1, 5), "acct": "Checking", "kind": "Expense", "cat": "Groceries", "amt": -124.35, "memo": "Harbor Market"},
    {"d": date(2025, 1, 7), "acct": "Credit Card", "kind": "Expense", "cat": "Dining", "amt": -58.20, "memo": "Riverside Café"},
    {"d": date(2025, 1, 10), "acct": "Checking", "kind": "Transfer Out", "cat": "Savings", "amt": -400.00, "memo": "To savings"},
    {"d": date(2025, 1, 10), "acct": "Savings", "kind": "Transfer In", "cat": "Checking", "amt": 400.00, "memo": "From checking"},
    {"d": date(2025, 1, 12), "acct": "Credit Card", "kind": "Expense", "cat": "Utilities", "amt": -142.88, "memo": "Grid power"},
    {"d": date(2025, 1, 15), "acct": "Checking", "kind": "Expense", "cat": "Rent", "amt": -1100.00, "memo": "Maple Court lease"},
    {"d": date(2025, 1, 18), "acct": "Checking", "kind": "Card Payment", "cat": "Credit Card", "amt": -175.00, "memo": "Card payoff"},
    {"d": date(2025, 1, 18), "acct": "Credit Card", "kind": "Card Payment", "cat": "Checking", "amt": 175.00, "memo": "Payment received"},
    {"d": date(2025, 1, 22), "acct": "Checking", "kind": "Expense", "cat": "Transport", "amt": -62.50, "memo": "Transit pass"},
    {"d": date(2025, 1, 28), "acct": "Savings", "kind": "Income", "cat": "Interest", "amt": 8.75, "memo": "Savings yield"},
    # February
    {"d": date(2025, 2, 4), "acct": "Checking", "kind": "Income", "cat": "Salary", "amt": 2850.00, "memo": "Acme Studio payroll"},
    {"d": date(2025, 2, 6), "acct": "Credit Card", "kind": "Expense", "cat": "Groceries", "amt": -98.44, "memo": "Harbor Market"},
    {"d": date(2025, 2, 9), "acct": "Checking", "kind": "Expense", "cat": "Healthcare", "amt": -45.00, "memo": "Clinic copay"},
    {"d": date(2025, 2, 12), "acct": "Checking", "kind": "Transfer Out", "cat": "Savings", "amt": -350.00, "memo": "To savings"},
    {"d": date(2025, 2, 12), "acct": "Savings", "kind": "Transfer In", "cat": "Checking", "amt": 350.00, "memo": "From checking"},
    {"d": date(2025, 2, 15), "acct": "Checking", "kind": "Expense", "cat": "Rent", "amt": -1100.00, "memo": "Maple Court lease"},
    {"d": date(2025, 2, 18), "acct": "Credit Card", "kind": "Expense", "cat": "Dining", "amt": -72.15, "memo": "Riverside Café"},
    {"d": date(2025, 2, 20), "acct": "Checking", "kind": "Card Payment", "cat": "Credit Card", "amt": -200.08, "memo": "Card payoff"},
    {"d": date(2025, 2, 20), "acct": "Credit Card", "kind": "Card Payment", "cat": "Checking", "amt": 200.08, "memo": "Payment received"},
    {"d": date(2025, 2, 25), "acct": "Checking", "kind": "Expense", "cat": "Groceries", "amt": -131.22, "memo": "Harbor Market"},
    # March
    {"d": date(2025, 3, 3), "acct": "Checking", "kind": "Income", "cat": "Salary", "amt": 2850.00, "memo": "Acme Studio payroll"},
    {"d": date(2025, 3, 5), "acct": "Checking", "kind": "Income", "cat": "Freelance", "amt": 425.00, "memo": "Vector Audio mix"},
    {"d": date(2025, 3, 8), "acct": "Credit Card", "kind": "Expense", "cat": "Shopping", "amt": -89.99, "memo": "Northwind Books"},
    {"d": date(2025, 3, 11), "acct": "Checking", "kind": "Expense", "cat": "Utilities", "amt": -118.60, "memo": "Grid power"},
    {"d": date(2025, 3, 14), "acct": "Checking", "kind": "Transfer Out", "cat": "Savings", "amt": -500.00, "memo": "To savings"},
    {"d": date(2025, 3, 14), "acct": "Savings", "kind": "Transfer In", "cat": "Checking", "amt": 500.00, "memo": "From checking"},
    {"d": date(2025, 3, 15), "acct": "Checking", "kind": "Expense", "cat": "Rent", "amt": -1100.00, "memo": "Maple Court lease"},
    {"d": date(2025, 3, 19), "acct": "Checking", "kind": "Expense", "cat": "Transport", "amt": -62.50, "memo": "Transit pass"},
    {"d": date(2025, 3, 22), "acct": "Credit Card", "kind": "Expense", "cat": "Dining", "amt": -41.30, "memo": "Riverside Café"},
    {"d": date(2025, 3, 28), "acct": "Checking", "kind": "Card Payment", "cat": "Credit Card", "amt": -131.29, "memo": "Card payoff"},
    {"d": date(2025, 3, 28), "acct": "Credit Card", "kind": "Card Payment", "cat": "Checking", "amt": 131.29, "memo": "Payment received"},
    {"d": date(2025, 3, 30), "acct": "Savings", "kind": "Income", "cat": "Interest", "amt": 9.10, "memo": "Savings yield"},
]

BUDGET_BUCKETS = [
    ("Rent", 1100),
    ("Groceries", 350),
    ("Dining", 150),
    ("Transport", 125),
    ("Utilities", 160),
    ("Healthcare", 75),
    ("Shopping", 100),
]

OPENING_BALANCES = {"Checking": 2400.00, "Savings": 5200.00, "Credit Card": -380.00}

HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
TITLE_FONT = Font(bold=True, size=12)
LOCKED = Protection(locked=True)
UNLOCKED = Protection(locked=False)


def protect_calc_column(ws, col: int, first_row: int, last_row: int) -> None:
    for r in range(first_row, last_row + 1):
        cell = ws.cell(row=r, column=col)
        cell.protection = LOCKED


def build_transactions(ws) -> int:
    headers = ["Date", "Account", "Kind", "Category", "Amount", "CountsAsSpending", "Memo"]
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        ws.cell(1, c).font = Font(bold=True)
        ws.cell(1, c).fill = HEADER_FILL
    kinds = '"Income,Expense,Transfer Out,Transfer In,Card Payment"'
    dv_acct = DataValidation(type="list", formula1='"Checking,Savings,Credit Card"', allow_blank=False)
    dv_kind = DataValidation(type="list", formula1=kinds, allow_blank=False)
    ws.add_data_validation(dv_acct)
    ws.add_data_validation(dv_kind)

    row = 2
    for tx in TRANSACTIONS:
        ws.cell(row, 1, tx["d"])
        ws.cell(row, 2, tx["acct"])
        ws.cell(row, 3, tx["kind"])
        ws.cell(row, 4, tx["cat"])
        ws.cell(row, 5, tx["amt"])
        ws.cell(
            row,
            6,
            f'=IF(D{row}="Expense",IF(OR(C{row}="Expense"),"Y","N"),"N")',
        )
        ws.cell(row, 7, tx["memo"])
        dv_acct.add(ws.cell(row, 2))
        dv_kind.add(ws.cell(row, 3))
        for c in range(1, 8):
            ws.cell(row, c).protection = UNLOCKED
        row += 1
    last = row - 1
    # Fix CountsAsSpending to use Kind only (cleaner formula)
    for r in range(2, last + 1):
        ws.cell(r, 6, f'=IF(C{r}="Expense","Y","N")')
        ws.cell(r, 6).protection = LOCKED
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["G"].width = 28
    return last


def build_accounts(ws, tx_last: int) -> None:
    ws["A1"] = "Account"
    ws["B1"] = "Opening balance"
    ws["C1"] = "Current balance"
    for c in range(1, 4):
        ws.cell(1, c).font = Font(bold=True)
        ws.cell(1, c).fill = HEADER_FILL
    accounts = ["Checking", "Savings", "Credit Card"]
    for i, name in enumerate(accounts, start=2):
        ws.cell(i, 1, name)
        ws.cell(i, 2, OPENING_BALANCES[name])
        ws.cell(i, 2).protection = UNLOCKED
        ws.cell(
            i,
            3,
            f"=B{i}+SUMIF(Transactions!$B$2:$B${tx_last},A{i},Transactions!$E$2:$E${tx_last})",
        )
        ws.cell(i, 3).protection = LOCKED
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 16


def build_budget(ws, tx_last: int) -> None:
    ws["A1"] = "Category"
    ws["B1"] = "Monthly target"
    ws["C1"] = "Jan spend"
    ws["D1"] = "Feb spend"
    ws["E1"] = "Mar spend"
    for c in range(1, 6):
        ws.cell(1, c).font = Font(bold=True)
        ws.cell(1, c).fill = HEADER_FILL
    for i, (cat, target) in enumerate(BUDGET_BUCKETS, start=2):
        ws.cell(i, 1, cat)
        ws.cell(i, 2, target)
        ws.cell(i, 2).protection = UNLOCKED
        for col, month in ((3, 1), (4, 2), (5, 3)):
            letter_amt = "E"
            formula = (
                f'=-SUMIFS(Transactions!${letter_amt}$2:${letter_amt}${tx_last},'
                f'Transactions!$D$2:$D${tx_last},$A{i},'
                f'Transactions!$F$2:$F${tx_last},"Y",'
                f'Transactions!$A$2:$A${tx_last},">="&DATE(2025,{month},1),'
                f'Transactions!$A$2:$A${tx_last},"<="&EOMONTH(DATE(2025,{month},1),0))'
            )
            ws.cell(i, col, formula)
            ws.cell(i, col).protection = LOCKED
    ws.column_dimensions["A"].width = 14


def month_formulas(ws, tx_last: int, row: int, month: int) -> None:
    ws.cell(row, 1, date(2025, month, 1))
    ws.cell(row, 1).number_format = "mmm yyyy"
    ws.cell(
        row,
        2,
        f'=SUMIFS(Transactions!$E$2:$E${tx_last},Transactions!$C$2:$C${tx_last},"Income",'
        f'Transactions!$A$2:$A${tx_last},">="&DATE(2025,{month},1),'
        f'Transactions!$A$2:$A${tx_last},"<="&EOMONTH(DATE(2025,{month},1),0))',
    )
    ws.cell(
        row,
        3,
        f'=-SUMIFS(Transactions!$E$2:$E${tx_last},Transactions!$F$2:$F${tx_last},"Y",'
        f'Transactions!$A$2:$A${tx_last},">="&DATE(2025,{month},1),'
        f'Transactions!$A$2:$A${tx_last},"<="&EOMONTH(DATE(2025,{month},1),0))',
    )
    ws.cell(
        row,
        4,
        f'=SUMIFS(Transactions!$E$2:$E${tx_last},Transactions!$C$2:$C${tx_last},"Transfer Out",'
        f'Transactions!$A$2:$A${tx_last},">="&DATE(2025,{month},1),'
        f'Transactions!$A$2:$A${tx_last},"<="&EOMONTH(DATE(2025,{month},1),0))',
    )
    for c in range(2, 5):
        ws.cell(row, c).protection = LOCKED


def build_summary(ws, tx_last: int) -> None:
    ws["A1"] = "Monthly summary (synthetic example)"
    ws["A1"].font = TITLE_FONT
    headers = ["Month", "Income", "Spending", "Transfers out"]
    for i, h in enumerate(headers, start=1):
        ws.cell(3, i, h)
        ws.cell(3, i).font = Font(bold=True)
        ws.cell(3, i).fill = HEADER_FILL
    for m, row in enumerate(range(4, 7), start=1):
        month_formulas(ws, tx_last, row, m)
    ws["A8"] = "Net worth (accounts)"
    ws["B8"] = "=SUM(Accounts!C2:C4)"
    ws["B8"].protection = LOCKED
    ws["A9"] = "Opening net worth"
    ws["B9"] = "=SUM(Accounts!B2:B4)"
    ws["A10"] = "Change in net worth"
    ws["B10"] = "=B8-B9"
    ws["A11"] = "Income minus spending (period)"
    ws["B11"] = "=SUM(B4:B6)-SUM(C4:C6)"  # spending stored as positive
    for r in (8, 9, 10, 11):
        ws.cell(r, 2).protection = LOCKED
    ws["A13"] = "Transfer net (should be 0)"
    ws["B13"] = f"=SUMIF(Transactions!$C$2:$C${tx_last},\"Transfer Out\",Transactions!$E$2:$E${tx_last})+SUMIF(Transactions!$C$2:$C${tx_last},\"Transfer In\",Transactions!$E$2:$E${tx_last})"
    ws["B13"].protection = LOCKED
    ws.column_dimensions["A"].width = 28


def apply_sheet_protection(wb: Workbook) -> None:
    for name in ("Transactions", "Accounts", "Budget", "Summary"):
        ws = wb[name]
        ws.protection = SheetProtection(sheet=True, password=None)
        ws.protection.formatCells = False
        ws.protection.formatColumns = False
        ws.protection.formatRows = False
        ws.protection.insertRows = False


def generate(path: Path) -> Path:
    wb = Workbook()
    tx_ws = wb.active
    tx_ws.title = "Transactions"
    tx_last = build_transactions(tx_ws)

    acct_ws = wb.create_sheet("Accounts")
    build_accounts(acct_ws, tx_last)

    budget_ws = wb.create_sheet("Budget")
    build_budget(budget_ws, tx_last)

    summary_ws = wb.create_sheet("Summary")
    build_summary(summary_ws, tx_last)

    apply_sheet_protection(wb)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate personal budget workbook.")
    parser.add_argument("-o", "--output", type=Path, default=OUTPUT_DEFAULT)
    args = parser.parse_args()
    out = generate(args.output)
    print(f"Wrote {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
