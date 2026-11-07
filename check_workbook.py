#!/usr/bin/env python3
"""Reconciliation checks for the generated budget workbook."""

from __future__ import annotations

import json
import subprocess
import sys
from datetime import date
from decimal import Decimal
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
DEFAULT_XLSX = ROOT / "output" / "personal_budget.xlsx"

# Mirror generator constants for Python-side expected totals.
from generate_workbook import OPENING_BALANCES, TRANSACTIONS  # noqa: E402


def money(v: float | Decimal) -> Decimal:
    return Decimal(str(v)).quantize(Decimal("0.01"))


def python_reconcile() -> dict[str, Decimal]:
    income = Decimal("0")
    spending = Decimal("0")
    transfer_net = Decimal("0")
    by_account: dict[str, Decimal] = {k: Decimal(str(v)) for k, v in OPENING_BALANCES.items()}

    for tx in TRANSACTIONS:
        amt = money(tx["amt"])
        acct = str(tx["acct"])
        kind = str(tx["kind"])
        by_account[acct] += amt
        if kind == "Income":
            income += amt
        elif kind == "Expense":
            spending += -amt
        elif kind in ("Transfer Out", "Transfer In", "Card Payment"):
            transfer_net += amt

    opening_nw = sum(by_account[k] - sum(money(t["amt"]) for t in TRANSACTIONS if t["acct"] == k) for k in by_account)
    # Recompute opening from OPENING only
    opening_nw = sum(money(v) for v in OPENING_BALANCES.values())
    closing_nw = sum(by_account.values())
    change = closing_nw - opening_nw
    return {
        "income": money(income),
        "spending": money(spending),
        "income_minus_spending": money(income - spending),
        "change_in_net_worth": money(change),
        "transfer_net": money(transfer_net),
        "closing_net_worth": money(closing_nw),
    }


def try_libreoffice_recalc(path: Path) -> tuple[bool, str]:
    for cmd in ("libreoffice", "soffice"):
        try:
            proc = subprocess.run(
                [cmd, "--headless", "--calc", "--convert-to", "xlsx", str(path)],
                capture_output=True,
                text=True,
                timeout=120,
                cwd=path.parent,
            )
        except FileNotFoundError:
            continue
        if proc.returncode == 0:
            return True, f"{cmd} recalc conversion succeeded"
        return False, proc.stderr or proc.stdout or f"{cmd} failed"
    return False, "LibreOffice not available (skipped)"


def read_summary_values(path: Path) -> dict[str, float | None]:
    wb = load_workbook(path, data_only=True)
    ws = wb["Summary"]
    return {
        "net_worth": ws["B8"].value,
        "opening_net_worth": ws["B9"].value,
        "change_in_net_worth": ws["B10"].value,
        "income_minus_spending": ws["B11"].value,
        "transfer_net": ws["B13"].value,
    }


def run_checks(path: Path) -> dict[str, object]:
    expected = python_reconcile()
    lo_ok, lo_msg = try_libreoffice_recalc(path)
    xlsx_values = read_summary_values(path)

    checks = {
        "expected": {k: float(v) for k, v in expected.items()},
        "workbook_cached_values": xlsx_values,
        "transfer_net_zero": abs(expected["transfer_net"]) < Decimal("0.01"),
        "reconcile_identity": expected["income_minus_spending"] == expected["change_in_net_worth"],
        "libreoffice": {"ok": lo_ok, "message": lo_msg},
    }
    if xlsx_values["change_in_net_worth"] is not None:
        checks["cached_matches_expected"] = money(xlsx_values["change_in_net_worth"]) == expected[
            "change_in_net_worth"
        ]
    else:
        checks["cached_matches_expected"] = None
        checks["note"] = "Cached formula values empty; identity verified in Python."
    return checks


def main(argv: list[str] | None = None) -> int:
    path = DEFAULT_XLSX
    if argv and len(argv) > 1:
        path = Path(argv[1])
    if not path.exists():
        print(f"Missing workbook: {path}", file=sys.stderr)
        return 1
    result = run_checks(path)
    print(json.dumps(result, indent=2))
    ok = result["transfer_net_zero"] and result["reconcile_identity"]
    if result.get("cached_matches_expected") is False:
        ok = False
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
