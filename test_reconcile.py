#!/usr/bin/env python3
import json
import sys
import unittest
from decimal import Decimal
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from check_workbook import money, python_reconcile, run_checks  # noqa: E402
from generate_workbook import generate  # noqa: E402


class ReconcileTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.xlsx = ROOT / "output" / "personal_budget.xlsx"
        generate(cls.xlsx)

    def test_transfer_net_is_zero(self):
        r = python_reconcile()
        self.assertEqual(r["transfer_net"], Decimal("0.00"))

    def test_income_minus_spending_equals_net_worth_change(self):
        r = python_reconcile()
        self.assertEqual(r["income_minus_spending"], r["change_in_net_worth"])

    def test_expected_totals_documented(self):
        r = python_reconcile()
        self.assertEqual(r["income"], Decimal("8992.85"))
        self.assertEqual(r["spending"], Decimal("4347.13"))
        self.assertEqual(r["income_minus_spending"], Decimal("4645.72"))
        self.assertEqual(r["closing_net_worth"], Decimal("11865.72"))

    def test_check_script_passes(self):
        result = run_checks(self.xlsx)
        self.assertTrue(result["transfer_net_zero"])
        self.assertTrue(result["reconcile_identity"])


if __name__ == "__main__":
    unittest.main()
