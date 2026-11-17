#!/usr/bin/env python3
"""Budget-vs-actual checks. Expected values are worked out by hand from the
synthetic transactions in generate_workbook.py (see README for the arithmetic)."""
import shutil
import subprocess
import sys
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from openpyxl import load_workbook  # noqa: E402

from check_workbook import budget_vs_actual  # noqa: E402
from generate_workbook import BUDGET_BUCKETS, generate  # noqa: E402

D = Decimal


def libreoffice_cmd():
    return shutil.which("libreoffice") or shutil.which("soffice")


def recalc_copy(src: Path, outdir: Path) -> Path:
    """Open the workbook headless in LibreOffice and save a recalculated copy."""
    cmd = libreoffice_cmd()
    subprocess.run(
        [cmd, "--headless", "--calc", "--convert-to", "xlsx", "--outdir", str(outdir), str(src)],
        check=True,
        capture_output=True,
        timeout=180,
    )
    return outdir / src.name


def cents(value) -> Decimal:
    return D(str(round(value, 2))).quantize(D("0.01"))


class BudgetVsActualPythonTests(unittest.TestCase):
    def test_hand_checked_quarter_spend(self):
        rows = budget_vs_actual()
        # Groceries: 124.35 (Jan) + 98.44 + 131.22 (Feb) = 354.01
        self.assertEqual(rows["Groceries"]["months"], [D("124.35"), D("229.66"), D("0.00")])
        self.assertEqual(rows["Groceries"]["spend"], D("354.01"))
        # Utilities: 142.88 (Jan) + 118.60 (Mar) = 261.48
        self.assertEqual(rows["Utilities"]["spend"], D("261.48"))
        # Rent: 3 x 1100 = 3300, exactly the quarter target
        self.assertEqual(rows["Rent"]["spend"], D("3300.00"))
        self.assertEqual(rows["Rent"]["status"], "On target")

    def test_dining_over_quarter_target(self):
        dining = budget_vs_actual()["Dining"]
        # 58.20 + 72.15 + 41.30 = 171.65 against 3 x 55 = 165.00
        self.assertEqual(dining["spend"], D("171.65"))
        self.assertEqual(dining["target"], D("165.00"))
        self.assertEqual(dining["variance"], D("-6.65"))
        self.assertEqual(dining["status"], "Over")
        self.assertEqual(dining["months_over"], [1, 2])

    def test_utilities_month_over_but_quarter_under(self):
        util = budget_vs_actual()["Utilities"]
        # Jan 142.88 > 130 monthly target; quarter 261.48 < 390
        self.assertEqual(util["months_over"], [1])
        self.assertEqual(util["status"], "Under")
        self.assertEqual(util["variance"], D("128.52"))

    def test_totals_tie_to_total_spending(self):
        rows = budget_vs_actual()
        spend = sum((r["spend"] for r in rows.values()), D("0"))
        target = sum((r["target"] for r in rows.values()), D("0"))
        self.assertEqual(spend, D("4347.13"))  # same as Summary spending: nothing unbudgeted
        self.assertEqual(target, D("5805.00"))  # 3 x 1935
        self.assertEqual(target - spend, D("1457.87"))

    def test_what_if_targets(self):
        custom = [(c, 100 if c == "Groceries" else t) for c, t in BUDGET_BUCKETS]
        g = budget_vs_actual(custom)["Groceries"]
        # Jan 124.35 and Feb 229.66 > 100; quarter 354.01 > 300
        self.assertEqual(g["months_over"], [1, 2])
        self.assertEqual(g["status"], "Over")
        self.assertEqual(g["variance"], D("-54.01"))


@unittest.skipUnless(libreoffice_cmd(), "LibreOffice not installed")
class BudgetSheetFormulaTests(unittest.TestCase):
    """Recalculate the real workbook formulas and compare them with the Python mirror."""

    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        src = Path(cls.tmp.name) / "src" / "personal_budget.xlsx"
        generate(src)
        out = Path(cls.tmp.name) / "recalc"
        out.mkdir()
        cls.ws = load_workbook(recalc_copy(src, out), data_only=True)["Budget"]

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_rows_match_python(self):
        expected = budget_vs_actual()
        for r in range(2, 2 + len(BUDGET_BUCKETS)):
            cat = self.ws.cell(r, 1).value
            exp = expected[cat]
            with self.subTest(category=cat):
                self.assertEqual(cents(self.ws.cell(r, 6).value), exp["spend"])
                self.assertEqual(cents(self.ws.cell(r, 7).value), exp["target"])
                self.assertEqual(cents(self.ws.cell(r, 8).value), exp["variance"])
                self.assertEqual(self.ws.cell(r, 10).value, exp["status"])

    def test_total_row_and_unbudgeted_check(self):
        total_row = 2 + len(BUDGET_BUCKETS)
        self.assertEqual(self.ws.cell(total_row, 1).value, "Total")
        self.assertEqual(cents(self.ws.cell(total_row, 6).value), D("4347.13"))
        self.assertEqual(cents(self.ws.cell(total_row, 7).value), D("5805.00"))
        self.assertEqual(cents(self.ws.cell(total_row, 8).value), D("1457.87"))
        self.assertEqual(self.ws.cell(total_row, 10).value, "Under")
        self.assertEqual(cents(self.ws.cell(total_row + 2, 6).value), D("0.00"))


if __name__ == "__main__":
    unittest.main()
