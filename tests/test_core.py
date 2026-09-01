import unittest

from playwright_writer import ValueComparator
from transaction_extractor import amount_to_wan_text, format_date
from utils import extract_portfolio_from_text, to_number


class CoreTests(unittest.TestCase):
    def test_to_number(self):
        self.assertEqual(to_number("1,234.50"), 1234.5)
        self.assertEqual(to_number(None), 0.0)

    def test_portfolio_detection(self):
        self.assertEqual(extract_portfolio_from_text("valuation_Portfolio_B_20260630.xlsx"), "Portfolio_B")

    def test_date_and_amount_format(self):
        self.assertEqual(format_date("2026-06-30"), "20260630")
        self.assertEqual(amount_to_wan_text(20_000_000), "2000万")

    def test_value_comparator(self):
        self.assertTrue(ValueComparator.equals("14,887", 14887.0))
        self.assertTrue(ValueComparator.equals("2026/6/30", "2026/06/30"))
        self.assertFalse(ValueComparator.equals("10", "11"))


if __name__ == "__main__":
    unittest.main()
