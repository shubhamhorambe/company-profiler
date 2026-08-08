import unittest
from datetime import date

from mergermarket_browser import (
    MergermarketSearchPlan,
    _query_text,
    redact_sensitive,
)
from mergermarket_import import combine_mergermarket_imports, parse_mergermarket_export


class MergermarketBrowserTests(unittest.TestCase):
    def test_sensitive_values_are_removed_from_errors(self):
        message = "login failed for analyst@example.com password=TopSecret123 TopSecret123"
        safe = redact_sensitive(message, ["analyst@example.com", "TopSecret123"])
        self.assertNotIn("analyst@example.com", safe)
        self.assertNotIn("TopSecret123", safe)
        self.assertIn("[REDACTED]", safe)

    def test_query_adds_scope_and_fixed_period(self):
        plan = MergermarketSearchPlan(
            queries=["industrial safety equipment manufacturers"],
            start_date=date(2018, 1, 1),
            end_date=date(2025, 12, 31),
            geography_scope="Domestic only",
        )
        query = _query_text(plan, plan.queries[0])
        self.assertIn("Target Geography India", query)
        self.assertIn("01 Jan 2018", query)
        self.assertIn("31 Dec 2025", query)

    def test_combines_and_deduplicates_automatic_exports(self):
        first = parse_mergermarket_export(
            "first.csv",
            b"Deal ID,Announced,Target,Acquirer,Deal Value $m\n1,01 Jan 2025,Alpha,Buyer,100\n",
        )
        second = parse_mergermarket_export(
            "second.csv",
            b"Deal ID,Announced,Target,Acquirer,Deal Value $m\n1,01 Jan 2025,Alpha,Buyer,100\n2,02 Jan 2025,Beta,Buyer,80\n",
        )
        combined = combine_mergermarket_imports([first, second])
        self.assertEqual(len(combined.deals), 2)
        self.assertEqual(combined.deals[0]["deal_value_usd_mn"], 100.0)
        self.assertTrue(any("across automatic searches" in warning for warning in combined.warnings))


if __name__ == "__main__":
    unittest.main()
