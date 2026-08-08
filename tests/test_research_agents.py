import unittest

from research_agents import ResearchOrchestrator, clean_credit_reports, clean_source_dossier, normalize_url


class _ScreeningOrchestrator(ResearchOrchestrator):
    def __init__(self):
        pass

    def _plain_json(self, prompt: str, max_output_tokens: int = 3000):
        return {
            "classifications": [
                {"row_ref": "2", "relevance_tier": "Core", "rationale": "Direct product overlap"},
                {"row_ref": "3", "relevance_tier": "Excluded", "rationale": "Unrelated software"},
            ]
        }


class _PlanningOrchestrator(ResearchOrchestrator):
    def __init__(self):
        pass

    def _plain_json(self, prompt: str, max_output_tokens: int = 3000):
        return {
            "queries": [
                "industrial safety equipment manufacturers",
                "industrial safety equipment manufacturers",
                "fall protection and personal protective equipment",
                "this fourth query must be ignored",
            ]
        }


class _CapturingResponses:
    def __init__(self):
        self.kwargs = None

    def create(self, **kwargs):
        self.kwargs = kwargs
        return type("Response", (), {"output_text": "{}"})()


class _CapturingClient:
    def __init__(self):
        self.responses = _CapturingResponses()


class ResearchAgentCleaningTests(unittest.TestCase):
    def test_normalize_url_rejects_unsafe_schemes_and_local_hosts(self):
        self.assertEqual(normalize_url("javascript:alert(1)"), "")
        self.assertEqual(normalize_url("http://localhost/private"), "")
        self.assertEqual(normalize_url("https://Example.com/report.pdf#page=2"), "https://example.com/report.pdf")

    def test_source_dossier_deduplicates_and_normalizes_types(self):
        result = clean_source_dossier(
            {
                "official_website": "https://Example.com",
                "sources": [
                    {"source_type": "annual_report", "url": "https://example.com/a.pdf", "title": "A"},
                    {"source_type": "unknown", "url": "https://example.com/a.pdf", "title": "Duplicate"},
                    {"source_type": "unknown", "url": "https://example.com/b", "title": "B"},
                ],
            }
        )
        self.assertEqual(result["official_website"], "https://example.com/")
        self.assertEqual(len(result["sources"]), 2)
        self.assertEqual(result["sources"][1]["source_type"], "other")

    def test_credit_reports_only_keep_official_rating_agency_domains(self):
        result = clean_credit_reports(
            {
                "reports": [
                    {
                        "agency": "ICRA",
                        "rating": "[ICRA]A",
                        "report_url": "https://www.icra.in/Rating/ShowRationalReportFilePdf/1",
                    },
                    {
                        "agency": "Unverified",
                        "rating": "AAA",
                        "report_url": "https://random-directory.example/rating",
                    },
                ]
            }
        )
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["agency"], "ICRA")

    def test_transaction_screening_preserves_deal_facts(self):
        deals = [
            {"source_row": 2, "target": "Alpha Safety", "description": "Safety equipment"},
            {"source_row": 3, "target": "Beta Software", "description": "Software"},
        ]
        result = _ScreeningOrchestrator().classify_transactions("Subject Co", {}, deals)
        self.assertEqual(result[0]["relevance_tier"], "Core")
        self.assertEqual(result[0]["target"], "Alpha Safety")
        self.assertEqual(result[1]["relevance_tier"], "Excluded")

    def test_mergermarket_search_plan_is_bounded_and_deduplicated(self):
        queries = _PlanningOrchestrator().build_mergermarket_search_plan(
            "Subject Co",
            {"primary_market": "Industrial safety", "keywords": ["PPE", "fall protection"]},
        )
        self.assertEqual(len(queries), 3)
        self.assertIn("fall protection", queries[1])

    def test_web_research_request_does_not_combine_search_with_json_mode(self):
        orchestrator = ResearchOrchestrator.__new__(ResearchOrchestrator)
        orchestrator.client = _CapturingClient()
        orchestrator.model = "gpt-5.6-terra"
        orchestrator._web_json("Return only JSON.")
        self.assertNotIn("text", orchestrator.client.responses.kwargs)
        self.assertEqual(
            orchestrator.client.responses.kwargs["tools"],
            [{"type": "web_search"}],
        )


if __name__ == "__main__":
    unittest.main()
