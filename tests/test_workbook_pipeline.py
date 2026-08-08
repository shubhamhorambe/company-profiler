import unittest

from pipeline_core_v2 import (
    _responses_create,
    clear_table_sheet_keep_header,
    evidence_gated_research_flat,
    ensure_credit_rating_sheet,
    ensure_source_dossier_sheet,
    load_template_workbook,
    write_evidence_register,
    write_credit_reports,
    write_source_dossier,
)


class _FakeResponses:
    def __init__(self):
        self.kwargs = None

    def create(self, **kwargs):
        self.kwargs = kwargs
        return object()


class _FakeClient:
    def __init__(self):
        self.responses = _FakeResponses()


class _QueuedResponses:
    def __init__(self, outputs):
        self.outputs = iter(outputs)

    def create(self, **kwargs):
        output = next(self.outputs)
        return type("Response", (), {"output_text": output})()


class _QueuedClient:
    def __init__(self, outputs):
        self.responses = _QueuedResponses(outputs)


class WorkbookPipelineTests(unittest.TestCase):
    def test_current_responses_parameters_are_normalized(self):
        client = _FakeClient()
        _responses_create(
            client,
            model="gpt-5.6-luna",
            input="test",
            response_format={"type": "json_object"},
            temperature=0,
        )
        self.assertNotIn("temperature", client.responses.kwargs)
        self.assertEqual(client.responses.kwargs["text"]["format"]["type"], "json_object")

    def test_rewrite_cannot_introduce_a_source_not_in_the_evidence_set(self):
        client = _QueuedClient(
            [
                '{"facts":[{"claim":"Verified fact","evidence_url":"https://example.com/fact","evidence_quality":"official"}]}',
                "DESC:\n- Verified fact\nSOURCES:\nhttps://invented.example/not-evidence",
            ]
        )
        description, sources = evidence_gated_research_flat(
            "Example Limited",
            "About the company",
            {},
            client,
            {},
        )
        self.assertIn("Verified fact", description)
        self.assertEqual(sources, "https://example.com/fact")

    def test_new_audit_sheets_preserve_headers_and_write_rows(self):
        workbook = load_template_workbook()
        source_sheet = ensure_source_dossier_sheet(workbook)
        rating_sheet = ensure_credit_rating_sheet(workbook)
        clear_table_sheet_keep_header(source_sheet)
        clear_table_sheet_keep_header(rating_sheet)

        write_source_dossier(
            source_sheet,
            {
                "legal_name": "Example Limited",
                "official_website": "https://example.com/",
                "sources": [
                    {
                        "source_type": "annual_report",
                        "publisher": "Example Limited",
                        "title": "Annual report",
                        "published_date": "2025",
                        "url": "https://example.com/ar.pdf",
                        "notes": "Primary source",
                    }
                ],
            },
        )
        write_credit_reports(
            rating_sheet,
            [
                {
                    "agency": "ICRA",
                    "report_date": "2025-01-01",
                    "rating": "[ICRA]A",
                    "report_url": "https://www.icra.in/report.pdf",
                }
            ],
        )

        self.assertEqual(source_sheet.cell(1, 1).value, "Source Type")
        self.assertEqual(source_sheet.cell(2, 1).value, "official_website")
        self.assertEqual(source_sheet.cell(3, 1).value, "annual_report")
        self.assertEqual(rating_sheet.cell(1, 1).value, "Agency")
        self.assertEqual(rating_sheet.cell(2, 1).value, "ICRA")

    def test_evidence_register_preserves_claim_level_sources(self):
        workbook = load_template_workbook()
        write_evidence_register(
            workbook,
            {
                ("example limited", "about the company"): [
                    {
                        "claim": "The company manufactures safety equipment.",
                        "evidence_url": "https://example.com/about",
                        "evidence_quality": "official",
                    }
                ]
            },
        )
        sheet = workbook["Evidence Register"]
        self.assertEqual(sheet["A2"].value, "about the company")
        self.assertEqual(sheet["B2"].value, "The company manufactures safety equipment.")
        self.assertEqual(sheet["D2"].value, "https://example.com/about")


if __name__ == "__main__":
    unittest.main()
