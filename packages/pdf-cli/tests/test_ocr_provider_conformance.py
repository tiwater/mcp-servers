import json
import unittest
from types import SimpleNamespace

from tiwater_pdf.cli import (
    _call_vision_page_with_retry,
    _extract_table_logical_rows,
    _parse_vision_page_response,
)


def response_with(payload):
    return SimpleNamespace(
        choices=[SimpleNamespace(message=SimpleNamespace(content=json.dumps(payload, ensure_ascii=False)))]
    )


class OcrProviderConformanceTest(unittest.TestCase):
    def test_preserves_complete_unseen_multilingual_table_structure(self):
        cover = _parse_vision_page_response(response_with({
            "text": "封面 Cover sheet",
            "tables": [],
            "fields": [],
            "orientation_degrees": 0,
            "warnings": [],
        }), 1)
        evidence = _parse_vision_page_response(response_with({
            "text": "检测观察 Observation ledger",
            "tables": [
                "| 检测观察 Observation ledger | | |\n"
                "| 样本标识 Specimen identity | 数值结果 Numeric result | 处置 Disposition |\n"
                "|---|---|---|\n"
                "| NX-77-Z | 18.42 mg/L | ☑接受 Accepted □拒绝 Rejected |\n"
                "| | | |"
            ],
            "fields": [],
            "orientation_degrees": 0,
            "warnings": [],
        }), 2)

        rows = evidence["table_rows"]
        logical_rows = _extract_table_logical_rows([cover, evidence])

        self.assertEqual([row["row_id"] for row in rows], [
            "page-2-table-0-row-0",
            "page-2-table-0-row-1",
            "page-2-table-0-row-2",
            "page-2-table-0-row-3",
        ])
        self.assertTrue(rows[0]["is_header"])
        self.assertTrue(rows[1]["is_header"])
        self.assertEqual(rows[1]["cells"], [
            "样本标识 Specimen identity", "数值结果 Numeric result", "处置 Disposition",
        ])
        identity_rows = [row for row in rows if row["cells"][0] == "NX-77-Z"]
        self.assertEqual(len(identity_rows), 1)
        self.assertEqual(identity_rows[0]["cells"][1], "18.42 mg/L")
        self.assertEqual(rows[3]["cells"], ["", "", ""])
        self.assertEqual(
            [row["logical_row_id"] for row in logical_rows],
            [row["row_id"] for row in rows],
        )

    def test_retries_then_rejects_rows_split_into_incomplete_table_fragments(self):
        fragmented = response_with({
            "text": "检测观察 Observation ledger",
            "tables": [
                "| 检测观察 Observation ledger | | |",
                "| 样本标识 Specimen identity | 数值结果 Numeric result | 处置 Disposition |",
                "|---|---|---|",
                "| NX-77-Z | 18.42 mg/L | ☑接受 Accepted □拒绝 Rejected |",
                "| | | |",
            ],
            "fields": [],
            "orientation_degrees": 0,
            "warnings": [],
        })
        calls = []

        with self.assertRaisesRegex(ValueError, "table 2 has a separator without a header"):
            _call_vision_page_with_retry(
                lambda use_response_format: calls.append(use_response_format) or fragmented,
                lambda response: _parse_vision_page_response(response, 2),
                attempts=3,
                sleep_fn=lambda _seconds: None,
            )

        self.assertEqual(len(calls), 3)

    def test_rejects_multiline_table_without_separator(self):
        with self.assertRaisesRegex(ValueError, "table 0 has multiple rows but no separator"):
            _parse_vision_page_response(response_with({
                "text": "Inventory",
                "tables": ["| Key | Reading |\n| novel-1 | 7.3 |"],
                "fields": [],
                "orientation_degrees": 0,
                "warnings": [],
            }), 4)

    def test_evidence_free_page_requires_explicit_blank_page_observation(self):
        with self.assertRaisesRegex(ValueError, "no content and is not declared blank"):
            _parse_vision_page_response(response_with({
                "text": "",
                "tables": [],
                "fields": [],
                "orientation_degrees": 0,
                "warnings": [],
            }), 3)

        page = _parse_vision_page_response(response_with({
            "text": "",
            "tables": [],
            "fields": [],
            "orientation_degrees": 0,
            "warnings": ["blank_page"],
        }), 3)
        self.assertEqual(page["warnings"], ["blank_page"])


if __name__ == "__main__":
    unittest.main()
