import json
import unittest
from types import SimpleNamespace

from tiwater_pdf.cli import (
    OCR_PAGE_PROMPT,
    _call_vision_page_with_retry,
    _extract_table_logical_rows,
    _parse_vision_page_response,
    _split_markdown_table_row,
    _table_unit_choices,
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

    def test_normalizes_multiline_table_without_separator(self):
        page = _parse_vision_page_response(response_with({
            "text": "Inventory",
            "tables": ["| Key | Reading |\n| novel-1 | 7.3 |\n| novel-2 | 8.1 |"],
            "fields": [],
            "orientation_degrees": 0,
            "warnings": [],
        }), 4)
        self.assertEqual(len(page["tables"]), 1)
        self.assertIn("| --- | --- |", page["tables"][0])
        self.assertEqual(len(page["table_rows"]), 3)
        self.assertTrue(page["table_rows"][0]["is_header"])
        self.assertEqual(page["table_rows"][0]["cells"], ["Key", "Reading"])
        self.assertFalse(page["table_rows"][1]["is_header"])
        self.assertEqual(page["table_rows"][1]["cells"], ["novel-1", "7.3"])
        self.assertEqual(page["table_rows"][2]["cells"], ["novel-2", "8.1"])

        with self.assertRaisesRegex(ValueError, "table 0 has multiple rows but no separator"):
            _parse_vision_page_response(response_with({
                "text": "Inventory",
                "tables": ["| OnlyOneCol |\n| AnotherCol |"],
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

    def test_extracts_separatorless_pipe_table_from_text_without_duplicating_tables(self):
        dilution_table = (
            "| Well | Sample | Dil | A | B | C | D | E | F | G | H |\n"
            "|---|---|---|---|---|---|---|---|---|---|---|\n"
            "| A1 | LOT-901 | 10 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 |"
        )
        page = _parse_vision_page_response(response_with({
            "text": (
                "Assay Summary Report\n"
                "Operator: Jane | Instrument: Reader-02\n"
                "Sample Code | Concentration (mg/mL) | Impurity (ng/mL) | Reported (ng/mg)\n"
                "LOT-901-A | 115.20 | 345.6 | 3\n"
                "LOT-901-B | 108.40 | 216.8 |\n"
                "\n"
                "Well | Sample | Dil | A | B | C | D | E | F | G | H\n"
                "A1 | LOT-901 | 10 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8\n"
                "\n"
                "All acceptance criteria verified.\n"
            ),
            "tables": [dilution_table],
            "fields": [],
            "orientation_degrees": 0,
            "warnings": [],
        }), 1)

        # The 11-column dilution table is NOT duplicated; the 4-column separatorless
        # summary table IS extracted alongside it; the 2-column single metadata line is ignored.
        self.assertEqual(len(page["tables"]), 2)
        summary_rows = [row for row in page["table_rows"] if row["table_index"] == 1]
        self.assertEqual(len(summary_rows), 3)
        self.assertTrue(summary_rows[0]["is_header"])
        self.assertFalse(summary_rows[1]["is_header"])
        self.assertFalse(summary_rows[2]["is_header"])
        self.assertEqual(summary_rows[0]["cells"], [
            "Sample Code", "Concentration (mg/mL)", "Impurity (ng/mL)", "Reported (ng/mg)",
        ])
        self.assertEqual(summary_rows[1]["cells"], ["LOT-901-A", "115.20", "345.6", "3"])
        self.assertEqual(summary_rows[2]["cells"], ["LOT-901-B", "108.40", "216.8", ""])
        self.assertIn("All acceptance criteria verified.", page["text"])

    def test_extracts_separatorless_table_with_outer_pipes_and_preserves_escaped_pipes(self):
        page = _parse_vision_page_response(response_with({
            "text": (
                "| Specimen | Formula | Result |\n"
                "| S-1 | A \\| B | 4.2 |\n"
                "| S-2 | C \\\\ D | 5.1 |"
            ),
            "tables": [],
            "fields": [],
            "orientation_degrees": 0,
            "warnings": [],
        }), 2)

        self.assertEqual(len(page["tables"]), 1)
        self.assertEqual(len(page["table_rows"]), 3)
        self.assertTrue(page["table_rows"][0]["is_header"])
        self.assertEqual(page["table_rows"][1]["cells"], ["S-1", "A | B", "4.2"])
        self.assertEqual(page["table_rows"][2]["cells"], ["S-2", r"C \ D", "5.1"])

    def test_extracts_back_to_back_markdown_tables_in_text_without_blank_line(self):
        page = _parse_vision_page_response(response_with({
            "text": (
                "| Col A | Col B |\n"
                "|---|---|\n"
                "| v1 | v2 |\n"
                "| Col C | Col D |\n"
                "|---|---|\n"
                "| v3 | v4 |"
            ),
            "tables": [],
            "fields": [],
            "orientation_degrees": 0,
            "warnings": [],
        }), 1)

        self.assertEqual(len(page["tables"]), 2)
        t0 = [r for r in page["table_rows"] if r["table_index"] == 0]
        t1 = [r for r in page["table_rows"] if r["table_index"] == 1]
        self.assertEqual([r["cells"] for r in t0], [["Col A", "Col B"], ["v1", "v2"]])
        self.assertEqual([r["cells"] for r in t1], [["Col C", "Col D"], ["v3", "v4"]])

    def test_non_pipe_prose_and_single_or_mismatched_pipe_lines_do_not_become_tables(self):
        page = _parse_vision_page_response(response_with({
            "text": (
                "Analytical Summary Page\n"
                "Sample Suitability = Yes\n"
                "Reported result = 4 ng/mg\n"
                "Document No: SOP-77 | Rev 3\n"
                "Page 1 | Section A | Confidential | Controlled Copy\n"
                "Single line with escaped pipe only: A \\| B\n"
                "Another line with escaped pipe only: C \\| D\n"
            ),
            "tables": [],
            "fields": [],
            "orientation_degrees": 0,
            "warnings": [],
        }), 1)

        self.assertEqual(page["tables"], [])
        self.assertEqual(page["table_rows"], [])

    def test_split_markdown_table_row_preserves_trailing_empty_cell_without_outer_pipes(self):
        self.assertEqual(_split_markdown_table_row("a | b |"), ["a", "b", ""])
        self.assertEqual(_split_markdown_table_row("| a | b |"), ["a", "b"])
        self.assertEqual(_split_markdown_table_row("| a | b | |"), ["a", "b", ""])
        self.assertEqual(_split_markdown_table_row(r"| a | b\| |"), ["a", "b|"])
        self.assertEqual(_split_markdown_table_row(r"| a | b\\|"), ["a", "b\\"])

    def test_ocr_prompt_guides_independent_checkboxes_bilingual_lines_and_column_isolation(self):
        self.assertIn("Inspect every checkbox glyph independently", OCR_PAGE_PROMPT)
        self.assertIn("Never invert adjacent options", OCR_PAGE_PROMPT)
        self.assertIn("handwritten added option lines", OCR_PAGE_PROMPT)
        self.assertIn("Respect vertical table column borders strictly", OCR_PAGE_PROMPT)
        self.assertIn("never bleeding or concatenating them into a neighboring option or result cell", OCR_PAGE_PROMPT)
        self.assertIn("never leave a visible table only as pipe-delimited lines in text", OCR_PAGE_PROMPT)

    def test_table_unit_choices_handles_tight_checkboxes_and_ignores_math_radicals(self):
        tight = _table_unit_choices("☑是(Yes)□否(No)")
        self.assertEqual(tight["options"], [
            {"label": "是(Yes)", "selected": True},
            {"label": "否(No)", "selected": False},
        ])
        self.assertEqual(tight["selected_options"], ["是(Yes)"])
        self.assertEqual(tight["selection_status"], "selected")

        math_text = _table_unit_choices("Area = 2√3 + √2")
        self.assertEqual(math_text["options"], [])
        self.assertEqual(math_text["selection_status"], "not-applicable")

    def test_bilingual_table_cell_units_keep_chinese_english_checkmarks_and_audit_column_separate(self):
        page = _parse_vision_page_response(response_with({
            "text": "Visual Inspection",
            "tables": [
                "| Sample | Color | Clarity | Audit Trail |\n"
                "|---|---|---|---|\n"
                "| LOT-88-01 | ☑ 浅于标准比色液 Y7<br>☑ Less colored than EP Y7 reference solution | ☑ 低于 3 号浊度标准液<br>☑ Lower than reference suspension III | PW A.L. 2026-03-11 |"
            ],
            "fields": [],
            "orientation_degrees": 0,
            "warnings": [],
        }), 2)

        color_units = [u for u in page["table_cell_units"] if u["row_index"] == 1 and u["cell_index"] == 1]
        clarity_units = [u for u in page["table_cell_units"] if u["row_index"] == 1 and u["cell_index"] == 2]
        audit_units = [u for u in page["table_cell_units"] if u["row_index"] == 1 and u["cell_index"] == 3]

        self.assertEqual([u["selected_options"] for u in color_units], [
            ["浅于标准比色液 Y7"],
            ["Less colored than EP Y7 reference solution"],
        ])
        self.assertEqual([u["selected_options"] for u in clarity_units], [
            ["低于 3 号浊度标准液"],
            ["Lower than reference suspension III"],
        ])
        self.assertEqual([u["text"] for u in audit_units], ["PW A.L. 2026-03-11"])

    def test_extracts_labeled_checkbox_form_fields_from_text_and_respects_ambiguity_and_dedup(self):
        page = _parse_vision_page_response(response_with({
            "text": (
                "System suitability valid (系统适用性是否有效): ☑ 是 (Yes) □ 否 (No)\n"
                "Appearance check：□ 合格 (Pass) ☑ 不合格 (Fail)\n"
                "Ambiguous check: ☑ Yes ☑ No\n"
                "Unselected check: □ Yes □ No\n"
                "Already provided: ☑ Yes □ No\n"
                "Note: verify □ mark later\n"
                "Sample ID: LOT-55\n"
                "| Col1 | Col2 |\n"
                "|---|---|\n"
                "| Row | ☑ A □ B |\n"
            ),
            "tables": [],
            "fields": [
                {
                    "label": "Already provided",
                    "value": "Yes",
                    "raw_text": "Already provided: ☑ Yes □ No",
                    "options": [],
                }
            ],
            "orientation_degrees": 0,
            "warnings": [],
        }), 5)

        fields_by_label = {field["label"]: field for field in page["form_fields"]}
        self.assertEqual(set(fields_by_label.keys()), {
            "Already provided",
            "System suitability valid (系统适用性是否有效)",
            "Appearance check",
            "Ambiguous check",
            "Unselected check",
        })
        # Model-provided field had options=[] enriched from raw_text and was not duplicated.
        self.assertEqual(fields_by_label["Already provided"]["selection_status"], "selected")
        self.assertEqual(fields_by_label["Already provided"]["selected_options"], ["Yes"])

        suitability = fields_by_label["System suitability valid (系统适用性是否有效)"]
        self.assertEqual(suitability["value"], "是 (Yes)")
        self.assertEqual(suitability["selected_options"], ["是 (Yes)"])
        self.assertEqual(suitability["selection_status"], "selected")
        self.assertEqual(suitability["options"], [
            {"label": "是 (Yes)", "selected": True},
            {"label": "否 (No)", "selected": False},
        ])

        appearance = fields_by_label["Appearance check"]
        self.assertEqual(appearance["value"], "不合格 (Fail)")
        self.assertEqual(appearance["selection_status"], "selected")

        ambiguous = fields_by_label["Ambiguous check"]
        self.assertEqual(ambiguous["value"], "")
        self.assertEqual(ambiguous["selected_options"], ["Yes", "No"])
        self.assertEqual(ambiguous["selection_status"], "ambiguous")

        unselected = fields_by_label["Unselected check"]
        self.assertEqual(unselected["value"], "")
        self.assertEqual(unselected["selected_options"], [])
        self.assertEqual(unselected["selection_status"], "unselected")


if __name__ == "__main__":
    unittest.main()
