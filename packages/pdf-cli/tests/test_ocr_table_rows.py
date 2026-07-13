import unittest

from tiwater_pdf.cli import _extract_markdown_table_rows, _extract_table_cell_lines, _extract_table_cell_units, _extract_table_logical_rows


class OcrTableRowsTest(unittest.TestCase):
    def test_exports_stable_rows_and_preserves_blank_continuation_cells(self):
        tables = [
            """| Group | Item | Criterion | Method |
|---|---|---|---|
| Alpha | Item One | First value | M-1 |
| | | Second value | |
| | Item Two | Third value | M-2 |"""
        ]

        rows = _extract_markdown_table_rows(tables, 5)

        self.assertEqual([row["row_id"] for row in rows], [
            "page-5-table-0-row-0",
            "page-5-table-0-row-1",
            "page-5-table-0-row-2",
            "page-5-table-0-row-3",
        ])
        self.assertTrue(rows[0]["is_header"])
        self.assertFalse(rows[1]["is_header"])
        self.assertEqual(rows[2]["cells"], ["", "", "Second value", ""])

    def test_handles_escaped_pipes_and_br_markup(self):
        tables = [
            """| Name | Value |
|---|---|
| A \\| B | line one<br>line two |"""
        ]

        rows = _extract_markdown_table_rows(tables, 2)

        self.assertEqual(rows[1]["cells"], ["A | B", "line one\nline two"])

    def test_ignores_non_string_table_entries(self):
        self.assertEqual(_extract_markdown_table_rows([None, {"rows": []}], 1), [])

    def test_exports_each_multiline_cell_value_as_stable_line_evidence(self):
        rows = _extract_markdown_table_rows([
            """| Item | Criterion |
|---|---|
| One | First<br>Second<br>Third |"""
        ], 4)

        lines = [line for line in _extract_table_cell_lines(rows) if line["cell_index"] == 1 and line["row_index"] == 1]

        self.assertEqual([line["line_id"] for line in lines], [
            "page-4-table-0-row-1-cell-1-line-0",
            "page-4-table-0-row-1-cell-1-line-1",
            "page-4-table-0-row-1-cell-1-line-2",
        ])
        self.assertEqual([line["text"] for line in lines], ["First", "Second", "Third"])

    def test_keeps_independent_multiline_values_as_separate_semantic_units(self):
        rows = _extract_markdown_table_rows([
            """| Item | Criterion |
|---|---|
| One | Main value at least 98%<br>High variant at most 2%<br>Low variant at most 0.5% |"""
        ], 4)

        units = [unit for unit in _extract_table_cell_units(rows) if unit["cell_index"] == 1 and unit["row_index"] == 1]

        self.assertEqual([unit["text"] for unit in units], [
            "Main value at least 98%",
            "High variant at most 2%",
            "Low variant at most 0.5%",
        ])
        self.assertTrue(all(len(unit["source_line_ids"]) == 1 for unit in units))

    def test_joins_a_wrapped_measurement_unit_to_the_preceding_value(self):
        rows = _extract_markdown_table_rows([
            """| Item | Criterion |
|---|---|
| One | Count must not exceed 3<br>cfu/30 ml |"""
        ], 7)

        units = [unit for unit in _extract_table_cell_units(rows) if unit["cell_index"] == 1 and unit["row_index"] == 1]

        self.assertEqual(len(units), 1)
        self.assertEqual(units[0]["text"], "Count must not exceed 3 cfu/30 ml")
        self.assertEqual(units[0]["source_line_ids"], [
            "page-7-table-0-row-1-cell-1-line-0",
            "page-7-table-0-row-1-cell-1-line-1",
        ])

    def test_does_not_merge_a_standalone_word_with_the_previous_line(self):
        rows = _extract_markdown_table_rows([
            """| Item | Criterion |
|---|---|
| One | First criterion<br>Pending |"""
        ], 8)

        units = [unit for unit in _extract_table_cell_units(rows) if unit["cell_index"] == 1 and unit["row_index"] == 1]

        self.assertEqual([unit["text"] for unit in units], ["First criterion", "Pending"])

    def test_joins_a_suffix_only_next_page_row_to_its_logical_owner(self):
        page_one_rows = _extract_markdown_table_rows(["""| Group | Item | Criterion | Method |
|---|---|---|---|
| Physical | pH | 5.5 +/- 0.3 | Pharmacopeia |"""], 3)
        page_two_rows = _extract_markdown_table_rows(["""| | | | General rule 0631<br>SOP-2021 |
|---|---|---|---|
| | Osmolality | 240 - 360 | SOP-2025 |"""], 4)

        rows = _extract_table_logical_rows([
            {"page": 3, "table_rows": page_one_rows},
            {"page": 4, "table_rows": page_two_rows},
        ])

        ph = next(row for row in rows if "pH" in row["cells"])
        self.assertEqual(ph["cells"][3], "Pharmacopeia\nGeneral rule 0631\nSOP-2021")
        self.assertEqual(ph["source_row_ids"], ["page-3-table-0-row-1", "page-4-table-0-row-0"])
        self.assertFalse(any(row["row_id"] == "page-4-table-0-row-0" for row in rows))

    def test_does_not_join_a_next_page_row_that_starts_a_new_item(self):
        page_one_rows = _extract_markdown_table_rows(["| Group | Item | Criterion | Method |\n|---|---|---|---|\n| A | One | C1 | M1 |"], 1)
        page_two_rows = _extract_markdown_table_rows(["| Group | Item | Criterion | Method |\n|---|---|---|---|\n| B | Two | C2 | M2 |"], 2)
        rows = _extract_table_logical_rows([{"page": 1, "table_rows": page_one_rows}, {"page": 2, "table_rows": page_two_rows}])
        self.assertEqual(len(rows), 4)

    def test_joins_suffix_after_a_blank_continuation_header_row(self):
        page_one_rows = _extract_markdown_table_rows(["""| Group | Item | Criterion | Method |
|---|---|---|---|
| Physical | pH | 5.5 +/- 0.3 | Pharmacopeia 2020 |"""], 3)
        page_two_rows = _extract_markdown_table_rows(["""| | | | |
|---|---|---|---|
| | | | General rule 0631<br>SOP03-5-2021 |
| | Osmolality | 240 - 360 | SOP03-5-2025 |"""], 4)

        rows = _extract_table_logical_rows([
            {"page": 3, "table_rows": page_one_rows},
            {"page": 4, "table_rows": page_two_rows},
        ])

        ph = next(row for row in rows if "pH" in row["cells"])
        self.assertEqual(ph["cells"][3], "Pharmacopeia 2020\nGeneral rule 0631\nSOP03-5-2021")
        self.assertEqual(ph["source_row_ids"], ["page-3-table-0-row-1", "page-4-table-0-row-1"])
        self.assertTrue(any(row["row_id"] == "page-4-table-0-row-0" for row in rows))
        self.assertFalse(any(row["row_id"] == "page-4-table-0-row-1" for row in rows))

    def test_does_not_skip_a_new_item_after_a_blank_continuation_header(self):
        page_one_rows = _extract_markdown_table_rows(["| Group | Item | Criterion | Method |\n|---|---|---|---|\n| A | One | C1 | M1 |"], 1)
        page_two_rows = _extract_markdown_table_rows(["| | | | |\n|---|---|---|---|\n| B | Two | C2 | M2 |"], 2)

        rows = _extract_table_logical_rows([
            {"page": 1, "table_rows": page_one_rows},
            {"page": 2, "table_rows": page_two_rows},
        ])

        one = next(row for row in rows if "One" in row["cells"])
        two = next(row for row in rows if "Two" in row["cells"])
        self.assertEqual(one["source_row_ids"], ["page-1-table-0-row-1"])
        self.assertEqual(two["source_row_ids"], ["page-2-table-0-row-1"])


if __name__ == "__main__":
    unittest.main()
