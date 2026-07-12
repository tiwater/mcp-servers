import unittest

from tiwater_pdf.cli import _extract_markdown_table_rows


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


if __name__ == "__main__":
    unittest.main()
