import os
import tempfile
import unittest
from pathlib import Path

import fitz

from tiwater_pdf.cli import extract_table_details


class ExtractTableDetailsTest(unittest.TestCase):
    def test_exports_cell_bboxes_text_spans_and_colors(self):
        path = self._create_table_pdf()

        result = extract_table_details(path)

        self.assertEqual(result["tables_found"], 1)
        table = result["tables"][0]
        self.assertEqual(table["rowCount"], 2)
        self.assertEqual(table["columnCount"], 2)
        self.assertEqual(table["expectedGridCellCount"], 4)
        self.assertEqual(table["detectedCellCount"], 4)
        self.assertGreater(len(table["lineSegments"]), 0)

        sequence_cell = table["detailRows"][1]["cells"][0]
        self.assertEqual(sequence_cell["extractedText"], "QVQLV")
        self.assertIn("#FF0000", sequence_cell["spanColors"])
        self.assertTrue(sequence_cell["bbox"])
        self.assertTrue(any(span["text"] == "Q" and span["color"] == "#FF0000" for span in sequence_cell["spans"]))

    @staticmethod
    def _create_table_pdf() -> Path:
        fd, filename = tempfile.mkstemp(suffix=".pdf")
        os.close(fd)
        Path(filename).unlink(missing_ok=True)
        path = Path(filename)

        doc = fitz.open()
        page = doc.new_page(width=240, height=160)
        xs = [20, 120, 220]
        ys = [20, 70, 120]
        for x in xs:
            page.draw_line((x, ys[0]), (x, ys[-1]), color=(0, 0, 0), width=1)
        for y in ys:
            page.draw_line((xs[0], y), (xs[-1], y), color=(0, 0, 0), width=1)
        page.insert_text((45, 50), "Header", fontsize=10)
        page.insert_text((145, 50), "Value", fontsize=10)
        page.insert_text((45, 100), "QV", fontsize=10)
        page.insert_text((57, 100), "Q", fontsize=10, color=(1, 0, 0))
        page.insert_text((65, 100), "LV", fontsize=10)
        page.insert_text((145, 100), "99.7", fontsize=10)
        doc.save(path)
        doc.close()
        return path


if __name__ == "__main__":
    unittest.main()
