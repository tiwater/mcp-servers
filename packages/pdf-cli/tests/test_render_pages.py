import hashlib
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

import pymupdf as fitz

from tiwater_pdf.cli import render_pages


class RenderPagesTests(unittest.TestCase):
    def make_pdf(self, root: Path) -> Path:
        pdf = root / "input.pdf"
        with fitz.open() as document:
            first = document.new_page(width=200, height=300)
            first.insert_text((20, 40), "first page")
            second = document.new_page(width=320, height=180)
            second.insert_text((20, 40), "second page")
            document.save(pdf)
        return pdf

    def test_renders_complete_ordered_revision_bound_page_set(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            pdf = self.make_pdf(root)
            output = root / "pages"

            result = render_pages(pdf, output, zoom=1.5)

            self.assertEqual(result["schema"], "tiwater.pdf-render-pages/v1")
            self.assertEqual(result["input_sha256"], hashlib.sha256(pdf.read_bytes()).hexdigest())
            self.assertEqual(result["page_count"], 2)
            self.assertEqual([page["page"] for page in result["pages"]], [1, 2])
            self.assertEqual([Path(page["path"]).name for page in result["pages"]], [
                "page-0001.png", "page-0002.png",
            ])
            self.assertNotEqual(result["pages"][0]["height"], result["pages"][1]["height"])
            for page in result["pages"]:
                rendered = Path(page["path"])
                self.assertTrue(rendered.is_file())
                self.assertEqual(page["bytes"], rendered.stat().st_size)
                self.assertEqual(page["sha256"], hashlib.sha256(rendered.read_bytes()).hexdigest())
            self.assertEqual(json.loads((output / "manifest.json").read_text()), result)

    def test_rejects_existing_output_without_changing_it(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            pdf = self.make_pdf(root)
            output = root / "pages"
            output.mkdir()
            retained = output / "retained.txt"
            retained.write_text("keep")

            with self.assertRaises(FileExistsError):
                render_pages(pdf, output)

            self.assertEqual(retained.read_text(), "keep")
            self.assertEqual(list(output.iterdir()), [retained])

    def test_rejects_non_positive_zoom_before_creating_output(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            pdf = self.make_pdf(root)
            output = root / "pages"

            with self.assertRaises(ValueError):
                render_pages(pdf, output, zoom=0)

            self.assertFalse(output.exists())


if __name__ == "__main__":
    unittest.main()
