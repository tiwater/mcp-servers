import json
import io
import tempfile
import unittest
from contextlib import redirect_stderr
from pathlib import Path
from unittest.mock import patch

import fitz

from tiwater_pdf.cli import _run_ocr_batch, llm_ocr


class OcrBatchTest(unittest.TestCase):
    def test_llm_ocr_emits_page_progress_before_the_document_finishes(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "two-pages.pdf"
            with fitz.open() as document:
                document.new_page()
                document.new_page()
                document.save(source)

            page_result = {
                "page": 1,
                "text": "visible",
                "table_rows": [],
                "table_cell_lines": [],
                "table_cell_units": [],
                "form_fields": [],
            }
            stream = io.StringIO()
            with patch("tiwater_pdf.cli._resolve_llm_client", return_value=object()), patch(
                "tiwater_pdf.cli._run_with_orientation_correction",
                return_value=(page_result, 0, [1]),
            ), redirect_stderr(stream):
                result = llm_ocr(
                    source,
                    api_key="test-key",
                    base_url="https://example.invalid/v1",
                    llm_model="unseen-model",
                    enable_thinking=False,
                    max_page_parallel=1,
                )

            self.assertEqual(result["page_count"], 2)
            self.assertRegex(
                stream.getvalue(),
                r"^\[tiwater-pdf-ocr-page\] completed=1/2 page=1\n"
                r"\[tiwater-pdf-ocr-page\] completed=2/2 page=1\n$",
            )

    def test_batch_emits_one_flushed_progress_line_per_completed_input(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            inputs = [tmp_path / "first.pdf", tmp_path / "second.pdf"]
            for source in inputs:
                source.write_bytes(b"%PDF-1.4\n")

            stream = io.StringIO()
            with redirect_stderr(stream):
                manifest = _run_ocr_batch(
                    inputs,
                    output_dir=tmp_path / "ocr",
                    max_parallel=2,
                    pages=[1],
                    ocr_func=lambda input_path, _pages: {
                        "file": str(input_path),
                        "model": "unseen-model",
                        "page_count": 1,
                        "pages": [],
                        "text": input_path.stem,
                    },
                    model="unseen-model",
                    provider="llm",
                    enable_thinking=False,
                )

            lines = stream.getvalue().splitlines()
            self.assertEqual(manifest["success_count"], 2)
            self.assertEqual(len(lines), 2)
            self.assertRegex(lines[0], r"^\[tiwater-pdf-ocr\] completed=1/2 status=success duration_ms=\d+$")
            self.assertRegex(lines[1], r"^\[tiwater-pdf-ocr\] completed=2/2 status=success duration_ms=\d+$")

    def test_batch_writes_per_file_outputs_and_manifest(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            first = tmp_path / "first scan.pdf"
            second = tmp_path / "second.pdf"
            first.write_bytes(b"%PDF-1.4\n")
            second.write_bytes(b"%PDF-1.4\n")
            output_dir = tmp_path / "ocr"
            calls = []

            def fake_ocr(input_path, pages):
                calls.append((input_path.name, pages))
                return {
                    "file": str(input_path),
                    "model": "qwen3.7-plus",
                    "page_count": 1,
                    "pages": [{"page": pages[0], "text": input_path.stem, "tables": [], "warnings": []}],
                    "text": input_path.stem,
                }

            manifest = _run_ocr_batch(
                [first, second],
                output_dir=output_dir,
                max_parallel=2,
                pages=[1],
                ocr_func=fake_ocr,
                model="qwen3.7-plus",
                provider="llm",
                enable_thinking=False,
            )

            self.assertEqual(manifest["file_count"], 2)
            self.assertEqual(manifest["success_count"], 2)
            self.assertEqual(manifest["failure_count"], 0)
            self.assertNotIn("text", manifest)
            self.assertNotIn("text", manifest["files"][0])
            self.assertEqual(sorted(name for name, _ in calls), ["first scan.pdf", "second.pdf"])

            for item in manifest["files"]:
                self.assertEqual(item["status"], "success")
                self.assertEqual(item["model"], "qwen3.7-plus")
                self.assertEqual(item["provider"], "llm")
                self.assertIs(item["enable_thinking"], False)
                self.assertGreaterEqual(item["duration_ms"], 0)
                self.assertTrue(Path(item["output"]).exists())
                self.assertTrue(Path(item["status_path"]).exists())
                written = json.loads(Path(item["output"]).read_text())
                self.assertEqual(written["model"], "qwen3.7-plus")

    def test_batch_fails_file_when_any_page_ocr_is_incomplete(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "scan.pdf"
            source.write_bytes(b"%PDF-1.4\n")

            def failed_page_ocr(_input_path, _pages):
                raise RuntimeError("OCR page 10 failed after bounded retries")

            stream = io.StringIO()
            with redirect_stderr(stream):
                manifest = _run_ocr_batch(
                    [source],
                    output_dir=tmp_path / "ocr",
                    max_parallel=1,
                    pages=None,
                    ocr_func=failed_page_ocr,
                    model="qwen3.7-plus",
                    provider="llm",
                    enable_thinking=False,
                )

            self.assertEqual(manifest["success_count"], 0)
            self.assertEqual(manifest["failure_count"], 1)
            self.assertEqual(manifest["files"][0]["status"], "failed")
            self.assertIn("OCR page 10 failed", manifest["files"][0]["error"])
            self.assertRegex(
                stream.getvalue().strip(),
                r"^\[tiwater-pdf-ocr\] completed=1/1 status=failed duration_ms=\d+$",
            )


if __name__ == "__main__":
    unittest.main()
