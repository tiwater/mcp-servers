import json
import tempfile
import unittest
from pathlib import Path

from tiwater_pdf.cli import _run_ocr_batch


class OcrBatchTest(unittest.TestCase):
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


if __name__ == "__main__":
    unittest.main()
