from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
import threading
import os
import time
import unittest

import pymupdf as fitz
from tiwater_pdf import cli


class OcrSchedulingTests(unittest.TestCase):
    def run_ocr(self, root, *, fail=False, parallel=1, selected=None):
        pdf = Path(root) / 'synthetic.pdf'
        with fitz.open() as document:
            for _ in range(8):
                document.new_page(width=100, height=150)
            document.save(pdf)
        seen = []
        local = threading.local()
        progress = []

        def render(document, index, **kwargs):
            local.index = index
            seen.append(index)
            return b'synthetic-image'

        def vision(*args, **kwargs):
            if fail and local.index == 0:
                raise ValueError('frozen-provider-failure')
            time.sleep(0.01)
            return {'page': local.index + 1, 'text': f'page-{local.index + 1}', 'orientation_degrees': 0}, 1

        with patch.dict(os.environ, {'SUPEN_LLM_TOKEN': 'synthetic', 'SUPEN_LLM_GATEWAY_URL': 'https://invalid.example/v1'}), patch('openai.OpenAI'), patch.object(cli, '_render_page_image', side_effect=render), patch.object(cli, '_call_vision_page_with_retry', side_effect=vision):
            if fail:
                with self.assertRaisesRegex(RuntimeError, 'OCR page 1 failed') as caught:
                    cli.llm_ocr(pdf, max_page_parallel=parallel)
                self.assertEqual(caught.exception.page_number, 1)
                return seen
            result = cli.llm_ocr(pdf, max_page_parallel=parallel, page_numbers=selected, page_progress_func=lambda done, total, page: progress.append((done, total)))
            return result, progress

    def test_terminal_failure_does_not_start_remaining_pages(self):
        with TemporaryDirectory() as root:
            self.assertEqual(self.run_ocr(root, fail=True), [0])

    def test_success_preserves_selected_page_identity_order_and_progress(self):
        for selected, parallel, expected in [([2, 5], 1, [2, 5]), (None, 3, list(range(1, 9))), ([99], 2, [])]:
            with self.subTest(selected=selected), TemporaryDirectory() as root:
                result, progress = self.run_ocr(root, parallel=parallel, selected=selected)
                self.assertEqual([page['page'] for page in result['pages']], expected)
                self.assertEqual(result['text'], '\n\n'.join(f'page-{number}' for number in expected))
                self.assertEqual(progress, [(number, len(expected)) for number in range(1, len(expected) + 1)])


if __name__ == '__main__':
    unittest.main()
