import json
import unittest
from types import SimpleNamespace

from tiwater_pdf.cli import (
    _VisionResponseFormatError,
    _call_vision_page_with_retry,
    _call_vision_with_retry,
    _is_retryable_vision_page_error,
    _parse_vision_page_response,
)


def response_with(content):
    return SimpleNamespace(
        choices=[SimpleNamespace(message=SimpleNamespace(content=content))]
    )


class VisionPageRetryTest(unittest.TestCase):
    def test_response_format_gateway_error_is_classified_for_compatible_retry(self):
        response = SimpleNamespace(
            choices=None,
            error={
                "code": "invalid_parameter_error",
                "message": "Model output became abnormal while generating a JSON response for response_format.",
            },
        )

        with self.assertRaises(_VisionResponseFormatError):
            _parse_vision_page_response(response, 2)

    def test_response_format_error_retries_without_response_format(self):
        uses_response_format = []

        def request(use_response_format):
            uses_response_format.append(use_response_format)
            if use_response_format:
                return SimpleNamespace(
                    choices=None,
                    error={
                        "code": "invalid_parameter_error",
                        "message": "JSON generation failed for response_format",
                    },
                )
            return response_with(json.dumps({"text": "251", "tables": [], "warnings": []}))

        page, attempts = _call_vision_page_with_retry(
            request,
            lambda response: _parse_vision_page_response(response, 2),
            sleep_fn=lambda _seconds: None,
        )

        self.assertEqual(uses_response_format, [True, False])
        self.assertEqual(attempts, 2)
        self.assertEqual(page["text"], "251")

    def test_malformed_page_response_is_retried_then_preserves_valid_evidence(self):
        responses = iter([
            SimpleNamespace(choices=None),
            response_with(json.dumps({
                "text": "Sample 240471-01",
                "tables": [
                    "| Sample | Result |\n|---|---|\n| 240471-01 | 69.8 |"
                ],
                "warnings": [],
            })),
        ])

        page, attempts = _call_vision_with_retry(
            lambda: _parse_vision_page_response(next(responses), 10),
            attempts=3,
            sleep_fn=lambda _seconds: None,
            retryable=_is_retryable_vision_page_error,
        )

        self.assertEqual(attempts, 2)
        self.assertEqual(page["text"], "Sample 240471-01")
        self.assertEqual(page["table_rows"][1]["cells"], ["240471-01", "69.8"])

    def test_persistently_malformed_page_response_fails_after_bound(self):
        calls = 0

        def malformed():
            nonlocal calls
            calls += 1
            return _parse_vision_page_response(SimpleNamespace(choices=None), 10)

        with self.assertRaisesRegex(ValueError, "vision response contains no choices"):
            _call_vision_with_retry(
                malformed,
                attempts=3,
                sleep_fn=lambda _seconds: None,
                retryable=_is_retryable_vision_page_error,
            )

        self.assertEqual(calls, 3)


if __name__ == "__main__":
    unittest.main()
