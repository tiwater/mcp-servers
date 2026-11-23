import json
import os
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import unittest
from unittest.mock import patch

import httpx
import openai
import pymupdf as fitz

from tiwater_pdf import cli


def _valid_vision_response(text: str = "page content"):
    payload = {
        "text": text,
        "tables": [],
        "fields": [],
        "orientation_degrees": 0,
        "warnings": [],
    }
    return SimpleNamespace(
        choices=[SimpleNamespace(message=SimpleNamespace(content=json.dumps(payload)))]
    )


def _http_response(status_code: int) -> httpx.Response:
    request = httpx.Request("POST", "https://openrouter.ai/api/v1/chat/completions")
    return httpx.Response(status_code=status_code, request=request)


class OcrTransientRetryTests(unittest.TestCase):
    def test_retries_openai_api_connection_error_until_success(self):
        request = httpx.Request("POST", "https://openrouter.ai/api/v1/chat/completions")
        responses = [
            openai.APIConnectionError(request=request),
            openai.APIConnectionError(request=request),
            _valid_vision_response("recovered page"),
        ]
        seen_response_format = []
        malformed_calls = []
        sleeps = []

        def fake_request(use_response_format: bool):
            seen_response_format.append(use_response_format)
            outcome = responses.pop(0)
            if isinstance(outcome, Exception):
                raise outcome
            return outcome

        result, attempts = cli._call_vision_page_with_retry(
            fake_request,
            lambda response: cli._parse_vision_page_response(response, 7),
            attempts=3,
            sleep_fn=sleeps.append,
            on_malformed_response=lambda: malformed_calls.append(True),
        )

        self.assertEqual(attempts, 3)
        self.assertEqual(result["text"], "recovered page")
        self.assertEqual(sleeps, [2, 4])
        self.assertEqual(seen_response_format, [True, True, True])
        self.assertEqual(malformed_calls, [])

    def test_retries_timeout_rate_limit_5xx_and_wrapped_transport_errors(self):
        request = httpx.Request("POST", "https://openrouter.ai/api/v1/chat/completions")
        retryable_errors = [
            openai.APIConnectionError(request=request),
            openai.APITimeoutError(request=request),
            openai.RateLimitError("Too many requests", response=_http_response(429), body=None),
            openai.InternalServerError("Upstream 502", response=_http_response(502), body=None),
            httpx.ConnectError("connect failed", request=request),
            httpx.ReadTimeout("read timed out", request=request),
            httpx.RemoteProtocolError("server disconnected", request=request),
            ConnectionResetError("peer reset"),
            TimeoutError("socket timed out"),
            RuntimeError("Connection error."),
        ]
        for error in retryable_errors:
            with self.subTest(error=type(error).__name__):
                self.assertTrue(cli._is_retryable_vision_error(error))

        wrapped = RuntimeError("gateway wrapper")
        wrapped.__cause__ = httpx.ConnectError("dns blip", request=request)
        self.assertTrue(cli._is_retryable_vision_error(wrapped))

    def test_does_not_retry_non_retryable_client_auth_errors(self):
        non_retryable_errors = [
            openai.AuthenticationError("Invalid API key", response=_http_response(401), body=None),
            openai.PermissionDeniedError("Forbidden", response=_http_response(403), body=None),
            openai.NotFoundError("Model not found", response=_http_response(404), body=None),
            openai.BadRequestError("Unsupported image", response=_http_response(400), body=None),
            RuntimeError("permanent contract error"),
        ]
        for error in non_retryable_errors:
            with self.subTest(error=type(error).__name__):
                self.assertFalse(cli._is_retryable_vision_error(error))
                calls = []
                sleeps = []
                with self.assertRaises(type(error)) as caught:
                    cli._call_vision_page_with_retry(
                        lambda _use_rf: calls.append(1) or (_ for _ in ()).throw(error),
                        lambda response: cli._parse_vision_page_response(response, 1),
                        attempts=3,
                        sleep_fn=sleeps.append,
                    )
                self.assertEqual(len(calls), 1)
                self.assertEqual(sleeps, [])
                self.assertEqual(caught.exception.request_attempts, 1)

    def test_llm_ocr_retries_per_page_api_connection_error_and_reports_exhausted_attempts(self):
        request = httpx.Request("POST", "https://openrouter.ai/api/v1/chat/completions")
        with TemporaryDirectory() as root:
            pdf = Path(root) / "scan.pdf"
            with fitz.open() as doc:
                doc.new_page(width=100, height=150)
                doc.save(pdf)

            # First verify recovery on attempt 2 inside llm_ocr.
            outcomes = [
                openai.APIConnectionError(request=request),
                _valid_vision_response("page 1 recovered"),
            ]

            def fake_create(**_kwargs):
                outcome = outcomes.pop(0)
                if isinstance(outcome, Exception):
                    raise outcome
                return outcome

            fake_client = SimpleNamespace(
                chat=SimpleNamespace(completions=SimpleNamespace(create=fake_create))
            )
            with (
                patch.dict(os.environ, {"OPENROUTER_API_KEY": "synthetic"}, clear=True),
                patch("openai.OpenAI", return_value=fake_client),
                patch.object(cli.time, "sleep"),
            ):
                result = cli.llm_ocr(pdf, max_page_parallel=1)
            self.assertEqual(result["pages"][0]["text"], "page 1 recovered")
            self.assertEqual(result["pages"][0]["request_attempts"], 2)

            # Next verify bounded exhaustion attaches page_number and request_attempts.
            def always_fail(**_kwargs):
                raise openai.APIConnectionError(request=request)

            failing_client = SimpleNamespace(
                chat=SimpleNamespace(completions=SimpleNamespace(create=always_fail))
            )
            with (
                patch.dict(os.environ, {"OPENROUTER_API_KEY": "synthetic"}, clear=True),
                patch("openai.OpenAI", return_value=failing_client),
                patch.object(cli.time, "sleep"),
            ):
                with self.assertRaisesRegex(RuntimeError, "OCR page 1 failed: APIConnectionError") as caught:
                    cli.llm_ocr(pdf, max_page_parallel=1)
            self.assertEqual(caught.exception.page_number, 1)
            self.assertEqual(caught.exception.request_attempts, cli.DEFAULT_VISION_REQUEST_ATTEMPTS)


if __name__ == "__main__":
    unittest.main()
