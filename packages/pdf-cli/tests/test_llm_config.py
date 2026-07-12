import os
import unittest
from unittest.mock import patch

from tiwater_pdf.cli import DEFAULT_LLM_TIMEOUT_SECONDS, DEFAULT_OCR_MODEL, _call_vision_with_retry, _resolve_llm_config, _resolve_llm_enable_thinking, llm_ocr


class ResolveLlmConfigTest(unittest.TestCase):
    def test_retries_transient_gateway_invalid_url_but_not_other_bad_requests(self):
        attempts = []

        def transient():
            attempts.append(1)
            if len(attempts) < 3:
                raise RuntimeError("invalid_parameter_error: provided URL does not appear to be valid")
            return "ok"

        result, count = _call_vision_with_retry(transient, sleep_fn=lambda _: None)
        self.assertEqual((result, count), ("ok", 3))
        with self.assertRaisesRegex(RuntimeError, "unsupported format"):
            _call_vision_with_retry(lambda: (_ for _ in ()).throw(RuntimeError("unsupported format")), sleep_fn=lambda _: None)

    def test_builtin_ocr_default_is_qwen37_plus(self):
        self.assertEqual(DEFAULT_OCR_MODEL, "qwen3.7-plus")
        self.assertEqual(llm_ocr.__defaults__[3], "qwen3.7-plus")

    def test_builtin_llm_timeout_covers_smai_gateway_latency(self):
        self.assertEqual(DEFAULT_LLM_TIMEOUT_SECONDS, 180.0)

    def test_uses_supen_gateway_env(self):
        env = {
            "SUPEN_LLM_TOKEN": "gateway-token",
            "SUPEN_LLM_GATEWAY_URL": "http://127.0.0.1:2755/api/llm/v1",
        }

        with patch.dict(os.environ, env, clear=True):
            api_key, base_url = _resolve_llm_config()

        self.assertEqual(api_key, "gateway-token")
        self.assertEqual(base_url, "http://127.0.0.1:2755/api/llm/v1")

    def test_accepts_supen_base_url_alias(self):
        env = {
            "SUPEN_LLM_API_KEY": "session-token",
            "SUPEN_LLM_BASE_URL": "http://127.0.0.1:2755/api/llm/v1",
        }

        with patch.dict(os.environ, env, clear=True):
            api_key, base_url = _resolve_llm_config()

        self.assertEqual(api_key, "session-token")
        self.assertEqual(base_url, "http://127.0.0.1:2755/api/llm/v1")

    def test_keeps_openrouter_default_only_for_openrouter_env(self):
        with patch.dict(os.environ, {"OPENROUTER_API_KEY": "openrouter-token"}, clear=True):
            api_key, base_url = _resolve_llm_config()

        self.assertEqual(api_key, "openrouter-token")
        self.assertEqual(base_url, "https://openrouter.ai/api/v1")

    def test_explicit_args_win_over_environment(self):
        env = {
            "SUPEN_LLM_TOKEN": "gateway-token",
            "SUPEN_LLM_GATEWAY_URL": "http://127.0.0.1:2755/api/llm/v1",
        }

        with patch.dict(os.environ, env, clear=True):
            api_key, base_url = _resolve_llm_config(
                api_key="explicit-token",
                base_url="https://llm.example/v1",
            )

        self.assertEqual(api_key, "explicit-token")
        self.assertEqual(base_url, "https://llm.example/v1")

    def test_auto_disables_thinking_for_aliyun_qwen37(self):
        value = _resolve_llm_enable_thinking(
            "auto",
            llm_model="qwen3.7-plus",
            base_url="https://example.cn-beijing.maas.aliyuncs.com/compatible-mode/v1",
        )

        self.assertIs(value, False)

    def test_auto_disables_thinking_for_bare_aliyun_qwen37_behind_gateway(self):
        value = _resolve_llm_enable_thinking(
            "auto",
            llm_model="qwen3.7-plus",
            base_url="https://hub.supen.ai/api/llm/v1",
        )

        self.assertIs(value, False)

    def test_auto_does_not_send_vendor_parameter_for_other_models(self):
        value = _resolve_llm_enable_thinking(
            "auto",
            llm_model="gpt-4o-mini",
            base_url="https://api.openai.com/v1",
        )

        self.assertIsNone(value)

    def test_auto_does_not_treat_openrouter_owner_prefixed_qwen_as_aliyun(self):
        value = _resolve_llm_enable_thinking(
            "auto",
            llm_model="qwen/qwen3.7-plus",
            base_url="https://openrouter.ai/api/v1",
        )

        self.assertIsNone(value)

    def test_explicit_enable_thinking_override(self):
        value = _resolve_llm_enable_thinking(
            "true",
            llm_model="qwen3.7-plus",
            base_url="https://example.cn-beijing.maas.aliyuncs.com/compatible-mode/v1",
        )

        self.assertIs(value, True)


if __name__ == "__main__":
    unittest.main()
