import os
import unittest
from unittest.mock import patch

from tiwater_pdf.cli import _resolve_llm_config


class ResolveLlmConfigTest(unittest.TestCase):
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


if __name__ == "__main__":
    unittest.main()
