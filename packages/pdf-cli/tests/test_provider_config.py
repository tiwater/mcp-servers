import os
import unittest
from unittest.mock import patch

from tiwater_pdf import cli


class ProviderConfigTests(unittest.TestCase):
    def test_openrouter_key_selects_public_endpoint(self):
        with patch.dict(os.environ, {"OPENROUTER_API_KEY": "private-key"}, clear=True):
            self.assertEqual(
                cli._resolve_llm_config(),
                ("private-key", "https://openrouter.ai/api/v1"),
            )

    def test_standard_openai_environment_is_supported(self):
        values = {
            "OPENAI_API_KEY": "private-key",
            "OPENAI_BASE_URL": "https://provider.example/v1",
        }
        with patch.dict(os.environ, values, clear=True):
            self.assertEqual(
                cli._resolve_llm_config(),
                ("private-key", "https://provider.example/v1"),
            )

    def test_missing_provider_key_fails_closed(self):
        with patch.dict(os.environ, {}, clear=True):
            with self.assertRaisesRegex(RuntimeError, "OPENROUTER_API_KEY"):
                cli._resolve_llm_config()

    def test_default_ocr_model_and_openrouter_thinking_auto_disable(self):
        self.assertEqual(cli.DEFAULT_OCR_MODEL, "qwen/qwen3.7-plus")
        self.assertIs(
            cli._resolve_llm_enable_thinking(
                "auto",
                llm_model="qwen/qwen3.7-plus",
                base_url="https://openrouter.ai/api/v1",
            ),
            False,
        )
        self.assertIsNone(
            cli._resolve_llm_enable_thinking(
                "auto",
                llm_model="openai/gpt-5.6-luna",
                base_url="https://openrouter.ai/api/v1",
            )
        )
        self.assertEqual(
            cli._build_llm_extra_body(
                False,
                base_url="https://openrouter.ai/api/v1",
            ),
            {"enable_thinking": False, "reasoning": {"enabled": False}},
        )
        self.assertEqual(
            cli._build_llm_extra_body(
                False,
                base_url="https://dashscope.aliyuncs.com/compatible-mode/v1",
            ),
            {"enable_thinking": False},
        )

    def test_validate_vision_markdown_tables_ignores_empty_strings(self):
        valid_table = "| A | B |\n|---|---|\n| 1 | 2 |"
        self.assertEqual(
            cli._validate_vision_markdown_tables([valid_table, "", "  \n "]),
            [valid_table],
        )


if __name__ == "__main__":
    unittest.main()
