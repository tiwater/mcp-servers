import os
from unittest.mock import patch

import pytest

from tiwater_pdf import cli


def test_openrouter_key_selects_public_endpoint():
    with patch.dict(os.environ, {"OPENROUTER_API_KEY": "private-key"}, clear=True):
        assert cli._resolve_llm_config() == ("private-key", "https://openrouter.ai/api/v1")


def test_standard_openai_environment_is_supported():
    values = {"OPENAI_API_KEY": "private-key", "OPENAI_BASE_URL": "https://provider.example/v1"}
    with patch.dict(os.environ, values, clear=True):
        assert cli._resolve_llm_config() == ("private-key", "https://provider.example/v1")


def test_missing_provider_key_fails_closed():
    with patch.dict(os.environ, {}, clear=True), pytest.raises(RuntimeError, match="OPENROUTER_API_KEY"):
        cli._resolve_llm_config()
