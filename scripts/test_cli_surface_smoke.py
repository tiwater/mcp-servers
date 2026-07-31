#!/usr/bin/env python3
"""Minimal proofs for the CLI public-surface publish smoke gate."""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from cli_surface_smoke import (
    evaluate_surfaces,
    find_forbidden,
    has_command,
    has_render_surface,
)


def valid_surfaces() -> dict[str, str]:
    return {
        "tiwater-docx": "Usage:\n  inspect <input.docx>\n  edit <input.docx> <ops.json> <out.docx>\n",
        "tiwater-xlsx": "Usage:\n  inspect <input.xlsx>\n  edit <input.xlsx> <ops.json> <out.xlsx>\n",
        "tiwater-pptx": "Usage:\n  inspect <input.pptx>\n  apply-format-edits <input.pptx> <plan.json> <out.pptx>\n",
        "tiwater-convert": (
            "Usage:\n"
            "  tiwater-convert xls-to-xlsx <in.xls> <out.xlsx>\n"
            "  tiwater-convert docx-to-pdf <in.docx> <out.pdf>\n"
        ),
        "tiwater-pdf": "usage: tiwater-pdf [-h] {inspect,ocr} ...\n  inspect\n  ocr\n",
    }


class CliSurfaceSmokeTests(unittest.TestCase):
    def test_forbidden_orchestration_term_fails(self) -> None:
        surfaces = valid_surfaces()
        surfaces["tiwater-docx"] += "  provider-contract-manifest --output x.json\n"
        failures = evaluate_surfaces(surfaces)
        self.assertTrue(any("provider-contract-manifest" in item for item in failures))

    def test_forbidden_framework_word_fails(self) -> None:
        surfaces = valid_surfaces()
        surfaces["tiwater-pdf"] += "powered by lucid runtime\n"
        failures = evaluate_surfaces(surfaces)
        self.assertTrue(any("lucid" in item for item in failures))

    def test_missing_inspect_fails(self) -> None:
        surfaces = valid_surfaces()
        surfaces["tiwater-xlsx"] = "Usage:\n  edit <input.xlsx> <ops.json> <out.xlsx>\n"
        failures = evaluate_surfaces(surfaces)
        self.assertTrue(any("missing required command 'inspect'" in item for item in failures))

    def test_valid_surfaces_pass(self) -> None:
        self.assertEqual(evaluate_surfaces(valid_surfaces()), [])

    def test_helpers(self) -> None:
        self.assertEqual(find_forbidden("derive-operation now"), ["derive-operation"])
        self.assertEqual(find_forbidden("derive-template-migration"), [])
        self.assertEqual(find_forbidden("format-evidence blob"), ["format-evidence"])
        self.assertEqual(find_forbidden("inspect-evidence blob"), ["inspect-evidence"])
        self.assertEqual(find_forbidden("execute-effect x"), ["execute-effect"])
        self.assertEqual(find_forbidden("validate-effect x"), ["validate-effect"])
        self.assertEqual(find_forbidden("Schema Set 15"), ["schema set"])
        self.assertEqual(find_forbidden("workflow step"), ["workflow"])
        self.assertEqual(find_forbidden("scenario id"), ["scenario"])
        self.assertEqual(find_forbidden("conformance suite"), ["conformance"])
        self.assertEqual(find_forbidden("release gate"), ["release"])
        self.assertTrue(has_command("  inspect <input.docx>\n", "inspect"))
        self.assertTrue(has_render_surface("docx-to-pdf <in> <out>"))


if __name__ == "__main__":
    unittest.main()
