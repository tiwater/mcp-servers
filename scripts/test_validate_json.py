#!/usr/bin/env python3
"""Minimal proofs for the strict JSON publish gate."""

from __future__ import annotations

import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

SCRIPT = Path(__file__).resolve().with_name("validate_json.py")


def run_gate(*args: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, str(SCRIPT), *args],
        capture_output=True,
        text=True,
    )


class ValidateJsonGateTests(unittest.TestCase):
    def test_trailing_comma_fails_the_gate(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            bad = root / "broken.json"
            bad.write_text('{"ok": true,}\n', encoding="utf-8")
            result = run_gate(str(bad))
            self.assertNotEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertIn("invalid JSON", result.stderr)

    def test_valid_json_passes_the_gate(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            good = root / "ok.json"
            good.write_text('{"ok": true}\n', encoding="utf-8")
            result = run_gate(str(good))
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_trailing_comma_inside_packed_archive_fails(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            archive = root / "sample.nupkg"
            with zipfile.ZipFile(archive, "w") as packed:
                packed.writestr("contracts/broken.json", '{"items": [1,],}\n')
            result = run_gate("--archives-from", str(root))
            self.assertNotEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertIn("invalid JSON", result.stderr)


if __name__ == "__main__":
    unittest.main()
