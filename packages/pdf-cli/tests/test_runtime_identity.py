import hashlib
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import fitz

from tiwater_pdf.runtime_identity import capabilities, identify


class RuntimeIdentityTest(unittest.TestCase):
    def test_capabilities_describe_only_non_mutating_pdf_identity_commands(self):
        descriptor = capabilities()

        self.assertEqual(descriptor["schemaVersion"], "1.0.0")
        self.assertEqual(descriptor["package"], {"name": "tiwater-pdf", "version": "0.16.0"})
        self.assertEqual(
            descriptor["runtime"],
            {"family": "pdf", "name": "tiwater-pdf", "version": "0.16.0"},
        )
        self.assertEqual(
            descriptor["supportedKinds"],
            [
                {
                    "fileKind": "pdf",
                    "mediaTypes": ["application/pdf"],
                    "signatureKinds": ["pdf-header-pymupdf-open"],
                }
            ],
        )
        self.assertTrue(all(not command["mutates"] for command in descriptor["commands"]))

    def test_identify_reads_once_and_hashes_the_exact_bytes_opened_by_pymupdf(self):
        with tempfile.TemporaryDirectory() as temporary:
            source = Path(temporary) / "renamed.payload"
            self._create_pdf(source)
            source_bytes = source.read_bytes()
            original_read_bytes = Path.read_bytes
            reads = []

            def counted_read_bytes(path):
                reads.append(path)
                return original_read_bytes(path)

            with patch.object(Path, "read_bytes", counted_read_bytes):
                evidence = identify(source)

        self.assertEqual(reads, [source.resolve()])
        self.assertEqual(evidence["status"], "supported")
        self.assertEqual(evidence["source"]["sizeBytes"], len(source_bytes))
        self.assertEqual(evidence["source"]["sha256"], hashlib.sha256(source_bytes).hexdigest())
        self.assertEqual(evidence["file"]["signature"]["status"], "matched")

    def test_identify_distinguishes_unsupported_content_from_source_read_failure(self):
        with tempfile.TemporaryDirectory() as temporary:
            fake = Path(temporary) / "fake.pdf"
            fake.write_bytes(b"not a PDF")
            missing = Path(temporary) / "missing.pdf"

            unsupported = identify(fake)
            failed = identify(missing)

        self.assertEqual(unsupported["status"], "unsupported")
        self.assertIsNotNone(unsupported["source"])
        self.assertIsNone(unsupported["failureStage"])
        self.assertEqual(unsupported["errors"], [])
        self.assertEqual(failed["status"], "failed")
        self.assertEqual(failed["failureStage"], "source-read")
        self.assertIsNone(failed["source"])
        self.assertEqual(failed["file"]["signature"]["status"], "not-checked")

    def test_identify_supports_encrypted_pdf_identity_without_authentication(self):
        with tempfile.TemporaryDirectory() as temporary:
            source = Path(temporary) / "encrypted.bin"
            self._create_pdf(source, encrypted=True)

            evidence = identify(source)

        self.assertEqual(evidence["status"], "supported")
        self.assertIs(evidence["payload"]["encrypted"], True)
        self.assertIn("pdf:encrypted=true", evidence["file"]["signature"]["evidence"])

    @staticmethod
    def _create_pdf(path: Path, *, encrypted: bool = False) -> None:
        document = fitz.open()
        document.new_page().insert_text((72, 72), "runtime identity fixture")
        options = {}
        if encrypted:
            options = {
                "encryption": fitz.PDF_ENCRYPT_AES_256,
                "owner_pw": "owner-secret",
                "user_pw": "user-secret",
            }
        document.save(path, **options)
        document.close()


if __name__ == "__main__":
    unittest.main()
