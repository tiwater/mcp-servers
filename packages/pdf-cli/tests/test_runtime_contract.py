import json
import tempfile
import unittest
from pathlib import Path

from tiwater_pdf.runtime_contract import (
    canonical_json_bytes,
    derived_identity,
    identify_canonical_json_artifact,
    identify_file,
    native_identity,
    normalize_evidence,
)


class RuntimeContractTest(unittest.TestCase):
    def test_file_identity_binds_path_size_hash_and_content_id(self):
        with tempfile.TemporaryDirectory() as temp:
            source = Path(temp) / "source.bin"
            source.write_bytes(b"abc")

            identity = identify_file(source)

        self.assertEqual(identity["path"], str(source.resolve()))
        self.assertEqual(identity["sizeBytes"], 3)
        self.assertEqual(
            identity["sha256"],
            "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad",
        )
        self.assertEqual(identity["contentId"], f"sha256:{identity['sha256']}")

    def test_canonical_artifact_is_independent_of_object_key_order(self):
        first = identify_canonical_json_artifact(
            {"b": 2, "a": 1},
            schema_id="tiwater.test-payload",
            schema_version="1.0.0",
        )
        second = identify_canonical_json_artifact(
            {"a": 1, "b": 2},
            schema_id="tiwater.test-payload",
            schema_version="1.0.0",
        )

        self.assertEqual(first, second)
        self.assertEqual(canonical_json_bytes({"b": 2, "a": 1}), b'{"a":1,"b":2}')
        self.assertEqual(first["artifactId"], f"sha256:{first['sha256']}")

    def test_canonical_json_matches_the_cross_language_utf8_fixture(self):
        value = {"text": "中文", "items": [2, 1], "nested": {"b": True, "a": None}}

        payload = canonical_json_bytes(value)

        self.assertEqual(
            payload,
            '{"items":[2,1],"nested":{"a":null,"b":true},"text":"中文"}'.encode(),
        )

    def test_canonical_json_matches_shared_adversarial_vectors(self):
        fixture_path = (
            Path(__file__).resolve().parents[3]
            / "contracts"
            / "runtime"
            / "fixtures"
            / "canonical-json-vectors.json"
        )
        vectors = json.loads(fixture_path.read_text(encoding="utf-8"))["vectors"]

        for vector in vectors:
            with self.subTest(vector=vector["name"]):
                self.assertEqual(
                    canonical_json_bytes(vector["value"]),
                    vector["canonical"].encode(),
                )

    def test_canonical_json_rejects_non_integer_numbers(self):
        with self.assertRaisesRegex(ValueError, "integer"):
            canonical_json_bytes({"value": 1.5})

    def test_canonical_json_rejects_shared_lossy_numeric_vectors(self):
        fixture_path = (
            Path(__file__).resolve().parents[3]
            / "contracts"
            / "runtime"
            / "fixtures"
            / "canonical-json-negative-vectors.json"
        )
        vectors = json.loads(fixture_path.read_text(encoding="utf-8"))["vectors"]

        for vector in vectors:
            with self.subTest(vector=vector["name"]):
                with self.assertRaisesRegex(ValueError, "integer"):
                    canonical_json_bytes(json.loads(vector["json"]))

    def test_native_and_derived_identity_are_disjoint(self):
        native = native_identity("pdf-object-xref", "42")
        derived = derived_identity(
            "source-version-structural-locator",
            ["sha256:source", "page[0]", "table[0]"],
        )

        self.assertEqual(native["kind"], "native")
        self.assertEqual(native["nativeId"], "42")
        self.assertNotIn("derivation", native)
        self.assertEqual(derived["kind"], "derived")
        self.assertNotIn("nativeId", derived)
        self.assertEqual(derived["inputs"], ["sha256:source", "page[0]", "table[0]"])

    def test_identity_helpers_reject_empty_contract_fields(self):
        with self.assertRaisesRegex(ValueError, "namespace"):
            native_identity("", "42")
        with self.assertRaisesRegex(ValueError, "inputs"):
            derived_identity("source-version-structural-locator", [])

    def test_normalized_evidence_has_stable_nodes_and_containment(self):
        result = normalize_evidence(
            {
                "file": "/renamed/source.pdf",
                "tables": [{"rows": [{"cells": [{"text": "结果", "confidence": 0.95}]}]}],
            }
        )

        nodes = {node["runtimeNodeId"]: node for node in result["payload"]["nodes"]}
        self.assertEqual(nodes["/tables/0/rows/0/cells/0/text"]["value"], "结果")
        self.assertEqual(nodes["/tables/0/rows/0/cells/0/confidence"]["value"], "0.95")
        self.assertEqual(
            nodes["/tables/0/rows/0/cells/0/text"]["containedBy"],
            "/tables/0/rows/0/cells/0",
        )
        self.assertNotIn("/file", nodes)
        object_ids = {item["objectId"] for item in result["objects"]}
        self.assertEqual(len(object_ids), len(result["objects"]))
        self.assertTrue(all(item["root"] or item["parentObjectId"] in object_ids for item in result["objects"]))
        canonical_json_bytes(result["payload"])


if __name__ == "__main__":
    unittest.main()
