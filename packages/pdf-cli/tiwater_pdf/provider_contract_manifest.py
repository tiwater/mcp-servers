"""Set-15 ``lucid.provider-contract-manifest`` producer and independent validator.

The PDF provider is observation-only: the manifest declares exactly one observe
port bound to the real ``inspect-evidence-v2`` / ``validate-inspect-evidence-v2``
commands and their published v2 contract family. It never declares derive,
execute, render, or reobserve ports. ``manifestId`` is recomputed over the
provider/runtime/port declarations and the deployed bytes of every declared
contract schema, so the independent validator re-derives all of it from the
installed package rather than trusting the manifest under validation.
"""

from __future__ import annotations

import json
from pathlib import Path

from .format_evidence import _contract_path, atomic_write, canonical, file_sha, runtime_identity, sha

MANIFEST_SCHEMA_ID = "lucid.provider-contract-manifest"
VERDICT_SCHEMA_ID = "tiwater.provider-contract-manifest-verdict/v1"
PROVIDER_ID = "tiwater-pdf"
PORT_VALIDATOR_ID = "tiwater-pdf-validator"
MANIFEST_VALIDATOR_ID = "tiwater-pdf:provider-contract-manifest:validator"

REQUEST_SCHEMA_ID = "tiwater.format-evidence-request/v2"
RESULT_SCHEMA_ID = "tiwater.format-evidence/v2"
VERDICT_PORT_SCHEMA_ID = "tiwater.format-evidence-verdict/v2"
EXTRACTION_OPTIONS_SCHEMA_ID = "tiwater.format-extraction-options/v1"

# Declared contract family: every schema id referenced by the manifest must
# exist as deployed package bytes; manifestId binds their sha256.
CONTRACT_SCHEMAS = (
    ("tiwater.format-evidence-request-v2.schema.json", REQUEST_SCHEMA_ID),
    ("tiwater.format-evidence-v2.schema.json", RESULT_SCHEMA_ID),
    ("tiwater.format-evidence-verdict-v2.schema.json", VERDICT_PORT_SCHEMA_ID),
    ("tiwater.format-extraction-options-v1.schema.json", EXTRACTION_OPTIONS_SCHEMA_ID),
    ("tiwater.provider-contract-manifest-verdict-v1.schema.json", VERDICT_SCHEMA_ID),
)

OBSERVATION_ONLY_FORBIDDEN_KINDS = ("derive", "execute", "render", "reobserve")
MANIFEST_FIELDS = {"schema", "schemaSetVersion", "manifestId", "provider", "runtime", "ports"}


def _schema_binding(file: str, contract_id: str) -> dict:
    return {"id": contract_id, "sha256": file_sha(_contract_path(file))}


def build_manifest(schema_set_version: int, version: str) -> dict:
    """Build the deterministic Set-15 manifest for this package version."""
    if isinstance(schema_set_version, bool) or not isinstance(schema_set_version, int) or schema_set_version < 1:
        raise ValueError("--schema-set-version must be a positive integer")
    provider = {"id": PROVIDER_ID, "version": version}
    port_validator = {"id": PORT_VALIDATOR_ID, "version": version}
    port = {
        "kind": "observe",
        "producer": {**provider, "adapterIdentity": dict(provider)},
        "validator": {**port_validator, "adapterIdentity": dict(port_validator)},
        "requestSchema": REQUEST_SCHEMA_ID,
        "validatorRequestSchema": REQUEST_SCHEMA_ID,
        "resultSchema": RESULT_SCHEMA_ID,
        "verdictSchema": VERDICT_PORT_SCHEMA_ID,
        "options": [{"name": "facets", "valueSchema": EXTRACTION_OPTIONS_SCHEMA_ID}],
        "cacheKeyComposition": ["schemaSetVersion", "bytesSha256", "provider", "optionsSha256"],
        "resourceDeclarations": [{"resourceKey": "pdf:document", "access": "read"}],
        "sideEffect": {"kind": "read-only", "idempotent": True},
        "attemptBudget": 1,
    }
    manifest = {
        "schema": MANIFEST_SCHEMA_ID,
        "schemaSetVersion": schema_set_version,
        "provider": provider,
        "runtime": runtime_identity(version),
        "ports": [port],
    }
    material = {
        "contractSchemas": sorted(
            (_schema_binding(file, contract_id) for file, contract_id in CONTRACT_SCHEMAS),
            key=lambda binding: binding["id"],
        ),
        "manifest": manifest,
    }
    return {**manifest, "manifestId": f"manifest-{sha(canonical(material))}"}


def validate_manifest(manifest: dict, version: str) -> list[dict]:
    """Independently recompute every identity, hash, and port declaration."""
    findings: list[dict] = []

    def check(condition: bool, code: str, message: str) -> None:
        if not condition:
            findings.append({"code": code, "message": message})

    check(
        isinstance(manifest, dict) and set(manifest) == MANIFEST_FIELDS,
        "manifest-fields",
        "manifest must close over exactly schema, schemaSetVersion, manifestId, provider, runtime, ports",
    )
    if not isinstance(manifest, dict):
        return findings
    check(
        manifest.get("schema") == MANIFEST_SCHEMA_ID,
        "manifest-schema",
        "manifest schema identity is wrong",
    )
    schema_set_version = manifest.get("schemaSetVersion")
    check(
        isinstance(schema_set_version, int) and not isinstance(schema_set_version, bool) and schema_set_version >= 1,
        "schema-set-version",
        "schemaSetVersion must be a positive integer",
    )
    ports = manifest.get("ports")
    forbidden = [
        port.get("kind")
        for port in ports
        if isinstance(port, dict) and port.get("kind") in OBSERVATION_ONLY_FORBIDDEN_KINDS
    ] if isinstance(ports, list) else []
    check(
        not forbidden,
        "port-kind-forbidden",
        "PDF provider is observation-only: derive/execute/render/reobserve ports must never be declared",
    )
    if not findings:
        expected = build_manifest(schema_set_version, version)
        check(
            manifest.get("provider") == expected["provider"],
            "provider-identity-mismatch",
            "provider identity does not match the package",
        )
        check(
            manifest.get("runtime") == expected["runtime"],
            "runtime-identity-mismatch",
            "runtime identity does not match the package",
        )
        check(
            ports == expected["ports"],
            "port-declaration-mismatch",
            "observe port declaration does not match the package: expected exactly one read-only observe port "
            "with the deployed adapter identities, v2 schema family, options, cache composition, "
            "resource declarations, and attempt budget",
        )
        check(
            manifest.get("manifestId") == expected["manifestId"],
            "manifest-id-mismatch",
            "manifestId does not bind the package identities, port declarations, and deployed contract schema bytes",
        )
    return findings


def run_producer(output: Path, schema_set_version: int, version: str) -> int:
    if output.exists():
        raise ValueError("provider contract manifest output must be fresh")
    atomic_write(output, build_manifest(schema_set_version, version))
    return 0


def run_validator(manifest_path: Path, output: Path, version: str) -> int:
    if output.exists():
        raise ValueError("provider contract manifest verdict output must be fresh")
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    if not isinstance(manifest, dict):
        raise ValueError("provider contract manifest is not an object")
    findings = validate_manifest(manifest, version)
    verdict = {
        "schema": VERDICT_SCHEMA_ID,
        "manifestSha256": file_sha(manifest_path),
        "validator": {"id": MANIFEST_VALIDATOR_ID, "version": version},
        "decision": "pass" if not findings else "fail",
        "findings": findings,
    }
    atomic_write(output, verdict)
    return 0 if not findings else 1
