"""Non-mutating runtime capability and PDF identity evidence probes."""

from __future__ import annotations

import hashlib
from pathlib import Path
import re
from typing import Any

import fitz

from . import __version__
from .runtime_contract import identify_canonical_json_artifact, normalize_evidence


CONTRACT_VERSION = "1.0.0"
CAPABILITIES_SCHEMA_ID = (
    "https://tiwater.dev/contracts/runtime/runtime-capabilities.schema.json"
)
EVIDENCE_SCHEMA_ID = (
    "https://tiwater.dev/contracts/runtime/runtime-evidence-envelope.schema.json"
)
IDENTIFY_PAYLOAD_SCHEMA_ID = "tiwater.runtime.identify-payload"
PACKAGE_NAME = "tiwater-pdf"
RUNTIME_NAME = "tiwater-pdf"
SIGNATURE_KIND = "pdf-header-pymupdf-open"
PDF_MEDIA_TYPE = "application/pdf"
_PDF_HEADER = re.compile(rb"\A%PDF-(?P<version>(?:1\.[0-7]|2\.0))(?:\r\n|\r|\n)")


def capabilities() -> dict[str, Any]:
    """Return the versioned, non-mutating PDF runtime descriptor."""

    return {
        "schemaVersion": CONTRACT_VERSION,
        "descriptorType": "runtime-capabilities",
        "package": _package_identity(),
        "runtime": _runtime_identity(),
        "evidenceSchema": _schema_identity(EVIDENCE_SCHEMA_ID),
        "descriptorCommand": {
            "command": "capabilities",
            "arguments": ["--json"],
            "mutates": False,
        },
        "identifyProbe": {
            "command": "identify",
            "arguments": ["<input>", "--json"],
            "mutates": False,
            "outcomes": ["supported", "unsupported", "failed"],
        },
        "supportedKinds": [
            {
                "fileKind": "pdf",
                "mediaTypes": [PDF_MEDIA_TYPE],
                "signatureKinds": [SIGNATURE_KIND],
            }
        ],
        "commands": [
            {
                "name": "capabilities",
                "mutates": False,
                "outputSchema": _schema_identity(CAPABILITIES_SCHEMA_ID),
            },
            {
                "name": "identify",
                "mutates": False,
                "outputSchema": _schema_identity(EVIDENCE_SCHEMA_ID),
            },
            {
                "name": "extract-evidence",
                "mutates": False,
                "outputSchema": _schema_identity(EVIDENCE_SCHEMA_ID),
            },
        ],
        "identityPolicy": {
            "nativeIds": "runtime-native-only",
            "derivedIds": "deterministic-and-explicit",
            "containment": "parent-object-id-required-for-non-root",
        },
    }


def identify(file_path: str | Path) -> dict[str, Any]:
    """Identify PDF bytes without OCR, network access, or source mutation."""

    try:
        resolved_path = Path(file_path).expanduser().resolve(strict=True)
        source_bytes = resolved_path.read_bytes()
    except (OSError, RuntimeError, ValueError):
        return _evidence_envelope(
            status="failed",
            failure_stage="source-read",
            source=None,
            file_evidence=_file_evidence(
                signature_status="not-checked",
                signature_evidence=[],
            ),
            payload={"failureClass": "source-read-error"},
            errors=[
                {
                    "code": "source-read-failed",
                    "message": "The source bytes could not be read.",
                }
            ],
        )

    source = _source_identity(resolved_path, source_bytes)
    header = _PDF_HEADER.match(source_bytes)
    if header is None:
        return _unsupported(
            source,
            reason="pdf-header-mismatch",
            evidence=["pdf-header:mismatched"],
        )

    version = header.group("version").decode("ascii")
    try:
        with fitz.open(stream=source_bytes, filetype="pdf") as document:
            if not document.is_pdf:
                return _unsupported(
                    source,
                    reason="pymupdf-not-pdf",
                    evidence=[f"pdf-header:version={version}", "pymupdf:not-pdf"],
                )
            if document.is_repaired:
                return _unsupported(
                    source,
                    reason="pymupdf-repair-required",
                    evidence=[
                        f"pdf-header:version={version}",
                        "pymupdf:pdf-repair-required",
                    ],
                )
            encrypted = bool(document.needs_pass or document.is_encrypted)
    except (fitz.FileDataError, RuntimeError, TypeError, ValueError):
        return _unsupported(
            source,
            reason="pymupdf-open-rejected",
            evidence=[f"pdf-header:version={version}", "pymupdf:pdf-rejected"],
        )

    return _evidence_envelope(
        status="supported",
        failure_stage=None,
        source=source,
        file_evidence=_file_evidence(
            file_kind="pdf",
            media_type=PDF_MEDIA_TYPE,
            signature_status="matched",
            signature_evidence=[
                f"pdf-header:version={version}",
                "pymupdf:pdf-opened",
                f"pdf:encrypted={str(encrypted).lower()}",
            ],
        ),
        payload={"recognized": True, "fileKind": "pdf", "encrypted": encrypted},
        errors=[],
    )


def extraction_evidence(
    identify_evidence: dict[str, Any],
    report: dict[str, Any] | None,
) -> dict[str, Any]:
    """Bind normalized extraction nodes to a previously computed identity probe."""

    if identify_evidence["status"] == "supported":
        normalized = normalize_evidence(report or {})
        payload = normalized["payload"]
        objects = normalized["objects"]
    else:
        payload = {
            "identifyStatus": identify_evidence["status"],
            "reason": "source-not-supported-for-extraction",
            "nodes": [],
        }
        objects = []

    return _evidence_envelope(
        status=identify_evidence["status"],
        failure_stage=identify_evidence["failureStage"],
        source=identify_evidence["source"],
        file_evidence=identify_evidence["file"],
        payload=payload,
        errors=identify_evidence["errors"],
        probe="extract-evidence",
        objects=objects,
        payload_schema_id="tiwater.runtime.normalized-evidence",
    )


def _unsupported(
    source: dict[str, Any],
    *,
    reason: str,
    evidence: list[str],
) -> dict[str, Any]:
    return _evidence_envelope(
        status="unsupported",
        failure_stage=None,
        source=source,
        file_evidence=_file_evidence(
            signature_status="mismatched",
            signature_evidence=evidence,
        ),
        payload={"recognized": False, "reason": reason},
        errors=[],
    )


def _evidence_envelope(
    *,
    status: str,
    failure_stage: str | None,
    source: dict[str, Any] | None,
    file_evidence: dict[str, Any],
    payload: dict[str, Any],
    errors: list[dict[str, Any]],
    probe: str = "identify",
    objects: list[dict[str, Any]] | None = None,
    payload_schema_id: str = IDENTIFY_PAYLOAD_SCHEMA_ID,
) -> dict[str, Any]:
    return {
        "schemaVersion": CONTRACT_VERSION,
        "envelopeType": "runtime-evidence",
        "probe": probe,
        "status": status,
        "failureStage": failure_stage,
        "package": _package_identity(),
        "runtime": _runtime_identity(),
        "evidenceSchema": _schema_identity(EVIDENCE_SCHEMA_ID),
        "source": source,
        "file": file_evidence,
        "artifact": identify_canonical_json_artifact(
            payload,
            schema_id=payload_schema_id,
            schema_version=CONTRACT_VERSION,
        ),
        "payload": payload,
        "objects": objects or [],
        "warnings": [],
        "errors": errors,
    }


def _source_identity(resolved_path: Path, source_bytes: bytes) -> dict[str, Any]:
    sha256 = hashlib.sha256(source_bytes).hexdigest()
    return {
        "path": str(resolved_path),
        "sizeBytes": len(source_bytes),
        "sha256": sha256,
        "contentId": f"sha256:{sha256}",
    }


def _file_evidence(
    *,
    signature_status: str,
    signature_evidence: list[str],
    file_kind: str | None = None,
    media_type: str | None = None,
) -> dict[str, Any]:
    return {
        "fileKind": file_kind,
        "mediaType": media_type,
        "signature": {
            "status": signature_status,
            "kind": SIGNATURE_KIND,
            "evidence": signature_evidence,
        },
    }


def _package_identity() -> dict[str, str]:
    return {"name": PACKAGE_NAME, "version": __version__}


def _runtime_identity() -> dict[str, str]:
    return {"family": "pdf", "name": RUNTIME_NAME, "version": __version__}


def _schema_identity(schema_id: str) -> dict[str, str]:
    return {"id": schema_id, "version": CONTRACT_VERSION}
