"""Closed published format-evidence producer and independent validator."""

from __future__ import annotations

import hashlib
import json
import math
import os
from pathlib import Path
import tempfile
from typing import Callable


def canonical(value) -> str:
    if value is None:
        return "null"
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, str):
        return json.dumps(value, ensure_ascii=False, separators=(",", ":"))
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if not math.isfinite(value):
            raise ValueError("canonical number is non-finite")
        if value == 0:
            return "0"
        if value.is_integer():
            return str(int(value))
        return json.dumps(value, allow_nan=False, separators=(",", ":")).replace("e-0", "e-").replace("e+0", "e+")
    if isinstance(value, list):
        return "[" + ",".join(canonical(item) for item in value) + "]"
    if isinstance(value, dict):
        return "{" + ",".join(f"{canonical(key)}:{canonical(value[key])}" for key in sorted(value)) + "}"
    raise TypeError(f"unsupported canonical type: {type(value).__name__}")


def sha(value) -> str:
    data = value if isinstance(value, str) else canonical(value)
    return hashlib.sha256(data.encode("utf-8")).hexdigest()


def file_sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def atomic_write(path: Path, value: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary = tempfile.mkstemp(prefix=f".{path.name}.", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8") as stream:
            stream.write(canonical(value) + "\n")
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
        directory = os.open(path.parent, os.O_RDONLY)
        try:
            os.fsync(directory)
        finally:
            os.close(directory)
    except Exception:
        Path(temporary).unlink(missing_ok=True)
        raise


def validate_request(request: dict, output: Path, *, validator: bool) -> None:
    fields = {"schema", "requestId", "runId", "subject", "artifact", "extraction", "expectedEvidenceSchema", "outputPath"}
    if set(request) != fields or request["schema"] != "tiwater.format-evidence-request/v1" or request["expectedEvidenceSchema"] != "lucid.published-format-evidence/v1":
        raise ValueError("request contract invalid")
    if not validator and Path(request["outputPath"]).resolve() != output.resolve():
        raise ValueError("output path mismatch")
    artifact = request["artifact"]
    source = Path(artifact["path"])
    if not source.is_absolute() or artifact["format"] != "pdf" or file_sha(source) != artifact["bytesSha256"]:
        raise ValueError("artifact authority mismatch")
    extraction = request["extraction"]
    if sha(extraction["options"]) != extraction["optionsSha256"]:
        raise ValueError("extraction options mismatch")


def build_evidence(request: dict, inspect: Callable[[Path], object], version: str) -> dict:
    artifact, extraction = request["artifact"], request["extraction"]
    epoch_material = {"bytesSha256": artifact["bytesSha256"], "runtimeTool": "tiwater-pdf", "runtimeSchema": extraction["schema"], "runtimeVersion": version, "extractionOptions": extraction["options"]}
    evidence = {
        "schema": "lucid.published-format-evidence/v1", "requestId": request["requestId"], "subject": request["subject"], "artifactVersionId": artifact["artifactVersionId"],
        "provider": {"tool": "tiwater-pdf", "toolVersion": version, "capabilityId": "inspect-evidence", "capabilityVersion": "1", "outputSchema": "lucid.published-format-evidence/v1"},
        "source": {"bytesSha256": artifact["bytesSha256"], "format": "pdf"}, "extraction": extraction,
        "epoch": {"epochId": f"ep-{sha(epoch_material)}", "bytesSha256": artifact["bytesSha256"], "runtimeTool": "tiwater-pdf", "runtimeSchema": extraction["schema"], "runtimeVersion": version, "extractionOptionsSha256": extraction["optionsSha256"]},
        "entities": [{"entityId": "document-1", "kind": "pdf-document", "provenance": {"source": "runtime", "pointer": "/inspection"}}],
        "observations": [{"observationId": "inspection-1", "entityId": "document-1", "semanticField": "pdf.inspection", "use": "structure", "value": inspect(Path(artifact["path"])), "parentObservationIds": [], "provenance": {"source": "runtime", "pointer": "/inspection"}}],
    }
    evidence["evidenceId"] = f"evidence-{sha(evidence)}"
    return evidence


def run(request_path: Path, output: Path, inspect: Callable[[Path], object], version: str, evidence_path: Path | None = None) -> int:
    request = json.loads(request_path.read_text(encoding="utf-8"))
    validator = evidence_path is not None
    try:
        validate_request(request, output, validator=validator)
        expected = build_evidence(request, inspect, version)
        if not validator:
            result = expected
        else:
            evidence = json.loads(evidence_path.read_text(encoding="utf-8"))
            passed = canonical(evidence) == canonical(expected)
            result = {"schema": "lucid.published-format-evidence-verdict/v1", "requestId": request["requestId"], "subject": request["subject"], "artifactVersionId": request["artifact"]["artifactVersionId"], "epochId": expected["epoch"]["epochId"], "evidenceRef": {"evidenceId": evidence.get("evidenceId", "missing"), "sha256": sha(evidence)}, "validator": {"tool": "tiwater-pdf", "toolVersion": version, "capabilityId": "validate-inspect-evidence", "capabilityVersion": "1"}, "recomputedSemanticHash": sha({"entities": expected["entities"], "observations": expected["observations"]}), "pass": passed, "findings": [] if passed else [{"code": "inspect-evidence-recomputation-mismatch", "owner": "validator"}]}
    except Exception:
        artifact = request.get("artifact", {})
        result = {"schema": "tiwater.format-evidence-error/v1", "requestId": request.get("requestId", "unknown"), "subject": request.get("subject", {"kind": "input", "inputId": "unknown"}), "artifactVersionId": artifact.get("artifactVersionId", "unknown"), "code": "inspect-evidence-invalid", "category": "evidence", "retryable": False, "provider": {"tool": "tiwater-pdf", "toolVersion": version, "capabilityId": "validate-inspect-evidence" if validator else "inspect-evidence"}, "refs": []}
    atomic_write(output, result)
    return 0
