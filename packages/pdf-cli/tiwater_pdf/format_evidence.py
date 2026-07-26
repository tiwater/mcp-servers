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


def _contract_path(file: str) -> Path:
    packaged = Path(__file__).with_name("contracts") / file
    if packaged.is_file():
        return packaged
    development = Path(__file__).resolve().parents[2] / "format-evidence" / "contracts" / file
    if development.is_file():
        return development
    raise ValueError(f"provider contract missing: {file}")


def _contract_ref(file: str, contract_id: str) -> dict:
    return {"id": contract_id, "sha256": file_sha(_contract_path(file))}


def _typed_value(file: str, contract_id: str, value: object) -> dict:
    return {"schema": _contract_ref(file, contract_id), "value": value, "sha256": sha(canonical(value))}


def _scalar_kind(value: object) -> str:
    if value is None:
        return "null"
    if isinstance(value, str):
        return "string"
    if isinstance(value, bool):
        return "boolean"
    return "number"


def _scalar_fields(value: object) -> list[dict]:
    if isinstance(value, dict):
        return [
            {"name": name, "kind": _scalar_kind(item), "value": item, "sha256": sha(canonical(item))}
            for name, item in sorted(value.items())
            if not isinstance(item, (dict, list))
        ]
    if isinstance(value, list):
        return [{"name": "length", "kind": "number", "value": len(value), "sha256": sha(canonical(len(value)))}]
    return [{"name": "value", "kind": _scalar_kind(value), "value": value, "sha256": sha(canonical(value))}]


def _escape_pointer(value: str) -> str:
    return value.replace("~", "~0").replace("/", "~1")


def _inventory_candidates(
    inspection: object,
    artifact_version_id: str,
    provider: dict,
    inspection_sha256: str,
) -> list[dict]:
    candidates: list[dict] = []

    def walk(value: object, pointer: str) -> None:
        candidate_value_sha256 = sha(canonical(value))
        material = {
            "artifactVersionId": artifact_version_id,
            "provider": provider,
            "inspectionSha256": inspection_sha256,
            "pointer": pointer,
        }
        candidates.append({
            "candidateId": f"candidate-{sha(material)}",
            "candidateKind": "object" if isinstance(value, dict) else "array" if isinstance(value, list) else "scalar",
            "pointer": pointer,
            "fields": _scalar_fields(value),
            "candidateValueSha256": candidate_value_sha256,
            "dispositionInput": "available",
        })
        if isinstance(value, dict):
            for name, child in sorted(value.items()):
                walk(child, f"{pointer}/{_escape_pointer(name)}")
        elif isinstance(value, list):
            for index, child in enumerate(value):
                walk(child, f"{pointer}/{index}")

    walk(inspection, "")
    return sorted(candidates, key=lambda item: item["pointer"])


def _validate_request_v2(request: dict, version: str) -> None:
    required = {
        "schema", "requestId", "runId", "subject", "artifact", "provider",
        "validator", "runtime", "extraction", "expectedEvidenceContract",
    }
    if set(request) != required or request.get("schema") != "tiwater.format-evidence-request/v2":
        raise ValueError("v2 request contract invalid")
    provider = {"id": "tiwater-pdf", "version": version}
    validator = {"id": "tiwater-pdf-validator", "version": version}
    if request["provider"] != provider or request["validator"] != validator or request["runtime"] != provider:
        raise ValueError("v2 provider identity mismatch")
    expected_evidence = _contract_ref("tiwater.format-evidence-v2.schema.json", "tiwater.format-evidence/v2")
    if request["expectedEvidenceContract"] != expected_evidence:
        raise ValueError("v2 evidence contract mismatch")
    extraction = request["extraction"]
    expected_extraction = _contract_ref("tiwater.format-extraction-options-v1.schema.json", "tiwater.format-extraction-options/v1")
    if extraction.get("schema") != expected_extraction or extraction.get("sha256") != sha(extraction.get("value")):
        raise ValueError("v2 extraction authority mismatch")
    if extraction.get("value") != {"facets": ["format-summary"]}:
        raise ValueError("v2 extraction options invalid")
    artifact = request["artifact"]
    source = Path(artifact["path"])
    if not source.is_absolute() or source.suffix.lower() != ".pdf" or file_sha(source) != artifact["bytesSha256"]:
        raise ValueError("v2 artifact authority mismatch")


def _build_evidence_v2(request: dict, inspect: Callable[[Path], object]) -> dict:
    artifact = request["artifact"]
    inspection = inspect(Path(artifact["path"]))
    inspection_sha256 = sha(canonical(inspection))
    if isinstance(inspection, dict):
        facets = [{"facetId": name, "sha256": sha(canonical(value))} for name, value in sorted(inspection.items())]
    else:
        facets = [{"facetId": "inspection", "sha256": inspection_sha256}]
    observation_schema = _contract_ref(
        "tiwater.provider-document-observation-v2.schema.json",
        "tiwater.provider-document-observation/v2",
    )
    epoch_material = {
        "sourceBytesSha256": artifact["bytesSha256"],
        "provider": request["provider"],
        "runtime": request["runtime"],
        "extractionSha256": request["extraction"]["sha256"],
        "observationSchema": observation_schema,
    }
    epoch_id = f"epoch-{sha(epoch_material)}"
    inventory_base = {
        "artifactVersionId": artifact["artifactVersionId"],
        "inspectionSha256": inspection_sha256,
        "candidates": _inventory_candidates(
            inspection,
            artifact["artifactVersionId"],
            request["provider"],
            inspection_sha256,
        ),
    }
    inventory_sha256 = sha(inventory_base)
    target_base = {
        "artifactVersionId": artifact["artifactVersionId"],
        "epochId": epoch_id,
        "inspectionSha256": inspection_sha256,
        "candidates": [],
    }
    target_sha256 = sha(target_base)
    observation_value = {
        "format": "pdf",
        "artifactVersionId": artifact["artifactVersionId"],
        "epochId": epoch_id,
        "inspectionSha256": inspection_sha256,
        "facets": facets,
        "inventoryUniverse": {
            "universeId": f"inventory-{inventory_sha256}",
            **inventory_base,
            "universeSha256": inventory_sha256,
        },
        "targetUniverse": {
            "universeId": f"targets-{target_sha256}",
            **target_base,
            "universeSha256": target_sha256,
        },
    }
    provenance_value = {
        "kind": "provider-inspection",
        "artifactVersionId": artifact["artifactVersionId"],
        "sourceBytesSha256": artifact["bytesSha256"],
        "inspectionSha256": inspection_sha256,
        "provider": request["provider"],
        "runtime": request["runtime"],
        "extractionSha256": request["extraction"]["sha256"],
    }
    evidence = {
        "schema": "tiwater.format-evidence/v2",
        "requestId": request["requestId"],
        "subject": request["subject"],
        "artifactVersionId": artifact["artifactVersionId"],
        "source": {"bytesSha256": artifact["bytesSha256"], "mediaType": artifact["mediaType"]},
        "format": "pdf",
        "provider": request["provider"],
        "runtime": request["runtime"],
        "extractionSha256": request["extraction"]["sha256"],
        "observation": {
            "schema": observation_schema,
            "value": observation_value,
            "sha256": sha(observation_value),
        },
        "provenance": [
            _typed_value(
                "tiwater.format-provenance-v1.schema.json",
                "tiwater.format-provenance/v1",
                provenance_value,
            )
        ],
    }
    evidence["evidenceId"] = f"evidence-{sha(evidence)}"
    return evidence


def run_v2(
    request_path: Path,
    output: Path,
    inspect: Callable[[Path], object],
    version: str,
    evidence_path: Path | None = None,
) -> int:
    request = json.loads(request_path.read_text(encoding="utf-8"))
    validator = evidence_path is not None
    try:
        _validate_request_v2(request, version)
        expected = _build_evidence_v2(request, inspect)
        if not validator:
            result = expected
        else:
            evidence = json.loads(evidence_path.read_text(encoding="utf-8"))
            passed = canonical(evidence) == canonical(expected)
            result = {
                "schema": "tiwater.format-evidence-verdict/v2",
                "requestId": request["requestId"],
                "subject": request["subject"],
                "artifactVersionId": request["artifact"]["artifactVersionId"],
                "evidence": {"evidenceId": evidence.get("evidenceId", "missing"), "sha256": sha(evidence)},
                "validator": request["validator"],
                "recomputedSourceBytesSha256": request["artifact"]["bytesSha256"],
                "recomputedObservationSha256": expected["observation"]["sha256"],
                "recomputedProvenanceSha256": sha(expected["provenance"]),
                "decision": "pass" if passed else "failed",
                "findings": [] if passed else [{"code": "format-evidence-recomputation-mismatch", "severity": "error"}],
            }
    except Exception as error:
        print(str(error), file=os.sys.stderr)
        artifact = request.get("artifact", {})
        result = {
            "schema": "tiwater.format-evidence-error/v1",
            "requestId": request.get("requestId", "unknown"),
            "subject": request.get("subject", {"kind": "input", "inputId": "unknown"}),
            "artifactVersionId": artifact.get("artifactVersionId", "unknown"),
            "code": "format-evidence-v2-invalid",
            "category": "evidence",
            "retryable": False,
            "provider": {
                "tool": "tiwater-pdf",
                "toolVersion": version,
                "capabilityId": "validate-inspect-evidence-v2" if validator else "inspect-evidence-v2",
            },
            "refs": [],
        }
    atomic_write(output, result)
    return 0


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
