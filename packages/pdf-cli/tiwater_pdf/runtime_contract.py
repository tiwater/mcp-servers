"""Format-neutral helpers for Tiwater runtime evidence contracts.

This module owns byte and JSON identity only. It intentionally does not inspect
PDF signatures, choose probe outcomes, or expose a CLI command.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import Any, Iterable


def _require_text(value: str, label: str) -> str:
    normalized = str(value or "").strip()
    if not normalized:
        raise ValueError(f"{label} must be non-empty")
    return normalized


def identify_file(file_path: str | Path) -> dict[str, Any]:
    """Return path, size, SHA-256, and content id for exact file bytes."""

    resolved = Path(file_path).expanduser().resolve(strict=True)
    digest = hashlib.sha256()
    size_bytes = 0
    with resolved.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            size_bytes += len(chunk)
            digest.update(chunk)
    sha256 = digest.hexdigest()
    return {
        "path": str(resolved),
        "sizeBytes": size_bytes,
        "sha256": sha256,
        "contentId": f"sha256:{sha256}",
    }


def canonical_json_bytes(value: Any) -> bytes:
    """Serialize JSON with sorted object keys, stable arrays, and no whitespace."""

    return json.dumps(
        _canonicalize_json(value),
        ensure_ascii=False,
        allow_nan=False,
        separators=(",", ":"),
    ).encode("utf-8")


def _canonicalize_json(value: Any) -> Any:
    if value is None or isinstance(value, (bool, str)):
        return value
    if isinstance(value, int):
        if value < -9_007_199_254_740_991 or value > 9_007_199_254_740_991:
            raise ValueError("canonical JSON integers must fit the cross-language safe integer range")
        return value
    if isinstance(value, float):
        raise ValueError(
            "canonical JSON v1 accepts native safe integer values only; encode exact decimal values as strings"
        )
    if isinstance(value, (list, tuple)):
        return [_canonicalize_json(item) for item in value]
    if isinstance(value, dict):
        if not all(isinstance(key, str) for key in value):
            raise ValueError("canonical JSON object keys must be strings")
        ordered_keys = sorted(value, key=lambda key: key.encode("utf-16-be", "surrogatepass"))
        return {key: _canonicalize_json(value[key]) for key in ordered_keys}
    raise ValueError(f"unsupported canonical JSON value: {type(value).__name__}")


def identify_canonical_json_artifact(
    value: Any,
    *,
    schema_id: str,
    schema_version: str,
) -> dict[str, Any]:
    """Hash canonical JSON payload bytes without creating a circular self-hash."""

    schema_id = _require_text(schema_id, "schema_id")
    schema_version = _require_text(schema_version, "schema_version")
    payload = canonical_json_bytes(value)
    sha256 = hashlib.sha256(payload).hexdigest()
    return {
        "artifactId": f"sha256:{sha256}",
        "sizeBytes": len(payload),
        "sha256": sha256,
        "mediaType": "application/json",
        "encoding": "canonical-json",
        "schema": {"id": schema_id, "version": schema_version},
    }


def native_identity(namespace: str, native_id: str) -> dict[str, Any]:
    """Describe only a real format-native id."""

    return {
        "kind": "native",
        "namespace": _require_text(namespace, "namespace"),
        "nativeId": _require_text(native_id, "native_id"),
    }


def derived_identity(derivation: str, inputs: Iterable[str]) -> dict[str, Any]:
    """Describe a deterministic derived id without fabricating ``nativeId``."""

    derivation = _require_text(derivation, "derivation")
    normalized_inputs = [_require_text(value, "inputs item") for value in inputs]
    if not normalized_inputs:
        raise ValueError("inputs must be non-empty")
    return {
        "kind": "derived",
        "derivation": derivation,
        "inputs": normalized_inputs,
    }


def normalize_evidence(report: Any) -> dict[str, Any]:
    """Project a runtime-owned JSON report into stable generic evidence nodes."""

    nodes: list[dict[str, Any]] = []
    objects: list[dict[str, Any]] = []

    def visit(value: Any, node_id: str, parent_id: str | None, kind: str, depth: int) -> None:
        nodes.append(
            {
                "runtimeNodeId": node_id,
                "kind": kind,
                "valueType": _evidence_value_type(value),
                "value": _evidence_scalar(value),
                "locator": node_id,
                "derivedFrom": [],
                "containedBy": parent_id,
            }
        )
        objects.append(
            {
                "objectId": node_id,
                "objectType": kind,
                "root": parent_id is None,
                "parentObjectId": parent_id,
                "identity": derived_identity("normalized-json-pointer-v1", [node_id]),
            }
        )

        if isinstance(value, dict):
            for key in sorted(value, key=lambda item: item.encode("utf-16-be", "surrogatepass")):
                if depth == 0 and key.lower() in {"file", "input", "output"}:
                    continue
                child_id = f"/{_json_pointer_escape(key)}" if node_id == "$" else f"{node_id}/{_json_pointer_escape(key)}"
                visit(value[key], child_id, node_id, key, depth + 1)
        elif isinstance(value, (list, tuple)):
            item_kind = _singular(kind)
            for index, item in enumerate(value):
                visit(item, f"{node_id}/{index}", node_id, item_kind, depth + 1)

    visit(report, "$", None, "document", 0)
    return {
        "payload": {"schemaVersion": "1.0.0", "nodes": nodes},
        "objects": objects,
    }


def _evidence_scalar(value: Any) -> Any:
    if value is None or isinstance(value, (bool, str, int)):
        return value
    if isinstance(value, float):
        return repr(value)
    return None


def _evidence_value_type(value: Any) -> str:
    if value is None:
        return "null"
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, str):
        return "string"
    if isinstance(value, int):
        return "integer"
    if isinstance(value, float):
        return "decimal-string"
    if isinstance(value, dict):
        return "object"
    if isinstance(value, (list, tuple)):
        return "array"
    raise ValueError(f"unsupported evidence value: {type(value).__name__}")


def _json_pointer_escape(value: str) -> str:
    return value.replace("~", "~0").replace("/", "~1")


def _singular(value: str) -> str:
    if value == "children":
        return "child"
    if value.endswith("ies"):
        return f"{value[:-3]}y"
    if value.endswith("s") and len(value) > 1:
        return value[:-1]
    return "item"
