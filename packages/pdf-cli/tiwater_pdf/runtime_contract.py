"""Format-neutral helpers for Tiwater runtime evidence contracts.

This module owns byte and JSON identity only. It intentionally does not inspect
PDF signatures, choose probe outcomes, or expose a CLI command.
"""

from __future__ import annotations

import hashlib
import json
import math
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
        if (
            not math.isfinite(value)
            or not value.is_integer()
            or value < -9_007_199_254_740_991
            or value > 9_007_199_254_740_991
        ):
            raise ValueError(
                "canonical JSON v1 accepts safe integer numbers only; encode exact decimal values as strings"
            )
        return int(value)
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
