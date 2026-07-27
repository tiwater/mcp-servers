"""Contract-first tests for the Set-15 observation-only provider contract manifest.

The manifest fixture is a byte copy of the real lucid schema-set 15
``provider-contract-manifest.schema.json``; expected semantics come from that
frozen schema and the deployed contract bytes, never from the producer.
"""

import hashlib
import json
from pathlib import Path
import sys

import pytest
from jsonschema import Draft202012Validator

from tiwater_pdf import __version__
from tiwater_pdf.cli import main
from tiwater_pdf.provider_contract_manifest import (
    MANIFEST_SCHEMA_ID,
    VERDICT_SCHEMA_ID,
    run_producer,
    run_validator,
)

SET15_SCHEMA_PATH = Path(__file__).parent / "fixtures" / "provider-contract-manifest.schema.json"
VERDICT_SCHEMA_PATH = (
    Path(__file__).parents[2] / "format-evidence" / "contracts"
    / "tiwater.provider-contract-manifest-verdict-v1.schema.json"
)

SET15_VALIDATOR = Draft202012Validator(json.loads(SET15_SCHEMA_PATH.read_text(encoding="utf-8")))
VERDICT_VALIDATOR = Draft202012Validator(json.loads(VERDICT_SCHEMA_PATH.read_text(encoding="utf-8")))


def produce(tmp_path, schema_set_version=15):
    manifest_path = tmp_path / "manifest.json"
    assert run_producer(manifest_path, schema_set_version, __version__) == 0
    return manifest_path, json.loads(manifest_path.read_text(encoding="utf-8"))


def validate(tmp_path, manifest_path):
    verdict_path = tmp_path / "verdict.json"
    exit_code = run_validator(manifest_path, verdict_path, __version__)
    return exit_code, json.loads(verdict_path.read_text(encoding="utf-8"))


def rewrite(manifest_path, manifest):
    manifest_path.write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def test_manifest_satisfies_real_set15_schema_and_is_observation_only(tmp_path):
    _, manifest = produce(tmp_path)
    assert SET15_VALIDATOR.is_valid(manifest), list(SET15_VALIDATOR.iter_errors(manifest))
    assert manifest["schema"] == MANIFEST_SCHEMA_ID == "lucid.provider-contract-manifest"
    assert manifest["schemaSetVersion"] == 15
    assert manifest["provider"] == {"id": "tiwater-pdf", "version": __version__}
    assert manifest["runtime"] == {"id": "tiwater-pdf", "version": __version__}

    ports = manifest["ports"]
    assert [port["kind"] for port in ports] == ["observe"]
    assert not {"derive", "execute", "render", "reobserve"} & {port["kind"] for port in ports}
    port = ports[0]
    assert port["producer"]["adapterIdentity"] == {"id": "tiwater-pdf", "version": __version__}
    assert port["validator"]["adapterIdentity"] == {"id": "tiwater-pdf-validator", "version": __version__}
    assert port["requestSchema"] == "tiwater.format-evidence-request/v2"
    assert port["validatorRequestSchema"] == "tiwater.format-evidence-request/v2"
    assert port["resultSchema"] == "tiwater.format-evidence/v2"
    assert port["verdictSchema"] == "tiwater.format-evidence-verdict/v2"
    assert port["options"] == [{"name": "facets", "valueSchema": "tiwater.format-extraction-options/v1"}]
    assert {"schemaSetVersion", "bytesSha256", "provider", "optionsSha256"} <= set(port["cacheKeyComposition"])
    assert port["sideEffect"] == {"kind": "read-only", "idempotent": True}
    assert port["attemptBudget"] >= 1

    # Determinism: same package and schema set version yield identical bytes.
    second_path = tmp_path / "manifest-second.json"
    assert run_producer(second_path, 15, __version__) == 0
    assert second_path.read_bytes() == (tmp_path / "manifest.json").read_bytes()


def test_declared_schema_names_exist_as_deployed_contract_bytes(tmp_path):
    _, manifest = produce(tmp_path)
    contracts = Path(__file__).parents[2] / "format-evidence" / "contracts"
    port = manifest["ports"][0]
    declared = {
        port["requestSchema"],
        port["validatorRequestSchema"],
        port["resultSchema"],
        port["verdictSchema"],
        *(option["valueSchema"] for option in port["options"]),
    }
    for name in declared:
        base, _, version = name.partition("/v")
        assert version.isdigit(), name
        schema_file = contracts / f"{base}-v{version}.schema.json"
        assert schema_file.is_file(), f"declared schema missing from deployed contracts: {name}"
        payload = json.loads(schema_file.read_bytes())
        assert payload.get("$id") == name, f"deployed schema $id drifted from declared name: {name}"
        assert len(hashlib.sha256(schema_file.read_bytes()).hexdigest()) == 64


def test_validator_passes_fresh_manifest_and_verdict_matches_packaged_schema(tmp_path):
    manifest_path, _ = produce(tmp_path)
    exit_code, verdict = validate(tmp_path, manifest_path)
    assert exit_code == 0
    assert verdict["schema"] == VERDICT_SCHEMA_ID
    assert verdict["decision"] == "pass"
    assert verdict["findings"] == []
    assert VERDICT_VALIDATOR.is_valid(verdict), list(VERDICT_VALIDATOR.iter_errors(verdict))


@pytest.mark.parametrize(
    "mutate",
    [
        pytest.param(lambda m: m["ports"][0].update(kind="render"), id="observe-port-becomes-render"),
        pytest.param(lambda m: m["ports"].append({**m["ports"][0], "kind": "derive"}), id="derive-port-added"),
        pytest.param(lambda m: m["ports"].append({**m["ports"][0], "kind": "execute"}), id="execute-port-added"),
        pytest.param(lambda m: m["ports"].append({**m["ports"][0], "kind": "reobserve"}), id="reobserve-port-added"),
        pytest.param(lambda m: m["ports"][0]["producer"].update(version="0.0.1"), id="producer-version-drift"),
        pytest.param(lambda m: m["ports"][0]["producer"]["adapterIdentity"].update(id="invented-adapter"), id="producer-adapter-identity-drift"),
        pytest.param(lambda m: m["ports"][0]["validator"]["adapterIdentity"].update(version="0.0.1"), id="validator-adapter-identity-drift"),
        pytest.param(lambda m: m["ports"][0].update(requestSchema="tiwater.invented-request/v9"), id="request-schema-drift"),
        pytest.param(lambda m: m["ports"][0].update(validatorRequestSchema="tiwater.invented-request/v9"), id="validator-request-schema-drift"),
        pytest.param(lambda m: m["ports"][0].update(verdictSchema="tiwater.invented-verdict/v9"), id="verdict-schema-drift"),
        pytest.param(lambda m: m["ports"][0].update(options=[{"name": "model", "valueSchema": "tiwater.format-extraction-options/v1"}]), id="undeclared-option"),
        pytest.param(lambda m: m["ports"][0]["cacheKeyComposition"].remove("bytesSha256"), id="cache-component-dropped"),
        pytest.param(lambda m: m["ports"][0].update(resourceDeclarations=[{"resourceKey": "pdf:document", "access": "exclusive-write"}]), id="resource-access-escalation"),
        pytest.param(lambda m: m["ports"][0].update(sideEffect={"kind": "read-only", "idempotent": False}), id="side-effect-drift"),
        pytest.param(lambda m: m["ports"][0].update(attemptBudget=99), id="attempt-budget-drift"),
        pytest.param(lambda m: m.update(schemaSetVersion=14), id="schema-set-version-drift"),
        pytest.param(lambda m: m.update(manifestId="manifest-" + "0" * 64), id="manifest-id-forgery"),
        pytest.param(lambda m: m["provider"].update(version="0.0.1"), id="provider-version-drift"),
        pytest.param(lambda m: m["runtime"].update(id="invented-runtime"), id="runtime-identity-drift"),
    ],
)
def test_independent_validator_rejects_mutations(tmp_path, mutate):
    manifest_path, manifest = produce(tmp_path)
    mutate(manifest)
    rewrite(manifest_path, manifest)
    exit_code, verdict = validate(tmp_path, manifest_path)
    assert exit_code == 1
    assert verdict["decision"] == "fail"
    assert verdict["findings"]
    assert VERDICT_VALIDATOR.is_valid(verdict), list(VERDICT_VALIDATOR.iter_errors(verdict))


def test_independent_validator_rejects_dropped_port_even_with_forged_field_set(tmp_path):
    manifest_path, manifest = produce(tmp_path)
    del manifest["ports"][0]
    rewrite(manifest_path, manifest)
    exit_code, verdict = validate(tmp_path, manifest_path)
    assert exit_code == 1
    assert verdict["decision"] == "fail"


def test_producer_fails_closed_without_schema_set_version(tmp_path, monkeypatch):
    monkeypatch.setattr(sys, "argv", ["tiwater-pdf", "provider-contract-manifest", "--output", str(tmp_path / "manifest.json")])
    with pytest.raises(SystemExit) as error:
        main()
    assert error.value.code == 2
    assert not (tmp_path / "manifest.json").exists()


def test_cli_producer_and_validator_round_trip(tmp_path, monkeypatch):
    manifest_path = tmp_path / "manifest.json"
    verdict_path = tmp_path / "verdict.json"
    monkeypatch.setattr(
        sys,
        "argv",
        ["tiwater-pdf", "provider-contract-manifest", "--schema-set-version", "15", "--output", str(manifest_path)],
    )
    assert main() == 0
    monkeypatch.setattr(
        sys,
        "argv",
        ["tiwater-pdf", "validate-provider-contract-manifest", "--manifest", str(manifest_path), "--output", str(verdict_path)],
    )
    assert main() == 0
    assert json.loads(verdict_path.read_text(encoding="utf-8"))["decision"] == "pass"


def test_outputs_must_be_fresh(tmp_path):
    manifest_path, _ = produce(tmp_path)
    with pytest.raises(ValueError, match="fresh"):
        run_producer(manifest_path, 15, __version__)
    verdict_path = tmp_path / "verdict.json"
    assert run_validator(manifest_path, verdict_path, __version__) == 0
    with pytest.raises(ValueError, match="fresh"):
        run_validator(manifest_path, verdict_path, __version__)
