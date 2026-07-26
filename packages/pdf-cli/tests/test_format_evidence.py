import hashlib
import json
from pathlib import Path
import sys

import fitz

from tiwater_pdf import __version__
from tiwater_pdf.cli import main
from tiwater_pdf.format_evidence import canonical, run, run_v2, sha


def contract_ref(name, contract_id):
    path = Path(__file__).parents[2] / "format-evidence" / "contracts" / name
    return {"id": contract_id, "sha256": hashlib.sha256(path.read_bytes()).hexdigest()}


def v2_request(source):
    extraction_value = {"facets": ["format-summary"]}
    return {
        "schema": "tiwater.format-evidence-request/v2",
        "requestId": "request-v2",
        "runId": "run-v2",
        "subject": {"kind": "input", "inputId": "input-v2"},
        "artifact": {
            "artifactVersionId": "av-v2",
            "path": str(source),
            "bytesSha256": hashlib.sha256(source.read_bytes()).hexdigest(),
            "mediaType": "application/pdf",
        },
        "provider": {"id": "tiwater-pdf", "version": __version__},
        "validator": {"id": "tiwater-pdf-validator", "version": __version__},
        "runtime": {"id": "tiwater-pdf", "version": __version__},
        "extraction": {
            "schema": contract_ref("tiwater.format-extraction-options-v1.schema.json", "tiwater.format-extraction-options/v1"),
            "value": extraction_value,
            "sha256": sha(extraction_value),
        },
        "expectedEvidenceContract": contract_ref("tiwater.format-evidence-v2.schema.json", "tiwater.format-evidence/v2"),
    }


def make_pdf(path):
    document = fitz.open()
    document.new_page().insert_text((72, 72), "Published evidence v2")
    document.save(path)
    document.close()


def test_pdf_v2_evidence_is_read_only_and_tampering_is_rejected(tmp_path):
    source = tmp_path / "source.pdf"
    make_pdf(source)
    request_path = tmp_path / "request.json"
    evidence = tmp_path / "evidence.json"
    verdict = tmp_path / "verdict.json"
    request_path.write_text(canonical(v2_request(source)), encoding="utf-8")
    inspect = lambda path: {"document": {"pages": len(fitz.open(path))}, "tables": []}

    assert run_v2(request_path, evidence, inspect, __version__) == 0
    value = json.loads(evidence.read_text())
    observation = value["observation"]["value"]
    assert value["schema"] == "tiwater.format-evidence/v2"
    assert observation["inventoryUniverse"]["candidates"]
    assert observation["targetUniverse"]["candidates"] == []
    assert run_v2(request_path, verdict, inspect, __version__, evidence) == 0
    assert json.loads(verdict.read_text())["decision"] == "pass"

    value["observation"]["value"]["targetUniverse"]["candidates"].append({"unauthorized": True})
    evidence.write_text(canonical(value), encoding="utf-8")
    verdict.unlink()
    assert run_v2(request_path, verdict, inspect, __version__, evidence) == 0
    changed = json.loads(verdict.read_text())
    assert changed["decision"] == "failed"
    assert changed["findings"] == [{"code": "format-evidence-recomputation-mismatch", "severity": "error"}]


def test_pdf_v2_rejects_provider_drift_and_changed_source_bytes(tmp_path):
    source = tmp_path / "source.pdf"
    make_pdf(source)
    inspect = lambda path: {"pages": len(fitz.open(path))}
    output = tmp_path / "output.json"
    request_path = tmp_path / "request.json"
    request = v2_request(source)
    request["provider"]["version"] = "invented"
    request_path.write_text(canonical(request), encoding="utf-8")
    assert run_v2(request_path, output, inspect, __version__) == 0
    assert json.loads(output.read_text())["schema"] == "tiwater.format-evidence-error/v1"

    request = v2_request(source)
    request_path.write_text(canonical(request), encoding="utf-8")
    source.write_bytes(source.read_bytes() + b"tampered")
    output.unlink()
    assert run_v2(request_path, output, inspect, __version__) == 0
    assert json.loads(output.read_text())["code"] == "format-evidence-v2-invalid"


def test_pdf_cli_exposes_v2_producer_and_validator(tmp_path, monkeypatch):
    source = tmp_path / "source.pdf"
    make_pdf(source)
    request_path = tmp_path / "request.json"
    evidence = tmp_path / "evidence.json"
    verdict = tmp_path / "verdict.json"
    request_path.write_text(canonical(v2_request(source)), encoding="utf-8")

    monkeypatch.setattr(sys, "argv", ["tiwater-pdf", "inspect-evidence-v2", "--request", str(request_path), "--output", str(evidence)])
    assert main() == 0
    monkeypatch.setattr(sys, "argv", ["tiwater-pdf", "validate-inspect-evidence-v2", "--request", str(request_path), "--evidence", str(evidence), "--output", str(verdict)])
    assert main() == 0
    assert json.loads(evidence.read_text())["provider"] == {"id": "tiwater-pdf", "version": __version__}
    assert json.loads(verdict.read_text())["decision"] == "pass"


def test_pdf_evidence_is_recomputed_and_tampering_is_rejected(tmp_path):
    source = tmp_path / "source.pdf"
    document = fitz.open()
    page = document.new_page()
    page.insert_text((72, 72), "Published evidence")
    document.save(source)
    document.close()
    evidence, verdict, request_path = tmp_path / "evidence.json", tmp_path / "verdict.json", tmp_path / "request.json"
    options = {}
    request = {"schema": "tiwater.format-evidence-request/v1", "requestId": "request-1", "runId": "run-1", "subject": {"kind": "input", "inputId": "input-1"}, "artifact": {"artifactVersionId": "av-1", "path": str(source), "bytesSha256": hashlib.sha256(source.read_bytes()).hexdigest(), "format": "pdf"}, "extraction": {"schema": "tiwater.pdf.inspect/v1", "options": options, "optionsSha256": sha(options)}, "expectedEvidenceSchema": "lucid.published-format-evidence/v1", "outputPath": str(evidence)}
    request_path.write_text(canonical(request), encoding="utf-8")
    inspect = lambda path: {"pages": len(fitz.open(path))}
    assert run(request_path, evidence, inspect, __version__) == 0
    assert run(request_path, verdict, inspect, __version__, evidence) == 0
    assert json.loads(verdict.read_text())["pass"] is True
    changed = json.loads(evidence.read_text())
    changed["observations"][0]["value"]["pages"] = 99
    evidence.write_text(canonical(changed), encoding="utf-8")
    verdict.unlink()
    assert run(request_path, verdict, inspect, __version__, evidence) == 0
    assert json.loads(verdict.read_text())["pass"] is False


def test_pdf_cli_evidence_reports_installed_package_version(tmp_path, monkeypatch):
    source = tmp_path / "source.pdf"
    document = fitz.open()
    document.new_page().insert_text((72, 72), "Published evidence")
    document.save(source)
    document.close()
    evidence = tmp_path / "evidence.json"
    verdict = tmp_path / "verdict.json"
    request_path = tmp_path / "request.json"
    options = {}
    request = {"schema": "tiwater.format-evidence-request/v1", "requestId": "request-1", "runId": "run-1", "subject": {"kind": "input", "inputId": "input-1"}, "artifact": {"artifactVersionId": "av-1", "path": str(source), "bytesSha256": hashlib.sha256(source.read_bytes()).hexdigest(), "format": "pdf"}, "extraction": {"schema": "tiwater.pdf.inspect/v1", "options": options, "optionsSha256": sha(options)}, "expectedEvidenceSchema": "lucid.published-format-evidence/v1", "outputPath": str(evidence)}
    request_path.write_text(canonical(request), encoding="utf-8")

    monkeypatch.setattr(sys, "argv", ["tiwater-pdf", "inspect-evidence", "--request", str(request_path), "--output", str(evidence)])
    assert main() == 0
    monkeypatch.setattr(sys, "argv", ["tiwater-pdf", "validate-inspect-evidence", "--request", str(request_path), "--evidence", str(evidence), "--output", str(verdict)])
    assert main() == 0

    assert json.loads(evidence.read_text())["provider"]["toolVersion"] == __version__
    verdict_value = json.loads(verdict.read_text())
    assert verdict_value["pass"] is True
    assert verdict_value["validator"]["toolVersion"] == __version__
