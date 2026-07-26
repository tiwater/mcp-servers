import hashlib
import json
import sys

import fitz

from tiwater_pdf import __version__
from tiwater_pdf.cli import main
from tiwater_pdf.format_evidence import canonical, run, sha


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
