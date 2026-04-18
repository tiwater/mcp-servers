#!/usr/bin/env python3
"""Minimal PPTX CLI for inspect/export/fill workflows."""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import zipfile
from io import BytesIO
from pathlib import Path
import xml.etree.ElementTree as ET

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PLACEHOLDER_RE = re.compile(r"{{\s*([^{}]+?)\s*}}")


def _load_zip(path: str) -> zipfile.ZipFile:
    return zipfile.ZipFile(path, "r")


def _slide_paths(names: list[str]) -> list[str]:
    return sorted([n for n in names if n.startswith("ppt/slides/slide") and n.endswith(".xml")])


def _notes_paths(names: list[str]) -> list[str]:
    return sorted([n for n in names if n.startswith("ppt/notesSlides/notesSlide") and n.endswith(".xml")])


def _extract_text_nodes(xml_bytes: bytes) -> list[str]:
    root = ET.fromstring(xml_bytes)
    texts: list[str] = []
    for node in root.findall(f".//{{{A_NS}}}t"):
        if node.text:
            texts.append(node.text)
    return texts


def _extract_placeholders(texts: list[str]) -> list[str]:
    values: list[str] = []
    seen: set[str] = set()
    for text in texts:
        for match in PLACEHOLDER_RE.findall(text):
            key = match.strip()
            if key and key not in seen:
                seen.add(key)
                values.append(key)
    return values


def inspect_pptx(input_path: str) -> dict:
    with _load_zip(input_path) as zf:
        names = zf.namelist()
        slides = []
        all_placeholders: set[str] = set()
        for idx, slide_path in enumerate(_slide_paths(names), start=1):
            texts = _extract_text_nodes(zf.read(slide_path))
            placeholders = _extract_placeholders(texts)
            all_placeholders.update(placeholders)
            slides.append(
                {
                    "slideNumber": idx,
                    "path": slide_path,
                    "textCount": len(texts),
                    "placeholders": placeholders,
                }
            )

    return {
        "file": input_path,
        "slideCount": len(slides),
        "placeholders": sorted(all_placeholders),
        "slides": slides,
    }


def export_json(input_path: str) -> dict:
    with _load_zip(input_path) as zf:
        names = zf.namelist()
        slides = []
        for idx, slide_path in enumerate(_slide_paths(names), start=1):
            texts = _extract_text_nodes(zf.read(slide_path))
            slides.append(
                {
                    "slideNumber": idx,
                    "path": slide_path,
                    "texts": texts,
                    "placeholders": _extract_placeholders(texts),
                }
            )

        notes = []
        for idx, notes_path in enumerate(_notes_paths(names), start=1):
            texts = _extract_text_nodes(zf.read(notes_path))
            notes.append({"notesNumber": idx, "path": notes_path, "texts": texts})

    return {"file": input_path, "slides": slides, "notes": notes}


def _replace_text(text: str, mapping: dict[str, str]) -> str:
    output = text
    for key, value in mapping.items():
        output = output.replace(f"{{{{{key}}}}}", value)
    return output


def _update_xml_placeholders(xml_bytes: bytes, mapping: dict[str, str]) -> tuple[bytes, int]:
    root = ET.fromstring(xml_bytes)
    changed = 0
    for node in root.findall(f".//{{{A_NS}}}t"):
        if not node.text:
            continue
        updated = _replace_text(node.text, mapping)
        if updated != node.text:
            node.text = updated
            changed += 1
    if changed == 0:
        return xml_bytes, 0
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), changed


def fill_template(template_path: str, data_path: str, output_path: str) -> dict:
    payload = json.loads(Path(data_path).read_text(encoding="utf-8"))
    if isinstance(payload, dict) and isinstance(payload.get("textValues"), dict):
        mapping = payload["textValues"]
    elif isinstance(payload, dict):
        mapping = payload
    else:
        raise ValueError("Fill data must be a JSON object or an object with textValues")

    mapping = {str(k): str(v) for k, v in mapping.items()}

    with _load_zip(template_path) as source_zip:
        in_memory = BytesIO()
        with zipfile.ZipFile(in_memory, "w", compression=zipfile.ZIP_DEFLATED) as target_zip:
            changed_slides = 0
            changed_notes = 0
            for name in source_zip.namelist():
                data = source_zip.read(name)
                if name.startswith("ppt/slides/slide") and name.endswith(".xml"):
                    data, count = _update_xml_placeholders(data, mapping)
                    if count:
                        changed_slides += 1
                elif name.startswith("ppt/notesSlides/notesSlide") and name.endswith(".xml"):
                    data, count = _update_xml_placeholders(data, mapping)
                    if count:
                        changed_notes += 1
                target_zip.writestr(name, data)

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    Path(output_path).write_bytes(in_memory.getvalue())

    return {
        "template": template_path,
        "output": output_path,
        "changedSlides": changed_slides,
        "changedNotes": changed_notes,
        "placeholderCount": len(mapping),
    }


def _print_json(payload: dict) -> None:
    print(json.dumps(payload, ensure_ascii=False))


def run(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(prog="tiwater-pptx")
    subparsers = parser.add_subparsers(dest="command", required=True)

    inspect_parser = subparsers.add_parser("inspect")
    inspect_parser.add_argument("input")
    inspect_parser.add_argument("--json", action="store_true")

    export_parser = subparsers.add_parser("export-json")
    export_parser.add_argument("input")
    export_parser.add_argument("output", nargs="?")

    fill_parser = subparsers.add_parser("fill-template")
    fill_parser.add_argument("template")
    fill_parser.add_argument("data")
    fill_parser.add_argument("output")

    args = parser.parse_args(argv)

    if args.command == "inspect":
        report = inspect_pptx(args.input)
        if args.json:
            _print_json(report)
        else:
            print(f"File: {report['file']}")
            print(f"Slides: {report['slideCount']}")
            print(f"Placeholders: {', '.join(report['placeholders'])}")
        return 0

    if args.command == "export-json":
        report = export_json(args.input)
        if args.output:
            Path(args.output).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        else:
            _print_json(report)
        return 0

    if args.command == "fill-template":
        result = fill_template(args.template, args.data, args.output)
        _print_json(result)
        return 0

    parser.error(f"Unknown command: {args.command}")
    return 1


if __name__ == "__main__":
    try:
        raise SystemExit(run(sys.argv[1:]))
    except Exception as exc:
        print(str(exc), file=sys.stderr)
        raise SystemExit(1)
