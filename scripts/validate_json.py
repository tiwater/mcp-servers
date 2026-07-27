#!/usr/bin/env python3
"""Strict JSON syntax gate for tracked sources and packed artifacts.

Rejects trailing commas, truncated payloads, and any other text that is not
RFC 8259 JSON. This is a publish gate only; it does not change package
behavior or contract semantics.
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
import tarfile
import tempfile
import zipfile
from pathlib import Path


ARCHIVE_SUFFIXES = (".nupkg", ".whl", ".zip", ".tar.gz", ".tgz")


def repo_root() -> Path:
    return Path(__file__).resolve().parents[1]


def tracked_json_files(root: Path) -> list[Path]:
    result = subprocess.run(
        ["git", "-C", str(root), "ls-files", "-z", "--", "*.json"],
        check=True,
        capture_output=True,
    )
    rels = [item.decode() for item in result.stdout.split(b"\0") if item]
    return sorted(root / rel for rel in rels)


def iter_archive_json(archive: Path):
    name = archive.name.lower()
    if name.endswith(".tar.gz") or name.endswith(".tgz"):
        with tarfile.open(archive, "r:*") as archive_file:
            for member in archive_file.getmembers():
                if not member.isfile() or not member.name.lower().endswith(".json"):
                    continue
                extracted = archive_file.extractfile(member)
                if extracted is None:
                    continue
                yield member.name, extracted.read()
        return

    with zipfile.ZipFile(archive) as archive_file:
        for member in archive_file.namelist():
            if member.endswith("/") or not member.lower().endswith(".json"):
                continue
            yield member, archive_file.read(member)


def parse_bytes(label: str, payload: bytes) -> None:
    try:
        text = payload.decode("utf-8")
    except UnicodeDecodeError as error:
        raise ValueError(f"{label}: not valid UTF-8 ({error})") from error
    try:
        json.loads(text)
    except json.JSONDecodeError as error:
        raise ValueError(
            f"{label}: invalid JSON at line {error.lineno} column {error.colno}: {error.msg}"
        ) from error


def display_path(path: Path, root: Path | None = None) -> str:
    if root is None:
        return str(path)
    try:
        return str(path.resolve().relative_to(root.resolve()))
    except ValueError:
        return str(path)


def validate_paths(paths: list[Path], root: Path | None = None) -> list[str]:
    failures: list[str] = []
    for path in paths:
        label = display_path(path, root)
        try:
            parse_bytes(label, path.read_bytes())
        except ValueError as error:
            failures.append(str(error))
    return failures


def validate_archives(archives: list[Path]) -> list[str]:
    failures: list[str] = []
    for archive in archives:
        try:
            members = list(iter_archive_json(archive))
        except (OSError, tarfile.TarError, zipfile.BadZipFile) as error:
            failures.append(f"{archive}: unreadable archive ({error})")
            continue
        if not members:
            continue
        for member_name, payload in members:
            label = f"{archive}!{member_name}"
            try:
                parse_bytes(label, payload)
            except ValueError as error:
                failures.append(str(error))
    return failures


def discover_archives(directory: Path) -> list[Path]:
    found: list[Path] = []
    for path in sorted(directory.rglob("*")):
        if not path.is_file():
            continue
        lower = path.name.lower()
        if lower.endswith(ARCHIVE_SUFFIXES) or lower.endswith(".nupkg"):
            found.append(path)
    return found


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--root",
        type=Path,
        default=None,
        help="Repository root (defaults to the parent of scripts/)",
    )
    parser.add_argument(
        "--tracked",
        action="store_true",
        help="Validate every version-controlled *.json file",
    )
    parser.add_argument(
        "--archives-from",
        type=Path,
        action="append",
        default=[],
        help="Directory whose packed artifacts are scanned for embedded *.json",
    )
    parser.add_argument(
        "paths",
        nargs="*",
        type=Path,
        help="Explicit files or directories to validate",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    root = (args.root or repo_root()).resolve()
    failures: list[str] = []
    checked = 0
    selected = bool(args.tracked or args.paths or args.archives_from)

    if not selected:
        print("json-gate: no JSON inputs selected", file=sys.stderr)
        return 2

    if args.tracked:
        tracked = tracked_json_files(root)
        checked += len(tracked)
        failures.extend(validate_paths(tracked, root=root))

    explicit_files: list[Path] = []
    archives: list[Path] = []
    for path in args.paths:
        resolved = path if path.is_absolute() else root / path
        if resolved.is_dir():
            for candidate in sorted(resolved.rglob("*.json")):
                if candidate.is_file():
                    explicit_files.append(candidate)
            archives.extend(discover_archives(resolved))
        elif resolved.is_file():
            lower = resolved.name.lower()
            if lower.endswith(".json"):
                explicit_files.append(resolved)
            elif any(lower.endswith(suffix) for suffix in ARCHIVE_SUFFIXES):
                archives.append(resolved)
            else:
                failures.append(f"{resolved}: unsupported path for JSON gate")
        else:
            failures.append(f"{resolved}: path does not exist")

    for directory in args.archives_from:
        resolved = directory if directory.is_absolute() else root / directory
        if not resolved.is_dir():
            failures.append(f"{resolved}: archive directory does not exist")
            continue
        found = discover_archives(resolved)
        if not found:
            failures.append(f"{resolved}: no packed archives found")
            continue
        archives.extend(found)

    if explicit_files:
        checked += len(explicit_files)
        failures.extend(validate_paths(explicit_files, root=root))

    unique_archives = sorted(set(archives))
    if unique_archives:
        member_count = 0
        for archive in unique_archives:
            try:
                member_count += sum(1 for _ in iter_archive_json(archive))
            except (OSError, tarfile.TarError, zipfile.BadZipFile) as error:
                failures.append(f"{archive}: unreadable archive ({error})")
        checked += member_count
        failures.extend(validate_archives(unique_archives))

    if failures:
        print(f"json-gate: {len(failures)} failure(s)", file=sys.stderr)
        for failure in failures:
            print(failure, file=sys.stderr)
        return 1

    print(f"json-gate: ok ({checked} JSON file(s))")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
