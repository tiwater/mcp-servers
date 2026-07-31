#!/usr/bin/env python3
"""Publish smoke gate for the five CLI public surfaces.

Installs tools from local pack/build artifacts, enumerates each CLI surface,
requires the generic inspect/edit/render/OCR entry points, and fails if
framework-oriented or specialized orchestration terms appear in the public
help text.
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path

DOTNET_PACKAGES = (
    ("tiwater-xlsx", "packages/xlsx-cli", "xlsx.csproj"),
    ("tiwater-docx", "packages/docx-cli", "docx.csproj"),
    ("tiwater-pptx", "packages/pptx-cli", "pptx.csproj"),
    ("tiwater-convert", "packages/convert-cli", "convert.csproj"),
)

# Exact phrases that must never appear in public CLI surface text.
FORBIDDEN_PHRASES = (
    "provider-contract-manifest",
    "format-evidence",
    "inspect-evidence",
    "derive-operation",
    "execute-effect",
    "validate-effect",
    "schema set",
    "schema-set",
    "schemaset",
)

# Single-token bans matched on word boundaries.
FORBIDDEN_WORDS = (
    "lucid",
    "workflow",
    "scenario",
    "conformance",
    "release",
)

REQUIRED_COMMANDS = {
    "tiwater-docx": ("inspect", "edit"),
    "tiwater-xlsx": ("inspect", "edit"),
    "tiwater-pptx": ("inspect", "apply-format-edits"),
    "tiwater-pdf": ("inspect", "ocr"),
}


def repo_root() -> Path:
    return Path(__file__).resolve().parents[1]


def run(
    command: list[str],
    *,
    cwd: Path | None = None,
    env: dict[str, str] | None = None,
    check: bool = True,
) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        command,
        cwd=str(cwd) if cwd is not None else None,
        env=env,
        text=True,
        capture_output=True,
        check=check,
    )


def read_csproj_identity(csproj: Path) -> tuple[str, str]:
    root = ET.parse(csproj).getroot()
    package_id = None
    version = None
    for element in root.iter():
        tag = element.tag.rsplit("}", 1)[-1]
        if tag == "PackageId" and element.text:
            package_id = element.text.strip()
        elif tag == "Version" and element.text:
            version = element.text.strip()
    if not package_id or not version:
        raise RuntimeError(f"{csproj}: missing PackageId/Version")
    return package_id, version


def pack_dotnet(root: Path) -> dict[str, Path]:
    sources: dict[str, Path] = {}
    for command_name, relative_dir, csproj_name in DOTNET_PACKAGES:
        package_dir = root / relative_dir
        nupkg_dir = package_dir / "nupkg"
        if nupkg_dir.exists():
            shutil.rmtree(nupkg_dir)
        nupkg_dir.mkdir(parents=True)
        result = run(
            ["dotnet", "pack", "-c", "Release", "-o", "./nupkg"],
            cwd=package_dir,
            check=False,
        )
        if result.returncode != 0:
            raise RuntimeError(
                f"dotnet pack failed for {package_dir}:\n{result.stdout}\n{result.stderr}"
            )
        package_id, version = read_csproj_identity(package_dir / csproj_name)
        artifact = nupkg_dir / f"{package_id}.{version}.nupkg"
        if not artifact.is_file():
            matches = sorted(nupkg_dir.glob(f"{package_id}.*.nupkg"))
            if not matches:
                raise RuntimeError(f"{nupkg_dir}: packed nupkg missing for {package_id}")
            artifact = matches[-1]
        sources[command_name] = nupkg_dir.resolve()
    return sources


def build_pdf(root: Path) -> Path:
    dist = root / "packages" / "pdf-cli" / "dist"
    if dist.exists():
        shutil.rmtree(dist)
    dist.mkdir(parents=True)
    result = run(
        [sys.executable, "-m", "build", "--outdir", str(dist), str(root / "packages" / "pdf-cli")],
        cwd=root,
        check=False,
    )
    if result.returncode != 0:
        raise RuntimeError(f"python build failed for pdf-cli:\n{result.stdout}\n{result.stderr}")
    wheels = sorted(dist.glob("*.whl"))
    if not wheels:
        raise RuntimeError(f"{dist}: wheel missing")
    return dist.resolve()


def write_nuget_config(path: Path, sources: dict[str, Path]) -> None:
    lines = [
        '<?xml version="1.0" encoding="utf-8"?>',
        "<configuration>",
        "  <packageSources>",
        "    <clear />",
    ]
    for name, directory in sources.items():
        lines.append(
            f'    <add key="{name}" value="{directory}" />'
        )
    lines.extend(["  </packageSources>", "</configuration>", ""])
    path.write_text("\n".join(lines), encoding="utf-8")


def install_dotnet_tools(
    root: Path, tool_path: Path, sources: dict[str, Path]
) -> dict[str, Path]:
    tool_path.mkdir(parents=True, exist_ok=True)
    config = tool_path / "nuget.config"
    write_nuget_config(config, sources)
    binaries: dict[str, Path] = {}
    for command_name, relative_dir, csproj_name in DOTNET_PACKAGES:
        package_id, version = read_csproj_identity(root / relative_dir / csproj_name)
        run(
            [
                "dotnet",
                "tool",
                "uninstall",
                package_id,
                "--tool-path",
                str(tool_path),
            ],
            check=False,
        )
        result = run(
            [
                "dotnet",
                "tool",
                "install",
                package_id,
                "--version",
                version,
                "--tool-path",
                str(tool_path),
                "--configfile",
                str(config),
                "--no-http-cache",
            ],
            check=False,
        )
        if result.returncode != 0:
            raise RuntimeError(
                f"dotnet tool install failed for {package_id}:\n{result.stdout}\n{result.stderr}"
            )
        binary = tool_path / command_name
        if not binary.is_file():
            raise RuntimeError(f"missing installed binary: {binary}")
        binaries[command_name] = binary
    return binaries


def install_pdf(dist: Path, venv_dir: Path) -> Path:
    if venv_dir.exists():
        shutil.rmtree(venv_dir)
    run([sys.executable, "-m", "venv", str(venv_dir)])
    pip = venv_dir / ("Scripts" if os.name == "nt" else "bin") / "pip"
    python = venv_dir / ("Scripts" if os.name == "nt" else "bin") / "python"
    wheel = sorted(dist.glob("*.whl"))[-1]
    # Install the local wheel; runtime deps (e.g. pymupdf) may resolve from PyPI.
    result = run(
        [str(pip), "install", str(wheel)],
        check=False,
    )
    if result.returncode != 0:
        raise RuntimeError(f"pip install failed for {wheel}:\n{result.stdout}\n{result.stderr}")
    binary = venv_dir / ("Scripts" if os.name == "nt" else "bin") / "tiwater-pdf"
    if not binary.is_file():
        # console script may resolve via python -m; prefer entrypoint path
        probe = run([str(python), "-c", "import shutil; print(shutil.which('tiwater-pdf') or '')"])
        which = probe.stdout.strip()
        if not which:
            raise RuntimeError("tiwater-pdf entrypoint missing after wheel install")
        binary = Path(which)
    return binary


def capture_surface(binary: Path) -> str:
    chunks: list[str] = []
    invocations = [[], ["--help"], ["-h"]]
    for args in invocations:
        result = run([str(binary), *args], check=False)
        chunks.append(result.stdout or "")
        chunks.append(result.stderr or "")
    return "\n".join(chunks)


def find_forbidden(surface: str) -> list[str]:
    lowered = surface.lower()
    hits: list[str] = []
    for phrase in FORBIDDEN_PHRASES:
        if phrase in lowered:
            hits.append(phrase)
    for word in FORBIDDEN_WORDS:
        if re.search(rf"\b{re.escape(word)}\b", lowered):
            hits.append(word)
    return sorted(set(hits))


def has_command(surface: str, command: str) -> bool:
    pattern = rf"(?m)^\s*(?:\S+\s+)?{re.escape(command)}\b"
    return re.search(pattern, surface) is not None or re.search(
        rf"\b{re.escape(command)}\b", surface
    ) is not None


def has_render_surface(convert_surface: str) -> bool:
    return "-to-pdf" in convert_surface.lower()


def evaluate_surfaces(surfaces: dict[str, str]) -> list[str]:
    failures: list[str] = []
    for command_name, surface in surfaces.items():
        hits = find_forbidden(surface)
        if hits:
            failures.append(
                f"{command_name}: forbidden public-surface term(s): {', '.join(hits)}"
            )
        for required in REQUIRED_COMMANDS.get(command_name, ()):
            if not has_command(surface, required):
                failures.append(
                    f"{command_name}: missing required command {required!r}"
                )
    if "tiwater-convert" in surfaces and not has_render_surface(
        surfaces["tiwater-convert"]
    ):
        failures.append("tiwater-convert: missing render surface (*-to-pdf)")
    expected = {item[0] for item in DOTNET_PACKAGES} | {"tiwater-pdf"}
    missing = sorted(expected - set(surfaces))
    if missing:
        failures.append(f"missing CLI surface(s): {', '.join(missing)}")
    return failures


def ensure_build_module() -> None:
    try:
        import build  # noqa: F401
    except ImportError as error:
        raise RuntimeError(
            "Python package 'build' is required (pip install build)"
        ) from error


def smoke(root: Path, work_dir: Path) -> int:
    ensure_build_module()
    print("cli-surface-smoke: packing .NET tools")
    nuget_sources = pack_dotnet(root)
    print("cli-surface-smoke: building pdf wheel")
    pdf_dist = build_pdf(root)
    tool_path = work_dir / "dotnet-tools"
    print("cli-surface-smoke: installing .NET tools from local nupkg only")
    binaries = install_dotnet_tools(root, tool_path, nuget_sources)
    print("cli-surface-smoke: installing pdf wheel from local dist only")
    binaries["tiwater-pdf"] = install_pdf(pdf_dist, work_dir / "pdf-venv")

    surfaces: dict[str, str] = {}
    for name, binary in binaries.items():
        print(f"cli-surface-smoke: enumerating {name}")
        surfaces[name] = capture_surface(binary)
        # Always execute the bare command and --help/-h paths already done in capture.

    failures = evaluate_surfaces(surfaces)
    if failures:
        print(f"cli-surface-smoke: {len(failures)} failure(s)", file=sys.stderr)
        for failure in failures:
            print(failure, file=sys.stderr)
        return 1
    print("cli-surface-smoke: ok")
    for name, surface in sorted(surfaces.items()):
        commands = sorted(
            {
                match.group(1)
                for match in re.finditer(
                    r"(?m)^\s*(?:tiwater-\S+\s+)?([a-z0-9][a-z0-9\-|<>]+)",
                    surface,
                )
                if match.group(1)
                not in {"usage:", "usage", "options:", "positional", "available"}
            }
        )
        print(f"  {name}: {', '.join(commands[:12])}{'...' if len(commands) > 12 else ''}")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--root",
        type=Path,
        default=None,
        help="Repository root (defaults to parent of scripts/)",
    )
    parser.add_argument(
        "--work-dir",
        type=Path,
        default=None,
        help="Scratch directory for installs (default: temporary)",
    )
    args = parser.parse_args(argv)
    root = (args.root or repo_root()).resolve()
    if args.work_dir is not None:
        work_dir = args.work_dir.resolve()
        work_dir.mkdir(parents=True, exist_ok=True)
        return smoke(root, work_dir)
    with tempfile.TemporaryDirectory(prefix="cli-surface-smoke-") as temporary:
        return smoke(root, Path(temporary))


if __name__ == "__main__":
    raise SystemExit(main())
