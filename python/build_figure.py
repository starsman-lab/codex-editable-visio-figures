#!/usr/bin/env python3
"""Lightweight Python wrapper for the local Visio paper figure skill."""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPTS_DIR = REPO_ROOT / "scripts"
DEFAULT_SEED = REPO_ROOT / "assets" / "seed-paper-figure.vsdx"


class WrapperError(Exception):
    pass


def decode_output(data: bytes | None) -> str:
    if not data:
        return ""

    for encoding in ("utf-8", "gbk", sys.getfilesystemencoding()):
        try:
            return data.decode(encoding)
        except (UnicodeDecodeError, LookupError):
            continue

    return data.decode("utf-8", errors="replace")


def parse_json_output(stdout: str, script_name: str) -> dict:
    cleaned = stdout.strip()
    if not cleaned:
        return {}

    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        pass

    lines = [line for line in cleaned.splitlines() if line.strip()]
    for start in range(len(lines)):
        candidate = "\n".join(lines[start:])
        try:
            return json.loads(candidate)
        except json.JSONDecodeError:
            continue

    raise WrapperError(f"Expected JSON output from {script_name}, got: {cleaned}")


def normalize_formats(values: list[str] | None) -> list[str]:
    if not values:
        return []

    normalized: list[str] = []
    for value in values:
        parts = [part.strip() for part in value.split(",")]
        normalized.extend(part for part in parts if part)
    return normalized


def run_powershell(script: Path, args: list[str]) -> dict:
    cmd = [
        "powershell",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(script),
        *args,
    ]
    result = subprocess.run(cmd, capture_output=True)
    stdout = decode_output(result.stdout)
    stderr = decode_output(result.stderr)
    if result.returncode != 0:
        raise WrapperError(stderr.strip() or stdout.strip() or f"Command failed: {script.name}")
    return parse_json_output(stdout, script.name)


def cmd_probe(_: argparse.Namespace) -> int:
    data = run_powershell(SCRIPTS_DIR / "visio_probe.ps1", [])
    print(json.dumps(data, indent=2))
    return 0


def cmd_export(args: argparse.Namespace) -> int:
    vsdx = Path(args.vsdx).resolve()
    output_dir = Path(args.output_dir).resolve() if args.output_dir else vsdx.parent
    formats = ",".join(normalize_formats(args.formats))
    data = run_powershell(
        SCRIPTS_DIR / "visio_export.ps1",
        [
            "-VsdxPath",
            str(vsdx),
            "-OutputDir",
            str(output_dir),
            "-Formats",
            formats,
            *( ["-PageName", args.page_name] if args.page_name else [] ),
        ],
    )
    print(json.dumps(data, indent=2))
    return 0


def cmd_build(args: argparse.Namespace) -> int:
    spec = Path(args.spec).resolve()
    target = Path(args.output).resolve()
    seed = Path(args.seed).resolve() if args.seed else DEFAULT_SEED.resolve()
    output_dir = Path(args.output_dir).resolve() if args.output_dir else target.parent / "exports"
    formats = ",".join(normalize_formats(args.formats))

    ps_args = [
        "-SpecPath",
        str(spec),
        "-VsdxPath",
        str(target),
        "-SeedVsdxPath",
        str(seed),
        "-OutputDir",
        str(output_dir),
    ]
    if formats:
        ps_args.extend(["-ExportFormats", formats])
    if args.replace_page:
        ps_args.append("-ReplacePage")

    data = run_powershell(SCRIPTS_DIR / "visio_apply_spec.ps1", ps_args)
    print(json.dumps(data, indent=2))
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Python wrapper for the local Visio paper figure skill.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    probe = subparsers.add_parser("probe", help="Probe local Visio COM availability")
    probe.set_defaults(func=cmd_probe)

    export = subparsers.add_parser("export", help="Export PNG/SVG/PDF from an existing VSDX")
    export.add_argument("--vsdx", required=True, help="Path to the source .vsdx file")
    export.add_argument("--output-dir", help="Directory for exported files")
    export.add_argument("--formats", nargs="+", default=["png"], help="Export formats, e.g. png pdf")
    export.add_argument("--page-name", help="Optional page name")
    export.set_defaults(func=cmd_export)

    build = subparsers.add_parser("build", help="Build or revise a figure from a JSON spec")
    build.add_argument("--spec", required=True, help="Path to the JSON spec file")
    build.add_argument("--output", required=True, help="Path to the target .vsdx file")
    build.add_argument("--seed", help="Optional seed .vsdx file; defaults to assets/seed-paper-figure.vsdx")
    build.add_argument("--output-dir", help="Directory for exported files")
    build.add_argument("--formats", nargs="*", default=["png", "pdf"], help="Export formats, e.g. png pdf")
    build.add_argument("--replace-page", action="store_true", help="Delete existing shapes before drawing")
    build.set_defaults(func=cmd_build)

    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    try:
        return args.func(args)
    except WrapperError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
