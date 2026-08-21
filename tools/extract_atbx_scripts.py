"""
extract_atbx_scripts.py
=======================

Refresh the readable Python source mirror in ``src/`` from the ArcGIS Pro
toolbox (``.atbx``) in the repository root.

Why this exists
---------------
An ArcGIS Pro ``.atbx`` toolbox is a ZIP container. The Python that actually
runs the tool lives *inside* that container, which means GitHub treats the
whole toolbox as an opaque binary: the code cannot be viewed, searched,
diffed in pull requests, or indexed by GitHub code search.

This script unpacks the toolbox and writes the embedded scripts out as plain
``.py`` files so the logic is visible to anyone browsing the repo.

The ``.atbx`` remains the authoritative, runnable artifact. The files in
``src/`` are a read-only mirror for humans and search engines.

Usage
-----
Run from the repository root::

    python tools/extract_atbx_scripts.py

Or point it at a specific toolbox and output directory::

    python tools/extract_atbx_scripts.py --atbx "My Toolbox.atbx" --out src

Check whether the mirror is stale without writing anything (useful in CI or
as a pre-commit sanity check; exits non-zero if out of date)::

    python tools/extract_atbx_scripts.py --check

Requires only the Python standard library. Does not require arcpy, and can be
run outside the ArcGIS Pro Python environment.
"""

from __future__ import annotations

import argparse
import sys
import zipfile
from pathlib import Path

# Map the script roles found inside an .atbx tool folder to friendlier
# output filenames. Keys are matched against the end of the archive member path.
SCRIPT_ROLES = {
    "tool.script.execute.py": "grouped_excel_export.py",
    "tool.script.validate.py": "tool_validator.py",
}

BANNER_TEMPLATE = """\
# =============================================================================
# SOURCE MIRROR - READ ONLY REFERENCE COPY
# -----------------------------------------------------------------------------
# Extracted from: {atbx_name}
#   archive path: {member}
#
# Published so the code is readable and searchable on GitHub. The .atbx is the
# runnable artifact -- editing this file does NOT change the tool. Edit the
# script inside the toolbox in ArcGIS Pro, then re-run:
#     python tools/extract_atbx_scripts.py
# =============================================================================

"""


def find_toolbox(root: Path) -> Path:
    """Return the single .atbx in *root*, or raise a helpful error."""
    candidates = sorted(root.glob("*.atbx"))
    if not candidates:
        raise SystemExit(
            f"No .atbx toolbox found in {root}. "
            "Run this from the repository root, or pass --atbx explicitly."
        )
    if len(candidates) > 1:
        names = ", ".join(c.name for c in candidates)
        raise SystemExit(
            f"Multiple toolboxes found ({names}). Pass --atbx to choose one."
        )
    return candidates[0]


def collect_scripts(atbx_path: Path) -> dict[str, tuple[str, str]]:
    """
    Read *atbx_path* and return ``{output_filename: (member_path, source)}``.

    Only members whose names match SCRIPT_ROLES are returned. Any other Python
    found in the toolbox is reported to stderr so it is not silently dropped.
    """
    collected: dict[str, tuple[str, str]] = {}
    unmatched: list[str] = []

    with zipfile.ZipFile(atbx_path) as archive:
        for member in archive.namelist():
            if not member.endswith(".py"):
                continue

            role = next(
                (r for r in SCRIPT_ROLES if member.endswith(r)),
                None,
            )
            if role is None:
                unmatched.append(member)
                continue

            source = archive.read(member).decode("utf-8")
            collected[SCRIPT_ROLES[role]] = (member, source)

    for member in unmatched:
        print(f"  note: skipped unrecognized script -> {member}", file=sys.stderr)

    return collected


def render(atbx_path: Path, member: str, source: str) -> str:
    """Prepend the provenance banner to *source*."""
    banner = BANNER_TEMPLATE.format(atbx_name=atbx_path.name, member=member)
    return banner + source


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Extract embedded Python from an ArcGIS Pro .atbx toolbox."
    )
    parser.add_argument(
        "--atbx",
        type=Path,
        default=None,
        help="Path to the .atbx toolbox (default: the only .atbx in the current directory).",
    )
    parser.add_argument(
        "--out",
        type=Path,
        default=Path("src"),
        help="Directory to write the extracted .py files into (default: src).",
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="Do not write files; exit 1 if the mirror is out of date.",
    )
    args = parser.parse_args(argv)

    atbx_path = args.atbx or find_toolbox(Path.cwd())
    if not atbx_path.is_file():
        raise SystemExit(f"Toolbox not found: {atbx_path}")

    print(f"Reading {atbx_path.name}")
    scripts = collect_scripts(atbx_path)

    if not scripts:
        raise SystemExit("No recognized Python scripts found inside the toolbox.")

    stale: list[str] = []

    for out_name, (member, source) in sorted(scripts.items()):
        rendered = render(atbx_path, member, source)
        target = args.out / out_name

        existing = (
            target.read_text(encoding="utf-8") if target.is_file() else None
        )

        if args.check:
            if existing != rendered:
                stale.append(out_name)
                print(f"  STALE   {target}")
            else:
                print(f"  ok      {target}")
            continue

        args.out.mkdir(parents=True, exist_ok=True)
        target.write_text(rendered, encoding="utf-8")
        status = "unchanged" if existing == rendered else "written"
        line_count = rendered.count("\n") + 1
        print(f"  {status:9} {target}  ({line_count} lines)")

    if args.check and stale:
        print(
            f"\n{len(stale)} file(s) out of date. "
            "Run: python tools/extract_atbx_scripts.py",
            file=sys.stderr,
        )
        return 1

    print("Done.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
