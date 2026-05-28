#!/usr/bin/env python3
"""Extract VBA modules from all .xlsm workbooks into build/vba_modules/."""

import argparse
import os
import sys
from pathlib import Path

from oletools.olevba import VBA_Parser

REPO_ROOT = Path(__file__).resolve().parent.parent
OUTPUT_DIR = REPO_ROOT / "build" / "vba_modules"

WORKBOOKS = [
    "Loan amortization.xlsm",
    "Mean variance portofolio analysis Github.xlsm",
    "VBA project.xlsm",
]


def extract(verbose: bool = False) -> dict[str, list[str]]:
    """Extract VBA from each workbook. Returns {workbook: [module_paths]}."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    results: dict[str, list[str]] = {}

    for wb_name in WORKBOOKS:
        wb_path = REPO_ROOT / wb_name
        if not wb_path.exists():
            print(f"WARNING: {wb_name} not found, skipping.", file=sys.stderr)
            continue

        safe_name = wb_name.replace(" ", "_").replace(".xlsm", "")
        wb_out = OUTPUT_DIR / safe_name
        wb_out.mkdir(parents=True, exist_ok=True)

        vba = VBA_Parser(str(wb_path))
        modules: list[str] = []

        for filename, stream_path, vba_filename, vba_code in vba.extract_macros():
            out_file = wb_out / vba_filename
            out_file.write_text(vba_code, encoding="utf-8")
            modules.append(str(out_file))
            if verbose:
                print(f"  Extracted: {safe_name}/{vba_filename} ({len(vba_code.splitlines())} lines)")

        vba.close()
        results[wb_name] = modules
        print(f"Extracted {len(modules)} modules from {wb_name}")

    return results


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract VBA modules from .xlsm workbooks")
    parser.add_argument("--verbose", "-v", action="store_true", help="Show detailed output")
    args = parser.parse_args()

    results = extract(verbose=args.verbose)
    total = sum(len(v) for v in results.values())
    print(f"\nDone: {total} modules extracted from {len(results)} workbooks into {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
