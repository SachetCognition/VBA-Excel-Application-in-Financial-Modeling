#!/usr/bin/env python3
"""Basic VBA linting: check for common issues in extracted VBA modules."""

import re
import sys
from pathlib import Path

from oletools.olevba import VBA_Parser

REPO_ROOT = Path(__file__).resolve().parent.parent

WORKBOOKS = [
    "Loan amortization.xlsm",
    "Mean variance portofolio analysis Github.xlsm",
    "VBA project.xlsm",
]

BALANCED_PAIRS = [
    ("Sub ", "End Sub"),
    ("Function ", "End Function"),
    ("If ", "End If"),
    ("For ", "Next"),
]


def lint_module(vba_code: str, module_name: str) -> list[str]:
    """Run basic lint checks on a VBA module. Returns list of warnings."""
    warnings: list[str] = []
    lines = vba_code.splitlines()

    for open_kw, close_kw in BALANCED_PAIRS:
        opens = sum(
            1 for line in lines
            if re.match(rf"^\s*(Public |Private )?{re.escape(open_kw)}", line)
        )
        closes = sum(1 for line in lines if line.strip().startswith(close_kw))
        if opens != closes:
            warnings.append(
                f"{module_name}: Unbalanced '{open_kw.strip()}' blocks "
                f"(opened {opens}, closed {closes})"
            )

    for i, line in enumerate(lines, 1):
        if len(line) > 500:
            warnings.append(f"{module_name}:{i}: Line exceeds 500 characters ({len(line)})")

    return warnings


def main() -> int:
    all_warnings: list[str] = []

    for wb_name in WORKBOOKS:
        wb_path = REPO_ROOT / wb_name
        if not wb_path.exists():
            print(f"WARNING: {wb_name} not found, skipping.", file=sys.stderr)
            continue

        vba = VBA_Parser(str(wb_path))
        for filename, stream_path, vba_filename, vba_code in vba.extract_macros():
            warnings = lint_module(vba_code, f"{wb_name}/{vba_filename}")
            all_warnings.extend(warnings)
        vba.close()

    if all_warnings:
        print(f"VBA lint found {len(all_warnings)} warning(s):")
        for w in all_warnings:
            print(f"  {w}")
        return 1

    print("VBA lint: all checks passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
