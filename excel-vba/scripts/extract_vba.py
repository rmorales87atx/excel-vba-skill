#!/usr/bin/env python3
"""Read-only VBA extraction. Works on any OS; does not need Excel.

Usage:
    python extract_vba.py <workbook> [--out DIR]

Without --out, prints every module to stdout.
With --out, writes one file per module (Windows-1252, CRLF).
Output keeps Attribute headers. It is reference material only:
never feed it to Import-Vba.ps1.

Requires: pip install oletools
"""
import argparse
import os
import sys

try:
    from oletools.olevba import VBA_Parser
except ImportError:
    sys.exit("MISSING_DEPENDENCY: install oletools (pip install oletools).")


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbook")
    ap.add_argument("--out", help="directory for one file per module")
    args = ap.parse_args()

    if not os.path.isfile(args.workbook):
        print(f"FILE_NOT_FOUND: {args.workbook}", file=sys.stderr)
        return 1

    parser = VBA_Parser(args.workbook)
    try:
        if not parser.detect_vba_macros():
            print("NO_VBA: the file contains no VBA project.")
            return 0

        if args.out:
            os.makedirs(args.out, exist_ok=True)

        seen = set()
        for _container, _stream, vba_filename, code in parser.extract_macros():
            if vba_filename in seen:
                continue
            seen.add(vba_filename)
            if isinstance(code, bytes):
                code = code.decode("cp1252", errors="replace")
            if args.out:
                path = os.path.join(args.out, os.path.basename(vba_filename))
                text = code.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
                with open(path, "w", encoding="cp1252", errors="replace", newline="") as fh:
                    fh.write(text)
                print(f"EXTRACTED: {path}")
            else:
                print(f"===== {vba_filename} =====")
                print(code)
        return 0
    finally:
        parser.close()


if __name__ == "__main__":
    sys.exit(main())
