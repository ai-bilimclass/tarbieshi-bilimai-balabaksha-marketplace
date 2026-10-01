#!/usr/bin/env python3
"""Verify, or normalize and verify, landscape orientation in a DOCX."""

import argparse
import sys

from docx import Document
from docx.enum.section import WD_ORIENT


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx")
    parser.add_argument("--fix", action="store_true", help="Set every section to landscape before auditing")
    args = parser.parse_args()
    document = Document(args.docx)
    if args.fix:
        for section in document.sections:
            width, height = section.page_width, section.page_height
            section.orientation = WD_ORIENT.LANDSCAPE
            section.page_width = max(width, height)
            section.page_height = min(width, height)
        document.save(args.docx)
        document = Document(args.docx)
    bad = [index for index, section in enumerate(document.sections, start=1)
           if section.orientation != WD_ORIENT.LANDSCAPE or section.page_width <= section.page_height]
    if bad:
        print(f"FAIL: sections without landscape orientation: {bad}", file=sys.stderr)
        return 1
    print(f"PASS: all {len(document.sections)} sections are landscape")
    return 0


if __name__ == "__main__":
    sys.exit(main())
