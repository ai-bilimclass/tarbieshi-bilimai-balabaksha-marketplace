#!/usr/bin/env python3
"""Set and verify landscape print layout in an OOXML planning workbook."""

from __future__ import annotations

import argparse
import os
import re
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile


NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
TAG = lambda local: f"{{{NS}}}{local}"
SHEET = re.compile(r"xl/worksheets/sheet\d+\.xml$")
ET.register_namespace("", NS)


def patch_sheet(blob: bytes) -> bytes:
    root = ET.fromstring(blob)
    sheet_pr = root.find(TAG("sheetPr"))
    if sheet_pr is None:
        sheet_pr = ET.Element(TAG("sheetPr"))
        root.insert(0, sheet_pr)
    setup_pr = sheet_pr.find(TAG("pageSetUpPr"))
    if setup_pr is None:
        setup_pr = ET.SubElement(sheet_pr, TAG("pageSetUpPr"))
    setup_pr.set("fitToPage", "1")

    setup = root.find(TAG("pageSetup"))
    if setup is None:
        setup = ET.Element(TAG("pageSetup"))
        before = {TAG(name) for name in (
            "headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches",
            "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF",
            "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst")}
        index = next((i for i, child in enumerate(root) if child.tag in before), len(root))
        root.insert(index, setup)
    setup.set("orientation", "landscape")
    setup.set("fitToWidth", "1")
    setup.set("fitToHeight", "0")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def audit(path: Path) -> int:
    count = 0
    with ZipFile(path) as archive:
        for name in archive.namelist():
            if not SHEET.fullmatch(name):
                continue
            count += 1
            root = ET.fromstring(archive.read(name))
            setup = root.find(TAG("pageSetup"))
            if setup is None or setup.get("orientation") != "landscape":
                raise ValueError(f"{name}: требуется альбомная ориентация печати")
    if count == 0:
        raise ValueError("В XLSX не найдено ни одного листа")
    return count


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("xlsx", type=Path)
    args = parser.parse_args()
    source = args.xlsx.resolve()
    fd, temp_name = tempfile.mkstemp(prefix=".landscape-", suffix=".xlsx", dir=source.parent)
    os.close(fd)
    try:
        with ZipFile(source) as original, ZipFile(temp_name, "w") as updated:
            for info in original.infolist():
                blob = original.read(info.filename)
                if SHEET.fullmatch(info.filename):
                    blob = patch_sheet(blob)
                updated.writestr(info, blob)
        audit(Path(temp_name))
        os.replace(temp_name, source)
    finally:
        if os.path.exists(temp_name):
            os.unlink(temp_name)
    print(f"PASS: {audit(source)} XLSX sheet(s) use landscape print orientation")


if __name__ == "__main__":
    main()
