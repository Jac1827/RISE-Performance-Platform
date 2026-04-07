#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
import shutil
import tempfile
import zipfile
from pathlib import Path


REQUIRED_DECLARATIONS = {
    "xl/worksheets/sheet2.xml": {
        "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "xmlns:xr2": "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
        "xmlns:xr3": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3",
    },
    "xl/worksheets/sheet3.xml": {
        "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "xmlns:xr2": "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
        "xmlns:xr3": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3",
    },
    "xl/styles.xml": {
        "xmlns:x16r2": "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main",
        "xmlns:xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
    },
}


ROOT_TAG_PATTERNS = {
    "xl/worksheets/sheet2.xml": re.compile(r"(<worksheet\b[^>]*)(>)"),
    "xl/worksheets/sheet3.xml": re.compile(r"(<worksheet\b[^>]*)(>)"),
    "xl/styles.xml": re.compile(r"(<styleSheet\b[^>]*)(>)"),
}


def add_missing_declarations(xml_text: str, *, pattern: re.Pattern[str], declarations: dict[str, str]) -> str:
    match = pattern.search(xml_text)
    if match is None:
        raise ValueError("Could not find document root tag")

    root_tag = match.group(1)
    updated_tag = root_tag
    for attr_name, attr_value in declarations.items():
        if f'{attr_name}="' not in updated_tag:
            updated_tag += f' {attr_name}="{attr_value}"'

    return xml_text[: match.start(1)] + updated_tag + xml_text[match.end(1) :]


def build_backup_path(workbook_path: Path) -> Path:
    candidate = workbook_path.with_name(f"{workbook_path.stem}.namespace-backup{workbook_path.suffix}")
    counter = 2
    while candidate.exists():
        candidate = workbook_path.with_name(f"{workbook_path.stem}.namespace-backup{counter}{workbook_path.suffix}")
        counter += 1
    return candidate


def repair_workbook(workbook_path: Path, *, backup: bool) -> Path | None:
    entries: dict[str, bytes] = {}
    with zipfile.ZipFile(workbook_path, "r") as src:
        for info in src.infolist():
            entries[info.filename] = src.read(info.filename)

    changed = False
    for xml_path, declarations in REQUIRED_DECLARATIONS.items():
        original_text = entries[xml_path].decode("utf-8")
        updated_text = add_missing_declarations(
            original_text,
            pattern=ROOT_TAG_PATTERNS[xml_path],
            declarations=declarations,
        )
        if updated_text != original_text:
            entries[xml_path] = updated_text.encode("utf-8")
            changed = True

    if not changed:
        return None

    backup_path = None
    if backup:
        backup_path = build_backup_path(workbook_path)
        shutil.copy2(workbook_path, backup_path)

    with tempfile.NamedTemporaryFile(dir=workbook_path.parent, suffix=".xlsx", delete=False) as tmp:
        temp_path = Path(tmp.name)

    try:
        with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as dst:
            for filename, data in entries.items():
                dst.writestr(filename, data)
        temp_path.replace(workbook_path)
    finally:
        if temp_path.exists():
            temp_path.unlink()

    return backup_path


def main() -> int:
    parser = argparse.ArgumentParser(description="Repair missing OOXML namespace declarations in an Excel workbook.")
    parser.add_argument("workbook", type=Path, help="Path to the .xlsx workbook")
    parser.add_argument("--backup", action="store_true", help="Save a backup copy before modifying the workbook")
    args = parser.parse_args()

    workbook_path = args.workbook.expanduser().resolve()
    if not workbook_path.exists():
        raise SystemExit(f"Workbook not found: {workbook_path}")

    backup_path = repair_workbook(workbook_path, backup=args.backup)
    if backup_path is None:
        print(f"No namespace changes were needed in {workbook_path}")
        return 0

    print(f"Repaired workbook: {workbook_path}")
    print(f"Backup saved: {backup_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
