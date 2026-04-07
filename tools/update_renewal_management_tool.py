#!/usr/bin/env python3
from __future__ import annotations

import argparse
import copy
import re
import shutil
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
X15_NS = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
XR_NS = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
XR6_NS = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6"
XR10_NS = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10"
XR2_NS = "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"
XCALCF_NS = "http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures"
LOEXT_NS = "http://schemas.libreoffice.org/"
X14AC_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"


ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", REL_NS)
ET.register_namespace("mc", MC_NS)
ET.register_namespace("x15", X15_NS)
ET.register_namespace("xr", XR_NS)
ET.register_namespace("xr6", XR6_NS)
ET.register_namespace("xr10", XR10_NS)
ET.register_namespace("xr2", XR2_NS)
ET.register_namespace("xcalcf", XCALCF_NS)
ET.register_namespace("loext", LOEXT_NS)
ET.register_namespace("x14ac", X14AC_NS)


TARGET_SHEETS = ("xl/worksheets/sheet2.xml", "xl/worksheets/sheet3.xml")
TARGET_DIMENSION = "A1:AB60"
HEADER_ROW_RANGE = "1:28"
VALIDATION_LIST = '"Investor Override,Offer 1,Offer 2,Offer 3"'
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

HEADER_STYLE = "34"
INPUT_HINT_STYLE = "35"
AUTO_HINT_STYLE = "36"
TEXT_INPUT_STYLE = "18"
NAME_INPUT_STYLE = "20"
DATE_INPUT_STYLE = "37"
DATE_AUTO_STYLE = "38"
DEPOSIT_INPUT_STYLE = "39"
RATE_INPUT_STYLE = "19"
LOOKUP_TEXT_STYLE = "40"
CURRENCY_STYLE = "41"
PERCENT_STYLE = "42"
AVG_LABEL_STYLE = "43"
AVG_PERCENT_STYLE = "44"
FORMULA_TYPE_STRING = "str"

LAYOUT_ORIGINAL = "original"
LAYOUT_RIGHT_OVERRIDE = "right_override"
LAYOUT_INVESTOR = "investor"
LAYOUT_SELECTED = "selected"

SOURCE_COLUMN_INDEX = {
    LAYOUT_ORIGINAL: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, None, None, 11, 12, 13, None, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, None],
    LAYOUT_RIGHT_OVERRIDE: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 25, 26, 11, 12, 13, None, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, None],
    LAYOUT_INVESTOR: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, None, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, None],
    LAYOUT_SELECTED: list(range(1, 29)),
}

INPUT_SOURCE_MAP = {
    LAYOUT_ORIGINAL: {"override_pct": None, "selected_offer": None, "renewal": "N", "transfer": "O", "ntv": "P", "phone": "Q", "notes": "R"},
    LAYOUT_RIGHT_OVERRIDE: {"override_pct": "Y", "selected_offer": None, "renewal": "N", "transfer": "O", "ntv": "P", "phone": "Q", "notes": "R"},
    LAYOUT_INVESTOR: {"override_pct": "K", "selected_offer": None, "renewal": "P", "transfer": "Q", "ntv": "R", "phone": "S", "notes": "T"},
    LAYOUT_SELECTED: {"override_pct": "K", "selected_offer": "P", "renewal": "Q", "transfer": "R", "ntv": "S", "phone": "T", "notes": "U"},
}

HEADER_LABELS = {
    "A": "Name",
    "B": "Unit",
    "C": "Unit\nType",
    "D": "Expiration\nDate",
    "E": "90-Day\nNotice",
    "F": "60-Day\nNotice",
    "G": "30-Day\nNotice",
    "H": "Deposit\nHeld",
    "I": "Current\nRate",
    "J": "Recommended\nOffer",
    "K": "Investor Override\n% Increase",
    "L": "Investor Override\nOffer",
    "M": "Offer 1\nConservative",
    "N": "Offer 2\nBalanced",
    "O": "Offer 3\nAggressive",
    "P": "Signed\nOffer",
    "Q": "Renewal\nSigned",
    "R": "Transfer",
    "S": "NTV\nRcv'd",
    "T": "Phone",
    "U": "Notes",
    "V": "Market\nRate",
    "W": "Budget",
    "X": "Occupancy\n(Under/Over)",
    "Y": "Rent Growth\nOffer 1",
    "Z": "Rent Growth\nOffer 2",
    "AA": "Rent Growth\nOffer 3",
    "AB": "Signed\nRent Growth",
}

HEADER_HINTS = {
    "A": ("▼ Enter", INPUT_HINT_STYLE),
    "B": ("▼ Enter", INPUT_HINT_STYLE),
    "C": ("▼ Enter", INPUT_HINT_STYLE),
    "D": ("▼ Enter", INPUT_HINT_STYLE),
    "E": ("◀ Auto", AUTO_HINT_STYLE),
    "F": ("◀ Auto", AUTO_HINT_STYLE),
    "G": ("◀ Auto", AUTO_HINT_STYLE),
    "H": ("▼ Enter", INPUT_HINT_STYLE),
    "I": ("▼ Enter", INPUT_HINT_STYLE),
    "J": ("◀ Auto", AUTO_HINT_STYLE),
    "K": ("▼ Enter", INPUT_HINT_STYLE),
    "L": ("◀ Auto", AUTO_HINT_STYLE),
    "M": ("◀ Auto", AUTO_HINT_STYLE),
    "N": ("◀ Auto", AUTO_HINT_STYLE),
    "O": ("◀ Auto", AUTO_HINT_STYLE),
    "P": ("▼ Select", INPUT_HINT_STYLE),
    "Q": ("▼ Enter", INPUT_HINT_STYLE),
    "R": ("▼ Enter", INPUT_HINT_STYLE),
    "S": ("▼ Enter", INPUT_HINT_STYLE),
    "T": ("▼ Enter", INPUT_HINT_STYLE),
    "U": ("▼ Enter", INPUT_HINT_STYLE),
    "V": ("◀ Auto", AUTO_HINT_STYLE),
    "W": ("◀ Auto", AUTO_HINT_STYLE),
    "X": ("◀ Auto", AUTO_HINT_STYLE),
    "Y": ("◀ Auto", AUTO_HINT_STYLE),
    "Z": ("◀ Auto", AUTO_HINT_STYLE),
    "AA": ("◀ Auto", AUTO_HINT_STYLE),
    "AB": ("◀ Auto", AUTO_HINT_STYLE),
}

DEFAULT_INPUT_STYLES = {
    "A": NAME_INPUT_STYLE,
    "B": TEXT_INPUT_STYLE,
    "C": TEXT_INPUT_STYLE,
    "D": DATE_INPUT_STYLE,
    "H": DEPOSIT_INPUT_STYLE,
    "I": RATE_INPUT_STYLE,
    "P": TEXT_INPUT_STYLE,
    "Q": TEXT_INPUT_STYLE,
    "R": TEXT_INPUT_STYLE,
    "S": TEXT_INPUT_STYLE,
    "T": TEXT_INPUT_STYLE,
    "U": TEXT_INPUT_STYLE,
}


def qn(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


def parse_xml(data: bytes) -> ET.Element:
    return ET.fromstring(data)


def xml_bytes(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def restore_required_namespace_declarations(xml_path: str, xml_data: bytes) -> bytes:
    declarations = REQUIRED_DECLARATIONS.get(xml_path)
    pattern = ROOT_TAG_PATTERNS.get(xml_path)
    if declarations is None or pattern is None:
        return xml_data

    xml_text = xml_data.decode("utf-8")
    match = pattern.search(xml_text)
    if match is None:
        return xml_data

    root_tag = match.group(1)
    updated_tag = root_tag
    for attr_name, attr_value in declarations.items():
        if f'{attr_name}="' not in updated_tag:
            updated_tag += f' {attr_name}="{attr_value}"'

    if updated_tag == root_tag:
        return xml_data

    updated_text = xml_text[: match.start(1)] + updated_tag + xml_text[match.end(1) :]
    return updated_text.encode("utf-8")


def find_row(sheet_root: ET.Element, row_number: int) -> ET.Element:
    sheet_data = sheet_root.find(qn(MAIN_NS, "sheetData"))
    if sheet_data is None:
        raise ValueError("Worksheet is missing sheetData")
    for row in sheet_data.findall(qn(MAIN_NS, "row")):
        if row.attrib.get("r") == str(row_number):
            return row
    raise ValueError(f"Worksheet is missing row {row_number}")


def find_cell(row: ET.Element, ref: str) -> ET.Element | None:
    for cell in row.findall(qn(MAIN_NS, "c")):
        if cell.attrib.get("r") == ref:
            return cell
    return None


def set_row_cells(row: ET.Element, cells: list[ET.Element]) -> None:
    for cell in list(row.findall(qn(MAIN_NS, "c"))):
        row.remove(cell)
    for cell in cells:
        row.append(cell)


def inline_string_cell(ref: str, style: str, text: str) -> ET.Element:
    cell = ET.Element(qn(MAIN_NS, "c"), {"r": ref, "s": style, "t": "inlineStr"})
    is_el = ET.SubElement(cell, qn(MAIN_NS, "is"))
    text_el = ET.SubElement(is_el, qn(MAIN_NS, "t"))
    text_el.text = text
    return cell


def blank_cell(ref: str, style: str) -> ET.Element:
    return ET.Element(qn(MAIN_NS, "c"), {"r": ref, "s": style})


def clone_cell(source_cell: ET.Element | None, target_ref: str, fallback_style: str | None = None) -> ET.Element:
    if source_cell is None:
        if fallback_style is None:
            raise ValueError(f"Missing source cell for {target_ref}")
        return blank_cell(target_ref, fallback_style)
    cloned = copy.deepcopy(source_cell)
    cloned.attrib["r"] = target_ref
    return cloned


def shared_formula_cell(
    ref: str,
    style: str,
    shared_index: str,
    *,
    formula: str | None = None,
    shared_range: str | None = None,
    cell_type: str | None = None,
) -> ET.Element:
    attrs = {"r": ref, "s": style}
    if cell_type is not None:
        attrs["t"] = cell_type
    cell = ET.Element(qn(MAIN_NS, "c"), attrs)
    formula_attrs = {"t": "shared", "si": shared_index}
    if shared_range is not None:
        formula_attrs["ref"] = shared_range
    formula_el = ET.SubElement(cell, qn(MAIN_NS, "f"), formula_attrs)
    if formula is not None:
        formula_el.text = formula
    return cell


def plain_formula_cell(ref: str, style: str, formula: str, *, cell_type: str | None = None) -> ET.Element:
    attrs = {"r": ref, "s": style}
    if cell_type is not None:
        attrs["t"] = cell_type
    cell = ET.Element(qn(MAIN_NS, "c"), attrs)
    formula_el = ET.SubElement(cell, qn(MAIN_NS, "f"))
    formula_el.text = formula
    return cell


def ensure_percent_input_style(styles_root: ET.Element) -> str:
    cell_xfs = styles_root.find(qn(MAIN_NS, "cellXfs"))
    if cell_xfs is None:
        raise ValueError("styles.xml is missing cellXfs")

    for index, xf in enumerate(cell_xfs.findall(qn(MAIN_NS, "xf"))):
        alignment = xf.find(qn(MAIN_NS, "alignment"))
        if (
            xf.attrib.get("numFmtId") == "167"
            and xf.attrib.get("fontId") == "4"
            and xf.attrib.get("fillId") == "4"
            and xf.attrib.get("borderId") == "2"
            and xf.attrib.get("applyNumberFormat") == "1"
            and alignment is not None
            and alignment.attrib.get("horizontal") == "center"
            and alignment.attrib.get("vertical") == "center"
        ):
            return str(index)

    new_xf = ET.Element(
        qn(MAIN_NS, "xf"),
        {
            "numFmtId": "167",
            "fontId": "4",
            "fillId": "4",
            "borderId": "2",
            "xfId": "0",
            "applyNumberFormat": "1",
            "applyFont": "1",
            "applyFill": "1",
            "applyBorder": "1",
            "applyAlignment": "1",
        },
    )
    ET.SubElement(new_xf, qn(MAIN_NS, "alignment"), {"horizontal": "center", "vertical": "center"})
    cell_xfs.append(new_xf)
    cell_xfs.attrib["count"] = str(len(cell_xfs.findall(qn(MAIN_NS, "xf"))))
    return str(len(cell_xfs.findall(qn(MAIN_NS, "xf"))) - 1)


def detect_layout(sheet_root: ET.Element) -> str:
    row2 = find_row(sheet_root, 2)
    if find_cell(row2, "AB2") is not None:
        return LAYOUT_SELECTED
    if find_cell(row2, "Y2") is not None:
        return LAYOUT_RIGHT_OVERRIDE
    if find_cell(row2, "K2") is not None and find_cell(row2, "L2") is not None:
        k2 = find_cell(row2, "K2")
        l2 = find_cell(row2, "L2")
        if k2 is not None and l2 is not None and k2.attrib.get("t") == "inlineStr" and l2.attrib.get("t") == "inlineStr":
            return LAYOUT_INVESTOR
    return LAYOUT_ORIGINAL


def update_column_layout(sheet_root: ET.Element, layout: str) -> None:
    cols = sheet_root.find(qn(MAIN_NS, "cols"))
    if cols is None:
        raise ValueError("Worksheet is missing cols")

    per_column_attrs: dict[int, dict[str, str]] = {}
    for col in cols.findall(qn(MAIN_NS, "col")):
        attrs = {key: value for key, value in col.attrib.items() if key not in {"min", "max"}}
        for col_index in range(int(col.attrib["min"]), int(col.attrib["max"]) + 1):
            per_column_attrs[col_index] = attrs

    default_widths = {
        11: {"width": "16", "customWidth": "1"},
        12: {"width": "15", "customWidth": "1"},
        16: {"width": "16", "customWidth": "1"},
        28: {"width": "18", "customWidth": "1"},
    }

    for col in list(cols.findall(qn(MAIN_NS, "col"))):
        cols.remove(col)

    for target_index, source_index in enumerate(SOURCE_COLUMN_INDEX[layout], start=1):
        attrs = dict(per_column_attrs.get(source_index, default_widths.get(target_index, {"width": "12", "customWidth": "1"})))
        attrs["min"] = str(target_index)
        attrs["max"] = str(target_index)
        cols.append(ET.Element(qn(MAIN_NS, "col"), attrs))


def rebuild_header_rows(sheet_root: ET.Element) -> None:
    row2 = find_row(sheet_root, 2)
    row3 = find_row(sheet_root, 3)

    set_row_cells(
        row2,
        [inline_string_cell(f"{col}2", HEADER_STYLE, label) for col, label in HEADER_LABELS.items()],
    )
    set_row_cells(
        row3,
        [inline_string_cell(f"{col}3", style, text) for col, (text, style) in HEADER_HINTS.items()],
    )


def notice_formula(row_number: int, days: int) -> str:
    return f'IF(D{row_number}="","",D{row_number}-{days})'


def recommended_formula(row_number: int) -> str:
    return (
        f'IF(I{row_number}="","",'
        f'IF(ABS(N{row_number}-W{row_number})<=ABS(M{row_number}-W{row_number}),'
        f'IF(ABS(N{row_number}-W{row_number})<=ABS(O{row_number}-W{row_number}),"Offer 2","Offer 3"),'
        f'IF(ABS(M{row_number}-W{row_number})<=ABS(O{row_number}-W{row_number}),"Offer 1","Offer 3")))'
    )


def override_offer_formula(row_number: int) -> str:
    return f'IF(OR(I{row_number}="",K{row_number}=""),"",CEILING(I{row_number}*(1+K{row_number}),5))'


def offer_1_formula(row_number: int) -> str:
    return (
        f'IF(I{row_number}="","",'
        f'IF(X{row_number}="Under",'
        f'CEILING(MAX(I{row_number}*1.02,MIN(V{row_number},W{row_number})*0.95),5),'
        f'CEILING(MAX(I{row_number}*1.03,MIN(V{row_number},W{row_number})*0.97),5)))'
    )


def offer_2_formula(row_number: int) -> str:
    return (
        f'IF(I{row_number}="","",'
        f'IF(X{row_number}="Under",'
        f'CEILING(MAX(I{row_number}*1.03,(V{row_number}+W{row_number})/2*0.96),5),'
        f'CEILING(MAX(I{row_number}*1.05,(V{row_number}+W{row_number})/2*0.98),5)))'
    )


def offer_3_formula(row_number: int) -> str:
    return (
        f'IF(I{row_number}="","",'
        f'IF(X{row_number}="Under",'
        f'CEILING(MAX(I{row_number}*1.05,MIN(V{row_number},W{row_number}*1.01)),5),'
        f'CEILING(MAX(I{row_number}*1.08,MIN(V{row_number},W{row_number}*1.02)),5)))'
    )


def market_formula(row_number: int) -> str:
    return f'IF(C{row_number}="","",IFERROR(VLOOKUP(C{row_number},SETUP!$A$9:$D$18,2,0),"⚠ Check SETUP"))'


def budget_formula(row_number: int) -> str:
    return f'IF(C{row_number}="","",IFERROR(VLOOKUP(C{row_number},SETUP!$A$9:$D$18,3,0),"⚠ Check SETUP"))'


def occupancy_formula(row_number: int) -> str:
    return f'IF(C{row_number}="","",IFERROR(VLOOKUP(C{row_number},SETUP!$A$9:$D$18,4,0),"⚠ Check SETUP"))'


def growth_formula(row_number: int, offer_col: str) -> str:
    return f'IF(OR(I{row_number}="",{offer_col}{row_number}=""),"",({offer_col}{row_number}-I{row_number})/I{row_number})'


def selected_growth_formula(row_number: int) -> str:
    return (
        f'IF(OR(Q{row_number}="",S{row_number}<>"",P{row_number}="",I{row_number}=""),"",'
        f'IF(P{row_number}="Investor Override",IF(L{row_number}="","",(L{row_number}-I{row_number})/I{row_number}),'
        f'IF(P{row_number}="Offer 1",IF(M{row_number}="","",(M{row_number}-I{row_number})/I{row_number}),'
        f'IF(P{row_number}="Offer 2",IF(N{row_number}="","",(N{row_number}-I{row_number})/I{row_number}),'
        f'IF(P{row_number}="Offer 3",IF(O{row_number}="","",(O{row_number}-I{row_number})/I{row_number}),"")))))'
    )


def rebuild_data_rows(sheet_root: ET.Element, layout: str, percent_input_style: str, template_mode: bool) -> None:
    numeric_formula_type = FORMULA_TYPE_STRING if template_mode else None
    string_formula_type = FORMULA_TYPE_STRING
    sources = INPUT_SOURCE_MAP[layout]

    for row_number in range(4, 34):
        row = find_row(sheet_root, row_number)
        cells: list[ET.Element] = []

        cells.append(clone_cell(find_cell(row, f"A{row_number}"), f"A{row_number}", NAME_INPUT_STYLE))
        cells.append(clone_cell(find_cell(row, f"B{row_number}"), f"B{row_number}", TEXT_INPUT_STYLE))
        cells.append(clone_cell(find_cell(row, f"C{row_number}"), f"C{row_number}", TEXT_INPUT_STYLE))
        cells.append(clone_cell(find_cell(row, f"D{row_number}"), f"D{row_number}", DATE_INPUT_STYLE))

        cells.append(
            shared_formula_cell(
                f"E{row_number}",
                DATE_AUTO_STYLE,
                "0",
                formula=notice_formula(row_number, 90) if row_number == 4 else None,
                shared_range="E4:E33" if row_number == 4 else None,
                cell_type=numeric_formula_type,
            )
        )
        cells.append(
            shared_formula_cell(
                f"F{row_number}",
                DATE_AUTO_STYLE,
                "1",
                formula=notice_formula(row_number, 60) if row_number == 4 else None,
                shared_range="F4:F33" if row_number == 4 else None,
                cell_type=numeric_formula_type,
            )
        )
        cells.append(
            shared_formula_cell(
                f"G{row_number}",
                DATE_AUTO_STYLE,
                "2",
                formula=notice_formula(row_number, 30) if row_number == 4 else None,
                shared_range="G4:G33" if row_number == 4 else None,
                cell_type=numeric_formula_type,
            )
        )

        cells.append(clone_cell(find_cell(row, f"H{row_number}"), f"H{row_number}", DEPOSIT_INPUT_STYLE))
        cells.append(clone_cell(find_cell(row, f"I{row_number}"), f"I{row_number}", RATE_INPUT_STYLE))

        cells.append(
            shared_formula_cell(
                f"J{row_number}",
                LOOKUP_TEXT_STYLE,
                "3",
                formula=recommended_formula(row_number) if row_number == 4 else None,
                shared_range="J4:J33" if row_number == 4 else None,
                cell_type=string_formula_type,
            )
        )

        override_source = find_cell(row, f"{sources['override_pct']}{row_number}") if sources["override_pct"] is not None else None
        selected_offer_source = find_cell(row, f"{sources['selected_offer']}{row_number}") if sources["selected_offer"] is not None else None

        cells.append(clone_cell(override_source, f"K{row_number}", percent_input_style))
        cells.append(
            shared_formula_cell(
                f"L{row_number}",
                CURRENCY_STYLE,
                "10",
                formula=override_offer_formula(row_number) if row_number == 4 else None,
                shared_range="L4:L33" if row_number == 4 else None,
                cell_type=numeric_formula_type,
            )
        )
        cells.append(
            shared_formula_cell(
                f"M{row_number}",
                CURRENCY_STYLE,
                "4",
                formula=offer_1_formula(row_number) if row_number == 4 else None,
                shared_range="M4:M33" if row_number == 4 else None,
                cell_type=numeric_formula_type,
            )
        )
        cells.append(
            shared_formula_cell(
                f"N{row_number}",
                CURRENCY_STYLE,
                "5",
                formula=offer_2_formula(row_number) if row_number == 4 else None,
                shared_range="N4:N33" if row_number == 4 else None,
                cell_type=numeric_formula_type,
            )
        )
        cells.append(
            shared_formula_cell(
                f"O{row_number}",
                CURRENCY_STYLE,
                "6",
                formula=offer_3_formula(row_number) if row_number == 4 else None,
                shared_range="O4:O33" if row_number == 4 else None,
                cell_type=numeric_formula_type,
            )
        )

        cells.append(clone_cell(selected_offer_source, f"P{row_number}", TEXT_INPUT_STYLE))
        cells.append(clone_cell(find_cell(row, f"{sources['renewal']}{row_number}"), f"Q{row_number}", TEXT_INPUT_STYLE))
        cells.append(clone_cell(find_cell(row, f"{sources['transfer']}{row_number}"), f"R{row_number}", TEXT_INPUT_STYLE))
        cells.append(clone_cell(find_cell(row, f"{sources['ntv']}{row_number}"), f"S{row_number}", TEXT_INPUT_STYLE))
        cells.append(clone_cell(find_cell(row, f"{sources['phone']}{row_number}"), f"T{row_number}", TEXT_INPUT_STYLE))
        cells.append(clone_cell(find_cell(row, f"{sources['notes']}{row_number}"), f"U{row_number}", TEXT_INPUT_STYLE))

        cells.append(plain_formula_cell(f"V{row_number}", CURRENCY_STYLE, market_formula(row_number), cell_type=numeric_formula_type))
        cells.append(plain_formula_cell(f"W{row_number}", CURRENCY_STYLE, budget_formula(row_number), cell_type=numeric_formula_type))
        cells.append(plain_formula_cell(f"X{row_number}", LOOKUP_TEXT_STYLE, occupancy_formula(row_number), cell_type=string_formula_type))

        cells.append(
            shared_formula_cell(
                f"Y{row_number}",
                PERCENT_STYLE,
                "7",
                formula=growth_formula(row_number, "M") if row_number == 4 else None,
                shared_range="Y4:Y33" if row_number == 4 else None,
                cell_type=numeric_formula_type,
            )
        )
        cells.append(
            shared_formula_cell(
                f"Z{row_number}",
                PERCENT_STYLE,
                "8",
                formula=growth_formula(row_number, "N") if row_number == 4 else None,
                shared_range="Z4:Z33" if row_number == 4 else None,
                cell_type=numeric_formula_type,
            )
        )
        cells.append(
            shared_formula_cell(
                f"AA{row_number}",
                PERCENT_STYLE,
                "9",
                formula=growth_formula(row_number, "O") if row_number == 4 else None,
                shared_range="AA4:AA33" if row_number == 4 else None,
                cell_type=numeric_formula_type,
            )
        )
        cells.append(
            shared_formula_cell(
                f"AB{row_number}",
                PERCENT_STYLE,
                "11",
                formula=selected_growth_formula(row_number) if row_number == 4 else None,
                shared_range="AB4:AB33" if row_number == 4 else None,
                cell_type=numeric_formula_type,
            )
        )

        set_row_cells(row, cells)


def rebuild_portfolio_average_row(sheet_root: ET.Element, template_mode: bool) -> None:
    row = find_row(sheet_root, 34)
    formula_type = FORMULA_TYPE_STRING if template_mode else None
    set_row_cells(
        row,
        [
            inline_string_cell("X34", AVG_LABEL_STYLE, "Portfolio Avg\n(Excl. NTV) →"),
            plain_formula_cell("Y34", AVG_PERCENT_STYLE, 'IFERROR(AVERAGEIFS(Y4:Y33,S4:S33,""),"")', cell_type=formula_type),
            plain_formula_cell("Z34", AVG_PERCENT_STYLE, 'IFERROR(AVERAGEIFS(Z4:Z33,S4:S33,""),"")', cell_type=formula_type),
            plain_formula_cell("AA34", AVG_PERCENT_STYLE, 'IFERROR(AVERAGEIFS(AA4:AA33,S4:S33,""),"")', cell_type=formula_type),
            plain_formula_cell("AB34", AVG_PERCENT_STYLE, 'IFERROR(AVERAGEIF(AB4:AB33,"<>"),"")', cell_type=formula_type),
        ],
    )


def update_summary_ranges(sheet_root: ET.Element) -> None:
    updated_formulas = {
        39: "COUNTA(Q4:Q33)",
        40: "COUNTA(S4:S33)",
        41: "COUNTA(R4:R33)",
    }
    for row_number, formula in updated_formulas.items():
        row = find_row(sheet_root, row_number)
        cell = find_cell(row, f"B{row_number}")
        if cell is None:
            raise ValueError(f"Missing B{row_number}")
        formula_el = cell.find(qn(MAIN_NS, "f"))
        if formula_el is None:
            formula_el = ET.SubElement(cell, qn(MAIN_NS, "f"))
        formula_el.text = formula
        value_el = cell.find(qn(MAIN_NS, "v"))
        if value_el is not None:
            cell.remove(value_el)

    conditional_formatting = sheet_root.find(qn(MAIN_NS, "conditionalFormatting"))
    if conditional_formatting is None:
        raise ValueError("Worksheet is missing conditional formatting")
    conditional_formatting.attrib["sqref"] = "A4:U33"

    formulas = ["$S4<>\"\"", "$R4<>\"\"", "$Q4<>\"\""]
    rules = conditional_formatting.findall(qn(MAIN_NS, "cfRule"))
    if len(rules) != len(formulas):
        raise ValueError("Unexpected conditional formatting rule count")
    for rule, formula in zip(rules, formulas):
        formula_el = rule.find(qn(MAIN_NS, "formula"))
        if formula_el is None:
            formula_el = ET.SubElement(rule, qn(MAIN_NS, "formula"))
        formula_el.text = formula


def ensure_selected_offer_validation(sheet_root: ET.Element) -> None:
    existing = sheet_root.find(qn(MAIN_NS, "dataValidations"))
    if existing is not None:
        sheet_root.remove(existing)

    data_validations = ET.Element(qn(MAIN_NS, "dataValidations"), {"count": "1"})
    validation = ET.SubElement(
        data_validations,
        qn(MAIN_NS, "dataValidation"),
        {
            "type": "list",
            "allowBlank": "1",
            "showInputMessage": "1",
            "showErrorMessage": "1",
            "sqref": "P4:P33",
        },
    )
    formula1 = ET.SubElement(validation, qn(MAIN_NS, "formula1"))
    formula1.text = VALIDATION_LIST

    children = list(sheet_root)
    insertion_index = len(children)
    for idx, child in enumerate(children):
        local_name = child.tag.split("}")[-1]
        if local_name == "pageMargins":
            insertion_index = idx
            break
    sheet_root.insert(insertion_index, data_validations)


def reshape_sheet(sheet_root: ET.Element, percent_input_style: str, template_mode: bool) -> None:
    layout = detect_layout(sheet_root)

    dimension = sheet_root.find(qn(MAIN_NS, "dimension"))
    if dimension is not None:
        dimension.attrib["ref"] = TARGET_DIMENSION

    update_column_layout(sheet_root, layout)

    merge_cells = sheet_root.find(qn(MAIN_NS, "mergeCells"))
    if merge_cells is not None:
        for merge_cell in merge_cells.findall(qn(MAIN_NS, "mergeCell")):
            if merge_cell.attrib.get("ref") in {"A1:X1", "A1:Z1", "A1:AB1"}:
                merge_cell.attrib["ref"] = "A1:AB1"

    for selection in sheet_root.findall(f".//{qn(MAIN_NS, 'selection')}"):
        if selection.attrib.get("sqref") in {"A1:X1", "A1:Z1", "A1:AB1"}:
            selection.attrib["sqref"] = "A1:AB1"

    for row_number in range(1, 35):
        find_row(sheet_root, row_number).attrib["spans"] = HEADER_ROW_RANGE

    rebuild_header_rows(sheet_root)
    rebuild_data_rows(sheet_root, layout, percent_input_style, template_mode)
    rebuild_portfolio_average_row(sheet_root, template_mode)
    update_summary_ranges(sheet_root)
    ensure_selected_offer_validation(sheet_root)


def mark_workbook_for_recalc(workbook_xml: bytes) -> bytes:
    text = workbook_xml.decode("utf-8")

    def upsert_attr(tag_text: str, name: str, value: str) -> str:
        pattern = rf'{name}="[^"]*"'
        replacement = f'{name}="{value}"'
        if re.search(pattern, tag_text):
            return re.sub(pattern, replacement, tag_text)
        return tag_text[:-2] + f' {replacement}/>'

    match = re.search(r"<calcPr\b[^>]*/>", text)
    if match is None:
        insert_at = text.rfind("</workbook>")
        if insert_at == -1:
            raise ValueError("workbook.xml is missing </workbook>")
        calc_pr = '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>'
        text = text[:insert_at] + calc_pr + text[insert_at:]
        return text.encode("utf-8")

    calc_tag = match.group(0)
    calc_tag = upsert_attr(calc_tag, "calcMode", "auto")
    calc_tag = upsert_attr(calc_tag, "fullCalcOnLoad", "1")
    calc_tag = upsert_attr(calc_tag, "forceFullCalc", "1")
    text = text[: match.start()] + calc_tag + text[match.end() :]
    return text.encode("utf-8")


def remove_calc_chain(entries: dict[str, bytes]) -> None:
    if "xl/calcChain.xml" in entries:
        del entries["xl/calcChain.xml"]

    rels_root = parse_xml(entries["xl/_rels/workbook.xml.rels"])
    changed = False
    for relationship in list(rels_root):
        if relationship.attrib.get("Type", "").endswith("/calcChain"):
            rels_root.remove(relationship)
            changed = True
    if changed:
        entries["xl/_rels/workbook.xml.rels"] = xml_bytes(rels_root)

    content_types_root = parse_xml(entries["[Content_Types].xml"])
    changed = False
    for override in list(content_types_root):
        if override.attrib.get("PartName") == "/xl/calcChain.xml":
            content_types_root.remove(override)
            changed = True
    if changed:
        entries["[Content_Types].xml"] = xml_bytes(content_types_root)


def build_backup_path(workbook_path: Path) -> Path:
    candidate = workbook_path.with_name(f"{workbook_path.stem}.backup{workbook_path.suffix}")
    counter = 2
    while candidate.exists():
        candidate = workbook_path.with_name(f"{workbook_path.stem}.backup{counter}{workbook_path.suffix}")
        counter += 1
    return candidate


def patch_workbook(workbook_path: Path, make_backup: bool) -> Path | None:
    entries: dict[str, bytes] = {}
    with zipfile.ZipFile(workbook_path, "r") as src:
        for info in src.infolist():
            entries[info.filename] = src.read(info.filename)

    styles_root = parse_xml(entries["xl/styles.xml"])
    percent_input_style = ensure_percent_input_style(styles_root)
    entries["xl/styles.xml"] = restore_required_namespace_declarations("xl/styles.xml", xml_bytes(styles_root))

    entries["xl/workbook.xml"] = mark_workbook_for_recalc(entries["xl/workbook.xml"])

    for sheet_path in TARGET_SHEETS:
        sheet_root = parse_xml(entries[sheet_path])
        reshape_sheet(sheet_root, percent_input_style, template_mode=sheet_path.endswith("sheet3.xml"))
        entries[sheet_path] = restore_required_namespace_declarations(sheet_path, xml_bytes(sheet_root))

    remove_calc_chain(entries)

    backup_path = None
    if make_backup:
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
    parser = argparse.ArgumentParser(description="Update the renewal workbook with investor override and signed-offer averages.")
    parser.add_argument("workbook", type=Path, help="Path to the .xlsx workbook")
    parser.add_argument("--backup", action="store_true", help="Save a backup copy before modifying the workbook")
    args = parser.parse_args()

    workbook_path = args.workbook.expanduser().resolve()
    if not workbook_path.exists():
        raise SystemExit(f"Workbook not found: {workbook_path}")

    backup_path = patch_workbook(workbook_path, args.backup)
    print(f"Updated workbook: {workbook_path}")
    if backup_path is not None:
        print(f"Backup saved: {backup_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
