#!/usr/bin/env python3
"""Convert XML data to an Excel .xlsx file without external dependencies."""

from __future__ import annotations

import argparse
import sys
import xml.etree.ElementTree as ET
from collections import Counter
from pathlib import Path
from typing import Dict, Iterable, List, Tuple
from zipfile import ZIP_DEFLATED, ZipFile


def strip_ns(tag: str) -> str:
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def escape_xml(value: str) -> str:
    return (
        value.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def flatten_element(element: ET.Element, prefix: str = "") -> Dict[str, str]:
    row: Dict[str, str] = {}
    current = f"{prefix}.{strip_ns(element.tag)}" if prefix else strip_ns(element.tag)

    for attr, value in element.attrib.items():
        row[f"{current}.@{strip_ns(attr)}"] = value

    children = list(element)
    if not children:
        text = (element.text or "").strip()
        if text:
            row[current] = text
        return row

    for child in children:
        row.update(flatten_element(child, current))
    return row


def detect_rows(root: ET.Element) -> List[ET.Element]:
    children = list(root)
    if not children:
        return [root]

    direct_tags = [strip_ns(c.tag) for c in children]
    direct_counts = Counter(direct_tags)
    repeated_direct = {tag for tag, count in direct_counts.items() if count > 1}
    if repeated_direct:
        return [c for c in children if strip_ns(c.tag) in repeated_direct]

    best_rows: List[ET.Element] = []
    for child in children:
        grandchildren = list(child)
        if not grandchildren:
            continue
        gc_tags = [strip_ns(gc.tag) for gc in grandchildren]
        gc_counts = Counter(gc_tags)
        repeated_gc = {tag for tag, count in gc_counts.items() if count > 1}
        if repeated_gc:
            candidate = [gc for gc in grandchildren if strip_ns(gc.tag) in repeated_gc]
            if len(candidate) > len(best_rows):
                best_rows = candidate

    return best_rows or children


def to_table(rows: Iterable[ET.Element]) -> Tuple[List[str], List[Dict[str, str]]]:
    flattened = [flatten_element(row) for row in rows]
    columns = sorted({key for item in flattened for key in item})
    return columns, flattened


def col_name(index: int) -> str:
    name = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def build_sheet_xml(columns: List[str], rows: List[Dict[str, str]]) -> str:
    lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        "<sheetData>",
    ]

    all_rows: List[List[str]] = [columns] + [[row.get(c, "") for c in columns] for row in rows]
    for r_idx, values in enumerate(all_rows, start=1):
        lines.append(f'<row r="{r_idx}">')
        for c_idx, value in enumerate(values, start=1):
            ref = f"{col_name(c_idx)}{r_idx}"
            text = escape_xml(str(value))
            lines.append(f'<c r="{ref}" t="inlineStr"><is><t>{text}</t></is></c>')
        lines.append("</row>")

    lines.extend(["</sheetData>", "</worksheet>"])
    return "".join(lines)


def write_xlsx(columns: List[str], rows: List[Dict[str, str]], output: Path) -> None:
    sheet_xml = build_sheet_xml(columns, rows)

    content_types = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">
  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>
  <Default Extension=\"xml\" ContentType=\"application/xml\"/>
  <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>
  <Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>
</Types>"""

    rels = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>
</Relationships>"""

    workbook = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">
  <sheets>
    <sheet name=\"XML Data\" sheetId=\"1\" r:id=\"rId1\"/>
  </sheets>
</workbook>"""

    workbook_rels = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>
</Relationships>"""

    with ZipFile(output, "w", ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("xl/workbook.xml", workbook)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)


def convert_xml_to_excel(input_xml: Path, output_xlsx: Path) -> None:
    root = ET.parse(input_xml).getroot()
    row_elements = detect_rows(root)
    columns, flattened_rows = to_table(row_elements)
    if not columns:
        raise ValueError("Không tìm thấy dữ liệu để ghi ra Excel.")
    write_xlsx(columns, flattened_rows, output_xlsx)


def main() -> int:
    parser = argparse.ArgumentParser(description="Chuyển file XML sang Excel (.xlsx)")
    parser.add_argument("input", type=Path, help="File XML đầu vào")
    parser.add_argument("output", type=Path, nargs="?", help="File Excel đầu ra (.xlsx)")
    args = parser.parse_args()

    output = args.output or args.input.with_suffix(".xlsx")
    try:
        convert_xml_to_excel(args.input, output)
    except Exception as exc:
        print(f"Lỗi khi chuyển đổi: {exc}", file=sys.stderr)
        return 1

    print(f"Đã tạo file Excel: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
