#!/usr/bin/env python3
"""Chuyển XML BCTC sang Excel gần giống file mẫu."""

from __future__ import annotations

import argparse
import re
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Tuple
from zipfile import ZIP_DEFLATED, ZipFile


def escape_xml(value: str) -> str:
    return value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;").replace("'", "&apos;")


def col_name(index: int) -> str:
    name = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def split_tag(tag: str) -> Tuple[str, str]:
    if "_" not in tag:
        return tag, "value"
    left, right = tag.rsplit("_", 1)
    return left, right


def to_metric_table(element: ET.Element, add_section: str | None = None) -> Tuple[List[str], List[List[str]]]:
    cols: List[str] = []
    rows_map: Dict[str, Dict[str, str]] = {}
    for c in list(element):
        metric, col = split_tag(c.tag)
        if col not in cols:
            cols.append(col)
        rows_map.setdefault(metric, {})[col] = (c.text or "").strip()

    headers = (["section"] if add_section is not None else []) + ["chi_tieu"] + cols
    rows = []
    for metric in sorted(rows_map):
        row = ([add_section] if add_section is not None else []) + [metric] + [rows_map[metric].get(c, "") for c in cols]
        rows.append(row)
    return headers, rows


def report_to_table(report: ET.Element) -> Tuple[List[str], List[List[str]]]:
    children = list(report)
    if children and all(len(list(c)) == 0 for c in children):
        return to_metric_table(report)

    # report có section lồng nhau: gom về 1 sheet, thêm cột section
    headers: List[str] = []
    all_rows: List[List[str]] = []
    for section in children:
        sec_headers, sec_rows = to_metric_table(section, add_section=section.tag)
        if not headers:
            headers = sec_headers
        all_rows.extend(sec_rows)
    return headers, all_rows


def extract_sheets(root: ET.Element) -> List[Tuple[str, List[str], List[List[str]]]]:
    sheets: List[Tuple[str, List[str], List[List[str]]]] = []

    tt_chung = root.find("TTChung")
    if tt_chung is not None:
        headers = [c.tag for c in tt_chung]
        row = [(c.text or "").strip() for c in tt_chung]
        sheets.append(("TTChung", headers, [row]))

    tt_bao_cao = root.find("TTBaoCao")
    if tt_bao_cao is not None:
        for report in tt_bao_cao:
            headers, rows = report_to_table(report)
            sheets.append((report.tag, headers, rows))

    return sheets


def sheet_name_safe(name: str, used: set[str]) -> str:
    n = re.sub(r"[\\/*?:\[\]]", "_", name)[:31] or "Sheet"
    base = n
    i = 1
    while n in used:
        s = f"_{i}"
        n = (base[: 31 - len(s)] + s) if len(base) + len(s) > 31 else base + s
        i += 1
    used.add(n)
    return n


def build_sheet_xml(headers: List[str], rows: List[List[str]]) -> str:
    lines = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">', "<sheetData>"]
    for r_idx, vals in enumerate([headers] + rows, start=1):
        lines.append(f'<row r="{r_idx}">')
        for c_idx, v in enumerate(vals, start=1):
            lines.append(f'<c r="{col_name(c_idx)}{r_idx}" t="inlineStr"><is><t>{escape_xml(str(v))}</t></is></c>')
        lines.append("</row>")
    lines += ["</sheetData>", "</worksheet>"]
    return "".join(lines)


def write_xlsx(sheets: List[Tuple[str, List[str], List[List[str]]]], output: Path) -> None:
    used = set()
    names = [sheet_name_safe(n, used) for n, _, _ in sheets]

    cts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">', '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>', '<Default Extension="xml" ContentType="application/xml"/>', '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>']
    for i in range(len(sheets)):
        cts.append(f'<Override PartName="/xl/worksheets/sheet{i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
    cts.append("</Types>")

    workbook = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">', '<sheets>']
    rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
    for i, name in enumerate(names, start=1):
        workbook.append(f'<sheet name="{escape_xml(name)}" sheetId="{i}" r:id="rId{i}"/>')
        rels.append(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>')
    workbook += ['</sheets>', '</workbook>']
    rels.append('</Relationships>')

    root_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'

    with ZipFile(output, "w", ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", "".join(cts))
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("xl/workbook.xml", "".join(workbook))
        zf.writestr("xl/_rels/workbook.xml.rels", "".join(rels))
        for i, (_, headers, rows) in enumerate(sheets, start=1):
            zf.writestr(f"xl/worksheets/sheet{i}.xml", build_sheet_xml(headers, rows))


def default_output_name(input_xml: Path) -> Path:
    # 2025-BCTC_24-830.27694.1136772-13052026-1778667433778.xml -> 2025_BCTC_24_830.27694.1136772.xlsx
    m = re.match(r"^(\d{4})-([^-]+)-([^-]+)-\d{8}-\d+\.xml$", input_xml.name)
    if not m:
        return input_xml.with_suffix('.xlsx')
    return input_xml.with_name(f"{m.group(1)}_{m.group(2)}_{m.group(3)}.xlsx")


def convert_xml_to_excel(input_xml: Path, output_xlsx: Path) -> None:
    root = ET.parse(input_xml).getroot()
    sheets = extract_sheets(root)
    if not sheets:
        raise ValueError("Không tìm thấy dữ liệu báo cáo")
    write_xlsx(sheets, output_xlsx)


def main() -> int:
    p = argparse.ArgumentParser(description="Chuyển XML BCTC sang Excel")
    p.add_argument("input", type=Path)
    p.add_argument("output", type=Path, nargs="?")
    args = p.parse_args()
    output = args.output or default_output_name(args.input)
    try:
        convert_xml_to_excel(args.input, output)
    except Exception as e:
        print(f"Lỗi khi chuyển đổi: {e}", file=sys.stderr)
        return 1
    print(f"Đã tạo file Excel: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
