#!/usr/bin/env python3
"""Chuyển XML sang Excel nhiều sheet theo từng báo cáo."""

from __future__ import annotations

import argparse
import re
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Tuple
from zipfile import ZIP_DEFLATED, ZipFile


def escape_xml(value: str) -> str:
    return (
        value.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def col_name(index: int) -> str:
    name = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def sheet_name_safe(name: str, used: set[str]) -> str:
    clean = re.sub(r"[\\/*?:\[\]]", "_", name)[:31] or "Sheet"
    base = clean
    i = 1
    while clean in used:
        suffix = f"_{i}"
        clean = (base[: 31 - len(suffix)] + suffix) if len(base) + len(suffix) > 31 else base + suffix
        i += 1
    used.add(clean)
    return clean


def split_tag(tag: str) -> Tuple[str, str]:
    if "_" not in tag:
        return tag, "value"
    metric, col = tag.rsplit("_", 1)
    return metric, col


def element_to_table(element: ET.Element) -> Tuple[List[str], List[List[str]]]:
    # Nếu tất cả con đều là lá: pivot theo tiền tố tag (CTxxx) thành từng hàng.
    children = list(element)
    if children and all(len(list(c)) == 0 for c in children):
        rows_map: Dict[str, Dict[str, str]] = {}
        cols: List[str] = []
        for c in children:
            metric, col = split_tag(c.tag)
            if col not in cols:
                cols.append(col)
            rows_map.setdefault(metric, {})[col] = (c.text or "").strip()

        headers = ["chi_tieu"] + cols
        rows = [[metric] + [rows_map[metric].get(c, "") for c in cols] for metric in sorted(rows_map)]
        return headers, rows

    # Trường hợp có cấu trúc lồng: flatten 1 cấp: mỗi node con là 1 hàng.
    headers_set = set(["muc"])
    data_rows: List[Dict[str, str]] = []
    for c in children:
        row = {"muc": c.tag}
        if len(list(c)) == 0:
            row["value"] = (c.text or "").strip()
            headers_set.add("value")
        else:
            for gc in c:
                row[gc.tag] = (gc.text or "").strip()
                headers_set.add(gc.tag)
        data_rows.append(row)

    headers = ["muc"] + sorted(h for h in headers_set if h != "muc")
    rows = [[r.get(h, "") for h in headers] for r in data_rows]
    return headers, rows


def build_sheet_xml(headers: List[str], rows: List[List[str]]) -> str:
    lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        "<sheetData>",
    ]
    all_rows = [headers] + rows
    for r_idx, values in enumerate(all_rows, start=1):
        lines.append(f'<row r="{r_idx}">')
        for c_idx, value in enumerate(values, start=1):
            ref = f"{col_name(c_idx)}{r_idx}"
            lines.append(f'<c r="{ref}" t="inlineStr"><is><t>{escape_xml(str(value))}</t></is></c>')
        lines.append("</row>")
    lines += ["</sheetData>", "</worksheet>"]
    return "".join(lines)


def extract_reports(root: ET.Element) -> List[Tuple[str, List[str], List[List[str]]]]:
    reports: List[Tuple[str, List[str], List[List[str]]]] = []

    tt_chung = root.find("TTChung")
    if tt_chung is not None:
        headers = [c.tag for c in tt_chung]
        row = [(c.text or "").strip() for c in tt_chung]
        reports.append(("TTChung", headers, [row]))

    tt_bao_cao = root.find("TTBaoCao")
    if tt_bao_cao is None:
        return reports

    for report in tt_bao_cao:
        if len(list(report)) == 0:
            reports.append((report.tag, ["value"], [[(report.text or "").strip()]]))
            continue

        # Nếu report chứa section lồng nhau (vd TMBCTC), tách mỗi section thành một sheet.
        if any(len(list(c)) > 0 for c in report):
            for section in report:
                headers, rows = element_to_table(section)
                reports.append((f"{report.tag}_{section.tag}", headers, rows))
        else:
            headers, rows = element_to_table(report)
            reports.append((report.tag, headers, rows))

    return reports


def write_xlsx_multi(reports: List[Tuple[str, List[str], List[List[str]]]], output: Path) -> None:
    used: set[str] = set()
    safe_names = [sheet_name_safe(name, used) for name, _, _ in reports]

    content_types = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
    ]
    for i in range(len(reports)):
        content_types.append(
            f'<Override PartName="/xl/worksheets/sheet{i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
    content_types.append("</Types>")

    root_rels = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>
</Relationships>"""

    wb_lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        "<sheets>",
    ]
    wb_rels = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    ]

    for i, name in enumerate(safe_names, start=1):
        wb_lines.append(f'<sheet name="{escape_xml(name)}" sheetId="{i}" r:id="rId{i}"/>')
        wb_rels.append(
            f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
        )

    wb_lines += ["</sheets>", "</workbook>"]
    wb_rels.append("</Relationships>")

    with ZipFile(output, "w", ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", "".join(content_types))
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("xl/workbook.xml", "".join(wb_lines))
        zf.writestr("xl/_rels/workbook.xml.rels", "".join(wb_rels))
        for i, (_, headers, rows) in enumerate(reports, start=1):
            zf.writestr(f"xl/worksheets/sheet{i}.xml", build_sheet_xml(headers, rows))


def convert_xml_to_excel(input_xml: Path, output_xlsx: Path) -> None:
    root = ET.parse(input_xml).getroot()
    reports = extract_reports(root)
    if not reports:
        raise ValueError("Không tìm thấy báo cáo trong XML")
    write_xlsx_multi(reports, output_xlsx)


def main() -> int:
    parser = argparse.ArgumentParser(description="Chuyển XML sang Excel theo từng báo cáo/sheet")
    parser.add_argument("input", type=Path)
    parser.add_argument("output", type=Path, nargs="?")
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
