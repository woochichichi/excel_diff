#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
실제 .xlsx 목업 생성기 (외부 라이브러리 0 — 표준 zipfile 만 사용).

목적:
 - CSV 목업(tests/data)과 동일한 데이터를 진짜 .xlsx 로 만들어,
   Excel COM 경로(ExcelComReader)와 실제 DRM 흐름까지 로컬에서 테스트.
 - Sheet2 에는 진짜 수식(=B2*0.1 등)을 넣어 '수식 diff' 도 검증 가능.

산출:
   tests/data/old.xlsx  (시트: Sheet1, Sheet2, OnlyOld)
   tests/data/new.xlsx  (시트: Sheet1, Sheet2, OnlyNew)

실행:  python3 tests/make_xlsx.py
"""
import os
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "data")

# 셀 표현:
#   str            → 텍스트(인라인 문자열)
#   int/float      → 숫자
#   ("=식", 캐시값) → 수식(+ 캐시된 계산값)
#   None           → 빈 셀
OLD = {
    "Sheet1": [
        ["ID", "Name", "Qty", "Price", "Note"],
        [1, "Apple", 10, 100, "fresh"],
        [2, "Banana", 5, 200, "ripe"],
        [3, "Cherry", 7, 300, None],
        [4, "Date", 2, 400, "dry"],
        [5, "Elderberry", 9, 500, "rare"],
    ],
    "Sheet2": [
        ["Item", "Base", "Tax", "Total"],
        ["A", 100, ("=B2*0.1", 10), ("=B2+C2", 110)],
        ["B", 200, ("=B3*0.1", 20), ("=B3+C3", 220)],
        ["C", 300, ("=B4*0.1", 30), ("=B4+C4", 330)],
    ],
    "OnlyOld": [
        ["P", "Q"],
        [9, 8],
    ],
}

NEW = {
    "Sheet1": [
        ["ID", "Name", "Qty", "Price", "Note"],
        [1, "Apple", 10, 150, "fresh"],
        [2, "Banana", 5, 200, "ripe"],
        [25, "Blueberry", 3, 250, "new"],
        [3, "Cherry", 9, 300, "tart"],
        [5, "Elderberry", 9, 500, None],
    ],
    "Sheet2": [
        ["Item", "Base", "Tax", "Total"],
        ["A", 100, ("=B2*0.15", 15), ("=B2+C2", 115)],
        ["B", 250, ("=B3*0.1", 25), ("=B3+C3", 275)],
        ["C", 300, ("=B4*0.1", 30), ("=B4+C4", 330)],
    ],
    "OnlyNew": [
        ["X", "Y"],
        [1, 2],
    ],
}

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def col_letter(col):  # 1-based → A, B, ... AA
    s = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        s = chr(ord("A") + rem) + s
    return s


def xml_escape(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
             .replace('"', "&quot;"))


def cell_xml(ref, value):
    if value is None:
        return ""
    if isinstance(value, tuple):  # 수식
        formula, cached = value
        f = formula[1:] if formula.startswith("=") else formula  # xlsx <f> 는 '=' 제외
        return '<c r="%s"><f>%s</f><v>%s</v></c>' % (ref, xml_escape(f), cached)
    if isinstance(value, bool):
        return '<c r="%s" t="b"><v>%d</v></c>' % (ref, 1 if value else 0)
    if isinstance(value, (int, float)):
        return '<c r="%s"><v>%s</v></c>' % (ref, value)
    # 텍스트 → 인라인 문자열
    return '<c r="%s" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (
        ref, xml_escape(str(value)))


def sheet_xml(rows):
    parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
             '<worksheet xmlns="%s"><sheetData>' % NS_MAIN]
    for r, row in enumerate(rows, start=1):
        cells = []
        for c, val in enumerate(row, start=1):
            cells.append(cell_xml(col_letter(c) + str(r), val))
        parts.append('<row r="%d">%s</row>' % (r, "".join(cells)))
    parts.append("</sheetData></worksheet>")
    return "".join(parts)


def content_types_xml(nsheets):
    overrides = ['<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>']
    for i in range(1, nsheets + 1):
        overrides.append('<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % i)
    overrides.append('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="%s">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '%s</Types>' % (NS_CT, "".join(overrides)))


def root_rels_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="%s">'
            '<Relationship Id="rId1" Type="%s/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>' % (NS_PKG_REL, NS_R))


def workbook_xml(sheet_names):
    sheets = []
    for i, name in enumerate(sheet_names, start=1):
        sheets.append('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (xml_escape(name), i, i))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="%s" xmlns:r="%s"><sheets>%s</sheets></workbook>'
            % (NS_MAIN, NS_R, "".join(sheets)))


def workbook_rels_xml(nsheets):
    rels = []
    for i in range(1, nsheets + 1):
        rels.append('<Relationship Id="rId%d" Type="%s/worksheet" Target="worksheets/sheet%d.xml"/>' % (i, NS_R, i))
    rels.append('<Relationship Id="rId%d" Type="%s/styles" Target="styles.xml"/>' % (nsheets + 1, NS_R))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="%s">%s</Relationships>' % (NS_PKG_REL, "".join(rels)))


def styles_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="%s">'
            '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
            '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
            '<borders count="1"><border/></borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
            '</styleSheet>' % NS_MAIN)


def build_xlsx(path, wb):
    names = list(wb.keys())
    n = len(names)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types_xml(n))
        z.writestr("_rels/.rels", root_rels_xml())
        z.writestr("xl/workbook.xml", workbook_xml(names))
        z.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml(n))
        z.writestr("xl/styles.xml", styles_xml())
        for i, name in enumerate(names, start=1):
            z.writestr("xl/worksheets/sheet%d.xml" % i, sheet_xml(wb[name]))


def main():
    if not os.path.isdir(OUT):
        os.makedirs(OUT)
    build_xlsx(os.path.join(OUT, "old.xlsx"), OLD)
    build_xlsx(os.path.join(OUT, "new.xlsx"), NEW)
    print("생성 완료:")
    print("  ", os.path.join(OUT, "old.xlsx"))
    print("  ", os.path.join(OUT, "new.xlsx"))


if __name__ == "__main__":
    main()
