"""Reproduce the small Excel navigation fixture using only Python's standard library."""
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED, ZipInfo

ROOT = Path(__file__).resolve().parents[1]
DESTINATION = ROOT / "tests" / "fixtures" / "Navigation.xlsx"
SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def xml(element):
    return ET.tostring(element, encoding="utf-8", xml_declaration=True)


def worksheet(cells):
    sheet = ET.Element("worksheet", xmlns=SHEET_NS)
    cols = ET.SubElement(sheet, "cols")
    ET.SubElement(cols, "col", min="1", max="4", width="30", customWidth="1")
    data = ET.SubElement(sheet, "sheetData")
    rows = {}
    for address, value, formula, kind in cells:
        row_number = int("".join(char for char in address if char.isdigit()))
        if row_number not in rows:
            rows[row_number] = ET.SubElement(data, "row", r=str(row_number))
        cell = ET.SubElement(rows[row_number], "c", r=address)
        if kind:
            cell.set("t", kind)
        if formula is not None:
            ET.SubElement(cell, "f").text = formula.lstrip("=")
        ET.SubElement(cell, "v").text = str(value)
    return xml(sheet)


def main():
    workbook = ET.Element("workbook", {"xmlns": SHEET_NS, "xmlns:r": REL_NS})
    sheets = ET.SubElement(workbook, "sheets")
    for index, name in enumerate(["Inputs", "Summary", "O'Brien"], 1):
        ET.SubElement(sheets, "sheet", name=name, sheetId=str(index), attrib={"r:id": "rId" + str(index)})
    names = ET.SubElement(workbook, "definedNames")
    ET.SubElement(names, "definedName", name="BaseCell").text = "Inputs!$A$1"
    ET.SubElement(workbook, "calcPr", calcId="191029", fullCalcOnLoad="1")
    relationships = ET.Element("Relationships", xmlns=PACKAGE_REL_NS)
    for index in range(1, 4):
        ET.SubElement(relationships, "Relationship", Id="rId" + str(index),
                      Type=REL_NS + "/worksheet", Target="worksheets/sheet" + str(index) + ".xml")
    package_relationships = ET.Element("Relationships", xmlns=PACKAGE_REL_NS)
    ET.SubElement(package_relationships, "Relationship", Id="rId1", Type=REL_NS + "/officeDocument", Target="xl/workbook.xml")
    types = ET.Element("Types", xmlns="http://schemas.openxmlformats.org/package/2006/content-types")
    ET.SubElement(types, "Default", Extension="rels", ContentType="application/vnd.openxmlformats-package.relationships+xml")
    ET.SubElement(types, "Default", Extension="xml", ContentType="application/xml")
    ET.SubElement(types, "Override", PartName="/xl/workbook.xml", ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
    for index in range(1, 4):
        ET.SubElement(types, "Override", PartName="/xl/worksheets/sheet" + str(index) + ".xml",
                      ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
    inputs = [
        ("A1", 10, None, None), ("B1", 30, "A1+A1*2", None),
        ("C1", "A1 is text", '"A1 is text"', "str"),
        ("A2", 20, None, None), ("B2", 20, "IF(A1>5,OFFSET(A1,1,0),A3)", None),
        ("C2", 3, "IF(A1>5,1,0)+IF(A1>5,2,0)", None),
        ("A3", 30, None, None), ("B3", 20, 'INDIRECT("A"&2)', None),
        ("C3", "#DIV/0!", "1/0", "e"),
        ("B4", 60, "SUM(A1:A3)", None),
        ("B5", 21, "'O''Brien'!A1", None),
        ("B6", 20, "BaseCell*2", None),
        ("B7", 30, 'IF(A1>5,B2+10,"none")', None),
        ("B8", 0, "IF(TRUE,0,A3)", None),
        ("B9", 9, "SUM(1,2)*3", None),
    ]
    parts = {
        "[Content_Types].xml": xml(types), "_rels/.rels": xml(package_relationships),
        "xl/workbook.xml": xml(workbook), "xl/_rels/workbook.xml.rels": xml(relationships),
        "xl/worksheets/sheet1.xml": worksheet(inputs),
        "xl/worksheets/sheet2.xml": worksheet([("A1", 40, "Inputs!B2*2", None), ("A2", 20, "Inputs!A2", None)]),
        "xl/worksheets/sheet3.xml": worksheet([("A1", 21, "Inputs!A2+1", None)]),
    }
    DESTINATION.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(DESTINATION, "w") as archive:
        for name, contents in sorted(parts.items()):
            ET.fromstring(contents)
            info = ZipInfo(name, date_time=(2026, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, contents)
    print(DESTINATION)


if __name__ == "__main__":
    main()
