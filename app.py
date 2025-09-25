
import io
import zipfile
import requests
from fastapi import FastAPI, Response
from fastapi.middleware.cors import CORSMiddleware
import xml.etree.ElementTree as ET

# Namespaces for Excel XML
NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
ET.register_namespace("", NS["a"])  # default namespace

TEMPLATE_URL = "https://fd9e47be-8bae-4028-9abb-e122237a79d5.usrfiles.com/ugd/fd9e47_c53aa7592925425dbb3e70ec9f45a74d.xlsx"

app = FastAPI(title="LSRA Generator (Logo-Preserving)")

# CORS: in production you can restrict allow_origins to your Wix site
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

def _ensure_row(root, r_str: str):
    """Ensure a <row r='15'> exists and return it."""
    sheetData = root.find("a:sheetData", NS)
    if sheetData is None:
        sheetData = ET.SubElement(root, f"{{{NS['a']}}}sheetData")
    # Try find row
    for row in sheetData.findall("a:row", NS):
        if row.get("r") == r_str:
            return row
    # Not found: create in numeric order (append is fine for Excel)
    row = ET.Element(f"{{{NS['a']}}}row", {"r": r_str})
    sheetData.append(row)
    return row

def _find_cell(row_elem, cell_ref: str):
    """Find a cell <c r='A15'> inside the given row; return None if missing."""
    for c in row_elem.findall("a:c", NS):
        if c.get("r") == cell_ref:
            return c
    return None

def _make_run(text: str, bold: bool=False, italic: bool=False):
    """Create an <r> run with optional bold/italic, Calibri 11, and text."""
    r = ET.Element(f"{{{NS['a']}}}r")
    rPr = ET.SubElement(r, f"{{{NS['a']}}}rPr")
    # Font props
    if bold:
        ET.SubElement(rPr, f"{{{NS['a']}}}b")
    if italic:
        ET.SubElement(rPr, f"{{{NS['a']}}}i")
    ET.SubElement(rPr, f"{{{NS['a']}}}rFont", {"val": "Calibri"})
    ET.SubElement(rPr, f"{{{NS['a']}}}sz", {"val": "11"})
    ET.SubElement(rPr, f"{{{NS['a']}}}family", {"val": "2"})
    # Text (preserve spaces/newlines)
    t = ET.SubElement(r, f"{{{NS['a']}}}t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    return r

def build_inline_string(date_val: str, addr_val: str, inspector_val: str):
    """Return <is> element containing rich text runs for A15 with line breaks."""
    is_elem = ET.Element(f"{{{NS['a']}}}is")
    # Date:
    is_elem.append(_make_run("Date: ", bold=True))
    is_elem.append(_make_run((date_val or "") + "\n", italic=True))
    # Address:
    is_elem.append(_make_run("Location Address: ", bold=True))
    is_elem.append(_make_run((addr_val or "") + "\n", italic=True))
    # Actions:
    is_elem.append(_make_run("Action(s) Taken: ", bold=True))
    is_elem.append(_make_run("Creation of Corrective Action Plan, notified engineering of deficiencies\n", italic=True))
    # Person:
    is_elem.append(_make_run("Person Completing Life Safety Risk Matrix: ", bold=True))
    is_elem.append(_make_run((inspector_val or "") + "\n", italic=True))
    # ILSM:
    is_elem.append(_make_run("ILSM Required? YES", bold=True))
    return is_elem

def modify_xlsx_bytes(template_bytes: bytes, date_val: str, addr_val: str, inspector_val: str) -> bytes:
    """
    Modify only xl/worksheets/sheet1.xml cell A15 to an inline rich text string.
    All other files in the XLSX zip are copied verbatim to preserve logos/formatting.
    """
    in_mem = io.BytesIO(template_bytes)
    out_mem = io.BytesIO()
    with zipfile.ZipFile(in_mem, "r") as zin, zipfile.ZipFile(out_mem, "w", zipfile.ZIP_DEFLATED) as zout:
        # Read worksheet XML
        sheet_name = "xl/worksheets/sheet1.xml"
        xml = zin.read(sheet_name)
        root = ET.fromstring(xml)

        # Ensure row 15 and cell A15 exist
        row15 = _ensure_row(root, "15")
        cell = _find_cell(row15, "A15")
        if cell is None:
            cell = ET.SubElement(row15, f"{{{NS['a']}}}c", {"r": "A15"})

        # Set inline string type and content
        cell.set("t", "inlineStr")
        # Remove existing v/is children
        for child in list(cell):
            cell.remove(child)
        cell.append(build_inline_string(date_val, addr_val, inspector_val))

        # Write back modified worksheet
        new_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        # Copy all entries; replace sheet1.xml with modified bytes
        for item in zin.infolist():
            data = new_xml if item.filename == sheet_name else zin.read(item.filename)
            zout.writestr(item, data)
    return out_mem.getvalue()

@app.post("/generate-lsra")
def generate_lsra(payload: dict):
    """
    Payload JSON:
    {{
      "dateOfInspection": "...",
      "address": "...",
      "inspector": "..."
    }}
    """
    date_val = payload.get("dateOfInspection", "")
    addr_val = payload.get("address", "")
    inspector_val = payload.get("inspector", "")

    # Download template from Wix-hosted URL
    r = requests.get(TEMPLATE_URL, timeout=20)
    r.raise_for_status()

    # Modify minimally (only A15 as inline rich text)
    out_bytes = modify_xlsx_bytes(r.content, date_val, addr_val, inspector_val)

    return Response(
        content=out_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="LSRA_Report_Generated.xlsx"'}
    )

@app.get("/")
def root():
    return {"ok": True, "service": "LSRA Generator", "template_source": TEMPLATE_URL}
