from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Alignment

app = Flask(__name__)
CORS(app)

# Find whichever template filename you committed
HERE = os.path.dirname(__file__)
TEMPLATE_CANDIDATES = ["LSRA_TEMPLATE.xlsx", "LSRA - TEMPLATE.xlsx"]
TEMPLATE_PATH = None
for name in TEMPLATE_CANDIDATES:
    p = os.path.join(HERE, name)
    if os.path.exists(p):
        TEMPLATE_PATH = p
        break

@app.route("/")
def index():
    return jsonify({"status": "ok", "message": "LSRA server running"})

@app.route("/generate", methods=["POST"])
def generate_lsra():
    try:
        data = request.get_json(force=True)
        print("🔹 Incoming LSRA request:", data)

        if not TEMPLATE_PATH or not os.path.exists(TEMPLATE_PATH):
            return jsonify({"error": "Template not found"}), 500

        wb = openpyxl.load_workbook(TEMPLATE_PATH)
        # Your template has two sheets; the main one is typically named "Tool" or "LSRA Tool"
        # Use the sheet named "Tool" if present; otherwise use the active sheet.
        ws = wb["Tool"] if "Tool" in wb.sheetnames else wb.active
        print("✅ Template loaded; sheets:", wb.sheetnames)

        # ---------- Fix merged block that collides with A15:B19 ----------
        # Your template merges A15:K20; writing to any cell in that range would crash unless we unmerge.
        for rng in list(ws.merged_cells.ranges):
            if rng.min_row <= 19 and rng.max_row >= 15:
                print(f"🔎 Unmerging {str(rng)} (overlaps rows 15–19)")
                ws.unmerge_cells(str(rng))

        # ---------- Column widths so footer values don’t overflow ----------
        # Adjust just A and B; leave the matrix columns as-is.
        try:
            ws.column_dimensions["A"].width = 18
            ws.column_dimensions["B"].width = 70
        except Exception as e:
            print("⚠️ Could not set column widths:", e)

        # ---------- Styles for footer ----------
        bold = Font(name="Calibri", size=11, bold=True)
        italic = Font(name="Calibri", size=11, italic=True)
        wrap_top = Alignment(wrap_text=True, vertical="top")

        # Footer rows (exactly like your template: labels in A, values in B, no borders)
        rows = [
            ("Date:", data.get("dateOfInspection", "")),
            ("Location Address:", data.get("address", "")),
            ("Action(s) Taken:", "Creation of Corrective Action Plan, ILSM created, notified engineering."),
            ("Person Completing Life Safety Risk Matrix:", data.get("inspector", "")),
            ("ILSM Required?", "YES"),
        ]

        start_row = 15
        for r, (label, value) in enumerate(rows, start=start_row):
            a = ws[f"A{r}"]; b = ws[f"B{r}"]
            a.value = label
            a.font = bold
            a.alignment = wrap_top

            b.value = value
            b.font = italic
            b.alignment = wrap_top

        # Optional: set print area to keep one-page look (tweak to your exact template height/width)
        try:
            ws.print_area = "A1:K30"
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 1
            ws.sheet_properties.pageSetUpPr.fitToPage = True
        except Exception as e:
            print("⚠️ Print area/page setup not applied:", e)

        # ---------- Save and return ----------
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        facility = (data.get("facilityName", "Facility") or "Facility").replace(" ", "_")
        floor = (data.get("floorName", "Floor") or "Floor").replace(" ", "_")
        filename = f"LSRA_{facility}_{floor}.xlsx"

        print("📤 Sending:", filename, "| bytes:", len(output.getvalue()))
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        print("❌ LSRA generation failed:", e)
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
