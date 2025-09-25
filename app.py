from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Alignment

app = Flask(__name__)
CORS(app)

# Look for either LSRA_TEMPLATE.xlsx or LSRA - TEMPLATE.xlsx
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

        # Main sheet: usually "Tool" or "LSRA Tool"
        ws = wb["Tool"] if "Tool" in wb.sheetnames else wb.active
        print("✅ Template loaded; sheets:", wb.sheetnames)

        # ---------------- Unmerge overlapping ranges ----------------
        for rng in list(ws.merged_cells.ranges):
            if rng.min_row <= 25 and rng.max_row >= 21:  # covers our footer rows
                print(f"🔎 Unmerging {str(rng)} (overlaps footer)")
                ws.unmerge_cells(str(rng))

        # ---------------- Column widths ----------------
        try:
            ws.column_dimensions["A"].width = 22
            ws.column_dimensions["B"].width = 80
        except Exception as e:
            print("⚠️ Could not set column widths:", e)

        # ---------------- Styles ----------------
        bold = Font(name="Calibri", size=11, bold=True)
        italic = Font(name="Calibri", size=11, italic=True)
        wrap_top = Alignment(wrap_text=True, vertical="top")

        # ---------------- Footer (start at row 21) ----------------
        rows = [
            ("Date:", data.get("dateOfInspection", "")),
            ("Location Address:", data.get("address", "")),
            ("Action(s) Taken:", "Creation of Corrective Action Plan, ILSM created, notified engineering."),
            ("Person Completing Life Safety Risk Matrix:", data.get("inspector", "")),
            ("ILSM Required?", "YES"),
        ]

        start_row = 21
        for r, (label, value) in enumerate(rows, start=start_row):
            a = ws[f"A{r}"]
            b = ws[f"B{r}"]

            a.value = label
            a.font = bold
            a.alignment = wrap_top

            b.value = value
            b.font = italic
            b.alignment = wrap_top

        # ---------------- Reuse logo from Instructions ----------------
        try:
            ws.oddHeader.left.text = "&[Picture]"
            ws.oddHeader.left.size = 12
            ws.oddHeader.left.font = "Calibri,Bold"
            print("✅ Header updated to include ASHE logo reference")
        except Exception as e:
            print("⚠️ Could not set header logo:", e)

        # ---------------- Save workbook ----------------
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
