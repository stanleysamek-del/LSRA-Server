from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage

app = Flask(__name__)
CORS(app)

# Paths
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "LSRA_TEMPLATE.xlsx")
LOGO_PATH     = os.path.join(os.path.dirname(__file__), "ASHE_logo.jpg")

@app.route("/")
def index():
    return jsonify({"status": "ok", "message": "LSRA server running"})

@app.route("/generate", methods=["POST"])
def generate_lsra():
    try:
        data = request.get_json(force=True)
        print("🔹 Incoming LSRA request:", data)

        if not os.path.exists(TEMPLATE_PATH):
            return jsonify({"error": "Template not found"}), 500

        wb = openpyxl.load_workbook(TEMPLATE_PATH)
        ws = wb.active

        # ---- Insert logo if needed ----
        if os.path.exists(LOGO_PATH):
            try:
                img = XLImage(LOGO_PATH)
                ws.add_image(img, "A1")
                print("✅ ASHE logo placed at A1")
            except Exception as e:
                print("⚠️ Could not insert logo:", e)

        # ---- Define styles ----
        bold = Font(name="Calibri", size=11, bold=True)
        italic = Font(name="Calibri", size=11, italic=True)
        border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )
        align = Alignment(wrap_text=True, vertical="top")

        # ---- Write rows 15–19 (A=label, B=value) ----
        rows = [
            ("Date:", data.get("dateOfInspection", "")),
            ("Location Address:", data.get("address", "")),
            ("Action(s) Taken:", "Creation of Corrective Action Plan, ILSM created, notified engineering."),
            ("Person Completing Life Safety Risk Matrix:", data.get("inspector", "")),
            ("ILSM Required?", "YES"),
        ]

        start_row = 15
        for i, (label, value) in enumerate(rows, start=start_row):
            ws[f"A{i}"].value = label
            ws[f"A{i}"].font = bold
            ws[f"A{i}"].alignment = align
            ws[f"A{i}"].border = border

            ws[f"B{i}"].value = value
            ws[f"B{i}"].font = italic
            ws[f"B{i}"].alignment = align
            ws[f"B{i}"].border = border

        # ---- Save to memory ----
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        # Naming
        facility = (data.get("facilityName", "Facility") or "Facility").replace(" ", "_")
        floor = (data.get("floorName", "Floor") or "Floor").replace(" ", "_")
        filename = f"LSRA_{facility}_{floor}.xlsx"

        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        print("❌ LSRA generation failed:", e)
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
