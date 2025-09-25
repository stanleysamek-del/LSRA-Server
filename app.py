from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
from io import BytesIO
import openpyxl
from openpyxl.styles import Font

app = Flask(__name__)
CORS(app)

# Local template path (keep LSRA_TEMPLATE.xlsx in your repo root)
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "LSRA_TEMPLATE.xlsx")

@app.route("/")
def index():
    return jsonify({"status": "ok", "message": "LSRA server running"})

@app.route("/generate", methods=["POST"])
def generate_lsra():
    try:
        data = request.get_json(force=True)
        print("🔹 Incoming LSRA request:", data)

        print("📂 Looking for template at:", TEMPLATE_PATH)
        print("📂 Directory contents:", os.listdir(os.path.dirname(__file__)))

        if not os.path.exists(TEMPLATE_PATH):
            print("❌ Template not found at", TEMPLATE_PATH)
            return jsonify({"error": "Template not found"}), 500

        wb = openpyxl.load_workbook(TEMPLATE_PATH)
        print("✅ Template loaded successfully")

        # Fill rows 15–19 with formatted text
        ws["A15"] = "Date:"
        ws["A15"].font = Font(bold=True, name="Calibri", size=11)
        ws["B15"] = data.get("dateOfInspection", "")
        ws["B15"].font = Font(italic=True, name="Calibri", size=11)

        ws["A16"] = "Location Address:"
        ws["A16"].font = Font(bold=True, name="Calibri", size=11)
        ws["B16"] = data.get("address", "")
        ws["B16"].font = Font(italic=True, name="Calibri", size=11)

        ws["A17"] = "Action(s) Taken:"
        ws["A17"].font = Font(bold=True, name="Calibri", size=11)
        ws["B17"] = "Creation of Corrective Action Plan, notified engineering of deficiencies"
        ws["B17"].font = Font(italic=True, name="Calibri", size=11)

        ws["A18"] = "Person Completing Life Safety Risk Matrix:"
        ws["A18"].font = Font(bold=True, name="Calibri", size=11)
        ws["B18"] = data.get("inspector", "")
        ws["B18"].font = Font(italic=True, name="Calibri", size=11)

        ws["A19"] = "ILSM Required? YES"
        ws["A19"].font = Font(bold=True, name="Calibri", size=11)

        # Save workbook into memory
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        # File naming convention
        facility = data.get("facilityName", "Facility").replace(" ", "_")
        floor = data.get("floorName", "Floor").replace(" ", "_")
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
    app.run(host="0.0.0.0", port=5000)
