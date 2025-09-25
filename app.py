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

        # Pick the active worksheet
        ws = wb.active

        # ---- Fill rows 15–19 with formatted text (safe for merged cells) ----
        ws["A15"] = f"Date: {data.get('dateOfInspection', '')}"
        ws["A15"].font = Font(name="Calibri", size=11)

        ws["A16"] = f"Location Address: {data.get('address', '')}"
        ws["A16"].font = Font(name="Calibri", size=11)

        ws["A17"] = "Action(s) Taken: Creation of Corrective Action Plan, notified engineering of deficiencies"
        ws["A17"].font = Font(name="Calibri", size=11)

        ws["A18"] = f"Person Completing Life Safety Risk Matrix: {data.get('inspector', '')}"
        ws["A18"].font = Font(name="Calibri", size=11)

        ws["A19"] = "ILSM Required? YES"
        ws["A19"].font = Font(name="Calibri", size=11)

        # Save workbook into memory
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        # File naming convention
        facility = data.get("facilityName", "Facility").replace(" ", "_")
        floor = data.get("floorName", "Floor").replace(" ", "_")
        filename = f"LSRA_{facility}_{floor}.xlsx"

        # 🔎 Debug: confirm final file before sending
        print("📤 Preparing to send file:", filename)
        print("📦 File size in memory:", len(output.getvalue()), "bytes")

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
