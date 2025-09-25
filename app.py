from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import io
import openpyxl
import os

app = Flask(__name__)
CORS(app)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "LSRA_TEMPLATE.xlsx")

@app.route("/")
def index():
    return jsonify({"ok": True, "service": "LSRA Generator"})

@app.route("/generate", methods=["POST"])
def generate_lsra():
    try:
        data = request.get_json(force=True)
        print("🔹 Incoming LSRA request:", data)

        # Load template
        if not os.path.exists(TEMPLATE_PATH):
            return jsonify({"error": "Template not found"}), 500

        wb = openpyxl.load_workbook(TEMPLATE_PATH)
        ws = wb.active

        # Example: write into row 15 merged block (A15:K19)
        ws["A15"] = (
            f"Date: {data.get('dateOfInspection', '')}\n"
            f"Location Address: {data.get('address', '')}\n"
            "Action(s) Taken: Creation of Corrective Action Plan, "
            "notified engineering of deficiencies\n"
            f"Person Completing Life Safety Risk Matrix: {data.get('inspector', '')}\n"
            "ILSM Required? YES"
        )

        # Save into memory
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        # Build filename
        safe_facility = data.get("facilityName", "Facility").replace(" ", "_")
        safe_floor = data.get("floorName", "Floor").replace(" ", "_")
        filename = f"LSRA_{safe_facility}_{safe_floor}.xlsx"

        # ✅ Send as binary file
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        print("❌ Error:", e)
        return jsonify({"error": str(e)}), 500
