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

        # --- Always fetch LSRA template from Wix ---
        import requests
        from io import BytesIO

        TEMPLATE_URL = "https://fd9e47be-8bae-4028-9abb-e122237a79d5.usrfiles.com/ugd/fd9e47_c53aa7592925425dbb3e70ec9f45a74d.xlsx"

        resp = requests.get(TEMPLATE_URL)
        resp.raise_for_status()
        wb = openpyxl.load_workbook(BytesIO(resp.content))
        ws = wb.active

        # --- Write data into merged block A15:K19 ---
        ws["A15"] = (
            f"Date: {data.get('dateOfInspection', '')}\n"
            f"Location Address: {data.get('address', '')}\n"
            "Action(s) Taken: Creation of Corrective Action Plan, "
            "notified engineering of deficiencies\n"
            f"Person Completing Life Safety Risk Matrix: {data.get('inspector', '')}\n"
            "ILSM Required? YES"
        )

        # --- Build filename ---
        facility = data.get("facilityName", "UnknownFacility")
        floor = data.get("floorName", "UnknownFloor")
        safe_facility = facility.replace(" ", "_")
        safe_floor = floor.replace(" ", "_")
        file_name = f"LSRA - {safe_facility} - {safe_floor}.xlsx"

        # --- Save to memory ---
        from io import BytesIO
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        # --- Send as response ---
        return send_file(
            output,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=file_name
        )

    except Exception as e:
        print("❌ LSRA generation failed:", str(e))
        return jsonify({"error": str(e)}), 500

