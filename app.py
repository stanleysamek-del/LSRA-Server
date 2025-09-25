from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import requests
from io import BytesIO
import openpyxl
from openpyxl.styles import Font

app = Flask(__name__)
CORS(app)

# Template URL (hosted XLSX)
TEMPLATE_URL = "https://fd9e47be-8bae-4028-9abb-e122237a79d5.usrfiles.com/ugd/fd9e47_c53aa7592925425dbb3e70ec9f45a74d.xlsx"

@app.route("/")
def index():
    return jsonify({"status": "ok", "message": "LSRA server running"})

@app.route("/generate", methods=["POST"])
def generate_lsra():
    try:
        data = request.get_json(force=True)
        print("🔹 Incoming LSRA request:", data)

        # Download template
        resp = requests.get(TEMPLATE_URL)
        resp.raise_for_status()
        wb = openpyxl.load_workbook(BytesIO(resp.content))
        ws = wb.active

        # Fill rows 15–19 with formatted text
        # Row 15 - Date
        ws["A15"] = "Date:"
        ws["A15"].font = Font(bold=True, name="Calibri", size=11)
        ws["B15"] = data.get("dateOfInspection", "")
        ws["B15"].font = Font(italic=True, name="Calibri", size=11)

        # Row 16 - Location
        ws["A16"] = "Location Address:"
        ws["A16"].font = Font(bold=True, name="Calibri", size=11)
        ws["B16"] = data.get("address", "")
        ws["B16"].font = Font(italic=True, name="Calibri", size=11)

        # Row 17 - Actions
        ws["A17"] = "Action(s) Taken:"
        ws["A17"].font = Font(bold=True, name="Calibri", size=11)
        ws["B17"] = "Creation of Corrective Action Plan, notified engineering of deficiencies"
        ws["B17"].font = Font(italic=True, name="Calibri", size=11)

        # Row 18 - Inspector
        ws["A18"] = "Person Completing Life Safety Risk Matrix:"
        ws["A18"].font = Font(bold=True, name="Calibri", size=11)
        ws["B18"] = data.get("inspector", "")
        ws["B18"].font = Font(italic=True, name="Calibri", size=11)

        # Row 19 - ILSM
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
