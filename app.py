from flask import Flask, request, send_file, jsonify
import io
import openpyxl
from datetime import datetime
import os

app = Flask(__name__)

# Path to LSRA template (make sure LSRA_TEMPLATE.xlsx is in the same repo or set full path)
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "LSRA_TEMPLATE.xlsx")

@app.route("/")
def index():
    return jsonify({"status": "ok", "message": "LSRA server running"})

@app.route("/generate", methods=["POST"])
def generate_lsra():
    try:
        data = request.get_json(force=True)

        # Extract values from request body
        date_val = data.get("dateOfInspection", "")
        address_val = data.get("address", "")
        inspector_val = data.get("inspector", "")
        facility_name = data.get("facilityName", "Unknown")
        floor_name = data.get("floorName", "")

        # Debug logging
        print("🔹 Incoming LSRA request:", data)

        # Load workbook template
        if not os.path.exists(TEMPLATE_PATH):
            return jsonify({"error": "Template not found"}), 500

        wb = openpyxl.load_workbook(TEMPLATE_PATH)
        ws = wb.active

        # Row 15:K19 — simplified as merged into A15
        # (Adjust if your template has specific merged ranges)
        ws["A15"] = (
            f"Date: {date_val}\n"
            f"Location Address: {address_val}\n"
            "Action(s) Taken: Creation of Corrective Action Plan, "
            "notified engineering of deficiencies\n"
            f"Person Completing Life Safety Risk Matrix: {inspector_val}\n"
            "ILSM Required? YES"
        )

        # Build filename
        safe_facility = facility_name.replace(" ", "_")
        safe_floor = floor_name.replace(" ", "_")
        file_name = f"LSRA - {safe_facility} - {safe_floor}.xlsx"

        # Save workbook into memory
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        # Return file as download
        return send_file(
            output,
            as_attachment=True,
            download_name=file_name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        print("❌ Error in /generate:", str(e))
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
