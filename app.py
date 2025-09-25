from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import requests
import openpyxl
from io import BytesIO

app = Flask(__name__)
CORS(app)  # ✅ Enable CORS for Wix

# URL to LSRA template hosted on Wix
TEMPLATE_URL = "https://fd9e47be-8bae-4028-9abb-e122237a79d5.usrfiles.com/ugd/fd9e47_c53aa7592925425dbb3e70ec9f45a74d.xlsx"


@app.route("/", methods=["GET"])
def index():
    return jsonify({"status": "ok", "message": "LSRA server running"})


@app.route("/generate", methods=["POST"])
def generate_lsra():
    try:
        data = request.get_json(force=True)
        print("🔹 Incoming LSRA request:", data)

        # Download fresh template from Wix
        resp = requests.get(TEMPLATE_URL)
        resp.raise_for_status()
        wb = openpyxl.load_workbook(BytesIO(resp.content))
        ws = wb.active

        # Fill row A15 with formatted info
        ws["A15"] = (
            f"Date: {data.get('dateOfInspection', '')}\n"
            f"Location Address: {data.get('address', '')}\n"
            "Action(s) Taken: Creation of Corrective Action Plan, "
            "notified engineering of deficiencies\n"
            f"Person Completing Life Safety Risk Matrix: {data.get('inspector', '')}\n"
            "ILSM Required? YES"
        )

        # Save workbook into memory
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        # Build filename
        safe_facility = data.get("facilityName", "Facility").replace(" ", "_")
        safe_floor = data.get("floorName", "Floor").replace(" ", "_")
        filename = f"LSRA_{safe_facility}_{safe_floor}.xlsx"

        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        print("❌ Error:", e)
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
