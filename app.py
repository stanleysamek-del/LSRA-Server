from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
from io import BytesIO
import openpyxl

app = Flask(__name__)
CORS(app)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "LSRA_TEMPLATE.xlsx")

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
        ws = wb["Tool"]

        # --- DEBUG: list all merged ranges ---
        print("🔎 Listing merged ranges in 'Tool' sheet:")
        for rng in ws.merged_cells.ranges:
            print("   ", rng.coord)

        # Don’t try writing yet, just return the template back
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        facility = data.get("facilityName", "Facility").replace(" ", "_")
        floor = data.get("floorName", "Floor").replace(" ", "_")
        filename = f"DEBUG_MergedRanges_{facility}_{floor}.xlsx"

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
