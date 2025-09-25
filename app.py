from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
from io import BytesIO
import xlsxwriter

app = Flask(__name__)
CORS(app)

# Logo path (must be in repo root and committed to GitHub)
LOGO_PATH = os.path.join(os.path.dirname(__file__), "ASHE_logo.png")

@app.route("/")
def index():
    return jsonify({"status": "ok", "message": "LSRA server running"})

@app.route("/generate", methods=["POST"])
def generate_lsra():
    try:
        data = request.get_json(force=True)
        print("🔹 Incoming LSRA request:", data)

        # Prepare in-memory output
        output = BytesIO()

        # Create a new workbook in memory
        wb = xlsxwriter.Workbook(output, {'in_memory': True})
        ws = wb.add_worksheet("LSRA")

        # Define formats
        bold = wb.add_format({'bold': True, 'font_name': 'Calibri', 'font_size': 11})
        italic = wb.add_format({'italic': True, 'font_name': 'Calibri', 'font_size': 11})
        normal = wb.add_format({'font_name': 'Calibri', 'font_size': 11})
        wrap = wb.add_format({'text_wrap': True, 'font_name': 'Calibri', 'font_size': 11})

        # Adjust column width
        ws.set_column("A:A", 80, wrap)

        # Insert ASHE logo if available
        if os.path.exists(LOGO_PATH):
            ws.insert_image("A1", LOGO_PATH, {"x_scale": 0.5, "y_scale": 0.5})
            print("✅ ASHE logo inserted")
        else:
            print("⚠️ ASHE logo not found at", LOGO_PATH)

        # Title row
        ws.write("A3", "LIFE SAFETY RISK ASSESSMENT TOOL", bold)

        # Row 15: Date
        ws.write_rich_string(
            14, 0,  # row 15 (zero-indexed)
            bold, "Date: ",
            italic, data.get("dateOfInspection", ""),
            normal, ""
        )

        # Row 16: Location
        ws.write_rich_string(
            15, 0,
            bold, "Location Address: ",
            italic, data.get("address", ""),
            normal, ""
        )

        # Row 17: Action(s) Taken
        ws.write_rich_string(
            16, 0,
            bold, "Action(s) Taken: ",
            italic, "Creation of Corrective Action Plan, notified engineering of deficiencies",
            normal, ""
        )

        # Row 18: Inspector
        ws.write_rich_string(
            17, 0,
            bold, "Person Completing Life Safety Risk Matrix: ",
            italic, data.get("inspector", ""),
            normal, ""
        )

        # Row 19: ILSM
        ws.write_rich_string(
            18, 0,
            bold, "ILSM Required? ",
            italic, "YES",
            normal, ""
        )

        # Close workbook
        wb.close()
        output.seek(0)

        # File naming convention
        facility = data.get("facilityName", "Facility").replace(" ", "_")
        floor = data.get("floorName", "Floor").replace(" ", "_")
        filename = f"LSRA_{facility}_{floor}.xlsx"

        # Debug info
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
