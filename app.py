from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
from io import BytesIO
import xlsxwriter

app = Flask(__name__)
CORS(app)

LOGO_PATH = os.path.join(os.path.dirname(__file__), "ASHE_logo.jpg")


@app.route("/")
def index():
    return jsonify({"status": "ok", "message": "LSRA server running"})


@app.route("/generate", methods=["POST"])
def generate_lsra():
    try:
        data = request.get_json(force=True)
        print("🔹 Incoming LSRA request:", data)

        output = BytesIO()
        wb = xlsxwriter.Workbook(output, {'in_memory': True})

        # ---------------- Tool Tab ----------------
        ws = wb.add_worksheet("Tool")

        # Header with logo (left) + title (center)
        if os.path.exists(LOGO_PATH):
            ws.set_header('&L&G&C&"Calibri,Bold"&14LIFE SAFETY RISK ASSESSMENT TOOL',
                          {'image_left': LOGO_PATH})
        else:
            ws.set_header('&C&"Calibri,Bold"&14LIFE SAFETY RISK ASSESSMENT TOOL')

        # Styles
        bold = wb.add_format({'bold': True, 'font_name': 'Calibri', 'font_size': 11})
        italic = wb.add_format({'italic': True, 'font_name': 'Calibri', 'font_size': 11})
        wrap_left = wb.add_format({'align': 'left', 'valign': 'top', 'text_wrap': True,
                                   'font_name': 'Calibri', 'font_size': 11, 'border': 1})
        wrap_center = wb.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
                                     'font_name': 'Calibri', 'font_size': 11, 'border': 1})
        red = wb.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
                             'border': 1, 'fg_color': '#FF0000', 'font_color': 'white'})
        yellow = wb.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
                                'border': 1, 'fg_color': '#FFFF00'})
        green = wb.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
                               'border': 1, 'fg_color': '#92D050'})
        orange = wb.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
                                'border': 1, 'fg_color': '#F79646'})

        # Column widths
        ws.set_column("A:A", 25)
        ws.set_column("B:E", 30)

        # Matrix header
        ws.merge_range("A3:A4", "Risk Tolerance", wrap_center)
        ws.merge_range("B2:E2", "Severity of Occurrence", wrap_center)

        headers = [
            "Category 1\nLife safety deficiency likely to cause major injury or death (1)",
            "Category 2\nLife safety deficiency likely to cause minor injury (2)",
            "Category 3\nLife safety deficiency not likely to cause injury (3)",
            "Category 4\nLife safety deficiency likely has minimal impact to safety (4)"
        ]
        for i, h in enumerate(headers):
            ws.write(3, i + 1, h, wrap_center)

        impacts = [
            ("Facility Wide\nImpacts the entire facility (1)", ["High", "High", "Medium", "Low"]),
            ("Multiple Units/Floors\nImpacts multiple smoke compartments (2)", ["High", "High", "Medium", "Low"]),
            ("Local/Single Unit\nImpacts a single smoke compartment or area (3)", ["High", "Medium", "Medium", "Low"]),
            ("Short Duration\nCorrection can be performed during shift identified (4)", ["", "No ILSM Required -", "", ""]),
        ]

        start_row = 4
        for i, (impact, values) in enumerate(impacts):
            ws.write(start_row + i, 0, impact, wrap_left)
            for j, val in enumerate(values):
                if val == "High":
                    ws.write(start_row + i, j + 1, "High\nNeeds Specific Remediation Actions", red)
                elif val == "Medium":
                    ws.write(start_row + i, j + 1, "Medium\nNeeds Remedial Action", yellow)
                elif val == "Low":
                    ws.write(start_row + i, j + 1, "Low\nRisk Acceptable Remedial Action Discretionary", orange)
                elif "No ILSM" in val:
                    ws.write(start_row + i, j + 1, val, green)
                else:
                    ws.write(start_row + i, j + 1, "", wrap_center)

        # Footer (row 21+)
        footer_rows = [
            ("Date:", data.get("dateOfInspection", "")),
            ("Location Address:", data.get("address", "")),
            ("Action(s) Taken:", "Creation of Corrective Action Plan, ILSM created, notified engineering."),
            ("Person Completing Life Safety Risk Matrix:", data.get("inspector", "")),
            ("ILSM Required?", "YES"),
        ]

        start_row = 21
        for i, (label, value) in enumerate(footer_rows, start=start_row):
            ws.write(i, 0, label, bold)
            ws.write(i, 1, value, italic)

        # Fit to 1 page
        ws.fit_to_pages(1, 1)

        # ---------------- Instructions Tab ----------------
        inst = wb.add_worksheet("Instructions")
        inst.set_column("A:A", 100)
        inst.write("A1", "Instructions for Life Safety Risk Assessment Tool", bold)
        inst.write("A3", "1. Use the Tool tab to complete the Life Safety Risk Assessment.")
        inst.write("A4", "2. Fill in the Date, Location, Actions Taken, Inspector, and ILSM status.")
        inst.write("A5", "3. The Risk Tolerance matrix helps determine severity and impact.")
        inst.write("A6", "4. Save the generated file as part of your compliance documentation.")
        inst.write("A7", "5. For questions, contact the Safety/Compliance department.")

        wb.close()
        output.seek(0)

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
