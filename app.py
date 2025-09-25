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

        # Header with logo + title
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

        ws.fit_to_pages(1, 1)

        # ---------------- Instructions Tab ----------------
        inst = wb.add_worksheet("Instructions")
        inst.set_column("A:A", 100)
        inst.set_row(0, 40)

        # Logo
        if os.path.exists(LOGO_PATH):
            inst.insert_image("A1", LOGO_PATH, {"x_scale": 0.5, "y_scale": 0.5})

        # Title
        inst.write("C2", "Life Safety Risk Assessment Tool", wb.add_format({
            'bold': True, 'font_name': 'Calibri', 'font_size': 16, 'align': 'center'
        }))

        # Overview
        overview = [
            "Brief Overview of the Tool & How to Use It",
            "The Life Safety Risk Assessment Tool determines the risk tolerance of a life safety deficiency. "
            "The tool is used by determining the severity of occurrence and the impact of the deficiency.",
            "Severity of Occurrence is determined on the severity of the deficiency and the probability that the occurrence "
            "will have based on the following four categories:",
            "• Category 1: Life safety deficiency likely to cause major injury or death",
            "• Category 2: Life safety deficiency likely to cause minor injury",
            "• Category 3: Life safety deficiency not likely to cause injury",
            "• Category 4: Life safety deficiency likely has minimal impact to safety",
            "The Impact of Deficiency is determined on the extent of the deficiency based on the following four impacts:",
            "• Impact 1: Facility Wide – the deficiency impacts the entire facility",
            "• Impact 2: Multiple Units/Floors – the deficiency impacts more than one smoke compartment and/or multiple floors of the facility",
            "• Impact 3: Local/Single Unit – the deficiency only impacts a single smoke compartment or an area within a single smoke compartment",
            "• Impact 4: Short Duration – the correction of the deficiency can be performed during the shift that the deficiency was identified",
            "By plotting these two factors in the LSRA Tool an organization can determine the risk tolerance of the deficiency. "
            "The risk tolerance is dependent on the following tolerances:",
            "• High – The deficiency needs specific remediation actions as outlined in LS.01.02.01 EPs 2-15.",
            "• Medium – The deficiency needs remedial action as outlined in LS.01.02.01 EPs 2-15.",
            "• Low – The deficiency poses an acceptable risk and remedial actions are discretionary.",
            "• Short Duration – The deficiency can be corrected within the shift it has been identified in and no ILSM action is required."
        ]
        row = 4
        for line in overview:
            fmt = bold if line.startswith("Brief") else wb.add_format({'font_name': 'Calibri', 'font_size': 11, 'text_wrap': True})
            inst.write(f"A{row}", line, fmt)
            row += 1

        # Deficiency Examples table
        inst.write("U10", "Deficiency Example", bold)
        inst.write("Y10", "Impact", bold)
        inst.write("Z10", "Severity", bold)
        inst.write("AA10", "Tolerance", bold)

        examples = [
            ("Fire/Smoke Doors", "", "", ""),
            ("Door latching problem not immediately repairable", 3, 4, "Low"),
            ("Hardware not fire rated on doors throughout stairwell", 2, 4, "Low"),
            ("Excessive gap between door leafs", 4, 3, "No ILSM"),
            ("Fire Alarms", "", "", ""),
            ("Pull station mounted too high", 3, 3, "Medium"),
            ("Smoke detector missing above fire alarm panel", 3, 2, "Medium"),
            ("Damaged devices/appliances across smoke zones", 2, 2, "High"),
            ("Sprinkler System", "", "", ""),
            ("Missing escutcheon plate in multiple smoke zones", 2, 3, "Medium"),
            ("Storage within 18\" of sprinkler head deflector", 4, 4, "No ILSM"),
            ("Items attached or supported by sprinkler system throughout facility", 1, 2, "High"),
            ("Fire/Smoke Barriers", "", "", ""),
            ("Improperly protected vertical opening", 2, 1, "High"),
            ("Unprotected penetrations in fire or smoke barriers", 2, 2, "High"),
            ("Hazardous area not properly protected", 2, 2, "High"),
            ("Means of Egress", "", "", ""),
            ("Storage within exit enclosure", 2, 2, "High"),
            ("Blocking of an exit due to construction activities", 3, 1, "High"),
            ("Excessive travel distance to an approved exit", 3, 1, "High"),
        ]

        row = 11
        for ex in examples:
            inst.write(f"U{row}", ex[0], wb.add_format({'font_name': 'Calibri', 'font_size': 11, 'text_wrap': True}))
            if ex[1] != "":
                inst.write(f"Y{row}", ex[1], wrap_center)
                inst.write(f"Z{row}", ex[2], wrap_center)
                inst.write(f"AA{row}", ex[3], wrap_center)
            row += 1

        # Closing note
        inst.write("U30", "Of course these examples are based on the knowledge of those identifying the deficiency and it is highly "
                           "recommended that when possible the assessment be performed by a multidisciplinary group such as the "
                           "environment of care or safety committee, but due to the nature of these deficiencies sometimes this will "
                           "not be possible. Making sure that the tool, and its use, is part of the organization’s ILSM policy is also important.",
                   wb.add_format({'font_name': 'Calibri', 'font_size': 11, 'text_wrap': True}))

        # Footer disclaimer
        inst.write("A60", "The ASHE advocacy team works to monitor and fight the many overlapping codes and standards relating "
                          "the health care physical environment allowing health care facilities to optimize their physical environment "
                          "and focus more of their valuable resources on patient care.", wb.add_format({'font_name': 'Calibri', 'font_size': 8, 'italic': True, 'text_wrap': True}))
        inst.write("A63", "© The American Society for Healthcare Engineering (ASHE) of the American Hospital Association\n"
                          "155 North Wacker Drive, Suite 400 | Chicago, IL 60606 | Phone: 312-422-3800 | Email: ashe@aha.org | Web: www.ashe.org\n\n"
                          "Disclaimer: This document is provided by ASHE as a service to its members. The information provided may not apply "
                          "to a reader’s specific situation and is not a substitute for application of the reader’s own independent judgment "
                          "or the advice of a competent professional. ASHE does not make any guaranty or warranty as to the accuracy or completeness "
                          "of any information contained in this document. ASHE and the authors disclaim liability for personal injury, property damage, "
                          "or other damages of any kind, whether special, indirect, consequential, or compensatory, that may result from the use of or "
                          "reliance on this document.", wb.add_format({'font_name': 'Calibri', 'font_size': 8, 'text_wrap': True}))

        # Close workbook
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
