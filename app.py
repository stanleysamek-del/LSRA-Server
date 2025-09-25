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

        output = BytesIO()
        wb = xlsxwriter.Workbook(output, {"in_memory": True})

        # ------------- TOOL TAB (clone-like styling) -------------
        ws = wb.add_worksheet("Tool")

        # Header: ASHE logo (left) + title (center)
        if os.path.exists(LOGO_PATH):
            ws.set_header('&L&G&C&"Calibri,Bold"&14LIFE SAFETY RISK ASSESSMENT TOOL',
                          {"image_left": LOGO_PATH})
        else:
            ws.set_header('&C&"Calibri,Bold"&14LIFE SAFETY RISK ASSESSMENT TOOL')

        # ---- Formats (match template look) ----
        # General
        f10 = {"font_name": "Calibri", "font_size": 10}
        f11 = {"font_name": "Calibri", "font_size": 11}
        border = {"border": 1}
        wrap_c = {"align": "center", "valign": "vcenter", "text_wrap": True}
        wrap_l = {"align": "left",   "valign": "top",    "text_wrap": True}

        # Headers / labels
        hdr_merge = wb.add_format({**f11, **wrap_c, **border, "bold": True, "bg_color": "#E6E6E6"})
        rt_hdr     = wb.add_format({**f11, **wrap_c, **border, "bold": True, "bg_color": "#E6E6E6"})
        cat_hdr    = wb.add_format({**f10, **wrap_c, **border, "bold": True, "bg_color": "#E6E6E6"})
        left_lbl   = wb.add_format({**f10, **wrap_l, **border})

        # Risk cells
        red   = wb.add_format({**f10, **wrap_c, **border, "bg_color": "#FF0000", "font_color": "#FFFFFF"})
        yellow= wb.add_format({**f10, **wrap_c, **border, "bg_color": "#FFFF00", "font_color": "#000000"})
        orange= wb.add_format({**f10, **wrap_c, **border, "bg_color": "#F79646", "font_color": "#000000"})
        green = wb.add_format({**f10, **wrap_c, **border, "bg_color": "#92D050", "font_color": "#000000"})

        # Footer (labels/values)
        bold11  = wb.add_format({**f11, "bold": True})
        italic11= wb.add_format({**f11, "italic": True})

        # ---- Layout: columns/rows tuned to template proportions ----
        # Column A = left labels (wide). B..E = four risk columns.
        ws.set_column("A:A", 36)   # impact labels
        ws.set_column("B:E", 22)   # categories
        # Leave plenty of vertical space for readability
        for r in range(4, 9):
            ws.set_row(r, 36)      # main matrix rows height

        # ---- Matrix headers ----
        ws.merge_range("A3:A4", "Risk Tolerance", rt_hdr)           # left top header
        ws.merge_range("B2:E2", "Severity of Occurrence", hdr_merge)

        headers = [
            "Category 1\nLife safety deficiency likely to cause major\ninjury or death (1)",
            "Category 2\nLife safety deficiency likely to cause\nminor injury (2)",
            "Category 3\nLife safety deficiency not likely to\ncause injury (3)",
            "Category 4\nLife safety deficiency likely has minimal\nimpact to safety (4)",
        ]
        # Category headers row (row 4 in Excel terms is index 3): place in row 4 (index 3)
        for i, h in enumerate(headers):
            ws.write(3, i + 1, h, cat_hdr)

        # ---- Impact labels down the left ----
        impacts = [
            "Facility Wide\nImpacts the entire facility (1)",
            "Multiple Units/Floors\nImpacts multiple smoke compartments (2)",
            "Local/Single Unit\nImpacts a single smoke compartment or area (3)",
            "Short Duration\nCorrection can be performed during shift identified (4)",
        ]
        start_row = 4  # first data row index
        for i, lbl in enumerate(impacts):
            ws.write(start_row + i, 0, lbl, left_lbl)

        # ---- Risk grid ----
        # Rows correspond to impacts in the same order
        # For Short Duration row, the template shows a single green merged cell across all categories.
        # Row indices: start_row..start_row+3  (4 rows)
        # 1) Facility Wide
        ws.write(start_row + 0, 1, "High\nNeeds Specific Remediation Actions", red)
        ws.write(start_row + 0, 2, "High\nNeeds Specific Remediation Actions", red)
        ws.write(start_row + 0, 3, "Medium\nNeeds Remedial Action", yellow)
        ws.write(start_row + 0, 4, "Low\nRisk Acceptable\nRemedial Action\nDiscretionary", orange)
        # 2) Multiple Units/Floors
        ws.write(start_row + 1, 1, "High\nNeeds Specific Remediation Actions", red)
        ws.write(start_row + 1, 2, "High\nNeeds Specific Remediation Actions", red)
        ws.write(start_row + 1, 3, "Medium\nNeeds Remedial Action", yellow)
        ws.write(start_row + 1, 4, "Low\nRisk Acceptable\nRemedial Action\nDiscretionary", orange)
        # 3) Local/Single Unit
        ws.write(start_row + 2, 1, "High\nNeeds Specific Remediation Actions", red)
        ws.write(start_row + 2, 2, "Medium\nNeeds Remedial Action", yellow)
        ws.write(start_row + 2, 3, "Medium\nNeeds Remedial Action", yellow)
        ws.write(start_row + 2, 4, "Low\nRisk Acceptable\nRemedial Action\nDiscretionary", orange)
        # 4) Short Duration -> merge across B:E
        ws.merge_range(start_row + 3, 1, start_row + 3, 4, "No ILSM Required -", green)

        # ---- Footer block (row 21+) with extracted data ----
        footer_rows = [
            ("Date:", data.get("dateOfInspection", "")),
            ("Location Address:", data.get("address", "")),
            ("Action(s) Taken:", "Creation of Corrective Action Plan, ILSM created, notified engineering."),
            ("Person Completing Life Safety Risk Matrix:", data.get("inspector", "")),
            ("ILSM Required?", "YES"),
        ]
        start_footer = 21
        ws.set_row(start_footer,     18)
        ws.set_row(start_footer + 1, 18)
        ws.set_row(start_footer + 2, 18)
        ws.set_row(start_footer + 3, 18)
        ws.set_row(start_footer + 4, 18)
        ws.set_column("B:B", 80)  # wide values column so long addresses wrap nicely
        for r, (label, value) in enumerate(footer_rows, start=start_footer):
            ws.write(r, 0, label, bold11)
            ws.write(r, 1, value, italic11)

        # Page setup to mimic template print
        ws.set_margins(left=0.3, right=0.3, top=0.5, bottom=0.5)
        ws.center_horizontally()
        ws.fit_to_pages(1, 2)  # grid often spans two pages horizontally; adjust if you want 1x1

        # ------------- INSTRUCTIONS TAB (static) -------------
        inst = wb.add_worksheet("Instructions")
        inst.set_column("A:A", 100)
        inst.set_row(0, 40)

        # Logo in instructions (top-left)
        if os.path.exists(LOGO_PATH):
            inst.insert_image("A1", LOGO_PATH, {"x_scale": 0.5, "y_scale": 0.5})

        title_fmt = wb.add_format({"font_name": "Calibri", "font_size": 16, "bold": True, "align": "center"})
        inst.merge_range("C2:G2", "Life Safety Risk Assessment Tool", title_fmt)

        # Overview + bullets (same text as previous draft; update if you need any exact wording tweaked)
        p = wb.add_format({"font_name": "Calibri", "font_size": 11, "text_wrap": True})
        h = wb.add_format({"font_name": "Calibri", "font_size": 11, "bold": True})

        row = 4
        inst.write(row, 0, "Brief Overview of the Tool & How to Use It", h); row += 1
        inst.write(row, 0, "The Life Safety Risk Assessment Tool determines the risk tolerance of a life safety deficiency. "
                           "The tool is used by determining the severity of occurrence and the impact of the deficiency.", p); row += 2
        inst.write(row, 0, "Severity of Occurrence is determined on the severity of the deficiency and the probability that the "
                           "occurrence will have based on the following four categories:", p); row += 1
        for line in [
            "• Category 1: Life safety deficiency likely to cause major injury or death",
            "• Category 2: Life safety deficiency likely to cause minor injury",
            "• Category 3: Life safety deficiency not likely to cause injury",
            "• Category 4: Life safety deficiency likely has minimal impact to safety",
        ]:
            inst.write(row, 0, line, p); row += 1
        row += 1
        inst.write(row, 0, "The Impact of Deficiency is determined on the extent of the deficiency based on the following four impacts:", p); row += 1
        for line in [
            "• Impact 1: Facility Wide – the deficiency impacts the entire facility",
            "• Impact 2: Multiple Units/Floors – the deficiency impacts more than one smoke compartment and/or multiple floors of the facility",
            "• Impact 3: Local/Single Unit – the deficiency only impacts a single smoke compartment or an area within a single smoke compartment",
            "• Impact 4: Short Duration – the correction of the deficiency can be performed during the shift that the deficiency was identified",
        ]:
            inst.write(row, 0, line, p); row += 1
        row += 1
        inst.write(row, 0, "By plotting these two factors in the LSRA Tool an organization can determine the risk tolerance of the deficiency. "
                           "The risk tolerance is dependent on the following tolerances:", p); row += 1
        for line in [
            "• High – The deficiency needs specific remediation actions as outlined in LS.01.02.01 EPs 2-15.",
            "• Medium – The deficiency needs remedial action as outlined in LS.01.02.01 EPs 2-15.",
            "• Low – The deficiency poses an acceptable risk and remedial actions are discretionary.",
            "• Short Duration – The deficiency can be corrected within the shift it has been identified in and no ILSM action is required.",
        ]:
            inst.write(row, 0, line, p); row += 1

        # Right-side table (examples)
        inst.set_column("U:AA", 20)
        inst.write("U10", "Deficiency Example", h)
        inst.write("Y10", "Impact", h)
        inst.write("Z10", "Severity", h)
        inst.write("AA10", "Tolerance", h)

        tbl_c = wb.add_format({**f10, **wrap_c, "border": 1})
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
        r = 11
        for ex in examples:
            inst.write(f"U{r}", ex[0], wb.add_format({**f10, "text_wrap": True}))
            if ex[1] != "":
                inst.write(f"Y{r}", ex[1], tbl_c)
                inst.write(f"Z{r}", ex[2], tbl_c)
                inst.write(f"AA{r}", ex[3], tbl_c)
            r += 1

        # Closing paragraph
        inst.write("U30",
                   "Of course these examples are based on the knowledge of those identifying the deficiency and it is highly "
                   "recommended that when possible the assessment be performed by a multidisciplinary group such as the "
                   "environment of care or safety committee, but due to the nature of these deficiencies sometimes this will "
                   "not be possible. Making sure that the tool, and its use, is part of the organization’s ILSM policy is also important.",
                   p)

        # Footer disclaimer (small text)
        small = wb.add_format({"font_name": "Calibri", "font_size": 8, "text_wrap": True, "italic": True})
        inst.write("A60",
                   "The ASHE advocacy team works to monitor and fight the many overlapping codes and standards relating the health care physical "
                   "environment allowing health care facilities to optimize their physical environment and focus more of their valuable resources on patient care.",
                   small)
        inst.write("A63",
                   "© The American Society for Healthcare Engineering (ASHE) of the American Hospital Association\n"
                   "155 North Wacker Drive, Suite 400 | Chicago, IL 60606 | Phone: 312-422-3800 | Email: ashe@aha.org | Web: www.ashe.org\n\n"
                   "Disclaimer: This document is provided by ASHE as a service to its members. The information provided may not apply to a reader’s specific situation "
                   "and is not a substitute for application of the reader’s own independent judgment or the advice of a competent professional. ASHE does not make any "
                   "guaranty or warranty as to the accuracy or completeness of any information contained in this document. ASHE and the authors disclaim liability for "
                   "personal injury, property damage, or other damages of any kind, whether special, indirect, consequential, or compensatory, that may result from the "
                   "use of or reliance on this document.", small)

        # ---- finalize ----
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
