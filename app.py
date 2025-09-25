from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Alignment

app = Flask(__name__)
CORS(app)

# Paths (put both files in the repo root)
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "LSRA_TEMPLATE.xlsx")
LOGO_PATH     = os.path.join(os.path.dirname(__file__), "ASHE_logo.png")

@app.route("/")
def index():
    return jsonify({"status": "ok", "message": "LSRA server running"})

@app.route("/generate", methods=["POST"])
def generate_lsra():
    try:
        data = request.get_json(force=True)
        print("🔹 Incoming LSRA request:", data)

        print("📂 Looking for template at:", TEMPLATE_PATH)
        print("📂 Directory contents:", os.listdir(os.path.dirname(__file__)))

        if not os.path.exists(TEMPLATE_PATH):
            print("❌ Template not found at", TEMPLATE_PATH)
            return jsonify({"error": "Template not found"}), 500

        wb = openpyxl.load_workbook(TEMPLATE_PATH)
        print("✅ Template loaded successfully")

        ws = wb.active

        # ---- SAFETY: Unmerge anything touching rows 15–19 so we can write cleanly ----
        for row in range(15, 20):
            for rng in list(ws.merged_cells.ranges):
                if rng.min_row <= row <= rng.max_row:
                    print(f"🔎 Unmerging cells: {str(rng)}")
                    ws.unmerge_cells(str(rng))

        # ---- Header/logo + title (try real header first; fallback to sheet cells) ----
        def add_logo_and_title():
            # Try true print header image (may not be supported in some envs)
            try:
                from openpyxl.drawing.image import Image as XLImage
                if os.path.exists(LOGO_PATH):
                    img = XLImage(LOGO_PATH)
                    try:
                        # Try header image (left header) + &G placeholder
                        hf = ws.header_footer
                        hf.add_image(img, "L")        # left header image
                        hf.left_header = "&G"         # render that image
                        print("✅ ASHE logo placed in page header (left).")
                    except Exception as e:
                        print("⚠️ Header image failed, placing at A1 instead:", e)
                        ws.add_image(img, "A1")
                        # Add centered title in row 3 as fallback
                        try:
                            ws.merge_cells("A3:D3")
                        except Exception:
                            pass
                        ws["A3"] = "LIFE SAFETY RISK ASSESSMENT TOOL"
                        ws["A3"].font = Font(name="Calibri", size=14, bold=True)
                        ws["A3"].alignment = Alignment(horizontal="center", vertical="center")
                else:
                    print("⚠️ ASHE logo not found at", LOGO_PATH)
            except Exception as e:
                print("⚠️ Could not process ASHE logo:", e)

            # Try to set center header text too (safe even without image)
            try:
                ws.header_footer.center_header = "LIFE SAFETY RISK ASSESSMENT TOOL"
                print("✅ Title set in page header (center).")
            except Exception as e:
                print("⚠️ Header title failed; keeping in-sheet fallback:", e)
                try:
                    ws.merge_cells("A3:D3")
                except Exception:
                    pass
                ws["A3"] = "LIFE SAFETY RISK ASSESSMENT TOOL"
                ws["A3"].font = Font(name="Calibri", size=14, bold=True)
                ws["A3"].alignment = Alignment(horizontal="center", vertical="center")

        add_logo_and_title()

        # ---- Wrap & align the target cells ----
        for r in range(15, 20):
            a = ws[f"A{r}"]
            a.alignment = Alignment(wrap_text=True, vertical="top")
            a.font = Font(name="Calibri", size=11)

        # ---- Write labels + values into A15–A19 (single-cell per row; no rich runs) ----
        ws["A15"].value = f"Date: {data.get('dateOfInspection', '')}"
        ws["A16"].value = f"Location Address: {data.get('address', '')}"
        ws["A17"].value = "Action(s) Taken: Creation of Corrective Action Plan, notified engineering of deficiencies"
        ws["A18"].value = f"Person Completing Life Safety Risk Matrix: {data.get('inspector', '')}"
        ws["A19"].value = "ILSM Required? YES"

        # ---- Save to memory and return ----
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        facility = (data.get("facilityName", "Facility") or "Facility").replace(" ", "_")
        floor    = (data.get("floorName", "Floor") or "Floor").replace(" ", "_")
        filename = f"LSRA_{facility}_{floor}.xlsx"

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
    # Render typically sets PORT; local default 5000
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
