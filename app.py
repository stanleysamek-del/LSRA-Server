@app.route("/generate", methods=["POST"])
def generate_lsra():
    try:
        data = request.get_json(force=True)
        print("🔹 Incoming LSRA request:", data)

        # --- Fetch LSRA template from Wix (includes ASHE logo & formatting) ---
        import requests
        from io import BytesIO

        TEMPLATE_URL = "https://fd9e47be-8bae-4028-9abb-e122237a79d5.usrfiles.com/ugd/fd9e47_c53aa7592925425dbb3e70ec9f45a74d.xlsx"
        resp = requests.get(TEMPLATE_URL)
        resp.raise_for_status()

        wb = openpyxl.load_workbook(BytesIO(resp.content))
        ws = wb.active

        # --- Build formatted lines ---
        date_val = data.get("dateOfInspection", "")
        address_val = data.get("address", "")
        inspector_val = data.get("inspector", "")

        # Clear old value
        ws["A15"] = None

        from openpyxl.styles import Font
        from openpyxl.cell.rich_text import CellRichText, TextBlock

        # Create rich text for A15
        rich_text = CellRichText()

        rich_text.append(TextBlock(Font(bold=True), "Date: "))
        rich_text.append(TextBlock(Font(italic=True), date_val + "\n"))

        rich_text.append(TextBlock(Font(bold=True), "Location Address: "))
        rich_text.append(TextBlock(Font(italic=True), address_val + "\n"))

        rich_text.append(TextBlock(Font(bold=True), "Action(s) Taken: "))
        rich_text.append(TextBlock(Font(italic=True), "Creation of Corrective Action Plan, notified engineering of deficiencies\n"))

        rich_text.append(TextBlock(Font(bold=True), "Person Completing Life Safety Risk Matrix: "))
        rich_text.append(TextBlock(Font(italic=True), inspector_val + "\n"))

        rich_text.append(TextBlock(Font(bold=True), "ILSM Required? YES"))

        # Apply styled text to cell A15
        ws["A15"].value = rich_text

        # --- Build filename ---
        facility = data.get("facilityName", "Facility")
        floor = data.get("floorName", "Floor")
        safe_facility = facility.replace(" ", "_")
        safe_floor = floor.replace(" ", "_")
        filename = f"LSRA - {safe_facility} - {safe_floor}.xlsx"

        # --- Save to memory ---
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        return send_file(
            output,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        print("❌ LSRA generation failed:", str(e))
        return jsonify({"error": str(e)}), 500
