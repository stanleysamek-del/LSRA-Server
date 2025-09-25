from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import io
import openpyxl
import os

app = Flask(__name__)
CORS(app)  # ✅ Allow all origins

@app.route("/")
def index():
    return jsonify({
        "ok": True,
        "service": "LSRA Generator",
        "template_source": "https://fd9e47be-8bae-4028-9abb-e122237a79d5.usrfiles.com/ugd/fd9e47_c53aa7592925425dbb3e70ec9f45a74d.xlsx"
    })

@app.route("/generate", methods=["POST"])
def generate_lsra():
    data = request.get_json(force=True)
    print("🔹 Incoming LSRA request:", data)

    # Minimal echo back (for testing)
    return jsonify({
        "ok": True,
        "received": data
    })
