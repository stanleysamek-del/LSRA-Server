from flask import Flask, request, send_file, jsonify
import io
import openpyxl

app = Flask(__name__)

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
    print("🔹 Incoming request:", data)

    # Minimal test: just echo data back
    return jsonify({
        "ok": True,
        "received": data
    })

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
