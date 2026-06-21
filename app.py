"""
app.py — tiny Flask API wrapping generate_resume.py so n8n (or anything else)
can POST structured resume JSON and get a .docx file back.

Deploy this on Render (free tier) — see DEPLOY.md for exact steps.

Endpoints:
  GET  /              -> health check
  POST /generate       -> body: resume_json (the same shape as sample_input.json)
                           returns: the .docx file as a binary download
"""

import os
import io
import tempfile
from flask import Flask, request, send_file, jsonify
from generate_resume import build_resume

app = Flask(__name__)

# Simple shared-secret auth so randoms on the internet can't hit your endpoint.
# Set this in Render's environment variables; n8n must send the same value
# in the X-API-Key header.
API_KEY = os.environ.get("API_KEY", "")


@app.route("/", methods=["GET"])
def health():
    return jsonify({"status": "ok", "service": "resume-generator"}), 200


@app.route("/generate", methods=["POST"])
def generate():
    if API_KEY:
        provided = request.headers.get("X-API-Key", "")
        if provided != API_KEY:
            return jsonify({"error": "unauthorized"}), 401

    data = request.get_json(silent=True)
    if not data:
        return jsonify({"error": "missing or invalid JSON body"}), 400

    company = data.get("company_for_filename") or "Resume"
    safe_company = "".join(c if c.isalnum() else "_" for c in company)
    filename = f"Resume_{safe_company}.docx"

    try:
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
            output_path = tmp.name
        build_resume(data, output_path)

        with open(output_path, "rb") as f:
            file_bytes = f.read()
        os.remove(output_path)

        return send_file(
            io.BytesIO(file_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name=filename,
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
