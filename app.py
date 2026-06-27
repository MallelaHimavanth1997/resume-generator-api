@app.route("/generate", methods=["POST"])
def generate():
    if API_KEY:
        provided = request.headers.get("X-API-Key", "")
        if provided != API_KEY:
            return jsonify({"error": "unauthorized"}), 401

    data = request.get_json(silent=True)
    if not data:
        return jsonify({"error": "missing or invalid JSON body"}), 400

    # Normalize experience field names
    if "experience" in data:
        normalized = []
        for job in data["experience"]:
            normalized.append({
                "client": job.get("client") or job.get("company") or "Client",
                "dates": job.get("dates") or "",
                "role": job.get("role") or job.get("title") or "Data Engineer",
                "responsibilities": job.get("responsibilities") or [job.get("description") or ""],
                "environment": job.get("environment") or ""
            })
        data["experience"] = normalized

    # Normalize certifications
    if "certifications" in data:
        data["certifications"] = [
            {"name": c, "description": ""} if isinstance(c, str) else c
            for c in data["certifications"]
        ]

    # Normalize projects
    if "projects" in data:
        data["projects"] = [
            {"title": p.get("title") or p.get("name",""), 
             "bullets": p.get("bullets") or [p.get("description","")]}
            for p in data["projects"]
        ]

    company = data.get("company_for_filename") or data.get("name") or "Resume"
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
