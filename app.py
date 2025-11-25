import os, io, json
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import google.generativeai as genai
from report_generator import generate_report

# ---------------- CONFIG ----------------
MASTER_EXCEL_PATH = "EmployeeDetails24-Nov-2025.xlsx"

genai.configure(api_key=os.getenv("GEMINI_API_KEY"))
model = genai.GenerativeModel("gemini-2.5-flash")

app = Flask(__name__)
CORS(app, resources={r"/api/*": {"origins": "*"}})  # Allow frontend access

# ---------------- ROUTE -----------------
@app.post("/api/upload-images")
def process_multi_images():
    # 1) Collect uploaded files
    images = request.files.getlist("image_files")
    if not images:
        return jsonify({"error": "Upload image_files[]"}), 400

    extracted_list = []

    # 2) Extract text using Gemini
    for img in images:
        try:
            img_bytes = img.read()
            prompt = """
            Extract all employee CLOCK NO and NAME from this image.
            RULES (VERY IMPORTANT):
            - RETURN ONLY JSON
            - NO COMMENT, NO CODE BLOCK, NO MARKDOWN, NO EXTRA TEXT
            - ALWAYS RETURN A LIST [] (even if empty)
            - JSON FORMAT EXACTLY LIKE:
            [
              {"clock": "A21646", "name": "Sanjay Kumar"}
            ]
            Now return ONLY the JSON list:
            """

            res = model.generate_content(
                [prompt, {"mime_type": img.mimetype, "data": img_bytes}]
            )

            # Clean and parse raw output
            txt = res.text.strip()
            txt = txt.replace("```json", "").replace("```", "").strip()
            data = json.loads(txt)

            print("GEMINI RAW OUTPUT:", data)

            if isinstance(data, list):
                extracted_list.extend(data)

        except Exception as e:
            print("Gemini error:", e)

    # 3) Merge duplicates
    clean_map = {}
    for item in extracted_list:
        c = str(item.get("clock", "")).upper().zfill(6)
        n = item.get("name", "").strip()
        if c:
            clean_map[c] = n  # last name wins if repeated

    final_clock_data = [{"clock": c, "name": n} for c, n in clean_map.items()]

    # 4) Generate PDF
    try:
        with open(MASTER_EXCEL_PATH, "rb") as file:
            pdf_bytes = generate_report(file, final_clock_data)
    except Exception as e:
        return jsonify({"error": "PDF generation failed", "details": str(e)}), 500

    # 5) Send PDF back
    return send_file(
        io.BytesIO(pdf_bytes),
        mimetype="application/pdf",
        as_attachment=True,
        download_name="Generated_Report.pdf"
    )

# ---------------- MAIN ----------------
if __name__ == "__main__":
    app.run(debug=True)
