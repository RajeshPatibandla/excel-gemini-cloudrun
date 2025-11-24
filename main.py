from flask import Flask, request, jsonify
from flask_cors import CORS
import tempfile
import os
from oletools.olevba import VBA_Parser
import requests

app = Flask(__name__)
CORS(app)

GEMINI_API_KEY = 'AIzaSyDnROx1cvFKvVhFKcIckCFzNUpbYx9zyvc'
GEMINI_MODEL_ENDPOINT = (
    'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-pro:generateContent'
)

@app.route("/")
def home():
    return "Service is running!"

def extract_vba_macros(file_path):
    vba_macros = []
    parser = VBA_Parser(file_path)
    if parser.contains_macros:
        for (_, stream_path, vba_filename, vba_code) in parser.extract_macros():
            if vba_code:
                vba_macros.append({
                    'name': vba_filename or stream_path or 'UnknownMacro',
                    'vba': vba_code
                })
    parser.close()
    return vba_macros

def convert_vba_with_gemini(vba_code):
    prompt_text = (
        "Convert this Excel VBA macro to Google Apps Script for Google Sheets. "
        "List any parts that cannot be directly converted and suggest alternatives.\n\n"
        "VBA code:\n" + vba_code
    )
    headers = {"Content-Type": "application/json"}
    body = {
        "contents": [
            {
                "parts": [
                    { "text": prompt_text }
                ]
            }
        ]
    }
    response = requests.post(
        GEMINI_MODEL_ENDPOINT + "?key=" + GEMINI_API_KEY,
        headers=headers,
        json=body
    )
    data = response.json()
    text = ''
    try:
        text = (
            data["candidates"][0]["content"]["parts"][0]["text"]
            if "candidates" in data and len(data["candidates"]) > 0 else ""
        )
    except Exception as e:
        text = f"Error parsing Gemini response: {str(e)}"
    return text

@app.route("/convert-excel", methods=["POST"])
def convert_excel():
    if "file" not in request.files:
        return jsonify({"error": "No file field 'file' in request"}), 400

    excel_file = request.files["file"]
    filename = excel_file.filename
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsm") as tmp:
        tmp.write(excel_file.read())
        temp_path = tmp.name

    macros_out = []
    vba_macros = extract_vba_macros(temp_path)

    for macro in vba_macros:
        result = convert_vba_with_gemini(macro['vba'])
        macros_out.append({
            "name": macro['name'],
            "vba": macro['vba'],
            "apps_script": result,
            "notes": "See Gemini output for unsupported parts/suggestions."
        })

    os.remove(temp_path)

    if not macros_out:
        macros_out.append({
            "name": "NoMacrosFound",
            "vba": "",
            "apps_script": "",
            "notes": "No VBA macros found in file."
        })

    return jsonify({
        "message": f"File '{filename}' received and processed.",
        "filename": filename,
        "macros": macros_out
    })

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080, debug=True)

