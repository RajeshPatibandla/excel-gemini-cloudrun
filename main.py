print("STARTUP DEBUG: Flask app has started!")

import sys
print("Startup test print!", file=sys.stdout, flush=True)

from flask import Flask, request, jsonify
from flask_cors import CORS
import tempfile
import os
from oletools.olevba import VBA_Parser
import requests

app = Flask(__name__)
CORS(app)

GEMINI_API_KEY = 'AIzaSyDnROx1cvFKvVhFKcIckCFzNUpbYx9zyvc'  # <== Replace with your real Gemini API key!
GEMINI_MODEL_ENDPOINT = (
    'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-pro:generateContent'
)

@app.route("/")
def home():
    print("[DEBUG] / route called")
    return "Service is running!"

@app.route("/testlog")
def test_log():
    print("[DEBUG] /testlog route called!")
    return "Test log endpoint hit"

def extract_vba_macros(file_path):
    print("[DEBUG] Entered extract_vba_macros for file:", file_path)
    vba_macros = []
    parser = VBA_Parser(file_path)
    print("[DEBUG] Parser type:", type(parser))
    print("[DEBUG] Has VBA Macros?", parser.contains_vba_macros)
    if parser.contains_vba_macros:
        for (_, stream_path, vba_filename, vba_code) in parser.extract_macros():
            print("[DEBUG] Macro found:", vba_filename, stream_path)
            if vba_code:
                print("[DEBUG] VBA code (first 100 chars):", vba_code[:100], "...")
                vba_macros.append({
                    'name': vba_filename or stream_path or 'UnknownMacro',
                    'vba': vba_code
                })
    else:
        print("[DEBUG] No VBA macros found in file.")
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
    try:
        response = requests.post(
            GEMINI_MODEL_ENDPOINT + "?key=" + GEMINI_API_KEY,
            headers=headers,
            json=body
        )
        data = response.json()
        text = (
            data["candidates"][0]["content"]["parts"][0]["text"]
            if "candidates" in data and len(data["candidates"]) > 0 else ""
        )
    except Exception as e:
        print("[DEBUG] Error contacting Gemini:", str(e))
        text = f"Error: {str(e)}"
    return text

@app.route("/convert-excel", methods=["POST"])
def convert_excel():
    try:
        print("[DEBUG] ===== Received POST to /convert-excel =====")
        if "file" not in request.files:
            print("[DEBUG] No file field 'file' in request")
            return jsonify({"error": "No file field 'file' in request"}), 400

        excel_file = request.files["file"]
        filename = excel_file.filename
        print(f"[DEBUG] Received file: {filename}")

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsm") as tmp:
            tmp.write(excel_file.read())
            temp_path = tmp.name
            print("[DEBUG] Temp file written to:", temp_path)

        macros_out = []
        vba_macros = extract_vba_macros(temp_path)
        print("[DEBUG] Number of macros extracted:", len(vba_macros))

        for macro in vba_macros:
            result = convert_vba_with_gemini(macro['vba'])
            macros_out.append({
                "name": macro['name'],
                "vba": macro['vba'],
                "apps_script": result,
                "notes": "See Gemini output for unsupported parts/suggestions."
            })

        os.remove(temp_path)
        print("[DEBUG] Temp file deleted.")

        if not macros_out:
            macros_out.append({
                "name": "NoMacrosFound",
                "vba": "",
                "apps_script": "",
                "notes": "No VBA macros found in file."
            })
            print("[DEBUG] No macros found in file.")

        response_json = {
            "message": f"File '{filename}' received and processed.",
            "filename": filename,
            "macros": macros_out
        }
        print("[DEBUG] Returning JSON response.")
        return jsonify(response_json)

    except Exception as e:
        print("[DEBUG] Exception in convert_excel:", repr(e))
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080, debug=True)

