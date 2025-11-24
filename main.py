import sys
import os
import tempfile
import requests

from flask import Flask, request, jsonify
from flask_cors import CORS

# NOTE: You must ensure oletools is installed (pip install oletools>=0.60.2)
try:
    from oletools.olevba import VBA_Parser
except ImportError:
    print("FATAL ERROR: The 'oletools' package is required but not installed.", file=sys.stderr)
    sys.exit(1)


# --- CONFIGURATION ---
# Read API key from environment variable for security
GEMINI_API_KEY = os.environ.get('GEMINI_API_KEY')
GEMINI_MODEL_ENDPOINT = (
    'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-pro:generateContent'
)
# ---------------------


# --- FLASK APP SETUP ---
app = Flask(__name__)
CORS(app)

print("STARTUP DEBUG: Flask app is initializing...", file=sys.stdout, flush=True)

if not GEMINI_API_KEY:
    print("FATAL ERROR: GEMINI_API_KEY environment variable not found.", file=sys.stderr, flush=True)

# -----------------------


@app.route("/")
def home():
    """Simple health check endpoint."""
    print("[DEBUG] / route called (Health Check)", file=sys.stdout, flush=True)
    return "VBA to Apps Script Conversion Service is running!"


def extract_vba_macros(file_path):
    """Uses oletools to extract VBA code from an Excel file."""
    print(f"[DEBUG] Entered extract_vba_macros for file: {file_path}", file=sys.stdout, flush=True)
    vba_macros = []
    
    try:
        parser = VBA_Parser(file_path)
    except Exception as e:
        print(f"[ERROR] Failed to initialize VBA_Parser for {file_path}: {e}", file=sys.stderr, flush=True)
        return vba_macros # Return empty list on failure

    if parser.contains_vba_macros:
        print("[DEBUG] VBA Macros detected.", file=sys.stdout, flush=True)
        for (_, stream_path, vba_filename, vba_code) in parser.extract_macros():
            if vba_code:
                print(f"[DEBUG] Extracted macro: {vba_filename}", file=sys.stdout, flush=True)
                vba_macros.append({
                    'name': vba_filename or stream_path or 'UnknownMacro',
                    'vba': vba_code
                })
    else:
        print("[DEBUG] No VBA macros found in file.", file=sys.stdout, flush=True)
        
    parser.close()
    return vba_macros


def convert_vba_with_gemini(vba_code):
    """Calls the Gemini API to convert VBA code to Google Apps Script."""
    if not GEMINI_API_KEY:
        return "ERROR: Gemini API Key is not configured on the server."

    prompt_text = (
        "Convert this Excel VBA macro to Google Apps Script for Google Sheets. "
        "List any parts that cannot be directly converted and suggest alternatives in a 'NOTES' section at the end.\n\n"
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
        # Using a higher temperature for creative translation is usually fine here, but default is 0.9
        # "config": { "temperature": 0.5 } 
    }
    
    try:
        response = requests.post(
            GEMINI_MODEL_ENDPOINT + "?key=" + GEMINI_API_KEY,
            headers=headers,
            json=body
        )
        response.raise_for_status() # Raises an HTTPError if the status is 4xx or 5xx

        data = response.json()
        text = (
            data.get("candidates", [{}])[0]
            .get("content", {})
            .get("parts", [{}])[0]
            .get("text", "Conversion failed to return text.")
        )
        print("[DEBUG] Gemini API call successful.", file=sys.stdout, flush=True)
    except requests.exceptions.HTTPError as e:
        print(f"[ERROR] Gemini HTTP Error: {e.response.text}", file=sys.stderr, flush=True)
        text = f"API Request Failed (HTTP Error). Response: {e.response.text}"
    except Exception as e:
        print(f"[ERROR] Error contacting Gemini: {str(e)}", file=sys.stderr, flush=True)
        text = f"API Request Failed: {str(e)}"
        
    return text


@app.route("/convert-excel", methods=["POST"])
def convert_excel():
    """Handles file upload, extracts macros, and sends them for Gemini conversion."""
    temp_path = None
    try:
        # 🌟 CRITICAL DEBUG LOG 🌟
        print("[DEBUG] === STARTING /convert-excel FUNCTION ===", file=sys.stdout, flush=True)
        
        # 1. File Check
        if "file" not in request.files:
            print("[DEBUG] ERROR: No file field 'file' in request (Check Postman 'form-data' key: 'file')", file=sys.stderr, flush=True)
            return jsonify({"error": "No file field 'file' in request"}), 400

        excel_file = request.files["file"]
        filename = excel_file.filename
        print(f"[DEBUG] Received file: {filename}", file=sys.stdout, flush=True)

        # 2. Save Temporary File
        # Use .read() once to get the file content
        file_content = excel_file.read()

        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(filename)[1] or ".xlsm") as tmp:
            tmp.write(file_content)
            temp_path = tmp.name
        
        file_size_bytes = os.path.getsize(temp_path)
        print(f"[DEBUG] Temp file saved to {temp_path}. Size: {file_size_bytes} bytes", file=sys.stdout, flush=True)


        # 3. Extract Macros
        vba_macros = extract_vba_macros(temp_path)
        print(f"[DEBUG] Total number of macros extracted: {len(vba_macros)}", file=sys.stdout, flush=True)
        macros_out = []

        # 4. Convert Macros
        for macro in vba_macros:
            result = convert_vba_with_gemini(macro['vba'])
            macros_out.append({
                "name": macro['name'],
                "vba": macro['vba'],
                "apps_script": result,
                "notes": "See Gemini output for unsupported parts/suggestions."
            })

        # 5. Handle No Macros Found
        if not macros_out:
            macros_out.append({
                "name": "NoMacrosFound",
                "vba": "",
                "apps_script": "",
                "notes": "No VBA macros found in file. Ensure you uploaded an unencrypted .XLSM file."
            })
            
        response_json = {
            "message": f"File '{filename}' processed.",
            "filename": filename,
            "macros": macros_out
        }
        
        print("[DEBUG] Returning final JSON response.", file=sys.stdout, flush=True)
        return jsonify(response_json)

    except Exception as e:
        # Ensure error message is logged to stderr in production environments
        print(f"[FATAL EXCEPTION] in convert_excel: {repr(e)}", file=sys.stderr, flush=True)
        return jsonify({"error": str(e), "message": "Internal server error during processing."}), 500

    finally:
        # 6. Cleanup Temporary File
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
                print(f"[DEBUG] Temp file deleted: {temp_path}", file=sys.stdout, flush=True)
            except Exception as e:
                print(f"[WARNING] Could not delete temp file {temp_path}: {str(e)}", file=sys.stderr, flush=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080, debug=True)
