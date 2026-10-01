import os
import tempfile
from pathlib import Path
from flask import Flask, request, send_file, render_template_string
from openpyxl import Workbook

from ais_to_excel import (
    read_ais,
    AisParserError,
    sanitize_sheet_name,
    sheet_base_name,
    write_summary_sheet,
    write_category_sheet,
)

app = Flask(__name__)

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>AIS to Excel Converter</title>
    <style>
        :root {
            --bg-color: #0e1117;
            --card-bg: #1e2638;
            --input-bg: #262730;
            --accent: #4f8bf9;
            --btn-bg: #ff4b4b;
            --btn-hover: #ff3333;
            --text-main: #ffffff;
            --text-muted: #b0b8c4;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            background-color: var(--bg-color);
            color: var(--text-main);
            display: flex;
            justify-content: center;
            align-items: center;
            min-height: 100vh;
            margin: 0;
            padding: 1rem;
        }

        .container {
            width: 100%;
            max-width: 540px;
            background-color: var(--bg-color);
        }

        h1 {
            font-size: 1.75rem;
            margin-bottom: 0.5rem;
        }

        .privacy-banner {
            background-color: var(--card-bg);
            border-left: 4px solid var(--accent);
            padding: 0.85rem 1rem;
            border-radius: 6px;
            font-size: 0.88rem;
            color: var(--text-muted);
            margin-bottom: 1.5rem;
        }

        .form-group {
            margin-bottom: 1.25rem;
        }

        label {
            display: block;
            font-size: 0.88rem;
            font-weight: 600;
            color: var(--text-muted);
            margin-bottom: 0.4rem;
        }

        input[type="file"],
        input[type="text"],
        input[type="password"] {
            width: 100%;
            box-sizing: border-box;
            background-color: var(--input-bg);
            border: 1px solid #363945;
            color: var(--text-main);
            padding: 0.75rem;
            border-radius: 6px;
            font-size: 0.95rem;
            font-family: monospace;
        }

        input[type="file"] {
            font-family: inherit;
        }

        .row {
            display: flex;
            gap: 1rem;
        }

        .row .form-group {
            flex: 1;
        }

        button {
            width: 100%;
            background-color: var(--btn-bg);
            color: white;
            border: none;
            padding: 0.85rem;
            font-size: 1rem;
            font-weight: 600;
            border-radius: 6px;
            cursor: pointer;
            transition: background-color 0.2s ease;
            margin-top: 0.5rem;
        }

        button:hover {
            background-color: var(--btn-hover);
        }

        .error-msg {
            background-color: #3d1c1c;
            border: 1px solid #ff4b4b;
            color: #ff8080;
            padding: 0.75rem;
            border-radius: 6px;
            margin-bottom: 1rem;
            font-size: 0.9rem;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>📊 AIS to Excel Converter</h1>
        <div class="privacy-banner">
            🔒 <strong>100% Secure:</strong> File parsing occurs in memory. No data is stored.
        </div>

        {% if error %}
        <div class="error-msg">⚠️ {{ error }}</div>
        {% endif %}

        <form action="/" method="post" enctype="multipart/form-data">
            <div class="form-group">
                <label for="json_file">1. Upload AIS JSON File</label>
                <input type="file" id="json_file" name="json_file" accept=".json" required>
            </div>

            <div class="row">
                <div class="form-group">
                    <label for="pan">PAN</label>
                    <input type="text" id="pan" name="pan" placeholder="ABCDE1234F" maxlength="10" required>
                </div>
                <div class="form-group">
                    <label for="dob">Date of Birth (DDMMYYYY)</label>
                    <input type="password" id="dob" name="dob" placeholder="01011990" maxlength="8" required>
                </div>
            </div>

            <button type="submit">Convert & Download Excel</button>
        </form>
    </div>
</body>
</html>
"""

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "GET":
        return render_template_string(HTML_TEMPLATE)

    file = request.files.get("json_file")
    pan = request.form.get("pan", "").strip().upper()
    dob = request.form.get("dob", "").strip()

    if not file or not file.filename.endswith(".json"):
        return render_template_string(HTML_TEMPLATE, error="Please upload a valid AIS .json file.")
    if not pan or len(pan) != 10:
        return render_template_string(HTML_TEMPLATE, error="Please enter a valid 10-character PAN.")
    if not dob or len(dob) != 8 or not dob.isdigit():
        return render_template_string(HTML_TEMPLATE, error="Please enter Date of Birth as 8 digits (DDMMYYYY).")

    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".json") as tmp:
            file.save(tmp.name)
            tmp_path = Path(tmp.name)

        password = f"{pan}{dob}"
        statement = read_ais(path=tmp_path, password=password)
        frames = statement.to_df()

        used_names = {"Summary"}
        sheet_name_by_key = {
            key: sanitize_sheet_name(sheet_base_name(key, df), used_names) if len(df) > 0 else None
            for key, df in sorted(frames.items())
        }

        wb = Workbook()
        write_summary_sheet(wb, statement, frames, sheet_name_by_key)

        for key, df in sorted(frames.items()):
            if len(df) == 0:
                continue
            rows = df.to_dict(orient="records")
            write_category_sheet(wb, sheet_name_by_key[key], list(df.columns), rows)

        out_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        wb.save(out_file.name)

        output_name = file.filename.replace(".json", ".xlsx")
        return send_file(
            out_file.name,
            as_attachment=True,
            download_name=output_name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except AisParserError as exc:
        return render_template_string(HTML_TEMPLATE, error=f"Decryption failed: {exc}. Please check your PAN and Date of Birth.")
    except Exception as exc:
        return render_template_string(HTML_TEMPLATE, error=f"Error processing AIS file: {exc}")
    finally:
        if 'tmp_path' in locals() and tmp_path.exists():
            tmp_path.unlink()

# Vercel entrypoint export
handler = app

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
