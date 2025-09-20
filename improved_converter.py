import os
import sys
import io
import traceback
from pathlib import Path
from werkzeug.utils import secure_filename

# Make project root importable
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from flask import Flask, jsonify, request, send_file, send_from_directory
from flask_cors import CORS

from improved_converter import SierraToWBSConverter, DATA, ORDER_TXT  # local module

# --- Flask setup ------------------------------------------------------------
app = Flask(__name__, static_folder=str(Path(__file__).resolve().parent / "static"))
app.config['SECRET_KEY'] = 'sierra-payroll-secret-key-2024'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB

# Robust upload folder for Railway (always exists)
UPLOAD_FOLDER = os.environ.get("UPLOAD_FOLDER", "/tmp/uploads")
Path(UPLOAD_FOLDER).mkdir(parents=True, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

CORS(app)

ALLOWED_EXT = {"xlsx", "xls"}

def _ok_ext(name: str) -> bool:
    return "." in name and name.rsplit(".", 1)[1].lower() in ALLOWED_EXT

# One shared converter (loads gold order + roster + template from /data)
converter = SierraToWBSConverter(str(ORDER_TXT))

# --- Routes ----------------------------------------------------------------
@app.route("/api/health", methods=["GET"])
def health():
    """Simple health probe + what we loaded."""
    try:
        return jsonify({
            "status": "ok",
            "converter": "improved",
            "gold_master_loaded": len(converter.gold_master_order) > 0,
            "gold_master_count": len(converter.gold_master_order),
            "data_dir": str(DATA),
        })
    except Exception as e:
        return jsonify({"status": "error", "error": str(e)}), 500


@app.route("/api/employees", methods=["GET"])
def get_employees():
    """Expose the gold master order (names only) for the Employees tab."""
    try:
        return jsonify([{"id": i + 1, "name": nm} for i, nm in enumerate(converter.gold_master_order)])
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/validate-sierra-file", methods=["POST"])
def validate_sierra_file():
    """
    Lightweight parse to show “N employees, H hours” immediately after upload.
    Uses the same parser as final convert. No file is saved permanently.
    """
    try:
        if "file" not in request.files:
            return jsonify({"valid": False, "error": "No file provided"})

        f = request.files["file"]
        if not f.filename:
            return jsonify({"valid": False, "error": "No file selected"})

        if not _ok_ext(f.filename):
            return jsonify({"valid": False, "error": "File must be .xlsx or .xls"})

        # Read into memory and hand to pandas via a temp file
        tmp = Path(app.config['UPLOAD_FOLDER']) / secure_filename(f.filename)
        f.save(tmp)

        try:
            df = converter.parse_sierra_file(str(tmp))   # <- single source of truth
        finally:
            try: tmp.unlink(missing_ok=True)
            except Exception: pass

        # Counts
        unique_names = int(df["__canon"].nunique()) if not df.empty else 0
        total_hours = float(df[["REGULAR", "OVERTIME", "DOUBLETIME"]].sum().sum()) if not df.empty else 0.0

        return jsonify({
            "valid": True,
            "employees": unique_names,
            "total_hours": round(total_hours, 3),
            "total_entries": int(len(df)) if not df.empty else 0
        })
    except Exception as e:
        app.logger.error("ERROR in validate: %s\n%s", str(e), traceback.format_exc())
        return jsonify({"valid": False, "error": str(e)})


@app.route("/api/process-payroll", methods=["POST"])
def process_payroll():
    """Full convert -> returns the WBS Excel file (based on your /data/wbs_template.xlsx)."""
    try:
        if "file" not in request.files:
            return jsonify({"error": "No file provided"}), 400

        f = request.files["file"]
        if not f.filename:
            return jsonify({"error": "No file selected"}), 400

        if not _ok_ext(f.filename):
            return jsonify({"error": "File must be .xlsx or .xls"}), 400

        # Save uploaded
        in_path = Path(app.config['UPLOAD_FOLDER']) / secure_filename(f.filename)
        f.save(in_path)

        out_name = f"WBS_Payroll_{Path(f.filename).stem}.xlsx"
        out_path = Path(app.config['UPLOAD_FOLDER']) / out_name

        result = converter.convert(str(in_path), str(out_path))

        try: in_path.unlink(missing_ok=True)
        except Exception: pass

        if not result.get("success"):
            return jsonify({"error": f"File format error - {result.get('error','unknown')}"}), 422

        return send_file(str(out_path),
                         as_attachment=True,
                         download_name=out_name,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        app.logger.error("ERROR in process: %s\n%s", str(e), traceback.format_exc())
        return jsonify({"error": str(e)}), 500


# Serve static SPA (if you’re using the built-in)
@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def serve(path):
    static_dir = app.static_folder
    if path and Path(static_dir, path).exists():
        return send_from_directory(static_dir, path)
    index_html = Path(static_dir, "index.html")
    if index_html.exists():
        return send_from_directory(static_dir, "index.html")
    return "index.html not found", 404


if __name__ == "__main__":
    print("Starting Sierra Payroll System...")
    print(f"Gold Master Order loaded: {len(converter.gold_master_order)} employees")
    port = int(os.environ.get("PORT", "8080"))
    app.run(host="0.0.0.0", port=port, debug=False)
