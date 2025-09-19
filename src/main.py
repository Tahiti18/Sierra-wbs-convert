# -*- coding: utf-8 -*-
import os
import traceback
from pathlib import Path
from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
from werkzeug.utils import secure_filename

# local import
from improved_converter import SierraToWBSConverter, DATA, ORDER_TXT

# ---- app & config ----
app = Flask(__name__, static_folder=str(Path(__file__).resolve().parent / "static"))
CORS(app)

# uploads (fixes earlier 'UPLOAD_FOLDER' errors)
UPLOAD_FOLDER = Path("/tmp/uploads")
UPLOAD_FOLDER.mkdir(parents=True, exist_ok=True)
app.config["UPLOAD_FOLDER"] = str(UPLOAD_FOLDER)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16MB

ALLOWED_EXT = {"xlsx", "xls"}

# converter
converter = SierraToWBSConverter(str(ORDER_TXT if ORDER_TXT.exists() else ""))

def _ok_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXT

def _coerce_num(series_like):
    import pandas as pd
    return pd.to_numeric(series_like, errors="coerce").fillna(0.0)

# ---- routes ----
@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({
        "status": "ok",
        "version": "2.0.0",
        "converter": "improved_template_writer",
        "gold_master_loaded": bool(converter.gold_master_order),
        "gold_master_count": len(converter.gold_master_order)
    })

@app.route("/api/employees", methods=["GET"])
def employees():
    return jsonify([
        {"id": i+1, "name": n, "ssn": "***-**-****", "department": "", "status": "A"}
        for i, n in enumerate(converter.gold_master_order)
    ])

@app.route("/api/validate-sierra-file", methods=["POST"])
def validate_sierra_file():
    try:
        if "file" not in request.files:
            return jsonify({"valid": False, "error": "No file provided"}), 400
        f = request.files["file"]
        if f.filename == "":
            return jsonify({"valid": False, "error": "No file selected"}), 400
        if not _ok_file(f.filename):
            return jsonify({"valid": False, "error": "File must be .xlsx or .xls"}), 400

        tmp_path = UPLOAD_FOLDER / ("tmp_" + secure_filename(f.filename))
        f.save(tmp_path)

        try:
            df = converter.parse_sierra_file(str(tmp_path))
            if df.empty:
                return jsonify({"valid": False, "employees": 0, "total_hours": 0.0, "error": "No valid rows"})

            # Total hours (REG+OT+DT) robustly
            total_hours = float(df[["REGULAR", "OVERTIME", "DOUBLETIME"]].sum().sum())

            # distinct employees from Sierra present in file
            emp_count = int(df["Name"].astype(str).str.strip().replace({"": None}).dropna().nunique())

            return jsonify({
                "valid": True,
                "employees": emp_count,
                "total_hours": round(total_hours, 2),
                "total_entries": int(len(df)),
                "error": None
            })
        finally:
            try:
                tmp_path.unlink(missing_ok=True)
            except Exception:
                pass

    except Exception as e:
        app.logger.error("validate_sierra_file failed: %s", e)
        app.logger.error(traceback.format_exc())
        return jsonify({"valid": False, "error": str(e), "employees": 0, "total_hours": 0.0}), 500

@app.route("/api/process-payroll", methods=["POST"])
def process_payroll():
    try:
        if "file" not in request.files:
            return jsonify({"error": "No file provided"}), 400
        f = request.files["file"]
        if f.filename == "":
            return jsonify({"error": "No file selected"}), 400
        if not _ok_file(f.filename):
            return jsonify({"error": "File must be .xlsx or .xls"}), 400

        in_path  = UPLOAD_FOLDER / secure_filename(f.filename)
        out_name = f"WBS_Payroll_{os.path.splitext(os.path.basename(f.filename))[0]}.xlsx"
        out_path = UPLOAD_FOLDER / out_name

        f.save(in_path)
        try:
            result = converter.convert(str(in_path), str(out_path))
            if not result.get("success"):
                return jsonify({"error": result.get("error", "Conversion failed")}), 422

            return send_file(
                str(out_path),
                as_attachment=True,
                download_name=out_name,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        finally:
            try: in_path.unlink(missing_ok=True)
            except Exception: pass
            # don't unlink out_path; user will download it
    except Exception as e:
        app.logger.error("process_payroll failed: %s", e)
        app.logger.error(traceback.format_exc())
        return jsonify({"error": f"Processing failed: {e}"}), 500

# Static fall-through
@app.route("/", defaults={"path": ""})
@app.route("/<path:path>")
def serve(path):
    static_path = Path(app.static_folder)
    if path and (static_path / path).exists():
        return send_from_directory(static_path, path)
    index_path = static_path / "index.html"
    if index_path.exists():
        return send_from_directory(static_path, "index.html")
    return "index.html not found", 404

if __name__ == "__main__":
    print("Starting Sierra Payroll System...")
    print(f"Gold Master Order loaded: {len(converter.gold_master_order)} employees")
    port = int(os.environ.get("PORT", 8080))
    # Host 0.0.0.0 for Railway
    app.run(host="0.0.0.0", port=port, debug=False)
