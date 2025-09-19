# src/main.py
import os
import sys
import traceback
from pathlib import Path

from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
from werkzeug.utils import secure_filename

# ----- paths/imports -----
HERE = Path(__file__).resolve().parent        # repo/src
ROOT = HERE.parent                            # repo/
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))             # allow import from repo root

try:
    from improved_converter import SierraToWBSConverter
except ModuleNotFoundError:
    from .improved_converter import SierraToWBSConverter  # type: ignore

DATA_DIR = (ROOT / "data") if (ROOT / "data").exists() else (HERE / "data")
ORDER_TXT = DATA_DIR / "gold_master_order.txt"

# ----- app -----
app = Flask(__name__, static_folder=str(HERE / "static"))
CORS(app)

UPLOAD_FOLDER = Path("/tmp/uploads")
UPLOAD_FOLDER.mkdir(parents=True, exist_ok=True)
app.config["UPLOAD_FOLDER"] = str(UPLOAD_FOLDER)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024

ALLOWED = {"xlsx", "xls"}

def _ok(name: str) -> bool:
    return "." in name and name.rsplit(".", 1)[1].lower() in ALLOWED

converter = SierraToWBSConverter(str(ORDER_TXT) if ORDER_TXT.exists() else None)

# ----- routes -----
@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({
        "status": "ok",
        "version": "2.1.0",
        "gold_master_loaded": bool(getattr(converter, "gold_master_order", [])),
        "gold_master_count": len(getattr(converter, "gold_master_order", [])),
        "data_dir": str(DATA_DIR),
    })

@app.route("/api/employees", methods=["GET"])
def employees():
    roster = []
    for i, name in enumerate(getattr(converter, "gold_master_order", []), start=1):
        roster.append({
            "id": i, "name": name, "ssn": f"***-**-{i:04d}",
            "department": "UNKNOWN", "pay_rate": 0.0, "status": "A",
        })
    return jsonify(roster)

@app.route("/api/validate-sierra-file", methods=["POST"])
def validate():
    try:
        if "file" not in request.files:
            return jsonify({"error": "No file provided"}), 400
        f = request.files["file"]
        if not f.filename:
            return jsonify({"error": "No file selected"}), 400
        if not _ok(f.filename):
            return jsonify({"error": "File must be .xlsx or .xls"}), 400

        tmp = UPLOAD_FOLDER / f"tmp_{secure_filename(f.filename)}"
        f.save(str(tmp))
        try:
            df = converter.parse_sierra_file(str(tmp))
            total_hours = 0.0
            employees = 0
            if df is not None and not df.empty:
                # prefer REG/OT/DT; else Hours
                cols = [c for c in ["REGULAR", "OVERTIME", "DOUBLETIME"] if c in df.columns]
                if cols:
                    total_hours = float(df[cols].sum().sum())
                elif "Hours" in df.columns:
                    total_hours = float(df["Hours"].sum())
                if "Name" in df.columns:
                    employees = int(df["Name"].dropna().astype(str).str.strip().replace("", None).dropna().nunique())

            return jsonify({
                "valid": True,
                "employees": employees,
                "total_hours": round(total_hours, 2),
                "total_entries": int(len(df)) if df is not None else 0,
                "employee_names": (df["Name"].dropna().astype(str).tolist()[:10] if "Name" in df.columns else [])
            })
        finally:
            try: tmp.unlink(missing_ok=True)
            except Exception: pass
    except Exception as e:
        app.logger.error("validate error: %s", e)
        app.logger.error(traceback.format_exc())
        return jsonify({"valid": False, "error": str(e), "employees": 0, "total_hours": 0.0})

@app.route("/api/process-payroll", methods=["POST"])
def process():
    try:
        if "file" not in request.files:
            return jsonify({"error": "No file provided"}), 400
        f = request.files["file"]
        if not f.filename:
            return jsonify({"error": "No file selected"}), 400
        if not _ok(f.filename):
            return jsonify({"error": "File must be .xlsx or .xls"}), 400

        in_path  = UPLOAD_FOLDER / secure_filename(f.filename)
        out_path = UPLOAD_FOLDER / f"WBS_Payroll_{Path(f.filename).stem}.xlsx"
        f.save(str(in_path))
        try:
            res = converter.convert(str(in_path), str(out_path))
            if not res.get("success"):
                return jsonify({"error": res.get("error", "Conversion failed")}), 422
            return send_file(
                str(out_path),
                as_attachment=True,
                download_name=out_path.name,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        finally:
            try: in_path.unlink(missing_ok=True)
            except Exception: pass
    except Exception as e:
        app.logger.error("process error: %s", e)
        app.logger.error(traceback.format_exc())
        return jsonify({"error": f"Processing failed: {e}"}), 500

@app.route("/", defaults={"path": ""})
@app.route("/<path:path>")
def serve(path):
    static_path = Path(app.static_folder)
    if path and (static_path / path).exists():
        return send_from_directory(static_path, path)
    if (static_path / "index.html").exists():
        return send_from_directory(static_path, "index.html")
    return "index.html not found", 404

if __name__ == "__main__":
    print("Starting Sierra Payroll System...")
    print(f"Gold Master Order loaded: {len(getattr(converter, 'gold_master_order', []))} employees")
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=False)
