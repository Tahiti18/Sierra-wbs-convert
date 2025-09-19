# src/main.py
import os
import sys
import traceback
from pathlib import Path

from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
from werkzeug.utils import secure_filename

# --------- Paths / imports (robust to your layout) ----------
HERE = Path(__file__).resolve().parent            # .../repo/src
ROOT = HERE.parent                                # .../repo
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))                 # allow import of root-level modules

# try to import whether improved_converter.py is in root or in src
try:
    from improved_converter import SierraToWBSConverter
except ModuleNotFoundError:
    # fallback if file was placed inside src/
    from .improved_converter import SierraToWBSConverter  # type: ignore

# DATA directory: prefer repo/data, otherwise src/data
DATA_DIR = (ROOT / "data") if (ROOT / "data").exists() else (HERE / "data")
ORDER_TXT = DATA_DIR / "gold_master_order.txt"

# --------- Flask app ----------
app = Flask(__name__, static_folder=str(HERE / "static"))
CORS(app)

# uploads (fixes previous UPLOAD_FOLDER KeyError / missing dir)
UPLOAD_FOLDER = Path("/tmp/uploads")
UPLOAD_FOLDER.mkdir(parents=True, exist_ok=True)
app.config["UPLOAD_FOLDER"] = str(UPLOAD_FOLDER)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16MB

ALLOWED_EXT = {"xlsx", "xls"}

def _allowed(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXT

# instantiate converter (ok if file missing; we report in /health)
converter = SierraToWBSConverter(str(ORDER_TXT) if ORDER_TXT.exists() else None)

# ---------------- API ----------------
@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({
        "status": "ok",
        "version": "2.0.0",
        "converter": "improved_converter",
        "gold_master_loaded": bool(getattr(converter, "gold_master_order", [])),
        "gold_master_count": len(getattr(converter, "gold_master_order", [])),
        "data_dir": str(DATA_DIR),
    })

@app.route("/api/employees", methods=["GET"])
def employees():
    # expose the gold order as a simple employee list (masked SSNs if you want later)
    try:
        roster = []
        for i, name in enumerate(getattr(converter, "gold_master_order", []), start=1):
            roster.append({
                "id": i,
                "name": name,
                "ssn": f"***-**-{i:04d}",
                "department": "UNKNOWN",
                "pay_rate": 0.0,
                "status": "A",
            })
        return jsonify(roster)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/validate-sierra-file", methods=["POST"])
def validate_sierra_file():
    try:
        if "file" not in request.files:
            return jsonify({"error": "No file provided"}), 400
        f = request.files["file"]
        if not f.filename:
            return jsonify({"error": "No file selected"}), 400
        if not _allowed(f.filename):
            return jsonify({"error": "File must be .xlsx or .xls"}), 400

        tmp_name = f"temp_{secure_filename(f.filename)}"
        tmp_path = UPLOAD_FOLDER / tmp_name
        f.save(str(tmp_path))

        try:
            df = converter.parse_sierra_file(str(tmp_path))
            # hours: prefer REG/OT/DT, else Hours column
            total_hours = 0.0
            if not df.empty:
                cols = [c for c in ["REGULAR", "OVERTIME", "DOUBLETIME"] if c in df.columns]
                if cols:
                    total_hours = float(df[cols].sum().sum())
                elif "Hours" in df.columns:
                    total_hours = float(df["Hours"].sum())
            employees = int(df["Name"].nunique()) if ("Name" in df.columns and not df.empty) else 0

            return jsonify({
                "valid": True,
                "employees": employees,
                "total_hours": round(total_hours, 2),
                "total_entries": int(len(df)) if df is not None else 0,
                "employee_names": (df["Name"].dropna().astype(str).tolist()[:10] if "Name" in df.columns else [])
            })
        finally:
            try:
                tmp_path.unlink(missing_ok=True)
            except Exception:
                pass
    except Exception as e:
        app.logger.error("ERROR in validate_sierra_file: %s", e)
        app.logger.error(traceback.format_exc())
        return jsonify({"valid": False, "error": str(e), "employees": 0, "total_hours": 0.0})

@app.route("/api/process-payroll", methods=["POST"])
def process_payroll():
    try:
        if "file" not in request.files:
            return jsonify({"error": "No file provided"}), 400
        f = request.files["file"]
        if not f.filename:
            return jsonify({"error": "No file selected"}), 400
        if not _allowed(f.filename):
            return jsonify({"error": "File must be .xlsx or .xls"}), 400

        in_name = secure_filename(f.filename)
        in_path = UPLOAD_FOLDER / in_name
        f.save(str(in_path))

        out_name = f"WBS_Payroll_{Path(in_name).stem}.xlsx"
        out_path = UPLOAD_FOLDER / out_name

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
            try:
                in_path.unlink(missing_ok=True)
            except Exception:
                pass
    except Exception as e:
        app.logger.error("ERROR in process_payroll: %s", e)
        app.logger.error(traceback.format_exc())
        return jsonify({"error": f"Processing failed: {e}"}), 500

# --------- Frontend files (optional) ----------
@app.route("/", defaults={"path": ""})
@app.route("/<path:path>")
def serve(path):
    static_folder_path = app.static_folder
    index_path = Path(static_folder_path) / "index.html"
    if path and (Path(static_folder_path) / path).exists():
        return send_from_directory(static_folder_path, path)
    if index_path.exists():
        return send_from_directory(static_folder_path, "index.html")
    return "index.html not found", 404

if __name__ == "__main__":
    print("Starting Sierra Payroll System...")
    print(f"Gold Master Order loaded: {len(getattr(converter, 'gold_master_order', []))} employees")
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=False)
