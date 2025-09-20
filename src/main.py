# src/main.py
import os
import sys
import tempfile
import traceback
from pathlib import Path
from werkzeug.utils import secure_filename
from flask import Flask, send_from_directory, request, jsonify, send_file
from flask_cors import CORS

# -------- robust import/path setup --------
THIS_FILE = Path(__file__).resolve()
SRC_DIR = THIS_FILE.parent               # .../src
PROJECT_ROOT = SRC_DIR.parent            # repo root

# Allow importing from both src/ and repo root (handles improved_converter in either place)
for p in (SRC_DIR, PROJECT_ROOT):
    sp = str(p)
    if sp not in sys.path:
        sys.path.insert(0, sp)

# Try both import styles (no crash if file is in src/ or in root)
try:
    from improved_converter import SierraToWBSConverter
except ModuleNotFoundError:
    # last resort (namespaced)
    from src.improved_converter import SierraToWBSConverter  # type: ignore

# -------- app setup --------
app = Flask(__name__, static_folder=str(SRC_DIR / "static"))
app.config["SECRET_KEY"] = "sierra-payroll-secret-key-2024"
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB

CORS(app)

# uploads
UPLOAD_FOLDER = "/tmp/uploads"
ALLOWED_EXTENSIONS = {"xlsx", "xls"}
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

# data files live under repo/data
DATA_DIR = PROJECT_ROOT / "data"
GOLD_MASTER_PATH = str(DATA_DIR / "gold_master_order.txt")

# converter (exposes .gold_master_order used by /api/health)
converter = SierraToWBSConverter(GOLD_MASTER_PATH if Path(GOLD_MASTER_PATH).exists() else None)

def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

# -------- API --------
@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({
        "status": "ok",
        "version": "2.0.0",
        "converter": "improved_converter",
        "gold_master_loaded": bool(getattr(converter, "gold_master_order", [])),
        "gold_master_count": len(getattr(converter, "gold_master_order", [])),
    })

@app.route("/api/employees", methods=["GET"])
def get_employees():
    try:
        emps = []
        for i, name in enumerate(getattr(converter, "gold_master_order", [])):
            emps.append({
                "id": i + 1,
                "name": name,
                "ssn": f"***-**-{str(i).zfill(4)}",
                "department": "UNKNOWN",
                "pay_rate": 0.0,
                "status": "A",
            })
        return jsonify(emps)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ---- helpers for validation math (no crashes) ----
def _to_num_series(s):
    import pandas as pd
    return pd.to_numeric(s, errors="coerce").fillna(0.0)

def _sum_hours(df):
    import pandas as pd
    if df is None or getattr(df, "empty", True):
        return 0.0
    cols = [c for c in ["REGULAR", "OVERTIME", "DOUBLETIME", "VACATION", "SICK", "HOLIDAY"] if c in df.columns]
    if not cols:
        return 0.0
    safe = {c: _to_num_series(df[c]) for c in cols}
    return float(pd.DataFrame(safe).sum().sum())

@app.route("/api/validate-sierra-file", methods=["POST"])
def validate_sierra_file():
    try:
        if "file" not in request.files:
            return jsonify({"valid": False, "error": "No file provided", "employees": 0, "total_hours": 0.0})
        file = request.files["file"]
        if not file.filename:
            return jsonify({"valid": False, "error": "No file selected", "employees": 0, "total_hours": 0.0})
        if not allowed_file(file.filename):
            return jsonify({"valid": False, "error": "File must be Excel (.xlsx/.xls)", "employees": 0, "total_hours": 0.0})

        with tempfile.NamedTemporaryFile(dir=UPLOAD_FOLDER, delete=False, suffix=Path(file.filename).suffix) as tmp:
            file.save(tmp.name)
            temp_path = tmp.name

        try:
            df = converter.parse_sierra_file(temp_path)
            total_hours = _sum_hours(df)
            uniq = int(df["__canon"].nunique()) if "__canon" in df.columns else int(df["Name"].nunique())
            return jsonify({
                "valid": True,
                "employees": uniq,
                "total_hours": total_hours,
                "total_entries": int(len(df)),
                "employee_names": df["Name"].astype(str).head(10).tolist()
            })
        finally:
            try: os.remove(temp_path)
            except: pass

    except Exception as e:
        app.logger.error(f"Error validating file: {e}")
        app.logger.error(traceback.format_exc())
        return jsonify({"valid": False, "error": str(e), "employees": 0, "total_hours": 0.0})

@app.route("/api/process-payroll", methods=["POST"])
def process_payroll():
    try:
        if "file" not in request.files:
            return jsonify({"error": "No file provided"}), 400
        file = request.files["file"]
        if not file.filename:
            return jsonify({"error": "No file selected"}), 400
        if not allowed_file(file.filename):
            return jsonify({"error": "File must be Excel (.xlsx/.xls)"}), 400

        filename = secure_filename(file.filename)
        in_path = str(Path(UPLOAD_FOLDER) / filename)
        file.save(in_path)

        out_filename = f"WBS_Payroll_{Path(filename).stem}.xlsx"
        out_path = str(Path(UPLOAD_FOLDER) / out_filename)

        result = converter.convert(in_path, out_path)

        try: os.remove(in_path)
        except: pass

        if not result.get("success"):
            return jsonify({"error": f"Conversion failed: {result.get('error','unknown')}"}), 422

        return send_file(
            out_path,
            as_attachment=True,
            download_name=out_filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        app.logger.error(f"Error processing payroll: {e}")
        app.logger.error(traceback.format_exc())
        return jsonify({"error": f"Processing failed: {str(e)}"}), 500

# ---- serve frontend (unchanged) ----
@app.route("/", defaults={"path": ""})
@app.route("/<path:path>")
def serve(path):
    static_folder_path = app.static_folder
    if static_folder_path is None:
        return "Static folder not configured", 404
    if path != "" and os.path.exists(os.path.join(static_folder_path, path)):
        return send_from_directory(static_folder_path, path)
    index_path = os.path.join(static_folder_path, "index.html")
    if os.path.exists(index_path):
        return send_from_directory(static_folder_path, "index.html")
    return "index.html not found", 404

if __name__ == "__main__":
    print("Starting Sierra Payroll System...")
    print(f"Gold Master Order loaded: {len(getattr(converter, 'gold_master_order', []))} employees")
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port, debug=False)
