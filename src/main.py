import os
import sys
import traceback
from pathlib import Path

from flask import Flask, request, jsonify, send_from_directory, send_file
from flask_cors import CORS
from werkzeug.utils import secure_filename

# --- paths / imports ---------------------------------------------------------
REPO_ROOT = Path(__file__).resolve().parents[1]
APP_DIR   = Path(__file__).resolve().parent
DATA_DIR  = REPO_ROOT / "data"

sys.path.insert(0, str(REPO_ROOT))

from improved_converter import SierraToWBSConverter  # import from repo root

# --- flask app ---------------------------------------------------------------
app = Flask(__name__, static_folder=str(APP_DIR / "static"))
app.config["SECRET_KEY"] = "sierra-payroll-secret-key-2024"
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16MB
CORS(app)

# uploads live in ephemeral /tmp on Railway
UPLOAD_FOLDER = Path("/tmp/uploads")
UPLOAD_FOLDER.mkdir(parents=True, exist_ok=True)
app.config["UPLOAD_FOLDER"] = str(UPLOAD_FOLDER)

ALLOWED_EXTENSIONS = {"xlsx", "xls"}

# converter (loads gold order from data/gold_master_order.txt)
GOLD_MASTER_PATH = str(DATA_DIR / "gold_master_order.txt")
converter = SierraToWBSConverter(GOLD_MASTER_PATH if Path(GOLD_MASTER_PATH).exists() else None)

def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

# --- routes ------------------------------------------------------------------
@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({
        "status": "ok",
        "version": "2.2.0",
        "converter": "improved",
        "gold_master_loaded": bool(converter.gold_master_order),
        "gold_master_count": len(converter.gold_master_order)
    })

@app.route("/api/employees", methods=["GET"])
def get_employees():
    try:
        return jsonify([{"id": i+1, "name": nm} for i, nm in enumerate(converter.gold_master_order)])
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/validate-sierra-file", methods=["POST"])
def validate_sierra_file():
    try:
        if "file" not in request.files:
            return jsonify({"valid": False, "error": "No file provided"})
        f = request.files["file"]
        if f.filename == "":
            return jsonify({"valid": False, "error": "No file selected"})
        if not allowed_file(f.filename):
            return jsonify({"valid": False, "error": "Upload .xlsx or .xls"})

        tmp_name = secure_filename(f.filename)
        tmp_path = UPLOAD_FOLDER / f"validate_{tmp_name}"
        f.save(str(tmp_path))

        try:
            df = converter.parse_sierra_file(str(tmp_path))
            employees = int((df["Name"].str.strip() != "").sum()) if not df.empty else 0
            total_hours = float(df[["REGULAR", "OVERTIME", "DOUBLETIME"]].sum().sum()) if not df.empty else 0.0
            return jsonify({"valid": True, "employees": employees, "total_hours": round(total_hours, 3)})
        finally:
            try:
                tmp_path.unlink(missing_ok=True)
            except Exception:
                pass
    except Exception as e:
        app.logger.error("ERROR in validate: %s", traceback.format_exc())
        return jsonify({"valid": False, "error": str(e)}), 200  # keep 200 so UI shows banner

@app.route("/api/process-payroll", methods=["POST"])
def process_payroll():
    try:
        if "file" not in request.files:
            return jsonify({"error": "No file provided"}), 400
        f = request.files["file"]
        if f.filename == "":
            return jsonify({"error": "No file selected"}), 400
        if not allowed_file(f.filename):
            return jsonify({"error": "Upload .xlsx or .xls"}), 400

        in_name = secure_filename(f.filename)
        in_path = UPLOAD_FOLDER / in_name
        f.save(str(in_path))

        out_name = f"WBS_Payroll_{in_name}"
        out_path = UPLOAD_FOLDER / out_name

        result = converter.convert(str(in_path), str(out_path))

        try:
            in_path.unlink(missing_ok=True)
        except Exception:
            pass

        if not result.get("success"):
            return jsonify({"error": f"Conversion failed: {result.get('error','unknown')}"}), 422

        return send_file(
            str(out_path),
            as_attachment=True,
            download_name=out_name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        app.logger.error("ERROR in process: %s", traceback.format_exc())
        return jsonify({"error": str(e)}), 500

# static + SPA
@app.route("/", defaults={"path": ""})
@app.route("/<path:path>")
def serve(path):
    static_folder_path = app.static_folder
    if path and Path(static_folder_path, path).exists():
        return send_from_directory(static_folder_path, path)
    index_path = Path(static_folder_path, "index.html")
    return (send_from_directory(static_folder_path, "index.html")
            if index_path.exists() else ("index.html not found", 404))

if __name__ == "__main__":
    print("Starting Sierra Payroll System...")
    print(f"Gold Master Order loaded: {len(converter.gold_master_order)} employees")
    port = int(os.environ.get("PORT", "8080"))
    app.run(host="0.0.0.0", port=port, debug=False)
