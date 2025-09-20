# src/main.py
import os
import sys
from pathlib import Path
from flask import Flask, jsonify, request
from werkzeug.utils import secure_filename

# ensure src directory is on path so improved_converter can be imported
ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

# import converter (file: src/improved_converter.py)
from improved_converter import SierraToWBSConverter

# Configure environment defaults if not present (prevents KeyError crashes)
UPLOAD_FOLDER = os.environ.get("UPLOAD_FOLDER", "/tmp/uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Data paths inside repo (adjust if different)
DATA_DIR = ROOT.parent / "data"
GOLD_ORDER = DATA_DIR / "gold_master_order.txt"
ROSTER_CSV = DATA_DIR / "gold_master_roster.csv"
TEMPLATE_XLSX = DATA_DIR / "wbs_template.xlsx"

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

converter = SierraToWBSConverter(gold_master_order_path=str(GOLD_ORDER))

@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})

@app.route("/api/validate-sierra-file", methods=["POST"])
def validate_sierra_file():
    # This endpoint expects file uploaded as form-data 'file'
    f = request.files.get("file")
    if not f:
        return jsonify({"success": False, "error": "no file uploaded"}), 400
    filename = secure_filename(f.filename)
    tmp_path = Path(app.config['UPLOAD_FOLDER']) / filename
    f.save(tmp_path)
    try:
        df = converter.parse_sierra_file(str(tmp_path))
        # count employees and total hours
        emp_count = int(df.shape[0])
        total_hours = float(df[["REGULAR","OVERTIME","DOUBLETIME"]].sum().sum()) if emp_count > 0 else 0.0
        return jsonify({"success": True, "employees": emp_count, "total_hours": total_hours})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/api/process-payroll", methods=["POST"])
def process_payroll():
    # same upload behavior
    f = request.files.get("file")
    if not f:
        return jsonify({"success": False, "error": "no file uploaded"}), 400
    filename = secure_filename(f.filename)
    tmp_path = Path(app.config['UPLOAD_FOLDER']) / filename
    f.save(tmp_path)

    out_path = Path(app.config['UPLOAD_FOLDER']) / f"wbs_output_{filename}"
    result = converter.convert(str(tmp_path), str(out_path))
    if not result.get("success"):
        return jsonify(result), 500
    # return link path (the front end expects success)
    return jsonify({"success": True, "out_file": str(out_path), **result})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
