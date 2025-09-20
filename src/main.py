# src/main.py — FINAL (stable endpoints; pairs with improved_converter.py)
import os
import sys
import traceback
from pathlib import Path
from flask import Flask, request, jsonify, send_from_directory, send_file
from flask_cors import CORS
from werkzeug.utils import secure_filename

# import converter from repo root
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from improved_converter import SierraToWBSConverter, DATA, ORDER_TXT

app = Flask(__name__, static_folder=str(Path(__file__).resolve().parent / "static"))
app.config['SECRET_KEY'] = 'sierra-payroll-secret-key-2024'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB
CORS(app)

UPLOAD_FOLDER = os.environ.get("UPLOAD_FOLDER", "/tmp/uploads")
Path(UPLOAD_FOLDER).mkdir(parents=True, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

ALLOWED = {"xlsx","xls"}
def ok_ext(name:str) -> bool:
    return "." in name and name.rsplit(".",1)[1].lower() in ALLOWED

# shared converter
converter = SierraToWBSConverter(str(ORDER_TXT if ORDER_TXT.exists() else ""))

@app.route('/api/health', methods=['GET'])
def health():
    try:
        return jsonify({
            "status": "ok",
            "converter": "improved_final",
            "gold_master_loaded": len(converter.gold_master_order) > 0,
            "gold_master_count": len(converter.gold_master_order),
            "data_dir": str(DATA)
        })
    except Exception as e:
        return jsonify({"status":"error","error":str(e)}), 500

@app.route('/api/employees', methods=['GET'])
def employees():
    try:
        return jsonify([{"id":i+1,"name":nm} for i, nm in enumerate(converter.gold_master_order)])
    except Exception as e:
        return jsonify({"error":str(e)}), 500

@app.route('/api/validate-sierra-file', methods=['POST'])
def validate():
    try:
        if 'file' not in request.files:
            return jsonify({"valid":False,"error":"No file provided"})
        f = request.files['file']
        if not f.filename:
            return jsonify({"valid":False,"error":"No file selected"})
        if not ok_ext(f.filename):
            return jsonify({"valid":False,"error":"File must be .xlsx or .xls"})

        tmp = Path(app.config['UPLOAD_FOLDER']) / secure_filename(f"val_{f.filename}")
        f.save(tmp)
        try:
            df = converter.parse_sierra_file(str(tmp))
            emps = int(df["__canon"].nunique()) if not df.empty else 0
            hours = float(df[["REGULAR","OVERTIME","DOUBLETIME"]].sum().sum()) if not df.empty else 0.0
            return jsonify({"valid": True, "employees": emps, "total_hours": round(hours, 3)})
        finally:
            try: tmp.unlink(missing_ok=True)
            except Exception: pass
    except Exception as e:
        app.logger.error("validate error: %s\n%s", str(e), traceback.format_exc())
        return jsonify({"valid":False,"error":str(e)})

@app.route('/api/process-payroll', methods=['POST'])
def process_payroll():
    try:
        if 'file' not in request.files:
            return jsonify({"error":"No file provided"}), 400
        f = request.files['file']
        if not f.filename:
            return jsonify({"error":"No file selected"}), 400
        if not ok_ext(f.filename):
            return jsonify({"error":"File must be .xlsx or .xls"}), 400

        in_path  = Path(app.config['UPLOAD_FOLDER']) / secure_filename(f.filename)
        f.save(in_path)

        out_name = f"WBS_Payroll_{Path(f.filename).stem}.xlsx"
        out_path = Path(app.config['UPLOAD_FOLDER']) / out_name

        try:
            result = converter.convert(str(in_path), str(out_path))
            if not result.get("success"):
                return jsonify({"error": f"File format error - {result.get('error','unknown')}"}), 422

            return send_file(str(out_path),
                             as_attachment=True,
                             download_name=out_name,
                             mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        finally:
            try: in_path.unlink(missing_ok=True)
            except Exception: pass
            # do not remove out_path (it’s being streamed)
    except Exception as e:
        app.logger.error("process error: %s\n%s", str(e), traceback.format_exc())
        return jsonify({"error":str(e)}), 500

@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def serve(path):
    static_dir = app.static_folder
    if path and Path(static_dir, path).exists():
        return send_from_directory(static_dir, path)
    idx = Path(static_dir) / "index.html"
    if idx.exists():
        return send_from_directory(static_dir, "index.html")
    return "index.html not found", 404

if __name__ == "__main__":
    print("Starting Sierra Payroll System...")
    print(f"Gold Master Order loaded: {len(converter.gold_master_order)} employees")
    port = int(os.environ.get("PORT", "8080"))
    app.run(host="0.0.0.0", port=port, debug=False)
