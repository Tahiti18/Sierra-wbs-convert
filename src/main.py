# src/main.py — stable Flask app that pairs with improved_converter.py
import os
import traceback
from pathlib import Path
from flask import Flask, request, jsonify, send_from_directory, send_file
from flask_cors import CORS
from werkzeug.utils import secure_filename

# import converter from repo root
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from improved_converter import SierraToWBSConverter, DATA, ORDER_TXT

app = Flask(__name__, static_folder=os.path.join(os.path.dirname(__file__), 'static'))
app.config['SECRET_KEY'] = 'sierra-payroll-secret-key-2024'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB

CORS(app)

UPLOAD_FOLDER = '/tmp/uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Converter
converter = SierraToWBSConverter(str(ORDER_TXT if ORDER_TXT.exists() else ""))

def allowed_file(filename: str) -> bool:
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/api/health', methods=['GET'])
def health():
    return jsonify({
        "status": "ok",
        "version": "2.0.0",
        "converter": "improved_flask",
        "gold_master_loaded": len(converter.gold_master_order) > 0,
        "gold_master_count": len(converter.gold_master_order)
    })

@app.route('/api/employees', methods=['GET'])
def get_employees():
    try:
        return jsonify([{"id": i+1, "name": n} for i, n in enumerate(converter.gold_master_order)])
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/validate-sierra-file', methods=['POST'])
def validate_sierra_file():
    try:
        if 'file' not in request.files:
            return jsonify({"valid": False, "error": "No file provided"})
        f = request.files['file']
        if f.filename == '':
            return jsonify({"valid": False, "error": "No file selected"})
        if not allowed_file(f.filename):
            return jsonify({"valid": False, "error": "File must be .xlsx or .xls"})

        tmp_name = secure_filename(f"val_{f.filename}")
        tmp_path = os.path.join(app.config['UPLOAD_FOLDER'], tmp_name)
        f.save(tmp_path)

        try:
            df = converter.parse_sierra_file(tmp_path)
            total_hours = float(df[["REGULAR","OVERTIME","DOUBLETIME"]].sum().sum()) if not df.empty else 0.0
            unique_emps = int(df["__canon"].nunique()) if not df.empty else 0
            return jsonify({"valid": True, "employees": unique_emps, "total_hours": total_hours})
        finally:
            try: os.remove(tmp_path)
            except: pass
    except Exception as e:
        app.logger.error("validate_sierra_file failed: %s\n%s", str(e), traceback.format_exc())
        return jsonify({"valid": False, "error": str(e)})

@app.route('/api/process-payroll', methods=['POST'])
def process_payroll():
    try:
        if 'file' not in request.files:
            return jsonify({"error": "No file provided"}), 400
        f = request.files['file']
        if f.filename == '':
            return jsonify({"error": "No file selected"}), 400
        if not allowed_file(f.filename):
            return jsonify({"error": "File must be .xlsx or .xls"}), 400

        in_name = secure_filename(f.filename)
        in_path = os.path.join(app.config['UPLOAD_FOLDER'], in_name)
        f.save(in_path)

        out_name = f"WBS_Payroll_{in_name}"
        out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_name)

        try:
            result = converter.convert(in_path, out_path)
            if not result.get('success'):
                return jsonify({"error": f"File format error – {result.get('error','unknown')}"}), 422

            return send_file(out_path,
                             as_attachment=True,
                             download_name=out_name,
                             mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        finally:
            try: os.remove(in_path)
            except: pass
            # do NOT remove out_path (it’s being streamed)
    except Exception as e:
        app.logger.error("process_payroll failed: %s\n%s", str(e), traceback.format_exc())
        return jsonify({"error": str(e)}), 500

# Frontend static (unchanged)
@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def serve(path):
    static_path = app.static_folder
    if path != "" and os.path.exists(os.path.join(static_path, path)):
        return send_from_directory(static_path, path)
    index_path = os.path.join(static_path, 'index.html')
    if os.path.exists(index_path):
        return send_from_directory(static_path, 'index.html')
    return "index.html not found", 404

if __name__ == '__main__':
    print("Starting Sierra Payroll System...")
    print(f"Gold Master Order loaded: {len(converter.gold_master_order)} employees")
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
