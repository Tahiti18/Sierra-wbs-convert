import os
import sys
import traceback
from pathlib import Path

from flask import Flask, send_from_directory, request, jsonify, send_file
from flask_cors import CORS
from werkzeug.utils import secure_filename

# import converter (lives at repo root)
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from improved_converter import SierraToWBSConverter, DATA, ORDER_TXT  # noqa

app = Flask(__name__, static_folder=os.path.join(os.path.dirname(__file__), 'static'))
CORS(app)

# --- config ---
app.config['SECRET_KEY'] = 'sierra-payroll-secret-key-2024'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB
UPLOAD_FOLDER = '/tmp/uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
ALLOWED = {'xlsx', 'xls'}

# converter with gold order
converter = SierraToWBSConverter(str(ORDER_TXT))

def _allowed(fn: str) -> bool:
    return '.' in fn and fn.rsplit('.', 1)[1].lower() in ALLOWED

# ---------- api ----------
@app.route('/api/health', methods=['GET'])
def health():
    return jsonify({
        "status": "ok",
        "version": "2.2.0",
        "gold_master_loaded": len(converter.gold_master_order) > 0,
        "gold_master_count": len(converter.gold_master_order),
        "data_dir": str(DATA)
    })

@app.route('/api/employees', methods=['GET'])
def employees():
    return jsonify([
        {"id": i + 1, "name": n, "ssn": "***-**-{:04d}".format(i), "department": "UNKNOWN", "pay_rate": 0.0, "status": "A"}
        for i, n in enumerate(converter.gold_master_order)
    ])

@app.route('/api/validate-sierra-file', methods=['POST'])
def validate_sierra_file():
    try:
        if 'file' not in request.files:
            return jsonify({"valid": False, "error": "No file provided"}), 400
        f = request.files['file']
        if f.filename == '':
            return jsonify({"valid": False, "error": "No file selected"}), 400
        if not _allowed(f.filename):
            return jsonify({"valid": False, "error": "File must be .xlsx or .xls"}), 400

        tmp_name = secure_filename(f"val_{f.filename}")
        tmp_path = os.path.join(app.config['UPLOAD_FOLDER'], tmp_name)
        f.save(tmp_path)

        try:
            df = converter.parse_sierra_file(tmp_path)
            employees = int(df['__canon'].nunique()) if not df.empty else 0
            hours = float(df[['REGULAR', 'OVERTIME', 'DOUBLETIME']].sum().sum()) if not df.empty else 0.0
            return jsonify({"valid": True, "employees": employees, "total_hours": hours})
        finally:
            try: os.remove(tmp_path)
            except Exception: pass

    except Exception as e:
        app.logger.error("ERROR in validate: %s", traceback.format_exc())
        return jsonify({"valid": False, "error": str(e), "employees": 0, "total_hours": 0.0})

@app.route('/api/process-payroll', methods=['POST'])
def process_payroll():
    try:
        if 'file' not in request.files:
            return jsonify({"error": "No file provided"}), 400
        f = request.files['file']
        if f.filename == '':
            return jsonify({"error": "No file selected"}), 400
        if not _allowed(f.filename):
            return jsonify({"error": "File must be .xlsx or .xls"}), 400

        in_name = secure_filename(f.filename)
        in_path = os.path.join(app.config['UPLOAD_FOLDER'], in_name)
        f.save(in_path)

        out_name = f"WBS_Payroll_{in_name}"
        out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_name)

        try:
            result = converter.convert(in_path, out_path)
            if not result.get('success'):
                return jsonify({"error": f"Conversion failed: {result.get('error')}"}), 422

            return send_file(
                out_path,
                as_attachment=True,
                download_name=out_name,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        finally:
            try: os.remove(in_path)
            except Exception: pass

    except Exception as e:
        app.logger.error("ERROR in process: %s", traceback.format_exc())
        return jsonify({"error": str(e)}), 500

# ---------- static ----------
@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def serve(path):
    static_folder_path = app.static_folder
    if path != "" and os.path.exists(os.path.join(static_folder_path, path)):
        return send_from_directory(static_folder_path, path)
    index_path = os.path.join(static_folder_path, 'index.html')
    if os.path.exists(index_path):
        return send_from_directory(static_folder_path, 'index.html')
    return "index.html not found", 404

if __name__ == '__main__':
    print("Starting Sierra Payroll System...")
    print(f"Gold Master Order loaded: {len(converter.gold_master_order)} employees")
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
