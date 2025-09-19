# src/main.py
import os, sys, traceback
from pathlib import Path
from werkzeug.utils import secure_filename
from flask import Flask, send_from_directory, request, jsonify, send_file
from flask_cors import CORS

# DON'T CHANGE THIS !!!
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from improved_converter import SierraToWBSConverter
from roster_enforcer import enforce_roster
from gold_roster import load_order

app = Flask(__name__, static_folder=os.path.join(os.path.dirname(__file__), 'static'))
app.config['SECRET_KEY'] = 'sierra-payroll-secret-key-2024'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
CORS(app)

UPLOAD_FOLDER = '/tmp/uploads'
ALLOWED_EXTENSIONS = {'xlsx','xls'}
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

DATA_DIR = Path(os.path.dirname(os.path.dirname(__file__))) / "data"
GOLD_MASTER_PATH = DATA_DIR / 'gold_master_order.txt'

converter = SierraToWBSConverter(str(GOLD_MASTER_PATH) if GOLD_MASTER_PATH.exists() else None)

def allowed_file(fn: str) -> bool:
    return '.' in fn and fn.rsplit('.',1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/api/health', methods=['GET'])
def health():
    try:
        order_count = len(load_order())
    except Exception:
        order_count = 0
    return jsonify({
        "status":"ok",
        "version":"2.1.0",
        "gold_master_count": order_count
    })

@app.route('/api/employees', methods=['GET'])
def employees():
    try:
        order = load_order()
        return jsonify([{"id":i+1,"name":n} for i,n in enumerate(order)])
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/validate-sierra-file', methods=['POST'])
def validate():
    try:
        if 'file' not in request.files: return jsonify({"error":"No file provided"}), 400
        f = request.files['file']
        if f.filename == '': return jsonify({"error":"No file selected"}), 400
        if not allowed_file(f.filename): return jsonify({"error":"File must be .xlsx or .xls"}), 400
        # Save temp and use converter parser for count/hours
        temp_path = os.path.join(app.config['UPLOAD_FOLDER'], f"temp_{secure_filename(f.filename)}")
        f.save(temp_path)
        try:
            sierra_df = converter.parse_sierra_file(temp_path)
            valid = not sierra_df.empty
            return jsonify({
                "valid": bool(valid),
                "employees": int(sierra_df['Name'].nunique()) if valid else 0,
                "total_hours": float(sierra_df['Hours'].sum()) if valid else 0.0,
                "total_entries": int(len(sierra_df)) if valid else 0
            })
        finally:
            try: os.remove(temp_path)
            except: pass
    except Exception as e:
        app.logger.error(traceback.format_exc())
        return jsonify({"valid": False, "error": str(e), "employees": 0, "total_hours": 0.0})

@app.route('/api/process-payroll', methods=['POST'])
def process_payroll():
    try:
        if 'file' not in request.files: return jsonify({"error":"No file provided"}), 400
        f = request.files['file']
        if f.filename == '': return jsonify({"error":"No file selected"}), 400
        if not allowed_file(f.filename): return jsonify({"error":"File must be .xlsx or .xls"}), 400

        input_name = secure_filename(f.filename)
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], input_name)
        f.save(input_path)

        output_name = f"WBS_Payroll_{input_name}"
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_name)

        # 1) Convert
        result = converter.convert(input_path, output_path)
        if not result.get('success', False):
            try: os.remove(input_path)
            except: pass
            return jsonify({"error": f"Conversion failed: {result.get('error','unknown')}"}), 422

        # 2) Enforce gold roster order + SSNs + layout
        try:
            enforce_roster(output_path, input_path, str(GOLD_MASTER_PATH))
        except Exception as e:
            app.logger.error(f"Roster enforcement warning: {e}")

        try: os.remove(input_path)
        except: pass

        return send_file(output_path, as_attachment=True, download_name=output_name,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        app.logger.error(traceback.format_exc())
        return jsonify({"error": f"Processing failed: {str(e)}"}), 500

# -------- Static (leave as-is) --------
@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def serve(path):
    static_folder_path = app.static_folder
    if static_folder_path is None:
        return "Static folder not configured", 404
    if path != "" and os.path.exists(os.path.join(static_folder_path, path)):
        return send_from_directory(static_folder_path, path)
    index_path = os.path.join(static_folder_path, 'index.html')
    return send_from_directory(static_folder_path, 'index.html') if os.path.exists(index_path) else ("index.html not found", 404)

if __name__ == '__main__':
    print("Starting Sierra Payroll System...")
    print(f"Gold Master Order loaded: {len(load_order())} employees")
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
