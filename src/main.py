import os
import sys
import tempfile
import traceback
from pathlib import Path
from werkzeug.utils import secure_filename

# DON'T CHANGE THIS !!!
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from flask import Flask, send_from_directory, request, jsonify, send_file
from flask_cors import CORS

# Import converter + roster enforcer
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from improved_converter import SierraToWBSConverter
from roster_enforcer import enforce_roster

app = Flask(__name__, static_folder=os.path.join(os.path.dirname(__file__), 'static'))
app.config['SECRET_KEY'] = 'sierra-payroll-secret-key-2024'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB

CORS(app)

UPLOAD_FOLDER = '/tmp/uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# GOLD MASTER ORDER (one name per line) — used to lock the roster order
GOLD_MASTER_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'data', 'gold_master_order.txt')
converter = SierraToWBSConverter(GOLD_MASTER_PATH if Path(GOLD_MASTER_PATH).exists() else None)

def allowed_file(filename: str) -> bool:
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# -------- API --------
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
        return jsonify([{"id": i+1, "name": n, "ssn": f"***-**-{str(i).zfill(4)}", "department": "UNKNOWN", "status": "A"}
                        for i, n in enumerate(converter.gold_master_order)])
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/validate-sierra-file', methods=['POST'])
def validate_sierra_file():
    try:
        if 'file' not in request.files: return jsonify({"error":"No file provided"}), 400
        f = request.files['file']
        if f.filename == '': return jsonify({"error":"No file selected"}), 400
        if not allowed_file(f.filename): return jsonify({"error":"File must be .xlsx or .xls"}), 400

        filename = secure_filename(f.filename)
        temp_path = os.path.join(app.config['UPLOAD_FOLDER'], f"temp_{filename}")
        f.save(temp_path)

        try:
            sierra_df = converter.parse_sierra_file(temp_path)  # returns normalized rows
            if sierra_df.empty:
                return jsonify({"valid": False, "error": "No valid employee data found", "employees": 0, "total_hours": 0.0})
            total_hours = float(sierra_df['Hours'].sum())
            unique_employees = int(sierra_df['Name'].nunique())
            return jsonify({"valid": True, "employees": unique_employees, "total_hours": total_hours,
                            "total_entries": len(sierra_df), "employee_names": sierra_df['Name'].unique().tolist()[:10]})
        finally:
            try: os.remove(temp_path)
            except Exception: pass
    except Exception as e:
        app.logger.error(f"Error validating file: {str(e)}")
        return jsonify({"valid": False, "error": str(e), "employees": 0, "total_hours": 0.0})

@app.route('/api/process-payroll', methods=['POST'])
def process_payroll():
    try:
        if 'file' not in request.files: return jsonify({"error":"No file provided"}), 400
        f = request.files['file']
        if f.filename == '': return jsonify({"error":"No file selected"}), 400
        if not allowed_file(f.filename): return jsonify({"error":"File must be .xlsx or .xls"}), 400

        # Save input
        filename = secure_filename(f.filename)
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        f.save(input_path)

        # Output path
        output_filename = f"WBS_Payroll_{filename}"
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)

        # Convert
        result = converter.convert(input_path, output_path)
        if not result['success']:
            try: os.remove(input_path)
            except Exception: pass
            return jsonify({"error": f"Conversion failed: {result['error']}"}), 422

        # ENFORCE GOLD ROSTER ORDER (adds missing names, preserves columns)
        try:
            enforce_roster(output_path, input_path, GOLD_MASTER_PATH)
        except Exception as e:
            app.logger.error(f"Roster enforcement warning: {e}")

        # Clean up input
        try: os.remove(input_path)
        except Exception: pass

        # Return file
        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        app.logger.error(f"Error processing payroll: {str(e)}")
        app.logger.error(traceback.format_exc())
        return jsonify({"error": f"Processing failed: {str(e)}"}), 500

@app.route('/api/conversion-stats', methods=['GET'])
def get_conversion_stats():
    return jsonify({"total_conversions": 0, "last_conversion": None, "average_employees": 0,
                    "average_hours": 0.0, "status": "operational"})

# -------- STATIC --------
@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def serve(path):
    static_folder_path = app.static_folder
    if static_folder_path is None: return "Static folder not configured", 404
    if path != "" and os.path.exists(os.path.join(static_folder_path, path)):
        return send_from_directory(static_folder_path, path)
    index_path = os.path.join(static_folder_path, 'index.html')
    return send_from_directory(static_folder_path, 'index.html') if os.path.exists(index_path) else ("index.html not found", 404)

if __name__ == '__main__':
    print("Starting Sierra Payroll System...")
    print(f"Gold Master Order loaded: {len(converter.gold_master_order)} employees")
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
