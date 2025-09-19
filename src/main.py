import os
import sys
import traceback
from pathlib import Path
from werkzeug.utils import secure_filename

# DON'T CHANGE THIS !!!
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from flask import Flask, send_from_directory, request, jsonify, send_file
from flask_cors import CORS
from improved_converter import SierraToWBSConverter
import pandas as pd

app = Flask(__name__, static_folder=os.path.join(os.path.dirname(__file__), 'static'))
app.config['SECRET_KEY'] = 'sierra-payroll-secret-key-2024'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
CORS(app)

# Configuration
UPLOAD_FOLDER = '/tmp/uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Initialize converter with gold master order
GOLD_MASTER_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'data', 'gold_master_order.txt')
converter = SierraToWBSConverter(GOLD_MASTER_PATH if Path(GOLD_MASTER_PATH).exists() else None)

def allowed_file(filename: str) -> bool:
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def _coerce_num(s):
    return pd.to_numeric(s, errors='coerce').fillna(0.0)

def _compute_total_hours(df: pd.DataFrame) -> float:
    """
    Robust hours calculator:
    - Prefer explicit 'Hours'
    - Else sum REG/OT/DT variants
    - Else sum A01/A02/A03 (template style)
    """
    if df is None or df.empty:
        return 0.0

    # 1) Single 'Hours' column
    for col in ['Hours', 'Hrs', 'Total Hours', 'Total']:
        if col in df.columns:
            tot = float(_coerce_num(df[col]).sum())
            if tot > 0:
                return tot

    # 2) Triplets that imply total hours
    triplet_aliases = [
        ('REGULAR', 'OVERTIME', 'DOUBLETIME'),
        ('Regular', 'Overtime', 'Double Time'),
        ('REG', 'OT', 'DT'),
        ('A01', 'A02', 'A03')
    ]
    for a, b, c in triplet_aliases:
        if a in df.columns and b in df.columns and c in df.columns:
            tot = float(_coerce_num(df[[a, b, c]]).sum().sum())
            if tot > 0:
                return tot

    # 3) Fallback: sum any column whose name includes 'hour'
    hourish = [c for c in df.columns if isinstance(c, str) and 'hour' in c.lower()]
    if hourish:
        tot = float(_coerce_num(df[hourish]).sum().sum())
        if tot > 0:
            return tot

    return 0.0

@app.route('/api/health', methods=['GET'])
def health():
    return jsonify({
        "status": "ok",
        "version": "2.0.0",
        "converter": "improved_flask",
        "gold_master_loaded": bool(getattr(converter, "gold_master_order", [])),
        "gold_master_count": len(getattr(converter, "gold_master_order", []))
    })

@app.route('/api/employees', methods=['GET'])
def get_employees():
    try:
        employees = []
        for i, name in enumerate(getattr(converter, "gold_master_order", [])):
            employees.append({
                "id": i + 1,
                "name": name,
                "ssn": f"***-**-{str(i).zfill(4)}",
                "department": "UNKNOWN",
                "pay_rate": 0.0,
                "status": "A"
            })
        return jsonify(employees)
    except Exception as e:
        app.logger.exception("employees endpoint failed")
        return jsonify({"error": str(e)}), 500

@app.route('/api/process-payroll', methods=['POST'])
def process_payroll():
    try:
        if 'file' not in request.files:
            return jsonify({"error": "No file provided"}), 400
        file = request.files['file']
        if file.filename == '':
            return jsonify({"error": "No file selected"}), 400
        if not allowed_file(file.filename):
            return jsonify({"error": "File must be Excel format (.xlsx or .xls)"}), 400

        in_name = secure_filename(file.filename)
        in_path = os.path.join(app.config['UPLOAD_FOLDER'], in_name)
        file.save(in_path)

        out_name = f"WBS_Payroll_{os.path.splitext(in_name)[0]}.xlsx"
        out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_name)

        result = converter.convert(in_path, out_path)

        try:
            os.remove(in_path)
        except Exception:
            pass

        if not result.get('success'):
            return jsonify({"error": f"Conversion failed: {result.get('error', 'unknown error')}"}), 422

        return send_file(
            out_path,
            as_attachment=True,
            download_name=out_name,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        app.logger.exception("process_payroll failed")
        return jsonify({"error": f"Processing failed: {str(e)}"}), 500

@app.route('/api/validate-sierra-file', methods=['POST'])
def validate_sierra_file():
    try:
        if 'file' not in request.files:
            return jsonify({"error": "No file provided"}), 400
        file = request.files['file']
        if file.filename == '':
            return jsonify({"error": "No file selected"}), 400
        if not allowed_file(file.filename):
            return jsonify({"error": "File must be Excel format (.xlsx or .xls)"}), 400

        tmp_name = secure_filename(file.filename)
        tmp_path = os.path.join(app.config['UPLOAD_FOLDER'], f"temp_{tmp_name}")
        file.save(tmp_path)

        try:
            sierra_df = converter.parse_sierra_file(tmp_path)
            # employee count (unique names with non-empty)
            emp_count = int(sierra_df['Name'].astype(str).str.strip().replace('', pd.NA).dropna().nunique()) if not sierra_df.empty else 0
            # robust total hours
            total_hours = _compute_total_hours(sierra_df)

            # If still zero, try to recompute from potentially named columns in raw excel
            # (Safety net — won’t throw if parse already cleaned)
            if total_hours == 0:
                try:
                    raw = pd.read_excel(tmp_path, sheet_name=0, header=0)
                    total_hours = _compute_total_hours(raw)
                except Exception:
                    pass

            return jsonify({
                "valid": emp_count > 0,
                "employees": emp_count,
                "total_hours": float(round(total_hours, 2)),
                "total_entries": int(len(sierra_df)) if sierra_df is not None else 0,
                "error": None if emp_count > 0 else "No valid employee rows detected"
            })
        finally:
            try:
                os.remove(tmp_path)
            except Exception:
                pass

    except Exception as e:
        app.logger.exception("validate_sierra_file failed")
        return jsonify({
            "valid": False,
            "error": str(e),
            "employees": 0,
            "total_hours": 0.0
        })

@app.route('/api/conversion-stats', methods=['GET'])
def get_conversion_stats():
    return jsonify({
        "total_conversions": 0,
        "last_conversion": None,
        "average_employees": 0,
        "average_hours": 0.0,
        "status": "operational"
    })

# Frontend serving routes
@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def serve(path):
    static_folder_path = app.static_folder
    if static_folder_path is None:
        return "Static folder not configured", 404

    file_path = os.path.join(static_folder_path, path)
    if path != "" and os.path.exists(file_path):
        return send_from_directory(static_folder_path, path)
    else:
        index_path = os.path.join(static_folder_path, 'index.html')
        if os.path.exists(index_path):
            return send_from_directory(static_folder_path, 'index.html')
        else:
            return "index.html not found", 404

if __name__ == '__main__':
    print("Starting Sierra Payroll System...")
    print(f"Gold Master Order loaded: {len(getattr(converter, 'gold_master_order', []))} employees")
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
