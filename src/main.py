#!/usr/bin/env python3
"""
Sierra Payroll API (stable)

Fixes:
- KeyError: 'UPLOAD_FOLDER'  -> app.config registered + mkdirs
- TypeError: arg must be a list/tuple/Series -> safe numeric coercion
- Validation totals hours from Hours OR REG/OT/DT OR A01/A02/A03 (with aliases)

Keeps existing endpoints/behavior so the frontend works as-is.
"""

import os
import sys
import traceback
from pathlib import Path
from werkzeug.utils import secure_filename

# Make project root importable (DON'T CHANGE)
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from flask import Flask, send_from_directory, request, jsonify, send_file
from flask_cors import CORS
import pandas as pd

from improved_converter import SierraToWBSConverter

# ---------------- Flask app ----------------
app = Flask(
    __name__,
    static_folder=os.path.join(os.path.dirname(__file__), 'static')
)
app.config['SECRET_KEY'] = 'sierra-payroll-secret-key-2024'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB
CORS(app)

# ---------------- Config ----------------
UPLOAD_FOLDER = '/tmp/uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)        # ✅ ensure exists
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER      # ✅ prevent KeyError

# Gold master order path
GOLD_MASTER_PATH = os.path.join(
    os.path.dirname(os.path.dirname(__file__)), 'data', 'gold_master_order.txt'
)

# Converter
converter = SierraToWBSConverter(
    GOLD_MASTER_PATH if Path(GOLD_MASTER_PATH).exists() else None
)

def allowed_file(filename: str) -> bool:
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# ---------------- Validation helpers ----------------
def _coerce_series(df: pd.DataFrame, col: str) -> pd.Series:
    """Return numeric Series for column if present, else a zero Series."""
    if df is None or df.empty:
        return pd.Series(dtype=float)
    if col in df.columns:
        return pd.to_numeric(df[col], errors='coerce').fillna(0.0)
    # zero Series aligned to df index so sums work
    return pd.Series([0.0] * len(df), index=df.index)

def _sum_cols(df: pd.DataFrame, cols) -> float:
    """Sum a list of columns safely; treats missing columns as zeros."""
    if df is None or df.empty:
        return 0.0
    total_series = None
    for c in cols:
        s = _coerce_series(df, c)
        total_series = s if total_series is None else (total_series + s)
    return float(0.0 if total_series is None else total_series.sum())

def _compute_total_hours(df: pd.DataFrame) -> float:
    """
    Robust hours calculator:
    - Prefer explicit 'Hours' / 'Hrs' / 'Total Hours' / 'Total' if that column is numeric hours.
    - Else sum REG/OT/DT variants (REGULAR/OVERTIME/DOUBLETIME, Regular/Overtime/Double Time, REG/OT/DT).
    - Else sum A01/A02/A03.
    - Else sum any columns whose name contains 'hour'.
    Never passes a scalar into a function expecting a Series (prevents TypeError).
    """
    if df is None or df.empty:
        return 0.0

    # 1) explicit single column totals
    for col in ['Hours', 'Hrs', 'Total Hours', 'Total']:
        if col in df.columns:
            val = float(_coerce_series(df, col).sum())
            if val > 0:
                return val

    # 2) triplets that imply total hours
    triplets = [
        ('REGULAR', 'OVERTIME', 'DOUBLETIME'),
        ('Regular', 'Overtime', 'Double Time'),
        ('REG', 'OT', 'DT'),
        ('A01', 'A02', 'A03')
    ]
    for a, b, c in triplets:
        if any(x in df.columns for x in (a, b, c)):
            val = _sum_cols(df, [a, b, c])
            if val > 0:
                return val

    # 3) any 'hour'-ish columns
    hourish = [c for c in df.columns if isinstance(c, str) and 'hour' in c.lower()]
    if hourish:
        val = _sum_cols(df, hourish)
        if val > 0:
            return val

    return 0.0

# ---------------- Routes ----------------
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
            # Use converter parser first (cleans names, filters junk)
            sierra_df = converter.parse_sierra_file(tmp_path)

            emp_count = 0
            if sierra_df is not None and not sierra_df.empty and 'Name' in sierra_df.columns:
                emp_count = int(
                    sierra_df['Name'].astype(str).str.strip().replace('', pd.NA).dropna().nunique()
                )

            total_hours = _compute_total_hours(sierra_df)

            # Safety net: if still zero, attempt raw sheet read (header row 0)
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

# ---------------- Static (frontend) ----------------
@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def serve(path):
    static_folder_path = app.static_folder
    if not static_folder_path:
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

# ---------------- Entrypoint ----------------
if __name__ == '__main__':
    print("Starting Sierra Payroll System...")
    print(f"Gold Master Order loaded: {len(getattr(converter, 'gold_master_order', []))} employees")
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
