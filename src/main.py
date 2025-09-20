import os
import sys
import io
import traceback
from pathlib import Path

# allow "from improved_converter import SierraToWBSConverter"
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
import pandas as pd

from improved_converter import SierraToWBSConverter

# --- app & config ---
app = Flask(__name__, static_folder=os.path.join(os.path.dirname(__file__), 'static'))
CORS(app)

app.config['SECRET_KEY'] = 'sierra-payroll-secret-key-2024'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB
app.config['UPLOAD_FOLDER'] = '/tmp/uploads'
Path(app.config['UPLOAD_FOLDER']).mkdir(parents=True, exist_ok=True)

# --- paths for data files (work both locally and on Railway) ---
HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
DATA = (ROOT / 'data')
if not DATA.exists():
    DATA = (HERE / 'data')
GOLD_MASTER_TXT = DATA / 'gold_master_order.txt'

converter = SierraToWBSConverter(str(GOLD_MASTER_TXT) if GOLD_MASTER_TXT.exists() else None)

# -------- small helpers --------
def _to_num(series_like):
    try:
        return pd.to_numeric(series_like, errors='coerce').fillna(0.0)
    except Exception:
        return pd.Series(dtype='float64')

def _read_sierra_for_validation(xlsx_path: str) -> pd.DataFrame:
    """
    Read Sierra sheet safely and return a frame with at least:
    ['Employee Name','REGULAR','OVERTIME','DOUBLETIME']
    """
    def _try(sheet, header):
        try:
            return pd.read_excel(xlsx_path, sheet_name=sheet, header=header)
        except Exception:
            return None

    # Try WEEKLY at row 8 (Excel row 8 => header=7)
    for sheet in ('WEEKLY', 0):
        for hdr in (7, 6, 0):
            df = _try(sheet, hdr)
            if df is None:
                continue
            df = df.dropna(how='all')

            # find a name column
            name_col = None
            for c in df.columns:
                n = str(c).strip().lower().replace(" ", "")
                if n in ("employeename", "name"):
                    name_col = c
                    break
            if name_col is None:
                continue

            # normalize numeric hour columns with aliases
            reg_col = next((c for c in df.columns if str(c).strip().lower().startswith('regular')), None)
            ot_col  = next((c for c in df.columns if 'overtime' in str(c).strip().lower() or str(c).strip().lower()=='ot'), None)
            dt_col  = next((c for c in df.columns if 'double' in str(c).strip().lower()), None)

            out = pd.DataFrame({
                'Employee Name': df[name_col].astype(str).str.strip()
            })
            out['REGULAR']   = _to_num(df[reg_col]) if reg_col else 0.0
            out['OVERTIME']  = _to_num(df[ot_col]) if ot_col else 0.0
            out['DOUBLETIME']= _to_num(df[dt_col]) if dt_col else 0.0

            # keep only rows that look like employees
            out = out[out['Employee Name'].str.len() > 0]
            # drop obvious header-echo rows
            out = out[out['Employee Name'].str.lower() != 'employee name']

            return out.reset_index(drop=True)

    # If everything failed:
    return pd.DataFrame(columns=['Employee Name','REGULAR','OVERTIME','DOUBLETIME'])

# --------- API routes ----------
@app.route('/api/health', methods=['GET'])
def health():
    return jsonify({
        "status": "ok",
        "version": "2.1.0",
        "converter": "improved_converter",
        "gold_master_loaded": len(converter.gold_master_order) > 0,
        "gold_master_count": len(converter.gold_master_order)
    })

@app.route('/api/employees', methods=['GET'])
def get_employees():
    try:
        return jsonify([{"id": i+1, "name": nm, "ssn": "***-**-{:04d}".format(i), "department":"", "pay_rate":0.0, "status":"A"}
                        for i, nm in enumerate(converter.gold_master_order)])
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/validate-sierra-file', methods=['POST'])
def validate_sierra_file():
    try:
        if 'file' not in request.files:
            return jsonify({"valid": False, "error": "No file provided", "employees":0, "total_hours":0.0})

        f = request.files['file']
        tmp = Path(app.config['UPLOAD_FOLDER']) / f"validate_{f.filename}"
        f.save(tmp)

        try:
            df = _read_sierra_for_validation(str(tmp))
            if df.empty:
                return jsonify({"valid": False, "error": "Could not locate employee rows", "employees":0, "total_hours":0.0})

            totals = (df[['REGULAR','OVERTIME','DOUBLETIME']].sum(numeric_only=True)).sum()
            return jsonify({
                "valid": True,
                "employees": int((df['Employee Name'].str.len() > 0).sum()),
                "total_hours": float(totals),
                "total_entries": int(len(df)),
                "employee_names": df['Employee Name'].dropna().astype(str).str.strip().tolist()[:10]
            })
        finally:
            try: tmp.unlink()
            except: pass

    except Exception as e:
        app.logger.error("Error validating file: %s", traceback.format_exc())
        return jsonify({"valid": False, "error": str(e), "employees":0, "total_hours":0.0})

@app.route('/api/process-payroll', methods=['POST'])
def process_payroll():
    try:
        if 'file' not in request.files:
            return jsonify({"error": "No file provided"}), 400

        f = request.files['file']
        in_path  = Path(app.config['UPLOAD_FOLDER']) / f.filename
        out_path = Path(app.config['UPLOAD_FOLDER']) / f"WBS_Payroll_{Path(f.filename).stem}.xlsx"
        f.save(in_path)

        try:
            result = converter.convert(str(in_path), str(out_path))
            if not result.get('success'):
                return jsonify({"error": f"Conversion failed: {result.get('error','unknown')}"}), 422

            return send_file(
                str(out_path),
                as_attachment=True,
                download_name=out_path.name,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        finally:
            try: in_path.unlink()
            except: pass

    except Exception as e:
        app.logger.error("Error processing payroll: %s", traceback.format_exc())
        return jsonify({"error": "Server error"}), 500

# serve SPA
@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def serve(path):
    static_folder_path = app.static_folder
    if path and os.path.exists(os.path.join(static_folder_path, path)):
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
