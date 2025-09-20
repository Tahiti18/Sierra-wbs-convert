#!/usr/bin/env python3
# Flask backend – robust, no hardcoded upload folder, safe temp files,
# counts employees/hours correctly, and uses the improved converter.

from __future__ import annotations
import io
import os
import sys
import tempfile
from pathlib import Path
from typing import Dict

from flask import Flask, jsonify, request, send_file
from flask_cors import CORS

# ----- paths / imports -------------------------------------------------------
ROOT = Path(__file__).resolve().parents[1]  # repo root
SRC  = ROOT / "src"
DATA = ROOT / "data"

# make sure we can import the converter from repo root
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

try:
    from improved_converter import SierraToWBSConverter  # noqa: E402
except Exception as e:
    # fail fast with a clear message in logs
    raise RuntimeError(f"Could not import improved_converter.py: {e}")

# single shared converter instance (loads order.txt once)
ORDER_TXT = DATA / "gold_master_order.txt"
converter = SierraToWBSConverter(str(ORDER_TXT))

app = Flask(__name__)
CORS(app)

# ----- helpers ---------------------------------------------------------------
def _ok(payload: Dict, status: int = 200):
    return jsonify(payload), status

def _err(msg: str, status: int = 400):
    return jsonify({"success": False, "error": msg}), status


# ----- routes ----------------------------------------------------------------
@app.route("/api/health", methods=["GET"])
def health():
    try:
        order_count = len(converter.gold_master_order)
        roster_path = DATA / "gold_master_roster.csv"
        template_ok = (DATA / "wbs_template.xlsx").exists()
        return _ok({
            "success": True,
            "status": "ok",
            "gold_order_count": order_count,
            "roster_present": roster_path.exists(),
            "template_present": template_ok
        })
    except Exception as e:
        return _err(f"health failed: {e}", 500)


@app.route("/api/validate-sierra-file", methods=["POST"])
def validate_sierra_file():
    try:
        if "file" not in request.files:
            return _err("no file part", 422)

        f = request.files["file"]
        if not f or f.filename == "":
            return _err("empty filename", 422)

        # write to a safe temp file (no UPLOAD_FOLDER dependency)
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tf:
            temp_in = Path(tf.name)
            f.save(temp_in)

        try:
            df = converter.parse_sierra_file(str(temp_in))
        finally:
            try:
                temp_in.unlink(missing_ok=True)
            except Exception:
                pass

        if df is None or df.empty:
            return _ok({"success": True, "employees": 0, "total_hours": 0.0})

        # count only rows with any hours
        df["Hours"] = df[["REGULAR", "OVERTIME", "DOUBLETIME"]].sum(axis=1)
        df = df[df["Hours"] > 0]

        employees = int(df.shape[0])
        total_hours = float(df["Hours"].sum())

        return _ok({
            "success": True,
            "employees": employees,
            "total_hours": round(total_hours, 3)
        })
    except Exception as e:
        return _err(f"validate_sierra_file failed: {e}", 500)


@app.route("/api/process-payroll", methods=["POST"])
def process_payroll():
    try:
        if "file" not in request.files:
            return _err("no file part", 422)
        f = request.files["file"]
        if not f or f.filename == "":
            return _err("empty filename", 422)

        # temp input and output
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tf_in:
            temp_in = Path(tf_in.name)
            f.save(temp_in)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tf_out:
            temp_out = Path(tf_out.name)

        try:
            result = converter.convert(str(temp_in), str(temp_out))
        finally:
            try:
                temp_in.unlink(missing_ok=True)
            except Exception:
                pass

        if not result.get("success"):
            try:
                temp_out.unlink(missing_ok=True)
            except Exception:
                pass
            return _err(result.get("error", "conversion failed"), 422)

        # stream the generated workbook back
        buf = io.BytesIO(temp_out.read_bytes())
        try:
            temp_out.unlink(missing_ok=True)
        except Exception:
            pass

        filename = f"WBS_Payroll_{os.path.basename(f.filename).rsplit('.',1)[0]}.xlsx"
        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return _err(f"process_payroll failed: {e}", 500)


if __name__ == "__main__":
    # for local runs
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8080")))
