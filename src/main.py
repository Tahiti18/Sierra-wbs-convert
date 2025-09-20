# src/main.py
import os
import sys
import traceback
from pathlib import Path
from typing import Optional, Dict

from flask import Flask, jsonify, request, send_file, send_from_directory
from flask_cors import CORS
from werkzeug.utils import secure_filename

# --- Resolve paths (repo root, /src, /data) ---
HERE = Path(__file__).resolve().parent
REPO_ROOT = HERE.parent if HERE.name == "src" else HERE
DATA_DIR = REPO_ROOT / "data"

# Make sure we can import modules placed at repo root (improved_converter.py)
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

# Import converter (must exist at repo root as improved_converter.py)
try:
    from improved_converter import SierraToWBSConverter  # type: ignore
except Exception as e:
    print("[BOOT] Failed to import improved_converter:", e)
    raise

# --- Flask app ---
app = Flask(
    __name__,
    static_folder=str((REPO_ROOT / "src" / "static") if (REPO_ROOT / "src" / "static").exists() else REPO_ROOT / "static"),
)
app.config["SECRET_KEY"] = "sierra-payroll-secret-key-2024"
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16MB

# Uploads
UPLOAD_DIR = Path(os.environ.get("UPLOAD_DIR", "/tmp/uploads"))
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
app.config["UPLOAD_FOLDER"] = str(UPLOAD_DIR)

# CORS
CORS(app)

ALLOWED_EXTENSIONS = {"xlsx", "xls"}

# --- Converter boot with gold order path (optional but preferred) ---
GOLD_ORDER = DATA_DIR / "gold_master_order.txt"
converter = SierraToWBSConverter(str(GOLD_ORDER) if GOLD_ORDER.exists() else None)

def _ok_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

# ---------- ROUTES ----------

@app.route("/api/health", methods=["GET"])
def health() -> tuple:
    """Simple health + config check."""
    try:
        # probe data files presence (FYI only; not required to run)
        roster_path = DATA_DIR / "gold_master_roster.csv"
        template_path = DATA_DIR / "wbs_template.xlsx"
        return jsonify({
            "status": "ok",
            "version": "2.0.0",
            "gold_master_loaded": bool(getattr(converter, "gold_master_order", [])),
            "gold_master_count": len(getattr(converter, "gold_master_order", [])),
            "has_roster_csv": roster_path.exists(),
            "has_template": template_path.exists(),
            "upload_folder": app.config.get("UPLOAD_FOLDER"),
        }), 200
    except Exception as e:
        return jsonify({"status": "degraded", "error": str(e)}), 200


@app.route("/api/employees", methods=["GET"])
def employees() -> tuple:
    """Expose the gold-master order as an employee list (masked SSNs unless roster present)."""
    try:
        names = list(getattr(converter, "gold_master_order", [])) or []
        roster_csv = DATA_DIR / "gold_master_roster.csv"
        roster: Dict[str, Dict[str, str]] = {}
        if roster_csv.exists():
            # Lightweight load w/o pandas dependency here
            import csv
            def _canon(s: str) -> str:
                s = (s or "").strip()
                return " ".join(s.replace(".", "").replace(" ,", ",").replace(",", ", ").split()).lower()
            with roster_csv.open("r", encoding="utf-8") as f:
                r = csv.DictReader(f)
                for row in r:
                    nm = row.get("Employee Name") or row.get("EmployeeName") or row.get("Name") or ""
                    roster[_canon(nm)] = row

        out = []
        for i, nm in enumerate(names, start=1):
            ssn = "***-**-{:04d}".format(i)  # masked fallback
            if roster:
                key = " ".join(nm.replace(".", "").replace(" ,", ",").replace(",", ", ").split()).lower()
                r = roster.get(key, {})
                raw = (r.get("SSN") or r.get("ssn") or "").strip()
                if raw:
                    # mask if it looks like 9 digits
                    raw_digits = "".join(ch for ch in raw if ch.isdigit())
                    if len(raw_digits) == 9:
                        ssn = f"***-**-{raw_digits[-4:]}"
                    else:
                        ssn = raw  # leave as-is if not 9 digits
            out.append({
                "id": i,
                "name": nm,
                "ssn": ssn,
                "department": r.get("Dept") if roster and (r := roster.get(key, {})) else "UNKNOWN",
                "status": (r.get("Status") or "A") if roster and r else "A",
                "pay_rate": float(r.get("Pay Rate") or r.get("PayRate") or 0.0) if roster and r else 0.0,
            })
        return jsonify(out), 200
    except Exception as e:
        app.logger.error("employees route failed: %s", e)
        return jsonify({"error": str(e)}), 500


@app.route("/api/validate-sierra-file", methods=["POST"])
def validate_sierra_file() -> tuple:
    """Parse the uploaded Sierra file robustly and return employee + hours summary."""
    try:
        if "file" not in request.files:
            return jsonify({"valid": False, "error": "No file provided", "employees": 0, "total_hours": 0.0}), 400

        f = request.files["file"]
        if not f.filename:
            return jsonify({"valid": False, "error": "No file selected", "employees": 0, "total_hours": 0.0}), 400
        if not _ok_file(f.filename):
            return jsonify({"valid": False, "error": "File must be .xlsx or .xls", "employees": 0, "total_hours": 0.0}), 400

        tmp_name = secure_filename(f"validate_{f.filename}")
        tmp_path = UPLOAD_DIR / tmp_name
        f.save(str(tmp_path))

        try:
            df = converter.parse_sierra_file(str(tmp_path))
            if df is None or df.empty:
                return jsonify({"valid": False, "error": "No rows detected", "employees": 0, "total_hours": 0.0}), 200

            # Total hours = REGULAR + OVERTIME + DOUBLETIME (row-wise then summed)
            hours_cols = [c for c in ["REGULAR", "OVERTIME", "DOUBLETIME"] if c in df.columns]
            df["__row_total"] = df[hours_cols].sum(axis=1) if hours_cols else 0.0
            total_hours = float(df["__row_total"].sum())
            unique_emps = int(df["Name"].astype(str).str.strip().str.lower().nunique())

            return jsonify({
                "valid": True,
                "employees": unique_emps,
                "total_hours": total_hours,
                "total_entries": int(len(df)),
                "sample_names": list(df["Name"].astype(str).head(10))
            }), 200
        finally:
            try:
                tmp_path.unlink(missing_ok=True)
            except Exception:
                pass

    except Exception as e:
        app.logger.error("validate_sierra_file failed: %s\n%s", e, traceback.format_exc())
        return jsonify({"valid": False, "error": str(e), "employees": 0, "total_hours": 0.0}), 200


@app.route("/api/process-payroll", methods=["POST"])
def process_payroll() -> tuple:
    """Run full conversion and return the resulting WBS file."""
    try:
        if "file" not in request.files:
            return jsonify({"error": "No file provided"}), 400

        f = request.files["file"]
        if not f.filename:
            return jsonify({"error": "No file selected"}), 400
        if not _ok_file(f.filename):
            return jsonify({"error": "File must be .xlsx or .xls"}), 400

        in_name = secure_filename(f.filename)
        in_path = UPLOAD_DIR / in_name
        f.save(str(in_path))

        out_name = f"WBS_Payroll_{Path(in_name).stem}.xlsx"
        out_path = UPLOAD_DIR / out_name

        # Execute conversion (converter enforces totals column formula if template cell is blank)
        result: Dict = converter.convert(str(in_path), str(out_path))

        # cleanup input
        try:
            in_path.unlink(missing_ok=True)
        except Exception:
            pass

        if not result.get("success"):
            return jsonify({"error": result.get("error", "Conversion failed")}), 422

        # Stream file back
        return send_file(
            str(out_path),
            as_attachment=True,
            download_name=out_name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        app.logger.error("process_payroll failed: %s\n%s", e, traceback.format_exc())
        return jsonify({"error": str(e)}), 500


# ---------- Static passthrough for single-page app builds ----------
@app.route("/", defaults={"path": ""})
@app.route("/<path:path>")
def serve_spa(path: Optional[str] = ""):
    static_dir = Path(app.static_folder) if app.static_folder else None
    if static_dir and path and (static_dir / path).exists():
        return send_from_directory(str(static_dir), path)
    index_path = static_dir / "index.html" if static_dir else None
    if index_path and index_path.exists():
        return send_from_directory(str(static_dir), "index.html")
    return "OK", 200


if __name__ == "__main__":
    print("Starting Sierra Payroll System...")
    print(f"Gold Master Order loaded: {len(getattr(converter, 'gold_master_order', []))} employees")
    port = int(os.environ.get("PORT", "8080"))
    app.run(host="0.0.0.0", port=port, debug=False)
