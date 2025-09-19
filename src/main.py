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

# Import our improved converter
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from improved_converter import SierraToWBSConverter

app = Flask(__name__, static_folder=os.path.join(os.path.dirname(__file__), 'static'))
app.config['SECRET_KEY'] = 'sierra-payroll-secret-key-2024'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Enable CORS for all routes
CORS(app)

# Configuration
UPLOAD_FOLDER = '/tmp/uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Ensure upload directory exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Initialize converter with gold master order
GOLD_MASTER_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'data', 'gold_master_order.txt')
converter = SierraToWBSConverter(GOLD_MASTER_PATH if Path(GOLD_MASTER_PATH).exists() else None)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# API Routes
@app.route('/api/health', methods=['GET'])
def health():
    """Health check endpoint"""
    return jsonify({
        "status": "ok",
        "version": "2.0.0",
        "converter": "improved_flask",
        "gold_master_loaded": len(converter.gold_master_order) > 0,
        "gold_master_count": len(converter.gold_master_order)
    })

@app.route('/api/employees', methods=['GET'])
def get_employees():
    """Get employee list from gold master order"""
    try:
        employees = []
        for i, name in enumerate(converter.gold_master_order):
            employees.append({
                "id": i + 1,
                "name": name,
                "ssn": f"***-**-{str(i).zfill(4)}",  # Masked SSN
                "department": "UNKNOWN",  # Would come from employee database
                "pay_rate": 0.0,  # Would come from employee database
                "status": "A"
            })
        
        return jsonify(employees)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/process-payroll', methods=['POST'])
def process_payroll():
    """Process Sierra payroll file and convert to WBS format"""
    try:
        # Check if file is present
        if 'file' not in request.files:
            return jsonify({"error": "No file provided"}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({"error": "No file selected"}), 400
        
        if not allowed_file(file.filename):
            return jsonify({"error": "File must be Excel format (.xlsx or .xls)"}), 400
        
        # Save uploaded file
        filename = secure_filename(file.filename)
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(input_path)
        
        # Create output file path
        output_filename = f"WBS_Payroll_{filename}"
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
        
        # Convert file
        result = converter.convert(input_path, output_path)
        
        # Clean up input file
        try:
            os.remove(input_path)
        except:
            pass
        
        if not result['success']:
            return jsonify({"error": f"Conversion failed: {result['error']}"}), 422
        
        # Return the converted file
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

@app.route('/api/validate-sierra-file', methods=['POST'])
def validate_sierra_file():
    """Validate Sierra file format without converting"""
    try:
        if 'file' not in request.files:
            return jsonify({"error": "No file provided"}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({"error": "No file selected"}), 400
        
        if not allowed_file(file.filename):
            return jsonify({"error": "File must be Excel format (.xlsx or .xls)"}), 400
        
        # Save uploaded file temporarily
        filename = secure_filename(file.filename)
        temp_path = os.path.join(app.config['UPLOAD_FOLDER'], f"temp_{filename}")
        file.save(temp_path)
        
        try:
            # Parse file to validate format
            sierra_data = converter.parse_sierra_file(temp_path)
            
            if sierra_data.empty:
                return jsonify({
                    "valid": False,
                    "error": "No valid employee data found",
                    "employees": 0,
                    "total_hours": 0.0
                })
            
            # Calculate stats
            total_hours = float(sierra_data['Hours'].sum())
            unique_employees = int(sierra_data['Name'].nunique())
            
            return jsonify({
                "valid": True,
                "employees": unique_employees,
                "total_hours": total_hours,
                "total_entries": len(sierra_data),
                "employee_names": sierra_data['Name'].unique().tolist()[:10]  # First 10 names
            })
            
        finally:
            # Clean up temp file
            try:
                os.remove(temp_path)
            except:
                pass
                
    except Exception as e:
        app.logger.error(f"Error validating file: {str(e)}")
        return jsonify({
            "valid": False,
            "error": str(e),
            "employees": 0,
            "total_hours": 0.0
        })

@app.route('/api/conversion-stats', methods=['GET'])
def get_conversion_stats():
    """Get statistics about conversions"""
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

    if path != "" and os.path.exists(os.path.join(static_folder_path, path)):
        return send_from_directory(static_folder_path, path)
    else:
        index_path = os.path.join(static_folder_path, 'index.html')
        if os.path.exists(index_path):
            return send_from_directory(static_folder_path, 'index.html')
        else:
            return "index.html not found", 404

if __name__ == '__main__':
    print("Starting Sierra Payroll System...")
    print(f"Gold Master Order loaded: {len(converter.gold_master_order)} employees")
    port = int(os.environ.get(\'PORT\', 5000))
    app.run(host=\'0.0.0.0\', port=port, debug=False)

