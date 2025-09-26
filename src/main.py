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

# Import our WBS accurate converter and multi-stage system
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from wbs_ordered_converter import WBSOrderedConverter as WBSAccurateConverter
from multi_stage_verification import MultiStagePayrollVerification

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

# Initialize WBS accurate converter with gold master order
GOLD_MASTER_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'data', 'gold_master_order.txt')
converter = WBSAccurateConverter()

# Initialize multi-stage verification system
multi_stage = MultiStagePayrollVerification(converter)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# API Routes
@app.route('/api/health', methods=['GET'])
def health():
    """Health check endpoint"""
    return jsonify({
        "status": "ok",
        "version": "2.1.0",
        "converter": "wbs_accurate_converter_v3",
        "gold_master_loaded": len(converter.wbs_order) > 0,
        "gold_master_count": len(converter.wbs_order),
        "employee_database_loaded": len(converter.employee_database) > 0,
        "employee_database_count": len(converter.employee_database),
        "features": {
            "view_mode": "Send format=json in POST to /api/process-payroll",
            "download_mode": "Default behavior of /api/process-payroll",
            "multi_stage_processing": "Available at /api/multi-stage endpoints"
        },
        "multi_stage_endpoints": {
            "full_pipeline": "/api/multi-stage/process-all",
            "stage_1": "/api/multi-stage/stage1-parse",
            "stage_2": "/api/multi-stage/stage2-consolidate", 
            "stage_3": "/api/multi-stage/stage3-overtime",
            "stage_4": "/api/multi-stage/stage4-mapping",
            "stage_5": "/api/multi-stage/stage5-wbs",
            "validation": "/api/multi-stage/validate"
        }
    })

@app.route('/api/employees', methods=['GET'])
def get_employees():
    """Get employee list from gold master order"""
    try:
        employees = []
        for i, name in enumerate(converter.wbs_order):
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
    """Process Sierra payroll file and convert to WBS format - supports both view and download"""
    try:
        # Check if file is present
        if 'file' not in request.files:
            return jsonify({"error": "No file provided"}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({"error": "No file selected"}), 400
        
        if not allowed_file(file.filename):
            return jsonify({"error": "File must be Excel format (.xlsx or .xls)"}), 400
        
        # Check if user wants JSON response (view mode)
        format_type = request.form.get('format', 'excel')  # Default to excel download
        
        # Save uploaded file
        filename = secure_filename(file.filename)
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(input_path)
        
        try:
            # If JSON format requested (view mode), return data instead of file
            if format_type == 'json':
                # Read Sierra file directly for viewing
                import pandas as pd
                df = pd.read_excel(input_path, header=0)
                df.columns = df.columns.astype(str).str.strip()
                
                # Find relevant columns
                name_col = None
                hours_col = None
                rate_col = None
                
                for col in df.columns:
                    if 'name' in col.lower() or 'employee' in col.lower():
                        name_col = col
                    if 'hours' in col.lower() or 'time' in col.lower():
                        hours_col = col
                    if 'rate' in col.lower() or 'pay' in col.lower():
                        rate_col = col
                
                if not all([name_col, hours_col, rate_col]):
                    return jsonify({
                        "error": f"Required columns not found. Found: {df.columns.tolist()}",
                        "name_col": name_col,
                        "hours_col": hours_col, 
                        "rate_col": rate_col
                    }), 400
                
                # Create WBS data for viewing
                wbs_data = []
                for _, row in df.iterrows():
                    employee_name = str(row[name_col]) if pd.notna(row[name_col]) else ""
                    hours = float(row[hours_col]) if pd.notna(row[hours_col]) else 0.0
                    rate = float(row[rate_col]) if pd.notna(row[rate_col]) else 0.0
                    
                    if not employee_name or hours <= 0:
                        continue
                    
                    # Get employee info
                    emp_info = converter.find_employee_info(employee_name)
                    
                    # Apply California overtime rules
                    pay_calc = converter.apply_california_overtime_rules(hours, rate)
                    
                    wbs_data.append({
                        "employee_name": employee_name,
                        "employee_number": emp_info['employee_number'],
                        "ssn": emp_info['ssn'],
                        "department": emp_info['department'],
                        "hours": float(hours),
                        "rate": float(rate),
                        "regular_hours": pay_calc['regular_hours'],
                        "ot15_hours": pay_calc['ot15_hours'],
                        "ot20_hours": pay_calc['ot20_hours'],
                        "regular_amount": pay_calc['regular_amount'],
                        "ot15_amount": pay_calc['ot15_amount'],
                        "ot20_amount": pay_calc['ot20_amount'],
                        "total_amount": pay_calc['total_amount']
                    })
                
                # Find specific employees for debugging
                dianne_data = [entry for entry in wbs_data if 'DIANNE' in entry['employee_name'].upper()]
                
                return jsonify({
                    "success": True,
                    "message": "WBS processing completed successfully - VIEW MODE",
                    "format": "view_only",
                    "total_employees": len(wbs_data),
                    "dianne_debug": dianne_data,
                    "summary": {
                        "total_hours": sum([emp['hours'] for emp in wbs_data]),
                        "total_amount": sum([emp['total_amount'] for emp in wbs_data])
                    },
                    "preview_data": wbs_data[:10],  # First 10 records
                    "full_wbs_data": wbs_data  # Complete data for analysis
                })
            
            # Otherwise, create and return Excel file (download mode)
            else:
                # Create output file path
                output_filename = f"WBS_Payroll_{filename}"
                output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
                
                # Convert file
                result = converter.convert(input_path, output_path)
                
                if not result['success']:
                    return jsonify({"error": f"Conversion failed: {result['error']}"}), 422
                
                # Return the converted file
                return send_file(
                    output_path,
                    as_attachment=True,
                    download_name=output_filename,
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
        
        finally:
            # Clean up input file
            try:
                os.remove(input_path)
            except:
                pass
        
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
            # Read Excel file directly for validation
            import pandas as pd
            df = pd.read_excel(temp_path, header=0)
            
            if df.empty:
                return jsonify({
                    "valid": False,
                    "error": "No valid employee data found",
                    "employees": 0,
                    "total_hours": 0.0
                })
            
            # Find relevant columns
            df.columns = df.columns.astype(str).str.strip()
            name_col = None
            hours_col = None
            
            for col in df.columns:
                if 'name' in col.lower() or 'employee' in col.lower():
                    name_col = col
                if 'hours' in col.lower() or 'time' in col.lower():
                    hours_col = col
            
            if not name_col or not hours_col:
                return jsonify({
                    "valid": False,
                    "error": "Required columns (Name, Hours) not found",
                    "employees": 0,
                    "total_hours": 0.0,
                    "columns_found": df.columns.tolist()
                })
            
            # Calculate stats
            valid_data = df.dropna(subset=[name_col, hours_col])
            total_hours = float(pd.to_numeric(valid_data[hours_col], errors='coerce').sum())
            unique_employees = int(valid_data[name_col].nunique())
            
            return jsonify({
                "valid": True,
                "employees": unique_employees,
                "total_hours": total_hours,
                "total_entries": len(valid_data),
                "employee_names": valid_data[name_col].unique().tolist()[:10],  # First 10 names
                "columns_found": df.columns.tolist(),
                "name_column": name_col,
                "hours_column": hours_col
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

@app.route('/api/view-wbs', methods=['POST'])
def view_wbs():
    """View WBS data without downloading - JSON response only"""
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
        
        try:
            # Parse Sierra file for viewing
            sierra_data = converter.parse_sierra_file(input_path)
            
            # Create WBS data for viewing
            wbs_data = []
            for _, row in sierra_data.iterrows():
                employee_name = row['Employee Name']
                hours = row['Hours']
                rate = row['Rate']
                
                # Get employee info
                emp_info = converter.find_employee_info(employee_name)
                
                # Apply California overtime rules
                pay_calc = converter.apply_california_overtime_rules(hours, rate)
                
                wbs_data.append({
                    "employee_name": employee_name,
                    "employee_number": emp_info['employee_number'],
                    "ssn": emp_info['ssn'],
                    "department": emp_info['department'],
                    "hours": float(hours),
                    "rate": float(rate),
                    "regular_hours": pay_calc['regular_hours'],
                    "ot15_hours": pay_calc['ot15_hours'],
                    "ot20_hours": pay_calc['ot20_hours'],
                    "regular_amount": pay_calc['regular_amount'],
                    "ot15_amount": pay_calc['ot15_amount'],
                    "ot20_amount": pay_calc['ot20_amount'],
                    "total_amount": pay_calc['total_amount']
                })
            
            # Find specific employees for debugging
            dianne_data = [entry for entry in wbs_data if 'DIANNE' in entry['employee_name'].upper()]
            
            return jsonify({
                "success": True,
                "message": "WBS data processed successfully - VIEW ONLY MODE",
                "endpoint": "/api/view-wbs",
                "total_employees": len(wbs_data),
                "dianne_debug": dianne_data,
                "summary": {
                    "total_hours": sum([emp['hours'] for emp in wbs_data]),
                    "total_amount": sum([emp['total_amount'] for emp in wbs_data])
                },
                "preview_data": wbs_data[:10],  # First 10 records
                "full_wbs_data": wbs_data,  # Complete data for analysis
                "note": "Use /api/process-payroll for Excel download"
            })
            
        finally:
            # Clean up input file
            try:
                os.remove(input_path)
            except:
                pass
                
    except Exception as e:
        app.logger.error(f"Error viewing WBS: {str(e)}")
        app.logger.error(traceback.format_exc())
        return jsonify({"error": f"WBS viewing failed: {str(e)}"}), 500

# MULTI-STAGE VERIFICATION ENDPOINTS

@app.route('/api/multi-stage/process-all', methods=['POST'])
def multi_stage_process_all():
    """Process all 5 stages with full verification"""
    try:
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
        
        # Create output path
        output_filename = f"MultiStage_WBS_{filename}"
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
        
        try:
            # Process all stages
            results = multi_stage.process_all_stages(input_path, output_path)
            
            # Check if user wants file download
            format_type = request.form.get('format', 'json')
            
            if format_type == 'excel' and results.get('final_status') == 'success':
                # Return Excel file
                return send_file(
                    output_path,
                    as_attachment=True,
                    download_name=output_filename,
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
            else:
                # Return JSON results
                return jsonify(results)
                
        finally:
            # Clean up input file
            try:
                os.remove(input_path)
            except:
                pass
                
    except Exception as e:
        app.logger.error(f"Multi-stage processing error: {str(e)}")
        return jsonify({"error": f"Multi-stage processing failed: {str(e)}"}), 500

@app.route('/api/multi-stage/stage1-parse', methods=['POST'])
def multi_stage_stage1():
    """Stage 1: Parse raw Sierra data only"""
    try:
        if 'file' not in request.files:
            return jsonify({"error": "No file provided"}), 400
        
        file = request.files['file']
        if not allowed_file(file.filename):
            return jsonify({"error": "File must be Excel format (.xlsx or .xls)"}), 400
        
        filename = secure_filename(file.filename)
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(input_path)
        
        try:
            result = multi_stage.stage1_parse_raw_sierra(input_path)
            return jsonify(result)
        finally:
            try:
                os.remove(input_path)
            except:
                pass
                
    except Exception as e:
        return jsonify({"error": f"Stage 1 failed: {str(e)}"}), 500

@app.route('/api/multi-stage/validate', methods=['GET'])
def multi_stage_validate():
    """Validate cross-stage consistency"""
    try:
        validation_results = multi_stage.validate_cross_stage_consistency()
        return jsonify(validation_results)
    except Exception as e:
        return jsonify({"error": f"Validation failed: {str(e)}"}), 500

@app.route('/api/conversion-stats', methods=['GET'])
def get_conversion_stats():
    """Get statistics about conversions"""
    return jsonify({
        "total_conversions": 0,
        "last_conversion": None,
        "average_employees": 0,
        "average_hours": 0.0,
        "status": "operational",
        "multi_stage_available": True
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
    print(f"WBS Master Order loaded: {len(converter.wbs_order)} employees")
    
    # Railway uses PORT, local development uses FLASK_PORT
    port = int(os.getenv('PORT', os.getenv('FLASK_PORT', 5000)))
    print(f"Starting on port: {port}")
    app.run(host='0.0.0.0', port=port, debug=False)

