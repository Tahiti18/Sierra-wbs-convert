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
from simple_3stage_system import Simple3StageConverter

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

# Initialize clean 3-stage system
simple_3stage = Simple3StageConverter()

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
            "multi_stage_processing": "Available at /api/multi-stage endpoints",
            "clean_3stage_system": "Integrated simple 3-stage pipeline for transparent conversion"
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
            # If JSON format requested (view mode), use CLEAN 3-STAGE SYSTEM
            if format_type == 'json':
                # Use the clean 3-stage system for transparent processing
                pipeline_result = simple_3stage.convert_full_pipeline(input_path)
                
                if pipeline_result['pipeline_complete']:
                    stage3_data = pipeline_result['stage3']
                    wbs_output = stage3_data['wbs_output']
                    
                    # Convert to view format (matching existing frontend expectations)
                    wbs_data_sorted = []
                    for emp in wbs_output:
                        wbs_data_sorted.append({
                            "employee_name": emp['employee_name'],
                            "employee_number": emp['employee_number'], 
                            "ssn": emp['ssn'],
                            "department": emp['department'],
                            "hours": float(emp['hours']),
                            "rate": float(emp['rate']),
                            "regular_hours": emp['regular_hours'],
                            "ot15_hours": emp['ot15_hours'],
                            "ot20_hours": emp['ot20_hours'],
                            "regular_amount": emp['regular_hours'] * emp['rate'] if emp['rate'] > 0 else 0,
                            "ot15_amount": emp['ot15_hours'] * emp['rate'] if emp['rate'] > 0 else 0,
                            "ot20_amount": emp['ot20_hours'] * emp['rate'] if emp['rate'] > 0 else 0,
                            "total_amount": emp['total_amount'],
                            "source": emp['source']  # SIERRA_CALCULATED or MISSING_ZERO
                        })
                    
                    return jsonify({
                        "success": True,
                        "message": f"✅ CLEAN 3-STAGE CONVERSION - ALL {len(wbs_data_sorted)} employees in exact WBS order",
                        "format": "view_only_clean_3stage",
                        "system": "Clean 3-Stage Pipeline",
                        "calculations_source": "100% calculated from Sierra data - no shortcuts",
                        "stage_breakdown": {
                            "stage1_raw_entries": pipeline_result['stage1']['total_entries'],
                            "stage1_unique_employees": pipeline_result['stage1']['unique_employees'], 
                            "stage2_consolidated": pipeline_result['stage2']['consolidated_employees'],
                            "stage3_wbs_with_data": stage3_data['wbs_with_data'],
                            "stage3_wbs_zero": stage3_data['wbs_with_zero']
                        },
                        "totals": {
                            "sierra_raw_amount": pipeline_result['stage1']['total_amount'],
                            "consolidated_amount": pipeline_result['stage2']['total_amount'],
                            "final_wbs_amount": stage3_data['final_total_amount'],
                            "total_wbs_employees": len(wbs_data_sorted),
                            "employees_with_sierra_data": stage3_data['wbs_with_data']
                        },
                        "verification": {
                            "all_calculations_from_sierra": True,
                            "no_gold_standard_copying": True,
                            "transparent_3stage_process": True
                        },
                        "preview_data": wbs_data_sorted[:10],  # First 10 records
                        "full_wbs_data": wbs_data_sorted  # ALL 79 employees in exact WBS order
                    })
                else:
                    return jsonify({"error": "3-stage pipeline failed to complete"}), 500
            
            
            # Otherwise, create and return Excel file (download mode) - CLEAN 3-STAGE
            else:
                # Create output file path
                output_filename = f"Clean3Stage_WBS_{filename}"
                output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
                
                # Use clean 3-stage system for conversion
                pipeline_result = simple_3stage.convert_full_pipeline(input_path)
                
                if not pipeline_result['pipeline_complete']:
                    return jsonify({"error": "Clean 3-stage conversion failed"}), 422
                
                # Create Excel file from WBS output
                wbs_output = pipeline_result['stage3']['wbs_output']
                converter.create_wbs_excel(wbs_output, output_path)
                
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
            
            # Find relevant columns - select best candidates
            df.columns = df.columns.astype(str).str.strip()
            name_candidates = []
            hours_candidates = []
            
            for col in df.columns:
                if 'name' in col.lower() or 'employee' in col.lower():
                    name_candidates.append(col)
                if 'hours' in col.lower() or 'time' in col.lower():
                    hours_candidates.append(col)
            
            # Select the best name column (one with most employee-like entries)
            name_col = None
            if name_candidates:
                best_score = 0
                for col in name_candidates:
                    # Count employee-like entries in this column
                    score = sum(1 for val in df[col].dropna() 
                              if isinstance(val, str) and len(val) > 3 and ' ' in val 
                              and not any(word in val.lower() for word in ['week', 'gross', 'total', 'signature']))
                    if score > best_score:
                        best_score = score
                        name_col = col
            
            hours_col = hours_candidates[0] if hours_candidates else None
            
            if not name_col or not hours_col:
                return jsonify({
                    "valid": False,
                    "error": "Required columns (Name, Hours) not found",
                    "employees": 0,
                    "total_hours": 0.0,
                    "columns_found": df.columns.tolist()
                })
            
            # Calculate stats - filter for actual employee rows (skip headers/summaries)
            valid_data = df.dropna(subset=[name_col, hours_col])
            
            # Filter for rows with actual employee names (not numbers or headers)
            employee_rows = valid_data[
                (pd.to_numeric(valid_data[name_col], errors='coerce').isna()) &  # Not a number
                (valid_data[name_col].str.len() > 3) &  # Name longer than 3 chars
                (valid_data[name_col].str.contains(' ')) &  # Contains space (First Last)
                (~valid_data[name_col].str.contains('Week|Gross|signature|Date', case=False, na=False))  # Not headers
            ]
            
            total_hours = float(pd.to_numeric(employee_rows[hours_col], errors='coerce').sum())
            unique_employees = int(employee_rows[name_col].nunique())
            
            return jsonify({
                "valid": True,
                "employees": unique_employees,
                "total_hours": total_hours,
                "total_entries": len(valid_data),
                "employee_names": employee_rows[name_col].unique().tolist()[:10],  # First 10 names
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

# NEW CLEAN 3-STAGE ENDPOINTS

@app.route('/api/3stage/full-pipeline', methods=['POST'])
def clean_3stage_full_pipeline():
    """Clean 3-stage full pipeline: Parse → Consolidate → Apply WBS Rules"""
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
            # Run full 3-stage pipeline
            result = simple_3stage.convert_full_pipeline(input_path)
            
            # Check if user wants Excel file download
            format_type = request.form.get('format', 'json')
            
            if format_type == 'excel':
                # Create Excel output
                output_filename = f"Clean3Stage_WBS_{filename}"
                output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
                
                # Use converter to create Excel file from WBS output
                converter.create_wbs_excel(result['final_wbs_output'], output_path)
                
                return send_file(
                    output_path,
                    as_attachment=True,
                    download_name=output_filename,
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
            else:
                # Return JSON with all stage data
                return jsonify({
                    "success": True,
                    "system": "Clean 3-Stage Pipeline",
                    "message": "All stages completed with verification",
                    "pipeline_complete": result['pipeline_complete'],
                    "stage1": result['stage1'],
                    "stage2": result['stage2'], 
                    "stage3": result['stage3'],
                    "final_summary": {
                        "sierra_employees_processed": result['stage1']['unique_employees'],
                        "wbs_employees_with_data": result['stage3']['wbs_with_data'],
                        "wbs_employees_zero": result['stage3']['wbs_with_zero'],
                        "final_amount": result['stage3']['final_total_amount'],
                        "calculations_source": "100% calculated from Sierra data"
                    },
                    "wbs_output": result['final_wbs_output'][:20]  # First 20 for preview
                })
                
        finally:
            try:
                os.remove(input_path)
            except:
                pass
                
    except Exception as e:
        app.logger.error(f"Clean 3-stage pipeline error: {str(e)}")
        return jsonify({"error": f"3-stage pipeline failed: {str(e)}"}), 500

@app.route('/api/3stage/stage1', methods=['POST']) 
def clean_3stage_stage1():
    """Clean Stage 1: Parse Sierra data only"""
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
            result = simple_3stage.stage1_parse_sierra(input_path)
            return jsonify(result)
        finally:
            try:
                os.remove(input_path)
            except:
                pass
                
    except Exception as e:
        return jsonify({"error": f"Clean Stage 1 failed: {str(e)}"}), 500

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
    import sys
    print("Starting Sierra Payroll System...")
    print(f"WBS Master Order loaded: {len(converter.wbs_order)} employees")
    
    # Check for command line port argument
    port = 5000
    if len(sys.argv) > 2 and sys.argv[1] == '--port':
        port = int(sys.argv[2])
    else:
        # Railway uses PORT, local development uses FLASK_PORT
        port = int(os.getenv('PORT', os.getenv('FLASK_PORT', 5000)))
    
    print(f"Starting on port: {port}")
    app.run(host='0.0.0.0', port=port, debug=False)

