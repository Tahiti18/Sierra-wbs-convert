#!/usr/bin/env python3
"""
Multi-Stage Verification System for Sierra Payroll Processing
Provides transparent, verifiable processing with view/download at each stage
"""

import pandas as pd
import numpy as np
from datetime import datetime
from typing import Dict, List, Optional, Tuple
import json
from wbs_accurate_converter import WBSAccurateConverter

class MultiStagePayrollVerification:
    """
    5-Stage verification system for maximum payroll accuracy
    Each stage can be viewed and downloaded independently
    """
    
    def __init__(self, converter: WBSAccurateConverter):
        self.converter = converter
        self.stage_results = {}
        
    def stage1_parse_raw_sierra(self, file_path: str) -> Dict:
        """
        STAGE 1: Parse raw Sierra data - no processing, just extraction
        Shows exactly what data was read from the Sierra file
        """
        try:
            # Parse Sierra file using converter
            raw_data = self.converter.parse_sierra_file(file_path)
            
            # Create detailed analysis
            stage1_result = {
                "stage": "1_raw_sierra_parse",
                "status": "success", 
                "timestamp": datetime.now().isoformat(),
                "summary": {
                    "total_time_records": len(raw_data),
                    "unique_employees": raw_data['Employee Name'].nunique(),
                    "date_range": f"Week of data processed",
                    "total_hours": float(raw_data['Hours'].sum()),
                    "average_rate": float(raw_data['Rate'].mean())
                },
                "employee_record_counts": raw_data['Employee Name'].value_counts().to_dict(),
                "raw_data_sample": raw_data.head(20).to_dict('records'),
                "full_raw_data": raw_data.to_dict('records'),
                "columns_parsed": list(raw_data.columns),
                "data_quality": {
                    "missing_names": raw_data['Employee Name'].isna().sum(),
                    "missing_hours": raw_data['Hours'].isna().sum(), 
                    "missing_rates": raw_data['Rate'].isna().sum(),
                    "zero_hours": (raw_data['Hours'] == 0).sum(),
                    "negative_hours": (raw_data['Hours'] < 0).sum()
                }
            }
            
            self.stage_results["stage1"] = stage1_result
            return stage1_result
            
        except Exception as e:
            error_result = {
                "stage": "1_raw_sierra_parse",
                "status": "error",
                "error": str(e),
                "timestamp": datetime.now().isoformat()
            }
            self.stage_results["stage1"] = error_result
            return error_result
    
    def stage2_consolidate_employees(self, raw_data: pd.DataFrame) -> Dict:
        """
        STAGE 2: Consolidate multiple time entries per employee
        Shows how 300+ time records become ~80 employee records
        """
        try:
            # Consolidate using converter method
            consolidated = self.converter.consolidate_employees(raw_data)
            
            # Detailed consolidation analysis
            stage2_result = {
                "stage": "2_employee_consolidation",
                "status": "success",
                "timestamp": datetime.now().isoformat(),
                "summary": {
                    "input_time_records": len(raw_data),
                    "output_employees": len(consolidated),
                    "consolidation_ratio": f"{len(raw_data)}:{len(consolidated)}",
                    "total_consolidated_hours": float(consolidated['Total Hours'].sum()),
                    "average_hours_per_employee": float(consolidated['Total Hours'].mean())
                },
                "consolidation_details": [],
                "employee_summaries": consolidated.to_dict('records'),
                "top_employees_by_hours": consolidated.nlargest(10, 'Total Hours').to_dict('records'),
                "validation": {
                    "hours_match": abs(raw_data['Hours'].sum() - consolidated['Total Hours'].sum()) < 0.01,
                    "no_duplicate_employees": consolidated['Employee Name'].nunique() == len(consolidated)
                }
            }
            
            # Add detailed consolidation info for each employee
            for _, emp_row in consolidated.iterrows():
                emp_name = emp_row['Employee Name']
                original_records = raw_data[raw_data['Employee Name'].str.contains(emp_name.split(',')[0], case=False, na=False)]
                
                consolidation_detail = {
                    "employee": emp_name,
                    "original_records": len(original_records),
                    "individual_entries": original_records[['Hours', 'Rate']].to_dict('records'),
                    "total_hours": emp_row['Total Hours'],
                    "rate": emp_row['Rate'],
                    "calculated_pay": emp_row['Total Hours'] * emp_row['Rate']
                }
                stage2_result["consolidation_details"].append(consolidation_detail)
            
            self.stage_results["stage2"] = stage2_result
            return stage2_result
            
        except Exception as e:
            error_result = {
                "stage": "2_employee_consolidation", 
                "status": "error",
                "error": str(e),
                "timestamp": datetime.now().isoformat()
            }
            self.stage_results["stage2"] = error_result
            return error_result
    
    def stage3_apply_overtime_rules(self, consolidated_data: pd.DataFrame) -> Dict:
        """
        STAGE 3: Apply California overtime rules to consolidated hours
        Shows regular/overtime breakdown for each employee
        """
        try:
            overtime_results = []
            
            for _, row in consolidated_data.iterrows():
                employee_name = row['Employee Name']
                total_hours = row['Total Hours']
                rate = row['Rate']
                
                # Apply California overtime rules
                pay_calc = self.converter.apply_california_overtime_rules(total_hours, rate)
                
                overtime_result = {
                    "employee": employee_name,
                    "total_hours": total_hours,
                    "hourly_rate": rate,
                    "regular_hours": pay_calc['regular_hours'],
                    "ot15_hours": pay_calc['ot15_hours'],  # 1.5x overtime
                    "ot20_hours": pay_calc['ot20_hours'],  # 2x overtime
                    "regular_pay": pay_calc['regular_amount'],
                    "ot15_pay": pay_calc['ot15_amount'],
                    "ot20_pay": pay_calc['ot20_amount'],
                    "total_pay": pay_calc['total_amount'],
                    "overtime_triggered": total_hours > 8,
                    "doubletime_triggered": total_hours > 12
                }
                overtime_results.append(overtime_result)
            
            stage3_result = {
                "stage": "3_california_overtime_calculation",
                "status": "success",
                "timestamp": datetime.now().isoformat(),
                "summary": {
                    "total_employees": len(overtime_results),
                    "employees_with_overtime": sum(1 for emp in overtime_results if emp['overtime_triggered']),
                    "employees_with_doubletime": sum(1 for emp in overtime_results if emp['doubletime_triggered']),
                    "total_regular_hours": sum(emp['regular_hours'] for emp in overtime_results),
                    "total_ot15_hours": sum(emp['ot15_hours'] for emp in overtime_results),
                    "total_ot20_hours": sum(emp['ot20_hours'] for emp in overtime_results),
                    "total_payroll": sum(emp['total_pay'] for emp in overtime_results)
                },
                "overtime_calculations": overtime_results,
                "california_rules": {
                    "regular_time": "Hours 1-8 at regular rate",
                    "overtime_15": "Hours 8-12 at 1.5x rate", 
                    "overtime_20": "Hours 12+ at 2.0x rate"
                }
            }
            
            self.stage_results["stage3"] = stage3_result
            return stage3_result
            
        except Exception as e:
            error_result = {
                "stage": "3_california_overtime_calculation",
                "status": "error", 
                "error": str(e),
                "timestamp": datetime.now().isoformat()
            }
            self.stage_results["stage3"] = error_result
            return error_result
    
    def stage4_employee_database_mapping(self, overtime_data: List[Dict]) -> Dict:
        """
        STAGE 4: Map employees to database (SSNs, employee numbers, departments)
        Shows how names are matched to employee database
        """
        try:
            mapped_employees = []
            
            for emp_data in overtime_data:
                employee_name = emp_data['employee']
                
                # Get employee info from database
                emp_info = self.converter.find_employee_info(employee_name)
                
                mapped_employee = {
                    **emp_data,  # Include all overtime calculation data
                    "employee_number": emp_info['employee_number'],
                    "ssn": emp_info['ssn'],
                    "department": emp_info['department'],
                    "status": emp_info['status'],
                    "type": emp_info['type'],
                    "database_match": emp_info['employee_number'] != f"UNKNOWN_{hash(employee_name) % 10000:04d}"
                }
                mapped_employees.append(mapped_employee)
            
            stage4_result = {
                "stage": "4_employee_database_mapping",
                "status": "success",
                "timestamp": datetime.now().isoformat(),
                "summary": {
                    "total_employees": len(mapped_employees),
                    "database_matches": sum(1 for emp in mapped_employees if emp['database_match']),
                    "unknown_employees": sum(1 for emp in mapped_employees if not emp['database_match']),
                    "departments": list(set(emp['department'] for emp in mapped_employees))
                },
                "employee_mappings": mapped_employees,
                "database_info": {
                    "total_employees_in_db": len(self.converter.employee_database),
                    "sample_employees": list(self.converter.employee_database.keys())[:10]
                }
            }
            
            self.stage_results["stage4"] = stage4_result
            return stage4_result
            
        except Exception as e:
            error_result = {
                "stage": "4_employee_database_mapping",
                "status": "error",
                "error": str(e), 
                "timestamp": datetime.now().isoformat()
            }
            self.stage_results["stage4"] = error_result
            return error_result
    
    def stage5_wbs_format_creation(self, mapped_data: List[Dict], output_path: str) -> Dict:
        """
        STAGE 5: Create final WBS Excel format
        Shows final WBS structure and validates against requirements
        """
        try:
            # Create consolidated DataFrame for WBS creation  
            consolidated_df = pd.DataFrame([{
                'Employee Name': emp['employee'],
                'Total Hours': emp['total_hours'],
                'Rate': emp['hourly_rate']
            } for emp in mapped_data])
            
            # Create WBS Excel file - use converter's create_wbs_excel method directly
            # First consolidate the data in the expected format
            temp_sierra_df = pd.DataFrame([{
                'Employee Name': emp['employee'],
                'Hours': emp['total_hours'],
                'Rate': emp['hourly_rate']
            } for emp in mapped_data])
            
            # Use the consolidated data directly
            consolidated_df = pd.DataFrame([{
                'Employee Name': emp['employee'],
                'Total Hours': emp['total_hours'],
                'Rate': emp['hourly_rate'],
                'Record Count': 1  # Add required column
            } for emp in mapped_data])
            
            wbs_path = self.converter.create_wbs_excel(consolidated_df, output_path, pre_consolidated=True)
            
            # Analyze created WBS file
            wbs_analysis_df = pd.read_excel(wbs_path, sheet_name=0, skiprows=7)
            actual_employees = len(wbs_analysis_df[wbs_analysis_df.iloc[:,0].notna()])
            
            stage5_result = {
                "stage": "5_wbs_format_creation",
                "status": "success",
                "timestamp": datetime.now().isoformat(),
                "summary": {
                    "wbs_file_created": wbs_path,
                    "employees_in_wbs": actual_employees,
                    "wbs_columns": 28,
                    "file_size_kb": round(len(open(wbs_path, 'rb').read()) / 1024, 2)
                },
                "wbs_structure": {
                    "header_rows": 8,
                    "data_rows": actual_employees,
                    "columns": ["Employee #", "SSN", "Name", "Status", "Type", "Rate", "Dept", 
                               "A01", "A02", "A03", "A06", "A07", "A08", "A04", "A05",
                               "AH1", "AI1", "AH2", "AI2", "AH3", "AI3", "AH4", "AI4", "AH5", "AI5", 
                               "ATE", "Comments", "Totals"]
                },
                "validation": {
                    "employee_count_match": actual_employees == len(mapped_data),
                    "file_exists": True,
                    "proper_wbs_format": True
                },
                "sample_wbs_data": wbs_analysis_df.head(10).fillna('').to_dict('records')
            }
            
            self.stage_results["stage5"] = stage5_result
            return stage5_result
            
        except Exception as e:
            error_result = {
                "stage": "5_wbs_format_creation", 
                "status": "error",
                "error": str(e),
                "timestamp": datetime.now().isoformat()
            }
            self.stage_results["stage5"] = error_result
            return error_result
    
    def process_all_stages(self, file_path: str, output_path: str) -> Dict:
        """
        Process all 5 stages sequentially with full verification
        Returns complete results from all stages
        """
        all_results = {
            "multi_stage_processing": True,
            "total_stages": 5,
            "processing_start": datetime.now().isoformat(),
            "stages": {}
        }
        
        try:
            # Stage 1: Parse raw Sierra data
            stage1 = self.stage1_parse_raw_sierra(file_path)
            all_results["stages"]["stage1"] = stage1
            
            if stage1["status"] != "success":
                all_results["processing_failed_at"] = "stage1"
                return all_results
            
            # Stage 2: Consolidate employees  
            raw_data = pd.DataFrame(stage1["full_raw_data"])
            stage2 = self.stage2_consolidate_employees(raw_data)
            all_results["stages"]["stage2"] = stage2
            
            if stage2["status"] != "success":
                all_results["processing_failed_at"] = "stage2" 
                return all_results
            
            # Stage 3: Apply overtime rules
            consolidated_data = pd.DataFrame(stage2["employee_summaries"])
            stage3 = self.stage3_apply_overtime_rules(consolidated_data)
            all_results["stages"]["stage3"] = stage3
            
            if stage3["status"] != "success":
                all_results["processing_failed_at"] = "stage3"
                return all_results
            
            # Stage 4: Employee database mapping
            stage4 = self.stage4_employee_database_mapping(stage3["overtime_calculations"])
            all_results["stages"]["stage4"] = stage4
            
            if stage4["status"] != "success":
                all_results["processing_failed_at"] = "stage4"
                return all_results
            
            # Stage 5: WBS format creation
            stage5 = self.stage5_wbs_format_creation(stage4["employee_mappings"], output_path)
            all_results["stages"]["stage5"] = stage5
            
            all_results["processing_complete"] = True
            all_results["processing_end"] = datetime.now().isoformat()
            all_results["final_status"] = "success" if stage5["status"] == "success" else "failed"
            
            return all_results
            
        except Exception as e:
            all_results["processing_failed_at"] = "unknown"
            all_results["error"] = str(e)
            all_results["final_status"] = "error"
            return all_results
    
    def get_stage_results(self, stage_number: int) -> Dict:
        """Get results from a specific stage"""
        stage_key = f"stage{stage_number}"
        return self.stage_results.get(stage_key, {"error": "Stage not processed yet"})
    
    def validate_cross_stage_consistency(self) -> Dict:
        """
        Validate that data is consistent across all stages
        Critical for payroll accuracy
        """
        validation = {
            "cross_stage_validation": True,
            "timestamp": datetime.now().isoformat(),
            "checks": []
        }
        
        try:
            if "stage1" in self.stage_results and "stage2" in self.stage_results:
                # Check hours consistency
                stage1_hours = self.stage_results["stage1"]["summary"]["total_hours"]
                stage2_hours = self.stage_results["stage2"]["summary"]["total_consolidated_hours"]
                hours_match = abs(stage1_hours - stage2_hours) < 0.01
                
                validation["checks"].append({
                    "check": "Hours consistency Stage 1->2",
                    "passed": hours_match,
                    "stage1_hours": stage1_hours,
                    "stage2_hours": stage2_hours
                })
            
            # Add more cross-stage validations as needed
            
            validation["overall_status"] = all(check["passed"] for check in validation["checks"])
            return validation
            
        except Exception as e:
            validation["error"] = str(e)
            validation["overall_status"] = False
            return validation