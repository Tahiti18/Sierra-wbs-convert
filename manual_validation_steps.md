# MANUAL VALIDATION STEPS FOR SIERRA TO WBS CONVERSION

Follow these exact steps to test and verify the conversion accuracy yourself.

## Step 1: Test the Live System

### A. Upload to Netlify Frontend
1. Go to: https://sierrapayrollapp.netlify.app/
2. Upload the Sierra file: `Sierra_Payroll_Sample_1.xlsx`
3. Click "Process Payroll"
4. Download the WBS output file

### B. Upload to Backend API Directly  
```bash
# Test the Railway backend (current production)
curl -X POST -F "file=@Sierra_Payroll_Sample_1.xlsx" \
  https://web-production-d09f2.up.railway.app/process-payroll \
  -o "Railway_WBS_Output.xlsx"

# Test our improved backend (if available)
curl -X POST -F "file=@Sierra_Payroll_Sample_1.xlsx" \
  https://8081-idrlbzy4bg2q93rmh2rr0-6532622b.e2b.dev/api/process-payroll \
  -o "Improved_WBS_Output.xlsx"
```

## Step 2: Manual Calculation Verification

### Pick Any Employee - Let's Use Dianne Robleza:

1. **Find in Sierra File:**
   - Open `Sierra_Payroll_Sample_1.xlsx`
   - Search for "Dianne" 
   - You should find: 4.0 hours at $28.00/hour
   - Manual calculation: 4 × $28 = $112

2. **Find in Expected WBS:**
   - Open `WBS_Payroll_Sample.xlsx`
   - Look for "Robleza, Dianne" (around row 8)
   - Expected: Rate=$28.00, Hours=4, Total=$112

3. **Find in Your Output:**
   - Open your downloaded WBS file
   - Search for "Dianne" or "Robleza"
   - Compare: Rate, Hours, and Total amounts

### Test California Overtime Rules:

1. **Find Employee with >8 Hours:**
   - In Sierra file, look for any employee with >8 hours in a day
   - Example: Alejandro Gonzalez has 9.0 hours at $43.00/hour

2. **Manual Overtime Calculation:**
   - Regular: 8 hours × $43.00 = $344.00
   - Overtime: 1 hour × $43.00 × 1.5 = $64.50
   - Total: $344.00 + $64.50 = $408.50

3. **Verify in Output:**
   - Check if your WBS output shows correct overtime calculations

## Step 3: Format Structure Verification

### Check WBS File Structure:
```bash
# Count columns in expected vs actual
python3 -c "
import pandas as pd
expected = pd.read_excel('WBS_Payroll_Sample.xlsx')
actual = pd.read_excel('YOUR_WBS_OUTPUT.xlsx')
print(f'Expected columns: {len(expected.columns)}')
print(f'Actual columns: {len(actual.columns)}')
print(f'Structure match: {len(expected.columns) == len(actual.columns)}')"
```

### Check Employee Data Positioning:
- Employee data should start at row 8 (index 7)
- Column structure should be:
  - Col 0: Employee ID (e.g., "0000662082")
  - Col 1: SSN (e.g., "626946016") 
  - Col 2: Name (e.g., "Robleza, Dianne")
  - Col 5: Pay Rate (e.g., "28.00")
  - Col 6: Department (e.g., "ADMIN")
  - Col 7: Hours (e.g., "4")
  - Col 27: Total Amount (e.g., "112")

## Step 4: Total Validation

### Calculate Grand Totals:
```bash
# Sierra input total
python3 -c "
import pandas as pd
df = pd.read_excel('Sierra_Payroll_Sample_1.xlsx')
valid_rows = df[(df['Hours'] > 0) & (df['Total'].notna())]
print(f'Sierra Total: ${valid_rows[\"Total\"].sum():,.2f}')
print(f'Sierra Hours: {valid_rows[\"Hours\"].sum()}')"

# Your WBS output total
python3 -c "
import pandas as pd
df = pd.read_excel('YOUR_WBS_OUTPUT.xlsx')
total = 0
for i in range(7, len(df)):
    try:
        val = df.iloc[i, 27]
        if pd.notna(val): total += float(val)
    except: pass
print(f'WBS Total: ${total:,.2f}')"
```

## Step 5: Specific Test Cases

### Test Case 1: Simple Employee (No Overtime)
- **Employee:** Dianne Robleza
- **Expected:** 4 hours × $28.00 = $112.00
- **Verify:** Your output matches exactly

### Test Case 2: Overtime Employee  
- **Employee:** Find one with >8 hours/day
- **Calculate:** 8 hrs regular + (extra hrs × 1.5) + (>12 hrs × 2.0)
- **Verify:** Your output applies CA overtime rules

### Test Case 3: Multiple Days Same Employee
- **Find:** Employee with entries on multiple days
- **Calculate:** Sum all daily totals (with daily overtime applied)
- **Verify:** Weekly aggregation is correct

## Expected Results

### ✅ PASS Criteria:
1. **Structure:** 28 columns exactly
2. **Employee Data:** Starts at row 8, proper column positioning
3. **Calculations:** Individual employee totals match manual calculations
4. **Overtime:** CA rules applied correctly (1.5x for 8-12hrs, 2x for >12hrs)
5. **Format:** Names in "Last, First" format
6. **Departments:** Correct ADMIN/GUTTR assignments

### ❌ FAIL Indicators:
1. Wrong number of columns
2. Employee data in wrong positions
3. Calculation errors (totals don't match manual calculations)
4. Missing employees
5. Incorrect overtime calculations
6. Wrong file format structure

## Troubleshooting

If tests fail:
1. Check which specific employee calculations are wrong
2. Verify the Sierra input data was parsed correctly
3. Check if CA overtime rules are being applied
4. Ensure employee sorting matches gold master order
5. Validate SSN and department assignments

Run these tests and report which specific validations pass or fail!
