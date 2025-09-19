# Sierra Payroll Automation System

Automated system to convert Sierra payroll spreadsheets to WBS payroll format with accurate California overtime calculations.

## Features

- **Automated Conversion**: Converts Sierra Excel files to WBS format
- **California Overtime Rules**: Applies proper overtime calculations (1.5x for 8-12 hours, 2x for >12 hours)
- **Employee Management**: Maintains employee order using gold master list
- **File Validation**: Validates Sierra files before processing
- **Professional UI**: Modern React-based interface with drag-and-drop upload

## Deployment

### Backend (Railway)
1. Push this repository to GitHub
2. Connect to Railway
3. Deploy automatically

### Frontend (Netlify) - Optional
The frontend is included in the Flask app, but for separate deployment:
1. Extract the `src/static/index.html` file
2. Update API_BASE_URL to point to your Railway backend
3. Deploy to Netlify

## Local Development

```bash
# Install dependencies
source venv/bin/activate
pip install -r requirements.txt

# Run the application
python src/main.py
```

Visit http://localhost:5000

## API Endpoints

- `GET /api/health` - Health check
- `GET /api/employees` - Get employee list
- `POST /api/process-payroll` - Convert Sierra file to WBS
- `POST /api/validate-sierra-file` - Validate Sierra file format

## File Structure

```
sierra-payroll-system/
├── src/
│   ├── main.py              # Flask application
│   └── static/
│       └── index.html       # Frontend interface
├── data/
│   └── gold_master_order.txt # Employee order list
├── improved_converter.py    # Core conversion logic
├── requirements.txt         # Python dependencies
├── Procfile                # Railway deployment config
└── README.md               # This file
```

## Conversion Process

1. **Parse Sierra File**: Extracts employee time data
2. **Apply Overtime Rules**: California daily overtime calculations
3. **Aggregate Weekly**: Sums daily hours to weekly totals
4. **Sort Employees**: Uses gold master order for consistent ordering
5. **Generate WBS**: Creates properly formatted Excel file

## Supported Formats

- **Input**: Sierra payroll Excel files (.xlsx, .xls)
- **Output**: WBS payroll Excel format with all required headers and calculations

## Technical Details

- **Backend**: Flask with CORS enabled
- **Frontend**: React with Tailwind CSS
- **File Processing**: pandas, openpyxl
- **Deployment**: Railway (backend), Netlify (optional frontend)

