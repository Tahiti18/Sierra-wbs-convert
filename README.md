# Sierra Payroll Automation System

Automated system to convert Sierra payroll spreadsheets to WBS payroll format with accurate California overtime calculations.

## Features

- **Automated Conversion**: Converts Sierra Excel files to WBS format
- **California Overtime Rules**: Applies proper overtime calculations (1.5x for 8-12 hours, 2x for >12 hours)
- **Employee Management**: Maintains employee order using gold master list
- **File Validation**: Validates Sierra files before processing
- **Professional UI**: Modern React-based interface with drag-and-drop upload

## Deployment

### Option 1: All-in-One Deployment (Recommended for Testing)
1. Deploy the entire app to Railway - includes both backend and frontend
2. Access the app at your Railway URL

### Option 2: Separate Deployment (Recommended for Production)

#### Backend (Railway)
1. Push this repository to GitHub
2. Connect Railway to your GitHub repository
3. Railway will automatically deploy the Flask backend
4. Note your Railway backend URL (e.g., `https://your-project.up.railway.app`)

#### Frontend (Netlify)
1. Create a new site on Netlify
2. Connect to the same GitHub repository
3. Set build configuration:
   - **Build command**: (leave empty)
   - **Publish directory**: `netlify-frontend`
4. Set environment variables:
   - **SIERRA_BACKEND_URL**: Your Railway backend URL
5. Deploy the site
6. The frontend will be accessible at your Netlify URL

### Quick Start Deployment

1. **Fork/Clone this repository**
2. **Deploy Backend to Railway**:
   - Connect Railway to your GitHub repo
   - Deploy automatically (Railway detects Flask app)
   - Copy the Railway URL
3. **Deploy Frontend to Netlify**:
   - Create new Netlify site from same GitHub repo
   - Set publish directory to `netlify-frontend`
   - Add environment variable `SIERRA_BACKEND_URL` with Railway URL
   - Deploy
4. **Configure Frontend**:
   - Visit your Netlify URL
   - Use the Config button to set backend URL if needed
   - Test connection and upload Sierra files

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
│   ├── main.py              # Flask backend application
│   └── static/
│       └── index.html       # Integrated frontend (Railway deployment)
├── netlify-frontend/
│   ├── index.html           # Standalone frontend (Netlify deployment)
│   ├── netlify.toml         # Netlify configuration
│   └── README.md           # Netlify deployment instructions
├── data/
│   └── gold_master_order.txt # Employee order list
├── improved_converter.py    # Core conversion logic
├── requirements.txt         # Python dependencies
├── Procfile                # Railway deployment config
├── railway.json            # Railway configuration
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

- **Backend**: Flask with CORS enabled for cross-domain requests
- **Frontend**: React SPA with Tailwind CSS and configurable API endpoints
- **File Processing**: pandas, openpyxl for Excel file manipulation
- **Deployment Options**:
  - **All-in-One**: Railway (backend + integrated frontend)
  - **Separate**: Railway (backend) + Netlify (standalone frontend)
- **CORS Configuration**: Enabled for frontend-backend communication
- **File Upload**: Supports .xlsx/.xls files up to 16MB
- **California Overtime**: Compliant with CA labor laws (8hr/12hr rules)

## Usage Instructions

1. **Access the Application**
   - Railway deployment: Visit your Railway URL
   - Netlify deployment: Visit your Netlify URL

2. **Upload Sierra File**
   - Click upload area or drag & drop Excel file
   - File is automatically validated
   - Shows employee count and total hours

3. **Process Payroll**
   - Click "Process Payroll" button
   - Backend applies CA overtime calculations
   - Generates WBS-formatted Excel file

4. **Download Results**
   - Download converted WBS payroll file
   - Ready for submission to payroll company

## Backend Configuration

The frontend can connect to different backend URLs:
- **Automatic**: Uses environment variables or defaults
- **Manual**: Configure through the UI settings
- **Stored**: Backend URL is saved in browser localStorage

## Troubleshooting

### Connection Issues
- Check that backend is deployed and accessible
- Verify backend URL in frontend configuration
- Check browser console for CORS errors
- Use manual configuration option in UI

### File Processing Issues
- Ensure Sierra Excel file has correct format
- Check file size is under 16MB limit
- Verify employee names and time data format
- Check Railway backend logs for processing errors

### Deployment Issues
- **Railway**: Check build logs and environment variables
- **Netlify**: Verify publish directory and environment variables
- **CORS**: Backend already configured for cross-domain requests

