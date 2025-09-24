# Sierra Payroll Frontend for Netlify

This is the Netlify-deployable frontend for the Sierra Payroll Automation System.

## Deployment to Netlify

1. **Deploy to Netlify:**
   - Drag and drop the contents of this folder to Netlify
   - Or connect your GitHub repository containing this folder
   - Netlify will automatically detect the `netlify.toml` configuration

2. **Backend API Configuration:**
   - The frontend is configured to try multiple API endpoints in order:
     1. `https://8080-idrlbzy4bg2q93rmh2rr0-6532622b.e2b.dev/api` (Current working backend)
     2. `https://web-production-d09f2.up.railway.app/api`
     3. `https://sierra-payroll-backend-production.up.railway.app/api`
     4. `https://sierra-roofing-backend.railway.app/api`
     5. `https://sierra-backend-production.up.railway.app/api`

## Features

- **Multi-Backend Support**: Automatically tests and connects to the first available backend API
- **File Upload**: Drag and drop Sierra payroll Excel files
- **Real-time Validation**: Validates files before processing
- **Progress Tracking**: Shows upload and processing progress
- **Employee Management**: View employee roster loaded from backend
- **Responsive Design**: Works on desktop and mobile devices

## API Endpoints Used

- `GET /api/health` - Backend health check
- `GET /api/employees` - Retrieve employee list
- `POST /api/validate-sierra-file` - Validate uploaded Sierra files
- `POST /api/process-payroll` - Convert Sierra files to WBS format

## File Flow

1. **Upload Sierra File**: User uploads Excel payroll file from Sierra system
2. **Validation**: Backend validates file structure and data
3. **Processing**: Backend applies California overtime rules and converts to WBS format
4. **Download**: User downloads the processed WBS payroll file

## Backend Requirements

The backend must support:
- CORS headers for cross-origin requests
- File upload endpoints
- Excel file processing with pandas/openpyxl
- Employee order management (gold master list)

## Current Backend Status

Backend is running at: `https://8080-idrlbzy4bg2q93rmh2rr0-6532622b.e2b.dev`

✅ Successfully tested with:
- 66 employees loaded
- 2052.5 total hours processed
- Sierra → WBS conversion working