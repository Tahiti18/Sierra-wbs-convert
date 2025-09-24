# Sierra Payroll System - Deployment Guide

Complete step-by-step guide for deploying the Sierra Payroll Automation System.

## Architecture Overview

```
┌─────────────────┐    HTTP Requests    ┌──────────────────┐
│  Netlify        │ ──────────────────> │  Railway         │
│  Frontend       │                     │  Backend         │
│  (React SPA)    │ <────────────────── │  (Flask API)     │
└─────────────────┘    JSON/File Data   └──────────────────┘
```

## Deployment Options

### Option A: Production Setup (Recommended)
- **Frontend**: Netlify (Fast CDN, automatic deployments)
- **Backend**: Railway (Reliable hosting, easy database integration)
- **Benefits**: Scalable, fast loading, separate deployments

### Option B: Simple Setup
- **All-in-One**: Railway (Backend + integrated frontend)
- **Benefits**: Single deployment, easier setup

---

## Option A: Production Setup (Netlify + Railway)

### Step 1: Prepare Repository

1. **Fork or clone this repository**
   ```bash
   git clone https://github.com/Tahiti18/Sierra-wbs-convert.git
   cd Sierra-wbs-convert
   ```

2. **Push to your GitHub account** (if cloned)
   ```bash
   git remote set-url origin https://github.com/YOUR_USERNAME/sierra-payroll-system.git
   git push -u origin main
   ```

### Step 2: Deploy Backend to Railway

1. **Visit [Railway.app](https://railway.app)**
2. **Create account** or sign in with GitHub
3. **Create new project**:
   - Click "New Project"
   - Select "Deploy from GitHub repo"
   - Choose your repository
4. **Configure deployment**:
   - Railway auto-detects Flask app
   - Uses `Procfile` for deployment commands
   - Installs dependencies from `requirements.txt`
5. **Get your Railway URL**:
   - After deployment, copy the Railway URL
   - Format: `https://your-project-name.up.railway.app`

### Step 3: Deploy Frontend to Netlify

1. **Visit [Netlify.com](https://netlify.com)**
2. **Create account** or sign in with GitHub
3. **Create new site**:
   - Click "New site from Git"
   - Choose GitHub and select your repository
4. **Configure build settings**:
   ```
   Build command: (leave empty)
   Publish directory: netlify-frontend
   ```
5. **Add environment variables**:
   - Go to Site settings → Environment variables
   - Add: `SIERRA_BACKEND_URL` = Your Railway URL
   - Example: `https://your-project-name.up.railway.app`
6. **Deploy site**:
   - Click "Deploy site"
   - Netlify builds and deploys automatically

### Step 4: Test the System

1. **Visit your Netlify URL**
2. **Check backend connection**:
   - Should show "Connected" status
   - If offline, use Config button to set backend URL
3. **Test file processing**:
   - Upload a Sierra payroll Excel file
   - Verify validation shows employee count
   - Process the file and download WBS format

---

## Option B: Simple Setup (Railway Only)

### Step 1: Deploy to Railway

1. **Visit [Railway.app](https://railway.app)**
2. **Create new project from GitHub**
3. **Railway automatically**:
   - Detects Flask application
   - Installs dependencies
   - Serves both backend API and frontend
4. **Access the application**:
   - Visit your Railway URL
   - Complete app available at one URL

---

## Configuration Options

### Frontend Configuration

The Netlify frontend supports multiple configuration methods:

1. **Environment Variables** (Recommended for production):
   ```
   SIERRA_BACKEND_URL=https://your-project-name.up.railway.app
   ```

2. **Manual Configuration** (Via UI):
   - Click "Config" button in app header
   - Enter backend URL
   - Click "Update" to test connection

3. **Code Configuration** (For custom deployments):
   ```javascript
   // In netlify-frontend/index.html, line ~45
   backendURL = 'https://your-custom-backend-url.com';
   ```

### Backend Configuration

The Railway backend includes:
- **CORS enabled** for cross-domain requests
- **File upload** support (16MB limit)
- **Employee management** with gold master order
- **Health check** endpoint for monitoring

---

## File Processing Workflow

```
1. User uploads Sierra Excel file
   ↓
2. Frontend validates file format
   ↓
3. File sent to backend via API
   ↓
4. Backend processes with CA overtime rules
   ↓
5. WBS Excel file generated
   ↓
6. User downloads converted file
```

---

## Monitoring and Maintenance

### Railway Backend Monitoring
- Check Railway dashboard for logs
- Monitor memory/CPU usage
- Set up alerts for downtime

### Netlify Frontend Monitoring
- Check build logs for deployment issues
- Monitor Core Web Vitals
- Set up form notifications

### Usage Analytics
- Railway provides built-in metrics
- Netlify provides traffic analytics
- Consider adding Google Analytics

---

## Troubleshooting

### Common Issues

#### 1. "Backend Offline" Error
**Symptoms**: Red indicator, cannot process files
**Solutions**:
- Check Railway deployment status
- Verify backend URL in Netlify environment variables
- Use manual configuration in frontend UI
- Check Railway logs for errors

#### 2. CORS Errors
**Symptoms**: Console errors about blocked requests
**Solutions**:
- Backend already configured for CORS
- Check that frontend uses correct backend URL
- Verify no trailing slashes in URLs

#### 3. File Upload Failures
**Symptoms**: Upload succeeds but processing fails
**Solutions**:
- Check file is valid Excel format (.xlsx, .xls)
- Verify file size is under 16MB
- Check Sierra file has correct column structure
- Review Railway logs for processing errors

#### 4. Deployment Failures

**Railway Issues**:
- Check `requirements.txt` for dependency errors
- Verify `Procfile` syntax
- Check Python version compatibility

**Netlify Issues**:
- Verify publish directory is `netlify-frontend`
- Check environment variables are set
- Ensure repository has correct folder structure

### Getting Support

1. **Check logs**:
   - Railway: Dashboard → Deployments → Logs
   - Netlify: Site dashboard → Functions → Logs

2. **Verify configuration**:
   - Environment variables set correctly
   - URLs don't have trailing slashes
   - File permissions are correct

3. **Test components separately**:
   - Access backend health endpoint directly
   - Test frontend with different backend URLs
   - Upload small test files

---

## Security Considerations

### Production Checklist

- [ ] Use HTTPS for all deployments
- [ ] Set proper CORS origins (currently set to * for development)
- [ ] Implement file size limits (currently 16MB)
- [ ] Add request rate limiting
- [ ] Monitor for suspicious file uploads
- [ ] Regular security updates for dependencies

### Environment Variables

Never commit these to version control:
- Database connection strings
- API keys
- Production backend URLs

---

## Performance Optimization

### Frontend (Netlify)
- Static files served via CDN
- Automatic GZIP compression
- Global edge network

### Backend (Railway)
- Consider upgrading plan for higher traffic
- Implement caching for employee data
- Optimize Excel processing for large files

### File Processing
- Current limit: 16MB files
- Processing time: ~2-5 seconds per file
- Memory usage: ~100MB per concurrent request

---

## Next Steps

After successful deployment:

1. **Test with real Sierra files**
2. **Train users on the interface**
3. **Set up monitoring alerts**
4. **Plan for scaling if needed**
5. **Consider adding user authentication**
6. **Implement audit logging**

---

## Support and Updates

### Repository
- **GitHub**: https://github.com/Tahiti18/Sierra-wbs-convert
- **Issues**: Use GitHub Issues for bug reports
- **Updates**: Pull latest changes and redeploy

### Deployment URLs
- **Railway Backend**: Your Railway project URL
- **Netlify Frontend**: Your Netlify site URL

Remember to update these URLs in your documentation and user training materials.