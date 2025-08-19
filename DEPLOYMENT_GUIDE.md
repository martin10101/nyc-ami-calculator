# NYC AMI Allocator - Deployment Guide

## Quick Start

### 1. Local Development
```bash
# Install dependencies
pip install -r requirements.txt

# Run the application
uvicorn app:app --host 0.0.0.0 --port 8000

# Access at http://localhost:8000
```

### 2. Railway Deployment
The application is pre-configured for Railway deployment:

1. **Push to Git Repository**
   ```bash
   git init
   git add .
   git commit -m "Initial commit - NYC AMI Allocator"
   ```

2. **Deploy to Railway**
   - Connect your Git repository to Railway
   - Railway will automatically detect the configuration files
   - The app will build and deploy automatically

3. **Environment Variables** (Optional)
   - `MILP_TIMELIMIT_SEC`: MILP solver timeout (default: 10 seconds)
   - `AMI_DB`: Database path for logging (default: /data/ami_log.sqlite)

### 3. Manual Server Deployment
```bash
# Using Gunicorn (production)
gunicorn -k uvicorn.workers.UvicornWorker app:app --bind 0.0.0.0:8000 --timeout 180

# Using Uvicorn (development)
uvicorn app:app --host 0.0.0.0 --port 8000
```

## File Upload Requirements

### Supported Formats
- Excel files (.xlsx, .xlsm, .xls, .xlsb)
- CSV files (.csv)
- Word documents (.docx) with tables

### Required Columns
The application automatically detects columns with these names (case-insensitive):
- **NET SF**: Unit square footage (required)
- **AMI**: Affordable housing selection flag (required)
- **FLOOR**: Floor number (recommended)
- **APT**: Apartment/unit number (recommended)
- **BED**: Number of bedrooms (recommended)

### AMI Selection Format
Units are selected for affordable housing if the AMI column contains:
- Numbers: 0.6, 0.7, 60%, 70%, etc.
- Text: "Yes", "Y", "✓", "✔", "X", etc.
- Any non-blank value

## Configuration Options

### Preview Settings
- **Fast Preview**: Uses heuristic algorithm (faster)
- **Full Optimization**: Uses MILP solver (more accurate)

### Constraints
- **Family Requirement**: Require ≥1 family unit (2BR+) at 40% AMI
- **Floor Limits**: Maximum 40% AMI units per floor
- **Top Floor Exemption**: Exempt top K floors from 40% AMI
- **Scenario Count**: Number of scenarios to generate (1-3)

### Compliance Rules
- **40% Share**: 20-21% of affordable SF must be at 40% AMI
- **Weighted Average**: Overall average must be 59-60% AMI
- **Band Limit**: Maximum 3 AMI bands per scenario
- **Allowed Bands**: 40%, 60%, 70%, 80%, 90%, 100% AMI

## Output Formats

### Preview Mode
- JSON response with scenario breakdown
- Shows unit distribution by AMI band
- Displays compliance metrics
- Identifies best scenario

### Excel Download
- **Master Sheet**: All units with AMI assignments
- **Breakdown Sheets**: Affordable units by scenario
- **Summary Sheet**: Scenario metrics and compliance

### Master Results
- Historical log of all runs
- Exportable Excel format
- Tracks performance over time

## Troubleshooting

### Common Issues
1. **No units selected**: Check AMI column format
2. **Compliance violations**: Adjust unit selection or constraints
3. **Slow performance**: Use fast preview mode
4. **Excel errors**: Check file format and column names

### Error Messages
- "No units selected for AMI allocation": AMI column is empty or incorrectly formatted
- "No valid scenarios could be generated": Constraints are too restrictive
- "Unsupported file type": Use supported file formats

### Performance Tips
- Use fast preview for initial testing
- Enable MILP optimization for final results
- Limit scenario count for faster processing
- Ensure clean data with proper column headers

## API Endpoints

### GET /
Main application interface

### POST /preview
Quick scenario preview
- Returns JSON with scenario breakdown
- Uses heuristic or MILP based on settings

### POST /allocate
Full Excel report generation
- Always uses MILP optimization
- Returns downloadable Excel file

### GET /export_master
Download historical results log
- Returns Excel file with all previous runs

### GET /health
Health check endpoint
- Returns {"ok": true} if application is running

## Support

For technical issues or questions:
1. Check the REBUILD_SUMMARY.md for detailed information
2. Review error messages in the application interface
3. Test with the provided sample data file
4. Verify compliance requirements match your project needs

The application is designed to be robust and user-friendly while maintaining strict compliance with NYC affordable housing regulations.

