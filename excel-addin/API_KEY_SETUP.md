# API Key Setup Guide

The AMI Optix Excel Add-in requires an API key for authentication with the optimization server.

## Overview

```
┌─────────────────────────────────────────────────────────────────┐
│                        SECURITY FLOW                            │
│                                                                 │
│  Excel Add-in                              Render API Server    │
│  ┌───────────────┐                        ┌───────────────┐    │
│  │ API Key       │ ──── X-API-Key ────►   │ Validates     │    │
│  │ (Registry)    │      Header            │ API Key       │    │
│  └───────────────┘                        └───────────────┘    │
│                                                                 │
│  Only requests with valid API key are processed                │
└─────────────────────────────────────────────────────────────────┘
```

---

## Step 1: Generate an API Key

Choose a strong, random API key. You can use:

### Option A: Online Generator
Use a password generator to create a 32+ character random string.

### Option B: Command Line (PowerShell)
```powershell
[System.Guid]::NewGuid().ToString() + [System.Guid]::NewGuid().ToString()
```

### Option C: Python
```python
import secrets
print(secrets.token_urlsafe(32))
```

**Example key**: `xK9mP2nQ8rT5vW3yB7cD4fG6hJ1kL0pR`

**Important**:
- Keep this key SECRET
- Don't share it publicly
- Use the SAME key in both Render and Excel

---

## Step 2: Configure API Key on Render

1. Go to your Render dashboard: https://dashboard.render.com
2. Select your **nyc-ami-calculator** service
3. Click **Environment** in the left menu
4. Add a new environment variable:
   - **Key**: `AMI_OPTIX_API_KEY`
   - **Value**: Your generated API key
5. Click **Save Changes**
6. The service will automatically redeploy

```
┌─────────────────────────────────────────────────────────────────┐
│  Render Dashboard > Environment Variables                       │
│                                                                 │
│  Key                    Value                                   │
│  ─────────────────────  ─────────────────────────────────────  │
│  AMI_OPTIX_API_KEY      xK9mP2nQ8rT5vW3yB7cD4fG6hJ1kL0pR       │
│                                                                 │
└─────────────────────────────────────────────────────────────────┘
```

---

## Step 3: Configure API Key in Excel Add-in

### First Time Setup

1. Open Excel with the AMI Optix add-in installed
2. Click **AMI Optix** tab in the ribbon
3. Click **API Key** button (or **Settings**)
4. Enter your API key in the dialog
5. Click **OK**

The key is stored securely in the Windows Registry and persists across Excel sessions.

### Updating the Key

1. Click **API Key** button
2. Enter the new key
3. Click **OK**

### Clearing the Key

1. Click **API Key** button
2. Type `CLEAR`
3. Click **OK**

---

## Step 4: Verify Setup

1. Open any Excel workbook with unit data
2. Click **Optimize** in the AMI Optix ribbon
3. If the API key is correct, optimization will run
4. If you see "Invalid API key" error, double-check both:
   - The key in Render environment variables
   - The key entered in Excel

---

## Where is the Key Stored?

### Server (Render)
- Environment variable: `AMI_OPTIX_API_KEY`
- Not visible in logs or code
- Only accessible to the application

### Client (Excel)
- Windows Registry: `HKEY_CURRENT_USER\Software\VB and VBA Program Settings\AMI_Optix\Settings\APIKey`
- Stored per Windows user
- Persists across Excel sessions
- Not stored in workbooks (so sharing workbooks doesn't expose the key)

---

## Security Notes

1. **Don't share the API key** - Treat it like a password
2. **One key per organization** - All users in the same company can use the same key
3. **Rotate periodically** - Change the key every 6-12 months
4. **If compromised** - Change the key immediately in both Render and Excel

---

## Troubleshooting

### "Invalid API key" Error
- Check the key matches exactly (no extra spaces)
- Verify Render service redeployed after adding the variable
- Try clearing and re-entering the key in Excel

### "API key not configured" Message
- Click **API Key** button and enter your key
- The add-in requires a key before optimization can run

### Web Dashboard Still Works Without Key
- The web dashboard (`/api/analyze`) doesn't require API key
- Only the new `/api/optimize` endpoint (used by Excel) requires authentication
- This is intentional - web users upload files, Excel sends raw data

---

## For Developers

### Testing Without API Key

During development, if `AMI_OPTIX_API_KEY` is not set (empty), the API allows all requests. This is useful for local testing.

```bash
# Local development (no auth)
flask run

# Production (with auth)
AMI_OPTIX_API_KEY=your-key-here flask run
```

### API Authentication Flow

```python
# Request from Excel
POST /api/optimize HTTP/1.1
Host: nyc-ami-calculator.onrender.com
Content-Type: application/json
X-API-Key: xK9mP2nQ8rT5vW3yB7cD4fG6hJ1kL0pR

{"units": [...], "utilities": {...}}
```

```python
# Server validation (app.py)
def _validate_api_key():
    if not API_KEY:  # No key configured = dev mode
        return None
    if request.headers.get('X-API-Key') != API_KEY:
        return jsonify({"error": "Invalid or missing API key"}), 401
    return None
```

---

*Last updated: January 2025*
