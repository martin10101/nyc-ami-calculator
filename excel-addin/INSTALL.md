# AMI Optix Excel Add-in Installation Guide

This guide explains how to install the AMI Optix add-in so it's available in **every** Excel workbook.

## Prerequisites

- Microsoft Excel 2016 or later (Windows)
- Internet connection (to reach the API)
- Trust Center settings configured (see Step 1)

---

## Step 1: Enable Macros in Excel

Before installing the add-in, ensure Excel allows macros:

1. Open Excel
2. Go to **File** > **Options** > **Trust Center**
3. Click **Trust Center Settings**
4. Select **Macro Settings**
5. Choose **"Enable all macros"** or **"Disable all macros with notification"**
6. Check **"Trust access to the VBA project object model"**
7. Click **OK** twice

---

## Step 2: Create the Add-in File

### Option A: Quick Setup (Import Modules)

1. Open Excel with a new blank workbook
2. Press **Alt + F11** to open VBA Editor
3. In VBA Editor, go to **File** > **Import File**
4. Import each `.bas` file from the `src/` folder:
   - `AMI_Optix_Main.bas`
   - `AMI_Optix_API.bas`
   - `AMI_Optix_DataReader.bas`
   - `AMI_Optix_ResultsWriter.bas`
5. Create the UserForm:
   - Right-click on your project > **Insert** > **UserForm**
   - Name it `frmUtilities`
   - Add controls as described in `forms/frmUtilities_DESIGN.txt`
   - Copy code from `forms/frmUtilities.frm` into the form's code window
6. Save as Add-in:
   - Go to **File** > **Save As**
   - Change "Save as type" to **Excel Add-in (.xlam)**
   - Name it `AMI_Optix.xlam`
   - Save to default Add-ins folder (usually auto-suggested)

### Option B: With Custom Ribbon (Advanced)

For the custom ribbon tab, you need to:

1. Complete Option A steps above
2. Save as `.xlam`
3. Close Excel
4. Use a tool like [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor):
   - Open `AMI_Optix.xlam` in the editor
   - Insert contents of `customUI/customUI.xml`
   - Save and close

---

## Step 3: Install the Add-in

1. Open Excel
2. Go to **File** > **Options** > **Add-ins**
3. At the bottom, select **Excel Add-ins** from the dropdown
4. Click **Go...**
5. Click **Browse...**
6. Navigate to your `AMI_Optix.xlam` file and select it
7. Check the box next to **AMI_Optix** in the list
8. Click **OK**

The add-in is now installed and will load automatically every time Excel starts.

---

## Step 4: Verify Installation

1. Open any Excel workbook
2. Look for the **AMI Optix** tab in the ribbon (if custom ribbon was added)
   - OR access via **Developer** tab > **Macros** > `RunOptimization`
3. Click **Optimize** to test

If you don't see the ribbon tab, you can run macros directly:
- Press **Alt + F8**
- Select `RunOptimization` and click **Run**

---

## Using the Add-in

### Configure API Key (Required - One-time setup)

Before using the add-in, you must configure your API key:

1. Click **API Key** button in the AMI Optix ribbon
2. Enter your API key (provided by your administrator)
3. Click **OK**

The key is stored securely in Windows Registry and persists across sessions.

**See [API_KEY_SETUP.md](API_KEY_SETUP.md) for detailed instructions.**

### Configure Utilities (One-time setup)

1. Click **Utilities** button (or run `frmUtilities.Show` macro)
2. Select payment responsibility for each utility:
   - **Tenant Pays**: Tenant is responsible
   - **N/A or owner pays**: Landlord covers it
3. Click **Save**

Settings are remembered for future sessions.

### Run Optimization

1. Open your workbook with unit data
2. Make sure it has columns for:
   - Unit ID (APT, UNIT, etc.)
   - Bedrooms (BED, BEDS, etc.)
   - Net SF (NET SF, SQFT, etc.)
   - AMI (for writing results)
3. Click **Optimize** button
4. Wait for results (may take 30-60 seconds on first run)
5. Results:
   - Best scenario AMI values written to your data
   - All scenarios shown on new "AMI Scenarios" sheet

---

## Troubleshooting

### "Macros are disabled"
- Enable macros in Trust Center (Step 1)

### "Cannot connect to server"
- Check internet connection
- Server may be starting up (cold start) - wait 30 seconds and retry
- Use web dashboard as backup: https://nyc-ami-calculator.onrender.com

### "No unit data found"
- Ensure your workbook has recognizable column headers
- Headers must include: Unit ID, Bedrooms, Net SF

### Add-in doesn't load on startup
- Re-check the add-in is enabled in **File** > **Options** > **Add-ins**

---

## Updating the Add-in

To update to a new version:

1. Close all Excel windows
2. Navigate to your Add-ins folder
3. Replace `AMI_Optix.xlam` with the new version
4. Reopen Excel

---

## Uninstalling

1. Go to **File** > **Options** > **Add-ins**
2. Select **Excel Add-ins** and click **Go...**
3. Uncheck **AMI_Optix**
4. Click **OK**
5. (Optional) Delete the `.xlam` file from the Add-ins folder

---

## Support

- **Web Dashboard Backup**: https://nyc-ami-calculator.onrender.com
- **Issues**: Contact your administrator

---

*AMI Optix v1.0 - NYC Affordable Housing Optimizer*
