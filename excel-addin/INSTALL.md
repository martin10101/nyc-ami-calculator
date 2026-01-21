# AMI Optix Excel Add-in Installation Guide

This guide explains how to build and install the AMI Optix add-in so it's available in **every** Excel workbook.

## Prerequisites

- Microsoft Excel 2016+ (Windows)
- Internet access to reach the API (Render)
- Excel Trust Center settings configured (Step 1)

---

## Step 1: Enable macros in Excel

1. Open Excel
2. Go to **File** > **Options** > **Trust Center**
3. Click **Trust Center Settings**
4. Select **Macro Settings**
5. Choose **Enable all macros** (or **Disable all macros with notification**)
6. Check **Trust access to the VBA project object model**
7. Click **OK** twice

---

## Step 2: Build the add-in file (`AMI_Optix.xlam`)

### Option A: Import VBA modules (recommended)

1. Open Excel (new blank workbook is fine)
2. Press **Alt + F11** to open the VBA editor
3. In the Project pane, find the project you are building into (ideally a new blank workbook)
4. (Strongly recommended) Remove any old `AMI_Optix_*` modules/classes first:
   - Right-click each old `AMI_Optix_*` item > **Remove**
   - Choose **No** when asked to export
5. In the VBA editor, go to **File** > **Import File...**
6. Import every file from `excel-addin/src/`:
   - `AMI_Optix_Main.bas`
   - `AMI_Optix_API.bas`
   - `AMI_Optix_DataReader.bas`
   - `AMI_Optix_ResultsWriter.bas`
   - `AMI_Optix_Ribbon.bas`
   - `AMI_Optix_EventHooks.bas`
   - `AMI_Optix_Setup.bas`
   - `AMI_Optix_Learning.bas` (AI learning + logging)
   - `AMI_Optix_Diagnostic.bas` (optional)
7. Import the class module from `excel-addin/src/`:
   - `AMI_Optix_AppEvents.cls` (must appear under **Class Modules**)
8. Compile:
   - VBA editor > **Debug** > **Compile VBAProject**
9. Save as add-in:
   - Excel > **File** > **Save As**
   - Save as type: **Excel Add-in (*.xlam)**
   - File name: `AMI_Optix.xlam`

### Option B: Add the custom ribbon (OfficeRibbonXEditor)

Excel stores Ribbon XML *inside* the `.xlam`. To add/update it:

1. Finish Option A and save `AMI_Optix.xlam`
2. Close Excel completely
3. Open **OfficeRibbonXEditor**
4. **File** > **Open** > select `AMI_Optix.xlam`
5. In the left tree:
   - If you see `customUI14.xml`: select it and replace its contents
   - If you do **not** see it: right-click > **Insert** > **Office 2010+ Custom UI Part**
6. Copy/paste the contents of `excel-addin/customUI/customUI14.xml` into the editor
7. Click **Save**
8. Close OfficeRibbonXEditor

---

## Step 3: Install the add-in in Excel

1. Open Excel
2. Go to **File** > **Options** > **Add-ins**
3. At the bottom, in **Manage**, select **Excel Add-ins** > click **Go...**
4. Click **Browse...**
5. Select your `AMI_Optix.xlam`
6. Check **AMI_Optix** in the list
7. Click **OK**

---

## Step 4: Verify it works

1. Restart Excel
2. Open a UAP workbook
3. Confirm you see the **AMI Optix** ribbon tab
4. Click **AMI Optix** > **API Settings** and confirm your API key is set
5. Click **Run UAP**
6. Confirm:
   - UAP AMI column updates
   - `AMI Scenarios` sheet is created/updated

---

## AI Learning (optional)

Learning changes only **soft preferences** (premium scoring weights). It does **not** relax compliance constraints.

Open: **AMI Optix** > **Learning**

- **OFF**: baseline behavior (no learning)
- **SHADOW**: runs learned + baseline compare, but **applies baseline** to the sheet (safe testing)
- **ON**: applies learned preferences (can still compare baseline and log diffs)

Logs:
- Default: `%USERPROFILE%\\Documents\\AMI_Optix_Learning`
- You can set a shared folder (e.g. `Z:\\AMI_Optix_Learning`) in Learning Settings
- Use **Open Logs** to open the current profile folder

---

## Troubleshooting

### "Wrong number of arguments" / random VBA errors
You have mixed module versions.

Fix:
1. Remove **all** `AMI_Optix_*` modules/classes from the VBA project
2. Re-import **all** files from `excel-addin/src/` (Step 2A)
3. Compile again (Debug > Compile)

### Ribbon buttons do nothing
Most common cause: Ribbon XML not saved into the `.xlam`.

Fix:
- Repeat Step 2B and save the `.xlam`, then restart Excel.

### "Cannot connect to server"
- The Render service may be cold-starting (wait 30–60s and retry)
- Confirm the API URL in code is correct:
  - `excel-addin/src/AMI_Optix_Main.bas` > `API_BASE_URL`

---

## Updating the add-in

1. Close all Excel windows
2. Replace `AMI_Optix.xlam` with the new version
3. Reopen Excel

