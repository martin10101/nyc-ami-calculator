# Step-by-step: Update the AMI Optix VBA + Ribbon (Windows Excel)

This is written to be *very* simple and safe.

## A) Get the latest code (GitHub)

**If you use Git (recommended):**
1. Open PowerShell in the repo folder.
2. Run:
   - `git fetch`
   - `git checkout feature/results-overhaul-2026-01-28`
   - `git pull`

**If you don’t use Git:**
1. On GitHub, switch to branch: `feature/results-overhaul-2026-01-28`
2. Download the ZIP of that branch.
3. Unzip it somewhere on your computer.

---

## B) Build a fresh add-in file (`AMI_Optix.xlam`)

1. **Close Excel completely** (all Excel windows).
2. Open Excel and create a **new blank workbook**.
3. Press **Alt + F11** (opens the VBA editor).
4. In the left “Project” panel, click your new workbook.
5. Remove old modules if you see any:
   - Delete anything named `AMI_Optix_*` (Modules + Class Modules).
6. Import the code files:
   - VBA editor menu: **File > Import File…**
   - Import **every** file inside: `excel-addin/src/`
     - `AMI_Optix_Main.bas`
     - `AMI_Optix_API.bas`
     - `AMI_Optix_DataReader.bas`
     - `AMI_Optix_ResultsWriter.bas`
     - `AMI_Optix_Ribbon.bas`
     - `AMI_Optix_EventHooks.bas`
     - `AMI_Optix_Learning.bas`
     - `AMI_Optix_Diagnostic.bas` (optional)
     - `AMI_Optix_Setup.bas` (optional)
   - Also import the **class module**:
     - `AMI_Optix_AppEvents.cls` (this must show up under “Class Modules”)
7. Compile (this catches missing files / mix-and-match versions):
   - VBA editor menu: **Debug > Compile VBAProject**
8. Save the add-in:
   - Excel: **File > Save As**
   - “Save as type”: **Excel Add-in (*.xlam)**
   - Name it: `AMI_Optix.xlam`

---

## C) Put the Ribbon into the add-in (OfficeRibbonXEditor)

Excel stores the Ribbon XML *inside* the `.xlam`.

1. **Close Excel completely** again.
2. Open **OfficeRibbonXEditor**.
3. **File > Open** and select your `AMI_Optix.xlam`.
4. In the left tree:
   - If you already see `customUI14.xml`, click it.
   - If you do **not** see it: right-click and choose **Insert > Office 2010+ Custom UI Part**.
5. Copy/paste the contents of:
   - `excel-addin/customUI/customUI14.xml`
6. Click **Save**, then close OfficeRibbonXEditor.

---

## D) Install/Enable the add-in in Excel

1. Open Excel.
2. **File > Options > Add-ins**
3. Bottom dropdown “Manage”: choose **Excel Add-ins** > click **Go…**
4. Click **Browse…** and select your `AMI_Optix.xlam`.
5. Make sure **AMI_Optix** is checked, then click **OK**.

---

## E) Quick “did it work?” checklist

1. Restart Excel.
2. Open your UAP or MIH workbook.
3. You should see the **AMI Optix** ribbon tab.
4. Click **Run UAP** or **Run MIH** and confirm:
   - UAP/MIH unit table updates
   - `AMI Scenarios` updates (with a “SCENARIO MANUAL (LIVE SYNC)” block)
5. Test Live Sync:
   - Change an AMI value in **UAP** (or **RentRoll** for MIH) → the “Scenario Manual” table updates.
   - Change an AMI value in the “Scenario Manual” table → the UAP/RentRoll AMIs update.
6. Test logging:
   - Click **Record Choice** → type a reason (or leave it blank) → it appends to the shared log file.

---

## F) MIH: where the file tells the solver what to do

The MIH solver reads these workbook cells:
- `Prog!K4` = MIH option (`Option 1` or `Option 4`)
- `MIH!J21` = residential square feet (denominator for share thresholds)
- `Prog!I4` = max AMI band factor (example: `1.35` = 135%)

