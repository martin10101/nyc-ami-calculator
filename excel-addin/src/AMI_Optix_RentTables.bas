Attribute VB_Name = "AMI_Optix_RentTables"
'===============================================================================
' AMI OPTIX - Rent Tables Cache (Fix-06c)
'
' Normalizes the selected year rent workbook into per-user CSV cache tables for:
'  - gross rent limits
'  - utility allowances
'
' Shared authoritative source:
'   Z:\AMI_Optix\RentRollYears\<YEAR>\
'
' Per-user fallback source:
'   %APPDATA%\AMI_Optix\RentRollYears\<YEAR>\
'
' Per-user normalized cache (inspectable):
'   %APPDATA%\AMI_Optix\RentTablesCache\<YEAR>\
'===============================================================================
Option Explicit

Private Const RENTROLL_SHARED_BASE As String = "Z:\AMI_Optix\RentRollYears"
Private Const RENTROLL_LOCAL_SUBPATH As String = "\AMI_Optix\RentRollYears"
Private Const RENTROLL_LOCAL_FILENAME_PREFIX As String = "RentCalculator_"
Private Const RENTROLL_LOCAL_FILENAME_SUFFIX As String = ".xlsx"

Private Const RENT_TABLES_CACHE_SUBPATH As String = "\AMI_Optix\RentTablesCache"
Private Const RENT_LIMITS_CSV_NAME As String = "rent_limits.csv"
Private Const UTILITY_ALLOWANCES_CSV_NAME As String = "utility_allowances.csv"
Private Const CACHE_META_NAME As String = "cache_meta.txt"

Private Const RENT_TABLES_ERR_CACHE As Long = vbObjectError + 620
Private Const RENT_TABLES_ERR_IMPORT As Long = vbObjectError + 621

Public Function ResolveYearWorkbookPath(year As Long, Optional ByRef sourceLabel As String = "") As String
    ' Z: first, then %APPDATA% fallback.
    Dim preferredName As String
    preferredName = RENTROLL_LOCAL_FILENAME_PREFIX & CStr(year) & RENTROLL_LOCAL_FILENAME_SUFFIX

    Dim sharedFolder As String
    sharedFolder = RENTROLL_SHARED_BASE & "\" & CStr(year)

    Dim sharedPreferred As String
    sharedPreferred = sharedFolder & "\" & preferredName
    If Dir$(sharedPreferred) <> "" Then
        sourceLabel = "Z:"
        ResolveYearWorkbookPath = sharedPreferred
        Exit Function
    End If

    Dim sharedAny As String
    sharedAny = FirstWorkbookInFolder(sharedFolder)
    If Trim$(sharedAny) <> "" Then
        sourceLabel = "Z:"
        ResolveYearWorkbookPath = sharedAny
        Exit Function
    End If

    Dim localFolder As String
    localFolder = Environ$("APPDATA") & RENTROLL_LOCAL_SUBPATH & "\" & CStr(year)

    Dim localPreferred As String
    localPreferred = localFolder & "\" & preferredName
    If Dir$(localPreferred) <> "" Then
        sourceLabel = "AppData"
        ResolveYearWorkbookPath = localPreferred
        Exit Function
    End If

    Dim localAny As String
    localAny = FirstWorkbookInFolder(localFolder)
    If Trim$(localAny) <> "" Then
        sourceLabel = "AppData"
        ResolveYearWorkbookPath = localAny
        Exit Function
    End If

    sourceLabel = ""
    ResolveYearWorkbookPath = ""
End Function

Public Function GetYearWorkbookFingerprint(path As String) As String
    On Error GoTo Fail
    If Trim$(path) = "" Then Err.Raise RENT_TABLES_ERR_CACHE, "AMI_Optix_RentTables.GetYearWorkbookFingerprint", "Missing path."

    Dim mtime As String
    mtime = Format$(FileDateTime(path), "yyyy-mm-dd hh:nn:ss")

    Dim sizeBytes As String
    sizeBytes = CStr(FileLen(path))

    GetYearWorkbookFingerprint = "mtime=" & mtime & ";size=" & sizeBytes
    Exit Function

Fail:
    Err.Raise RENT_TABLES_ERR_CACHE, "AMI_Optix_RentTables.GetYearWorkbookFingerprint", _
              "Could not fingerprint year workbook." & vbCrLf & vbCrLf & _
              "Path: " & path & vbCrLf & _
              "Error: " & Err.Description
End Function

Public Function GetRentTablesCacheFolder(year As Long) As String
    GetRentTablesCacheFolder = Environ$("APPDATA") & RENT_TABLES_CACHE_SUBPATH & "\" & CStr(year)
End Function

Public Function EnsureRentTablesCache(year As Long, Optional forceRefresh As Boolean = False, _
                                     Optional ByRef outSourcePath As String = "", _
                                     Optional ByRef outCacheFolder As String = "", _
                                     Optional ByRef outFingerprint As String = "") As Boolean
    ' Ensures %APPDATA% CSV cache exists and matches the authoritative source workbook.
    ' Silent on success (logs diagnostics); callers decide whether to MsgBox.
    On Error GoTo Fail

    EnsureRentTablesCache = False

    Dim prevEnableEvents As Boolean
    Dim prevScreenUpdating As Boolean
    prevEnableEvents = Application.EnableEvents
    prevScreenUpdating = Application.ScreenUpdating

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Dim sourceLabel As String
    sourceLabel = ""

    Dim sourcePath As String
    sourcePath = ResolveYearWorkbookPath(year, sourceLabel)
    If Trim$(sourcePath) = "" Then
        Err.Raise RENT_TABLES_ERR_CACHE, "AMI_Optix_RentTables.EnsureRentTablesCache", _
                  "Rent tables source workbook not found for year " & CStr(year) & "." & vbCrLf & vbCrLf & _
                  "Checked:" & vbCrLf & _
                  " - " & RENTROLL_SHARED_BASE & "\" & CStr(year) & vbCrLf & _
                  " - %APPDATA%" & RENTROLL_LOCAL_SUBPATH & "\" & CStr(year)
    End If

    Dim fingerprint As String
    fingerprint = GetYearWorkbookFingerprint(sourcePath)

    Dim cacheFolder As String
    cacheFolder = GetRentTablesCacheFolder(year)

    Dim baseFolder As String
    baseFolder = Environ$("APPDATA") & RENT_TABLES_CACHE_SUBPATH
    EnsureFolderExistsSafe baseFolder
    EnsureFolderExistsSafe cacheFolder

    Dim rentLimitsCsv As String
    rentLimitsCsv = cacheFolder & "\" & RENT_LIMITS_CSV_NAME

    Dim utilityCsv As String
    utilityCsv = cacheFolder & "\" & UTILITY_ALLOWANCES_CSV_NAME

    Dim metaPath As String
    metaPath = cacheFolder & "\" & CACHE_META_NAME

    Dim stale As Boolean
    stale = forceRefresh

    If Not stale Then
        If Dir$(rentLimitsCsv) = "" Or Dir$(utilityCsv) = "" Or Dir$(metaPath) = "" Then
            stale = True
        ElseIf Not CacheMetaMatches(metaPath, sourcePath, fingerprint) Then
            stale = True
        End If
    End If

    If stale Then
        BuildRentTablesCache year, sourcePath, fingerprint, cacheFolder
    End If

    ' Verify outputs exist after build/reuse.
    If Dir$(rentLimitsCsv) = "" Then
        Err.Raise RENT_TABLES_ERR_CACHE, "AMI_Optix_RentTables.EnsureRentTablesCache", _
                  "Rent tables cache missing after refresh." & vbCrLf & vbCrLf & _
                  "Missing: " & rentLimitsCsv
    End If
    If Dir$(utilityCsv) = "" Then
        Err.Raise RENT_TABLES_ERR_CACHE, "AMI_Optix_RentTables.EnsureRentTablesCache", _
                  "Rent tables cache missing after refresh." & vbCrLf & vbCrLf & _
                  "Missing: " & utilityCsv
    End If

    If Not CacheMetaMatches(metaPath, sourcePath, fingerprint) Then
        Err.Raise RENT_TABLES_ERR_CACHE, "AMI_Optix_RentTables.EnsureRentTablesCache", _
                  "Rent tables cache meta does not match the resolved year workbook." & vbCrLf & vbCrLf & _
                  "Year: " & CStr(year) & vbCrLf & _
                  "Workbook: " & sourcePath & vbCrLf & _
                  "Fingerprint: " & fingerprint & vbCrLf & _
                  "Meta: " & metaPath
    End If

    outSourcePath = sourcePath
    outCacheFolder = cacheFolder
    outFingerprint = fingerprint

    ' Diagnostics (no MsgBox on success).
    On Error Resume Next
    LogRentTablesDiagnostics year, sourcePath, cacheFolder, fingerprint, stale, sourceLabel
    On Error GoTo 0

    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents

    EnsureRentTablesCache = True
    Exit Function

Fail:
    On Error Resume Next
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    On Error GoTo 0
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'-------------------------------------------------------------------------------
' CACHE BUILD
'-------------------------------------------------------------------------------

Private Sub BuildRentTablesCache(year As Long, sourcePath As String, fingerprint As String, cacheFolder As String)
    On Error GoTo Fail

    Dim rentWb As Workbook
    Dim openedHere As Boolean
    openedHere = False

    Set rentWb = FindOpenWorkbookByPath(sourcePath)
    If rentWb Is Nothing Then
        Set rentWb = Workbooks.Open(sourcePath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False, Notify:=False)
        openedHere = True
        On Error Resume Next
        rentWb.Windows(1).Visible = False
        On Error GoTo Fail
    End If

    Dim rentRows As Collection
    Set rentRows = ExtractRentLimits(rentWb, year)

    Dim utilRows As Collection
    Set utilRows = ExtractUtilityAllowances(rentWb, year)

    ValidateRentLimitsCoverage rentRows, year, sourcePath
    ValidateUtilityAllowanceCoverage utilRows, year, sourcePath

    WriteRentLimitsCsv cacheFolder & "\" & RENT_LIMITS_CSV_NAME, rentRows
    WriteUtilityAllowancesCsv cacheFolder & "\" & UTILITY_ALLOWANCES_CSV_NAME, utilRows

    WriteCacheMeta cacheFolder & "\" & CACHE_META_NAME, year, sourcePath, fingerprint

    If openedHere Then
        rentWb.Close SaveChanges:=False
    End If

    Exit Sub

Fail:
    On Error Resume Next
    If openedHere Then
        If Not rentWb Is Nothing Then rentWb.Close SaveChanges:=False
    End If
    On Error GoTo 0
    Err.Raise RENT_TABLES_ERR_IMPORT, "AMI_Optix_RentTables.BuildRentTablesCache", _
              "Failed to build rent tables cache." & vbCrLf & vbCrLf & _
              "Year: " & CStr(year) & vbCrLf & _
              "Workbook: " & sourcePath & vbCrLf & _
              "Fingerprint: " & fingerprint & vbCrLf & _
              "Cache folder: " & cacheFolder & vbCrLf & vbCrLf & _
              "Error: " & Err.Description
End Sub

'-------------------------------------------------------------------------------
' EXTRACTION
'-------------------------------------------------------------------------------

Public Function ExtractRentLimits(rentWb As Workbook, Optional year As Long = 0) As Collection
    ' Returns rows: [Year, Program, Bedrooms, AMI, GrossRent]
    On Error GoTo Fail

    Dim rows As Collection
    Set rows = New Collection

    If TryExtractRentLimitsFromNamedTable(rentWb, year, rows) Then
        Set ExtractRentLimits = rows
        Exit Function
    End If

    ' Fallback: scrape "AMI & Rent" layout (import-only).
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = rentWb.Worksheets("AMI & Rent")
    On Error GoTo Fail
    If ws Is Nothing Then
        Err.Raise RENT_TABLES_ERR_IMPORT, "AMI_Optix_RentTables.ExtractRentLimits", "Sheet 'AMI & Rent' not found."
    End If

    Dim fingerprint As String
    Dim reason As String
    fingerprint = ""
    reason = ""
    If Not ValidateLocalRentWorkbookLayout(ws, fingerprint, reason) Then
        Err.Raise RENT_TABLES_ERR_IMPORT, "AMI_Optix_RentTables.ExtractRentLimits", _
                  "Local rent workbook layout not recognized." & vbCrLf & vbCrLf & _
                  "Year: " & CStr(year) & vbCrLf & _
                  "Workbook: " & rentWb.FullName & vbCrLf & _
                  "Sheet: AMI & Rent" & vbCrLf & _
                  "Reason: " & reason & vbCrLf & _
                  "Fingerprint: " & fingerprint
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).row ' col C
    If lastRow < 1 Then lastRow = 1

    Dim currentAmi As Double
    currentAmi = -1#

    Dim programs As Variant
    programs = Array("UAP", "MIH")

    Dim r As Long
    For r = 1 To Application.Min(lastRow, 5000)
        Dim cVal As Variant
        cVal = ws.Cells(r, 3).Value

        Dim marker As String
        marker = LCase$(Trim$(CStr(ws.Cells(r, 4).Value)))

        If IsNumeric(cVal) And marker = "of ami" Then
            currentAmi = CDbl(cVal)
        ElseIf currentAmi >= 0# Then
            Dim bedLabel As String
            bedLabel = BedroomLabelFromSheetLabel(CStr(cVal))
            If bedLabel <> "" Then
                Dim grossVal As Variant
                grossVal = ws.Cells(r, 7).Value ' col G
                If Not IsNumeric(grossVal) Then
                    Err.Raise RENT_TABLES_ERR_IMPORT, "AMI_Optix_RentTables.ExtractRentLimits", _
                              "Gross rent value is not numeric." & vbCrLf & vbCrLf & _
                              "Row: " & CStr(r) & vbCrLf & _
                              "AMI: " & CanonAmiKey(currentAmi) & vbCrLf & _
                              "Bedrooms: " & bedLabel
                End If

                Dim p As Long
                For p = LBound(programs) To UBound(programs)
                    rows.Add Array(CLng(year), CStr(programs(p)), bedLabel, CanonAmiKey(currentAmi), CDbl(grossVal))
                Next p
            End If
        End If
    Next r

    If rows.Count = 0 Then
        Err.Raise RENT_TABLES_ERR_IMPORT, "AMI_Optix_RentTables.ExtractRentLimits", _
                  "Could not extract any rent limit rows from 'AMI & Rent'."
    End If

    Set ExtractRentLimits = rows
    Exit Function

Fail:
    Err.Raise RENT_TABLES_ERR_IMPORT, "AMI_Optix_RentTables.ExtractRentLimits", _
              "ExtractRentLimits failed: " & Err.Description
End Function

Public Function ExtractUtilityAllowances(rentWb As Workbook, Optional year As Long = 0) As Collection
    ' Returns rows: [Year, UtilityType, UtilityVariant, Bedrooms, Allowance]
    On Error GoTo Fail

    Dim rows As Collection
    Set rows = New Collection

    If TryExtractUtilityAllowancesFromNamedTable(rentWb, year, rows) Then
        Set ExtractUtilityAllowances = rows
        Exit Function
    End If

    ' Fallback: scrape "AMI & Rent" layout (import-only).
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = rentWb.Worksheets("AMI & Rent")
    On Error GoTo Fail
    If ws Is Nothing Then
        Err.Raise RENT_TABLES_ERR_IMPORT, "AMI_Optix_RentTables.ExtractUtilityAllowances", "Sheet 'AMI & Rent' not found."
    End If

    Dim fingerprint As String
    Dim reason As String
    fingerprint = ""
    reason = ""
    If Not ValidateLocalRentWorkbookLayout(ws, fingerprint, reason) Then
        Err.Raise RENT_TABLES_ERR_IMPORT, "AMI_Optix_RentTables.ExtractUtilityAllowances", _
                  "Local rent workbook layout not recognized." & vbCrLf & vbCrLf & _
                  "Year: " & CStr(year) & vbCrLf & _
                  "Workbook: " & rentWb.FullName & vbCrLf & _
                  "Sheet: AMI & Rent" & vbCrLf & _
                  "Reason: " & reason & vbCrLf & _
                  "Fingerprint: " & fingerprint
    End If

    Dim bedLabels As Variant
    bedLabels = Array("studio", "1 BR", "2 BR", "3 BR", "4 BR", "5 BR")

    Dim currentCat As String
    currentCat = ""

    Dim lastCol As Long
    lastCol = ws.Cells(15, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then lastCol = 1

    Dim col As Long
    For col = 1 To Application.Min(lastCol, 200)
        Dim headerVal As String
        headerVal = Trim$(CStr(ws.Cells(15, col).Value))

        Dim headerCat As String
        headerCat = UtilityCategoryFromHeaderValue(headerVal)
        If headerCat <> "" Then currentCat = headerCat

        Dim optionVal As Variant
        optionVal = ws.Cells(16, col).Value
        If Trim$(CStr(optionVal)) = "" Then optionVal = ws.Cells(17, col).Value

        Dim optionLabel As String
        optionLabel = Trim$(CStr(optionVal))
        If optionLabel = "" Then GoTo NextCol
        If LCase$(optionLabel) = "select -->>" Then GoTo NextCol

        Dim optionCat As String
        optionCat = UtilityCategoryFromOptionLabel(optionLabel)
        If optionCat = "" Then optionCat = currentCat
        If optionCat = "" Then GoTo NextCol

        Dim variantCode As String
        variantCode = UtilityVariantCodeFromOptionLabel(optionCat, optionLabel)
        If variantCode = "" Then GoTo NextCol ' ignore N/A/owner-pays option columns

        Dim i As Long
        For i = LBound(bedLabels) To UBound(bedLabels)
            Dim amtVal As Variant
            amtVal = ws.Cells(18 + i, col).Value ' rows 18..23

            Dim amt As Double
            amt = 0#
            If IsNumeric(amtVal) Then
                amt = CDbl(amtVal)
            ElseIf Len(Trim$(CStr(amtVal))) > 0 Then
                On Error Resume Next
                amt = CDbl(amtVal)
                If Err.Number <> 0 Then
                    On Error GoTo Fail
                    Err.Raise RENT_TABLES_ERR_IMPORT, "AMI_Optix_RentTables.ExtractUtilityAllowances", _
                              "Allowance value is not numeric." & vbCrLf & vbCrLf & _
                              "UtilityType: " & optionCat & vbCrLf & _
                              "UtilityVariant: " & variantCode & vbCrLf & _
                              "Bedrooms: " & CStr(bedLabels(i)) & vbCrLf & _
                              "Value: " & CStr(amtVal)
                End If
                On Error GoTo Fail
            End If

            rows.Add Array(CLng(year), optionCat, variantCode, CStr(bedLabels(i)), Round(amt, 2))
        Next i

NextCol:
    Next col

    If rows.Count = 0 Then
        Err.Raise RENT_TABLES_ERR_IMPORT, "AMI_Optix_RentTables.ExtractUtilityAllowances", _
                  "Could not extract any utility allowance rows from 'AMI & Rent'."
    End If

    Set ExtractUtilityAllowances = rows
    Exit Function

Fail:
    Err.Raise RENT_TABLES_ERR_IMPORT, "AMI_Optix_RentTables.ExtractUtilityAllowances", _
              "ExtractUtilityAllowances failed: " & Err.Description
End Function

'-------------------------------------------------------------------------------
' NAMED TABLE EXTRACTION (PREFERRED WHEN AVAILABLE)
'-------------------------------------------------------------------------------

Private Function TryExtractRentLimitsFromNamedTable(rentWb As Workbook, year As Long, ByRef outRows As Collection) As Boolean
    On Error GoTo Fail
    TryExtractRentLimitsFromNamedTable = False
    If rentWb Is Nothing Then Exit Function

    Dim lo As ListObject
    Set lo = FindListObjectByName(rentWb, Array("rent_limits", "rentlimits", "tbl_rent_limits", "tblrentlimits"))
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim colProgram As Long, colBedrooms As Long, colAmi As Long, colGross As Long
    colProgram = ListObjectColumnIndex(lo, Array("program"))
    colBedrooms = ListObjectColumnIndex(lo, Array("bedrooms", "bedroom"))
    colAmi = ListObjectColumnIndex(lo, Array("ami"))
    colGross = ListObjectColumnIndex(lo, Array("grossrent", "gross_rent", "gross rent", "gross"))

    If colProgram = 0 Or colBedrooms = 0 Or colAmi = 0 Or colGross = 0 Then Exit Function

    Dim r As Long
    For r = 1 To lo.DataBodyRange.Rows.Count
        Dim program As String
        program = UCase$(Trim$(CStr(lo.DataBodyRange.Cells(r, colProgram).Value)))
        If program = "" Then program = "UAP"

        Dim bedLabel As String
        bedLabel = Trim$(CStr(lo.DataBodyRange.Cells(r, colBedrooms).Value))
        If bedLabel = "" Then GoTo NextR

        Dim amiVal As Variant
        amiVal = lo.DataBodyRange.Cells(r, colAmi).Value
        If Not IsNumeric(amiVal) Then GoTo NextR

        Dim grossVal As Variant
        grossVal = lo.DataBodyRange.Cells(r, colGross).Value
        If Not IsNumeric(grossVal) Then GoTo NextR

        outRows.Add Array(CLng(year), program, bedLabel, CanonAmiKey(CDbl(amiVal)), CDbl(grossVal))

NextR:
    Next r

    TryExtractRentLimitsFromNamedTable = (outRows.Count > 0)
    Exit Function

Fail:
    TryExtractRentLimitsFromNamedTable = False
End Function

Private Function TryExtractUtilityAllowancesFromNamedTable(rentWb As Workbook, year As Long, ByRef outRows As Collection) As Boolean
    On Error GoTo Fail
    TryExtractUtilityAllowancesFromNamedTable = False
    If rentWb Is Nothing Then Exit Function

    Dim lo As ListObject
    Set lo = FindListObjectByName(rentWb, Array("utility_allowances", "utilityallowances", "tbl_utility_allowances", "tblutilityallowances"))
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim colType As Long, colVariant As Long, colBedrooms As Long, colAmt As Long
    colType = ListObjectColumnIndex(lo, Array("utilitytype", "utility_type", "type"))
    colVariant = ListObjectColumnIndex(lo, Array("utilityvariant", "utility_variant", "variant"))
    colBedrooms = ListObjectColumnIndex(lo, Array("bedrooms", "bedroom"))
    colAmt = ListObjectColumnIndex(lo, Array("allowance", "amount"))

    If colType = 0 Or colVariant = 0 Or colBedrooms = 0 Or colAmt = 0 Then Exit Function

    Dim r As Long
    For r = 1 To lo.DataBodyRange.Rows.Count
        Dim uType As String
        uType = LCase$(Trim$(CStr(lo.DataBodyRange.Cells(r, colType).Value)))
        If uType = "" Then GoTo NextR

        Dim uVar As String
        uVar = LCase$(Trim$(CStr(lo.DataBodyRange.Cells(r, colVariant).Value)))
        If uVar = "" Then GoTo NextR

        Dim bed As String
        bed = Trim$(CStr(lo.DataBodyRange.Cells(r, colBedrooms).Value))
        If bed = "" Then GoTo NextR

        Dim amtVal As Variant
        amtVal = lo.DataBodyRange.Cells(r, colAmt).Value
        If Not IsNumeric(amtVal) Then GoTo NextR

        outRows.Add Array(CLng(year), uType, uVar, bed, CDbl(amtVal))

NextR:
    Next r

    TryExtractUtilityAllowancesFromNamedTable = (outRows.Count > 0)
    Exit Function

Fail:
    TryExtractUtilityAllowancesFromNamedTable = False
End Function

'-------------------------------------------------------------------------------
' CSV IO
'-------------------------------------------------------------------------------

Private Sub WriteRentLimitsCsv(path As String, rows As Collection)
    On Error GoTo Fail
    Dim f As Integer
    f = FreeFile
    Open path For Output As #f
    Print #f, "Year,Program,Bedrooms,AMI,GrossRent"

    Dim i As Long
    For i = 1 To rows.Count
        Dim r As Variant
        r = rows(i)
        Print #f, CsvEscape(CStr(r(0))) & "," & CsvEscape(CStr(r(1))) & "," & CsvEscape(CStr(r(2))) & "," & _
                  CsvEscape(CStr(r(3))) & "," & CsvEscape(InvariantNumberStr(CDbl(r(4))))
    Next i

    Close #f
    Exit Sub

Fail:
    On Error Resume Next
    If f <> 0 Then Close #f
    On Error GoTo 0
    Err.Raise RENT_TABLES_ERR_CACHE, "AMI_Optix_RentTables.WriteRentLimitsCsv", "WriteRentLimitsCsv failed: " & Err.Description
End Sub

Private Sub WriteUtilityAllowancesCsv(path As String, rows As Collection)
    On Error GoTo Fail
    Dim f As Integer
    f = FreeFile
    Open path For Output As #f
    Print #f, "Year,UtilityType,UtilityVariant,Bedrooms,Allowance"

    Dim i As Long
    For i = 1 To rows.Count
        Dim r As Variant
        r = rows(i)
        Print #f, CsvEscape(CStr(r(0))) & "," & CsvEscape(CStr(r(1))) & "," & CsvEscape(CStr(r(2))) & "," & _
                  CsvEscape(CStr(r(3))) & "," & CsvEscape(InvariantNumberStr(CDbl(r(4))))
    Next i

    Close #f
    Exit Sub

Fail:
    On Error Resume Next
    If f <> 0 Then Close #f
    On Error GoTo 0
    Err.Raise RENT_TABLES_ERR_CACHE, "AMI_Optix_RentTables.WriteUtilityAllowancesCsv", "WriteUtilityAllowancesCsv failed: " & Err.Description
End Sub

Private Function CsvEscape(value As String) As String
    Dim s As String
    s = CStr(value)
    If InStr(s, """") > 0 Then s = Replace(s, """", """""")
    If InStr(s, ",") > 0 Or InStr(s, vbCr) > 0 Or InStr(s, vbLf) > 0 Then
        CsvEscape = """" & s & """"
    Else
        CsvEscape = s
    End If
End Function

Private Function InvariantNumberStr(value As Double) As String
    Dim s As String
    s = Trim$(CStr(value))
    s = Replace(s, ",", ".")
    InvariantNumberStr = s
End Function

'-------------------------------------------------------------------------------
' CACHE META
'-------------------------------------------------------------------------------

Private Sub WriteCacheMeta(path As String, year As Long, sourcePath As String, fingerprint As String)
    On Error GoTo Fail

    Dim f As Integer
    f = FreeFile
    Open path For Output As #f
    Print #f, "year=" & CStr(year)
    Print #f, "source_path=" & sourcePath
    Print #f, "source_fingerprint=" & fingerprint
    Print #f, "generated_at=" & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Close #f
    Exit Sub

Fail:
    On Error Resume Next
    If f <> 0 Then Close #f
    On Error GoTo 0
    Err.Raise RENT_TABLES_ERR_CACHE, "AMI_Optix_RentTables.WriteCacheMeta", "WriteCacheMeta failed: " & Err.Description
End Sub

Private Function CacheMetaMatches(metaPath As String, sourcePath As String, fingerprint As String) As Boolean
    On Error GoTo Fail
    CacheMetaMatches = False
    If Dir$(metaPath) = "" Then Exit Function

    Dim metaSource As String
    Dim metaFp As String
    metaSource = ""
    metaFp = ""

    If Not TryReadMeta(metaPath, metaSource, metaFp) Then Exit Function

    CacheMetaMatches = (UCase$(Trim$(metaSource)) = UCase$(Trim$(sourcePath)) And Trim$(metaFp) = Trim$(fingerprint))
    Exit Function

Fail:
    CacheMetaMatches = False
End Function

Private Function TryReadMeta(metaPath As String, ByRef metaSourcePath As String, ByRef metaFingerprint As String) As Boolean
    On Error GoTo Fail

    TryReadMeta = False
    metaSourcePath = ""
    metaFingerprint = ""

    Dim f As Integer
    f = FreeFile
    Open metaPath For Input As #f

    Dim line As String
    Do While Not EOF(f)
        Line Input #f, line
        Dim p As Long
        p = InStr(1, line, "=", vbBinaryCompare)
        If p > 0 Then
            Dim k As String
            Dim v As String
            k = LCase$(Trim$(Left$(line, p - 1)))
            v = Mid$(line, p + 1)
            Select Case k
                Case "source_path": metaSourcePath = v
                Case "source_fingerprint": metaFingerprint = v
            End Select
        End If
    Loop

    Close #f

    TryReadMeta = (Trim$(metaSourcePath) <> "" And Trim$(metaFingerprint) <> "")
    Exit Function

Fail:
    On Error Resume Next
    If f <> 0 Then Close #f
    On Error GoTo 0
    TryReadMeta = False
End Function

'-------------------------------------------------------------------------------
' VALIDATION (HARD FAIL ON MISSING COVERAGE)
'-------------------------------------------------------------------------------

Private Sub ValidateRentLimitsCoverage(rows As Collection, year As Long, sourcePath As String)
    On Error GoTo Fail

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary") ' Program|Bedrooms|AMI -> Gross

    Dim requiredAmiKeys As Variant
    requiredAmiKeys = Array(CanonAmiKey(0.4), CanonAmiKey(0.6), CanonAmiKey(0.8), CanonAmiKey(1#))

    Dim requiredBeds As Variant
    requiredBeds = Array("studio", "1 BR", "2 BR", "3 BR", "4 BR", "5 BR")

    Dim programs As Object
    Set programs = CreateObject("Scripting.Dictionary")

    Dim amiKeys As Object
    Set amiKeys = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To rows.Count
        Dim r As Variant
        r = rows(i)

        Dim program As String
        program = UCase$(Trim$(CStr(r(1))))
        Dim bed As String
        bed = Trim$(CStr(r(2)))
        Dim amiKey As String
        amiKey = Trim$(CStr(r(3)))
        Dim gross As Double
        gross = CDbl(r(4))

        If gross <= 0# Then
            Err.Raise RENT_TABLES_ERR_IMPORT, "AMI_Optix_RentTables.ValidateRentLimitsCoverage", _
                      "Gross rent must be > 0." & vbCrLf & vbCrLf & _
                      "Program: " & program & vbCrLf & _
                      "Bedrooms: " & bed & vbCrLf & _
                      "AMI: " & amiKey & vbCrLf & _
                      "GrossRent: " & CStr(gross)
        End If

        Dim key As String
        key = program & "|" & bed & "|" & amiKey

        If seen.Exists(key) Then
            If CDbl(seen(key)) <> gross Then
                Err.Raise RENT_TABLES_ERR_IMPORT, "AMI_Optix_RentTables.ValidateRentLimitsCoverage", _
                          "Duplicate rent limit key with different values." & vbCrLf & vbCrLf & _
                          "Key: " & key & vbCrLf & _
                          "ValueA: " & CStr(seen(key)) & vbCrLf & _
                          "ValueB: " & CStr(gross)
            End If
        Else
            seen(key) = gross
        End If

        programs(program) = True
        amiKeys(amiKey) = True
    Next i

    Dim ra As Long
    For ra = LBound(requiredAmiKeys) To UBound(requiredAmiKeys)
        Dim reqAmi As String
        reqAmi = CStr(requiredAmiKeys(ra))
        If Not amiKeys.Exists(reqAmi) Then
            Err.Raise RENT_TABLES_ERR_IMPORT, "AMI_Optix_RentTables.ValidateRentLimitsCoverage", _
                      "Rent limits missing required AMI band: " & reqAmi & vbCrLf & vbCrLf & _
                      "Year: " & CStr(year) & vbCrLf & _
                      "Workbook: " & sourcePath
        End If
    Next ra

    Dim progKey As Variant
    For Each progKey In programs.Keys
        Dim b As Long
        For b = LBound(requiredBeds) To UBound(requiredBeds)
            Dim bed As String
            bed = CStr(requiredBeds(b))
            For ra = LBound(requiredAmiKeys) To UBound(requiredAmiKeys)
                Dim amiKey2 As String
                amiKey2 = CStr(requiredAmiKeys(ra))
                Dim fullKey As String
                fullKey = CStr(progKey) & "|" & bed & "|" & amiKey2
                If Not seen.Exists(fullKey) Then
                    Err.Raise RENT_TABLES_ERR_IMPORT, "AMI_Optix_RentTables.ValidateRentLimitsCoverage", _
                              "Rent limits missing required key." & vbCrLf & vbCrLf & _
                              "Key: " & fullKey & vbCrLf & _
                              "Year: " & CStr(year) & vbCrLf & _
                              "Workbook: " & sourcePath
                End If
            Next ra
        Next b
    Next progKey

    Exit Sub

Fail:
    Err.Raise RENT_TABLES_ERR_IMPORT, "AMI_Optix_RentTables.ValidateRentLimitsCoverage", _
              "Rent limits validation failed: " & Err.Description
End Sub

Private Sub ValidateUtilityAllowanceCoverage(rows As Collection, year As Long, sourcePath As String)
    On Error GoTo Fail

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary") ' UtilityType|Variant|Bedrooms -> Allowance

    Dim requiredBeds As Variant
    requiredBeds = Array("studio", "1 BR", "2 BR", "3 BR", "4 BR", "5 BR")

    Dim requiredVariants As Object
    Set requiredVariants = CreateObject("Scripting.Dictionary")
    requiredVariants("electricity") = Array("tenant_pays")
    requiredVariants("cooking") = Array("electric", "gas")
    requiredVariants("heat") = Array("electric_ccashp", "electric_other", "gas", "oil")
    requiredVariants("hot_water") = Array("electric_heat_pump", "electric_other", "gas", "oil")

    Dim i As Long
    For i = 1 To rows.Count
        Dim r As Variant
        r = rows(i)

        Dim uType As String
        uType = LCase$(Trim$(CStr(r(1))))
        Dim uVar As String
        uVar = LCase$(Trim$(CStr(r(2))))
        Dim bed As String
        bed = Trim$(CStr(r(3)))
        Dim amt As Double
        amt = CDbl(r(4))

        Dim key As String
        key = uType & "|" & uVar & "|" & bed

        If seen.Exists(key) Then
            If CDbl(seen(key)) <> amt Then
                Err.Raise RENT_TABLES_ERR_IMPORT, "AMI_Optix_RentTables.ValidateUtilityAllowanceCoverage", _
                          "Duplicate utility allowance key with different values." & vbCrLf & vbCrLf & _
                          "Key: " & key & vbCrLf & _
                          "ValueA: " & CStr(seen(key)) & vbCrLf & _
                          "ValueB: " & CStr(amt)
            End If
        Else
            seen(key) = amt
        End If
    Next i

    Dim uTypeKey As Variant
    For Each uTypeKey In requiredVariants.Keys
        Dim variants As Variant
        variants = requiredVariants(uTypeKey)
        Dim v As Long
        For v = LBound(variants) To UBound(variants)
            Dim variantCode As String
            variantCode = CStr(variants(v))

            Dim b As Long
            For b = LBound(requiredBeds) To UBound(requiredBeds)
                Dim bed As String
                bed = CStr(requiredBeds(b))

                Dim fullKey As String
                fullKey = CStr(uTypeKey) & "|" & variantCode & "|" & bed

                If Not seen.Exists(fullKey) Then
                    Err.Raise RENT_TABLES_ERR_IMPORT, "AMI_Optix_RentTables.ValidateUtilityAllowanceCoverage", _
                              "Utility allowances missing required key." & vbCrLf & vbCrLf & _
                              "Key: " & fullKey & vbCrLf & _
                              "Year: " & CStr(year) & vbCrLf & _
                              "Workbook: " & sourcePath
                End If
            Next b
        Next v
    Next uTypeKey

    Exit Sub

Fail:
    Err.Raise RENT_TABLES_ERR_IMPORT, "AMI_Optix_RentTables.ValidateUtilityAllowanceCoverage", _
              "Utility allowances validation failed: " & Err.Description
End Sub

'-------------------------------------------------------------------------------
' DIAGNOSTICS
'-------------------------------------------------------------------------------

Private Sub LogRentTablesDiagnostics(year As Long, sourcePath As String, cacheFolder As String, fingerprint As String, rebuilt As Boolean, sourceLabel As String)
    On Error Resume Next

    ' Lightweight always-on breadcrumb in debug log.
    DebugLog "RentTablesCache " & IIf(rebuilt, "REFRESH", "OK") & " | year=" & CStr(year) & _
             " | src=" & sourceLabel & " | wb=" & sourcePath & " | cache=" & cacheFolder & _
             " | fp=" & fingerprint, True

    ' Append-only JSONL diagnostics (reuses existing run log).
    Dim payload As String
    payload = "{"
    payload = payload & """year"":" & CStr(year) & ","
    payload = payload & """source_path"":""" & EscapeJsonString(sourcePath) & ""","
    payload = payload & """cache_folder"":""" & EscapeJsonString(cacheFolder) & ""","
    payload = payload & """source_fingerprint"":""" & EscapeJsonString(fingerprint) & ""","
    payload = payload & """rebuilt"":" & IIf(rebuilt, "true", "false") & ","
    payload = payload & """source_label"":""" & EscapeJsonString(sourceLabel) & """"
    payload = payload & "}"

    AppendRunLog "rent_tables_cache", payload
End Sub

Private Function EscapeJsonString(value As String) As String
    Dim result As String
    result = CStr(value)
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCr, "\r")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")
    EscapeJsonString = result
End Function

'-------------------------------------------------------------------------------
' HELPERS
'-------------------------------------------------------------------------------

Private Function FirstWorkbookInFolder(folderPath As String) As String
    On Error GoTo SafeExit

    If Trim$(folderPath) = "" Then GoTo SafeExit

    Dim f As String
    f = Dir$(folderPath & "\*.xlsx")
    If Trim$(f) <> "" Then
        FirstWorkbookInFolder = folderPath & "\" & f
        Exit Function
    End If

    f = Dir$(folderPath & "\*.xlsm")
    If Trim$(f) <> "" Then
        FirstWorkbookInFolder = folderPath & "\" & f
        Exit Function
    End If

SafeExit:
    FirstWorkbookInFolder = ""
End Function

Private Sub EnsureFolderExistsSafe(path As String)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Sub
    If Not fso.FolderExists(path) Then fso.CreateFolder path
End Sub

Private Function FindOpenWorkbookByPath(fullPath As String) As Workbook
    On Error GoTo SafeExit
    Set FindOpenWorkbookByPath = Nothing
    If Trim$(fullPath) = "" Then Exit Function

    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If UCase$(Trim$(wb.FullName)) = UCase$(Trim$(fullPath)) Then
            Set FindOpenWorkbookByPath = wb
            Exit Function
        End If
    Next wb

SafeExit:
    Set FindOpenWorkbookByPath = Nothing
End Function

Private Function FindListObjectByName(wb As Workbook, names As Variant) As ListObject
    On Error GoTo SafeExit
    Set FindListObjectByName = Nothing
    If wb Is Nothing Then Exit Function

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        Dim lo As ListObject
        For Each lo In ws.ListObjects
            Dim i As Long
            For i = LBound(names) To UBound(names)
                If LCase$(Trim$(lo.Name)) = LCase$(Trim$(CStr(names(i)))) Then
                    Set FindListObjectByName = lo
                    Exit Function
                End If
            Next i
        Next lo
    Next ws

SafeExit:
    Set FindListObjectByName = Nothing
End Function

Private Function ListObjectColumnIndex(lo As ListObject, names As Variant) As Long
    On Error GoTo SafeExit
    ListObjectColumnIndex = 0
    If lo Is Nothing Then Exit Function

    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        Dim header As String
        header = LCase$(Trim$(CStr(lc.Name)))
        Dim i As Long
        For i = LBound(names) To UBound(names)
            Dim n As String
            n = LCase$(Trim$(CStr(names(i))))
            If header = n Then
                ListObjectColumnIndex = lc.Index
                Exit Function
            End If
        Next i
    Next lc

SafeExit:
    ListObjectColumnIndex = 0
End Function

Private Function CanonAmiKey(ami As Double) As String
    CanonAmiKey = Format$(Round(CDbl(ami), 4), "0.0000")
End Function

Private Function BedroomLabelFromSheetLabel(label As String) As String
    Dim v As String
    v = LCase$(Trim$(CStr(label)))

    Select Case v
        Case "studio": BedroomLabelFromSheetLabel = "studio"
        Case "1 br": BedroomLabelFromSheetLabel = "1 BR"
        Case "2 br": BedroomLabelFromSheetLabel = "2 BR"
        Case "3 br": BedroomLabelFromSheetLabel = "3 BR"
        Case "4 br": BedroomLabelFromSheetLabel = "4 BR"
        Case "5 br": BedroomLabelFromSheetLabel = "5 BR"
        Case Else: BedroomLabelFromSheetLabel = ""
    End Select
End Function

Private Function UtilityCategoryFromHeaderValue(headerValue As String) As String
    Dim h As String
    h = LCase$(Trim$(CStr(headerValue)))

    Select Case h
        Case "apartment electricity only": UtilityCategoryFromHeaderValue = "electricity"
        Case "cooking": UtilityCategoryFromHeaderValue = "cooking"
        Case "heat": UtilityCategoryFromHeaderValue = "heat"
        Case "hot water": UtilityCategoryFromHeaderValue = "hot_water"
        Case Else: UtilityCategoryFromHeaderValue = UtilityCategoryFromOptionLabel(headerValue)
    End Select
End Function

Private Function UtilityCategoryFromOptionLabel(optionLabel As String) As String
    Dim v As String
    v = Trim$(CStr(optionLabel))

    Select Case v
        Case "Tenant Pays": UtilityCategoryFromOptionLabel = "electricity"
        Case "Electric Stove", "Gas Stove": UtilityCategoryFromOptionLabel = "cooking"
        Case "Electric Heat - Cold Climate Air Source Heat Pump (ccASHP)1", "Electric Heat - Other2", "Gas Heat", "Oil Heat": UtilityCategoryFromOptionLabel = "heat"
        Case "Electric Hot Water - Heat Pump", "Electric Hot Water - Other", "Gas Hot Water", "Oil Hot Water": UtilityCategoryFromOptionLabel = "hot_water"
        Case Else: UtilityCategoryFromOptionLabel = ""
    End Select
End Function

Private Function UtilityVariantCodeFromOptionLabel(category As String, optionLabel As String) As String
    Dim cat As String
    cat = LCase$(Trim$(CStr(category)))

    Dim label As String
    label = NormalizeOptionLabel(optionLabel)

    Select Case cat
        Case "electricity"
            If label = "tenant pays" Then UtilityVariantCodeFromOptionLabel = "tenant_pays"

        Case "cooking"
            If label = "electric stove" Then
                UtilityVariantCodeFromOptionLabel = "electric"
            ElseIf label = "gas stove" Then
                UtilityVariantCodeFromOptionLabel = "gas"
            End If

        Case "heat"
            If InStr(1, label, "cold climate air source heat pump", vbTextCompare) > 0 Or InStr(1, label, "ccashp", vbTextCompare) > 0 Then
                UtilityVariantCodeFromOptionLabel = "electric_ccashp"
            ElseIf label = "electric heat - other2" Or (InStr(1, label, "electric heat", vbTextCompare) > 0 And InStr(1, label, "other", vbTextCompare) > 0) Then
                UtilityVariantCodeFromOptionLabel = "electric_other"
            ElseIf label = "gas heat" Then
                UtilityVariantCodeFromOptionLabel = "gas"
            ElseIf label = "oil heat" Then
                UtilityVariantCodeFromOptionLabel = "oil"
            End If

        Case "hot_water"
            If label = "electric hot water - heat pump" Or InStr(1, label, "heat pump", vbTextCompare) > 0 Then
                UtilityVariantCodeFromOptionLabel = "electric_heat_pump"
            ElseIf label = "electric hot water - other" Or (InStr(1, label, "electric hot water", vbTextCompare) > 0 And InStr(1, label, "other", vbTextCompare) > 0) Then
                UtilityVariantCodeFromOptionLabel = "electric_other"
            ElseIf label = "gas hot water" Then
                UtilityVariantCodeFromOptionLabel = "gas"
            ElseIf label = "oil hot water" Then
                UtilityVariantCodeFromOptionLabel = "oil"
            End If
    End Select
End Function

Private Function NormalizeOptionLabel(optionLabel As String) As String
    Dim s As String
    s = Trim$(CStr(optionLabel))
    s = Replace(s, Chr$(160), " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormalizeOptionLabel = LCase$(s)
End Function
