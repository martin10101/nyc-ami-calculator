Attribute VB_Name = "AMI_Optix_Diagnostics"
'===============================================================================
' AMI OPTIX - Diagnostics
' Produces copy/paste friendly troubleshooting info without requiring the VBE.
'===============================================================================
Option Explicit

Private Const DIAG_SHEET_NAME As String = "AMI Optix Diagnostics"

Public Sub ShowAMIOptixDiagnostics()
    On Error GoTo Fail

    If ActiveWorkbook Is Nothing Then
        MsgBox "No workbook is open.", vbExclamation, "AMI Optix"
        Exit Sub
    End If

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim ws As Worksheet
    Set ws = GetOrCreateDiagnosticsSheet(wb)

    ws.Cells.Clear
    ws.Range("A:A").ColumnWidth = 34
    ws.Range("B:B").ColumnWidth = 110

    Dim r As Long
    r = 1

    ws.Cells(r, 1).Value = "AMI Optix Diagnostics"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Size = 14
    r = r + 2

    WriteKV ws, r, "Generated", Format$(Now, "yyyy-mm-dd hh:nn:ss")
    WriteKV ws, r, "User", Environ$("USERNAME")
    WriteKV ws, r, "Computer", Environ$("COMPUTERNAME")
    WriteKV ws, r, "Excel", Application.Version & " (" & Application.OperatingSystem & ")"
    WriteKV ws, r, "Workbook", wb.Name
    WriteKV ws, r, "Workbook Path", wb.FullName
    WriteKV ws, r, "API Base URL", API_BASE_URL
    r = r + 1

    ' -----------------------------------------------------------------------
    ' Rent Tables Status (Fix-06c)
    ' -----------------------------------------------------------------------

    ws.Cells(r, 1).Value = "Rent Tables Status"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1

    Dim year As Long
    year = 0
    On Error Resume Next
    year = CLng(GetSetting("AMI_Optix", "RentRollYears", "SelectedYear", "2025"))
    On Error GoTo 0
    If year <= 0 Then year = 2025

    WriteKV ws, r, "Selected Year", CStr(year)

    Dim rentStatus As Object
    Set rentStatus = Nothing
    Dim rentStatusErr As String
    rentStatusErr = ""

    On Error Resume Next
    Set rentStatus = GetRentTablesStatus(year)
    If Err.Number <> 0 Then rentStatusErr = Err.Description
    Err.Clear
    On Error GoTo 0

    If rentStatus Is Nothing Or rentStatusErr <> "" Then
        WriteKV ws, r, "Status", "ERROR: " & rentStatusErr
    Else
        WriteKV ws, r, "Cache Status", CStr(rentStatus("cache_status"))
        WriteKV ws, r, "Cache Folder", CStr(rentStatus("cache_folder"))
        WriteKV ws, r, "Cache Built At", CStr(rentStatus("cache_generated_at"))
        WriteKV ws, r, "Cache Built From", CStr(rentStatus("cache_source_label")) & " | " & CStr(rentStatus("cache_source_path"))
        WriteKV ws, r, "Cache Fingerprint", CStr(rentStatus("cache_source_fingerprint"))
        WriteKV ws, r, "Cache Build Reason", CStr(rentStatus("cache_build_reason"))
        WriteKV ws, r, "Source (Resolved Now)", CStr(rentStatus("resolved_source_label")) & " | " & CStr(rentStatus("resolved_source_path"))
        WriteKV ws, r, "Source Last Modified", CStr(rentStatus("resolved_source_last_modified"))
        WriteKV ws, r, "Source Fingerprint (Now)", CStr(rentStatus("resolved_fingerprint"))
        If CStr(rentStatus("resolved_source_error")) <> "" Then
            WriteKV ws, r, "Source Access Error", CStr(rentStatus("resolved_source_error"))
        End If
    End If

    r = r + 1

    ' -----------------------------------------------------------------------
    ' Verify Manual Rents (API) (Fix-06d)
    ' -----------------------------------------------------------------------

    ws.Cells(r, 1).Value = "Verify Manual Rents (API)"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1

    If g_AMIOptixLastVerify Is Nothing Then
        WriteKV ws, r, "Last Verify", "(not run this session)"
    Else
        On Error Resume Next
        WriteKV ws, r, "Last Verify At", CStr(g_AMIOptixLastVerify("timestamp"))
        WriteKV ws, r, "Result", CStr(g_AMIOptixLastVerify("result"))
        WriteKV ws, r, "Program", CStr(g_AMIOptixLastVerify("program"))
        WriteKV ws, r, "Year", CStr(g_AMIOptixLastVerify("rent_roll_year"))

        WriteKV ws, r, "Local Cache Status", CStr(g_AMIOptixLastVerify("local_cache_status"))
        WriteKV ws, r, "Local Cache Built At", CStr(g_AMIOptixLastVerify("local_cache_generated_at"))
        WriteKV ws, r, "Local Cache Built From", CStr(g_AMIOptixLastVerify("local_cache_source_label")) & " | " & CStr(g_AMIOptixLastVerify("local_cache_source_path"))

        Dim apiLine As String
        apiLine = ""
        If Trim$(CStr(g_AMIOptixLastVerify("api_rent_roll_year_used"))) <> "" Then apiLine = apiLine & "year_used=" & CStr(g_AMIOptixLastVerify("api_rent_roll_year_used"))
        If Trim$(CStr(g_AMIOptixLastVerify("api_calculator_filename"))) <> "" Then
            If apiLine <> "" Then apiLine = apiLine & " | "
            apiLine = apiLine & "calculator=" & CStr(g_AMIOptixLastVerify("api_calculator_filename"))
        End If
        If Trim$(CStr(g_AMIOptixLastVerify("api_rent_schedule_source"))) <> "" Then
            If apiLine <> "" Then apiLine = apiLine & " | "
            apiLine = apiLine & "source=" & CStr(g_AMIOptixLastVerify("api_rent_schedule_source"))
        End If
        If apiLine = "" Then apiLine = "(unknown)"
        WriteKV ws, r, "API", apiLine

        If Trim$(CStr(g_AMIOptixLastVerify("api_rent_schedule_warning"))) <> "" Then
            WriteKV ws, r, "API Warning", CStr(g_AMIOptixLastVerify("api_rent_schedule_warning"))
        End If

        WriteKV ws, r, "Local Net Monthly", CStr(g_AMIOptixLastVerify("local_net_monthly"))
        WriteKV ws, r, "API Net Monthly", CStr(g_AMIOptixLastVerify("api_net_monthly"))
        WriteKV ws, r, "Total Delta", CStr(g_AMIOptixLastVerify("total_delta"))

        WriteKV ws, r, "Compare Per-Unit", CStr(g_AMIOptixLastVerify("compare_per_unit"))
        WriteKV ws, r, "Compare Totals", CStr(g_AMIOptixLastVerify("compare_totals"))
        WriteKV ws, r, "Tolerance ($/unit)", CStr(g_AMIOptixLastVerify("tolerance_unit_dollars"))
        WriteKV ws, r, "Tolerance ($ total)", CStr(g_AMIOptixLastVerify("tolerance_total_dollars"))

        If Trim$(CStr(g_AMIOptixLastVerify("errors"))) <> "" Then
            WriteKV ws, r, "Errors", CStr(g_AMIOptixLastVerify("errors"))
        End If
        On Error GoTo 0

        ' Full mismatch list (if any).
        On Error Resume Next
        Dim mismatches As Collection
        Set mismatches = Nothing
        Set mismatches = g_AMIOptixLastVerify("mismatches")
        On Error GoTo 0

        If Not mismatches Is Nothing Then
            If mismatches.Count > 0 Then
                r = r + 1
                ws.Cells(r, 1).Value = "Mismatches"
                ws.Cells(r, 1).Font.Bold = True
                r = r + 1

                ws.Cells(r, 1).Value = "unit_id"
                ws.Cells(r, 2).Value = "local_monthly_rent"
                ws.Cells(r, 3).Value = "api_monthly_rent"
                ws.Cells(r, 4).Value = "delta"
                ws.Cells(r, 5).Value = "reason"
                ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Font.Bold = True
                r = r + 1

                Dim mi As Long
                For mi = 1 To mismatches.Count
                    Dim m As Object
                    Set m = mismatches(mi)
                    If Not m Is Nothing Then
                        ws.Cells(r, 1).Value = CStr(m("unit_id"))
                        ws.Cells(r, 2).Value = m("local_monthly_rent")
                        ws.Cells(r, 3).Value = m("api_monthly_rent")
                        ws.Cells(r, 4).Value = m("delta")
                        ws.Cells(r, 5).Value = CStr(m("reason"))
                        r = r + 1
                    End If
                Next mi
            End If
        End If
    End If

    r = r + 1

    ws.Cells(r, 1).Value = "Run Log"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1

    Dim rootPath As String
    rootPath = GetLearningLogRootPath()
    WriteKV ws, r, "Log Root", rootPath

    Dim runLogPath As String
    runLogPath = GetRunLogFilePath()
    WriteKV ws, r, "Run Log File", runLogPath

    Dim runLogFileName As String
    runLogFileName = Dir$(runLogPath)
    WriteKV ws, r, "Dir(Run Log File)", runLogFileName

    Dim lastErr As String
    lastErr = GetLastRunLogError()
    If Trim$(lastErr) <> "" Then
        WriteKV ws, r, "Last Log Error", lastErr
    Else
        WriteKV ws, r, "Last Log Error", "(none)"
    End If
    r = r + 1

    ws.Cells(r, 1).Value = "Last API Scenarios"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1

    If g_LastScenarios Is Nothing Then
        WriteKV ws, r, "g_LastScenarios", "(Nothing) - run the solver first"
    Else
        WriteKV ws, r, "g_LastScenarios", TypeName(g_LastScenarios)

        Dim scenarios As Object
        Set scenarios = Nothing
        On Error Resume Next
        If HasKey(g_LastScenarios, "scenarios") Then Set scenarios = g_LastScenarios("scenarios")
        On Error GoTo 0

        If scenarios Is Nothing Then
            WriteKV ws, r, "scenarios", "(missing)"
        Else
            WriteKV ws, r, "scenarios.Count", CStr(GetObjectCountSafe(scenarios))
            r = r + 1
            ws.Cells(r, 1).Value = "Scenario Keys"
            ws.Cells(r, 1).Font.Bold = True
            r = r + 1
            r = WriteKeysList(ws, r, scenarios)
            r = r + 1
        End If

        Dim notes As Object
        Set notes = Nothing
        On Error Resume Next
        If HasKey(g_LastScenarios, "notes") Then Set notes = g_LastScenarios("notes")
        On Error GoTo 0

        ws.Cells(r, 1).Value = "Solver Notes"
        ws.Cells(r, 1).Font.Bold = True
        r = r + 1
        If notes Is Nothing Then
            ws.Cells(r, 1).Value = "(none)"
            r = r + 1
        Else
            r = WriteNotesList(ws, r, notes)
        End If
    End If

    ws.Activate
    ws.Range("A1").Select
    Exit Sub

Fail:
    MsgBox "Diagnostics failed: " & Err.Description, vbExclamation, "AMI Optix"
End Sub

Private Sub WriteKV(ws As Worksheet, ByRef r As Long, label As String, value As String)
    ws.Cells(r, 1).Value = label
    ws.Cells(r, 2).Value = value
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1
End Sub

Private Function GetOrCreateDiagnosticsSheet(wb As Workbook) As Worksheet
    On Error GoTo CreateNew

    Dim ws As Worksheet
    Set ws = wb.Worksheets(DIAG_SHEET_NAME)
    Set GetOrCreateDiagnosticsSheet = ws
    Exit Function

CreateNew:
    On Error GoTo 0
    Dim newWs As Worksheet
    Set newWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    newWs.Name = DIAG_SHEET_NAME
    Set GetOrCreateDiagnosticsSheet = newWs
End Function

Private Function HasKey(d As Object, key As String) As Boolean
    On Error GoTo Fail
    If d Is Nothing Then Exit Function
    HasKey = CBool(d.Exists(key))
    Exit Function
Fail:
    HasKey = False
End Function

Private Function GetObjectCountSafe(obj As Object) As Long
    On Error GoTo Fail
    If obj Is Nothing Then Exit Function
    GetObjectCountSafe = CLng(obj.Count)
    Exit Function
Fail:
    GetObjectCountSafe = 0
End Function

Private Function WriteKeysList(ws As Worksheet, ByVal r As Long, scenarios As Object) As Long
    On Error GoTo Fail

    Dim wroteAny As Boolean
    wroteAny = False

    Dim k As Variant
    On Error Resume Next
    For Each k In scenarios.Keys
        ws.Cells(r, 1).Value = CStr(k)
        r = r + 1
        wroteAny = True
    Next k
    On Error GoTo 0

    If Not wroteAny Then
        ws.Cells(r, 1).Value = "(no keys)"
        r = r + 1
    End If

    WriteKeysList = r
    Exit Function

Fail:
    ws.Cells(r, 1).Value = "(could not enumerate keys)"
    WriteKeysList = r + 1
End Function

Private Function WriteNotesList(ws As Worksheet, ByVal r As Long, notes As Object) As Long
    On Error GoTo Fail

    Dim wroteAny As Boolean
    wroteAny = False

    Dim i As Long
    If TypeName(notes) = "Collection" Then
        For i = 1 To notes.Count
            ws.Cells(r, 1).Value = CStr(notes(i))
            r = r + 1
            wroteAny = True
        Next i
    Else
        ' Try 1-based indexing for other list-like objects
        For i = 1 To CLng(notes.Count)
            On Error Resume Next
            ws.Cells(r, 1).Value = CStr(notes(i))
            If Err.Number = 0 Then
                r = r + 1
                wroteAny = True
            End If
            Err.Clear
            On Error GoTo Fail
        Next i
    End If

    If Not wroteAny Then
        ws.Cells(r, 1).Value = "(no notes)"
        r = r + 1
    End If

    WriteNotesList = r
    Exit Function

Fail:
    ws.Cells(r, 1).Value = "(could not enumerate notes)"
    WriteNotesList = r + 1
End Function
