Attribute VB_Name = "AMI_Optix_API"
'===============================================================================
' AMI OPTIX - API Communication Module
' Handles HTTP requests to the optimization API and JSON parsing
'
' Compatible with both 32-bit and 64-bit Office
'===============================================================================
Option Explicit

'-------------------------------------------------------------------------------
' RENT CALCULATOR YEAR (Fix-03)
'-------------------------------------------------------------------------------

Private Const AMI_OPTIX_REGISTRY_PATH As String = "AMI_Optix"
Private Const RENTROLL_YEAR_REG_SECTION As String = "RentRollYears"
Private Const RENTROLL_YEAR_REG_KEY_SELECTED As String = "SelectedYear"

Private Const RENTROLL_YEAR_MIN As Long = 2022
Private Const RENTROLL_YEAR_MAX As Long = 2026
Private Const RENTROLL_YEAR_DEFAULT As Long = 2025

Private Const RENT_CALC_REMOTE_PREFIX As String = "AMI_Optix_Rent_Calculator_"
Private Const RENT_CALC_REMOTE_SUFFIX As String = ".xlsx"
Private Const RENT_CALC_REMOTE_DEFAULT_NAME As String = "default"

Private m_LastActivatedRentCalcName As String
Private m_LastActivationWarnedName As String
Private m_DefaultYearOverrideUnavailable As Boolean

'-------------------------------------------------------------------------------
' API CALL
'-------------------------------------------------------------------------------

Public Function EnsureSelectedRentRollYearActive(Optional showErrors As Boolean = False) As Boolean
    ' Ensures the API is using the rent calculator matching the user's selected Rent Roll Year.
    Dim year As Long
    year = GetSelectedRentRollYearSetting()
    EnsureSelectedRentRollYearActive = EnsureRentCalculatorYearActive(year, showErrors)
End Function

Public Function ActivateRentCalculatorByName(name As String, Optional showErrors As Boolean = True) As Boolean
    ActivateRentCalculatorByName = ActivateRentCalculatorByNameInternal(CStr(name), showErrors)

    If ActivateRentCalculatorByName Then
        m_LastActivatedRentCalcName = CStr(name)
        m_LastActivationWarnedName = ""
    End If
End Function

Public Function UploadRentCalculatorFile( _
    localFilePath As String, _
    remoteFileName As String, _
    Optional overwrite As Boolean = True, _
    Optional showErrors As Boolean = True _
) As Boolean
    ' Uploads a rent calculator workbook to the API storage (multipart/form-data).
    On Error GoTo Fail

    If Trim$(localFilePath) = "" Or Dir(localFilePath) = "" Then
        If showErrors Then
            MsgBox "File not found:" & vbCrLf & localFilePath, vbExclamation, "AMI Optix"
        End If
        Exit Function
    End If

    Dim boundary As String
    boundary = MakeMultipartBoundary()

    Dim body As Variant
    body = BuildRentCalculatorUploadBodyBytes(localFilePath, remoteFileName, overwrite, boundary)

    Dim http As Object
    Dim url As String
    Dim apiKey As String

    url = API_BASE_URL & "/api/rent-calculators/upload"
    apiKey = GetAPIKey()

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 5000, 30000, 30000, 240000

    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
    http.setRequestHeader "Accept", "application/json"
    If Len(apiKey) > 0 Then
        http.setRequestHeader "X-API-Key", apiKey
    End If

    http.send body

    If http.Status = 200 Then
        UploadRentCalculatorFile = True
    Else
        If showErrors Then
            Dim msg As String
            msg = "Rent calculator upload failed." & vbCrLf & _
                  "Status: " & http.Status & " - " & http.statusText
            If Len(http.responseText) > 0 Then msg = msg & vbCrLf & vbCrLf & http.responseText
            MsgBox msg, vbExclamation, "AMI Optix"
        End If
        UploadRentCalculatorFile = False
    End If
    Exit Function

Fail:
    If showErrors Then
        MsgBox "Rent calculator upload failed: " & Err.Description, vbExclamation, "AMI Optix"
    End If
    UploadRentCalculatorFile = False
End Function

Private Function GetSelectedRentRollYearSetting() As Long
    ' Uses the ribbon dropdown selection (registry). Defaults to 2025 if not set.
    Dim raw As String
    raw = GetSetting(AMI_OPTIX_REGISTRY_PATH, RENTROLL_YEAR_REG_SECTION, RENTROLL_YEAR_REG_KEY_SELECTED, CStr(RENTROLL_YEAR_DEFAULT))

    Dim y As Long
    y = RENTROLL_YEAR_DEFAULT
    On Error Resume Next
    y = CLng(raw)
    On Error GoTo 0

    If y < RENTROLL_YEAR_MIN Or y > RENTROLL_YEAR_MAX Then y = RENTROLL_YEAR_DEFAULT
    GetSelectedRentRollYearSetting = y
End Function

Private Function EnsureRentCalculatorYearActive(year As Long, showErrors As Boolean) As Boolean
    Dim targetName As String

    If year = RENTROLL_YEAR_DEFAULT Then
        ' Default year normally uses the built-in default calculator.
        ' If the user uploaded a 2025 override (same naming convention), prefer it when available.
        Dim overrideName As String
        overrideName = RentCalculatorRemoteNameForYear(year)

        If m_LastActivatedRentCalcName = overrideName Then
            EnsureRentCalculatorYearActive = True
            Exit Function
        End If

        If Not m_DefaultYearOverrideUnavailable Then
            If ActivateRentCalculatorByNameInternal(overrideName, False) Then
                m_LastActivatedRentCalcName = overrideName
                m_LastActivationWarnedName = ""
                EnsureRentCalculatorYearActive = True
                Exit Function
            End If
            m_DefaultYearOverrideUnavailable = True
        End If

        targetName = RENT_CALC_REMOTE_DEFAULT_NAME
    Else
        targetName = RentCalculatorRemoteNameForYear(year)
    End If

    If targetName = "" Then
        EnsureRentCalculatorYearActive = True
        Exit Function
    End If

    If m_LastActivatedRentCalcName = targetName Then
        EnsureRentCalculatorYearActive = True
        Exit Function
    End If

    Dim allowUI As Boolean
    allowUI = showErrors
    If showErrors Then
        If m_LastActivationWarnedName = targetName Then allowUI = False
    End If

    If ActivateRentCalculatorByNameInternal(targetName, allowUI) Then
        m_LastActivatedRentCalcName = targetName
        m_LastActivationWarnedName = ""
        EnsureRentCalculatorYearActive = True
        Exit Function
    End If

    If showErrors And m_LastActivationWarnedName <> targetName Then
        m_LastActivationWarnedName = targetName
    End If

    EnsureRentCalculatorYearActive = False
End Function

Private Function RentCalculatorRemoteNameForYear(year As Long) As String
    RentCalculatorRemoteNameForYear = RENT_CALC_REMOTE_PREFIX & CStr(year) & RENT_CALC_REMOTE_SUFFIX
End Function

Private Function ActivateRentCalculatorByNameInternal(name As String, showErrors As Boolean) As Boolean
    On Error GoTo Fail

    Dim http As Object
    Dim url As String
    Dim apiKey As String
    Dim payload As String

    url = API_BASE_URL & "/api/rent-calculators/activate"
    apiKey = GetAPIKey()
    payload = "{""name"": """ & EscapeJSON(CStr(name)) & """}"

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 5000, 30000, 30000, 60000

    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json"
    If Len(apiKey) > 0 Then
        http.setRequestHeader "X-API-Key", apiKey
    End If

    http.send payload

    If http.Status = 200 Then
        ActivateRentCalculatorByNameInternal = True
    Else
        If showErrors Then
            Dim msg As String
            msg = "Could not activate rent calculator." & vbCrLf & _
                  "Status: " & http.Status & " - " & http.statusText & vbCrLf & _
                  "Name: " & CStr(name)
            If http.Status = 401 Then
                msg = msg & vbCrLf & vbCrLf & _
                      "This endpoint requires admin authorization." & vbCrLf & _
                      "Render fix: set AMI_OPTIX_ADMIN_KEY equal to AMI_OPTIX_API_KEY, " & _
                      "or set AMI_OPTIX_ALLOW_API_KEY_FOR_ADMIN=1."
            End If
            If Len(http.responseText) > 0 Then msg = msg & vbCrLf & vbCrLf & http.responseText
            MsgBox msg, vbExclamation, "AMI Optix"
        End If
        ActivateRentCalculatorByNameInternal = False
    End If
    Exit Function

Fail:
    If showErrors Then
        MsgBox "Could not activate rent calculator: " & Err.Description, vbExclamation, "AMI Optix"
    End If
    ActivateRentCalculatorByNameInternal = False
End Function

Private Function MakeMultipartBoundary() As String
    Randomize
    MakeMultipartBoundary = "----AMIOptixBoundary" & Format$(Now, "yyyymmddhhnnss") & CStr(Int(Rnd() * 1000000#))
End Function

Private Function BuildRentCalculatorUploadBodyBytes( _
    localFilePath As String, _
    remoteFileName As String, _
    overwrite As Boolean, _
    boundary As String _
) As Variant
    Dim fileBytes As Variant
    fileBytes = ReadBinaryFileBytes(localFilePath)

    Dim safeName As String
    safeName = Replace(CStr(remoteFileName), """", "_")

    Dim prefix As String
    prefix = "--" & boundary & vbCrLf & _
             "Content-Disposition: form-data; name=""overwrite""" & vbCrLf & vbCrLf & _
             IIf(overwrite, "true", "false") & vbCrLf & _
             "--" & boundary & vbCrLf & _
             "Content-Disposition: form-data; name=""file""; filename=""" & safeName & """" & vbCrLf & _
             "Content-Type: application/octet-stream" & vbCrLf & vbCrLf

    Dim suffix As String
    suffix = vbCrLf & "--" & boundary & "--" & vbCrLf

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1 ' adTypeBinary
    stm.Open
    stm.Write StrToBytes(prefix)
    stm.Write fileBytes
    stm.Write StrToBytes(suffix)
    stm.Position = 0
    BuildRentCalculatorUploadBodyBytes = stm.Read
    stm.Close
End Function

Private Function ReadBinaryFileBytes(filePath As String) As Variant
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1 ' adTypeBinary
    stm.Open
    stm.LoadFromFile filePath
    ReadBinaryFileBytes = stm.Read
    stm.Close
End Function

Private Function StrToBytes(s As String) As Variant
    StrToBytes = StrConv(CStr(s), vbFromUnicode)
End Function

Public Function CallOptimizeAPI(payload As String) As String
    ' Makes POST request to /api/optimize endpoint
    ' Returns response body or empty string on failure
    ' Includes API key authentication header
    ' Uses ServerXMLHTTP for timeout support

    Dim http As Object
    Dim url As String
    Dim apiKey As String
    Dim t0 As Double

    On Error GoTo ErrorHandler

    If Not EnsureSelectedRentRollYearActive(True) Then
        CallOptimizeAPI = ""
        Exit Function
    End If

    url = API_BASE_URL & "/api/optimize"
    t0 = Timer

    ' Get API key from registry
    apiKey = GetAPIKey()

    ' Create HTTP object with timeout support
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    ' Configure timeouts (ms): Resolve, Connect, Send, Receive
    ' 5 sec resolve, 30 sec connect, 30 sec send, 240 sec receive (solver can take longer for edge cases)
    http.setTimeouts 5000, 30000, 30000, 240000

    ' Configure request
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json"

    ' Add API key authentication header
    If Len(apiKey) > 0 Then
        http.setRequestHeader "X-API-Key", apiKey
    End If

    DebugLog "HTTP POST /api/optimize: payload_len=" & Len(payload), True

    ' Send request
    http.send payload

    DebugLog "HTTP /api/optimize: status=" & http.Status & ", elapsed=" & Format$(ElapsedSeconds(t0), "0.00") & "s, resp_len=" & Len(http.responseText), True

    ' Check response
    If http.Status = 200 Then
        CallOptimizeAPI = http.responseText
    ElseIf http.Status = 401 Then
        ' Unauthorized - invalid API key
        MsgBox "Invalid API key." & vbCrLf & vbCrLf & _
               "Please check your API key in Settings.", _
               vbCritical, "AMI Optix - Authentication Failed"
        CallOptimizeAPI = ""
    ElseIf http.Status = 504 Or http.Status = 502 Then
        ' Gateway timeout - server might be cold starting
        MsgBox "Server is starting up (cold start). Please wait 30 seconds and try again.", _
               vbInformation, "AMI Optix"
        CallOptimizeAPI = ""
    Else
        ' Other error
        Debug.Print "API Error: " & http.Status & " - " & http.statusText
        Debug.Print "Response: " & http.responseText
        MsgBox "API Error: " & http.Status & " - " & http.statusText, _
               vbExclamation, "AMI Optix"
        CallOptimizeAPI = ""
    End If

    Exit Function

ErrorHandler:
    DebugLogError "CallOptimizeAPI"
    Debug.Print "HTTP Error: " & Err.Description
    MsgBox "Connection error: " & Err.Description & vbCrLf & vbCrLf & _
           "The server may be starting up. Please wait 30 seconds and try again.", _
           vbExclamation, "AMI Optix"
    CallOptimizeAPI = ""
End Function

Public Function CallHealthAPI() As Boolean
    ' Check if API is available
    Dim http As Object
    Dim url As String

    On Error GoTo ErrorHandler

    url = API_BASE_URL & "/healthz"  ' Fixed: was /health

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 5000, 10000, 10000, 10000
    http.Open "GET", url, False
    http.send

    CallHealthAPI = (http.Status = 200)
    Exit Function

ErrorHandler:
    CallHealthAPI = False
End Function

'-------------------------------------------------------------------------------
' JSON BUILDER
'-------------------------------------------------------------------------------

Public Function BuildAPIPayload(units As Collection, utilities As Object) As String
    ' Backward-compatible payload builder (defaults to UAP).
    BuildAPIPayload = BuildAPIPayloadV2(units, utilities, "UAP", "", 0, 0)
End Function

Public Function CallEvaluateAPI(payload As String) As String
    ' Makes POST request to /api/evaluate endpoint
    Dim http As Object
    Dim url As String
    Dim apiKey As String

    On Error GoTo ErrorHandler

    If Not EnsureSelectedRentRollYearActive(True) Then
        CallEvaluateAPI = ""
        Exit Function
    End If

    url = API_BASE_URL & "/api/evaluate"
    apiKey = GetAPIKey()

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 5000, 30000, 30000, 120000

    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json"
    If Len(apiKey) > 0 Then
        http.setRequestHeader "X-API-Key", apiKey
    End If

    http.send payload

    If http.Status = 200 Then
        CallEvaluateAPI = http.responseText
    ElseIf http.Status = 401 Then
        MsgBox "Invalid API key." & vbCrLf & vbCrLf & _
               "Please check your API key in Settings.", _
               vbCritical, "AMI Optix - Authentication Failed"
        CallEvaluateAPI = ""
    Else
        Debug.Print "API Error: " & http.Status & " - " & http.statusText
        Debug.Print "Response: " & http.responseText
        MsgBox "API Error: " & http.Status & " - " & http.statusText, _
               vbExclamation, "AMI Optix"
        CallEvaluateAPI = ""
    End If

    Exit Function

ErrorHandler:
    Debug.Print "HTTP Error: " & Err.Description
    MsgBox "Connection error: " & Err.Description, vbExclamation, "AMI Optix"
    CallEvaluateAPI = ""
End Function

Public Function CallEvaluateAPIStateless(payload As String, Optional ByRef outError As String = "") As String
    ' Makes ONE POST request to /api/evaluate endpoint without relying on server-global "active calculator" state.
    ' Intended for Fix-06d verification (no pre-activation call; selection comes from payload fields).
    Dim http As Object
    Dim url As String
    Dim apiKey As String

    outError = ""
    On Error GoTo ErrorHandler

    url = API_BASE_URL & "/api/evaluate"
    apiKey = GetAPIKey()

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 5000, 30000, 30000, 120000

    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json"

    If Len(apiKey) > 0 Then
        http.setRequestHeader "X-API-Key", apiKey
    End If

    DebugLog "HTTP POST /api/evaluate (stateless): payload_len=" & Len(payload), True

    http.send payload

    DebugLog "HTTP /api/evaluate (stateless): status=" & http.Status & ", resp_len=" & Len(http.responseText), True

    If http.Status = 200 Then
        CallEvaluateAPIStateless = http.responseText
    ElseIf http.Status = 401 Then
        outError = "Invalid API key." & vbCrLf & vbCrLf & "Please check your API key in Settings."
        CallEvaluateAPIStateless = ""
    Else
        outError = "API Error: " & http.Status & " - " & http.statusText
        If Len(http.responseText) > 0 Then outError = outError & vbCrLf & vbCrLf & http.responseText
        CallEvaluateAPIStateless = ""
    End If

    Exit Function

ErrorHandler:
    outError = "Connection error: " & Err.Description
    CallEvaluateAPIStateless = ""
End Function

Public Function CallManualCalculateAPI(payload As String) As String
    ' Makes POST request to /api/manual_calculate endpoint.
    ' This endpoint always returns computed rents/totals and diagnostics, even if the assignment is non-compliant.
    Dim http As Object
    Dim url As String
    Dim apiKey As String

    On Error GoTo ErrorHandler

    If Not EnsureSelectedRentRollYearActive(True) Then
        CallManualCalculateAPI = ""
        Exit Function
    End If

    url = API_BASE_URL & "/api/manual_calculate"
    apiKey = GetAPIKey()

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 5000, 30000, 30000, 120000

    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json"
    If Len(apiKey) > 0 Then
        http.setRequestHeader "X-API-Key", apiKey
    End If

    http.send payload

    If http.Status = 200 Then
        CallManualCalculateAPI = http.responseText
    ElseIf http.Status = 401 Then
        MsgBox "Invalid API key." & vbCrLf & vbCrLf & _
               "Please check your API key in Settings.", _
               vbCritical, "AMI Optix - Authentication Failed"
        CallManualCalculateAPI = ""
    Else
        Debug.Print "API Error: " & http.Status & " - " & http.statusText
        Debug.Print "Response: " & http.responseText
        MsgBox "API Error: " & http.Status & " - " & http.statusText, _
               vbExclamation, "AMI Optix"
        CallManualCalculateAPI = ""
    End If

    Exit Function

ErrorHandler:
    Debug.Print "HTTP Error: " & Err.Description
    MsgBox "Connection error: " & Err.Description, vbExclamation, "AMI Optix"
    CallManualCalculateAPI = ""
End Function

Public Function BuildAPIPayloadV2( _
    units As Collection, _
    utilities As Object, _
    program As String, _
    mihOption As String, _
    mihResidentialSF As Double, _
    mihMaxBandPercent As Long, _
    Optional projectOverridesJson As String = "", _
    Optional compareBaseline As Boolean = False _
) As String
    ' Builds JSON payload for API request (supports UAP/MIH).
    Dim json As String
    Dim unit As Object
    Dim i As Long

    Dim programNorm As String
    programNorm = UCase(Trim(program))
    If programNorm = "" Then programNorm = "UAP"

    json = "{"

    json = json & """program"": """ & EscapeJSON(programNorm) & """, "

    ' Rent-roll year from the ribbon dropdown. Without this field the server
    ' falls back to its global default calculator (2025), so the optimizer
    ' would price scenarios on a different year than Manual Calculate — the
    ' mixed-year bug of 2026-06. Same pattern as BuildEvaluatePayloadV2.
    Dim optimizeRentYear As Long
    optimizeRentYear = GetSelectedRentRollYearSetting()
    If optimizeRentYear > 0 Then
        json = json & """rent_roll_year"": " & CStr(optimizeRentYear) & ", "
    End If

    If programNorm = "MIH" Then
        If mihOption <> "" Then
            json = json & """mih_option"": """ & EscapeJSON(mihOption) & """, "
        End If
        If mihResidentialSF > 0 Then
            json = json & """mih_residential_sf"": " & Replace(CStr(mihResidentialSF), ",", "") & ", "
        End If
        If mihMaxBandPercent > 0 Then
            json = json & """mih_max_band_percent"": " & mihMaxBandPercent & ", "
        End If
    End If

    ' Units array
    json = json & """units"": ["
    For i = 1 To units.Count
        Set unit = units(i)
        If i > 1 Then json = json & ", "

        json = json & "{"
        json = json & """unit_id"": """ & EscapeJSON(CStr(unit("unit_id"))) & """, "
        json = json & """bedrooms"": " & unit("bedrooms") & ", "
        json = json & """net_sf"": " & unit("net_sf")

        If unit.Exists("floor") Then
            json = json & ", ""floor"": " & unit("floor")
        End If
        If unit.Exists("balcony") Then
            json = json & ", ""balcony"": " & IIf(unit("balcony"), "true", "false")
        End If
        If unit.Exists("client_ami") Then
            json = json & ", ""client_ami"": " & unit("client_ami")
        End If
        If unit.Exists("original_ami") Then
            ' Baseline snapshot of the user's true input (AMI_Optix_Baseline);
            ' the server prefers it over client_ami for the Original Scenario.
            json = json & ", ""original_ami"": " & unit("original_ami")
        End If

        json = json & "}"
    Next i
    json = json & "], "

    ' Utilities
    json = json & """utilities"": {"
    json = json & """electricity"": """ & utilities("electricity") & """, "
    json = json & """cooking"": """ & utilities("cooking") & """, "
    json = json & """heat"": """ & utilities("heat") & """, "
    json = json & """hot_water"": """ & utilities("hot_water") & """"
    json = json & "}"

    If Trim$(projectOverridesJson) <> "" Then
        json = json & ", ""project_overrides"": " & projectOverridesJson
    End If

    If compareBaseline Then
        json = json & ", ""compare_baseline"": true"
    End If

    json = json & "}"

    BuildAPIPayloadV2 = json
End Function

Public Function BuildEvaluatePayloadV2( _
    units As Collection, _
    utilities As Object, _
    program As String, _
    mihOption As String, _
    mihResidentialSF As Double, _
    mihMaxBandPercent As Long, _
    Optional rentRollYear As Long = 0, _
    Optional calculatorId As String = "" _
) As String
    ' Builds JSON payload for /api/evaluate (explicit assigned_ami per unit).
    Dim json As String
    Dim unit As Object
    Dim i As Long
    Dim programNorm As String

    programNorm = UCase(Trim(program))
    If programNorm = "" Then programNorm = "UAP"

    json = "{"

    Dim yearNorm As Long
    yearNorm = rentRollYear
    If yearNorm <= 0 Then yearNorm = GetSelectedRentRollYearSetting()

    Dim didPrefix As Boolean
    didPrefix = False
    If yearNorm > 0 Then
        json = json & """rent_roll_year"": " & CStr(yearNorm)
        didPrefix = True
    End If
    If Trim$(calculatorId) <> "" Then
        If didPrefix Then json = json & ", "
        json = json & """calculator_id"": """ & EscapeJSON(CStr(calculatorId)) & """"
        didPrefix = True
    End If
    If didPrefix Then json = json & ", "

    json = json & """program"": """ & EscapeJSON(programNorm) & """, "

    If programNorm = "MIH" Then
        If mihOption <> "" Then
            json = json & """mih_option"": """ & EscapeJSON(mihOption) & """, "
        End If
        If mihResidentialSF > 0 Then
            json = json & """mih_residential_sf"": " & Replace(CStr(mihResidentialSF), ",", "") & ", "
        End If
        If mihMaxBandPercent > 0 Then
            json = json & """mih_max_band_percent"": " & mihMaxBandPercent & ", "
        End If
    End If

    json = json & """utilities"": {"
    json = json & """electricity"": """ & utilities("electricity") & """, "
    json = json & """cooking"": """ & utilities("cooking") & """, "
    json = json & """heat"": """ & utilities("heat") & """, "
    json = json & """hot_water"": """ & utilities("hot_water") & """"
    json = json & "}, "

    json = json & """units"": ["
    For i = 1 To units.Count
        Set unit = units(i)
        If i > 1 Then json = json & ", "

        json = json & "{"
        json = json & """unit_id"": """ & EscapeJSON(CStr(unit("unit_id"))) & """, "
        json = json & """bedrooms"": " & unit("bedrooms") & ", "
        json = json & """net_sf"": " & unit("net_sf")

        If unit.Exists("floor") Then
            json = json & ", ""floor"": " & unit("floor")
        End If
        If unit.Exists("balcony") Then
            json = json & ", ""balcony"": " & IIf(unit("balcony"), "true", "false")
        End If

        If unit.Exists("client_ami") Then
            json = json & ", ""assigned_ami"": " & unit("client_ami")
        End If

        json = json & "}"
    Next i
    json = json & "]"

    json = json & "}"

    BuildEvaluatePayloadV2 = json
End Function

Private Function EscapeJSON(str As String) As String
    ' Escape special characters in JSON strings
    Dim result As String
    result = str
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCr, "\r")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")
    EscapeJSON = result
End Function

'-------------------------------------------------------------------------------
' JSON PARSER (Works with both 32-bit and 64-bit Office)
' This is a custom parser that doesn't rely on ScriptControl
'-------------------------------------------------------------------------------

Public Function ParseJSON(jsonString As String) As Object
    ' Parses JSON response into Dictionary/Collection structure
    ' Compatible with 64-bit Office (no ScriptControl dependency)

    Dim result As Object
    Dim pos As Long

    On Error GoTo ErrorHandler

    pos = 1
    SkipWhitespace jsonString, pos

    If Mid(jsonString, pos, 1) = "{" Then
        Set result = ParseObject(jsonString, pos)
    ElseIf Mid(jsonString, pos, 1) = "[" Then
        Set result = ParseArray(jsonString, pos)
    Else
        Set result = Nothing
    End If

    Set ParseJSON = result
    Exit Function

ErrorHandler:
    Debug.Print "JSON Parse Error at position " & pos & ": " & Err.Description
    Set ParseJSON = Nothing
End Function

Private Function ParseObject(jsonString As String, ByRef pos As Long) As Object
    ' Parses a JSON object into a Dictionary
    Dim dict As Object
    Dim key As String
    Dim val As Variant
    Dim peekChar As String

    Set dict = CreateObject("Scripting.Dictionary")

    ' Skip opening brace
    pos = pos + 1
    SkipWhitespace jsonString, pos

    ' Empty object
    If Mid(jsonString, pos, 1) = "}" Then
        pos = pos + 1
        Set ParseObject = dict
        Exit Function
    End If

    Do
        SkipWhitespace jsonString, pos

        ' Parse key
        key = ParseString(jsonString, pos)

        SkipWhitespace jsonString, pos

        ' Skip colon
        If Mid(jsonString, pos, 1) = ":" Then
            pos = pos + 1
        End If

        SkipWhitespace jsonString, pos

        ' Parse value - peek at type first to handle object assignment correctly
        peekChar = Mid(jsonString, pos, 1)

        If peekChar = "{" Then
            ' Object - use Set
            Set val = ParseObject(jsonString, pos)
            Set dict(key) = val
        ElseIf peekChar = "[" Then
            ' Array - use Set
            Set val = ParseArray(jsonString, pos)
            Set dict(key) = val
        Else
            ' Scalar (string, number, boolean, null)
            val = ParseScalarValue(jsonString, pos)
            dict(key) = val
        End If

        SkipWhitespace jsonString, pos

        ' Check for comma or end
        If Mid(jsonString, pos, 1) = "," Then
            pos = pos + 1
        ElseIf Mid(jsonString, pos, 1) = "}" Then
            pos = pos + 1
            Exit Do
        Else
            Exit Do
        End If
    Loop

    Set ParseObject = dict
End Function

Private Function ParseArray(jsonString As String, ByRef pos As Long) As Collection
    ' Parses a JSON array into a Collection
    Dim coll As Collection
    Dim val As Variant
    Dim peekChar As String

    Set coll = New Collection

    ' Skip opening bracket
    pos = pos + 1
    SkipWhitespace jsonString, pos

    ' Empty array
    If Mid(jsonString, pos, 1) = "]" Then
        pos = pos + 1
        Set ParseArray = coll
        Exit Function
    End If

    Do
        SkipWhitespace jsonString, pos

        ' Parse value - peek at type first to handle object assignment correctly
        peekChar = Mid(jsonString, pos, 1)

        If peekChar = "{" Then
            ' Object - use Set
            Set val = ParseObject(jsonString, pos)
            coll.Add val
        ElseIf peekChar = "[" Then
            ' Nested array - use Set
            Set val = ParseArray(jsonString, pos)
            coll.Add val
        Else
            ' Scalar (string, number, boolean, null)
            val = ParseScalarValue(jsonString, pos)
            coll.Add val
        End If

        SkipWhitespace jsonString, pos

        ' Check for comma or end
        If Mid(jsonString, pos, 1) = "," Then
            pos = pos + 1
        ElseIf Mid(jsonString, pos, 1) = "]" Then
            pos = pos + 1
            Exit Do
        Else
            Exit Do
        End If
    Loop

    Set ParseArray = coll
End Function

Private Function ParseScalarValue(jsonString As String, ByRef pos As Long) As Variant
    ' Parses scalar JSON values (string, number, boolean, null)
    ' Objects and arrays are handled separately in ParseObject/ParseArray
    Dim char As String

    SkipWhitespace jsonString, pos
    char = Mid(jsonString, pos, 1)

    If char = """" Then
        ParseScalarValue = ParseString(jsonString, pos)
    ElseIf char = "t" Then
        ' true
        pos = pos + 4
        ParseScalarValue = True
    ElseIf char = "f" Then
        ' false
        pos = pos + 5
        ParseScalarValue = False
    ElseIf char = "n" Then
        ' null
        pos = pos + 4
        ParseScalarValue = Null
    Else
        ' Number
        ParseScalarValue = ParseNumber(jsonString, pos)
    End If
End Function

Private Function ParseString(jsonString As String, ByRef pos As Long) As String
    ' Parses a JSON string
    Dim result As String
    Dim char As String
    Dim nextChar As String

    result = ""

    ' Skip opening quote
    pos = pos + 1

    Do While pos <= Len(jsonString)
        char = Mid(jsonString, pos, 1)

        If char = """" Then
            pos = pos + 1
            Exit Do
        ElseIf char = "\" Then
            pos = pos + 1
            nextChar = Mid(jsonString, pos, 1)
            Select Case nextChar
                Case """"
                    result = result & """"
                Case "\"
                    result = result & "\"
                Case "/"
                    result = result & "/"
                Case "b"
                    result = result & vbBack
                Case "f"
                    result = result & vbFormFeed
                Case "n"
                    result = result & vbLf
                Case "r"
                    result = result & vbCr
                Case "t"
                    result = result & vbTab
                Case "u"
                    ' Unicode escape - skip for now
                    pos = pos + 4
                Case Else
                    result = result & nextChar
            End Select
            pos = pos + 1
        Else
            result = result & char
            pos = pos + 1
        End If
    Loop

    ParseString = result
End Function

Private Function ParseNumber(jsonString As String, ByRef pos As Long) As Double
    ' Parses a JSON number
    Dim startPos As Long
    Dim char As String
    Dim numStr As String

    startPos = pos

    Do While pos <= Len(jsonString)
        char = Mid(jsonString, pos, 1)
        ' Check for valid number characters (digits, decimal, exponent, signs)
        ' Using explicit checks to avoid VBA Like operator ambiguity with hyphen
        If (char >= "0" And char <= "9") Or char = "." Or char = "e" Or _
           char = "E" Or char = "+" Or char = "-" Then
            pos = pos + 1
        Else
            Exit Do
        End If
    Loop

    numStr = Mid(jsonString, startPos, pos - startPos)

    ' Handle empty or invalid number string
    If Len(numStr) = 0 Then
        ParseNumber = 0
    Else
        ' Replace any locale-specific decimal separator issues
        numStr = Replace(numStr, ",", ".")
        On Error Resume Next
        ParseNumber = CDbl(numStr)
        If Err.Number <> 0 Then
            ParseNumber = 0
            Err.Clear
        End If
        On Error GoTo 0
    End If
End Function

Private Sub SkipWhitespace(jsonString As String, ByRef pos As Long)
    ' Skips whitespace characters
    Dim char As String

    Do While pos <= Len(jsonString)
        char = Mid(jsonString, pos, 1)
        If char = " " Or char = vbCr Or char = vbLf Or char = vbTab Then
            pos = pos + 1
        Else
            Exit Do
        End If
    Loop
End Sub
