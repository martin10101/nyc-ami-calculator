Attribute VB_Name = "AMI_Optix_API"
'===============================================================================
' AMI OPTIX - API Communication Module
' Handles HTTP requests to the optimization API and JSON parsing
'===============================================================================
Option Explicit

'-------------------------------------------------------------------------------
' API CALL
'-------------------------------------------------------------------------------

Public Function CallOptimizeAPI(payload As String) As String
    ' Makes POST request to /api/optimize endpoint
    ' Returns response body or empty string on failure

    Dim http As Object
    Dim url As String

    On Error GoTo ErrorHandler

    url = API_BASE_URL & "/api/optimize"

    ' Create HTTP object
    Set http = CreateObject("MSXML2.XMLHTTP")

    ' Configure request
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json"

    ' Set timeout (note: XMLHTTP doesn't have direct timeout, we handle via error)
    ' For longer timeout, use ServerXMLHTTP instead

    ' Send request
    http.send payload

    ' Check response
    If http.Status = 200 Then
        CallOptimizeAPI = http.responseText
    ElseIf http.Status = 504 Or http.Status = 502 Then
        ' Gateway timeout - server might be cold starting
        MsgBox "Server is starting up (cold start). Please wait 30 seconds and try again.", _
               vbInformation, "AMI Optix"
        CallOptimizeAPI = ""
    Else
        ' Other error
        Debug.Print "API Error: " & http.Status & " - " & http.statusText
        Debug.Print "Response: " & http.responseText
        CallOptimizeAPI = ""
    End If

    Exit Function

ErrorHandler:
    Debug.Print "HTTP Error: " & Err.Description
    CallOptimizeAPI = ""
End Function

Public Function CallHealthAPI() As Boolean
    ' Check if API is available
    Dim http As Object
    Dim url As String

    On Error GoTo ErrorHandler

    url = API_BASE_URL & "/health"

    Set http = CreateObject("MSXML2.XMLHTTP")
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
    ' Builds JSON payload for API request
    Dim json As String
    Dim unit As Object
    Dim i As Long

    ' Start JSON object
    json = "{"

    ' Add units array
    json = json & """units"": ["

    For i = 1 To units.Count
        Set unit = units(i)

        If i > 1 Then json = json & ", "

        json = json & "{"
        json = json & """unit_id"": """ & EscapeJSON(CStr(unit("unit_id"))) & """, "
        json = json & """bedrooms"": " & unit("bedrooms") & ", "
        json = json & """net_sf"": " & unit("net_sf")

        ' Optional fields
        If unit.Exists("floor") Then
            json = json & ", ""floor"": " & unit("floor")
        End If

        If unit.Exists("balcony") Then
            json = json & ", ""balcony"": " & IIf(unit("balcony"), "true", "false")
        End If

        json = json & "}"
    Next i

    json = json & "], "

    ' Add utilities
    json = json & """utilities"": {"
    json = json & """electricity"": """ & utilities("electricity") & """, "
    json = json & """cooking"": """ & utilities("cooking") & """, "
    json = json & """heat"": """ & utilities("heat") & """, "
    json = json & """hot_water"": """ & utilities("hot_water") & """"
    json = json & "}"

    ' Close JSON object
    json = json & "}"

    BuildAPIPayload = json
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
' JSON PARSER (Simple implementation for our specific response structure)
'-------------------------------------------------------------------------------

Public Function ParseJSON(jsonString As String) As Object
    ' Parses JSON response into Dictionary/Collection structure
    ' This is a simplified parser for our known response format

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    On Error GoTo ErrorHandler

    ' Use ScriptControl for JSON parsing (Windows only)
    Dim sc As Object
    Set sc = CreateObject("MSScriptControl.ScriptControl")
    sc.Language = "JScript"

    ' Add JSON parse helper
    sc.AddCode "function parseJSON(str) { return eval('(' + str + ')'); }"
    sc.AddCode "function getKeys(obj) { var keys = []; for(var k in obj) keys.push(k); return keys; }"
    sc.AddCode "function getValue(obj, key) { return obj[key]; }"
    sc.AddCode "function isArray(obj) { return Object.prototype.toString.call(obj) === '[object Array]'; }"
    sc.AddCode "function getLength(arr) { return arr.length; }"
    sc.AddCode "function getItem(arr, idx) { return arr[idx]; }"

    ' Parse JSON
    Dim jsObj As Object
    Set jsObj = sc.Run("parseJSON", jsonString)

    ' Convert to VBA structure
    Set result = JSObjectToDict(sc, jsObj)

    Set ParseJSON = result
    Exit Function

ErrorHandler:
    Debug.Print "JSON Parse Error: " & Err.Description
    Set ParseJSON = Nothing
End Function

Private Function JSObjectToDict(sc As Object, jsObj As Object) As Object
    ' Recursively convert JavaScript object to VBA Dictionary
    Dim result As Object
    Dim keys As Variant
    Dim i As Long
    Dim key As String
    Dim val As Variant

    On Error GoTo ErrorHandler

    ' Check if array
    If sc.Run("isArray", jsObj) Then
        Set result = New Collection
        Dim arrLen As Long
        arrLen = sc.Run("getLength", jsObj)

        For i = 0 To arrLen - 1
            val = sc.Run("getItem", jsObj, i)
            If IsObject(val) Then
                result.Add JSObjectToDict(sc, val)
            Else
                result.Add val
            End If
        Next i
    Else
        Set result = CreateObject("Scripting.Dictionary")

        ' Get keys
        keys = sc.Run("getKeys", jsObj)
        Dim keyLen As Long
        keyLen = sc.Run("getLength", keys)

        For i = 0 To keyLen - 1
            key = sc.Run("getItem", keys, i)
            val = sc.Run("getValue", jsObj, key)

            If IsObject(val) Then
                result(key) = JSObjectToDict(sc, val)
            Else
                result(key) = val
            End If
        Next i
    End If

    Set JSObjectToDict = result
    Exit Function

ErrorHandler:
    Set JSObjectToDict = CreateObject("Scripting.Dictionary")
End Function
