Attribute VB_Name = "AMI_Optix_API"
'===============================================================================
' AMI OPTIX - API Communication Module
' Handles HTTP requests to the optimization API and JSON parsing
'
' Compatible with both 32-bit and 64-bit Office
'===============================================================================
Option Explicit

'-------------------------------------------------------------------------------
' API CALL
'-------------------------------------------------------------------------------

Public Function CallOptimizeAPI(payload As String) As String
    ' Makes POST request to /api/optimize endpoint
    ' Returns response body or empty string on failure
    ' Includes API key authentication header
    ' Uses ServerXMLHTTP for timeout support

    Dim http As Object
    Dim url As String
    Dim apiKey As String

    On Error GoTo ErrorHandler

    url = API_BASE_URL & "/api/optimize"

    ' Get API key from registry
    apiKey = GetAPIKey()

    ' Create HTTP object with timeout support
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    ' Configure timeouts (ms): Resolve, Connect, Send, Receive
    ' 5 sec resolve, 30 sec connect, 30 sec send, 120 sec receive (for cold start)
    http.setTimeouts 5000, 30000, 30000, 120000

    ' Configure request
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json"

    ' Add API key authentication header
    If Len(apiKey) > 0 Then
        http.setRequestHeader "X-API-Key", apiKey
    End If

    ' Send request
    http.send payload

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
