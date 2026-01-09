Attribute VB_Name = "AMI_Optix_DataReader"
'===============================================================================
' AMI OPTIX - Data Reader Module
' Reads unit data from any workbook using fuzzy header matching
'===============================================================================
Option Explicit

' Column indices found by fuzzy matching
Private m_UnitIdCol As Long
Private m_BedroomsCol As Long
Private m_NetSFCol As Long
Private m_FloorCol As Long
Private m_AMICol As Long
Private m_BalconyCol As Long
Private m_HeaderRow As Long
Private m_DataSheet As Worksheet

'-------------------------------------------------------------------------------
' MAIN DATA READER
'-------------------------------------------------------------------------------

Public Function ReadUnitData() As Collection
    ' Reads unit data from active workbook
    ' Returns Collection of Dictionary objects (one per unit)

    Dim units As New Collection
    Dim ws As Worksheet
    Dim headerRow As Long
    Dim lastRow As Long
    Dim i As Long
    Dim unit As Object

    On Error GoTo ErrorHandler

    ' Try to find data in active sheet first, then search all sheets
    Set ws = FindDataSheet()

    If ws Is Nothing Then
        Set ReadUnitData = Nothing
        Exit Function
    End If

    Set m_DataSheet = ws

    ' Find header row and column mappings
    headerRow = FindHeaderRow(ws)
    If headerRow = 0 Then
        Set ReadUnitData = Nothing
        Exit Function
    End If

    m_HeaderRow = headerRow

    ' Map columns using fuzzy matching
    If Not MapColumns(ws, headerRow) Then
        Set ReadUnitData = Nothing
        Exit Function
    End If

    ' Find last row with data
    lastRow = FindLastDataRow(ws, headerRow)

    ' Read each unit
    For i = headerRow + 1 To lastRow
        Set unit = ReadUnitRow(ws, i)
        If Not unit Is Nothing Then
            units.Add unit
        End If
    Next i

    Set ReadUnitData = units
    Exit Function

ErrorHandler:
    Debug.Print "ReadUnitData Error: " & Err.Description
    Set ReadUnitData = Nothing
End Function

'-------------------------------------------------------------------------------
' FIND DATA SHEET
'-------------------------------------------------------------------------------

Private Function FindDataSheet() As Worksheet
    ' Finds the sheet containing unit data

    Dim ws As Worksheet
    Dim preferredNames As Variant
    Dim i As Long

    ' First try active sheet
    If HasUnitDataHeaders(ActiveSheet) Then
        Set FindDataSheet = ActiveSheet
        Exit Function
    End If

    ' Try preferred sheet names
    preferredNames = Array("UAP", "PROJECT WORKSHEET", "RentRoll", "Units", "Sheet1", "Data")

    For i = LBound(preferredNames) To UBound(preferredNames)
        On Error Resume Next
        Set ws = ActiveWorkbook.Worksheets(CStr(preferredNames(i)))
        On Error GoTo 0

        If Not ws Is Nothing Then
            If HasUnitDataHeaders(ws) Then
                Set FindDataSheet = ws
                Exit Function
            End If
        End If
        Set ws = Nothing
    Next i

    ' Search all sheets
    For Each ws In ActiveWorkbook.Worksheets
        If HasUnitDataHeaders(ws) Then
            Set FindDataSheet = ws
            Exit Function
        End If
    Next ws

    Set FindDataSheet = Nothing
End Function

Private Function HasUnitDataHeaders(ws As Worksheet) As Boolean
    ' Quick check if sheet has recognizable unit data headers
    Dim headerRow As Long
    headerRow = FindHeaderRow(ws)
    HasUnitDataHeaders = (headerRow > 0)
End Function

'-------------------------------------------------------------------------------
' HEADER ROW DETECTION
'-------------------------------------------------------------------------------

Private Function FindHeaderRow(ws As Worksheet) As Long
    ' Finds the row containing column headers
    ' Looks for rows with at least 3 recognized headers

    Dim row As Long
    Dim col As Long
    Dim cellValue As String
    Dim matches As Long
    Dim maxRow As Long

    maxRow = Application.Min(50, ws.UsedRange.Rows.Count)  ' Only check first 50 rows

    For row = 1 To maxRow
        matches = 0

        For col = 1 To Application.Min(20, ws.UsedRange.Columns.Count)
            cellValue = UCase(Trim(CStr(ws.Cells(row, col).Value)))

            If IsUnitIdHeader(cellValue) Then matches = matches + 1
            If IsBedroomsHeader(cellValue) Then matches = matches + 1
            If IsNetSFHeader(cellValue) Then matches = matches + 1
            If IsFloorHeader(cellValue) Then matches = matches + 1
            If IsAMIHeader(cellValue) Then matches = matches + 1
        Next col

        ' Need at least 3 matches (unit_id, bedrooms, net_sf)
        If matches >= 3 Then
            FindHeaderRow = row
            Exit Function
        End If
    Next row

    FindHeaderRow = 0
End Function

'-------------------------------------------------------------------------------
' FUZZY HEADER MATCHING
'-------------------------------------------------------------------------------

Private Function MapColumns(ws As Worksheet, headerRow As Long) As Boolean
    ' Maps columns using fuzzy matching
    ' Returns True if required columns found

    Dim col As Long
    Dim cellValue As String
    Dim maxCol As Long

    ' Reset
    m_UnitIdCol = 0
    m_BedroomsCol = 0
    m_NetSFCol = 0
    m_FloorCol = 0
    m_AMICol = 0
    m_BalconyCol = 0

    maxCol = Application.Min(30, ws.UsedRange.Columns.Count)

    For col = 1 To maxCol
        cellValue = UCase(Trim(CStr(ws.Cells(headerRow, col).Value)))

        ' Unit ID
        If m_UnitIdCol = 0 And IsUnitIdHeader(cellValue) Then
            m_UnitIdCol = col
        End If

        ' Bedrooms
        If m_BedroomsCol = 0 And IsBedroomsHeader(cellValue) Then
            m_BedroomsCol = col
        End If

        ' Net SF
        If m_NetSFCol = 0 And IsNetSFHeader(cellValue) Then
            m_NetSFCol = col
        End If

        ' Floor
        If m_FloorCol = 0 And IsFloorHeader(cellValue) Then
            m_FloorCol = col
        End If

        ' AMI (for writing results back)
        If m_AMICol = 0 And IsAMIHeader(cellValue) Then
            m_AMICol = col
        End If

        ' Balcony
        If m_BalconyCol = 0 And IsBalconyHeader(cellValue) Then
            m_BalconyCol = col
        End If
    Next col

    ' Required columns: unit_id, bedrooms, net_sf
    MapColumns = (m_UnitIdCol > 0 And m_BedroomsCol > 0 And m_NetSFCol > 0)

    ' Debug output
    Debug.Print "Column Mapping:"
    Debug.Print "  Unit ID: " & m_UnitIdCol
    Debug.Print "  Bedrooms: " & m_BedroomsCol
    Debug.Print "  Net SF: " & m_NetSFCol
    Debug.Print "  Floor: " & m_FloorCol
    Debug.Print "  AMI: " & m_AMICol
    Debug.Print "  Balcony: " & m_BalconyCol
End Function

'-------------------------------------------------------------------------------
' HEADER MATCHING FUNCTIONS
'-------------------------------------------------------------------------------

Private Function IsUnitIdHeader(header As String) As Boolean
    Dim patterns As Variant
    patterns = Array("APT #", "APT", "UNIT", "UNIT ID", "APARTMENT", "UNIT NO", "APT NUMBER")
    IsUnitIdHeader = MatchesAny(header, patterns)
End Function

Private Function IsBedroomsHeader(header As String) As Boolean
    Dim patterns As Variant
    patterns = Array("BED", "BEDS", "BEDROOMS", "NUMBER OF BEDROOMS", "BR", "BEDROOM")
    IsBedroomsHeader = MatchesAny(header, patterns)
End Function

Private Function IsNetSFHeader(header As String) As Boolean
    ' More specific patterns to avoid false positives
    ' Removed "SF" alone (too broad - could match STAFF, SATISFACTION, etc.)
    ' Removed "AREA" alone (too broad - could match PROJECT AREA, etc.)
    Dim patterns As Variant
    patterns = Array("NET SF", "NETSF", "NET S.F.", "SQFT", "SQ FT", "SQ. FT.", _
                     "NET SQUARE FEET", "SQUARE FEET", "NET SQFT", "UNIT SF", _
                     "UNIT SQFT", "APT SF", "APT SQFT", "RENTABLE SF", "RENTABLE SQFT")
    IsNetSFHeader = MatchesAny(header, patterns)
End Function

Private Function IsFloorHeader(header As String) As Boolean
    Dim patterns As Variant
    patterns = Array("FLOOR", "STORY", "LEVEL", "CONSTRUCTION STORY", "MARKETING STORY", "FLR")
    IsFloorHeader = MatchesAny(header, patterns)
End Function

Private Function IsAMIHeader(header As String) As Boolean
    ' Match AMI but NOT "AMI AFTER" or "AMI FOR 35"
    If InStr(header, "AMI AFTER") > 0 Then
        IsAMIHeader = False
        Exit Function
    End If

    If InStr(header, "AMI FOR 35") > 0 Then
        IsAMIHeader = False
        Exit Function
    End If

    Dim patterns As Variant
    patterns = Array("AMI", "AFFORDABILITY", "AFF %", "AFF", "AMI BAND")
    IsAMIHeader = MatchesAny(header, patterns)
End Function

Private Function IsBalconyHeader(header As String) As Boolean
    Dim patterns As Variant
    patterns = Array("BALCONY", "TERRACE", "OUTDOOR", "BALCONIES")
    IsBalconyHeader = MatchesAny(header, patterns)
End Function

Private Function MatchesAny(header As String, patterns As Variant) As Boolean
    ' Matches header against patterns using smart matching rules:
    ' 1. Exact match always succeeds
    ' 2. Contains match only if pattern is at word boundary

    Dim i As Long
    Dim pattern As String
    Dim pos As Long
    Dim charBefore As String
    Dim charAfter As String

    For i = LBound(patterns) To UBound(patterns)
        pattern = UCase(CStr(patterns(i)))

        ' Exact match
        If header = pattern Then
            MatchesAny = True
            Exit Function
        End If

        ' Contains match - but only at word boundaries to avoid false positives
        ' e.g., "NET SF" in "UNIT NET SF" is OK, but "SF" in "STAFF" is not
        pos = InStr(header, pattern)
        If pos > 0 Then
            ' Check character before pattern (should be start or space/punctuation)
            If pos = 1 Then
                charBefore = " "  ' Start of string is OK
            Else
                charBefore = Mid(header, pos - 1, 1)
            End If

            ' Check character after pattern (should be end or space/punctuation)
            If pos + Len(pattern) > Len(header) Then
                charAfter = " "  ' End of string is OK
            Else
                charAfter = Mid(header, pos + Len(pattern), 1)
            End If

            ' Word boundary check: before and after should not be alphanumeric
            If Not (charBefore Like "[A-Z0-9]") And Not (charAfter Like "[A-Z0-9]") Then
                MatchesAny = True
                Exit Function
            End If
        End If
    Next i

    MatchesAny = False
End Function

'-------------------------------------------------------------------------------
' DATA READING
'-------------------------------------------------------------------------------

Private Function FindLastDataRow(ws As Worksheet, headerRow As Long) As Long
    ' Find last row with data in the unit_id column
    Dim lastRow As Long

    If m_UnitIdCol > 0 Then
        lastRow = ws.Cells(ws.Rows.Count, m_UnitIdCol).End(xlUp).row
    Else
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    End If

    ' Make sure we're past the header
    If lastRow <= headerRow Then
        lastRow = headerRow
    End If

    FindLastDataRow = lastRow
End Function

Private Function ReadUnitRow(ws As Worksheet, row As Long) As Object
    ' Reads a single unit row into a Dictionary
    Dim unit As Object
    Dim unitId As String
    Dim bedrooms As Variant
    Dim netSF As Variant

    On Error GoTo ErrorHandler

    ' Read required fields
    unitId = Trim(CStr(ws.Cells(row, m_UnitIdCol).Value))
    bedrooms = ws.Cells(row, m_BedroomsCol).Value
    netSF = ws.Cells(row, m_NetSFCol).Value

    ' Skip empty rows
    If unitId = "" Or unitId = "0" Then
        Set ReadUnitRow = Nothing
        Exit Function
    End If

    ' Skip non-numeric bedrooms/sqft
    If Not IsNumeric(bedrooms) Or Not IsNumeric(netSF) Then
        Set ReadUnitRow = Nothing
        Exit Function
    End If

    ' Skip zero/negative values
    If CDbl(bedrooms) < 0 Or CDbl(netSF) <= 0 Then
        Set ReadUnitRow = Nothing
        Exit Function
    End If

    ' Create unit dictionary
    Set unit = CreateObject("Scripting.Dictionary")
    unit("unit_id") = unitId
    unit("bedrooms") = CDbl(bedrooms)
    unit("net_sf") = CDbl(netSF)
    unit("row") = row  ' Store row number for writing back

    ' Optional: Floor
    If m_FloorCol > 0 Then
        Dim floor As Variant
        floor = ws.Cells(row, m_FloorCol).Value
        If IsNumeric(floor) Then
            unit("floor") = CDbl(floor)
        End If
    End If

    ' Optional: Balcony
    If m_BalconyCol > 0 Then
        Dim balcony As Variant
        balcony = ws.Cells(row, m_BalconyCol).Value
        If balcony = "Yes" Or balcony = "Y" Or balcony = 1 Or balcony = True Then
            unit("balcony") = True
        Else
            unit("balcony") = False
        End If
    End If

    Set ReadUnitRow = unit
    Exit Function

ErrorHandler:
    Set ReadUnitRow = Nothing
End Function

'-------------------------------------------------------------------------------
' PUBLIC ACCESSORS
'-------------------------------------------------------------------------------

Public Function GetAMIColumn() As Long
    GetAMIColumn = m_AMICol
End Function

Public Function GetHeaderRow() As Long
    GetHeaderRow = m_HeaderRow
End Function

Public Function GetDataSheet() As Worksheet
    Set GetDataSheet = m_DataSheet
End Function
