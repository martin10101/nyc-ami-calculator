Attribute VB_Name = "AMI_Optix_RentCalcTables"
'===============================================================================
' AMI OPTIX - Table-driven Local Rent Calculation (Fix-06c)
'
' Loads normalized CSV cache tables (per-user) and computes net rent locally
' for the Manual Working Copy refresh path only.
'===============================================================================
Option Explicit

Private Const RENT_CALC_ERR_MISSING As Long = vbObjectError + 630
Private Const RENT_CALC_ERR_IO As Long = vbObjectError + 631

Private m_CacheYear As Long
Private m_RentLimits As Object ' Scripting.Dictionary: "Program|Bedrooms|AMI" -> GrossRent
Private m_UtilityAllowances As Object ' Scripting.Dictionary: "UtilityType|UtilityVariant|Bedrooms" -> Allowance

Public Function LoadRentLimitsCacheToDict(year As Long) As Object
    On Error GoTo Fail

    Dim cacheFolder As String
    cacheFolder = GetRentTablesCacheFolder(year)

    Dim path As String
    path = cacheFolder & "\rent_limits.csv"

    If Dir$(path) = "" Then
        Err.Raise RENT_CALC_ERR_IO, "AMI_Optix_RentCalcTables.LoadRentLimitsCacheToDict", _
                  "rent_limits.csv not found. Refresh rent tables cache for the selected year." & vbCrLf & vbCrLf & _
                  "Year: " & CStr(year) & vbCrLf & _
                  "Expected: " & path
    End If

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    Dim f As Integer
    f = FreeFile
    Open path For Input As #f

    Dim line As String
    Dim isHeader As Boolean
    isHeader = True

    Do While Not EOF(f)
        Line Input #f, line
        If isHeader Then
            isHeader = False
            GoTo NextLine
        End If

        line = Trim$(line)
        If line = "" Then GoTo NextLine

        Dim fields As Variant
        fields = ParseCsvLine(line)
        If IsEmpty(fields) Then GoTo NextLine
        If UBound(fields) < 4 Then GoTo NextLine

        Dim program As String
        program = UCase$(Trim$(CStr(fields(1))))

        Dim bed As String
        bed = Trim$(CStr(fields(2)))

        Dim amiKey As String
        amiKey = CanonAmiKey(ParseInvariantDouble(CStr(fields(3))))

        Dim gross As Double
        gross = ParseInvariantDouble(CStr(fields(4)))

        Dim key As String
        key = program & "|" & bed & "|" & amiKey

        If d.Exists(key) Then
            If CDbl(d(key)) <> gross Then
                Err.Raise RENT_CALC_ERR_IO, "AMI_Optix_RentCalcTables.LoadRentLimitsCacheToDict", _
                          "Duplicate rent limits key with different values in cache." & vbCrLf & vbCrLf & _
                          "Key: " & key & vbCrLf & _
                          "ValueA: " & CStr(d(key)) & vbCrLf & _
                          "ValueB: " & CStr(gross) & vbCrLf & _
                          "File: " & path
            End If
        Else
            d(key) = gross
        End If

NextLine:
    Loop

    Close #f

    Set m_RentLimits = d
    m_CacheYear = year
    Set LoadRentLimitsCacheToDict = d
    Exit Function

Fail:
    On Error Resume Next
    If f <> 0 Then Close #f
    On Error GoTo 0
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function LoadUtilityAllowancesCacheToDict(year As Long) As Object
    On Error GoTo Fail

    Dim cacheFolder As String
    cacheFolder = GetRentTablesCacheFolder(year)

    Dim path As String
    path = cacheFolder & "\utility_allowances.csv"

    If Dir$(path) = "" Then
        Err.Raise RENT_CALC_ERR_IO, "AMI_Optix_RentCalcTables.LoadUtilityAllowancesCacheToDict", _
                  "utility_allowances.csv not found. Refresh rent tables cache for the selected year." & vbCrLf & vbCrLf & _
                  "Year: " & CStr(year) & vbCrLf & _
                  "Expected: " & path
    End If

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    Dim f As Integer
    f = FreeFile
    Open path For Input As #f

    Dim line As String
    Dim isHeader As Boolean
    isHeader = True

    Do While Not EOF(f)
        Line Input #f, line
        If isHeader Then
            isHeader = False
            GoTo NextLine
        End If

        line = Trim$(line)
        If line = "" Then GoTo NextLine

        Dim fields As Variant
        fields = ParseCsvLine(line)
        If IsEmpty(fields) Then GoTo NextLine
        If UBound(fields) < 4 Then GoTo NextLine

        Dim uType As String
        uType = LCase$(Trim$(CStr(fields(1))))

        Dim uVar As String
        uVar = LCase$(Trim$(CStr(fields(2))))

        Dim bed As String
        bed = Trim$(CStr(fields(3)))

        Dim amt As Double
        amt = ParseInvariantDouble(CStr(fields(4)))

        Dim key As String
        key = uType & "|" & uVar & "|" & bed

        If d.Exists(key) Then
            If CDbl(d(key)) <> amt Then
                Err.Raise RENT_CALC_ERR_IO, "AMI_Optix_RentCalcTables.LoadUtilityAllowancesCacheToDict", _
                          "Duplicate utility allowance key with different values in cache." & vbCrLf & vbCrLf & _
                          "Key: " & key & vbCrLf & _
                          "ValueA: " & CStr(d(key)) & vbCrLf & _
                          "ValueB: " & CStr(amt) & vbCrLf & _
                          "File: " & path
            End If
        Else
            d(key) = amt
        End If

NextLine:
    Loop

    Close #f

    Set m_UtilityAllowances = d
    m_CacheYear = year
    Set LoadUtilityAllowancesCacheToDict = d
    Exit Function

Fail:
    On Error Resume Next
    If f <> 0 Then Close #f
    On Error GoTo 0
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function ComputeNetRent(year As Long, program As String, bedrooms As Variant, ami As Double, utilitiesDict As Object, Optional unitId As String = "") As Object
    ' Returns a Dictionary: gross_rent, allowance_total, monthly_rent, annual_rent, allowances(Collection)
    On Error GoTo Fail

    If m_RentLimits Is Nothing Or m_CacheYear <> year Then
        Set m_RentLimits = LoadRentLimitsCacheToDict(year)
    End If
    If m_UtilityAllowances Is Nothing Or m_CacheYear <> year Then
        Set m_UtilityAllowances = LoadUtilityAllowancesCacheToDict(year)
    End If

    Dim programNorm As String
    programNorm = UCase$(Trim$(CStr(program)))
    If programNorm = "" Then programNorm = "UAP"

    Dim amiNorm As Double
    amiNorm = CDbl(ami)
    If amiNorm > 2# Then amiNorm = amiNorm / 100#

    Dim amiKey As String
    amiKey = CanonAmiKey(amiNorm)

    Dim bedLabel As String
    bedLabel = BedroomLabelFromCount(bedrooms)

    Dim cacheFolder As String
    cacheFolder = GetRentTablesCacheFolder(year)

    Dim rentKey As String
    rentKey = programNorm & "|" & bedLabel & "|" & amiKey

    Dim gross As Double
    If Not m_RentLimits.Exists(rentKey) Then
        RaiseMissingKey "rent_limits.csv", year, cacheFolder, unitId, rentKey
    End If
    gross = CDbl(m_RentLimits(rentKey))

    Dim categories As Variant
    categories = Array("electricity", "cooking", "heat", "hot_water")

    Dim allowancesArr As Collection
    Set allowancesArr = New Collection

    Dim totalAllowance As Double
    totalAllowance = 0#

    Dim c As Long
    For c = LBound(categories) To UBound(categories)
        Dim cat As String
        cat = CStr(categories(c))

        Dim selectionKey As String
        selectionKey = "na"
        On Error Resume Next
        If Not utilitiesDict Is Nothing Then
            If utilitiesDict.Exists(cat) Then selectionKey = CStr(utilitiesDict(cat))
        End If
        On Error GoTo Fail

        selectionKey = LCase$(Trim$(selectionKey))

        Dim amt As Double
        amt = 0#

        Dim lookupKey As String
        lookupKey = ""

        If selectionKey <> "" And selectionKey <> "na" Then
            lookupKey = cat & "|" & selectionKey & "|" & bedLabel
            If Not m_UtilityAllowances.Exists(lookupKey) Then
                RaiseMissingKey "utility_allowances.csv", year, cacheFolder, unitId, lookupKey
            End If
            amt = CDbl(m_UtilityAllowances(lookupKey))
        End If

        Dim item As Object
        Set item = CreateObject("Scripting.Dictionary")
        item("category") = cat
        item("label") = FormatUtilityTypeGuideline(selectionKey, cat)
        item("amount") = Round(amt, 2)
        allowancesArr.Add item

        totalAllowance = totalAllowance + amt
    Next c

    Dim net As Double
    net = gross - totalAllowance
    If net < 0# Then net = 0#

    Dim out As Object
    Set out = CreateObject("Scripting.Dictionary")
    out("gross_rent") = Round(gross, 2)
    out("allowance_total") = Round(totalAllowance, 2)
    out("monthly_rent") = Round(net, 2)
    out("annual_rent") = Round(net * 12#, 2)
    Set out("allowances") = allowancesArr

    Set ComputeNetRent = out
    Exit Function

Fail:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub RaiseMissingKey(tableName As String, year As Long, cacheFolder As String, unitId As String, key As String)
    Err.Raise RENT_CALC_ERR_MISSING, "AMI_Optix_RentCalcTables.ComputeNetRent", _
              "Local rent calc blocked (missing lookup in " & tableName & ")." & vbCrLf & vbCrLf & _
              "Year: " & CStr(year) & vbCrLf & _
              "Cache folder: " & cacheFolder & vbCrLf & _
              "Unit: " & unitId & vbCrLf & _
              "Key: " & key
End Sub

Private Function BedroomLabelFromCount(bedrooms As Variant) As String
    Dim n As Long
    n = 0
    On Error Resume Next
    If IsNumeric(bedrooms) Then n = CLng(Round(CDbl(bedrooms), 0))
    On Error GoTo 0

    If n <= 0 Then
        BedroomLabelFromCount = "studio"
    ElseIf n >= 5 Then
        BedroomLabelFromCount = "5 BR"
    Else
        BedroomLabelFromCount = CStr(n) & " BR"
    End If
End Function

Private Function CanonAmiKey(ami As Double) As String
    CanonAmiKey = Format$(Round(CDbl(ami), 4), "0.0000")
End Function

Private Function ParseInvariantDouble(value As String) As Double
    Dim s As String
    s = Trim$(CStr(value))
    s = Replace(s, ",", ".")
    ParseInvariantDouble = Val(s)
End Function

Private Function ParseCsvLine(line As String) As Variant
    On Error GoTo Fail

    Dim fields() As String
    ReDim fields(0 To 0)

    Dim i As Long
    Dim ch As String
    Dim cur As String
    cur = ""

    Dim inQuotes As Boolean
    inQuotes = False

    Dim idx As Long
    idx = 0

    For i = 1 To Len(line)
        ch = Mid$(line, i, 1)
        If inQuotes Then
            If ch = """" Then
                If i < Len(line) And Mid$(line, i + 1, 1) = """" Then
                    cur = cur & """"
                    i = i + 1
                Else
                    inQuotes = False
                End If
            Else
                cur = cur & ch
            End If
        Else
            If ch = """" Then
                inQuotes = True
            ElseIf ch = "," Then
                fields(idx) = cur
                idx = idx + 1
                ReDim Preserve fields(0 To idx)
                cur = ""
            Else
                cur = cur & ch
            End If
        End If
    Next i

    fields(idx) = cur
    ParseCsvLine = fields
    Exit Function

Fail:
    ParseCsvLine = Empty
End Function

Private Function FormatUtilityTypeGuideline(value As String, category As String) As String
    Dim v As String
    Dim c As String
    v = LCase$(Trim$(CStr(value)))
    c = LCase$(Trim$(CStr(category)))

    Select Case c
        Case "electricity"
            If v = "tenant_pays" Then
                FormatUtilityTypeGuideline = "Tenant Pays"
            Else
                FormatUtilityTypeGuideline = "N/A or owner pays"
            End If
            Exit Function

        Case "cooking"
            Select Case v
                Case "electric", "electric_stove": FormatUtilityTypeGuideline = "Electric Stove"
                Case "gas": FormatUtilityTypeGuideline = "Gas Stove"
                Case Else: FormatUtilityTypeGuideline = "N/A or owner pays"
            End Select
            Exit Function

        Case "heat"
            Select Case v
                Case "electric_ccashp": FormatUtilityTypeGuideline = "Electric Heat - Cold Climate Air Source Heat Pump (ccASHP)1"
                Case "electric_other": FormatUtilityTypeGuideline = "Electric Heat - Other2"
                Case "gas": FormatUtilityTypeGuideline = "Gas Heat"
                Case "oil": FormatUtilityTypeGuideline = "Oil Heat"
                Case Else: FormatUtilityTypeGuideline = "N/A or owner pays"
            End Select
            Exit Function

        Case "hot_water"
            Select Case v
                Case "electric_heat_pump": FormatUtilityTypeGuideline = "Electric Hot Water - Heat Pump"
                Case "electric_other": FormatUtilityTypeGuideline = "Electric Hot Water - Other"
                Case "gas": FormatUtilityTypeGuideline = "Gas Hot Water"
                Case "oil": FormatUtilityTypeGuideline = "Oil Hot Water"
                Case Else: FormatUtilityTypeGuideline = "N/A or owner pays"
            End Select
            Exit Function
    End Select

    FormatUtilityTypeGuideline = CStr(value)
End Function
