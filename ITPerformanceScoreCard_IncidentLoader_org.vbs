'==========================================================
' ExcelIncidentLoader_Ini.vbs - ROBUST VERSION (VBScript)
' Handles diverse data formats gracefully
'==========================================================

Option Explicit

' === Adjust if your INI is elsewhere ===
Dim INI_PATH: INI_PATH = "S:\Reports\ITScoreCard\ITPerformanceScoreCard.ini"

' === Globals ===
Dim gCn, gExcel, gWb, gCfg
Dim gInserted, gSkipped, gErrors, gWarnings

'========================
' byREQUEST entry points
'========================
Sub StartDoc()
    LogLine "=== Starting Excel Incident Loader ==="
    gInserted = 0: gSkipped = 0: gErrors = 0: gWarnings = 0

    Set gCfg = ReadIniSafe(INI_PATH)
    If gCfg Is Nothing Then
        LogLine "ERROR: INI not found or unreadable: " & INI_PATH
        Exit Sub
    End If

    Dim connStr: connStr = GetRequired(gCfg, "Database.ConnStr")
    Dim excelPath: excelPath = GetRequired(gCfg, "Loader.ExcelPath")
    If Len(connStr)=0 Or Len(excelPath)=0 Then Exit Sub
    
    If Not FileExists(excelPath) Then
        LogLine "ERROR: ExcelPath not found: " & excelPath
        Exit Sub
    End If

    ' === Open database ===
    Set gCn = CreateObject("ADODB.Connection")
    gCn.Open connStr

    ' === Open Excel ===
    Set gExcel = CreateObject("Excel.Application")
    gExcel.Visible = False
    gExcel.DisplayAlerts = False
    
    On Error Resume Next
    Set gWb = gExcel.Workbooks.Open(excelPath, 0, True)
    If Err.Number <> 0 Or gWb Is Nothing Then
        LogLine "ERROR: Cannot open workbook: (" & Err.Number & ") " & Err.Description
        Err.Clear
        CleanupResources
        Exit Sub
    End If
    On Error GoTo 0

    ' === Process sheets ===
    Dim ws
    For Each ws In gWb.Worksheets
        If ws.Visible = -1 Then
            LogLine "Processing sheet: " & ws.Name
            ProcessSheet ws
        Else
            LogLine "Skipping hidden sheet: " & ws.Name
        End If
    Next

    ' === Cleanup ===
    CleanupResources
    
    LogLine "=== Summary: " & gInserted & " inserted, " & gSkipped & " skipped, " & gWarnings & " warnings, " & gErrors & " errors ==="
End Sub

Sub ProcessLine()
    ' No-op
End Sub

Sub CloseDoc()
    ' No-op (we clean up in StartDoc)
End Sub

Sub CleanupResources()
    On Error Resume Next
    If Not gWb Is Nothing Then gWb.Close False
    If Not gExcel Is Nothing Then gExcel.Quit
    If Not gCn Is Nothing Then gCn.Close
    Set gWb = Nothing
    Set gExcel = Nothing
    Set gCn = Nothing
    On Error GoTo 0
End Sub

'========================
'        Core Logic
'========================
Sub ProcessSheet(ws)
    Dim map, hdrRow
    Set map = CreateObject("Scripting.Dictionary")
    
    ' === Find headers ===
    hdrRow = FindHeaderRow(ws, map)
    If hdrRow = 0 Then
        LogLine "WARN: No headers found on sheet: " & ws.Name
        Exit Sub
    End If
    
    ' === Process data rows ===
    Dim lastRow, r, rowData
    lastRow = FindLastRow(ws, hdrRow + 1)
    LogLine "Processing " & (lastRow - hdrRow) & " data rows"
    
    For r = hdrRow + 1 To lastRow
        Set rowData = ExtractRowData(ws, r, map)
        If ShouldProcessRow(rowData) Then
            ProcessValidRow rowData
        End If
    Next
End Sub

Function FindHeaderRow(ws, map)
    Dim tryRow, lastCol, c, name
    For tryRow = 1 To 10
        lastCol = ws.Cells(tryRow, ws.Columns.Count).End(-4159).Column  ' xlToLeft
        
        For c = 1 To lastCol
            name = LCase(Trim(CStr(ws.Cells(tryRow, c).Value)))
            If Len(name) > 0 Then
                ' Flexible header matching
                Select Case True
                    Case InStr(name, "date of event") > 0 Or InStr(name, "dateofevent") > 0 
                        map("DateOfEvent") = c
                    Case InStr(name, "time down") > 0 Or InStr(name, "timedown") > 0 Or InStr(name, "time do") > 0 
                        map("TimeDown") = c
                    Case InStr(name, "date up") > 0 Or InStr(name, "dateup") > 0 
                        map("DateUp") = c
                    Case InStr(name, "time up") > 0 Or InStr(name, "timeup") > 0 Or (name="time" And Not map.Exists("TimeUp"))
                        map("TimeUp") = c
                    Case InStr(name, "product") > 0 
                        map("ProductAffected") = c
                    Case InStr(name, "department") > 0 
                        map("DepartmentAffected") = c
                    Case InStr(name, "member impact") > 0 Or InStr(name, "memberimpact") > 0 
                        map("MemberImpactYN") = c
                    Case InStr(name, "total") > 0 Or InStr(name, "dt") > 0 Or InStr(name, "min") > 0 
                        map("TotalDT_Min") = c
                    Case InStr(name, "severity") > 0 
                        map("Severity") = c
                    Case InStr(name, "notes") > 0 
                        map("Notes") = c
                    Case InStr(name, "source") > 0 
                        map("Source") = c
                End Select
            End If
        Next
        
        ' Require at least DateOfEvent + a couple others
        If map.Exists("DateOfEvent") And map.Count >= 3 Then
            LogLine "Headers found at row " & tryRow & " (" & map.Count & " columns mapped)"
            FindHeaderRow = tryRow
            Exit Function
        End If
        
        map.RemoveAll
    Next
    
    FindHeaderRow = 0
End Function

Function FindLastRow(ws, startRow)
    Dim maxRow, col, testRow
    maxRow = startRow
    
    ' Check multiple columns to find the real last row
    For col = 1 To 15
        On Error Resume Next
        testRow = ws.Cells(ws.Rows.Count, col).End(-4162).Row  ' xlUp
        If Err.Number = 0 And testRow > maxRow Then maxRow = testRow
        Err.Clear
        On Error GoTo 0
    Next
    
    FindLastRow = maxRow
End Function

Function ExtractRowData(ws, row, map)
    Dim data
    Set data = CreateObject("Scripting.Dictionary")
    
    data("Vendor") = SafeString(Trim(CStr(ws.Name)))
    data("RowNumber") = row
    data("HasErrors") = False
    data("Warnings") = ""
    
    ' Extract all fields with robust error handling
    data("DateOfEvent")        = ExtractDate(ws, row, map, "DateOfEvent", data)
    data("TimeDown")           = ExtractTime(ws, row, map, "TimeDown", data)
    data("DateUp")             = ExtractDate(ws, row, map, "DateUp", data)
    data("TimeUp")             = ExtractTime(ws, row, map, "TimeUp", data)
    data("ProductAffected")    = ExtractString(ws, row, map, "ProductAffected", data)
    data("DepartmentAffected") = ExtractString(ws, row, map, "DepartmentAffected", data)
    data("MemberImpactYN")     = ExtractString(ws, row, map, "MemberImpactYN", data)
    data("TotalDT_Min")        = ExtractNumber(ws, row, map, "TotalDT_Min", data)
    data("Severity")           = ExtractString(ws, row, map, "Severity", data)
    data("Notes")              = ExtractString(ws, row, map, "Notes", data)
    data("Source")             = ExtractString(ws, row, map, "Source", data)
    
    Set ExtractRowData = data
End Function

Function ShouldProcessRow(rowData)
    ' Skip if no vendor
    If Len(rowData("Vendor")) = 0 Then
        gSkipped = gSkipped + 1
        LogLine "SKIP Row " & rowData("RowNumber") & ": Missing Vendor"
        ShouldProcessRow = False
        Exit Function
    End If
    
    ' Skip if no date of event
    If IsNull(rowData("DateOfEvent")) Then
        gSkipped = gSkipped + 1
        LogLine "SKIP Row " & rowData("RowNumber") & ": Missing/invalid DateOfEvent"
        ShouldProcessRow = False
        Exit Function
    End If
    
    ' Skip if row has critical errors
    If rowData("HasErrors") Then
        gSkipped = gSkipped + 1
        LogLine "SKIP Row " & rowData("RowNumber") & ": Critical extraction errors"
        ShouldProcessRow = False
        Exit Function
    End If
    
    ShouldProcessRow = True
End Function

Sub ProcessValidRow(rowData)
    ' Log warnings if any
    If Len(rowData("Warnings")) > 0 Then
        LogLine "WARN Row " & rowData("RowNumber") & ": " & rowData("Warnings")
        gWarnings = gWarnings + 1
    End If
    
    ' Insert into database
    If InsertIncident( _
        rowData("Vendor"), _
        rowData("DateOfEvent"), _
        rowData("TimeDown"), _
        rowData("DateUp"), _
        rowData("TimeUp"), _
        rowData("ProductAffected"), _
        rowData("DepartmentAffected"), _
        rowData("MemberImpactYN"), _
        rowData("TotalDT_Min"), _
        rowData("Severity"), _
        rowData("Notes"), _
        rowData("Source")) Then
        
        gInserted = gInserted + 1
        LogLine "INSERTED Row " & rowData("RowNumber") & ": " & rowData("Vendor") & " | " & FormatDateTime(rowData("DateOfEvent"))
    Else
        gSkipped = gSkipped + 1
        LogLine "SKIPPED Row " & rowData("RowNumber") & ": Database insertion failed"
    End If
End Sub

'========================
'    Data Extraction
'========================
Function ExtractDate(ws, row, map, key, rowData)
    On Error Resume Next
    Dim rawValue, result, spacePos, datePart
    
    If Not map.Exists(key) Then
        ExtractDate = Null
        Exit Function
    End If
    
    rawValue = Trim(CStr(ws.Cells(row, map(key)).Value))
    If Len(rawValue) = 0 Then
        ExtractDate = Null
        Exit Function
    End If
    
    If IsDate(rawValue) Then
        result = CDate(rawValue)
        ExtractDate = result
        Exit Function
    End If
    
    spacePos = InStr(rawValue, " ")
    If spacePos > 0 Then
        datePart = Left(rawValue, spacePos - 1)
        If IsDate(datePart) Then
            result = CDate(datePart)
            ExtractDate = result
            AddWarning rowData, key & " extracted from complex string"
            Exit Function
        End If
    End If
    
    ' Final fallback
    AddWarning rowData, "Invalid date format for " & key & ": '" & rawValue & "'"
    ExtractDate = Null
    On Error GoTo 0
End Function

Function ExtractTime(ws, row, map, key, rowData)
    On Error Resume Next
    Dim rawValue, result, lc
    
    If Not map.Exists(key) Then
        ExtractTime = Null
        Exit Function
    End If
    
    rawValue = Trim(CStr(ws.Cells(row, map(key)).Value))
    If Len(rawValue) = 0 Then
        ExtractTime = Null
        Exit Function
    End If
    
    lc = LCase(rawValue)
    If lc = "unknown" Or lc = "during gm processing" Then
        AddWarning rowData, key & " contains non-time value: '" & rawValue & "'"
        ExtractTime = Null
        Exit Function
    End If
    
    rawValue = Replace(rawValue, ",", "")
    
    If IsDate(rawValue) Then
        result = CDate(rawValue)
        ExtractTime = result
    Else
        AddWarning rowData, "Invalid time format for " & key & ": '" & rawValue & "'"
        ExtractTime = Null
    End If
    
    On Error GoTo 0
End Function

Function ExtractNumber(ws, row, map, key, rowData)
    On Error Resume Next
    Dim rawValue
    
    If Not map.Exists(key) Then
        ExtractNumber = Null
        Exit Function
    End If
    
    rawValue = Trim(CStr(ws.Cells(row, map(key)).Value))
    If Len(rawValue) = 0 Then
        ExtractNumber = Null
        Exit Function
    End If
    
    rawValue = Replace(rawValue, "$", "")
    rawValue = Replace(rawValue, ",", "")
    rawValue = Replace(rawValue, "£", "")
    
    If IsNumeric(rawValue) Then
        ExtractNumber = CDbl(rawValue)
    Else
        AddWarning rowData, "Invalid number for " & key & ": '" & rawValue & "'"
        ExtractNumber = Null
    End If
    
    On Error GoTo 0
End Function

Function ExtractString(ws, row, map, key, rowData)
    On Error Resume Next
    Dim result, origLen
    
    If Not map.Exists(key) Then
        ExtractString = Null
        Exit Function
    End If
    
    result = SafeString(Trim(CStr(ws.Cells(row, map(key)).Value)))
    origLen = Len(result)
    
    If origLen > 1000 Then
        result = Left(result, 1000)
        AddWarning rowData, key & " truncated from " & origLen & " characters"
    End If
    
    ExtractString = result
    On Error GoTo 0
End Function

Function SafeString(val)
    If IsNull(val) Or IsEmpty(val) Then
        SafeString = ""
    Else
        SafeString = CStr(val)
    End If
End Function

Sub AddWarning(rowData, warning)
    If Len(rowData("Warnings")) > 0 Then
        rowData("Warnings") = rowData("Warnings") & "; " & warning
    Else
        rowData("Warnings") = warning
    End If
End Sub

'========================
'    Database Operations
'========================
Function InsertIncident(Vendor, DateOfEvent, TimeDown, DateUp, TimeUp, _
                        ProductAffected, DepartmentAffected, MemberImpactYN, _
                        TotalDT_Min, Severity, Notes, Source)
    On Error Resume Next
    
    Dim cmd
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = gCn
    cmd.CommandType = 4   ' adCmdStoredProc
    cmd.CommandText = "dbo.usp_IncidentStaging_Upsert"
    
    cmd.Parameters.Refresh
    If Err.Number <> 0 Then
        LogLine "ERROR: Cannot refresh stored procedure parameters: (" & Err.Number & ") " & Err.Description
        Err.Clear
        InsertIncident = False
        Exit Function
    End If
    
    ' Safe parameter setting
    SafeSetParam cmd, "@Vendor",             Vendor
    SafeSetParam cmd, "@DateOfEvent",        DateOnly(DateOfEvent)
    SafeSetParam cmd, "@TimeDown",           TimeStringOrNull(TimeDown)
    SafeSetParam cmd, "@DateUp",             DateOnly(DateUp)
    SafeSetParam cmd, "@TimeUp",             TimeStringOrNull(TimeUp)
    SafeSetParam cmd, "@ProductAffected",    ProductAffected
    SafeSetParam cmd, "@DepartmentAffected", DepartmentAffected
    SafeSetParam cmd, "@MemberImpactYN",     MemberImpactYN
    SafeSetParam cmd, "@TotalDT_Min",        TotalDT_Min
    SafeSetParam cmd, "@Severity",           Severity
    SafeSetParam cmd, "@Notes",              Notes
    SafeSetParam cmd, "@Source",             Source
    
    cmd.Execute
    If Err.Number <> 0 Then
        LogLine "DB ERROR: (" & Err.Number & ") " & Err.Description
        gErrors = gErrors + 1
        Err.Clear
        InsertIncident = False
    Else
        InsertIncident = True
    End If
    
    On Error GoTo 0
End Function

Sub SafeSetParam(cmd, name, val)
    On Error Resume Next
    If IsMissingOrNull(val) Then
        cmd.Parameters(name).Value = Null
    Else
        cmd.Parameters(name).Value = val
    End If
    If Err.Number <> 0 Then
        LogLine "ERROR: Setting parameter " & name & " failed: (" & Err.Number & ") " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Function IsMissingOrNull(v)
    If IsEmpty(v) Or IsNull(v) Then
        IsMissingOrNull = True
    ElseIf VarType(v) = vbString And Len(Trim(CStr(v))) = 0 Then
        IsMissingOrNull = True
    Else
        IsMissingOrNull = False
    End If
End Function

Function DateOnly(d)
    If IsMissingOrNull(d) Then
        DateOnly = Null
    Else
        Dim dt: dt = CDate(d)
        DateOnly = DateSerial(Year(dt), Month(dt), Day(dt))
    End If
End Function

Function TimeStringOrNull(t)
    If IsMissingOrNull(t) Then
        TimeStringOrNull = Null
    Else
        Dim dt: dt = CDate(t)
        TimeStringOrNull = Right("0" & Hour(dt), 2) & ":" & _
                           Right("0" & Minute(dt), 2) & ":" & _
                           Right("0" & Second(dt), 2)
    End If
End Function

'========================
'    Configuration & Logging
'========================
Function ReadIniSafe(path)
    On Error Resume Next
    Dim fso, ts, d, sect, line, p
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then Exit Function
    
    Set ts = fso.OpenTextFile(path, 1)
    Set d = CreateObject("Scripting.Dictionary")
    
    Do Until ts.AtEndOfStream
        line = Trim(ts.ReadLine)
        If Len(line) = 0 Or Left(line, 1) = ";" Then
            ' Skip
        ElseIf Left(line, 1) = "[" And Right(line, 1) = "]" Then
            sect = Mid(line, 2, Len(line) - 2)
        Else
            p = InStr(line, "=")
            If p > 0 Then d(sect & "." & Trim(Left(line, p - 1))) = Trim(Mid(line, p + 1))
        End If
    Loop
    
    ts.Close
    Set ReadIniSafe = d
    On Error GoTo 0
End Function

Function GetRequired(cfg, key)
    If Not (cfg Is Nothing) And cfg.Exists(key) And Len(cfg(key)) > 0 Then
        GetRequired = cfg(key)
    Else
        LogLine "ERROR: Missing required INI key: " & key
        GetRequired = ""
    End If
End Function

Function FileExists(p)
    FileExists = CreateObject("Scripting.FileSystemObject").FileExists(p)
End Function

Sub LogLine(msg)
    ' byREQUEST captures stdout; on desktop, use cscript/wscript host
    Fout.Writeline Now & " | " & msg
End Sub
