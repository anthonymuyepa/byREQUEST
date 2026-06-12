'==========================================================
' ExcelIncidentLoader_Ingest.vbs  (PURE INGEST)
' - Reads Excel -> calls dbo.usp_IncidentStaging_Upsert
' - Robust parsing; NULLs for blanks/unknowns
' - Logs to file (fallback to stdout)
' - NO EXCLUSIONS at load time - all vendors with valid headers ingested
'==========================================================

Option Explicit

' ---------- CONFIG ----------
Dim INI_PATH: INI_PATH = "S:\Reports\ITScoreCard\ITPerformanceScoreCard.ini"
Dim DEFAULT_LOG: DEFAULT_LOG = "S:\Reports\ITScoreCard\IncidentIngest.log"

' ---------- GLOBALS ----------
Dim gCn, gExcel, gWb, gCfg
Dim gInserted, gSkipped, gErrors, gWarnings
Dim FLog    ' log stream (optional)

'========================
' byREQUEST entry points
'========================
Sub StartDoc
    OpenLogger

    LogLine "=== Vendor DownTimes Incident Ingestion: START ==="
	Fout.WriteLine "=== Vendor DownTimes Incident Ingestion: START ==="
    gInserted = 0: gSkipped = 0: gErrors = 0: gWarnings = 0

    Set gCfg = ReadIniSafe(INI_PATH)
    If gCfg Is Nothing Then
        LogLine "ERROR: INI not found: " & INI_PATH
        EndRun
        Exit Sub
    End If

    ' REMOVED: All vendor exclusion logic - no filtering at load time

    Dim connStr: connStr = GetRequired(gCfg, "Database.ConnStr")
    Dim excelPath: excelPath = GetRequired(gCfg, "Loader.ExcelPath")
    If Len(connStr)=0 Or Len(excelPath)=0 Then
        EndRun
        Exit Sub
    End If

    If Not FileExists(excelPath) Then
        LogLine "ERROR: ExcelPath not found: " & excelPath
        EndRun
        Exit Sub
    End If

    ' --- DB ---
    Set gCn = CreateObject("ADODB.Connection")
    On Error Resume Next
    gCn.Open connStr
    If Err.Number <> 0 Then
        LogLine "ERROR: DB open failed (" & Err.Number & ") " & Err.Description
        Err.Clear
        On Error GoTo 0
        EndRun
        Exit Sub
    End If
    On Error GoTo 0

    ' --- Excel (headless) ---
    Set gExcel = CreateObject("Excel.Application")
    gExcel.Visible = False
    gExcel.DisplayAlerts = False

    On Error Resume Next
    Set gWb = gExcel.Workbooks.Open(excelPath, 0, True)
    If Err.Number <> 0 Or gWb Is Nothing Then
        LogLine "ERROR: Open workbook failed (" & Err.Number & ") " & Err.Description
        Err.Clear
        On Error GoTo 0
        EndRun
        Exit Sub
    End If
    On Error GoTo 0

    ' --- Ingest ONLY sheets with valid headers ---
    Dim ws
    For Each ws In gWb.Worksheets
        If ws.Visible = -1 Then
            LogLine "Processing sheet: " & ws.Name
            ProcessSheet ws
        Else
            LogLine "Skip hidden: " & ws.Name
        End If
    Next

    EndRun
End Sub

' Always call this before exiting StartDoc.
Sub EndRun()
    CleanupResources
    LogLine "=== Summary: inserted=" & gInserted & " skipped=" & gSkipped & _
            " warnings=" & gWarnings & " errors=" & gErrors & " ==="
			
	Fout.WriteLine "=== Summary: inserted=" & gInserted & " skipped=" & gSkipped & " warnings=" & gWarnings & " errors=" & gErrors & " ==="
    LogLine "=== Excel Incident Ingestion: END ==="
    CloseLogger
End Sub

Sub ProcessLine(): End Sub
Sub CloseDoc():   End Sub

'========================
'   CLEANUP / LOGGING
'========================
Sub CleanupResources()
    On Error Resume Next
    If Not gWb Is Nothing Then gWb.Close False
    If Not gExcel Is Nothing Then gExcel.Quit
    If Not gCn Is Nothing Then gCn.Close
    Set gWb = Nothing: Set gExcel = Nothing: Set gCn = Nothing
    On Error GoTo 0
End Sub

Sub OpenLogger()
    On Error Resume Next
    Dim logPath: logPath = DEFAULT_LOG
    If Not (gCfg Is Nothing) Then
        If gCfg.Exists("Loader.LogPath") And Len(gCfg("Loader.LogPath"))>0 Then
            logPath = gCfg("Loader.LogPath")
        End If
    End If
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
    Set FLog = fso.OpenTextFile(logPath, 8, True) ' Append/create
    On Error GoTo 0
End Sub

Sub CloseLogger()
    On Error Resume Next
    If Not FLog Is Nothing Then FLog.Close
    Set FLog = Nothing
    On Error GoTo 0
End Sub

Sub LogLine(msg)
    Dim line: line = Now & " | " & msg
    If Not FLog Is Nothing Then
      ' Uncomment next line to log every transaction.
      ' On Error Resume Next: Fout.WriteLine line: On Error GoTo 0
    Else
       ' WScript.Echo line
    End If
End Sub

'========================
'        CORE INGEST
'========================
Sub ProcessSheet(ws)
    Dim map, hdrRow
    Set map = CreateObject("Scripting.Dictionary")

    hdrRow = FindHeaderRow(ws, map)
    If hdrRow = 0 Then 
        LogLine "WARN: No headers found on sheet '" & ws.Name & "' - skipping"
        ' DO NOT PROCESS: Sheets without headers are not valid vendors
        Exit Sub
    End If

    Dim lastRow, r, rowData
    lastRow = FindLastRow(ws, hdrRow + 1)
    
    If lastRow <= hdrRow Then
        LogLine "INFO: Sheet '" & ws.Name & "' has headers but no data rows - vendor without incidents"
        ' PROCESS: This represents a vendor with no incidents yet  
        ' We'll insert minimal record to track the vendor existence
        ProcessVendorWithoutIncidents ws.Name
        Exit Sub
    End If

    LogLine "Data rows to process: " & (lastRow - hdrRow)

    ' Track if we processed ANY valid incidents
    Dim processedAnyIncidents
    processedAnyIncidents = False

    For r = hdrRow + 1 To lastRow
        Set rowData = ExtractRowData(ws, r, map)

        If ShouldProcessRow(rowData) Then
            ' We found at least one valid incident (has DateOfEvent)
            processedAnyIncidents = True
            
            If InsertIncident( _
                rowData("Vendor"), rowData("DateOfEvent"), rowData("TimeDown"), _
                rowData("DateUp"),   rowData("TimeUp"),   rowData("ProductAffected"), _
                rowData("DepartmentAffected"), rowData("MemberImpactYN"), _
                rowData("TotalDT_Min"), rowData("Severity"), rowData("Notes"), rowData("Source") ) Then

                gInserted = gInserted + 1
                LogLine "insert:- r" & r & " vendor=" & rowData("Vendor") & _
                        " date=" & SafeShow(rowData("DateOfEvent")) & _
                        " totalDT=" & SafeShow(rowData("TotalDT_Min"))

            Else
                gSkipped = gSkipped + 1
                LogLine "skiped:- r" & r & " DB upsert returned False vendor=" & rowData("Vendor")
            End If

            If Len(rowData("Warnings")) > 0 Then
                gWarnings = gWarnings + 1
                LogLine "WARN r" & r & ": " & rowData("Warnings")
                FLog.WriteLine "WARN " & ws.Name & " r" & r & ": " & rowData("Warnings")
            End If
        End If
    Next
    
    ' NEW: If we found headers and data rows but NO valid incidents (all DateOfEvent were NULL)
    If Not processedAnyIncidents Then
        LogLine "INFO: Sheet '" & ws.Name & "' has data rows but no valid incidents (all DateOfEvent NULL) - vendor without incidents"
        ProcessVendorWithoutIncidents ws.Name
    End If
End Sub

' Process vendor sheets that have headers but no data rows (vendors without incidents)
Sub ProcessVendorWithoutIncidents(vendorName)
    ' Insert a minimal record to track that this vendor exists
    ' This ensures all vendors are represented in the database
	MsgBox "Found1"
    If InsertMinimalVendorRecord(vendorName) Then
        LogLine "insert:- vendor='" & vendorName & "' (no incidents)"
        gInserted = gInserted + 1
    Else
        LogLine "skiped:- vendor='" & vendorName & "' minimal record failed"
        gSkipped = gSkipped + 1
    End If
End Sub

' Insert minimal record for vendors without incidents

Function InsertMinimalVendorRecord(vendorName)
    On Error Resume Next
    
    ' Use current month for the monitoring record
    Dim currentMonth: currentMonth = DateSerial(Year(Date), Month(Date), 1)
    
    ' Call your existing InsertIncident function with monitoring-only data
    If InsertIncident( _
        vendorName, currentMonth, Null, currentMonth, Null, _
        "", "", "", 0, "Low", "MONITORING-ONLY", "LOADER" ) Then
        
        InsertMinimalVendorRecord = True
        MsgBox "inserted monitoring-only record for vendor: " & vendorName
    Else
        InsertMinimalVendorRecord = False
        MsgBox "failed to insert monitoring-only record for vendor: " & vendorName
    End If
    On Error GoTo 0
End Function

Function SafeShow(v)
    If IsMissingOrNull(v) Then
        SafeShow = "NULL"
    Else
        SafeShow = CStr(v)
    End If
End Function

Function FindHeaderRow(ws, map)
    Dim tryRow, lastCol, c, name
    For tryRow = 1 To 10
        lastCol = ws.Cells(tryRow, ws.Columns.Count).End(-4159).Column
        For c = 1 To lastCol
            name = LCase(Trim(CStr(ws.Cells(tryRow, c).Value)))
            If Len(name) > 0 Then
                Select Case True
                    Case InStr(name,"date of event")>0 Or InStr(name,"dateofevent")>0: map("DateOfEvent") = c
                    Case InStr(name,"time down")>0 Or InStr(name,"timedown")>0 Or InStr(name,"time do")>0: map("TimeDown") = c
                    Case InStr(name,"date up")>0 Or InStr(name,"dateup")>0: map("DateUp") = c
                    Case InStr(name,"time up")>0 Or InStr(name,"timeup")>0 Or name="time": map("TimeUp") = c
                    Case InStr(name,"product")>0: map("ProductAffected") = c
                    Case InStr(name,"department")>0: map("DepartmentAffected") = c
                    Case InStr(name,"member impact")>0 Or InStr(name,"memberimpact")>0: map("MemberImpactYN") = c
                    Case InStr(name,"total")>0 Or InStr(name,"dt")>0 Or InStr(name,"min")>0: map("TotalDT_Min") = c
                    Case InStr(name,"severity")>0: map("Severity") = c
                    Case InStr(name,"notes")>0: map("Notes") = c
                    Case InStr(name,"source")>0: map("Source") = c
                End Select
            End If
        Next
        If map.Exists("DateOfEvent") And map.Count >= 3 Then FindHeaderRow = tryRow: Exit Function
        map.RemoveAll
    Next
    FindHeaderRow = 0
End Function

Function FindLastRow(ws, startRow)
    Dim maxRow, col, testRow
    maxRow = startRow
    For col = 1 To 15
        On Error Resume Next
        testRow = ws.Cells(ws.Rows.Count, col).End(-4162).Row
        If Err.Number = 0 And testRow > maxRow Then maxRow = testRow
        Err.Clear: On Error GoTo 0
    Next
    FindLastRow = maxRow
End Function

Function ExtractRowData(ws, row, map)
    Dim d: Set d = CreateObject("Scripting.Dictionary")
    d("Vendor") = SafeString(Trim(CStr(ws.Name)))
    d("RowNumber") = row
    d("HasErrors") = False
    d("Warnings") = ""

    d("DateOfEvent")        = ExtractDate(ws, row, map, "DateOfEvent", d)
    d("TimeDown")           = ExtractTime(ws, row, map, "TimeDown", d)
    d("DateUp")             = ExtractDate(ws, row, map, "DateUp", d)
    d("TimeUp")             = ExtractTime(ws, row, map, "TimeUp", d)
    d("ProductAffected")    = ExtractString(ws, row, map, "ProductAffected", d)
    d("DepartmentAffected") = ExtractString(ws, row, map, "DepartmentAffected", d)
    d("MemberImpactYN")     = ExtractString(ws, row, map, "MemberImpactYN", d)
    d("TotalDT_Min")        = ExtractNumber(ws, row, map, "TotalDT_Min", d)
    d("Severity")           = ExtractString(ws, row, map, "Severity", d)
    d("Notes")              = ExtractString(ws, row, map, "Notes", d)
    d("Source")             = ExtractString(ws, row, map, "Source", d)
    Set ExtractRowData = d
End Function

Function ShouldProcessRow(d)
    ' Check for required fields first
    If Len(d("Vendor")) = 0 Then
        gSkipped = gSkipped + 1
        LogLine "skiped:- r" & d("RowNumber") & " no Vendor"
        ShouldProcessRow = False
        Exit Function
    End If

    If IsNull(d("DateOfEvent")) Then
        gSkipped = gSkipped + 1
        LogLine "skiped:- r" & d("RowNumber") & " no DateOfEvent"
        ShouldProcessRow = False
        Exit Function
    End If

    ' REMOVED: All vendor exclusion logic - no filtering at load time

    If d("HasErrors") Then
        gSkipped = gSkipped + 1
        LogLine "skiped:- r" & d("RowNumber") & " extraction errors"
        ShouldProcessRow = False
        Exit Function
    End If

    ShouldProcessRow = True
End Function

'========================
'    EXTRACTION HELPERS
'========================
Function ExtractDate(ws, row, map, key, bag)
    On Error Resume Next
    Dim v, spacePos, part
    If Not map.Exists(key) Then ExtractDate=Null: Exit Function
    v = Trim(CStr(ws.Cells(row, map(key)).Value))
    If Len(v)=0 Then ExtractDate=Null: Exit Function
    If IsDate(v) Then ExtractDate=CDate(v): Exit Function
    spacePos = InStr(v," ")
    If spacePos>0 Then
        part = Left(v, spacePos-1)
        If IsDate(part) Then ExtractDate=CDate(part): AddWarning bag,key & " parsed from mixed value": Exit Function
    End If
    AddWarning bag,"Invalid date for " & key & ": '" & v & "'": ExtractDate=Null
    On Error GoTo 0
End Function

Function ExtractTime(ws, row, map, key, bag)
    On Error Resume Next
    Dim v, lc
    If Not map.Exists(key) Then
        ExtractTime = Null
        Exit Function
    End If

    v = Trim(CStr(ws.Cells(row, map(key)).Value))
    If Len(v) = 0 Then
        ExtractTime = Null
        Exit Function
    End If

    lc = LCase(v)
    If lc = "unknown" Or lc = "during gm processing" Or lc = "???" Or lc = "n/a" Then
        AddWarning bag, key & " placeholder: '" & v & "'"
        ExtractTime = Null
        Exit Function
    End If

    v = Replace(v, ",", "")

    If IsDate(v) Then
        ExtractTime = CDate(v)
    ElseIf IsNumeric(v) Then
        ExtractTime = CDate(CDbl(v))
    Else
        AddWarning bag, "Invalid time for " & key & ": '" & v & "'"
        ExtractTime = Null
    End If

    On Error GoTo 0
End Function

Function ExtractNumber(ws, row, map, key, bag)
    On Error Resume Next
    Dim v
    If Not map.Exists(key) Then ExtractNumber=Null: Exit Function
    v = Trim(CStr(ws.Cells(row, map(key)).Value))
    If Len(v)=0 Then ExtractNumber=Null: Exit Function
    v = Replace(Replace(Replace(v,"$",""),",",""),"£","")
    If IsNumeric(v) Then ExtractNumber=CDbl(v) Else AddWarning bag,"Invalid number for " & key & ": '" & v & "'": ExtractNumber=Null
    On Error GoTo 0
End Function

Function ExtractString(ws, row, map, key, bag)
    On Error Resume Next
    Dim s, n
    If Not map.Exists(key) Then ExtractString=Null: Exit Function
    s = SafeString(Trim(CStr(ws.Cells(row, map(key)).Value)))
    n = Len(s)
    If n>1000 Then AddWarning bag, key & " truncated from " & n & " chars": s = Left(s,1000)
    ExtractString = s
    On Error GoTo 0
End Function

Function SafeString(v)
    If IsNull(v) Or IsEmpty(v) Then SafeString = "" Else SafeString = CStr(v)
End Function

Sub AddWarning(bag, w)
    If Len(bag("Warnings"))>0 Then bag("Warnings") = bag("Warnings") & "; " & w Else bag("Warnings") = w
End Sub

'========================
'    DATABASE CALL
'========================
Function InsertIncident(Vendor, DateOfEvent, TimeDown, DateUp, TimeUp, _
                        ProductAffected, DepartmentAffected, MemberImpactYN, _
                        TotalDT_Min, Severity, Notes, Source)
    On Error Resume Next

    Dim cmd: Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = gCn
    cmd.CommandType = 4  ' adCmdStoredProc
    cmd.CommandText  = "dbo.usp_IncidentStaging_Upsert"

    cmd.Parameters.Refresh
    If Err.Number <> 0 Then
        LogLine "ERROR: Param refresh failed (" & Err.Number & ") " & Err.Description
        Err.Clear: gErrors = gErrors + 1: InsertIncident = False: Exit Function
    End If

    SafeSetParam cmd,"@Vendor",             Vendor
    SafeSetParam cmd,"@DateOfEvent",        DateOnly(DateOfEvent)
    SafeSetParam cmd,"@TimeDown",           TimeStringOrNull(TimeDown)
    SafeSetParam cmd,"@DateUp",             DateOnly(DateUp)
    SafeSetParam cmd,"@TimeUp",             TimeStringOrNull(TimeUp)
    SafeSetParam cmd,"@ProductAffected",    ProductAffected
    SafeSetParam cmd,"@DepartmentAffected", DepartmentAffected
    SafeSetParam cmd,"@MemberImpactYN",     MemberImpactYN
    SafeSetParam cmd,"@TotalDT_Min",        TotalDT_Min
    SafeSetParam cmd,"@Severity",           Severity
    SafeSetParam cmd,"@Notes",              Notes
    SafeSetParam cmd,"@Source",             Source

    cmd.Execute
    If Err.Number <> 0 Then
        LogLine "DB ERROR: (" & Err.Number & ") " & Err.Description
        Err.Clear: gErrors = gErrors + 1: InsertIncident = False
    Else
        InsertIncident = True
    End If
    On Error GoTo 0
End Function

Sub SafeSetParam(cmd, name, val)
    On Error Resume Next
    If IsMissingOrNull(val) Then cmd.Parameters(name).Value = Null Else cmd.Parameters(name).Value = val
    If Err.Number <> 0 Then LogLine "ERROR: Set param " & name & " (" & Err.Number & ") " & Err.Description: Err.Clear
    On Error GoTo 0
End Sub

Function IsMissingOrNull(v)
    If IsEmpty(v) Or IsNull(v) Then
        IsMissingOrNull = True
    ElseIf VarType(v)=vbString And Len(Trim(CStr(v)))=0 Then
        IsMissingOrNull = True
    Else
        IsMissingOrNull = False
    End If
End Function

Function DateOnly(d)
    If IsMissingOrNull(d) Then DateOnly = Null Else DateOnly = DateSerial(Year(CDate(d)), Month(CDate(d)), Day(CDate(d)))
End Function

Function TimeStringOrNull(t)
    If IsMissingOrNull(t) Then
        TimeStringOrNull = Null
    Else
        Dim dt: dt = CDate(t)
        TimeStringOrNull = Right("0"&Hour(dt),2) & ":" & Right("0"&Minute(dt),2) & ":" & Right("0"&Second(dt),2)
    End If
End Function

'========================
'    INI / UTIL
'========================
Function ReadIniSafe(path)
    On Error Resume Next
    Dim fso, ts, d, sect, line, p
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then Exit Function
    Set ts = fso.OpenTextFile(path,1)
    Set d = CreateObject("Scripting.Dictionary")
    Do Until ts.AtEndOfStream
        line = Trim(ts.ReadLine)
        If Len(line)=0 Or Left(line,1)=";" Then
        ElseIf Left(line,1)="[" And Right(line,1)="]" Then
            sect = Mid(line,2,Len(line)-2)
        Else
            p = InStr(line,"=")
            If p>0 Then d(sect & "." & Trim(Left(line,p-1))) = Trim(Mid(line,p+1))
        End If
    Loop
    ts.Close
    Set ReadIniSafe = d
    On Error GoTo 0
End Function

Function GetRequired(cfg, key)
    If Not (cfg Is Nothing) And cfg.Exists(key) And Len(cfg(key))>0 Then
        GetRequired = cfg(key)
    Else
        LogLine "ERROR: Missing INI key: " & key
        GetRequired = ""
    End If
End Function

Function FileExists(p)
    FileExists = CreateObject("Scripting.FileSystemObject").FileExists(p)
End Function