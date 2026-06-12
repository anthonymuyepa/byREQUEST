' =========================================================
' COB7_GL Period Detail Report (INI-driven)
' - Queries COB7_GL heap table (Run_Date is real DATE type)
' - Returns ALL raw rows in chronological order: ORDER BY Run_Date, ID
' =========================================================
Option Explicit

' ---------------- CONFIG ----------------
'Dim INI_PATH: INI_PATH = "C:\byREQUEST\Templates\COB7_GL_Monthly.ini"
Dim INI_PATH: INI_PATH = "S:\CarolinaTrust\byREQUEST\Templates\COB7_GL_Monthly.ini"
' ---------------- GLOBALS ----------------
Dim gMode, gRelOffset, gYear, gMonth, gQuarter
Dim gStartDateStr, gEndDateStr
Dim gConnStr, gFullTableName, gTable2Part, gObjIdName
Dim gDailyLagDays   ' Configurable lag for DAILY and MTD current
Dim conn, rs, cmd

' =========================================================
' StartDoc
' =========================================================
Private Sub StartDoc()
    On Error Resume Next
    Dim startDate, endDate, periodDisplay, errMsg
    Dim today : today = Date()
    
    If Not FileExists(INI_PATH) Then
        errMsg = "ERROR: INI file not found: " & INI_PATH
        LogLine errMsg : fout.WriteLine errMsg
        Exit Sub
    End If
    
    LoadParameters INI_PATH
    
    If Trim(gConnStr) = "" Then
        errMsg = "ERROR: ConnStr not defined in [Database] section."
        LogLine errMsg : fout.WriteLine errMsg
        Exit Sub
    End If
    
    BuildTableNames gFullTableName, gTable2Part, gObjIdName
    If Len(gTable2Part) = 0 Then
        errMsg = "ERROR: Could not parse TableName"
        LogLine errMsg : fout.WriteLine errMsg
        Exit Sub
    End If
    
    ComputeDateRange startDate, endDate
    
    If Not IsDate(startDate) Or Not IsDate(endDate) Or startDate > endDate Then
        errMsg = "ERROR: Invalid date range (" & TypeName(startDate) & ")"
        LogLine errMsg : fout.WriteLine errMsg
        Exit Sub
    End If
    
    periodDisplay = FormatPeriodDisplay(startDate, endDate, gMode)
    
    ' NEW: Get latest Run_Date for informational note (before main query)
    Dim latestRunDate : latestRunDate = Null
    Set conn = CreateObject("ADODB.Connection")
    Err.Clear
    conn.Open gConnStr
    If Err.Number <> 0 Then
        LogLine "ERROR opening DB connection: " & Err.Description & " (Error #" & Err.Number & ")"
        Exit Sub
    End If
    
    On Error Resume Next
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandType = 1
        .CommandText = "SELECT MAX(Run_Date) AS Latest FROM " & gTable2Part
        Set rs = .Execute
        If Not rs.EOF Then latestRunDate = rs("Latest")
        rs.Close
    End With
    Set rs = Nothing
    Set cmd = Nothing
    If Not IsNull(latestRunDate) Then
        LogLine "INFO: Latest available Run_Date in table = " & Format(latestRunDate, "yyyy-mm-dd")
    Else
        LogLine "INFO: No data found in table (MAX(Run_Date) returned NULL)"
    End If
    On Error GoTo 0
    
    ' Compute periodStr for User2 - consistent mmddyy format
    Dim periodStr
    Dim reportDate
    
    Select Case gMode
        Case "MTD"
            reportDate = endDate   ' lagged to yesterday for current MTD
            periodStr = "MTD_" & Format(reportDate, "mmddyy")
        Case "DAILY"
            reportDate = startDate   ' lagged if applicable
            periodStr = Format(reportDate, "mmddyy")
        Case "MONTHLY"
            If gRelOffset = 0 Then
                reportDate = today
                periodStr = Format(reportDate, "mmddyy")
            Else
                reportDate = startDate   ' first of historical month
                periodStr = Format(reportDate, "mmddyy")
            End If
        Case Else   ' CUSTOM, etc.
            If startDate = endDate Then
                periodStr = Format(startDate, "mmddyy")
            Else
                periodStr = Format(startDate, "mmddyy") & "-" & Format(endDate, "mmddyy")
            End If
    End Select
    
    ' Set Spoolfile vars
    Spoolfile.User1 = "COB7_GL " & UCase(Trim(gMode)) & " - " & periodDisplay
    Spoolfile.User2 = periodStr
    
    ' MAIN QUERY
    Dim sql
    sql = _
        "SELECT " & _
        " [Report], [Run_Date], [Run_Time], [GL_Number], [Branch], " & _
        " [Debit], [Credit], [Sequence], [Voucher], [Teller], " & _
        " [Tran_Nbr_Type1], [Mbr_Nbr], [BR], [Suffix], [Product_Type_Code], " & _
        " [Tran_Nbr_Type2], [MI_Code], [Addl_Detail] " & _
        "FROM " & gTable2Part & " " & _
        "WHERE Run_Date BETWEEN ? AND ? " & _
        "ORDER BY ID"

    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandType = 1
        .CommandText = sql
        .Parameters.Append .CreateParameter("@StartDate", 133, 1, , startDate)
        .Parameters.Append .CreateParameter("@EndDate",   133, 1, , endDate)
    End With

    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorType   = 0
    rs.LockType     = 1
    rs.Open cmd

    If Err.Number <> 0 Then
        fout.WriteLine "ERROR opening recordset: " & Err.Description & " (Error " & Err.Number & ")"
        Err.Clear
    End If

    ' Output
    If rs.EOF And rs.BOF Then
        fout.WriteLine "No COB7_GL records found for period: " & periodDisplay
    Else
        ' As-of note for current/lagging modes
        If gMode = "DAILY" Or gMode = "MTD" Or (gMode = "MONTHLY" And gRelOffset = 0) Then
            Dim asOfText
            asOfText = "Note: Data available only up to the latest loaded date (typical " & gDailyLagDays & "-business-day lag). " & _
                       "Latest Run_Date found in table: " & IIf(IsNull(latestRunDate), "None", Format(latestRunDate, "mmmm d, yyyy"))
            fout.WriteLine asOfText
            fout.WriteLine ""
        End If
        
        ' Header
        fout.WriteLine Pad("Report", 7) & _
            Pad("Run_Date", 11) & _
            Pad("Run_Time", 9) & _
            Pad("GL_Number", 11) & _
            Pad("Branch", 7) & _
            Pad("Debit", 16) & _
            Pad("Credit", 19) & _
            Pad("Sequence", 9) & _
            Pad("Voucher", 9) & _
            Pad("Teller", 8) & _
            Pad("NBR_TYPE1", 12) & _
            Pad("Mbr_Nbr", 10) & _
            Pad("BR", 4) & _
            Pad("SFX", 4) & _
            Pad("PTyp", 5) & _
            Pad("NBR_TYPE2", 10) & _
            Pad("MI_Code", 8) & _
            Pad("Addl_Detail", 50)
        
        Dim rowCount : rowCount = 0
        Do While Not rs.EOF
            rowCount = rowCount + 1
            Dim lineOut
            lineOut = Pad(rs("Report"), 7)
            lineOut = lineOut & Pad(CDate(CStr(rs("Run_Date"))), 11)
            lineOut = lineOut & Pad(rs("Run_Time"), 9)
            lineOut = lineOut & Pad(rs("GL_Number"), 11)
            lineOut = lineOut & Pad(rs("Branch"), 7)
            lineOut = lineOut & Pad(rs("Debit"), 16)
            lineOut = lineOut & Pad(rs("Credit"), 19)
            lineOut = lineOut & Pad(rs("Sequence"), 9)
            lineOut = lineOut & Pad(rs("Voucher"), 9)
            lineOut = lineOut & Pad(rs("Teller"), 8)
            lineOut = lineOut & Pad(rs("Tran_Nbr_Type1"), 12)
            lineOut = lineOut & Pad(rs("Mbr_Nbr"), 10)
            lineOut = lineOut & Pad(rs("BR"), 4)
            lineOut = lineOut & Pad(rs("Suffix"), 4)
            lineOut = lineOut & Pad(rs("Product_Type_Code"), 5)
            lineOut = lineOut & Pad(rs("Tran_Nbr_Type2"), 10)
            lineOut = lineOut & Pad(rs("MI_Code"), 8)
            lineOut = lineOut & Pad(rs("Addl_Detail"), 50)

            fout.WriteLine lineOut
            rs.MoveNext
        Loop
        
        fout.WriteLine ""
        fout.WriteLine "Total rows: " & rowCount
    End If

    ' Cleanup
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State <> 0 Then rs.Close: Set rs = Nothing
    If Not cmd Is Nothing Then Set cmd = Nothing
    If Not conn Is Nothing Then If conn.State <> 0 Then conn.Close: Set conn = Nothing
    On Error GoTo 0
End Sub

' =========================================================
' Helper functions (unchanged except LoadParameters and ComputeDateRange)
' =========================================================

Sub LoadParameters(iniPath)
    Dim cfg : Set cfg = ReadIniSafe(iniPath)
    If cfg Is Nothing Then Exit Sub
    
    gMode         = UCase(Trim(GetIniValue(cfg, "Report.Mode", "MONTHLY")))
    gYear         = GetIniValue(cfg, "Report.Year", "")
    gMonth        = GetIniValue(cfg, "Report.Month", "")
    gQuarter      = GetIniValue(cfg, "Report.Quarter", "")
    gStartDateStr = GetIniValue(cfg, "Report.StartDate", "")
    gEndDateStr   = GetIniValue(cfg, "Report.EndDate", "")
    
    Dim tempOffset : tempOffset = Trim(GetIniValue(cfg, "Report.RelativeOffset", "0"))
    If IsNumeric(tempOffset) Then gRelOffset = CLng(tempOffset) Else gRelOffset = 0
    
    Dim tempLag : tempLag = Trim(GetIniValue(cfg, "Report.DailyLagDays", "1"))
    gDailyLagDays = 1
    If IsNumeric(tempLag) And CLng(tempLag) >= 0 Then gDailyLagDays = CLng(tempLag)
    
    gConnStr      = GetIniValue(cfg, "Database.ConnStr", "")
    gFullTableName = GetIniValue(cfg, "Database.TableName", "[dbo].[COB7_GL]")
    
    LogLine "DEBUG: Loaded Mode=" & gMode & ", RelOffset=" & gRelOffset & ", DailyLagDays=" & gDailyLagDays
End Sub

Sub ComputeDateRange(ByRef dStart, ByRef dEnd)
    On Error Resume Next
    Dim today : today = Date()
    Dim y, m, baseDate, tmp
    
    Dim cleanMode : cleanMode = UCase(Trim(gMode))
    LogLine "DEBUG: ComputeDateRange - Mode=[" & cleanMode & "] RelOffset=" & gRelOffset
    
    Select Case cleanMode
        Case "MONTHLY"
            If gYear <> "" And IsNumeric(gYear) Then y = CInt(gYear) Else y = Year(today)
            If gMonth <> "" And IsNumeric(gMonth) Then m = CInt(gMonth) Else m = Month(today)
            If gRelOffset <> 0 Then
                baseDate = DateSerial(y, m, 1)
                tmp = DateAdd("m", gRelOffset, baseDate)
                y = Year(tmp): m = Month(tmp)
            End If
            dStart = DateSerial(y, m, 1)
            dEnd = DateAdd("d", -1, DateAdd("m", 1, dStart))
            If dEnd > today Then dEnd = today
            
        Case "MTD"
            dStart = DateSerial(Year(today), Month(today), 1)
            dEnd = today
            ' Auto-lag for current MTD (like DAILY)
            If gRelOffset = 0 Then
                dEnd = DateAdd("d", -gDailyLagDays, today)
                LogLine "DEBUG: MTD current → applied lag → end date = " & Format(dEnd, "yyyy-mm-dd")
            End If
            
        Case "DAILY"
            If gRelOffset = 0 Then
                gRelOffset = -gDailyLagDays
                LogLine "DEBUG: DAILY current → applied lag → offset = " & gRelOffset
            End If
            dStart = DateAdd("d", gRelOffset, today)
            dEnd = dStart
            
        Case "CUSTOM"
            If Trim(gStartDateStr) <> "" And IsDate(gStartDateStr) Then dStart = CDate(gStartDateStr) Else dStart = today
            If Trim(gEndDateStr) <> "" And IsDate(gEndDateStr) Then dEnd = CDate(gEndDateStr) Else dEnd = today
            
        Case Else
            dStart = DateSerial(Year(today), Month(today), 1)
            dEnd = today
    End Select
    
    If Not IsDate(dStart) Or Not IsDate(dEnd) Then
        dStart = today: dEnd = today
    End If
    LogLine "DEBUG: Range: " & Format(dStart, "yyyy-mm-dd") & " to " & Format(dEnd, "yyyy-mm-dd")
    On Error GoTo 0
End Sub

Function FormatPeriodDisplay(startDate, endDate, mode)
    Select Case UCase(Trim(mode))
        Case "DAILY":   FormatPeriodDisplay = Format(startDate, "mmmm d, yyyy")
        Case "MTD":     FormatPeriodDisplay = "Month-to-Date " & Format(startDate, "mmmm yyyy")
        Case "MONTHLY": FormatPeriodDisplay = Format(startDate, "mmmm yyyy")
        Case "CUSTOM":  FormatPeriodDisplay = Format(startDate, "mmm d, yyyy") & " to " & Format(endDate, "mmm d, yyyy")
        Case Else:      FormatPeriodDisplay = Format(startDate, "mmm d, yyyy") & " to " & Format(endDate, "mmm d, yyyy")
    End Select
End Function

' ... (Pad, FileExists, ReadIniSafe, BuildTableNames, RemoveBrackets, LogLine functions remain unchanged from previous versions)

Function Pad(val, width)
    Dim s : If IsNull(val) Or IsEmpty(val) Then s = "" Else s = Trim(CStr(val))
    Pad = Left(s & Space(width), width)
End Function

Function FileExists(path)
    Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(path)
End Function

Function ReadIniSafe(path)
    On Error Resume Next
    Dim fso, ts, d, sect, line, p
    Set ReadIniSafe = Nothing
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then Exit Function
    Set ts = fso.OpenTextFile(path, 1)
    Set d = CreateObject("Scripting.Dictionary")
    sect = ""
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

Sub BuildTableNames(fullName, ByRef out2Part, ByRef outObjIdName)
    Dim s, parts, schemaName, tableName
    s = RemoveBrackets(Trim(fullName))
    parts = Split(s, ".")
    schemaName = "dbo"
    If UBound(parts) >= 1 Then schemaName = Trim(parts(0)) : tableName = Trim(parts(1))
    If UBound(parts) >= 2 Then schemaName = Trim(parts(UBound(parts)-1)) : tableName = Trim(parts(UBound(parts)))
    If Len(tableName)=0 Then Exit Sub
    out2Part = "[" & schemaName & "].[" & tableName & "]"
    outObjIdName = schemaName & "." & tableName
End Sub

Function RemoveBrackets(s)
    RemoveBrackets = Replace(Replace(s, "[", ""), "]", "")
End Function

Sub LogLine(msg)
    On Error Resume Next
    ' fout.WriteLine msg  ' uncomment if needed
    On Error GoTo 0
End Sub

Function GetIniValue(cfg, key, defaultValue)
    If cfg.Exists(key) And Len(Trim(cfg(key))) > 0 Then
        GetIniValue = Trim(cfg(key))
    Else
        GetIniValue = defaultValue
    End If
End Function