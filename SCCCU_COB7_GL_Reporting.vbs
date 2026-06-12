' =========================================================
' COB7_GL Period Detail Report (INI-driven)
' - Queries COB7_GL heap table (Run_Date is real DATE type)
' - Returns ALL raw rows in chronological order: ORDER BY Run_Date, ID
' =========================================================
Option Explicit

' ---------------- CONFIG ----------------
Dim INI_PATH: INI_PATH = "S:\CarolinaTrust\byREQUEST\Templates\COB7_GL_Monthly.ini"

' ---------------- GLOBALS ----------------
Dim gMode, gRelOffset, gYear, gMonth, gQuarter
Dim gStartDateStr, gEndDateStr
Dim gConnStr, gFullTableName, gTable2Part, gObjIdName
Dim conn, rs, cmd

' =========================================================
' StartDoc
' =========================================================
Private Sub StartDoc()
    On Error Resume Next
    Dim startDate, endDate, periodDisplay, errMsg
    
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
    Spoolfile.User1 = "COB7_GL " & UCase(Trim(gMode)) & " - " & periodDisplay
    
	' Connect to database
	Set conn = CreateObject("ADODB.Connection")   ' ← This line MUST exist and come BEFORE .Open
	Err.Clear                                      ' Clear any previous errors
	conn.Open gConnStr

	If Err.Number <> 0 Then
		LogLine "ERROR opening DB connection: " & Err.Description & " (Error #" & Err.Number & ")"
		LogLine "Attempted ConnStr (partial): " & Left(gConnStr, 60) & "..."
		' Optional: exit gracefully or continue with empty result
		Exit Sub
	Else
		LogLine "DEBUG: Database connection opened successfully"
	End If
		
    ' =====================================================
	' MAIN QUERY SECTION 
	' =====================================================
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
		.CommandType = 1   ' adCmdText
		.CommandText = sql
		
		' adDBTimeStamp = 7 → best for SQL Server DATE columns
		.Parameters.Append .CreateParameter("@StartDate", 133, 1, , startDate)
		.Parameters.Append .CreateParameter("@EndDate",   133, 1, , endDate)
	End With

	' === Correct & standard way to open recordset with parameterized query ===
	Set rs = CreateObject("ADODB.Recordset")
	rs.CursorType   = 0   ' adOpenForwardOnly – efficient for reading only
	rs.LockType     = 1   ' adLockReadOnly
	rs.Open cmd       ' This opens the recordset using the command object

	If Err.Number <> 0 Then
		fout.WriteLine "ERROR opening recordset: " & Err.Description & " (Error " & Err.Number & ")"
		Err.Clear
		' Continue to cleanup
	End If

	If rs.EOF And rs.BOF Then
		fout.WriteLine "No COB7_GL records found for period: " & periodDisplay
	Else
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
			
		
		Dim periodStr
		periodStr = Format(dStart, "yyyymm") & "-" & Format(dEnd, "yyyymm")
		If gMode = "DAILY" Then periodStr = Format(dStart, "yyyymmdd")
		If gMode = "MTD" Then periodStr = "MTD_" & Format(Date, "yyyymm")

		SpoolFile.User1 = gMode 
		SpoolFile.User2 = periodStr  
	
		
	End If

	' Cleanup - no label needed
	On Error Resume Next
	If Not rs Is Nothing Then
		If rs.State <> 0 Then rs.Close
		Set rs = Nothing
	End If
	If Not cmd Is Nothing Then Set cmd = Nothing
	If Not conn Is Nothing Then
		If conn.State <> 0 Then conn.Close
		Set conn = Nothing
	End If
	On Error GoTo 0
	
End Sub

' =========================================================
' Required helper functions (all included)
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
    
    ' Safe handling of RelativeOffset (handles negatives properly)
    Dim tempOffset
    tempOffset = Trim(GetIniValue(cfg, "Report.RelativeOffset", "0"))
    If IsNumeric(tempOffset) Then
        gRelOffset = CLng(tempOffset)
    Else
        gRelOffset = 0
        LogLine "WARNING: RelativeOffset invalid: '" & tempOffset & "' - defaulting to 0"
    End If
    
    gConnStr      = GetIniValue(cfg, "Database.ConnStr", "")
    gFullTableName = GetIniValue(cfg, "Database.TableName", "[dbo].[COB7_GL]")
    
    ' Optional debug (remove later)
    LogLine "DEBUG: Loaded RelativeOffset = " & gRelOffset
End Sub

Function GetIniValue(cfg, key, defaultValue)
    If cfg.Exists(key) And Len(Trim(cfg(key))) > 0 Then
        GetIniValue = Trim(cfg(key))
    Else
        GetIniValue = defaultValue
    End If
End Function

Sub ComputeDateRange(ByRef dStart, ByRef dEnd)
    On Error Resume Next
    Dim today : today = Date()
    Dim y, m, baseDate, tmp
    
    ' Normalize mode
    Dim cleanMode : cleanMode = UCase(Trim(gMode))
    LogLine "DEBUG: Entering ComputeDateRange - Mode=[" & cleanMode & "] RelOffset=" & gRelOffset
    
    Select Case cleanMode
        Case "MONTHLY"
            LogLine "DEBUG: Processing MONTHLY mode"
            
            ' Determine base year/month
            If gYear <> "" And IsNumeric(gYear) Then
                y = CInt(gYear)
            Else
                y = Year(today)
            End If
            
            If gMonth <> "" And IsNumeric(gMonth) Then
                m = CInt(gMonth)
            Else
                m = Month(today)
            End If
            
            LogLine "DEBUG: Base y/m = " & y & "/" & m
            
            ' Apply relative offset if present
            If gRelOffset <> 0 Then
                LogLine "DEBUG: Applying RelativeOffset " & gRelOffset & " months"
                baseDate = DateSerial(y, m, 1)
                tmp = DateAdd("m", gRelOffset, baseDate)
                y = Year(tmp)
                m = Month(tmp)
                LogLine "DEBUG: After offset → y/m = " & y & "/" & m
            End If
            
            ' Set range
            dStart = DateSerial(y, m, 1)
            dEnd = DateAdd("d", -1, DateAdd("m", 1, dStart))
            
            If dEnd > today Then
                LogLine "DEBUG: Capping end date to today"
                dEnd = today
            End If
            
        Case "MTD"
            LogLine "DEBUG: Processing MTD mode"
            dStart = DateSerial(Year(today), Month(today), 1)
            dEnd = today
            
        Case "DAILY"
            LogLine "DEBUG: Processing DAILY mode"
            dStart = DateAdd("d", gRelOffset, today)
            dEnd = dStart
            
        Case "CUSTOM"
            LogLine "DEBUG: Processing CUSTOM mode"
            
            ' --- Safe conversion for StartDate ---
            If Trim(gStartDateStr) <> "" And IsDate(gStartDateStr) Then
                dStart = CDate(gStartDateStr)
                LogLine "DEBUG: Custom StartDate set to " & Format(dStart, "yyyy-mm-dd")
            Else
                dStart = today
                LogLine "DEBUG: Custom StartDate invalid or missing → using today"
            End If
            
            ' --- Safe conversion for EndDate ---
            If Trim(gEndDateStr) <> "" And IsDate(gEndDateStr) Then
                dEnd = CDate(gEndDateStr)
                LogLine "DEBUG: Custom EndDate set to " & Format(dEnd, "yyyy-mm-dd")
            Else
                dEnd = today
                LogLine "DEBUG: Custom EndDate invalid or missing → using today"
            End If
            
        Case Else
            LogLine "WARNING: Unknown mode '" & gMode & "' - defaulting to current month"
            dStart = DateSerial(Year(today), Month(today), 1)
            dEnd = today
    End Select
    
    ' Final validation
    If Err.Number <> 0 Then
        LogLine "ERROR in date calculation: " & Err.Description & " (Err=" & Err.Number & ")"
        Err.Clear
    End If
    
    If Not IsDate(dStart) Or Not IsDate(dEnd) Then
        LogLine "ERROR: Date range still invalid after calculation - forcing today"
        dStart = today
        dEnd = today
    End If
    
    LogLine "DEBUG: Final computed range: " & Format(dStart, "yyyy-mm-dd") & " to " & Format(dEnd, "yyyy-mm-dd")
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

Function Pad(val, width)
    Dim s
    If IsNull(val) Or IsEmpty(val) Then
        s = ""
    Else
        s = Trim(CStr(val))
    End If
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
    'fout.WriteLine msg
    On Error GoTo 0
End Sub