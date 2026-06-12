Option Explicit

' =========================================================
'  ACH Period Report (Monthly / Weekly / Quarterly / Custom / YTD)
'  - Reads parameters.ini
'  - Computes date range from Mode / RelativeOffset
'  - Queries ACH_EPSA_Staging table using DateCol (from spool filename)
'  - Writes fixed-width, column-aligned output
' =========================================================

' Globals
Dim gMode, gRelOffset, gYear, gMonth, gQuarter
Dim gStartDateStr, gEndDateStr
Dim gConnStr, gFullTableName, gTable2Part, gObjIdName

' ADO objects
Dim conn, rs

Dim INI_PATH: INI_PATH = "S:\CarolinaTrust\byREQUEST\Templates\ACH_Daily.ini"


' ---------------------------------------------------------
'  Simple debug helper
' ---------------------------------------------------------
Sub LogLine(msg)
    On Error Resume Next
    fout.WriteLine msg
    On Error GoTo 0
End Sub

' ---------------------------------------------------------
'  StartDoc: run entire report here
' ---------------------------------------------------------
Sub StartDoc


    On Error Resume Next

    Dim startKey, endKey
    Dim errMsg

    'LogLine "========================================="
    'LogLine "ACH EPSA Report Started: " & Now()
    'LogLine "INI_PATH = " & INI_PATH

    If Not FileExists(INI_PATH) Then
        errMsg = "ERROR: INI file not found: " & INI_PATH
        LogLine errMsg
        Exit Sub
    End If

    ' Load parameters from ini
    LoadParameters INI_PATH
    'LogLine "Mode=" & gMode & ", RelOffset=" & gRelOffset & ", Year=" & gYear & _
    '         ", Month=" & gMonth & ", Quarter=" & gQuarter & _
     '        ", StartDate=" & gStartDateStr & ", EndDate=" & gEndDateStr



    If Trim(gConnStr) = "" Then
        errMsg = "ERROR: ConnStr not defined in [Database] section."
        LogLine errMsg
        Exit Sub
    End If

    ' Parse table name
    BuildTableNames gFullTableName, gTable2Part, gObjIdName
    If Len(gTable2Part) = 0 Or Len(gObjIdName) = 0 Then
        LogLine "ERROR: Could not parse TableName: " & gFullTableName
        Exit Sub
    End If


    ' Compute date range (YYYYMMDD) for DateCol
    ComputeDateRange startKey, endKey
    'LogLine "Date range: " & startKey & " to " & endKey

    If startKey = "" Or endKey = "" Then
        errMsg = "ERROR: Could not compute date range from parameters."
        LogLine errMsg
        Exit Sub
    End If

    ' Header info
    Spoolfile.user1 = "ACH EPSA Report - Mode: " & gMode & "  RelativeOffset: " & gRelOffset & _
                   "  Date Range: " & startKey & " to " & endKey
    'LogLine Spoolfile.user1

    ' Open connection
    Set conn = CreateObject("ADODB.Connection")
    Err.Clear
    conn.Open gConnStr

    If Err.Number <> 0 Then
        errMsg = "ERROR opening DB connection: " & Err.Description & " (Err=" & Err.Number & ")"
        LogLine errMsg
        Exit Sub
    End If
	
	
    ' =====================================================
    ' Format the period for Spoolfile.User1
    ' =====================================================
    Dim periodDisplay
    Dim startDateObj, endDateObj
    
    ' Convert YYYYMMDD to date objects for display
    startDateObj = FormatKeyToDate(startKey)
    endDateObj = FormatKeyToDate(endKey)
    
    ' Format based on mode
    Select Case UCase(Trim(gMode))
        Case "DAILY"
            periodDisplay = Format(startDateObj, "dddd, mmmm d, yyyy")
        Case "WEEKLY"
            periodDisplay = "Week of " & Format(startDateObj, "mmm d") & " to " & Format(endDateObj, "mmm d, yyyy")
        Case "MONTHLY"
            periodDisplay = Format(startDateObj, "mmmm yyyy")
        Case "QUARTERLY"
            Dim quarterNum
            quarterNum = DatePart("q", startDateObj)
            periodDisplay = "Q" & quarterNum & " " & Year(startDateObj)
        Case "YTD"
            periodDisplay = "Year to Date " & Year(startDateObj)
        Case "CUSTOM"
            periodDisplay = Format(startDateObj, "mmm d, yyyy") & " to " & Format(endDateObj, "mmm d, yyyy")
        Case Else
            periodDisplay = Format(startDateObj, "mmm d, yyyy") & " to " & Format(endDateObj, "mmm d, yyyy")
    End Select
    
    ' Set Spoolfile.User1 with Mode and Period
    Spoolfile.User1 = gMode & " - " & periodDisplay
    'LogLine "Spoolfile.User1 = " & Spoolfile.User1
	
	
	
	
	
	
    'LogLine "SUCCESS: DB connection opened"

	' =====================================================
	' MAIN QUERY: Get data from ACH_EPSA table
	' All date columns are NVARCHAR, so we need to parse them
	' =====================================================
	Dim sql, cmd, rowCount

	' Convert our YYYYMMDD keys to the format stored in Post_Date (MM-DD-YY)
	Dim startDateFormatted, endDateFormatted
	startDateFormatted = FormatDateForPostDate(startKey)
	endDateFormatted = FormatDateForPostDate(endKey)

	'LogLine "Looking for Post_Date between: " & startDateFormatted & " and " & endDateFormatted

	' First try: Use TRY_CONVERT to convert Post_Date string to DATE for comparison
	sql = "SELECT " & _
		  "[Report], [Post_Date], [Post_Time], [Settlement_Date], [Company_Name], " & _
		  "[ACH_Batch_ID], [PCTL], [LCTL], [Status], [Tran_Type], [Tran_Code], " & _
		  "[Effective_Date], [Amount], [Member_Nbr], [Acct], [Mbr_Name], [Fee], " & _
		  "[ACH_Ind_Name], [DD], [Trace_Nbr], [Description], [Type], " & _
		  "[Post_Seq], [File_Post_Seq], [Source_Name], [Batch], [Check_Number], " & _
		  "[SDC_Number], [S_L], [RR], [RDFI_ID], [File_Acc_Number], [Individual_ID], " & _
		  "[MESSAGE], [HashKey], [Source_Message], [Import_Timestamp] " & _
		  "FROM " & gTable2Part & " " & _
		  "WHERE TRY_CONVERT(DATE, [Post_Date], 1) IS NOT NULL " & _
		  "AND TRY_CONVERT(DATE, [Post_Date], 1) BETWEEN ? AND ? " & _
		  "ORDER BY TRY_CONVERT(DATE, [Post_Date], 1), [Member_Nbr], [Trace_Nbr]"

	'LogLine "SQL Query (TRY_CONVERT): " & sql

	Set cmd = CreateObject("ADODB.Command")
	With cmd
		.ActiveConnection = conn
		.CommandText = sql
		.CommandType = 1   ' adCmdText
		
		' Add parameters as DATE type
		.Parameters.Append .CreateParameter("", 135, 1, , CDate(startDateFormatted)) ' adDBTimeStamp
		.Parameters.Append .CreateParameter("", 135, 1, , CDate(endDateFormatted))   ' adDBTimeStamp
	End With

	Err.Clear
	Set rs = cmd.Execute

	If Err.Number <> 0 Then
		
		errMsg = "ERROR executing TRY_CONVERT query: " & Err.Description & " (Err=" & Err.Number & ")"
		LogLine errMsg
		
		' Second try: Use CONVERT with style 1 (MM/DD/YY)
		sql = "SELECT " & _
			  "[Report], [Post_Date], [Post_Time], [Settlement_Date], [Company_Name], " & _
			  "[ACH_Batch_ID], [PCTL], [LCTL], [Status], [Tran_Type], [Tran_Code], " & _
			  "[Effective_Date], [Amount], [Member_Nbr], [Acct], [Mbr_Name], [Fee], " & _
			  "[ACH_Ind_Name], [DD], [Trace_Nbr], [Description], [Type], " & _
			  "[Post_Seq], [File_Post_Seq], [Source_Name], [Batch], [Check_Number], " & _
			  "[SDC_Number], [S_L], [RR], [RDFI_ID], [File_Acc_Number], [Individual_ID], " & _
			  "[MESSAGE], [HashKey], [Source_Message], [Import_Timestamp] " & _
			  "FROM " & gTable2Part & " " & _
			  "WHERE ISDATE([Post_Date]) = 1 " & _
			  "AND CONVERT(DATE, [Post_Date], 1) BETWEEN CONVERT(DATE, ?, 1) AND CONVERT(DATE, ?, 1) " & _
			  "ORDER BY CONVERT(DATE, [Post_Date], 1), [Member_Nbr], [Trace_Nbr]"
		
		'LogLine "SQL Query (CONVERT with ISDATE): " & sql
		
		Set cmd = CreateObject("ADODB.Command")
		With cmd
			.ActiveConnection = conn
			.CommandText = sql
			.CommandType = 1
			.Parameters.Append .CreateParameter("", 135, 1, , CDate(startDateFormatted))
			.Parameters.Append .CreateParameter("", 135, 1, , CDate(endDateFormatted))
		End With
		
		Err.Clear
		Set rs = cmd.Execute
		
		If Err.Number <> 0 Then
			errMsg = "ERROR executing CONVERT query: " & Err.Description & " (Err=" & Err.Number & ")"
			LogLine errMsg
			
			' Third try: Get all data and filter in VBScript
			LogLine "Falling back to get all data and filter in VBScript"
			sql = "SELECT " & _
				  "[Report], [Post_Date], [Post_Time], [Settlement_Date], [Company_Name], " & _
				  "[ACH_Batch_ID], [PCTL], [LCTL], [Status], [Tran_Type], [Tran_Code], " & _
				  "[Effective_Date], [Amount], [Member_Nbr], [Acct], [Mbr_Name], [Fee], " & _
				  "[ACH_Ind_Name], [DD], [Trace_Nbr], [Description], [Type], " & _
				  "[Post_Seq], [File_Post_Seq], [Source_Name], [Batch], [Check_Number], " & _
				  "[SDC_Number], [S_L], [RR], [RDFI_ID], [File_Acc_Number], [Individual_ID], " & _
				  "[MESSAGE], [HashKey], [Source_Message], [Import_Timestamp] " & _
				  "FROM " & gTable2Part & " " & _
				  "ORDER BY [Post_Date], [Member_Nbr], [Trace_Nbr]"
			
			LogLine "Getting all data: " & sql
			
			Set rs = conn.Execute(sql)
			
			If Err.Number <> 0 Then
				errMsg = "ERROR getting all data: " & Err.Description & " (Err=" & Err.Number & ")"
				LogLine errMsg
				Exit Sub
			End If
		End If
	End If

    If rs.EOF Then
        LogLine "No records found for the specified date range."
        fout.WriteLine "No records found for the specified date range."
        Exit Sub
    End If

    ' Column headers (aligned)
    fout.WriteLine _
        Pad("Report", 10) & _
        Pad("Post_Date", 12) & _
        Pad("Post_Time", 12) & _
        Pad("Settlement_Date", 16) & _
        Pad("Company_Name", 25) & _
        Pad("ACH_Batch_ID", 15) & _
        Pad("PCTL", 8) & _
        Pad("LCTL", 8) & _
        Pad("Status", 8) & _
        Pad("Tran_Type", 12) & _
        Pad("Tran_Code", 12) & _
        Pad("Effective_Date", 16) & _
        Pad("Amount", 12) & _
        Pad("Member_Nbr", 15) & _
        Pad("Acct", 8) & _
        Pad("Mbr_Name", 25) & _
        Pad("Fee", 10) & _
        Pad("ACH_Ind_Name", 25) & _
        Pad("DD", 6) & _
        Pad("Trace_Nbr", 20) & _
        Pad("Description", 30) & _
        Pad("Type", 8) & _
        Pad("Post_Seq", 10) & _
        Pad("File_Post_Seq", 15)


    ' Process rows
    rowCount = 0
    Dim totalAmount
    totalAmount = 0

    Do While Not rs.EOF
        rowCount = rowCount + 1
        
        ' Get amount for totaling
        Dim amount
        amount = 0
        If Not IsNull(rs("Amount")) Then
            If IsNumeric(rs("Amount")) Then
                amount = CDbl(rs("Amount"))
            End If
        End If
        totalAmount = totalAmount + amount

        ' Write row
        On Error Resume Next
        fout.WriteLine _
            Pad(rs("Report"), 10) & _
            Pad(rs("Post_Date"), 12) & _
            Pad(rs("Post_Time"), 12) & _
            Pad(rs("Settlement_Date"), 16) & _
            Pad(rs("Company_Name"), 25) & _
            Pad(rs("ACH_Batch_ID"), 15) & _
            Pad(rs("PCTL"), 8) & _
            Pad(rs("LCTL"), 8) & _
            Pad(rs("Status"), 8) & _
            Pad(rs("Tran_Type"), 12) & _
            Pad(rs("Tran_Code"), 12) & _
            Pad(rs("Effective_Date"), 16) & _
            Pad(FormatNumber(amount, 2), 12) & _
            Pad(rs("Member_Nbr"), 15) & _
            Pad(rs("Acct"), 8) & _
            Pad(rs("Mbr_Name"), 25) & _
            Pad(rs("Fee"), 10) & _
            Pad(rs("ACH_Ind_Name"), 25) & _
            Pad(rs("DD"), 6) & _
            Pad(rs("Trace_Nbr"), 20) & _
            Pad(rs("Description"), 30) & _
            Pad(rs("Type"), 8) & _
            Pad(rs("Post_Seq"), 10) & _
            Pad(rs("File_Post_Seq"), 15)

        If Err.Number <> 0 Then
            LogLine "ERROR writing row " & rowCount & ": " & Err.Description
            Err.Clear
        End If

        rs.MoveNext
    Loop

    ' Summary line
    'fout.WriteLine "Total Records: " & rowCount & " | Total Amount: " & FormatNumber(totalAmount, 2)
    'fout.WriteLine "Date Range: " & startKey & " to " & endKey & " | Mode: " & gMode & " | Relative Offset: " & gRelOffset

    ' Close objects
    rs.Close
    conn.Close
    
    Set rs = Nothing
    Set cmd = Nothing
    Set conn = Nothing

    'LogLine "Report completed. Total rows: " & rowCount & " | Total amount: " & FormatNumber(totalAmount, 2)
    'LogLine "ACH EPSA Report Completed: " & Now()

    On Error GoTo 0
End Sub

' ---------------------------------------------------------
'  ProcessLine: not needed for this pattern
' ---------------------------------------------------------
Sub ProcessLine
    ' No per-line processing required
End Sub

' ---------------------------------------------------------
'  CloseDoc: nothing to do (we closed everything in StartDoc)
' ---------------------------------------------------------
Sub CloseDoc
    ' Optional: could log footer here if you want
End Sub

' =========================================================
'  Helpers
' =========================================================

' --- Load parameters from INI file ---
Sub LoadParameters(iniPath)
    Dim cfg, key
    
    Set cfg = ReadIniSafe(iniPath)
    If cfg Is Nothing Then
        LogLine "ERROR: Could not read INI file"
        Exit Sub
    End If
    
    ' [Report] section
    gMode = GetIniValue(cfg, "Report.Mode", "MONTHLY")
    gRelOffset = GetIniValue(cfg, "Report.RelativeOffset", "0")
    gYear = GetIniValue(cfg, "Report.Year", "")
    gMonth = GetIniValue(cfg, "Report.Month", "")
    gQuarter = GetIniValue(cfg, "Report.Quarter", "")
    gStartDateStr = GetIniValue(cfg, "Report.StartDate", "")
    gEndDateStr = GetIniValue(cfg, "Report.EndDate", "")
    
    ' [Database] section
    gConnStr = GetIniValue(cfg, "Database.ConnStr", "")
    gFullTableName = GetIniValue(cfg, "Database.TableName", "[dbo].[ACH_EPSA_Staging]")
    
    ' Convert RelativeOffset to integer
    If IsNumeric(gRelOffset) Then
        gRelOffset = CInt(gRelOffset)
    Else
        gRelOffset = 0
    End If
End Sub

Function GetIniValue(cfg, key, defaultValue)
    If Not cfg Is Nothing And cfg.Exists(key) And Len(cfg(key)) > 0 Then
        GetIniValue = Trim(cfg(key))
    Else
        GetIniValue = defaultValue
    End If
End Function

' --- Compute date range (startKey, endKey) based on Mode & parameters ---
Sub ComputeDateRange(ByRef startKey, ByRef endKey)
    Dim mode, today, dStart, dEnd
    Dim y, m, q, tmp, dow, weekStart

    mode  = UCase(Trim(gMode))
    today = Date()

    'LogLine "ComputeDateRange: mode=" & mode & ", today=" & CStr(today) & ", RelOffset=" & gRelOffset

    Select Case mode

        Case "CUSTOM"
            If gStartDateStr <> "" Then
                dStart = CDate(gStartDateStr)
            Else
                dStart = today
            End If

            If gEndDateStr <> "" Then
                dEnd = CDate(gEndDateStr)
            Else
                dEnd = today
            End If

        Case "DAILY"
            ' Daily mode: get a specific day
            If gRelOffset <> 0 Then
                ' Use relative offset (0 = today, -1 = yesterday, 1 = tomorrow, etc.)
                dStart = DateAdd("d", gRelOffset, today)
            Else
                dStart = today
            End If
            dEnd = dStart ' Same day for daily mode

        Case "WEEKLY"
            dow       = Weekday(today, vbMonday)   ' 1=Mon, 7=Sun
            weekStart = DateAdd("d", 1 - dow, today)
            If gRelOffset <> 0 Then
                weekStart = DateAdd("ww", gRelOffset, weekStart)
            End If
            dStart = weekStart
            dEnd   = DateAdd("d", 6, dStart)

        Case "QUARTERLY"
            If gYear <> "" And IsNumeric(gYear) Then
                y = CInt(gYear)
            Else
                y = Year(today)
            End If

            If gQuarter <> "" And IsNumeric(gQuarter) Then
                q = CInt(gQuarter)
            Else
                q = DatePart("q", today) ' Get current quarter (1-4)
            End If

            If gRelOffset <> 0 Then
                ' Adjust quarter based on offset
                q = q + gRelOffset
                ' Handle year rollover
                While q > 4
                    q = q - 4
                    y = y + 1
                Wend
                While q < 1
                    q = q + 4
                    y = y - 1
                Wend
            End If

            m = (q - 1) * 3 + 1
            dStart = DateSerial(y, m, 1)
            dEnd   = DateAdd("d", -1, DateAdd("m", 3, dStart))

        Case "YTD"
            If gYear <> "" And IsNumeric(gYear) Then
                y = CInt(gYear)
            Else
                y = Year(today)
            End If
            dStart = DateSerial(y, 1, 1)
            dEnd   = today

        Case Else   ' MONTHLY default
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

            ' Apply RelativeOffset
            If gRelOffset <> 0 Then
                ' Add offset months
                tmp = DateAdd("m", gRelOffset, DateSerial(y, m, 1))
                y = Year(tmp)
                m = Month(tmp)
            End If

            dStart = DateSerial(y, m, 1)
            dEnd = DateAdd("d", -1, DateAdd("m", 1, dStart))
    End Select

    ' Ensure end date doesn't exceed today for non-CUSTOM modes
    If mode <> "CUSTOM" And mode <> "DAILY" And dEnd > today Then
        dEnd = today
    End If

    startKey = FormatDateKey(dStart)
    endKey   = FormatDateKey(dEnd)

    'LogLine "ComputeDateRange: dStart=" & CStr(dStart) & ", dEnd=" & CStr(dEnd) & ", startKey=" & startKey & ", endKey=" & endKey
End Sub

' --- Format VBScript Date as "YYYYMMDD" for DateCol comparison ---
Function FormatDateKey(d)
    FormatDateKey = CStr(Year(d)) & _
                    Right("0" & Month(d), 2) & _
                    Right("0" & Day(d),   2)
End Function

' --- Convert YYYYMMDD to date object for comparison ---
Function FormatDateForPostDate(yyyymmdd)
    If Len(yyyymmdd) = 8 Then
        Dim y, m, d
        y = CInt(Left(yyyymmdd, 4))
        m = CInt(Mid(yyyymmdd, 5, 2))
        d = CInt(Right(yyyymmdd, 2))
        
        ' Create a date object
        FormatDateForPostDate = DateSerial(y, m, d)
    Else
        FormatDateForPostDate = Date()
    End If
End Function

' --- Column padding helper: left-aligned, fixed width ---
Function Pad(val, width)
    Dim s
    If IsNull(val) Then
        s = ""
    Else
        s = CStr(val)
    End If
    s = Trim(s)
    Pad = Left(s & Space(width), width)
End Function

' --- Simple file existence helper ---
Function FileExists(path)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(path)
End Function

' --- INI reading helper (from previous code) ---
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
            ' skip
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

' --- Table name parsing helper (from previous code) ---
Sub BuildTableNames(ByVal fullName, ByRef out2Part, ByRef outObjIdName)
    Dim s, parts, schemaName, tableName
    s = RemoveBrackets(Trim(fullName))
    parts = Split(s, ".")

    schemaName = "dbo"
    tableName = ""

    If UBound(parts) = 0 Then
        tableName = Trim(parts(0))
    ElseIf UBound(parts) = 1 Then
        schemaName = Trim(parts(0))
        tableName  = Trim(parts(1))
    Else
        schemaName = Trim(parts(UBound(parts)-1))
        tableName  = Trim(parts(UBound(parts)))
    End If

    If Len(tableName) = 0 Then
        out2Part = ""
        outObjIdName = ""
        Exit Sub
    End If

    out2Part = "[" & schemaName & "].[" & tableName & "]"
    outObjIdName = schemaName & "." & tableName
End Sub

Function RemoveBrackets(ByVal s)
    s = Replace(s, "[", "")
    s = Replace(s, "]", "")
    RemoveBrackets = s
End Function