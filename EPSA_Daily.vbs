Option Explicit

' =========================================================
'  ACH Period Report (INI-driven) - Status='NP'
'  - Reads INI for Report params + Database.ConnStr + Database.TableName
'  - Computes date range from Mode / RelativeOffset
'  - Queries table by Post_Date BETWEEN start/end AND Status='NP'
'  - SELECT list matches ACH_EPSA columns (excluding id, Import_Timestamp)
' =========================================================

' ---------------- GLOBALS ----------------
Dim gMode, gRelOffset, gYear, gMonth, gQuarter
Dim gStartDateStr, gEndDateStr
Dim gConnStr, gFullTableName, gTable2Part, gObjIdName

Dim conn, rs
Dim INI_PATH: INI_PATH = "C:\byREQUEST\Templates\EPSA_DAILY.ini"
    INI_PATH = "S:\CarolinaTrust\byREQUEST\Templates\ACH_Trans.ini"

' ---------------------------------------------------------
'  Simple log helper
' ---------------------------------------------------------
Sub LogLine(msg)
    On Error Resume Next
    fout.WriteLine msg
    On Error GoTo 0
End Sub

' ---------------------------------------------------------
'  StartDoc
' ---------------------------------------------------------
Sub StartDoc
    On Error Resume Next

    Dim startKey, endKey
    Dim errMsg
    Dim sql, cmd
    Dim rowCount, totalAmount, amount
    Dim periodDisplay, startDateObj, endDateObj

    If Not FileExists(INI_PATH) Then
        LogLine "ERROR: INI file not found: " & INI_PATH
        Exit Sub
    End If

    LoadParameters INI_PATH

    If Trim(gConnStr) = "" Then
        LogLine "ERROR: ConnStr not defined in [Database] section."
        Exit Sub
    End If

    BuildTableNames gFullTableName, gTable2Part, gObjIdName
    If Len(gTable2Part) = 0 Then
        LogLine "ERROR: Could not parse TableName: " & gFullTableName
        Exit Sub
    End If

    ComputeDateRange startKey, endKey
    If startKey = "" Or endKey = "" Then
        LogLine "ERROR: Could not compute date range from parameters."
        Exit Sub
    End If

    startDateObj = FormatKeyToDate(startKey)
    endDateObj   = FormatKeyToDate(endKey)

    Select Case UCase(Trim(gMode))
        Case "DAILY"
            periodDisplay = Fmt(startDateObj, "dddd, mmmm d, yyyy")
        Case "WEEKLY"
            periodDisplay = "Week of " & Fmt(startDateObj, "mmm d") & " to " & Fmt(endDateObj, "mmm d, yyyy")
        Case "MONTHLY"
            periodDisplay = Fmt(startDateObj, "mmmm yyyy")
        Case "QUARTERLY"
            Dim quarterNum
            quarterNum = DatePart("q", startDateObj)
            periodDisplay = "Q" & quarterNum & " " & Year(startDateObj)
        Case "YTD"
            periodDisplay = "Year to Date " & Year(startDateObj)
        Case "CUSTOM"
            periodDisplay = Fmt(startDateObj, "mmm d, yyyy") & " to " & Fmt(endDateObj, "mmm d, yyyy")
        Case "MTD"
            periodDisplay = "MTD " & Fmt(startDateObj, "mmmm yyyy") & " (" & _
                            Fmt(startDateObj, "mmm d, yyyy") & " to " & Fmt(endDateObj, "mmm d, yyyy") & ")"
        Case Else
            periodDisplay = Fmt(startDateObj, "mmm d, yyyy") & " to " & Fmt(endDateObj, "mmm d, yyyy")
    End Select

    Spoolfile.User1 = UCase(gMode) & "-" & _
                      Right("0" & Month(endDateObj), 2) & _
                      Right("0" & Day(endDateObj), 2) & _
                      Year(endDateObj)

    'MsgBox Spoolfile.User1

    ' Open DB
    Set conn = CreateObject("ADODB.Connection")
    Err.Clear
    conn.Open gConnStr
    If Err.Number <> 0 Then
        LogLine "ERROR opening DB connection: " & Err.Description & " (Err=" & Err.Number & ")"
        Exit Sub
    End If

    ' =====================================================
    ' MAIN QUERY: Status='NP' (NO id, NO Import_Timestamp)
    ' =====================================================
    sql = "SELECT " & _
          "[Report],[Post_Date],[Post_Time],[Settlement_Date],[Company_Name]," & _
          "[ACH_Batch_ID],[PCTL],[LCTL],[Status],[Tran_Type],[Tran_Code],[Effective_Date]," & _
          "[Amount],[Member_Nbr],[Mbr_Name],[Fee],[ACH_Ind_Name],[DD],[Trace_Nbr],[Description]," & _
          "[Type],[Post_Seq],[File_Post_Seq],[Source_Name],[Check_Number],[SDC_Number],[S_L],[RR]," & _
          "[RDFI_ID],[DFI_ACCOUNT],[Individual_ID],[MESSAGE],[EPSA_File] " & _
          "FROM " & gTable2Part & " " & _
          "WHERE [Post_Date] BETWEEN ? AND ? " & _
          "AND ( [Status] <> 'NP' OR [Status] IS NULL ) " & _
          "ORDER BY [Post_Date], [Post_Time]"

    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = sql
        .CommandType = 1  ' adCmdText
        .Parameters.Append .CreateParameter("", 133, 1, , startDateObj) ' adDBDate
        .Parameters.Append .CreateParameter("", 133, 1, , endDateObj)
    End With

    Err.Clear
    Set rs = cmd.Execute
    If Err.Number <> 0 Then
        LogLine "ERROR executing query: " & Err.Description & " (Err=" & Err.Number & ")"
        CleanupObjects conn, rs, cmd
        Exit Sub
    End If

    If rs.EOF Then
        fout.WriteLine "No records found (Status <> 'NP' or Status IS NULL) for " & periodDisplay & "."
        CleanupObjects conn, rs, cmd
        Exit Sub
    End If

    ' Title
    'fout.WriteLine "ACH EPSA Report (Status <> NP) - " & periodDisplay

    ' Headers (fixed width)
    fout.WriteLine _
        Pad("Report", 10) & _
        Pad("Post_Date", 12) & _
        Pad("Post_Time", 12) & _
        Pad("Settlement", 16) & _
        Pad("Company_Name", 25) & _
        Pad("Batch_ID", 16) & _
        Pad("PCTL", 8) & _
        Pad("LCTL", 8) & _
        Pad("Status", 8) & _
        Pad("Tran_Type", 12) & _
        Pad("Tran_Code", 12) & _
        Pad("Effective", 16) & _
        Pad("Amount", 12) & _
        Pad("Member_Nbr", 15) & _
        Pad("Mbr_Name", 25) & _
        Pad("Fee", 10) & _
        Pad("ACH_Ind_Name", 25) & _
        Pad("DD", 4) & _
        Pad("Trace_Nbr", 20) & _
        Pad("Description", 30) & _
        Pad("Type", 8) & _
        Pad("Post_Seq", 10) & _
        Pad("File_Post_Seq", 15) & _
        Pad("Source_Name", 14) & _
        Pad("Check_Nbr", 12) & _
        Pad("SDC_Number", 12) & _
        Pad("S_L", 4) & _
        Pad("RR", 4) & _
        Pad("RDFI_ID", 10) & _
        Pad("DFI_ACCOUNT", 14) & _
        Pad("Individual_ID", 14) & _
        Pad("MESSAGE", 35) & _
        Pad("EPSA_File", 14)

    rowCount = 0
    totalAmount = 0

    Do While Not rs.EOF
        rowCount = rowCount + 1

        amount = 0
        If Not IsNull(rs("Amount")) And IsNumeric(rs("Amount")) Then
            amount = CDbl(rs("Amount"))
        End If
        totalAmount = totalAmount + amount

        fout.WriteLine _
            Pad(rs("Report"), 10) & _
            Pad(FmtDate8(rs("Post_Date")), 12) & _
            Pad(FmtTimeShort(rs("Post_Time")), 12) & _
            Pad(FmtDate8(rs("Settlement_Date")), 16) & _
            Pad(rs("Company_Name"), 25) & _
            Pad(rs("ACH_Batch_ID"), 16) & _
            Pad(rs("PCTL"), 8) & _
            Pad(rs("LCTL"), 8) & _
            Pad(rs("Status"), 8) & _
            Pad(rs("Tran_Type"), 12) & _
            Pad(rs("Tran_Code"), 12) & _
            Pad(FmtDate8(rs("Effective_Date")), 16) & _
            Pad(FormatNumber(amount, 2), 12) & _
            Pad(rs("Member_Nbr"), 15) & _
            Pad(rs("Mbr_Name"), 25) & _
            Pad(rs("Fee"), 10) & _
            Pad(rs("ACH_Ind_Name"), 25) & _
            Pad(rs("DD"), 4) & _
            Pad(rs("Trace_Nbr"), 20) & _
            Pad(rs("Description"), 30) & _
            Pad(rs("Type"), 8) & _
            Pad(rs("Post_Seq"), 10) & _
            Pad(rs("File_Post_Seq"), 15) & _
            Pad(rs("Source_Name"), 14) & _
            Pad(rs("Check_Number"), 12) & _
            Pad(rs("SDC_Number"), 12) & _
            Pad(rs("S_L"), 4) & _
            Pad(rs("RR"), 4) & _
            Pad(rs("RDFI_ID"), 10) & _
            Pad(rs("DFI_ACCOUNT"), 14) & _
            Pad(rs("Individual_ID"), 14) & _
            Pad(rs("MESSAGE"), 35) & _
            Pad(rs("EPSA_File"), 14)

        rs.MoveNext
    Loop

    fout.WriteLine ""
    'fout.WriteLine "Total Records: " & rowCount & " | Total Amount: " & FormatNumber(totalAmount, 2)

    ' ================================
    ' SECOND QUERY (NP matched to ERTR)
    ' Uses SAME startDateObj/endDateObj
    ' ================================
    Dim sql2, cmd2, rs2

    ' Using the first query version (with proper fields)
    sql2 = _
        "SELECT " & _
        "  'ERTR' AS Report, " & _
        "  ertr.Post_Date AS ERTR_Post_Date, " & _
        "  ertr.TC, " & _
        "  ertr.[RDFI_RT#], " & _
        "  ertr.[DFI_Account], " & _
        "  ertr.RTN_Amount, " & _
        "  ertr.Company_ID, " & _
        "  ertr.Individual_Name, " & _
        "  ertr.[DD] AS ERTR_DD, " & _
        "  ertr.A_Flag, " & _
        "  ertr.RRC, " & _
        "  ertr.ORC, " & _
        "  ertr.DRC, " & _
        "  ertr.[SRC_BH#], " & _
        "  ertr.[RTN_RDFI#], " & _
        "  ertr.DeathDate, " & _
        "  ertr.Info, " & _
        "  ertr.Trace_Number, " & _
        "  ertr.Company_Name AS ERTR_Company_Name, " & _
        "  ertr.Individual_ID AS ERTR_Individual_ID, " & _
        "  ertr.ERTR_File, " & _
        "  ertr.[RTN_Trace#], " & _                  
        "  'Matched return record' AS Match_Status " & _
        "FROM ( " & _
        "  SELECT Trace_Nbr, Post_Date, Amount, id, " & _
        "         ROW_NUMBER() OVER (PARTITION BY Trace_Nbr ORDER BY Post_Date DESC, id DESC) AS rn " & _
        "  FROM " & gTable2Part & " " & _
        "  WHERE Post_Date BETWEEN ? AND ? " & _
        ") epsa " & _
        "INNER JOIN [byREQUEST].[dbo].[ACH_ERTR] ertr " & _
        "  ON (ertr.Trace_Number = epsa.Trace_Nbr OR ertr.[RTN_Trace#] = epsa.Trace_Nbr) " & _
        "WHERE epsa.rn = 1 " & _
        "ORDER BY ertr.Post_Date DESC, ertr.Trace_Number"
		
		
		
	' --- TEST VERSION: Simple ERTR query  ---
		sql2 = "SELECT " & _
			  "  'ERTR' AS Report, " & _
			  "  ertr.[Post_Date] AS ERTR_Post_Date, " & _
			  "  ertr.[TC], " & _
			  "  ertr.[RDFI_RT#], " & _
			  "  ertr.[DFI_Account], " & _
			  "  ertr.[RTN_Amount], " & _
			  "  ertr.[Individual_ID] AS ERTR_Individual_ID, " & _
			  "  ertr.[Individual_Name], " & _
			  "  ertr.[DD] AS ERTR_DD, " & _
			  "  ertr.[A_Flag], " & _
			  "  ertr.[Trace_Number], " & _
			  "  ertr.[RRC], " & _
			  "  ertr.[ORC], " & _
			  "  ertr.[DRC], " & _
			  "  ertr.[SRC_BH#], " & _
			  "  ertr.[RTN_Trace#], " & _
			  "  ertr.[RTN_RDFI#], " & _
			  "  ertr.[DeathDate], " & _
			  "  ertr.[Info], " & _
			  "  ertr.[Company_Name] AS ERTR_Company_Name, " & _
			  "  ertr.[Company_ID], " & _
			  "  ertr.[ERTR_File] " & _
			  "FROM [byREQUEST].[dbo].[ACH_ERTR] ertr " & _
			  "WHERE ertr.[Post_Date] BETWEEN ? AND ? " & _
			  "ORDER BY ertr.[Post_Date] DESC"
		

    Set cmd2 = CreateObject("ADODB.Command")

    With cmd2
        .ActiveConnection = conn
        .CommandText = sql2
        .CommandType = 1  ' adCmdText
        .Parameters.Append .CreateParameter("", 133, 1, , startDateObj) ' adDBDate
        .Parameters.Append .CreateParameter("", 133, 1, , endDateObj)
    End With

    Err.Clear
    Set rs2 = cmd2.Execute
    If Err.Number <> 0 Then
        LogLine "ERROR executing Query2: " & Err.Description & " (Err=" & Err.Number & ")"
    Else
        'fout.WriteLine "ACH ERTR Report - " & periodDisplay

        If rs2.EOF Then
            fout.WriteLine "Report     No ERTR records found for the specified date range."
        Else
            ' --- ERTR Header (Clean ERTR-only) ---
            fout.WriteLine _
                Pad("Report", 10) & _
                Pad("Post_Date", 12) & _
                Pad("Company_ID", 28) & _
                Pad("Company_Name", 25) & _
                Pad("SRC_BH#", 16) & _
                Pad("TC", 8) & _
                Pad("DD", 8) & _
                Pad("A_Flag", 8) & _
                Pad("RRC", 12) & _
                Pad("ORC", 12) & _
                Pad("DRC", 16) & _
                Pad("RTN_Amount", 12) & _
                Pad("RTN_RDFI#", 15) & _
                Pad("Individual_ID", 35) & _
                Pad("Trace_Number", 29) & _
                Pad("RTN_Trace#", 20) & _
                Pad("Individual_Name", 48) & _
                Pad("RDFI_RT#", 15) & _
                Pad("DFI_Account", 15) & _
                Pad("DeathDate", 11) & _
                Pad("Info", 58) & _
                Pad("ERTR_File", 30)

            Do While Not rs2.EOF
                fout.WriteLine _
                    Pad(rs2("Report"), 10) & _
                    Pad(FmtDate8(rs2("ERTR_Post_Date")), 12) & _
                    Pad(rs2("Company_ID"), 28) & _
                    Pad(rs2("ERTR_Company_Name"), 25) & _
                    Pad(rs2("SRC_BH#"), 16) & _
                    Pad(rs2("TC"), 8) & _
                    Pad(rs2("ERTR_DD"), 8) & _
                    Pad(rs2("A_Flag"), 8) & _
                    Pad(rs2("RRC"), 12) & _
                    Pad(rs2("ORC"), 12) & _
                    Pad(rs2("DRC"), 16) & _
                    Pad(rs2("RTN_Amount"), 12) & _
                    Pad(rs2("RTN_RDFI#"), 15) & _
                    Pad(rs2("ERTR_Individual_ID"), 35) & _
                    Pad(rs2("Trace_Number"), 29) & _
                    Pad(rs2("RTN_Trace#"), 20) & _
                    Pad(rs2("Individual_Name"), 48) & _
                    Pad(rs2("RDFI_RT#"), 15) & _
                    Pad(rs2("DFI_Account"), 15) & _
                    Pad(rs2("DeathDate"), 11) & _
                    Pad(rs2("Info"), 58) & _
                    Pad(rs2("ERTR_File"), 30)

                rs2.MoveNext
            Loop
        End If
    End If

    CleanupObjects conn, rs, cmd
    CleanupObjects Nothing, rs2, cmd2

    On Error GoTo 0
End Sub

Sub ProcessLine
End Sub

Sub CloseDoc
End Sub

' =========================================================
'  Helpers
' =========================================================

Sub CleanupObjects(ByRef c, ByRef r, ByRef cm)
    On Error Resume Next
    If Not r Is Nothing Then
        If r.State = 1 Then r.Close
        Set r = Nothing
    End If
    If Not cm Is Nothing Then Set cm = Nothing
    If Not c Is Nothing Then
        If c.State = 1 Then c.Close
        Set c = Nothing
    End If
    On Error GoTo 0
End Sub

Sub LoadParameters(iniPath)
    Dim cfg
    Set cfg = ReadIniSafe(iniPath)
    If cfg Is Nothing Then
        LogLine "ERROR: Could not read INI file: " & iniPath
        Exit Sub
    End If

    gMode         = GetIniValue(cfg, "Report.Mode", "MONTHLY")
    gRelOffset    = GetIniValue(cfg, "Report.RelativeOffset", "0")
    gYear         = GetIniValue(cfg, "Report.Year", "")
    gMonth        = GetIniValue(cfg, "Report.Month", "")
    gQuarter      = GetIniValue(cfg, "Report.Quarter", "")
    gStartDateStr = GetIniValue(cfg, "Report.StartDate", "")
    gEndDateStr   = GetIniValue(cfg, "Report.EndDate", "")

    gConnStr       = GetIniValue(cfg, "Database.ConnStr", "")
    gFullTableName = GetIniValue(cfg, "Database.TableName", "[dbo].[ACH_EPSA]")

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

Sub ComputeDateRange(ByRef startKey, ByRef endKey)
    Dim mode, today, dStart, dEnd
    Dim y, m, q, tmp, dow, weekStart

    mode  = UCase(Trim(gMode))
    today = Date()

    Select Case mode
        Case "CUSTOM"
            If gStartDateStr <> "" Then dStart = CDate(gStartDateStr) Else dStart = today
            If gEndDateStr <> "" Then dEnd = CDate(gEndDateStr) Else dEnd = today

        Case "DAILY"
            dStart = DateAdd("d", gRelOffset, today)
            dEnd = dStart

        Case "WEEKLY"
            dow = Weekday(today, vbMonday)
            weekStart = DateAdd("d", 1 - dow, today)
            If gRelOffset <> 0 Then weekStart = DateAdd("ww", gRelOffset, weekStart)
            dStart = weekStart
            dEnd = DateAdd("d", 6, dStart)

        Case "QUARTERLY"
            If gYear <> "" And IsNumeric(gYear) Then y = CInt(gYear) Else y = Year(today)
            If gQuarter <> "" And IsNumeric(gQuarter) Then q = CInt(gQuarter) Else q = DatePart("q", today)

            If gRelOffset <> 0 Then
                q = q + gRelOffset
                While q > 4: q = q - 4: y = y + 1: Wend
                While q < 1: q = q + 4: y = y - 1: Wend
            End If

            m = (q - 1) * 3 + 1
            dStart = DateSerial(y, m, 1)
            dEnd = DateAdd("d", -1, DateAdd("m", 3, dStart))

        Case "YTD"
            If gYear <> "" And IsNumeric(gYear) Then y = CInt(gYear) Else y = Year(today)
            dStart = DateSerial(y, 1, 1)
            dEnd = today

		Case "MTD"
			Dim baseDate
			baseDate = today

			If gRelOffset <> 0 Then
				baseDate = DateAdd("m", gRelOffset, baseDate)
			End If

			dStart = DateSerial(Year(baseDate), Month(baseDate), 1)

			If Year(baseDate) = Year(today) And Month(baseDate) = Month(today) Then
				dEnd = today
			Else
				dEnd = DateAdd("d", -1, DateAdd("m", 1, dStart))
			End If


        Case Else ' MONTHLY
            If gYear <> "" And IsNumeric(gYear) Then y = CInt(gYear) Else y = Year(today)
            If gMonth <> "" And IsNumeric(gMonth) Then m = CInt(gMonth) Else m = Month(today)

            If gRelOffset <> 0 Then
                tmp = DateAdd("m", gRelOffset, DateSerial(y, m, 1))
                y = Year(tmp)
                m = Month(tmp)
            End If

            dStart = DateSerial(y, m, 1)
            dEnd = DateAdd("d", -1, DateAdd("m", 1, dStart))
    End Select

    If mode <> "CUSTOM" And mode <> "DAILY" And dEnd > today Then dEnd = today

    startKey = FormatDateKey(dStart)
    endKey   = FormatDateKey(dEnd)
End Sub

Function FormatDateKey(d)
    FormatDateKey = CStr(Year(d)) & Right("0" & Month(d), 2) & Right("0" & Day(d), 2)
End Function

Function FormatKeyToDate(ByVal yyyymmdd)
    Dim y, m, d
    If Len(yyyymmdd) <> 8 Then
        FormatKeyToDate = Date()
        Exit Function
    End If
    y = CInt(Left(yyyymmdd, 4))
    m = CInt(Mid(yyyymmdd, 5, 2))
    d = CInt(Right(yyyymmdd, 2))
    FormatKeyToDate = DateSerial(y, m, d)
End Function

Function Pad(val, width)
    Dim s
    If IsNull(val) Then s = "" Else s = CStr(val)
    s = Trim(s)
    Pad = Left(s & Space(width), width)
End Function

Function FmtDate8(val)
    On Error Resume Next
    If IsNull(val) Or Len(Trim(CStr(val))) = 0 Then
        FmtDate8 = ""
    ElseIf IsDate(val) Then
        FmtDate8 = Right("0" & Month(CDate(val)), 2) & "/" & _
                   Right("0" & Day(CDate(val)), 2) & "/" & _
                   CStr(Year(CDate(val)))
    Else
        FmtDate8 = Trim(CStr(val))
    End If
    On Error GoTo 0
End Function

Function FmtTimeShort(val)
    On Error Resume Next
    If IsNull(val) Then
        FmtTimeShort = ""
    Else
        FmtTimeShort = Trim(CStr(val))
    End If
    On Error GoTo 0
End Function

Function FileExists(path)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
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
        If Len(line) = 0 Or Left(line, 1) = ";" Then
            ' skip
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
        schemaName = Trim(parts(UBound(parts) - 1))
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


Function Fmt(d, mask)
    ' Tries VBA-style Format(); if not available, falls back to VBScript formatting
    On Error Resume Next

    Dim s
    s = ""

    ' Try VBA/host-provided Format first
    s = Format(d, mask)
    If Err.Number = 0 And Len(Trim(CStr(s))) > 0 Then
        Fmt = CStr(s)
        Exit Function
    End If

    ' Fallback
    Err.Clear
    If Not IsDate(d) Then
        Fmt = ""
        Exit Function
    End If

    Select Case LCase(mask)
        Case "dddd, mmmm d, yyyy"
            Fmt = WeekdayName(Weekday(d), False) & ", " & MonthName(Month(d), False) & " " & Day(d) & ", " & Year(d)
        Case "mmm d"
            Fmt = MonthName(Month(d), True) & " " & Day(d)
        Case "mmm d, yyyy"
            Fmt = MonthName(Month(d), True) & " " & Day(d) & ", " & Year(d)
        Case "mmmm yyyy"
            Fmt = MonthName(Month(d), False) & " " & Year(d)
        Case Else
            ' Generic safe default
            Fmt = MonthName(Month(d), True) & " " & Day(d) & ", " & Year(d)
    End Select

    On Error GoTo 0
End Function
