Option Explicit

' ---------------- GLOBALS ----------------
Dim gMode, gRelOffset, gYear, gMonth, gQuarter
Dim gStartDateStr, gEndDateStr
Dim gConnStr
Dim gFactor
Dim gDaysDQHigh, gDaysDQMidHigh, gDaysDQMidLow
Dim gPctHigh, gPctMidHigh, gPctMidLow, gPctLow

Dim conn
Dim INI_PATH: INI_PATH = "S:\CarolinaTrust\byREQUEST\Templates\DQ_ICR_Monthly.ini"

Sub LogLine(msg)
    On Error Resume Next
    fout.WriteLine msg
    On Error GoTo 0
End Sub

Sub StartDoc()
    On Error Resume Next

    Dim startKey, endKey
    Dim startDateObj, endDateObj

    If Not FileExists(INI_PATH) Then
        LogLine "ERROR: INI file not found: " & INI_PATH
        Exit Sub
    End If

    LoadParameters INI_PATH

    If Trim(gConnStr) = "" Then
        LogLine "ERROR: ConnStr not defined in [Database] section."
        Exit Sub
    End If

    If Not ValidateDaysDQConfig() Then Exit Sub

    ComputeDateRange startKey, endKey
    If startKey = "" Or endKey = "" Then
        LogLine "ERROR: Could not compute date range from parameters."
        Exit Sub
    End If

    startDateObj = FormatKeyToDate(startKey)
    endDateObj   = FormatKeyToDate(endKey)

    Set conn = CreateObject("ADODB.Connection")
    Err.Clear
    conn.Open gConnStr
    If Err.Number <> 0 Then
        LogLine "ERROR opening DB connection: " & Err.Description & " (Err=" & Err.Number & ")"
        Exit Sub
    End If
    On Error GoTo 0

    RunSection "SECURED DELINQUENCY", "dbo.DQ_Secured", True, startDateObj, endDateObj

    fout.WriteLine ""
    fout.WriteLine ""

    RunSection "UNSECURED DELINQUENCY", "dbo.DQ_UnSecured", False, startDateObj, endDateObj

    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
End Sub

Sub ProcessLine()
End Sub

Sub CloseDoc()
End Sub

Sub RunSection(ByVal sectionTitle, ByVal tableName, ByVal hasSFX, ByVal startDateObj, ByVal endDateObj)
    On Error Resume Next

    Dim sql, cmd, rs
    Dim sfxExpr
    Dim factorText, dqPctText
    Dim nm, acct, sfx, loanDesc, bal, dlp, daysDQ, dqPct
    Dim projDef, projLoss, rowSort

    If hasSFX Then
        sfxExpr = "ISNULL(LTRIM(RTRIM(IDValue)), '')"
    Else
        sfxExpr = "CAST('' AS NVARCHAR(20))"
    End If

  sql = ""
	sql = sql & "WITH base AS ("
	sql = sql & " SELECT "
	sql = sql & "   ISNULL(LTRIM(RTRIM(FullName)), '') AS FullName, "
	sql = sql & "   ISNULL(LTRIM(RTRIM(AccountNo)), '') AS AccountNo, "
	sql = sql & "   " & sfxExpr & " AS SFX, "
	sql = sql & "   ISNULL(LTRIM(RTRIM(LoanType)), '') AS LoanType, "
	sql = sql & "   CASE "
	sql = sql & "     WHEN TRY_CONVERT(int, LTRIM(RTRIM(LoanType))) BETWEEN 1 AND 10 OR TRY_CONVERT(int, LTRIM(RTRIM(LoanType))) BETWEEN 81 AND 85 THEN '1-10' "
	sql = sql & "     WHEN TRY_CONVERT(int, LTRIM(RTRIM(LoanType))) BETWEEN 11 AND 29 THEN '11-29' "
	sql = sql & "     WHEN TRY_CONVERT(int, LTRIM(RTRIM(LoanType))) BETWEEN 30 AND 80 OR TRY_CONVERT(int, LTRIM(RTRIM(LoanType))) BETWEEN 86 AND 100 THEN '30-120' "
	sql = sql & "     ELSE ISNULL(LTRIM(RTRIM(LoanType)), '') "
	sql = sql & "   END AS LoanTypeGroup, "
	sql = sql & "   ISNULL(LTRIM(RTRIM([Description])), '') AS LoanDescription, "
	sql = sql & "   TRY_CONVERT(decimal(18,2), REPLACE(REPLACE(REPLACE(LTRIM(RTRIM(BalanceRaw)), ',', ''), '$', ''), ' ', '')) AS BalanceAmt, "
	sql = sql & "   ISNULL(LTRIM(RTRIM(LastPaid)), '') AS DLP, "
	sql = sql & "   TRY_CONVERT(int, LTRIM(RTRIM(DaysDQ))) AS DaysDQNum, "
	sql = sql & "   CASE "
	sql = sql & "     WHEN TRY_CONVERT(int, LTRIM(RTRIM(DaysDQ))) > ? THEN ? "
	sql = sql & "     WHEN TRY_CONVERT(int, LTRIM(RTRIM(DaysDQ))) > ? AND TRY_CONVERT(int, LTRIM(RTRIM(DaysDQ))) <= ? THEN ? "
	sql = sql & "     WHEN TRY_CONVERT(int, LTRIM(RTRIM(DaysDQ))) >= ? AND TRY_CONVERT(int, LTRIM(RTRIM(DaysDQ))) <= ? THEN ? "
	sql = sql & "     ELSE ? "
	sql = sql & "   END AS DQPct "
	sql = sql & " FROM " & tableName
	sql = sql & " WHERE PeriodDate BETWEEN ? AND ? "
	sql = sql & "), finalrows AS ("
	sql = sql & " SELECT "
	sql = sql & "   LoanTypeGroup, "
	sql = sql & "   1 AS RowSort, "
	sql = sql & "   FullName, "
	sql = sql & "   AccountNo, "
	sql = sql & "   SFX, "
	sql = sql & "   LoanDescription, "
	sql = sql & "   BalanceAmt, "
	sql = sql & "   DQPct, "
	sql = sql & "   DLP, "
	sql = sql & "   DaysDQNum "
	sql = sql & " FROM base "
	sql = sql & " UNION ALL "
	sql = sql & " SELECT "
	sql = sql & "   LoanTypeGroup, "
	sql = sql & "   2 AS RowSort, "
	sql = sql & "   '' AS FullName, "
	sql = sql & "   '' AS AccountNo, "
	sql = sql & "   '' AS SFX, "
	sql = sql & "   LoanTypeGroup + ' TOTAL' AS LoanDescription, "
	sql = sql & "   SUM(ISNULL(BalanceAmt,0)) AS BalanceAmt, "
	sql = sql & "   NULL AS DQPct, "
	sql = sql & "   '' AS DLP, "
	sql = sql & "   NULL AS DaysDQNum "
	sql = sql & " FROM base "
	sql = sql & " GROUP BY LoanTypeGroup "
	sql = sql & ") "
	sql = sql & "SELECT "
	sql = sql & "  LoanTypeGroup, FullName, AccountNo, SFX, LoanDescription, BalanceAmt, DQPct, DLP, DaysDQNum, RowSort "
	sql = sql & "FROM finalrows "
	sql = sql & "ORDER BY LoanTypeGroup, RowSort, DaysDQNum DESC, FullName, AccountNo"
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = sql
        .CommandType = 1

        AddParameterInt cmd, gDaysDQHigh
        AddParameterDouble cmd, gPctHigh

        AddParameterInt cmd, gDaysDQMidHigh
        AddParameterInt cmd, gDaysDQHigh
        AddParameterDouble cmd, gPctMidHigh

        AddParameterInt cmd, gDaysDQMidLow
        AddParameterInt cmd, gDaysDQMidHigh
        AddParameterDouble cmd, gPctMidLow

        AddParameterDouble cmd, gPctLow

        .Parameters.Append .CreateParameter("", 133, 1, , startDateObj)
        .Parameters.Append .CreateParameter("", 133, 1, , endDateObj)
    End With

    Err.Clear
    Set rs = cmd.Execute
    If Err.Number <> 0 Then
        LogLine "ERROR executing query for " & sectionTitle & ": " & Err.Description & " (Err=" & Err.Number & ")"
        Set rs = Nothing
        Set cmd = Nothing
        Exit Sub
    End If

    fout.WriteLine sectionTitle
    fout.WriteLine ""

    fout.WriteLine _
        Pad("NAME", 28) & _
        Pad("ACCT#", 14) & _
        Pad("SFX", 8) & _
        Pad("LOAN DESCRIPTION", 40) & _
        Pad("BALANCE", 14) & _
        Pad("FACTOR", 10) & _
        Pad("PROJECTED DEFICIENCY", 22) & _
        Pad("DQ %", 10) & _
        Pad("PROJECTED LOSS", 18) & _
        Pad("DLP", 12) & _
        Pad("DAYS DQ", 10)

    If rs.EOF Then
        fout.WriteLine "No records found for this section."
        fout.WriteLine ""
        Set rs = Nothing
        Set cmd = Nothing
        Exit Sub
    End If

    Do While Not rs.EOF
        nm       = SafeString(rs("FullName"))
        acct     = SafeString(rs("AccountNo"))
        sfx      = SafeString(rs("SFX"))
        loanDesc = SafeString(rs("LoanDescription"))
        dlp      = SafeString(rs("DLP"))

        rowSort = 0
        If Not IsNull(rs("RowSort")) Then rowSort = CLng(rs("RowSort"))

        If IsNull(rs("BalanceAmt")) Or Len(Trim(CStr(rs("BalanceAmt")))) = 0 Then
            bal = 0
        Else
            bal = CDbl(rs("BalanceAmt"))
        End If

        If IsNull(rs("DaysDQNum")) Or Len(Trim(CStr(rs("DaysDQNum")))) = 0 Then
            daysDQ = ""
        Else
            daysDQ = CStr(rs("DaysDQNum"))
        End If

        If rowSort = 2 Then
            projDef = bal
            projLoss = 0
            factorText = ""
            dqPctText = ""
        Else
            If IsNull(rs("DQPct")) Or Len(Trim(CStr(rs("DQPct")))) = 0 Then
                dqPct = 0
            Else
                dqPct = CDbl(rs("DQPct"))
            End If

            projDef = bal * gFactor
            projLoss = projDef * dqPct
            factorText = FormatPercentValue(gFactor)
            dqPctText = FormatPercentValue(dqPct)
        End If

        fout.WriteLine _
            Pad(nm, 28) & _
            Pad(acct, 14) & _
            Pad(sfx, 8) & _
            Pad(loanDesc, 40) & _
            Pad(FormatNumber(bal, 2), 14) & _
            Pad(factorText, 10) & _
            Pad(FormatNumber(projDef, 2), 22) & _
            Pad(dqPctText, 10) & _
            Pad(FormatNumber(projLoss, 2), 18) & _
            Pad(dlp, 12) & _
            Pad(daysDQ, 10)

        rs.MoveNext
    Loop

    Set rs = Nothing
    Set cmd = Nothing
    On Error GoTo 0
End Sub

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

    gConnStr        = GetIniValue(cfg, "Database.ConnStr", "")
    gFactor         = CDbl(GetIniValue(cfg, "Report.Factor", "0"))

    gDaysDQHigh     = CLng(GetIniValue(cfg, "DaysDQ.High", "191"))
    gDaysDQMidHigh  = CLng(GetIniValue(cfg, "DaysDQ.MidHigh", "140"))
    gDaysDQMidLow   = CLng(GetIniValue(cfg, "DaysDQ.MidLow", "98"))

    gPctHigh        = CDbl(GetIniValue(cfg, "DQPercent.High", "1.00"))
    gPctMidHigh     = CDbl(GetIniValue(cfg, "DQPercent.MidHigh", "0.75"))
    gPctMidLow      = CDbl(GetIniValue(cfg, "DQPercent.MidLow", "0.50"))
    gPctLow         = CDbl(GetIniValue(cfg, "DQPercent.Low", "0.25"))

    If IsNumeric(gRelOffset) Then
        gRelOffset = CInt(gRelOffset)
    Else
        gRelOffset = 0
    End If
End Sub

Function ValidateDaysDQConfig()
    ValidateDaysDQConfig = False

    If gDaysDQHigh <= gDaysDQMidHigh Then
        LogLine "ERROR: DaysDQ.High must be greater than DaysDQ.MidHigh"
        Exit Function
    End If

    If gDaysDQMidHigh <= gDaysDQMidLow Then
        LogLine "ERROR: DaysDQ.MidHigh must be greater than DaysDQ.MidLow"
        Exit Function
    End If

    ValidateDaysDQConfig = True
End Function

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

        Case Else
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
    If IsNull(val) Then
        s = ""
    Else
        s = CStr(val)
    End If
    s = Trim(s)
    Pad = Left(s & Space(width), width)
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

Function SafeString(ByVal v)
    If IsNull(v) Then
        SafeString = ""
    Else
        SafeString = Trim(CStr(v))
    End If
End Function

Function FormatPercentValue(ByVal v)
    FormatPercentValue = CStr(Round(CDbl(v) * 100, 2)) & "%"
End Function

Sub AddParameterDouble(cmdObj, dblVal)
    Const adDouble = 5
    Dim param
    On Error Resume Next
    Set param = cmdObj.CreateParameter("", adDouble, 1, , CDbl(dblVal))
    cmdObj.Parameters.Append param
    On Error GoTo 0
End Sub

Sub AddParameterInt(cmdObj, intVal)
    Const adInteger = 3
    Dim param
    On Error Resume Next
    Set param = cmdObj.CreateParameter("", adInteger, 1, , CLng(intVal))
    cmdObj.Parameters.Append param
    On Error GoTo 0
End Sub