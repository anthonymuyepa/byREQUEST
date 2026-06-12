Option Explicit

' ---------------- GLOBALS ----------------
Dim gMode, gRelOffset, gYear, gMonth, gQuarter
Dim gStartDateStr, gEndDateStr
Dim gConnStr, gFullTableName, gTable2Part, gObjIdName

Dim gFactor
Dim gDaysDQHigh, gDaysDQMidHigh, gDaysDQMidLow
Dim gPctHigh, gPctMidHigh, gPctMidLow, gPctLow

Dim conn, rs
Dim INI_PATH: INI_PATH = "S:\CarolinaTrust\byREQUEST\Templates\DQ_ICR_Monthly.ini"

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
    Dim sql, cmd
    Dim startDateObj, endDateObj
    Dim sectionTitle

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

    sectionTitle = "SECURED DELINQUENCY"

    sql = ""
    sql = sql & "SELECT " & _
                "  ISNULL(LTRIM(RTRIM(FullName)), '') AS FullName, " & _
                "  ISNULL(LTRIM(RTRIM(AccountNo)), '') AS AccountNo, " & _
                "  ISNULL(LTRIM(RTRIM(IDValue)), '') AS SFX, " & _
                "  ISNULL(LTRIM(RTRIM(Description)), '') AS LoanDescription, " & _
                "  TRY_CONVERT(decimal(18,2), REPLACE(REPLACE(REPLACE(LTRIM(RTRIM(BalanceRaw)), ',', ''), '$', ''), ' ', '')) AS BalanceAmt, " & _
                "  ISNULL(LTRIM(RTRIM(LastPaid)), '') AS DLP, " & _
                "  TRY_CONVERT(int, LTRIM(RTRIM(DaysDQ))) AS DaysDQNum, " & _
                "  CASE " & _
                "    WHEN TRY_CONVERT(int, LTRIM(RTRIM(DaysDQ))) > ? THEN ? " & _
                "    WHEN TRY_CONVERT(int, LTRIM(RTRIM(DaysDQ))) > ? AND TRY_CONVERT(int, LTRIM(RTRIM(DaysDQ))) <= ? THEN ? " & _
                "    WHEN TRY_CONVERT(int, LTRIM(RTRIM(DaysDQ))) >= ? AND TRY_CONVERT(int, LTRIM(RTRIM(DaysDQ))) <= ? THEN ? " & _
                "    ELSE ? " & _
                "  END AS DQPct " & _
                "FROM " & gTable2Part & " " & _
                "WHERE PeriodDate BETWEEN ? AND ? " & _
                "ORDER BY TRY_CONVERT(int, LTRIM(RTRIM(DaysDQ))) DESC, FullName, AccountNo"

    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = sql
        .CommandType = 1   ' adCmdText

        AddParameterInt    cmd, gDaysDQHigh
        AddParameterDouble cmd, gPctHigh

        AddParameterInt    cmd, gDaysDQMidHigh
        AddParameterInt    cmd, gDaysDQHigh
        AddParameterDouble cmd, gPctMidHigh

        AddParameterInt    cmd, gDaysDQMidLow
        AddParameterInt    cmd, gDaysDQMidHigh
        AddParameterDouble cmd, gPctMidLow

        AddParameterDouble cmd, gPctLow

        .Parameters.Append .CreateParameter("", 133, 1, , startDateObj) ' adDBDate
        .Parameters.Append .CreateParameter("", 133, 1, , endDateObj)
    End With

    Err.Clear
    Set rs = cmd.Execute
    If Err.Number <> 0 Then
        LogLine "ERROR executing secured query: " & Err.Description & " (Err=" & Err.Number & ")"
        CleanupObjects conn, rs, cmd
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

    Do While Not rs.EOF
        Dim nm, acct, sfx, loanDesc, bal, dlp, daysDQ, dqPct
        Dim projDef, projLoss

        nm       = SafeString(rs("FullName"))
        acct     = SafeString(rs("AccountNo"))
        sfx      = SafeString(rs("SFX"))
        loanDesc = SafeString(rs("LoanDescription"))
        dlp      = SafeString(rs("DLP"))

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

        If IsNull(rs("DQPct")) Or Len(Trim(CStr(rs("DQPct")))) = 0 Then
            dqPct = 0
        Else
            dqPct = CDbl(rs("DQPct"))
        End If

        projDef  = bal * gFactor
        projLoss = projDef * dqPct

        fout.WriteLine _
            Pad(nm, 28) & _
            Pad(acct, 14) & _
            Pad(sfx, 8) & _
            Pad(loanDesc, 40) & _
            Pad(FormatNumber(bal, 2), 14) & _
            Pad(FormatPercentValue(gFactor), 10) & _
            Pad(FormatNumber(projDef, 2), 22) & _
            Pad(FormatPercentValue(dqPct), 10) & _
            Pad(FormatNumber(projLoss, 2), 18) & _
            Pad(dlp, 12) & _
            Pad(daysDQ, 10)

        rs.MoveNext
    Loop

    CleanupObjects conn, rs, cmd
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
    gFullTableName = GetIniValue(cfg, "Database.TableName", "[dbo].[DQ_Secured]")

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