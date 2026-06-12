Option Explicit

' =========================================================
'  ACH Period Report (Monthly / Weekly / Quarterly / Custom / YTD / Daily)
'  - Reads ACH_Monthly.ini
'  - Computes date range from Mode / RelativeOffset
'  - Queries ACH_EPSA_Staging by Post_Date (DATE)
'  - Writes fixed-width, column-aligned output
' =========================================================

' ---------------- GLOBALS ----------------
Dim gMode, gRelOffset, gYear, gMonth, gQuarter
Dim gStartDateStr, gEndDateStr
Dim gConnStr, gFullTableName, gTable2Part, gObjIdName

Dim conn, rs

Dim INI_PATH: INI_PATH = "S:\CarolinaTrust\byREQUEST\Templates\ACH_Monthly.ini"

' ---------------------------------------------------------
'  Simple log helper
' ---------------------------------------------------------
Sub LogLine(msg)
    On Error Resume Next
    fout.WriteLine msg
    On Error GoTo 0
End Sub

' ---------------------------------------------------------
'  StartDoc: run report
' ---------------------------------------------------------
Sub StartDoc
    On Error Resume Next

    Dim startKey, endKey
    Dim errMsg
    Dim sql, cmd
    Dim rowCount, totalAmount, amount
    Dim periodDisplay, startDateObj, endDateObj

    ' INI exists?
    If Not FileExists(INI_PATH) Then
        LogLine "ERROR: INI file not found: " & INI_PATH
        Exit Sub
    End If

    ' Load parameters
    LoadParameters INI_PATH

    If Trim(gConnStr) = "" Then
        LogLine "ERROR: ConnStr not defined in [Database] section."
        Exit Sub
    End If

    ' Parse table name
    BuildTableNames gFullTableName, gTable2Part, gObjIdName
    If Len(gTable2Part) = 0 Or Len(gObjIdName) = 0 Then
        LogLine "ERROR: Could not parse TableName: " & gFullTableName
        Exit Sub
    End If

    ' Compute date range keys (YYYYMMDD)
    ComputeDateRange startKey, endKey
    If startKey = "" Or endKey = "" Then
        LogLine "ERROR: Could not compute date range from parameters."
        Exit Sub
    End If

    ' Format the period for Spoolfile.User1
    startDateObj = FormatKeyToDate(startKey)
    endDateObj   = FormatKeyToDate(endKey)

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

    Spoolfile.User1 = gMode & " - " & periodDisplay

    ' Open DB connection
    Set conn = CreateObject("ADODB.Connection")
    Err.Clear
    conn.Open gConnStr
    If Err.Number <> 0 Then
        errMsg = "ERROR opening DB connection: " & Err.Description & " (Err=" & Err.Number & ")"
        LogLine errMsg
        Exit Sub
    End If

    ' =====================================================
    ' MAIN QUERY: Post_Date is DATE (parameterized)
    ' =====================================================
    sql = "SELECT " & _
          "[Report], [Post_Date], [Post_Time], [Settlement_Date], [Company_Name], " & _
          "[ACH_Batch_ID], [PCTL], [LCTL], [Status], [Tran_Type], [Tran_Code], " & _
          "[Effective_Date], [Amount], [Member_Nbr], [Acct], [Mbr_Name], [Fee], " & _
          "[ACH_Ind_Name], [DD], [Trace_Nbr], [Description], [Type], " & _
          "[Post_Seq], [File_Post_Seq], [EPSA_File], [Import_Timestamp] " & _
          "FROM " & gTable2Part & " " & _
          "WHERE [Post_Date] BETWEEN ? AND ? " & _
          "ORDER BY [id]"

    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = sql
        .CommandType = 1  ' adCmdText

        ' adDBDate = 133
        .Parameters.Append .CreateParameter("", 133, 1, , CDate(FormatKeyToDate(startKey)))
        .Parameters.Append .CreateParameter("", 133, 1, , CDate(FormatKeyToDate(endKey)))
    End With

    Err.Clear
    Set rs = cmd.Execute
    If Err.Number <> 0 Then
        errMsg = "ERROR executing query: " & Err.Description & " (Err=" & Err.Number & ")"
        LogLine errMsg
        CleanupObjects conn, rs, cmd
        Exit Sub
    End If

    If rs.EOF Then
        LogLine "No records found for the specified date range."
        fout.WriteLine "No records found for the specified date range."
        CleanupObjects conn, rs, cmd
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
        Pad("File_Post_Seq", 15) & _
        Pad("EPSA_File", 14)

    ' Process rows
    rowCount = 0
    totalAmount = 0

    Do While Not rs.EOF
        rowCount = rowCount + 1

        amount = 0
        If Not IsNull(rs("Amount")) Then
            If IsNumeric(rs("Amount")) Then
                amount = CDbl(rs("Amount"))
            End If
        End If
        totalAmount = totalAmount + amount

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
            Pad(rs("File_Post_Seq"), 15) & _
            Pad(rs("EPSA_File"), 14)

        rs.MoveNext
    Loop

    ' Optional summary
    'fout.WriteLine ""
    'fout.WriteLine "Total Records: " & rowCount & " | Total Amount: " & FormatNumber(totalAmount, 2)

    CleanupObjects conn, rs, cmd
    On Error GoTo 0
End Sub

' ---------------------------------------------------------
'  ProcessLine: not needed
' ---------------------------------------------------------
Sub ProcessLine
End Sub

' ---------------------------------------------------------
'  CloseDoc: nothing (we did it in StartDoc)
' ---------------------------------------------------------
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

' --- Load parameters from INI file ---
Sub LoadParameters(iniPath)
    Dim cfg

    Set cfg = ReadIniSafe(iniPath)
    If cfg Is Nothing Then
        LogLine "ERROR: Could not read INI file: " & iniPath
        Exit Sub
    End If

    ' [Report]
    gMode        = GetIniValue(cfg, "Report.Mode", "MONTHLY")
    gRelOffset   = GetIniValue(cfg, "Report.RelativeOffset", "0")
    gYear        = GetIniValue(cfg, "Report.Year", "")
    gMonth       = GetIniValue(cfg, "Report.Month", "")
    gQuarter     = GetIniValue(cfg, "Report.Quarter", "")
    gStartDateStr= GetIniValue(cfg, "Report.StartDate", "")
    gEndDateStr  = GetIniValue(cfg, "Report.EndDate", "")

    ' [Database]
    gConnStr      = GetIniValue(cfg, "Database.ConnStr", "")
    gFullTableName= GetIniValue(cfg, "Database.TableName", "[dbo].[ACH_EPSA_Staging]")

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

' --- Compute date range keys (startKey, endKey) based on Mode ---
Sub ComputeDateRange(ByRef startKey, ByRef endKey)
    Dim mode, today, dStart, dEnd
    Dim y, m, q, tmp, dow, weekStart

    mode  = UCase(Trim(gMode))
    today = Date()

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
            dStart = DateAdd("d", gRelOffset, today)
            dEnd   = dStart

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
                q = DatePart("q", today)
            End If

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
            dEnd   = DateAdd("d", -1, DateAdd("m", 3, dStart))

        Case "YTD"
            If gYear <> "" And IsNumeric(gYear) Then
                y = CInt(gYear)
            Else
                y = Year(today)
            End If
            dStart = DateSerial(y, 1, 1)
            dEnd   = today

        Case Else ' MONTHLY
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

            If gRelOffset <> 0 Then
                tmp = DateAdd("m", gRelOffset, DateSerial(y, m, 1))
                y = Year(tmp)
                m = Month(tmp)
            End If

            dStart = DateSerial(y, m, 1)
            dEnd   = DateAdd("d", -1, DateAdd("m", 1, dStart))
    End Select

    ' For non-custom (and non-daily), don't exceed today
    If mode <> "CUSTOM" And mode <> "DAILY" And dEnd > today Then
        dEnd = today
    End If

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
            ' skip
        ElseIf Left(line, 1) = "[" And Right(line, 1) = "]" Then
            sect = Mid(line, 2, Len(line) - 2)
        Else
            p = InStr(line, "=")
            If p > 0 Then
                d(sect & "." & Trim(Left(line, p - 1))) = Trim(Mid(line, p + 1))
            End If
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
