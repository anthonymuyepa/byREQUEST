Option Explicit

' =========================================================
'  ACH Period Report (Monthly / Weekly / Quarterly / Custom / YTD)
'  - Reads parameters.ini
'  - Computes date range from Mode / RelativeOffset
'  - Queries ACHCombine using DateCol (YYYYMMDD)
'  - Writes fixed-width, column-aligned output
'  - Now with heavy MsgBox + log debugging
' =========================================================

' Globals
Dim gMode, gRelOffset, gYear, gMonth, gQuarter
Dim gStartDateStr, gEndDateStr
Dim gConnStr

' ADO objects
Dim conn, rs

Const DEBUG_ON = FALSE

' ---------------------------------------------------------
'  Simple debug helper
' ---------------------------------------------------------
Sub DebugMsg(msg)
    If DEBUG_ON Then
        MsgBox msg
    End If
    On Error Resume Next
    'fout.WriteLine "DEBUG: " & msg
    On Error GoTo 0
End Sub

' ---------------------------------------------------------
'  StartDoc: run entire report here
' ---------------------------------------------------------
Sub StartDoc
    On Error Resume Next

    Dim paramsPath, startKey, endKey
    Dim errMsg

    DebugMsg "StartDoc: entered"

    ' parameters.ini location
    paramsPath = "S:\JHU\ACHParameters.ini"
    DebugMsg "StartDoc: parameters.ini path = " & paramsPath

    If Not FileExists(paramsPath) Then
        errMsg = "ERROR: parameters.ini not found: " & paramsPath
        DebugMsg errMsg
        fout.WriteLine errMsg
        Exit Sub
    End If

    ' Load parameters from ini
    DebugMsg "StartDoc: calling LoadParameters"
    LoadParameters paramsPath
    DebugMsg "StartDoc: after LoadParameters - Mode=" & gMode & _
             ", RelOffset=" & gRelOffset & ", Year=" & gYear & _
             ", Month=" & gMonth & ", Quarter=" & gQuarter & _
             ", StartDate=" & gStartDateStr & ", EndDate=" & gEndDateStr & _
             ", ConnStr=" & gConnStr

    If Trim(gConnStr) = "" Then
        errMsg = "ERROR: ConnStr not defined in [Database] section."
        DebugMsg errMsg
        fout.WriteLine errMsg
        Exit Sub
    End If

    ' Compute date range (YYYYMMDD) for DateCol
    DebugMsg "StartDoc: calling ComputeDateRange"
    ComputeDateRange startKey, endKey
    DebugMsg "StartDoc: after ComputeDateRange - startKey=" & startKey & ", endKey=" & endKey

    If startKey = "" Or endKey = "" Then
        errMsg = "ERROR: Could not compute date range from parameters."
        DebugMsg errMsg
        fout.WriteLine errMsg
        Exit Sub
    End If

    ' Header info
    fout.WriteLine "ACH Report " & " Mode= " & gMode & "  RelativeOffset= " & gRelOffset & _
                   "  From " & startKey & " To " & endKey
    Spoolfile.user1 =  "ACH Report " & " Mode= " & gMode & "  RelativeOffset= " & gRelOffset & _
                   "  From " & startKey & " To " & endKey

    ' Open connection
    DebugMsg "StartDoc: opening DB connection"
    Set conn = CreateObject("ADODB.Connection")
    Err.Clear
    conn.Open gConnStr

    If Err.Number <> 0 Then
        errMsg = "ERROR opening DB connection: " & Err.Description & " (Err=" & Err.Number & ")"
        DebugMsg errMsg
        fout.WriteLine errMsg
        Exit Sub
    End If
    DebugMsg "StartDoc: DB connection opened successfully"

    ' =====================================================
    ' 1) MAIN QUERY: exclude File Acc Number LIKE '%TOTAL%'
    ' =====================================================
    Dim sql, cmd, rowCount

    sql = "SELECT [DateCol], [RunDate], [Member Number], [Tran Code], " & _
          "       [File Acc Number], [Amount], [Trace Number], [Batch], [Source Name], " & _
          "       [ACTUALRC], [FILERC], [DIFFERENCERC], " & _
          "       [ACTUALTD], [FILETD], [DIFFERENCETD], " & _
          "       [ACTUALTC], [FILETC], [DIFFERENCETC] " & _
          "FROM   [byREQUEST].[dbo].[ACHCombine_original] " & _
          "WHERE  LTRIM(RTRIM([DateCol])) >= ? " & _
          "  AND  LTRIM(RTRIM([DateCol])) <= ? " & _
          "  AND  [File Acc Number] NOT LIKE '%TOTAL%' " & _
          "ORDER BY [DateCol], [Member Number], [Trace Number]"

    DebugMsg "StartDoc: MAIN SQL (excluding TOTAL) = " & sql

    Set cmd = CreateObject("ADODB.Command")

    With cmd
        .ActiveConnection = conn
        .CommandText = sql
        .CommandType = 1   ' adCmdText

        Err.Clear
        .Parameters.Append .CreateParameter("", 202, 1, 8, startKey)
        If Err.Number <> 0 Then
            errMsg = "ERROR adding first parameter (main): " & Err.Description & " (Err=" & Err.Number & ")"
            DebugMsg errMsg
            fout.WriteLine errMsg
            Exit Sub
        End If

        Err.Clear
        .Parameters.Append .CreateParameter("", 202, 1, 8, endKey)
        If Err.Number <> 0 Then
            errMsg = "ERROR adding second parameter (main): " & Err.Description & " (Err=" & Err.Number & ")"
            DebugMsg errMsg
            fout.WriteLine errMsg
            Exit Sub
        End If
    End With

    DebugMsg "StartDoc: parameters added for MAIN query. Executing..."

    Err.Clear
    Set rs = cmd.Execute

    If Err.Number <> 0 Then
        errMsg = "ERROR executing MAIN query: " & Err.Description & " (Err=" & Err.Number & ")"
        DebugMsg errMsg
        fout.WriteLine errMsg
        Exit Sub
    End If

    If rs.EOF Then
        DebugMsg "StartDoc: MAIN query returned no rows."
    Else
        DebugMsg "StartDoc: MAIN query returned at least one row."
    End If

    ' Column headers (aligned) – now with RC/TD/TC fields
    fout.WriteLine _
        Pad("DateCol", 10) & _
        Pad("RunDate", 10) & _
        Pad("Member", 15) & _
        Pad("Tran", 6) & _
        Pad("FileAcc", 15) & _
        Pad("Amount", 12) & _
        Pad("Trace", 10) & _
        Pad("Batch", 14) & _
        Pad("Source Name", 20) & _
        Pad("ActRC", 12) & _
        Pad("FileRC", 12) & _
        Pad("DiffRC", 12) & _
        Pad("ActTD", 12) & _
        Pad("FileTD", 12) & _
        Pad("DiffTD", 12) & _
        Pad("ActTC", 12) & _
        Pad("FileTC", 12) & _
        Pad("DiffTC", 12)



    ' MAIN rows
    rowCount = 0

    Do While Not rs.EOF
        rowCount = rowCount + 1

        On Error Resume Next
        fout.WriteLine _
            Pad(rs("DateCol"),         10) & _
            Pad(rs("RunDate"),         10) & _
            Pad(rs("Member Number"),   15) & _
            Pad(rs("Tran Code"),       6)  & _
            Pad(rs("File Acc Number"), 15) & _
            Pad(rs("Amount"),          12) & _
            Pad(rs("Trace Number"),    10) & _
            Pad(rs("Batch"),           14) & _
            Pad(rs("Source Name"),     20) & _
            Pad(rs("ACTUALRC"),        12)  & _
            Pad(rs("FILERC"),          12)  & _
            Pad(rs("DIFFERENCERC"),    12)  & _
            Pad(rs("ACTUALTD"),        12)  & _
            Pad(rs("FILETD"),          12)  & _
            Pad(rs("DIFFERENCETD"),    12)  & _
            Pad(rs("ACTUALTC"),        12)  & _
            Pad(rs("FILETC"),          12)  & _
            Pad(rs("DIFFERENCETC"),    12)

        If Err.Number <> 0 Then
            errMsg = "ERROR writing MAIN row " & rowCount & ": " & Err.Description & " (Err=" & Err.Number & ")"
            DebugMsg errMsg
            fout.WriteLine errMsg
            Err.Clear
        End If

        rs.MoveNext
    Loop

    DebugMsg "StartDoc: MAIN query finished. Total rows written = " & rowCount

    rs.Close
    Set rs = Nothing
    Set cmd = Nothing

    ' =====================================================
    ' 2) SECOND QUERY: ONLY rows with File Acc Number LIKE '%TOTAL%'
    ' =====================================================
    fout.WriteLine ""
    fout.WriteLine String(200, "=")
    fout.WriteLine "Rows with [File Acc Number] LIKE '%TOTAL%' (excluded from main detail):"


    sql = "SELECT [DateCol], [RunDate], [Member Number], [Tran Code], " & _
          "       [File Acc Number], [Amount], [Trace Number], [Batch], [Source Name], " & _
          "       [ACTUALRC], [FILERC], [DIFFERENCERC], " & _
          "       [ACTUALTD], [FILETD], [DIFFERENCETD], " & _
          "       [ACTUALTC], [FILETC], [DIFFERENCETC] " & _
          "FROM   [byREQUEST].[dbo].[ACHCombine_original] " & _
          "WHERE  LTRIM(RTRIM([DateCol])) >= ? " & _
          "  AND  LTRIM(RTRIM([DateCol])) <= ? " & _
          "  AND  [File Acc Number] LIKE '%TOTAL%' " & _
          "ORDER BY [DateCol], [Member Number], [Trace Number]"

    DebugMsg "StartDoc: TOTAL SQL = " & sql

    Set cmd = CreateObject("ADODB.Command")

    With cmd
        .ActiveConnection = conn
        .CommandText = sql
        .CommandType = 1   ' adCmdText

        Err.Clear
        .Parameters.Append .CreateParameter("", 202, 1, 8, startKey)
        If Err.Number <> 0 Then
            errMsg = "ERROR adding first parameter (TOTAL): " & Err.Description & " (Err=" & Err.Number & ")"
            DebugMsg errMsg
            fout.WriteLine errMsg
            Exit Sub
        End If

        Err.Clear
        .Parameters.Append .CreateParameter("", 202, 1, 8, endKey)
        If Err.Number <> 0 Then
            errMsg = "ERROR adding second parameter (TOTAL): " & Err.Description & " (Err=" & Err.Number & ")"
            DebugMsg errMsg
            fout.WriteLine errMsg
            Exit Sub
        End If
    End With

    DebugMsg "StartDoc: parameters added for TOTAL query. Executing..."

    Err.Clear
    Set rs = cmd.Execute

    If Err.Number <> 0 Then
        errMsg = "ERROR executing TOTAL query: " & Err.Description & " (Err=" & Err.Number & ")"
        DebugMsg errMsg
        fout.WriteLine errMsg
        Exit Sub
    End If

    rowCount = 0

    Do While Not rs.EOF
        rowCount = rowCount + 1

        On Error Resume Next
        fout.WriteLine _
            Pad(rs("DateCol"),         10) & _
            Pad(rs("RunDate"),         10) & _
            Pad(rs("Member Number"),   15) & _
            Pad(rs("Tran Code"),       6)  & _
            Pad(rs("File Acc Number"), 15) & _
            Pad(rs("Amount"),          12) & _
            Pad(rs("Trace Number"),    10) & _
            Pad(rs("Batch"),           14) & _
            Pad(rs("Source Name"),     20) & _
            Pad(rs("ACTUALRC"),        12)  & _
            Pad(rs("FILERC"),          12)  & _
            Pad(rs("DIFFERENCERC"),    12)  & _
            Pad(rs("ACTUALTD"),        12)  & _
            Pad(rs("FILETD"),          12)  & _
            Pad(rs("DIFFERENCETD"),    12)  & _
            Pad(rs("ACTUALTC"),        12)  & _
            Pad(rs("FILETC"),          12)  & _
            Pad(rs("DIFFERENCETC"),    12)

        If Err.Number <> 0 Then
            errMsg = "ERROR writing TOTAL row " & rowCount & ": " & Err.Description & " (Err=" & Err.Number & ")"
            DebugMsg errMsg
            fout.WriteLine errMsg
            Err.Clear
        End If

        rs.MoveNext
    Loop

    DebugMsg "StartDoc: TOTAL query finished. TOTAL rows written = " & rowCount

    ' Cleanup
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If

    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If

    DebugMsg "StartDoc: cleanup complete, exiting"

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

' --- Simple INI loader: fills global vars from [Report] and [Database] ---
Sub LoadParameters(path)
    Dim fso, ts, line, currentSection
    Dim p, key, value

    DebugMsg "LoadParameters: opening " & path

    ' defaults
    gMode         = "MONTHLY"
    gRelOffset    = 0
    gYear         = ""
    gMonth        = ""
    gQuarter      = ""
    gStartDateStr = ""
    gEndDateStr   = ""
    gConnStr      = ""

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts  = fso.OpenTextFile(path, 1)

    currentSection = ""

    Do Until ts.AtEndOfStream
        line = Trim(ts.ReadLine)

        If line = "" Then
            ' skip
        ElseIf Left(line, 1) = ";" Then
            ' comment
        ElseIf Left(line, 1) = "[" And Right(line, 1) = "]" Then
            currentSection = UCase(Mid(line, 2, Len(line) - 2))
        Else
            p = InStr(line, "=")
            If p > 0 Then
                key   = UCase(Trim(Left(line, p - 1)))
                value = Trim(Mid(line, p + 1))

                Select Case currentSection
                    Case "REPORT"
                        Select Case key
                            Case "MODE"
                                gMode = value

                            Case "RELATIVEOFFSET"
                                If value <> "" And IsNumeric(value) Then
                                    gRelOffset = CInt(value)
                                End If

                            Case "YEAR"
                                gYear = value

                            Case "MONTH"
                                gMonth = value

                            Case "QUARTER"
                                gQuarter = value

                            Case "STARTDATE"
                                gStartDateStr = value

                            Case "ENDDATE"
                                gEndDateStr = value
                        End Select

                    Case "DATABASE"
                        If key = "CONNSTR" Then gConnStr = value
                End Select
            End If
        End If
    Loop

    ts.Close

    DebugMsg "LoadParameters: done. Mode=" & gMode & ", RelOffset=" & gRelOffset
End Sub

' --- Compute date range (startKey, endKey) based on Mode & parameters ---
Sub ComputeDateRange(ByRef startKey, ByRef endKey)
    Dim mode, today, dStart, dEnd
    Dim y, m, q, tmp, dow, weekStart

    mode  = UCase(Trim(gMode))
    today = Date()

    DebugMsg "ComputeDateRange: mode=" & mode & ", today=" & CStr(today)

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
                q = ((Month(today) - 1) \ 3) + 1
            End If

            If gRelOffset <> 0 Then
                tmp = DateAdd("q", gRelOffset, DateSerial(y, (q - 1) * 3 + 1, 1))
                y   = Year(tmp)
                q   = ((Month(tmp) - 1) \ 3) + 1
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

            If gRelOffset <> 0 Then
                tmp = DateAdd("m", gRelOffset, DateSerial(y, m, 1))
                y   = Year(tmp)
                m   = Month(tmp)
            End If

            dStart = DateSerial(y, m, 1)
            dEnd   = DateAdd("d", -1, DateAdd("m", 1, dStart))
    End Select

    startKey = FormatDateKey(dStart)
    endKey   = FormatDateKey(dEnd)

    DebugMsg "ComputeDateRange: dStart=" & CStr(dStart) & ", dEnd=" & CStr(dEnd) & _
             ", startKey=" & startKey & ", endKey=" & endKey
End Sub

' --- Format VBScript Date as "YYYYMMDD" for DateCol comparison ---
Function FormatDateKey(d)
    FormatDateKey = CStr(Year(d)) & _
                    Right("0" & Month(d), 2) & _
                    Right("0" & Day(d),   2)
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
