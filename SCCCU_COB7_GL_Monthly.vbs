' =========================================================
' COB7_GL Period Detail Report (INI-driven)
' - Reads COB7_GL_Monthly.ini
' - Queries COB7_GL heap table (Run_Date is DATE type)
' - Returns ALL raw rows in chronological order: ORDER BY Run_Date, ID
' - Supports: DAILY (single day), MONTHLY (full month), MTD (month-to-date), CUSTOM
' - Fixed-width text output
' =========================================================
Option Explicit

' ---------------- CONFIG ----------------
Dim INI_PATH: INI_PATH = "S:\CarolinaTrust\byREQUEST\Templates\COB7_GL_Monthly.ini"

' ---------------- GLOBALS ----------------
Dim gMode, gRelOffset, gYear, gMonth, gQuarter
Dim gStartDateStr, gEndDateStr
Dim gConnStr, gFullTableName, gTable2Part, gObjIdName
Dim conn, rs

' =========================================================
' StartDoc - Main entry point
' =========================================================
Sub StartDoc
    On Error Resume Next
    Dim startDate, endDate, periodDisplay, errMsg
    
    If Not FileExists(INI_PATH) Then
        errMsg = "ERROR: INI file not found: " & INI_PATH
        LogLine errMsg
        fout.WriteLine errMsg
        Exit Sub
    End If
    
    ' Load INI
    LoadParameters INI_PATH
    
    If Trim(gConnStr) = "" Then
        errMsg = "ERROR: ConnStr not defined in [Database] section."
        LogLine errMsg
        fout.WriteLine errMsg
        Exit Sub
    End If
    
    BuildTableNames gFullTableName, gTable2Part, gObjIdName
    If Len(gTable2Part) = 0 Then
        errMsg = "ERROR: Could not parse TableName"
        LogLine errMsg
        fout.WriteLine errMsg
        Exit Sub
    End If
    
    ' Compute date range
    ComputeDateRange startDate, endDate
    
    If startDate > endDate Or IsEmpty(startDate) Or IsEmpty(endDate) Then
        errMsg = "ERROR: Invalid date range"
        LogLine errMsg
        fout.WriteLine errMsg
        Exit Sub
    End If
    
    periodDisplay = FormatPeriodDisplay(startDate, endDate, gMode)
    Spoolfile.User1 = "COB7_GL " & UCase(Trim(gMode)) & " - " & periodDisplay
    
    ' Connect to DB
    Set conn = CreateObject("ADODB.Connection")
    conn.Open gConnStr
    If Err.Number <> 0 Then
        LogLine "ERROR opening DB: " & Err.Description
        Exit Sub
    End If
    
    ' Main query
    Dim sql, cmd
    sql = _
        "SELECT " & _
        "   [Report], [Run_Date], [Run_Time], [GL_Number], [Branch], " & _
        "   [Debit], [Credit], [Sequence], [Voucher], [Teller], " & _
        "   [Tran_Nbr_Type1], [Mbr_Nbr], [BR], [Suffix], [Product_Type_Code], " & _
        "   [Tran_Nbr_Type2], [MI_Code], [Addl_Detail] " & _
        "FROM " & gTable2Part & " " & _
        "WHERE Run_Date BETWEEN ? AND ? " & _
        "ORDER BY Run_Date, ID"   ' <--- Stable chronological order
    
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandType = 1
        .CommandText = sql
        .Parameters.Append .CreateParameter("Start", 7, 1, , startDate)    ' adDBTimeStamp = 7 for DATE
        .Parameters.Append .CreateParameter("End",   7, 1, , endDate)
    End With
    
    Set rs = cmd.Execute
    
    If rs.EOF Then
        fout.WriteLine "No COB7_GL records found for period: " & periodDisplay
    Else
        ' Header
        fout.WriteLine _
            Pad("Report", 8) & _
            Pad("Run_Date", 10) & _
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
        
        ' Data rows
        Dim rowCount : rowCount = 0
        Do While Not rs.EOF
            rowCount = rowCount + 1
            fout.WriteLine _
                Pad(rs("Report"), 8) & _
                Pad(Format(rs("Run_Date"), "mm/dd/yyyy"), 10) & _   ' Consistent display
                Pad(rs("Run_Time"), 9) & _
                Pad(rs("GL_Number"), 11) & _
                Pad(rs("Branch"), 7) & _
                Pad(rs("Debit"), 16) & _
                Pad(rs("Credit"), 19) & _
                Pad(rs("Sequence"), 9) & _
                Pad(rs("Voucher"), 9) & _
                Pad(rs("Teller"), 8) & _
                Pad(rs("Tran_Nbr_Type1"), 12) & _
                Pad(rs("Mbr_Nbr"), 10) & _
                Pad(rs("BR"), 4) & _
                Pad(rs("Suffix"), 4) & _
                Pad(rs("Product_Type_Code"), 5) & _
                Pad(rs("Tran_Nbr_Type2"), 10) & _
                Pad(rs("MI_Code"), 8) & _
                Pad(rs("Addl_Detail"), 50)
            rs.MoveNext
        Loop
        LogLine "Rows returned: " & rowCount
    End If
    
    ' Cleanup
    If Not rs Is Nothing Then rs.Close
    If Not conn Is Nothing Then conn.Close
    Set rs = Nothing
    Set cmd = Nothing
    Set conn = Nothing
End Sub

' =========================================================
' Helpers (mostly unchanged, with minor improvements)
' =========================================================

Sub LoadParameters(iniPath)
    Dim cfg : Set cfg = ReadIniSafe(iniPath)
    If cfg Is Nothing Then Exit Sub
    
    gMode         = UCase(Trim(GetIniValue(cfg, "Report.Mode", "MONTHLY")))
    gRelOffset    = GetIniValue(cfg, "Report.RelativeOffset", "0")
    gYear         = GetIniValue(cfg, "Report.Year", "")
    gMonth        = GetIniValue(cfg, "Report.Month", "")
    gQuarter      = GetIniValue(cfg, "Report.Quarter", "")
    gStartDateStr = GetIniValue(cfg, "Report.StartDate", "")
    gEndDateStr   = GetIniValue(cfg, "Report.EndDate", "")
    
    If IsNumeric(gRelOffset) Then gRelOffset = CInt(gRelOffset) Else gRelOffset = 0
    
    gConnStr      = GetIniValue(cfg, "Database.ConnStr", "")
    gFullTableName = GetIniValue(cfg, "Database.TableName", "[dbo].[COB7_GL]")
End Sub

Sub ComputeDateRange(ByRef dStart, ByRef dEnd)
    Dim today : today = Date()
    Dim y, m
    
    Select Case gMode
        Case "DAILY"
            dStart = DateAdd("d", gRelOffset, today)
            dEnd   = dStart
            
        Case "MTD"   ' <--- New: current month-to-date
            dStart = DateSerial(Year(today), Month(today), 1)
            dEnd   = today
            
        Case "MONTHLY"
            If gYear <> "" And IsNumeric(gYear) Then y = CInt(gYear) Else y = Year(today)
            If gMonth <> "" And IsNumeric(gMonth) Then m = CInt(gMonth) Else m = Month(today)
            If gRelOffset <> 0 Then
                Dim tmp : tmp = DateAdd("m", gRelOffset, DateSerial(y, m, 1))
                y = Year(tmp)
                m = Month(tmp)
            End If
            dStart = DateSerial(y, m, 1)
            dEnd   = DateAdd("d", -1, DateAdd("m", 1, dStart))
            If dEnd > today Then dEnd = today   ' don't show future
            
        Case "CUSTOM"
            If gStartDateStr <> "" Then dStart = CDate(gStartDateStr) Else dStart = today
            If gEndDateStr   <> "" Then dEnd   = CDate(gEndDateStr)   Else dEnd   = today
            
        Case Else   ' fallback to monthly
            ComputeDateRange dStart, dEnd   ' recursive call to MONTHLY
    End Select
End Sub

' ... (keep your existing helper functions: FormatPeriodDisplay, Pad, FileExists, 
'     ReadIniSafe, GetIniValue, BuildTableNames, RemoveBrackets)

Function FormatPeriodDisplay(ByVal startDate, ByVal endDate, ByVal mode)
    Select Case UCase(Trim(mode))
        Case "DAILY"
            FormatPeriodDisplay = Format(startDate, "mmmm d, yyyy")
        Case "MTD"
            FormatPeriodDisplay = "Month-to-Date " & Format(startDate, "mmmm yyyy")
        Case "MONTHLY"
            FormatPeriodDisplay = Format(startDate, "mmmm yyyy")
        Case "CUSTOM"
            FormatPeriodDisplay = Format(startDate, "mmm d, yyyy") & " to " & Format(endDate, "mmm d, yyyy")
        Case Else
            FormatPeriodDisplay = Format(startDate, "mmm d, yyyy") & " to " & Format(endDate, "mmm d, yyyy")
    End Select
End Function

' ... (other unchanged helpers)
Function Pad(val, width)
    Dim s
    If IsNull(val) Then s = "" Else s = CStr(val)
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
        tableName = Trim(parts(1))
    Else
        schemaName = Trim(parts(UBound(parts)-1))
        tableName = Trim(parts(UBound(parts)))
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