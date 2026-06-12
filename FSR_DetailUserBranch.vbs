Option Explicit

Dim folderPath, connString, tablename, fso
Set fso = CreateObject("Scripting.FileSystemObject")

folderPath = "R:\byrequest\reports\Lobby Tracking\"
tablename = "FSR_DetailUserBranch"
connString = "Provider=SQLOLEDB;" & _
             "Data Source=GOLDMINE-SQL1;" & _
             "Initial Catalog=byREQUEST;" & _
             "Integrated Security=SSPI;"

Sub StartDoc()
    On Error Resume Next

    Dim months, m, currMonth, thisYear, prevYear
    Dim monthName, monthPath, entryYearForMonth, dtLastDay, entryDate
    Dim xl, conn, file

    currMonth = Month(Date)
    thisYear  = Year(Date)
    prevYear  = thisYear - 1

    months = Array("January","February","March","April","May","June", _
                   "July","August","September","October","November","December")

    ' Header (once)
    fout.WriteLine PadRight("DateCol", 12)  & PadRight("Month", 10) & PadRight("Branch", 20) & _
                   PadRight("EmpName", 24) & PadRight("AvgAssistTime", 18) & _
                   PadRight("DaysWorked", 12)

    ' DB connect (once)
    Dim createSQL
    Set conn = CreateObject("ADODB.Connection")
    conn.Open connString
    If conn.State <> 1 Then
        Call LogError("DB connection failed: " & Err.Description)
        Exit Sub
    End If

    createSQL = "IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '" & tablename & "') " & _
                "BEGIN CREATE TABLE " & tablename & " (" & _
                "DateCol NVARCHAR(8), Branch NVARCHAR(100), EmpName NVARCHAR(100), " & _
                "AvgAssistTime NVARCHAR(50), DaysWorked INT, Month NVARCHAR(10)) END"
    conn.Execute createSQL
    If Err.Number <> 0 Then Call LogError("Create table error: " & Err.Description)

    ' Excel (once)
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False

    ' ---- Traverse ALL available month folders ----
    For m = 5 To 9
        monthName = months(m - 1)
        monthPath = folderPath & monthName & "\"

        If fso.FolderExists(monthPath) Then
            ' Rolling year: months ahead of today belong to previous year
            If m <= currMonth Then
                entryYearForMonth = thisYear
            Else
                entryYearForMonth = prevYear
            End If

            ' Last day of this month for DateCol (YYYYMMDD)
            dtLastDay = DateSerial(entryYearForMonth, m + 1, 0)
            entryDate = Year(dtLastDay) & Right("0" & Month(dtLastDay), 2) & Right("0" & Day(dtLastDay), 2)

            ' Process matching files in this month
            For Each file In fso.GetFolder(monthPath).Files
                If LCase(fso.GetExtensionName(file.Name)) = "xlsx" _
                   And InStr(LCase(file.Name), "~") = 0 _
                   And InStr(LCase(file.Name), "detail user and branch overview report") > 0 Then

                    Call ProcessLobbyFile(file.Path, entryDate, xl, conn)
                End If
            Next
        Else
            Call LogInfo("Skipping missing folder: " & monthPath)
        End If
    Next

    ' Cleanup
    If Not conn Is Nothing Then conn.Close : Set conn = Nothing
    If Not xl Is Nothing Then xl.Quit : Set xl = Nothing
End Sub



Sub ProcessLobbyFile(excelPath, entryDate, xl, conn)
    Dim wb, ws, row, empName, avgTime, daysWorked
    Dim headerParts, dateParts, ColDate, Branch, upsertSQL
    Dim skipFile : skipFile = False

    Set wb = Nothing : Set ws = Nothing

    On Error Resume Next
    Set wb = xl.Workbooks.Open(excelPath)
    If Err.Number <> 0 Or wb Is Nothing Then
        Call LogError("Cannot open Excel: " & excelPath)
        skipFile = True
        Err.Clear
    End If
    On Error GoTo 0

    If skipFile Then Exit Sub

    Set ws = wb.Sheets(1)
    headerParts = Split(ws.Range("B6").Value, "|")
    Branch = Trim(Split(headerParts(3), ":")(1))
    dateParts = Split(headerParts(2), "to")
    Dim dt : dt = CDate(Trim(dateParts(1)))
	
	dim mnth
	mnth = Left(MonthName(Month(dt), True), 3)  ' "Mar"
	
    ColDate = Year(dt) & Right("0" & Month(dt), 2) & Right("0" & Day(dt), 2)

    Dim maxRows : maxRows = ws.Cells(ws.Rows.Count, 2).End(-4162).Row
    For row = 1 To maxRows
        If Trim(ws.Cells(row, 2).Value) = "Associate Stats" Then
            row = row + 1
            Exit For
        End If
    Next

    Do While row <= maxRows
        empName = ws.Cells(row, 2).Value
        If empName <> "" And InStr(1, ws.Cells(row, 7).Value, "Days Worked:", vbTextCompare) > 0 Then
            daysWorked = Trim(ws.Cells(row, 10).Value)
            avgTime = Trim(ws.Cells(row + 2, 7).Value)
            avgTime = ConvertToTimeFormat(avgTime)
            If Not IsNumeric(daysWorked) Then daysWorked = 0

            fout.WriteLine PadRight(ColDate, 12) & PadRight(mnth, 10) &  PadRight(Branch, 20) & _
                           PadRight(empName, 24) & PadRight(avgTime, 18) & _
                           PadRight(daysWorked, 12) 

            upsertSQL = "IF NOT EXISTS (SELECT 1 FROM " & tablename & " WHERE " & _
                        "DateCol = '" & SqlSafe(ColDate) & "' AND " & _
                        "Branch = '" & SqlSafe(Branch) & "' AND " & _
                        "EmpName = '" & SqlSafe(empName) & "') " & _
                        "BEGIN " & _
                        "INSERT INTO " & tablename & _
                        " (DateCol, Branch, EmpName, AvgAssistTime, DaysWorked, Month) VALUES (" & _
                        "'" & SqlSafe(ColDate) & "', '" & SqlSafe(Branch) & "', '" & SqlSafe(empName) & "', '" & SqlSafe(avgTime) & "', " & _
                        daysWorked & ", '" & SqlSafe(mnth) & "') " & _
                        "END"
            On Error Resume Next
            conn.Execute upsertSQL
            If Err.Number <> 0 Then Call LogError("Insert error: " & Err.Description): Err.Clear
            On Error GoTo 0

            row = row + 3
        Else
            row = row + 1
        End If
    Loop

    wb.Close False : Set wb = Nothing
End Sub

' === Helpers ===

Sub LogError(msg)
    If fout Is Nothing Then Set fout = fso.OpenTextFile("FSR_ErrorLog.txt", 8, True)
    fout.WriteLine Now & " | ERROR: " & msg
End Sub

Sub LogInfo(msg)
    If Not IsEmpty(fout) Then fout.WriteLine Now & " | INFO: " & msg
End Sub

Function SqlSafe(val)
    If IsNull(val) Or IsEmpty(val) Then
        SqlSafe = ""
    Else
        SqlSafe = Replace(CStr(val), "'", "''")
    End If
End Function

Function ConvertToMinutes(avgTimeText)
    Dim minutes, seconds, part, parts
    minutes = 0 : seconds = 0
    avgTimeText = LCase(Trim(avgTimeText))
    parts = Split(avgTimeText, ",")

    For Each part In parts
        part = Trim(part)
        If InStr(part, "minute") > 0 Then
            part = Replace(part, "minutes", ""): part = Replace(part, "minute", "")
            If IsNumeric(part) Then minutes = CLng(part)
        ElseIf InStr(part, "second") > 0 Then
            part = Replace(part, "seconds", ""): part = Replace(part, "second", "")
            If IsNumeric(part) Then seconds = CLng(part)
        End If
    Next

    ConvertToMinutes = Round(minutes + (seconds / 60), 2)
End Function

Function PadRight(val, width)
    PadRight = Left(val & Space(width), width)
End Function


Function ConvertToTimeFormat(rawTime)
    Dim minutes, seconds, matches, regex, hhmmss
    Dim match

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = "(\d+)\s*minutes?|(\d+)\s*seconds?"

    Set matches = regex.Execute(rawTime)
    minutes = 0
    seconds = 0

    For Each match In matches
        If InStr(LCase(match.Value), "minute") > 0 Then
            minutes = CLng(Trim(Split(match.Value, " ")(0)))
        ElseIf InStr(LCase(match.Value), "second") > 0 Then
            seconds = CLng(Trim(Split(match.Value, " ")(0)))
        End If
    Next

    hhmmss = "00:" & Right("0" & minutes, 2) & ":" & Right("0" & seconds, 2)
    ConvertToTimeFormat = hhmmss
End Function
