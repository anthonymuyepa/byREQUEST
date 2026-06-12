Option Explicit

Dim excelPath, basePath, tablename, connString, conn
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")

basePath = "R:\byRequest\Reports\phonesystem\Monthly\"
tablename = "FSR_AcdCalls"
connString = "Provider=SQLOLEDB;" & _
             "Data Source=GOLDMINE-SQL1;" & _
             "Initial Catalog=byREQUEST;" & _
             "Integrated Security=SSPI;"



' === MASTER ENTRY POINT ===
Sub StartDoc()
    Dim fso, folderPath, monthFolder, subfolder, filePath
    Dim tday, dt, fileYear, fileMonthName
    Dim excelPath

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' === Use previous month ===
    tday = Date
    dt = DateAdd("m", -1, tday)
    fileYear = Year(dt)
    fileMonthName = MonthName(Month(dt), False) ' "May", "June", etc.

    ' === Build monthly path: R:\...\Monthly\2025\May\ ===

    folderPath = basePath & fileYear & "\" & fileMonthName & "\"

    If Not fso.FolderExists(folderPath) Then
        Call LogError("Folder not found: " & folderPath)
        Exit Sub
    End If

    ' === Loop through all daily subfolders: 01\, 02\, ..., 31\ ===
    For Each subfolder In fso.GetFolder(folderPath).SubFolders
        filePath = subfolder.Path & "\Branch ACD Monthly Report.xls"

        If fso.FileExists(filePath) Then
            excelPath = filePath
            'Call LogInfo("Processing file: " & excelPath)
            Call ProcessAllACDReports(excelPath)
        Else
            Call LogInfo("File not found: " & filePath)
        End If
    Next
End Sub



' === ORCHESTRATOR ===
Sub ProcessAllACDReports(filePath)
    excelPath = filePath
    Dim xl, wb, ws, reportPeriod, DateCol
    Dim connReady

    connReady = InitializeConnection()
    If Not connReady Then Exit Sub

    If Not EnsureTableExists() Then Exit Sub

    Set xl = CreateObject("Excel.Application")
    xl.Visible = False

    Set wb = xl.Workbooks.Open(excelPath)
    Set ws = wb.Sheets(1)

    If ws Is Nothing Then
        Call LogError("Worksheet not found in: " & excelPath)
        wb.Close False : xl.Quit
        Exit Sub
    End If

    reportPeriod = FormatReportPeriod(ws.Range("F4").Value)
    	
	Dim yearVal, monthNum, lastDay
	
	yearVal = Year(Now)
	monthNum = Month(DateValue("1 " & reportPeriod & " " & yearVal))
	lastDay = DateSerial(yearVal, monthNum + 1, 0)
	DateCol = Year(lastDay) & Right("0" & Month(lastDay), 2) & Right("0" & Day(lastDay), 2)

    ' Delete existing records
    On Error Resume Next
    conn.Execute "DELETE FROM " & tablename & " WHERE DateCol = '" & DateCol & "' AND month = '" & reportPeriod & "'"
    If Err.Number <> 0 Then Call LogError("Error deleting existing data: " & Err.Description): Err.Clear
    On Error GoTo 0

    ' Header line
    fout.WriteLine PadRight("DateCol", 10) & PadRight("Month", 10) & PadRight("User", 24) & PadRight("Answered", 10)

    ' Insert parsed data
    Call InsertACDData(ws, DateCol, reportPeriod)


	Set ws = Nothing
    wb.Close False : xl.Quit
    Set wb = Nothing : Set xl = Nothing
    conn.Close : Set conn = Nothing
End Sub






Function InitializeConnection()
    Set conn = CreateObject("ADODB.Connection")
    On Error Resume Next
    conn.Open connString
    If conn.State <> 1 Then
        Call LogError("Failed to open database connection." & vbCrLf & "Error: " & Err.Description)
        Set conn = Nothing
        InitializeConnection = False
    Else
        InitializeConnection = True
    End If
    On Error GoTo 0
End Function

Function EnsureTableExists()
    Dim rs, tableExists
    tableExists = False
    On Error Resume Next
    Set rs = conn.Execute("SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '" & tablename & "'")
    If Not rs Is Nothing Then
        If Not rs.EOF Then tableExists = True
        rs.Close
    End If
    On Error GoTo 0

    If tableExists Then
        EnsureTableExists = True
        Exit Function
    End If

    Dim createSQL
    createSQL = "CREATE TABLE " & tablename & " (" & _
                "DateCol NVARCHAR(8)," & _
                "Month NVARCHAR(20)," & _
                "FSR_Name NVARCHAR(100)," & _
                "Answered NVARCHAR(20))"

    On Error Resume Next
    conn.Execute createSQL
    If Err.Number <> 0 Then
        Call LogError("Failed to create table '" & tablename & "': " & Err.Description)
        Err.Clear
        EnsureTableExists = False
    Else
        EnsureTableExists = True
    End If
    On Error GoTo 0
End Function

Sub InsertACDData(ws, DateCol, reportPeriod)
    Dim row, user, answered, insertSQL, countInserted
    row = 17 : countInserted = 0

    Do Until Trim(CStr(ws.Cells(row, 1).Value)) = "Summaries Per User And Queue"
        user = Trim(CStr(ws.Cells(row, 1).Value))
        answered = Trim(CStr(ws.Cells(row, 7).Value))

        If user <> "" And user <> "User" And answered <> "" And InStr(user, "/") = 0 Then
            fout.WriteLine PadRight(DateCol, 10) & PadRight(reportPeriod, 10) & PadRight(user, 24) & PadRight(answered, 10)

            insertSQL = "IF NOT EXISTS (SELECT 1 FROM " & tablename & _
                        " WHERE DateCol = '" & SqlSafe(DateCol) & "'" & _
                        " AND Month = '" & SqlSafe(reportPeriod) & "'" & _
                        " AND FSR_Name = '" & SqlSafe(user) & "') " & _
                        "BEGIN " & _
                        "INSERT INTO " & tablename & " (DateCol, Month, FSR_Name, Answered) VALUES (" & _
                        "'" & SqlSafe(DateCol) & "', '" & SqlSafe(reportPeriod) & "', '" & SqlSafe(user) & "', '" & SqlSafe(answered) & "') " & _
                        "END"

            On Error Resume Next
            conn.Execute insertSQL
            If Err.Number <> 0 Then
                Call LogError("Insert failed:" & vbCrLf & _
                              "Row: " & row & ", User: " & user & ", Answered: " & answered & vbCrLf & _
                              "SQL: " & insertSQL & vbCrLf & _
                              "Error: " & Err.Description)
                Err.Clear
            Else
                countInserted = countInserted + 1
            End If
            On Error GoTo 0
        End If
        row = row + 1
    Loop
End Sub

Function ExtractDateFromPath(path)
    Dim folderParts, yr, moName, day, monthNum, dtLastDay, arr
    folderParts = Split(path, "\")
    If UBound(folderParts) < 7 Then Exit Function

    yr = folderParts(5)
    moName = folderParts(6)
    day = folderParts(7)
    If Not IsNumeric(day) Then Exit Function

    monthNum = MonthNumberFromName(moName)
    dtLastDay = DateSerial(CInt(yr), monthNum + 1, 0)
    arr = Split(dtLastDay, "/")
    ExtractDateFromPath = Right("0" & arr(2), 4) & Right("0" & arr(0), 2) & Right("0" & arr(1), 2)
End Function

Function MonthNumberFromName(monthName)
    Dim dt : dt = CDate("01-" & monthName & "-2000")
    MonthNumberFromName = Month(dt)
End Function

Function FormatReportPeriod(periodText)
    Dim arr, dt
    arr = Split(periodText, " ")
    dt = CDate(arr(0))
    FormatReportPeriod = MonthName(Month(dt))
End Function

Function PadRight(val, width)
    PadRight = Left(val & Space(width), width)
End Function

Function SqlSafe(val)
    SqlSafe = Replace(Trim(val), "'", "''")
End Function

Sub LogError(msg)
    If IsEmpty(fout) Then
        MsgBox Now & " | ERROR: " & msg
    Else
        fout.WriteLine Now & " | ERROR: " & msg
    End If
End Sub

Sub LogInfo(msg)
    If Not IsEmpty(fout) Then
        fout.WriteLine Now & " | INFO: " & msg
    End If
End Sub
