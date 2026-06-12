Option Explicit

' === CONFIG ===
Dim excelPath, tablename, connString, conn

'connString = "Provider=SQLOLEDB;Data Source=BYRequest\SQLEXPRESS;Initial Catalog=byREQUEST;Integrated Security=SSPI;"


'	excelPath = "R:\byREQUEST\reports\Phonesystem\Monthly\2025\May\01\Branch ACD Monthly Report.xls"
	tablename = "acd_answered_stats"
	connString = "Provider=SQLOLEDB;" & _
				 "Data Source=GOLDMINE-SQL1,1433;" & _
				 "Initial Catalog=byREQUEST;" & _
				 "User ID=sa;" & _
				 "Password=H1llary;"


' === Initialize Logging ===
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")


Sub StartDoc() 
   Dim months, m, d, yrNow, filePath 
   Dim tday 
   tday = Date
    
   yrNow = Right(tday, 4)
	 

  months = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")

    For m = 0 To 11
        For d = 1 To 2
            filePath = "R:\byRequest\Reports\phonesystem\Monthly\" & yrNow & "\" & months(m) & "\" & "0"& d & "\Branch ACD Monthly Report.xls"
		MsgBox "Month exists " & filePath
            If fso.FileExists(filePath) Then
			MsgBox "EYs Month exists " & filePath
                excelPath = filePath
				
                On Error Resume Next
                Set conn = CreateObject("ADODB.Connection")
                conn.Open connString
                If conn.State <> 1 Then
                    Call LogError("Failed to open database connection." & vbCrLf & _
                                  "Error: " & Err.Description)
                    On Error GoTo 0
                    Exit Sub
                End If
                On Error GoTo 0

                Dim rs, tableExists
                tableExists = False
                On Error Resume Next
                Set rs = conn.Execute("SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '" & tablename & "'")
                If Not rs Is Nothing Then
                    If Not rs.EOF Then tableExists = True
                    rs.Close
                End If
                On Error GoTo 0

                If Not tableExists Then
                    Dim createSQL
                    createSQL = "CREATE TABLE " & tablename & " (" & _
                                "report_date NVARCHAR(8)," & _
                                "report_period NVARCHAR(20)," & _
                                "user_name NVARCHAR(100)," & _
                                "answered_count NVARCHAR(20))"
                    On Error Resume Next
                    conn.Execute createSQL
                    If Err.Number <> 0 Then
                        Call LogError("Failed to create table '" & tablename & "': " & Err.Description)
                        conn.Close : Set conn = Nothing
                        Exit Sub
                    End If
                    On Error GoTo 0
                End If

                Dim xl, wb, ws
                Set xl = CreateObject("Excel.Application")
                xl.Visible = False
                On Error Resume Next
                Set wb = xl.Workbooks.Open(excelPath)
                If Err.Number <> 0 Then
                    Call LogError("Failed to open Excel file: " & excelPath & vbCrLf & "Error: " & Err.Description)
                    xl.Quit : Set xl = Nothing
                    conn.Close : Set conn = Nothing
                    Exit Sub
                End If
				MsgBox "here"
                Set ws = wb.Sheets(1)
                If ws Is Nothing Then
                    Call LogError("Worksheet not found in Excel file: " & excelPath)
                    wb.Close False : xl.Quit
                    Set xl = Nothing
                    conn.Close : Set conn = Nothing
                    Exit Sub
                End If
                On Error GoTo 0

                Dim periodText, reportPeriod, reportDate
                periodText = ws.Range("F4").Value
                If periodText = "" Then
                    Call LogError("Missing period data in cell F4")
                    wb.Close False : xl.Quit : conn.Close
                    Exit Sub
                End If
                reportPeriod = FormatReportPeriod(periodText)

                Dim folderPath, folderParts, year, monthName, day, monthNum
                folderPath = Left(excelPath, InStrRev(excelPath, "\") - 1)
                folderParts = Split(folderPath, "\")
                If UBound(folderParts) < 7 Then
                    Call LogError("Unable to extract report date from path: " & excelPath)
                    wb.Close False : xl.Quit : conn.Close
                    Exit Sub
                End If

                year = folderParts(5)
                monthName = folderParts(6)
                day = folderParts(7)
                If Not IsNumeric(day) Then
                    Call LogError("Invalid day folder in path: " & day)
                    wb.Close False : xl.Quit : conn.Close
                    Exit Sub
                End If
                monthNum = MonthNumberFromName(monthName)
                reportDate = year & Right("0" & monthNum, 2) & Right("0" & day, 2)

                On Error Resume Next
                conn.Execute "DELETE FROM " & tablename & " WHERE report_date = '" & reportDate & "' AND report_period = '" & reportPeriod & "'"
                If Err.Number <> 0 Then
                    Call LogError("Error deleting existing data: " & Err.Description)
                    Err.Clear
                End If
                On Error GoTo 0

                Dim row, user, answered, insertSQL, countInserted
                row = 17 : countInserted = 0
                Do Until Trim(CStr(ws.Cells(row, 1).Value)) = "Summaries Per User And Queue"
                    user = Trim(CStr(ws.Cells(row, 1).Value))
                    answered = Trim(CStr(ws.Cells(row, 7).Value))
                    If user <> "" And user <> "User" And answered <> "" And InStr(user, "/") = 0 Then
                        insertSQL = "INSERT INTO " & tablename & _
                                    " (report_date, report_period, user_name, answered_count) VALUES (" & _
                                    "'" & reportDate & "', '" & reportPeriod & "', '" & Replace(user, "'", "''") & "', '" & Replace(answered, "'", "''") & "')"
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

                Call LogInfo("Import complete for " & excelPath & ". Rows inserted: " & countInserted)
                wb.Close False : xl.Quit
                Set ws = Nothing : Set wb = Nothing : Set xl = Nothing
                conn.Close : Set conn = Nothing
            End If
        Next
    Next
End Sub




' === Helpers ===

Function MonthNumberFromName(monthName)
    Dim dt : dt = CDate("01-" & monthName & "-2000")
    MonthNumberFromName = Month(dt)
End Function

Function FormatReportPeriod(periodText)
    Dim arr, dt
    arr = Split(periodText, " ")
    dt = CDate(arr(0))
    FormatReportPeriod = MonthName(Month(dt)) & " " & Year(dt)
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



Sub ProcessAllACDReports()
    Dim fso, basePath, yearFolder, monthFolder, dayFolder, filePath
    Set fso = CreateObject("Scripting.FileSystemObject")
    basePath = "R:\byRequest\Reports\phonesystem\Monthly"

    If Not fso.FolderExists(basePath) Then
        Call LogError("Base folder not found: " & basePath)
        Exit Sub
    End If

    Dim folderYear, folderMonth, folderDay, reportFile
    For Each folderYear In fso.GetFolder(basePath).SubFolders
        For Each folderMonth In folderYear.SubFolders
            For Each folderDay In folderMonth.SubFolders
                reportFile = folderDay.Path & "\Branch ACD Monthly Report.xls"
                If fso.FileExists(reportFile) Then
                    Call StartDoc(reportFile)
                Else
                    Call LogInfo("No report file found in: " & folderDay.Path)
                End If
            Next
        Next
    Next
End Sub
