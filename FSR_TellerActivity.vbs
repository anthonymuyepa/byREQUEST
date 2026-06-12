' FSR_TellerActivity Report Processor
Option Explicit

Dim folderPath, connString, tablename
folderPath = "R:\byRequest\Reports\Head Teller\Teller Activity\"  ' Root folder
connString = "Provider=SQLOLEDB;Data Source=GOLDMINE-SQL1;Initial Catalog=byREQUEST;Integrated Security=SSPI;"
tablename = "FSR_TellerActivity"

Sub StartDoc()
    Dim fso, xl, yearFolder, monthFolder, file, conn, wb, ws, lastRow, stats
    Dim DateCol, folderMonth, moName, branch, currYear, currMonth

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    Set conn = OpenConnection()
    If conn Is Nothing Then Exit Sub

    Call CreateTableIfNeeded(conn)

	 
	If Month(Date) = 1 Then
		currMonth = "12"
		currYear = Year(Date) - 1
	Else
		currMonth = Right("0" & (Month(Date) - 1), 2)
		currYear = Year(Date)
	End If

    Dim Mnth
For Mnth = 1 to 9

    Dim targetFolderPath
    targetFolderPath = folderPath & currYear & " Teller Activity Totals\" & currYear & "-0" & Mnth & "\"
'MsgBox targetFolderPath
    If fso.FolderExists(targetFolderPath) Then
        For Each file In fso.GetFolder(targetFolderPath).Files
           If InStr(LCase(file.Name), "teller activity") > 0 And _
			   Right(LCase(file.Name), 5) = ".xlsx" And _
			   InStr(file.Name, "~") = 0 Then

				'msgbox "File.Name :" & file.Name
              
                Set wb = xl.Workbooks.Open(file.Path)
                Set ws = wb.Sheets(1)
                branch = CleanBranch(ws.Range("C1").Value)
				
				' Remove trailing " Branch" if it exists (case-insensitive)
				If LCase(Right(branch, 7)) = " branch" Then
					branch = Left(branch, Len(branch) - 7)
				End If
				branch = Trim(branch)

				
				'msgbox "Branch :" & branch
                lastRow = ws.Cells(ws.Rows.Count, "A").End(-4162).Row
								
				Dim k1Val, startDate, endDate, dtLastDay

				k1Val = Trim(CStr(ws.Range("K1").Value))
				k1Val = Replace(k1Val, "'", "")  ' remove leading apostrophe if present
				k1Val = Replace(k1Val, "  ", " ") ' normalize extra spaces

				' Extract date range from format: "04/01/25-04/30/25"
				If InStr(k1Val, "-") > 0 Then
					startDate = Trim(Split(k1Val, "-")(0))
					endDate = Trim(Split(k1Val, "-")(1))

					If IsDate(endDate) Then
						dtLastDay = CDate(endDate)
						moName = MonthName(Month(dtLastDay), False)  ' e.g., April
						DateCol = Year(dtLastDay) & Right("0" & Month(dtLastDay), 2) & Right("0" & Day(dtLastDay), 2)  ' e.g., 20250430
					Else
						LogError "Invalid end date in K1: " & k1Val
						skipFile = True
					End If
				Else
					LogError "Invalid format in K1: " & k1Val
					skipFile = True
				End If
										
				'msgbox "moName :" & moName

				'msgbox "DateCol :" & DateCol
							
                Set stats = SummarizeTellerActivity(ws, 6, lastRow)
                Call InsertStats(conn, stats, DateCol, moName, branch)
	 'msgbox "After InsertStats :" & branch
                wb.Close False
                Set ws = Nothing
                Set wb = Nothing
            End If
        Next
    Else
        LogError "Folder not found: " & targetFolderPath
    End If
	
	Next
	
	

    xl.Quit
    conn.Close
    Set xl = Nothing
    Set conn = Nothing
    Set fso = Nothing
End Sub

Function GetLastDayOfMonth(moName)
    Dim dt, y, m
    y = Year(Date): m = Month(DateValue(moName & " 1"))
    dt = DateSerial(y, m + 1, 0)
    GetLastDayOfMonth = Year(dt) & Right("0" & Month(dt), 2) & Right("0" & Day(dt), 2)
End Function

Function OpenConnection()
    On Error Resume Next
    Dim c: Set c = CreateObject("ADODB.Connection")
    c.Open connString
    If Err.Number <> 0 Then LogError "DB error: " & Err.Description: Set c = Nothing
    On Error GoTo 0
    Set OpenConnection = c
End Function

Sub CreateTableIfNeeded(c)
    Dim sql
    sql = "IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '" & tablename & "') BEGIN " & _
          "CREATE TABLE " & tablename & " (" & _
          "DateCol NVARCHAR(8), " & _
          "Month NVARCHAR(20), " & _
          "Branch NVARCHAR(100), " & _
          "EmpID NVARCHAR(50), " & _
          "EmpName NVARCHAR(100), " & _
          "Total NVARCHAR(50)) END"
    On Error Resume Next
    c.Execute sql
    If Err.Number <> 0 Then LogError "CreateTable failed: " & Err.Description
    On Error GoTo 0
End Sub


Function SummarizeTellerActivity(ws, startRow, endRow)
    Dim dict: Set dict = CreateObject("Scripting.Dictionary")
    Dim r, empID, empName, totalVal

    For r = startRow To endRow
        empID = Trim(CStr(ws.Cells(r, "A").Value))
        empName = CleanName(ws.Cells(r, "C").Value)
        totalVal = Trim(CStr(ws.Cells(r, 28).Value)) ' Column AB

        If empID = "" Or Not IsNumeric(empID) Then
            ' skip
        ElseIf InStr(empID, "Total") > 0 Then
            Exit For
        Else
            If Not dict.Exists(empID) Then
                Dim s: Set s = CreateObject("Scripting.Dictionary")
                s.Add "EmpName", empName
                s.Add "Total", totalVal
                dict.Add empID, s
            End If
        End If
    Next

    Set SummarizeTellerActivity = dict
End Function




Function CleanBranch(val)
    Dim s: s = Trim(CStr(val))
    If Left(s, 1) = "'" Then s = Mid(s, 2)
    CleanBranch = s
End Function



Sub InsertStats(c, statsDict, dateVal, moName, br)
    Dim empID, s, sql
    On Error Resume Next

    c.Execute "DELETE FROM " & tablename & " WHERE DateCol = '" & dateVal & "' AND Branch = '" & br & "'"
    
    If Err.Number <> 0 Then
        LogError "Delete error for " & br & ": " & Err.Description
    End If
    On Error GoTo 0

    fout.WriteLine PadRight("DateCol", 10) & _
                   PadRight("Month", 10) & _
                   PadRight("Branch", 24) & _
                   PadRight("EmpID", 10) & _
                   PadRight("EmpName", 24) & _
                   PadRight("Total", 10)

    For Each empID In statsDict.Keys
        Set s = statsDict(empID)
        sql = "INSERT INTO " & tablename & " (DateCol, Month, Branch, EmpID, EmpName, Total) VALUES (" & _
              "'" & dateVal & "', '" & moName & "', '" & br & "', '" & empID & "', '" & s("EmpName") & "', '" & s("Total") & "')"

        On Error Resume Next
        fout.WriteLine PadRight(dateVal, 10) & _
                       PadRight(moName, 10) & _
                       PadRight(br, 24) & _
                       PadRight(empID, 10) & _
                       PadRight(s("EmpName"), 24) & _
                       PadRight(s("Total"), 10)

        c.Execute sql
        If Err.Number <> 0 Then LogError "Insert error (" & empID & "): " & Err.Description
        On Error GoTo 0
    Next
End Sub


' === LOGGING MODULES ===
Sub LogInfo(msg)
      fout.WriteLine Now & " | INFO: " & msg
End Sub

Sub LogError(msg)
    fout.WriteLine Now & " | ERROR: " & msg
End Sub

Function CleanString(val)
    Dim s
    s = Trim(CStr(val))
    If Left(s, 1) = "'" Then s = Mid(s, 2)
    If InStr(s, "----") > 0 Then s = ""
    CleanString = s
End Function

Function CleanNumber(val)
    Dim s
    s = Trim(CStr(val))
    If Left(s, 1) = "'" Then s = Mid(s, 2)
    If InStr(s, "----") > 0 Then
        CleanNumber = 0
    ElseIf IsNumeric(s) Then
        CleanNumber = CDbl(s)
    Else
        CleanNumber = 0
    End If
End Function

Function CleanName(val)
    Dim s, dashPos
    On Error Resume Next

    s = Trim(CStr(val))

    If Left(s, 1) = "'" Then s = Mid(s, 2)

    If Len(s) > 0 Then
        If s Like "*-#" Or s Like "*-##" Or s Like "*-###"Then
            dashPos = InStrRev(s, "-")
            If dashPos > 0 Then s = Left(s, dashPos - 1)
        End If
    End If

    On Error GoTo 0
    CleanName = Trim(s)
End Function



Function PadRight(val, width)
    PadRight = Left(val & Space(width), width)
End Function
