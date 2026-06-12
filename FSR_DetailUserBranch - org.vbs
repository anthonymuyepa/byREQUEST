Option Explicit

' === CONFIG ===
  
    Dim folderPath, connString, tablename
	Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
	'connString = "Provider=SQLOLEDB;Data Source=BYRequest\SQLEXPRESS;Initial Catalog=byREQUEST;Integrated Security=SSPI;"


    folderPath = "R:\byrequest\reports\Lobby Tracking\"
	tablename = "FSR_DetailUserBranch"
	connString = "Provider=SQLOLEDB;" & _
				 "Data Source=GOLDMINE-SQL1;" & _
				 "Initial Catalog=byREQUEST;" & _
				 "Integrated Security=SSPI;"




Sub StartDoc()
    On Error Resume Next

    Dim currMonth, entryYear, m, dtLastDay, entryDate, monthAbbr
    Dim conn, createSQL, file, wb, ws, row, lastRow
    Dim branch, empName, avgTime, daysWorked, upsertSQL, xl, excelPath

    currMonth = Month(Date)
    If currMonth = 1 Then
        m = 12 : entryYear = Year(Date) - 1
    Else
        m = currMonth - 1 : entryYear = Year(Date)
    End If

    dtLastDay = DateSerial(entryYear, m + 1, 0)
    entryDate = Year(dtLastDay) & Right("0" & Month(dtLastDay), 2) & Right("0" & Day(dtLastDay), 2)
    monthAbbr = MonthName(m, True)
	
	
	
						' === Write header once before loop ===
	fout.WriteLine PadRight("DateCol", 12) & _
               PadRight("Branch", 20) & _
               PadRight("EmpName", 24) & _
               PadRight("AvgAssistTime", 18) & _
               PadRight("DaysWorked", 12) & _
               PadRight("Month", 10)


    Set conn = CreateObject("ADODB.Connection")
    conn.Open connString
    If conn.State <> 1 Then
        Call LogError("DB connection failed: " & Err.Description)
        Exit Sub
    End If

    createSQL = _
        "IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '" & tablename & "') " & _
        "BEGIN " & _
        "CREATE TABLE " & tablename & " (" & _
        "DateCol NVARCHAR(8), " & _
        "Branch NVARCHAR(100), " & _
        "EmpName NVARCHAR(100), " & _
        "AvgAssistTime NVARCHAR(50), " & _
        "DaysWorked INT, " & _
        "Month NVARCHAR(10)) " & _
        "END"
    conn.Execute createSQL
    If Err.Number <> 0 Then Call LogError("Create table error: " & Err.Description)

    If Not fso.FolderExists(folderPath) Then
        Call LogError("Base path not found: " & folderPath)
        Exit Sub
    End If

    Set xl = CreateObject("Excel.Application")
    xl.Visible = False

    Dim months, folderMonth, monthPath
    months = Array("January", "February", "March", "April", "May", "June", _
                   "July", "August", "September", "October", "November", "December")

    For m = 0 To UBound(months)
        folderMonth = months(m)
        monthPath = folderPath & folderMonth & "\"

        If fso.FolderExists(monthPath) Then
            'Call LogInfo("Processing: " & monthPath)

            For Each file In fso.GetFolder(monthPath).Files
                If LCase(fso.GetExtensionName(file.Name)) = "xlsx" And _
                   InStr(LCase(file.Name), "~") = 0 And _
                   InStr(LCase(file.Name), "detail user and branch overview report") > 0 Then

                    excelPath = file.Path
                    'Call LogInfo("Reading: " & excelPath)

                    Dim skipFile : skipFile = False
                    Set wb = Nothing
                    Set ws = Nothing
                    On Error Resume Next
                    Set wb = xl.Workbooks.Open(excelPath)
                    If Err.Number <> 0 Or wb Is Nothing Then
                        Call LogError("Cannot open Excel: " & excelPath)
                        skipFile = True
                        Err.Clear
                    End If
                    On Error GoTo 0

                   If Not skipFile Then
					Set ws = wb.Sheets(1)

					' Get Branch and ColDate from B6
					Dim headerParts, dateParts, ColDate
					headerParts = Split(ws.Range("B6").Value, "|")
					Branch = Trim(Split(headerParts(3), ":")(1))
					dateParts = Split(headerParts(2), "to")
					
					Dim dt : dt = CDate(Trim(dateParts(1))) ' Last day of the range
					 
					ColDate = Year(dt) & Right("0" & Month(dt), 2) & Right("0" & Day(dt), 2)

					' Get last used row
					Dim maxRows
					maxRows = ws.Cells(ws.Rows.Count, 2).End(-4162).Row ' xlUp

					' Find start of "Associate Stats"
					For row = 1 To maxRows
						If Trim(ws.Cells(row, 2).Value) = "Associate Stats" Then
							row = row + 1
							Exit For
						End If
					Next

					' Extract rows under "Associate Stats"

					Do While row <= maxRows
						empName = ws.Cells(row, 2).Value

						If empName <> "" And InStr(1, ws.Cells(row, 7).Value, "Days Worked:", vbTextCompare) > 0 Then
							daysWorked = Trim(ws.Cells(row, 10).Value) ' Column J
							avgTime = Trim(ws.Cells(row + 2, 7).Value) ' Column G, 2 rows down
							avgTime = ConvertToMinutes(avgTime)

							If IsNumeric(daysWorked) = False Then daysWorked = 0
							
														' Log padded output
							fout.WriteLine PadRight(ColDate, 12) & _
										   PadRight(Branch, 20) & _
										   PadRight(empName, 24) & _
										   PadRight(avgTime, 18) & _
										   PadRight(daysWorked, 12) & _
										   PadRight(folderMonth, 10)


							' SQL-safe insert
							upsertSQL = "IF NOT EXISTS (SELECT 1 FROM " & tablename & " WHERE " & _
										"DateCol = '" & SqlSafe(ColDate) & "' AND " & _
										"Branch = '" & SqlSafe(Branch) & "' AND " & _
										"EmpName = '" & SqlSafe(empName) & "') " & _
										"BEGIN " & _
										"INSERT INTO " & tablename & _
										" (DateCol, Branch, EmpName, AvgAssistTime, DaysWorked, Month) VALUES (" & _
										"'" & SqlSafe(ColDate) & "', " & _
										"'" & SqlSafe(Branch) & "', " & _
										"'" & SqlSafe(empName) & "', " & _
										"'" & SqlSafe(avgTime) & "', " & _
										daysWorked & ", " & _
										"'" & SqlSafe(folderMonth) & "') " & _
										"END"


							conn.Execute upsertSQL
							If Err.Number <> 0 Then
								Call LogError("Insert error: " & Err.Description)
								Err.Clear
							End If

							row = row + 3 ' Skip to next employee
						Else
							row = row + 1
						End If
					Loop


					wb.Close False
					Set wb = Nothing
				End If


                    Set wb = Nothing : Set ws = Nothing
                End If
            Next
        Else
            'Call LogInfo("Skipping missing folder: " & monthPath)
        End If
    Next

    If Not conn Is Nothing Then conn.Close : Set conn = Nothing
    If Not xl Is Nothing Then xl.Quit : Set xl = Nothing
End Sub




Sub LogError(msg)
    If fout Is Nothing Then
        Set fout = fso.OpenTextFile("FSR_ErrorLog.txt", 8, True)
    End If
    fout.WriteLine Now & " | ERROR: " & msg
End Sub

Function IsEmpLine(val)
    If Len(val) > 4 And IsNumeric(Left(val, 3)) And Mid(val, 4, 1) = " " Then
        IsEmpLine = True
    Else
        IsEmpLine = False
    End If
End Function

Sub LogInfo(msg)
    If Not IsEmpty(fout) Then
        fout.WriteLine Now & " | INFO: " & msg
    End If
End Sub


' === Helper Function to Escape SQL Strings ===
Function SqlSafe(val)
    If IsNull(val) Or IsEmpty(val) Then
        SqlSafe = ""
    Else
        SqlSafe = Replace(CStr(val), "'", "''")
    End If
End Function


Function ConvertToMinutes(avgTimeText)
    Dim minutes, seconds, part, parts
    minutes = 0
    seconds = 0

    avgTimeText = LCase(Trim(avgTimeText))
    parts = Split(avgTimeText, ",")

    For Each part In parts
        part = Trim(part)
        If InStr(part, "minute") > 0 Then
            part = Replace(part, "minutes", "")
            part = Replace(part, "minute", "")
            If IsNumeric(Trim(part)) Then minutes = CLng(Trim(part))
        ElseIf InStr(part, "second") > 0 Then
            part = Replace(part, "seconds", "")
            part = Replace(part, "second", "")
            If IsNumeric(Trim(part)) Then seconds = CLng(Trim(part))
        End If
    Next

    ConvertToMinutes = Round(minutes + (seconds / 60), 2)
End Function


Function PadRight(val, width)
    PadRight = Left(val & Space(width), width)
End Function
