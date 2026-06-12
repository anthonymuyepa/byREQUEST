Option Explicit

' === CONFIG ===
  
    Dim folderPath, connString, tablename
	Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
	'connString = "Provider=SQLOLEDB;Data Source=BYRequest\SQLEXPRESS;Initial Catalog=byREQUEST;Integrated Security=SSPI;"


    folderPath = "E:\File Maintenance\"
	tablename = "FSR_BranchMaintenance"
	connString = "Provider=SQLOLEDB;" & _
				 "Data Source=GOLDMINE-SQL1;" & _
				 "Initial Catalog=byREQUEST;" & _
				 "Integrated Security=SSPI;"

Sub StartDoc()

    On Error Resume Next

    Dim currMonth, entryYear, m, dtLastDay, entryDate, monthAbbr
    Dim conn, createSQL, file, wb, ws, sheet, row, lastRow, cellValue
    Dim branch, empId, empName, totalEntries, upsertSQL
    Dim xl

    currMonth = Month(Date)

    If currMonth = 1 Then
        m = 12
        entryYear = Year(Date) - 1
    Else
        m = currMonth - 1
        entryYear = Year(Date)
    End If

    dtLastDay = DateSerial(entryYear, m + 1, 0)
    entryDate = Year(dtLastDay) & Right("0" & Month(dtLastDay), 2) & Right("0" & Day(dtLastDay), 2)
    monthAbbr = MonthName(m, True)

                ' === Write header once ===
	fout.WriteLine PadRight("DateCol", 10) & _
               PadRight("Month", 10) & _
               PadRight("Branch", 24) & _
               PadRight("EmpID", 10) & _
               PadRight("EmpName", 24) & _
               PadRight("TotalEntries", 14)




    Set conn = CreateObject("ADODB.Connection")
    conn.Open connString

    If conn.State <> 1 Then
        Call LogError("Failed to open database connection." & vbCrLf & "Error: " & Err.Description)
        Exit Sub
    End If

    createSQL = "IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '" & tablename & "') " & _
                "CREATE TABLE " & tablename & " (" & _
                "DateCol NVARCHAR(8)," & _
                "Branch NVARCHAR(100)," & _
                "EmpName NVARCHAR(100)," & _
                "EmpId NVARCHAR(10)," & _
                "TotalEntries NVARCHAR(10)," & _
                "Month NVARCHAR(10))"
    conn.Execute createSQL
    If Err.Number <> 0 Then
        Call LogError("Error creating table: " & Err.Description)
        Err.Clear
    End If
    On Error GoTo 0

    If Not fso.FolderExists(folderPath) Then
        Call LogError("Base path not found: " & folderPath)
        Exit Sub
    End If

    Set xl = CreateObject("Excel.Application")
    If xl Is Nothing Then
        fout.WriteLine "Failed to create Excel application"
        fout.Close
        WScript.Quit
    End If
    xl.Visible = False
 
    For Each file In fso.GetFolder(folderPath).Files
        If (LCase(fso.GetExtensionName(file.Name)) = "xlsx" Or LCase(fso.GetExtensionName(file.Name)) = "xls") And InStr(file.Name, "~") = 0 Then
            On Error Resume Next
            Set wb = xl.Workbooks.Open(file.Path)

            If wb Is Nothing Or Err.Number <> 0 Then
                fout.WriteLine "ERROR opening: " & file.Name & " → " & Err.Description
                Err.Clear
            Else

				Dim sheetFound
				sheetFound = False

				' Try to find a sheet that contains "File Maintenance" in its name (case-insensitive)
				For Each sheet In wb.Sheets
					If InStr(1, sheet.Name, "File Maintenance", vbTextCompare) > 0 _
					   Or wb.Sheets.Count = 1 Then
			
						Set ws = sheet
						sheetFound = True
						Exit For
					End If
				Next


				' If not found, fall back to the first visible sheet
				If Not sheetFound Then
					For Each sheet In wb.Sheets
						If sheet.Visible Then
							Set ws = sheet
							Exit For
						End If
					Next
				End If


                branch = CleanText(ws.Range("A2").Value)
                row = 3
                lastRow = ws.UsedRange.Rows.Count
				'fout.WriteLine file.Name & "Branch :" & branch & " Rows Count: " & lastRow
				

				Do While row <= lastRow
					cellValue = Trim(ws.Cells(row, 1).Value)

					If IsEmployeeHeader(cellValue) Then
						empId = ExtractFirstNumber(cellValue)
						empName = ExtractName(cellValue)

						totalEntries = 0
						row = row + 1

						Do While row <= lastRow
							cellValue = Trim(ws.Cells(row, 1).Value)

							If cellValue = "" Then
								row = row + 1
							ElseIf IsEmployeeHeader(cellValue) Then
								Exit Do
							Else
								If IsFMIssue(cellValue) Then totalEntries = totalEntries + 1
								row = row + 1
							End If
						Loop

						' === Write padded row to log ===
						fout.WriteLine PadRight(entryDate, 10) & _
									   PadRight(monthAbbr, 10) & _
									   PadRight(branch, 24) & _
									   PadRight(empId, 10) & _
									   PadRight(empName, 24) & _
									   PadRight(totalEntries, 14)

						' === Upsert SQL logic ===
						upsertSQL = _
							"IF EXISTS (SELECT * FROM " & tablename & _
							" WHERE EmpName = '" & Replace(empName, "'", "''") & "'" & _
							" AND Month = '" & monthAbbr & "'" & _
							" AND Branch = '" & Replace(branch, "'", "''") & "') " & _
							"BEGIN " & _
								"UPDATE " & tablename & " SET " & _
								"DateCol = '" & entryDate & "', " & _
								"EmpId = '" & empId & "', " & _
								"TotalEntries = " & totalEntries & " " & _
								"WHERE EmpName = '" & Replace(empName, "'", "''") & "' " & _
								"AND Month = '" & monthAbbr & "' " & _
								"AND Branch = '" & Replace(branch, "'", "''") & "' " & _
							"END " & _
							"ELSE " & _
							"BEGIN " & _
								"INSERT INTO " & tablename & _
								" (DateCol, Branch, EmpName, EmpId, TotalEntries, Month) VALUES (" & _
								"'" & entryDate & "', '" & Replace(branch, "'", "''") & "', '" & Replace(empName, "'", "''") & "', '" & empId & "', " & totalEntries & ", '" & monthAbbr & "') " & _
							"END"

						conn.Execute upsertSQL
						If Err.Number <> 0 Then
							Call LogError("SQL Error for " & empName & " (" & empId & ") in branch " & branch & ": " & Err.Description)
							Err.Clear
						End If
					Else
						row = row + 1
					End If
				Loop


                wb.Close False
                Set ws = Nothing
                Set wb = Nothing
            End If

            On Error GoTo 0
        Else
            'fout.WriteLine "Skipped temp file: " & file.Name
        End If
    Next

    On Error Resume Next
    If Not conn Is Nothing Then conn.Close : Set conn = Nothing
    If Not xl Is Nothing Then xl.Quit : Set xl = Nothing
    On Error GoTo 0

End Sub







Sub LogError(msg)
    'fout.WriteLine Now & " | ERROR: " & msg
End Sub

Function IsEmployeeHeader(val)
    val = Trim(val)

    ' Reject blank lines
    If val = "" Then
        IsEmployeeHeader = False
        Exit Function
    End If
	    ' Reject if the value is "Date" (case-insensitive)
    If LCase(val) = "date" Then
        IsEmployeeHeader = False
        Exit Function
    End If

    ' Reject if VBScript can parse it as a date
    If IsDate(val) Then
        IsEmployeeHeader = False
        Exit Function
    End If

    ' Reject if it matches common date-like text formats
    Dim regEx
    Set regEx = New RegExp
    regEx.IgnoreCase = True
    regEx.Global = False

    ' Pattern 1: MM/DD/YYYY or M/D/YYYY
    regEx.Pattern = "^\d{1,2}/\d{1,2}/\d{4}$"
    If regEx.Test(val) Then
        IsEmployeeHeader = False
        Exit Function
    End If

    ' Pattern 2: YYYY-MM-DD (ISO format)
    regEx.Pattern = "^\d{4}-\d{1,2}-\d{1,2}$"
    If regEx.Test(val) Then
        IsEmployeeHeader = False
        Exit Function
    End If

    ' Pattern 3: DD-MMM-YYYY (e.g., 28-Mar-2024)
    regEx.Pattern = "^\d{1,2}-[A-Za-z]{3}-\d{4}$"
    If regEx.Test(val) Then
        IsEmployeeHeader = False
        Exit Function
    End If

    ' Otherwise, it's a valid employee header
    IsEmployeeHeader = True
End Function



Sub LogInfo(msg)
    If Not IsEmpty(fout) Then
        'fout.WriteLine Now & " | INFO: " & msg
    End If
End Sub


Function CleanText(txt)
    If IsNull(txt) Then
        CleanText = ""
    Else
        ' Replace non-breaking space, remove tabs/newlines, trim
        CleanText = Trim(Replace(Replace(Replace(txt, Chr(160), " "), vbTab, ""), vbCrLf, ""))
    End If
End Function




Function ExtractFirstNumber(inputStr)
    Dim regEx, matches, result
    Set regEx = New RegExp
    regEx.Pattern = "^\d+"
    regEx.IgnoreCase = True
    regEx.Global = False

    If regEx.Test(inputStr) Then
        Set matches = regEx.Execute(inputStr)
        result = matches(0).Value
    Else
        result = ""
    End If

    ExtractFirstNumber = result
End Function



 
Function ExtractName(inputStr)
    Dim regEx, result
    Set regEx = New RegExp
    ' Match: number, optional / or // and number, then capture the rest
    regEx.Pattern = "^\d+(\/\/?\d+)?(.*)"
    regEx.IgnoreCase = True
    regEx.Global = False

    If regEx.Test(inputStr) Then
        Dim matches
        Set matches = regEx.Execute(inputStr)
        result = matches(0).SubMatches(1) ' Capture group 2: the name
    Else
        result = inputStr ' Return original if no match
		
    End If
	result = Replace(result, "/", "") 
    ExtractName = Trim(result)
End Function

Function IsFMIssue(val)
    val = Trim(val)

    ' Return True if blank
    If val = "" Then
        IsNotEmployeeHeader = True
        Exit Function
    End If

    ' Return True if VBScript can parse it as a date
    If IsDate(val) Then
        IsNotEmployeeHeader = True
        Exit Function
    End If

    ' Check common date-like formats
    Dim regEx
    Set regEx = New RegExp
    regEx.IgnoreCase = True
    regEx.Global = False

    ' Pattern 1: MM/DD/YYYY or M/D/YYYY
    regEx.Pattern = "^\d{1,2}/\d{1,2}/\d{4}$"
    If regEx.Test(val) Then
        IsNotEmployeeHeader = True
        Exit Function
    End If

    ' Pattern 2: YYYY-MM-DD (ISO format)
    regEx.Pattern = "^\d{4}-\d{1,2}-\d{1,2}$"
    If regEx.Test(val) Then
        IsNotEmployeeHeader = True
        Exit Function
    End If

    ' Pattern 3: DD-MMM-YYYY (e.g., 28-Mar-2024)
    regEx.Pattern = "^\d{1,2}-[A-Za-z]{3}-\d{4}$"
    If regEx.Test(val) Then
        IsNotEmployeeHeader = True
        Exit Function
    End If

    ' Otherwise, it's not a date-like or invalid header
    IsNotEmployeeHeader = False
End Function




Function PadRight(val, width)
    PadRight = Left(val & Space(width), width)
End Function
