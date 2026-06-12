Option Explicit

Dim folderPath, connString, tablename, thisYear, fso,lineHeader, lineReport


folderPath = "S:\EastTexas\"
connString = "Provider=SQLOLEDB;" & _
             "Data Source=GOLDMINE-SQL1;" & _
             "Initial Catalog=byREQUEST;" & _
             "Integrated Security=SSPI;"


Set fso = CreateObject("Scripting.FileSystemObject")

Sub StartDoc()
    Dim xl, file, branch
	Dim conn1, conn2, conn3
    Dim thisYear
    thisYear = Year(Date)


	Set conn1 = OpenConnection()
	Set conn2 = OpenConnection()
	Set conn3 = OpenConnection()

	Set conn1 = OpenConnection()
	If conn1 Is Nothing Then Exit Sub

	Set conn2 = OpenConnection()
	If conn2 Is Nothing Then
		conn1.Close : Set conn1 = Nothing
		Exit Sub
	End If

	Set conn3 = OpenConnection()
	If conn3 Is Nothing Then
		conn1.Close : Set conn1 = Nothing
		conn2.Close : Set conn2 = Nothing
		Exit Sub
	End If


    If Not fso.FolderExists(folderPath) Then
        MsgBox "Base folder not found"
        Call LogError("Folder not found: " & folderPath)
        Exit Sub
    End If
	
lineHeader = PadRight("Branch", 18) & _
			 PadRight("Month", 6) & _
             PadRight("EmpID", 12) & _
             PadRight("EmpName", 24) & _
             PadRight("ProdSold", 12) & _
             PadRight("PSProvided", 12) & _
             PadRight("AvgAstTime", 12) & _
             PadRight("DaysWorked", 12) & _
             PadRight("BookedLns", 12) & _
             PadRight("SumBkLoans", 12) & _
             PadRight("LoanRecap", 12) & _
             PadRight("SumRecap", 12) & _
             PadRight("CDLnElig", 12) & _
             PadRight("SumCDElig", 12) & _
             PadRight("CDLnProt", 12) & _
             PadRight("SumCDProt", 12) & _
             PadRight("CLLnElig", 12) & _
             PadRight("SumCLElig", 12) & _
             PadRight("CLLnProt", 12) & _
             PadRight("SumCLProt", 12) & _
             PadRight("MMP", 12) & _
             PadRight("GAP", 12) & _
             PadRight("NewSav", 12) & _
             PadRight("NewDrafts", 12) & _
             PadRight("TotHours", 12) & _
             PadRight("AnsCalls", 12) & _
             PadRight("TotEntrs", 12) & _
             PadRight("TellerTot", 12)


	
	
	
	
    ' Open Excel
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False


	


    ' Loop through matching files
    For Each file In fso.GetFolder(folderPath).Files
        If LCase(Right(file.Name, 5)) = ".xlsx" And _
           InStr(LCase(file.Name), "~") = 0 And _
		   InStr(LCase(file.Name), "copy") = 0 And _
           InStr(LCase(file.Name), LCase(thisYear & " fsr production report")) > 0 Then

           branch = ExtractBranchFromFilename(file.Name)
           Call ProcessWorkbook(file.Path, branch, xl)
        Else
           'fout.WriteLine "Skipped file (no match): " & file.Name
        End If
    Next

    fout.WriteLine Now & " | File scan complete." & vbCrLf

    ' Cleanup
    If Not conn Is Nothing Then conn.Close : Set conn = Nothing
    If Not xl Is Nothing Then xl.Quit : Set xl = Nothing
End Sub

Sub ProcessWorkbook(filePath, branch, xl)
    On Error Resume Next

    Dim wb, sheet, empMeta, empName, empID
    Dim colDate, mnth, m, targetMonth
    Dim rs

    Set wb = xl.Workbooks.Open(filePath, False, False)
    fout.WriteLine "Processing: " & fso.GetFileName(filePath)

	fout.WriteLine lineHeader
	
    For Each sheet In wb.Sheets

        empMeta = ParseEmpSheet(sheet.Name)
        If empMeta <> "" Then
            empName = Split(empMeta, "|")(0)
            empID = Split(empMeta, "|")(1)
            targetMonth = Month(Now) - 1

            For m = 1 To targetMonth
			
                colDate = GetYYYYMMDD(DateSerial(Year(Now), m + 1, 0))
                mnth = GetMonthAbbreviation(DateSerial(Year(Now), m + 1, 0))
				
				lineReport = PadRight(branch, 18) & PadRight(mnth, 6) & PadRight(empID, 12) & PadRight(empName, 24)



				
				
                ' === SECTION 1: Total Accounts Assisted Service Products, Members Assisted ===
                Set rs = Fetch_FSR_TotalAccountsAssisted(branch, empName, colDate)
                If Not rs Is Nothing Then WriteMonthlyTotalAccountsAssisted sheet, rs, mnth

                ' === SECTION 2: Detail User Branch, Avg Assist Time, Days Worked ===
                Set rs = Fetch_FSR_DetailUserBranch(branch, empName, colDate)
                If Not rs Is Nothing Then WriteMonthlyDetailUserBranch sheet, rs, mnth

                ' === SECTION 3: New Loan Report, Total Bookes Loans, Recaptured Loans, Credit Loans Eligible, Credit Life, MMP, GAP ===
                Set rs = Fetch_FSR_NewLoanReportbyBranch(branch, empID, colDate)
                If Not rs Is Nothing Then WriteMonthlyNewLoanReportbyBranch sheet, rs, mnth
				
				
				' === SECTION 4: New Savings ===
                Set rs = Fetch_FSR_SharesOpenedbyBranch(branch, empID, colDate)
                If Not rs Is Nothing Then WriteMonthlySharesOpenedbyBranch sheet, rs, mnth
				
				' === SECTION 5: New Checking ===
                Set rs = Fetch_FSR_ShareDFTSOpenedbyBranch(branch, empID, colDate)
                If Not rs Is Nothing Then WriteMonthlyShareDFTSOpenedbyBranch sheet, rs, mnth
				
				
				
				' === SECTION 6: Attendance TimeDetailReport ===
                Set rs = Fetch_FSR_TimeDetailReport(branch, empName, colDate)
                If Not rs Is Nothing Then WriteMonthlyTimeDetailReport sheet, rs, mnth
				
                ' === SECTION 7: Teller Transactions TellerActivity ===
                Set rs = Fetch_FSR_TellerActivity(branch, empID, colDate)
                If Not rs Is Nothing Then WriteMonthlyTellerActivity sheet, rs, mnth
								
				
                ' === SECTION 8: ACD Calls ===
                Set rs = Fetch_FSR_AcdCalls(branch, empName, colDate)
                If Not rs Is Nothing Then WriteMonthlyAcdCalls sheet, rs, mnth

                ' === SECTION 9: File Maintenance     Branch Maintenance ===
                Set rs = Fetch_FSR_BranchMaintenance(branch, empName, colDate)
                If Not rs Is Nothing Then WriteMonthlyBranchMaintenance sheet, rs, mnth

				fout.WriteLine lineReport
				lineReport =""
				


            Next
        End If
		
		
    Next


	' === Save updated workbook copy with versioning ===
    Call SaveWorkbookWithVersion(wb, filePath, mnth)
	
	
    wb.Close True
    Set wb = Nothing
    fout.WriteLine "Workbook complete: " & fso.GetFileName(filePath) & vbCrLf
End Sub








Function OpenConnection()
    On Error Resume Next
    Dim conn
    Set conn = CreateObject("ADODB.Connection")
    
    conn.ConnectionString = "Provider=SQLOLEDB;" & _
                            "Data Source=GOLDMINE-SQL1;" & _
                            "Initial Catalog=byREQUEST;" & _
                            "Integrated Security=SSPI;" & _
                            "OLE DB Services=-1;"  ' Disable pooling and caching

    conn.Open
    If Err.Number <> 0 Then
        fout.WriteLine "ERROR: Failed to open DB connection: " & Err.Description
        Set OpenConnection = Nothing
    Else
        Set OpenConnection = conn
    End If
    On Error GoTo 0
End Function



Sub LogError(msg)
    fout.WriteLine Now & " | ERROR: " & msg
End Sub


Function ExtractBranchFromFilename(filename)
    Dim nameOnly, splitPoint
    nameOnly = fso.GetBaseName(filename) ' strip extension

    splitPoint = InStr(1, nameOnly, "Report-", vbTextCompare)
    If splitPoint > 0 Then
        ExtractBranchFromFilename = Trim(Mid(nameOnly, splitPoint + 7)) ' "Report-" = 7 chars
    Else
        ExtractBranchFromFilename = ""
    End If
End Function


Sub ListWorkbookSheets(filePath, xl)
    Dim wb, sheet, empMeta, empName, empID
    Set wb = xl.Workbooks.Open(filePath)

    fout.WriteLine "Workbook: " & fso.GetFileName(filePath) & " | Sheets:"
    For Each sheet In wb.Sheets
        'fout.WriteLine "    " & sheet.Name

        empMeta = ParseEmpSheet(sheet.Name)
        If empMeta <> "" Then
            empName = Split(empMeta, "|")(0)
            empID = Split(empMeta, "|")(1)
            fout.WriteLine empName & " | ID: " & empID
        End If
    Next

    wb.Close True
    Set wb = Nothing
End Sub


Function ParseEmpSheet(sheetName)
    Dim trimmedName, dashPos, empName, empID

    trimmedName = Trim(sheetName)

    ' Skip known non-employee sheet patterns
    If InStr(1, trimmedName, "fsr", vbTextCompare) > 0 Or _
       InStr(1, trimmedName, "total", vbTextCompare) > 0 Then
        ParseEmpSheet = ""  ' Not an employee sheet
        Exit Function
    End If

    dashPos = InStrRev(trimmedName, "-")
    If dashPos > 0 Then
        empName = Trim(Left(trimmedName, dashPos - 1))
        empID = Trim(Mid(trimmedName, dashPos + 1))
    Else
        empName = trimmedName
        empID = ""
    End If

    ParseEmpSheet = empName & "|" & empID
End Function


Function Fetch_FSR_TotalAccountsAssisted(branch, empName, colDate)
    Dim conn, rs, sql
    Set conn = OpenConnection()
    If conn Is Nothing Then Exit Function
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3 ' Use client-side cursor

    sql = "SELECT ProductsSold, TotalPSProvided FROM dbo.FSR_TotalAccountsAssisted " & _
          "WHERE Branch = '" & SqlSafe(branch) & "' " & _
          "AND FSR_Name LIKE '%" & SqlSafe(empName) & "%' " & _
          "AND DateCol = '" & colDate & "'"

	
		rs.Open sql, conn, 1, 1

		' Check if recordset is valid and open
		If rs Is Nothing Or rs.State = 0 Then
			fout.WriteLine " SQL Fetch Failed for [" & empName & "] on [" & colDate & "]"
			fout.WriteLine "    Query: " & sql
			Set Fetch_FSR_TotalAccountsAssisted = Nothing
			Exit Function
		End If

		' Check if the recordset is empty
		If rs.EOF And rs.BOF Then
			'fout.WriteLine " No records found for [" & empName & "] on [" & colDate & "]"
			'fout.WriteLine "    Query: " & sql
			Set Fetch_FSR_TotalAccountsAssisted = Nothing
			Exit Function
		End If

    Set Fetch_FSR_TotalAccountsAssisted = rs
End Function





Sub WriteDataToSheet(ws, rs)
    Dim colCount, rowNum, colNum, i

    If rs.EOF Then
        fout.WriteLine "No records to write."
        Exit Sub
    End If

    rs.MoveFirst
    colCount = rs.Fields.Count
    rowNum = 1

    ' === Clear the sheet (optional) ===
    ws.Cells.ClearContents

    ' === Write headers ===
    For colNum = 1 To colCount
        ws.Cells(rowNum, colNum).Value = rs.Fields(colNum - 1).Name
    Next

    rowNum = rowNum + 1

    ' === Write records ===
    Do Until rs.EOF
        For colNum = 1 To colCount
            ws.Cells(rowNum, colNum).Value = rs.Fields(colNum - 1).Value
        Next
        rowNum = rowNum + 1
        rs.MoveNext
    Loop

    'fout.WriteLine "Wrote " & (rowNum - 2) & " rows to sheet: " & ws.Name
End Sub



Function GetYYYYMMDD(dateVal)
    GetYYYYMMDD = Year(dateVal) & _
                  Right("0" & Month(dateVal), 2) & _
                  Right("0" & Day(dateVal), 2)
End Function


Function SqlSafe(value)
    SqlSafe = Replace(value, "'", "''") ' Escape single quotes for SQL
End Function


Function GetLastDayOfPreviousMonth()
    Dim dt, colDate
    dt = DateSerial(Year(Date), Month(Date), 1) ' 1st of current month
    colDate = dt - 1 ' subtract 1 day = last day of previous month
    GetLastDayOfPreviousMonth = Year(colDate) & Right("0" & Month(colDate), 2) & Right("0" & Day(colDate), 2)
End Function


Function GetMonthAbbreviation(val)
    If IsDate(val) Then
        GetMonthAbbreviation = MonthName(Month(val), True)
    Else
        Dim mm
        mm = CInt(Mid(val, 5, 2))
        GetMonthAbbreviation = MonthName(mm, True)
    End If
End Function





Sub Test_FSR_Query(branch, empName, conn)
    Dim rs, sql, colDate

    colDate = GetLastDayOfPreviousMonth()

    sql = "SELECT ProductsSold, TotalPSProvided FROM dbo.FSR_TotalAccountsAssisted " & _
          "WHERE Branch = '" & SqlSafe(branch) & "' " & _
          "AND FSR_Name LIKE '%" & SqlSafe(empName) & "%' " & _
          "AND DateCol = '" & colDate & "'"

    fout.WriteLine "Test Query: " & sql

    On Error Resume Next
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 1, 1
    If Err.Number <> 0 Then
        fout.WriteLine "SQL Error: " & Err.Description
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    If rs.EOF Then
        fout.WriteLine "No results for: " & empName & " on " & colDate
    Else
        rs.MoveFirst
        fout.WriteLine "Query returned:"
        fout.WriteLine "   ProductsSold       = " & rs("ProductsSold")
        fout.WriteLine "   TotalPSProvided    = " & rs("TotalPSProvided")
    End If

    If Not rs Is Nothing Then rs.Close : Set rs = Nothing
End Sub

Sub WriteMonthlyTotalAccountsAssisted(ws, rsAssisted, mnth)
    Dim rowIndexAssist

    rowIndexAssist = GetMonthRowIndex(mnth, 7)
    If rowIndexAssist = -1 Then Exit Sub

    If rsAssisted Is Nothing Or (rsAssisted.EOF And rsAssisted.BOF) Then
        ' No data to write
        Exit Sub
    End If

    rsAssisted.MoveFirst

    ' Write ProductsSold to column B (2)
    On Error Resume Next
    ws.Cells(rowIndexAssist, 2).Value = Nz(rsAssisted("ProductsSold"), 0)
	lineReport = lineReport & PadRight(Nz(rsAssisted("ProductsSold"), 0), 12)

    ' Write TotalPSProvided to column D (4)
    ws.Cells(rowIndexAssist, 4).Value = Nz(rsAssisted("TotalPSProvided"), 0)
	lineReport = lineReport & PadRight(Nz(rsAssisted("TotalPSProvided"), 0), 12)

    On Error GoTo 0
End Sub



Function Fetch_FSR_DetailUserBranch(branch, empName, colDate)
    Dim conn, rs, sql
    Set conn = OpenConnection()
    If conn Is Nothing Then Exit Function
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3 ' Use client-side cursor

    sql = "SELECT AvgAssistTime, DaysWorked FROM dbo.FSR_DetailUserBranch " & _
          "WHERE Branch = '" & SqlSafe(branch) & "' " & _
          "AND EmpName LIKE '%" & SqlSafe(empName) & "%' " & _
          "AND DateCol = '" & colDate & "'"


    On Error Resume Next
    rs.Open sql, conn, 1, 1
    On Error GoTo 0
	
    ' Check if recordset is valid and open
    If rs Is Nothing Or rs.State = 0 Then
        fout.WriteLine " SQL Fetch Failed for [" & empName & "] on [" & colDate & "]"
        fout.WriteLine "    Query: " & sql
        Set Fetch_FSR_DetailUserBranch = Nothing
        Exit Function
    End If

    Set Fetch_FSR_DetailUserBranch = rs
End Function

Sub WriteMonthlyDetailUserBranch(ws, rsUserBranch, mnth)
    Dim rowIndex
    rowIndex = GetMonthRowIndex(mnth, 7)
    If rowIndex = -1 Then Exit Sub

    If rsUserBranch Is Nothing Or (rsUserBranch.EOF And rsUserBranch.BOF) Then
        Exit Sub
    End If

    rsUserBranch.MoveFirst

    On Error Resume Next
    ' AvgAssistTime → Column F (6)
    ws.Cells(rowIndex, 6).Value = Nz(rsUserBranch("AvgAssistTime"), 0)
	lineReport = lineReport & PadRight(Nz(rsUserBranch("AvgAssistTime"), 0),12)
    ' DaysWorked → Column H (8)
    ws.Cells(rowIndex, 8).Value = Nz(rsUserBranch("DaysWorked"), 0)
	lineReport = lineReport & PadRight(Nz(rsUserBranch("DaysWorked"), 0),12)
    On Error GoTo 0
End Sub




Function Nz(val, fallback)
    If IsNull(val) Then
        Nz = fallback
    Else
        Nz = val
    End If
End Function

Function GetMonthRowIndex(mnth, startRow)
    Dim monthMap, i
    monthMap = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", _
                     "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

    For i = 0 To 11
        If LCase(mnth) = LCase(monthMap(i)) Then
            GetMonthRowIndex = startRow + i
            Exit Function
        End If
    Next

    fout.WriteLine "Invalid month in GetMonthRowIndex: " & mnth
    GetMonthRowIndex = -1
End Function

Function Fetch_FSR_NewLoanReportbyBranch(branch, empID, colDate)
    Dim conn, rs, sql, thisbranch
    Set conn = OpenConnection()
    If conn Is Nothing Then Exit Function
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3 ' Use client-side cursor

    ' Normalize branch
	thisbranch = branch
    If LCase(Left(branch, 8)) = "longview" Then thisbranch = "Longview"

    ' Build SQL
    sql = "SELECT * FROM FSR_NewLoanReportbyBranch WHERE " & _
          "Branch LIKE '%" & Replace(thisbranch, "'", "''") & "%' AND " & _
          "EmpID = '" & Replace(empID, "'", "''") & "' AND " & _
          "DateCol = '" & colDate & "'"

   
    On Error Resume Next
    rs.Open sql, conn, 1, 1
    On Error GoTo 0
	
    ' Check if recordset is valid and open
    If rs Is Nothing Or rs.State = 0 Then
        fout.WriteLine " SQL Fetch Failed for [" & empId & "] on [" & colDate & "]"
        fout.WriteLine "    Query: " & sql
        Set Fetch_FSR_NewLoanReportbyBranch = Nothing
        Exit Function
    End If

    Set Fetch_FSR_NewLoanReportbyBranch = rs
End Function




Sub WriteMonthlyNewLoanReportbyBranch(ws, rs, mnth)
    Dim rowBlock1, rowBlock2

    rowBlock1 = GetMonthRowIndex(mnth, 7)
    If rowBlock1 = -1 Then Exit Sub

    If rs Is Nothing Or (rs.EOF And rs.BOF) Then Exit Sub

    rs.MoveFirst

    On Error Resume Next

	   ' === Block 1: Starting at rowBlock1 ===
	ws.Cells(rowBlock1, 10).Value = Nz(rs("TotalBookedLoans"), 0)       ' Column J
	lineReport = lineReport & PadRight(Nz(rs("TotalBookedLoans"), 0), 12)

	ws.Cells(rowBlock1, 11).Value = Nz(rs("SumTotalBookedLoans"), 0)    ' Column K
	lineReport = lineReport & PadRight(Nz(rs("SumTotalBookedLoans"), 0), 12)

	ws.Cells(rowBlock1, 13).Value = Nz(rs("TotalLoanRecap"), 0)         ' Column M
	lineReport = lineReport & PadRight(Nz(rs("TotalLoanRecap"), 0), 12)

	ws.Cells(rowBlock1, 15).Value = Nz(rs("SumTotalLoanRecap"), 0)      ' Column O
	lineReport = lineReport & PadRight(Nz(rs("SumTotalLoanRecap"), 0), 12)


	' === Block 2: Starting at rowBlock2 ===
	rowBlock2 = GetMonthRowIndex(mnth, 25)
	If rowBlock2 = -1 Then Exit Sub

	ws.Cells(rowBlock2, 2).Value = Nz(rs("CDLoansEligible"), 0)         ' Column B
	lineReport = lineReport & PadRight(Nz(rs("CDLoansEligible"), 0), 12)

	ws.Cells(rowBlock2, 3).Value = Nz(rs("SumCDLoansEligible"), 0)      ' Column C
	lineReport = lineReport & PadRight(Nz(rs("SumCDLoansEligible"), 0), 12)

	ws.Cells(rowBlock2, 4).Value = Nz(rs("CDLoansProtected"), 0)        ' Column D
	lineReport = lineReport & PadRight(Nz(rs("CDLoansProtected"), 0), 12)

	ws.Cells(rowBlock2, 5).Value = Nz(rs("SumCDLoansProtected"), 0)     ' Column E
	lineReport = lineReport & PadRight(Nz(rs("SumCDLoansProtected"), 0), 12)

	ws.Cells(rowBlock2, 9).Value = Nz(rs("CLLoansEligible"), 0)         ' Column I
	lineReport = lineReport & PadRight(Nz(rs("CLLoansEligible"), 0), 12)

	ws.Cells(rowBlock2, 10).Value = Nz(rs("SumCLLoansEligible"), 0)     ' Column J
	lineReport = lineReport & PadRight(Nz(rs("SumCLLoansEligible"), 0), 12)

	ws.Cells(rowBlock2, 11).Value = Nz(rs("CLLoansProtected"), 0)       ' Column K
	lineReport = lineReport & PadRight(Nz(rs("CLLoansProtected"), 0), 12)

	ws.Cells(rowBlock2, 12).Value = Nz(rs("SumCLLoansProtected"), 0)    ' Column L
	lineReport = lineReport & PadRight(Nz(rs("SumCLLoansProtected"), 0), 12)

	ws.Cells(rowBlock2, 16).Value = Nz(rs("MMP"), 0)                    ' Column P
	lineReport = lineReport & PadRight(Nz(rs("MMP"), 0), 12)

	ws.Cells(rowBlock2, 18).Value = Nz(rs("GAP"), 0)                    ' Column R
	lineReport = lineReport & PadRight(Nz(rs("GAP"), 0), 12)

    On Error GoTo 0
End Sub



Function PadRight(val, width)
    PadRight = Left(CStr(val) & Space(width), width)
End Function


Function Fetch_FSR_SharesOpenedbyBranch(branch, EmpID, colDate)
    Dim conn, rs, sql, thisbranch
    Set conn = OpenConnection()
    If conn Is Nothing Then Exit Function

    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3 ' Use client-side cursor
	
    ' Normalize branch
	thisbranch = branch
    If LCase(Left(branch, 8)) = "longview" Then thisbranch = "Longview"

    sql = "SELECT NewSavings FROM dbo.FSR_SharesOpenedbyBranch " & _
          "WHERE Branch = '" & SqlSafe(thisbranch) & "' " & _
          "AND EmpID = '" & SqlSafe(EmpID) & "' " & _
          "AND DateCol = '" & colDate & "'"

    On Error Resume Next
    rs.Open sql, conn, 1, 1
    On Error GoTo 0

    If rs Is Nothing Or rs.State = 0 Then
        fout.WriteLine " SQL Fetch Failed for EmpID [" & EmpID & "] on [" & colDate & "]"
        fout.WriteLine "    Query: " & sql
        Set Fetch_FSR_SharesOpenedbyBranch = Nothing
        Exit Function
    End If

    If rs.EOF And rs.BOF Then
        'fout.WriteLine " No records found for EmpID [" & EmpID & "] on [" & colDate & "]"
        'fout.WriteLine "    Query: " & sql
        Set Fetch_FSR_SharesOpenedbyBranch = Nothing
        Exit Function
    End If

    Set Fetch_FSR_SharesOpenedbyBranch = rs
End Function



Sub WriteMonthlySharesOpenedbyBranch(ws, rsShares, mnth)
    Dim rowIndex

    rowIndex = GetMonthRowIndex(mnth, 7)
    If rowIndex = -1 Then Exit Sub

    If rsShares Is Nothing Or (rsShares.EOF And rsShares.BOF) Then Exit Sub

    rsShares.MoveFirst

    ' Write NewSavings to column Q (17)
    On Error Resume Next
    ws.Cells(rowIndex, 17).Value = Nz(rsShares("NewSavings"), 0)
	lineReport = lineReport & PadRight(Nz(rsShares("NewSavings"), 0), 12)
	
    On Error GoTo 0
End Sub

Function Fetch_FSR_ShareDFTSOpenedbyBranch(branch, EmpID, colDate)
    Dim conn, rs, sql, thisbranch
    Set conn = OpenConnection()
    If conn Is Nothing Then Exit Function

    ' Normalize branch
	thisbranch = branch
    If LCase(Left(branch, 8)) = "longview" Then thisbranch = "Longview"

    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3 ' Use client-side cursor

    sql = "SELECT NewDrafts FROM dbo.FSR_ShareDFTSOpenedbyBranch " & _
          "WHERE Branch = '" & SqlSafe(thisbranch) & "' " & _
          "AND EmpID = '" & SqlSafe(EmpID) & "' " & _
          "AND DateCol = '" & colDate & "'"

    On Error Resume Next
    rs.Open sql, conn, 1, 1
    On Error GoTo 0

    If rs Is Nothing Or rs.State = 0 Then
        fout.WriteLine " SQL Fetch Failed for EmpID [" & EmpID & "] on [" & colDate & "]"
        fout.WriteLine "    Query: " & sql
        Set Fetch_FSR_ShareDFTSOpenedbyBranch = Nothing
        Exit Function
    End If

    If rs.EOF And rs.BOF Then
        'fout.WriteLine " No records found for EmpID [" & EmpID & "] on [" & colDate & "]"
        'fout.WriteLine "    Query: " & sql
        Set Fetch_FSR_ShareDFTSOpenedbyBranch = Nothing
        Exit Function
    End If

    Set Fetch_FSR_ShareDFTSOpenedbyBranch = rs
End Function



Sub WriteMonthlyShareDFTSOpenedbyBranch(ws, rsDrafts, mnth)
    Dim rowIndex
    rowIndex = GetMonthRowIndex(mnth, 7)
    If rowIndex = -1 Then Exit Sub

    If rsDrafts Is Nothing Or (rsDrafts.EOF And rsDrafts.BOF) Then Exit Sub

    rsDrafts.MoveFirst

    ' Write NewDrafts to column S (19)
    On Error Resume Next
	ws.Cells(rowIndex, 19).Value = Nz(rsDrafts("NewDrafts"), 0)          ' Column S
	lineReport = lineReport & PadRight(Nz(rsDrafts("NewDrafts"), 0), 12)

    On Error GoTo 0
End Sub

Function Fetch_FSR_TimeDetailReport(branch, empName, colDate)
    Dim conn, rs, sql, thisbranch
    Set conn = OpenConnection()
    If conn Is Nothing Then Exit Function

    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3 ' Use client-side cursor
	
    ' Normalize branch
	thisbranch = branch
    If LCase(Left(branch, 8)) = "longview" Then thisbranch = "Longview"

    sql = "SELECT TotalHours FROM dbo.FSR_TimeDetailReport " & _
          "WHERE Branch = '" & SqlSafe(thisbranch) & "' " & _
          "AND EmpName LIKE '%" & SqlSafe(empName) & "%' " & _
          "AND DateCol = '" & colDate & "'"

    On Error Resume Next
    rs.Open sql, conn, 1, 1
    On Error GoTo 0

    If rs Is Nothing Or rs.State = 0 Then
        fout.WriteLine " SQL Fetch Failed for [" & empName & "] on [" & colDate & "]"
        fout.WriteLine "    Query: " & sql
        Set Fetch_FSR_TimeDetailReport = Nothing
        Exit Function
    End If

    If rs.EOF And rs.BOF Then
        'fout.WriteLine " No records found for [" & empName & "] on [" & colDate & "]"
        'fout.WriteLine "    Query: " & sql
        Set Fetch_FSR_TimeDetailReport = Nothing
        Exit Function
    End If

    Set Fetch_FSR_TimeDetailReport = rs
End Function

Sub WriteMonthlyTimeDetailReport(ws, rsTime, mnth)
    Dim rowIndex
    rowIndex = GetMonthRowIndex(mnth, 43)
    If rowIndex = -1 Then Exit Sub

    If rsTime Is Nothing Or (rsTime.EOF And rsTime.BOF) Then Exit Sub

    rsTime.MoveFirst

    ' Write TotalHours to column B (2)
    On Error Resume Next
    ws.Cells(rowIndex, 2).Value = Nz(rsTime("TotalHours"), 0)
	lineReport = lineReport & PadRight(Nz(rsTime("TotalHours"), 0), 12)

    On Error GoTo 0
End Sub


Function Fetch_FSR_AcdCalls(branch, empName, colDate)
    Dim conn, rs, sql, thisbranch
	
    Set conn = OpenConnection()
    If conn Is Nothing Then Exit Function
	
    ' Normalize branch
	thisbranch = branch
    If LCase(Left(branch, 8)) = "longview" Then thisbranch = "Longview"
	
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3 ' Use client-side cursor

    sql = "SELECT Answered FROM dbo.FSR_AcdCalls " & _
          "WHERE FSR_Name LIKE '%" & SqlSafe(empName) & "%' " & _
          "AND DateCol = '" & colDate & "'"

    On Error Resume Next
    rs.Open sql, conn, 1, 1
    On Error GoTo 0

    If rs Is Nothing Or rs.State = 0 Then
        fout.WriteLine " SQL Fetch Failed for [" & empName & "] on [" & colDate & "]"
        fout.WriteLine "    Query: " & sql
        Set Fetch_FSR_AcdCalls = Nothing
        Exit Function
    End If

    If rs.EOF And rs.BOF Then
        'fout.WriteLine " No records found for [" & empName & "] on [" & colDate & "]"
        'fout.WriteLine "    Query: " & sql
        Set Fetch_FSR_AcdCalls = Nothing
        Exit Function
    End If

    Set Fetch_FSR_AcdCalls = rs
End Function

Sub WriteMonthlyAcdCalls(ws, rsAcd, mnth)
    Dim rowIndex
    rowIndex = GetMonthRowIndex(mnth, 43)
    If rowIndex = -1 Then Exit Sub

    If rsAcd Is Nothing Or (rsAcd.EOF And rsAcd.BOF) Then Exit Sub

    rsAcd.MoveFirst

    ' Write Answered count to column D (4)
    On Error Resume Next
	ws.Cells(rowIndex, 4).Value = Nz(rsAcd("Answered"), 0)               ' Column D
	lineReport = lineReport & PadRight(Nz(rsAcd("Answered"), 0), 12)

    On Error GoTo 0
End Sub


Function Fetch_FSR_BranchMaintenance(branch, empName, colDate)
    Dim conn, rs, sql, thisbranch
    Set conn = OpenConnection()
    If conn Is Nothing Then Exit Function

    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3 ' Use client-side cursor

    ' Normalize branch
	thisbranch = branch
    If LCase(Left(branch, 8)) = "longview" Then thisbranch = "Longview"

    sql = "SELECT TotalEntries FROM dbo.FSR_BranchMaintenance " & _
          "WHERE Branch = '" & SqlSafe(thisbranch) & "' " & _
          "AND empName LIKE '%" & SqlSafe(empName) & "%' " & _
          "AND DateCol = '" & colDate & "'"


    On Error Resume Next
    rs.Open sql, conn, 1, 1
    On Error GoTo 0

    If rs Is Nothing Or rs.State = 0 Then
        fout.WriteLine " SQL Fetch Failed for [" & empName & "] on [" & colDate & "]"
        fout.WriteLine "    Query: " & sql
        Set Fetch_FSR_BranchMaintenance = Nothing
        Exit Function
    End If



    If rs.EOF And rs.BOF Then
        'fout.WriteLine " No records found for [" & empName & "] on [" & colDate & "]"
		
		
        'fout.WriteLine "    Query: " & sql
        Set Fetch_FSR_BranchMaintenance = Nothing
        Exit Function
    End If

    Set Fetch_FSR_BranchMaintenance = rs
End Function

Sub WriteMonthlyBranchMaintenance(ws, rsBranch, mnth)
    Dim rowIndex
    rowIndex = GetMonthRowIndex(mnth, 43)
    If rowIndex = -1 Then Exit Sub

    If rsBranch Is Nothing Or (rsBranch.EOF And rsBranch.BOF) Then Exit Sub

    rsBranch.MoveFirst

    ' Write TotalEntries to column E (5)
    On Error Resume Next
	ws.Cells(rowIndex, 5).Value = Nz(rsBranch("TotalEntries"), 0)        ' Column E
	lineReport = lineReport & PadRight(Nz(rsBranch("TotalEntries"), 0), 12)

    On Error GoTo 0
End Sub


Function Fetch_FSR_TellerActivity(branch, empID, colDate)
    Dim conn, rs, sql, thisbranch
	
	
    Set conn = OpenConnection()
    If conn Is Nothing Then Exit Function

    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3 ' Use client-side cursor

    ' Normalize branch
	thisbranch = branch
    If LCase(Left(branch, 8)) = "longview" Then thisbranch = "Longview"
	
    sql = "SELECT Total FROM dbo.FSR_TellerActivity " & _
          "WHERE EmpID = '" & SqlSafe(empID) & "' " & _
          "AND DateCol = '" & colDate & "'"

    On Error Resume Next
    rs.Open sql, conn, 1, 1
    On Error GoTo 0

    If rs Is Nothing Or rs.State = 0 Then
        fout.WriteLine " SQL Fetch Failed for [" & empID & "] on [" & colDate & "]"
        fout.WriteLine "    Query: " & sql
        Set Fetch_FSR_TellerActivity = Nothing
        Exit Function
    End If

    If rs.EOF And rs.BOF Then
        ' No records found
        Set Fetch_FSR_TellerActivity = Nothing
        Exit Function
    End If

    Set Fetch_FSR_TellerActivity = rs
End Function


Sub WriteMonthlyTellerActivity(ws, rsTeller, mnth)
    Dim rowIndex
    rowIndex = GetMonthRowIndex(mnth, 43)
    If rowIndex = -1 Then Exit Sub

    If rsTeller Is Nothing Or (rsTeller.EOF And rsTeller.BOF) Then Exit Sub

    rsTeller.MoveFirst

    ' Write Total to column C (3)
    On Error Resume Next
	ws.Cells(rowIndex, 3).Value = Nz(rsTeller("Total"), 0)               ' Column C
	lineReport = lineReport & PadRight(Nz(rsTeller("Total"), 0), 12)

    On Error GoTo 0
End Sub


Sub SaveWorkbookWithVersion(wb, filePath, mnth)
    If wb Is Nothing Then
        Call LogError("Cannot save workbook: wb is Nothing")
        Exit Sub
    End If

    Dim fso, reportBaseFolder, baseName, ext, nameOnly, saveCopyPath, version
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    reportBaseFolder = "S:\Reports\"

    If Not fso.FolderExists(reportBaseFolder) Then
        fso.CreateFolder(reportBaseFolder)
    End If

    baseName = fso.GetFileName(filePath)
    ext = fso.GetExtensionName(baseName)
    nameOnly = Left(baseName, Len(baseName) - Len(ext) - 1) & "_" & mnth
    version = 0

    Do
        If version = 0 Then
            saveCopyPath = reportBaseFolder & nameOnly & ".xlsx"
        Else
            saveCopyPath = reportBaseFolder & nameOnly & "~" & version & ".xlsx"
        End If
        version = version + 1
    Loop While fso.FileExists(saveCopyPath)

    On Error Resume Next
    wb.SaveAs saveCopyPath
    If Err.Number <> 0 Then
        Call LogError("Failed to save updated copy: " & Err.Description)
        Err.Clear
    End If
    On Error GoTo 0
End Sub
