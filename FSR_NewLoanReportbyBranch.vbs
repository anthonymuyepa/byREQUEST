Option Explicit

Dim folderPath, connString, tablename
	folderPath = "R:\Branch Manager Hillary Reports\"  ' Root folder for all branches
	tablename = "FSR_NewLoanReportbyBranch"
	connString = "Provider=SQLOLEDB;" & _
				 "Data Source=GOLDMINE-SQL1;" & _
				 "Initial Catalog=byREQUEST;" & _
				 "Integrated Security=SSPI;"



Sub StartDoc()
    Dim fso, xl, folder, file, conn, wb, ws, lastRow
    Dim stats, fileDate, folderMonth, DateCol, branchName, branchFolder
    Dim skipFile

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    Set conn = OpenConnection()
    If conn Is Nothing Then Exit Sub

    Call CreateTableIfNeeded(conn)
	

	' === Log header ===
	fout.WriteLine PadRight("DateCol", 10) & _
				   PadRight("Month", 10) & _
				   PadRight("Branch", 24) & _
				   PadRight("EmpID", 10) & _
				   PadRight("Booked", 8) & _
				   PadRight("BookedSum", 14) & _
				   PadRight("Recap", 8) & _
				   PadRight("RecapSum", 14) & _
				   PadRight("CDLE", 8) & _
				   PadRight("SumCDLE", 14) & _
				   PadRight("CDLP", 8) & _
				   PadRight("SumCDLP", 14) & _
				   PadRight("CLLE", 8) & _
				   PadRight("SumCLLE", 14) & _
				   PadRight("CLLEP", 8) & _
				   PadRight("SumCLLEP", 14) & _
				   PadRight("MMP", 6) & _
				   PadRight("GAP", 6)




    For Each branchFolder In fso.GetFolder(folderPath).SubFolders

        branchName = branchFolder.Name
        Dim subFolderPath
        subFolderPath = branchFolder.Path & "\Hillary Reports\"

        On Error Resume Next
        If fso.FolderExists(subFolderPath) Then
            On Error GoTo 0
            For Each file In fso.GetFolder(subFolderPath).Files
                skipFile = False
 
				If InStr(LCase(file.Name), "new_loan_report_by_branch-") > 0 And _
				   Right(LCase(file.Name), 5) = ".xlsx" And _
				   InStr(file.Name, "~") = 0 Then
                    ' Extract fileDate from filename
					'msgbox "Call ExtractDate: "& file.name
                    fileDate = ExtractFileDate(file.Name)
	
'MsgBox "File Date: " & fileDate
                    If Not skipFile Then
                        DateCol = fileDate
						
                        folderMonth = MonthName(Month(ParseYYYYMMDD(fileDate)), False)

                        ' Open Excel + process employee metrics
                        Set wb = xl.Workbooks.Open(file.Path)
                        Set ws = wb.Sheets(1)
                        lastRow = ws.Cells(ws.Rows.Count, "E").End(-4162).Row
                        Set stats = Summarize(ws, 5, lastRow)
						'msgbox "Call InsertStats"
					
						' Remove trailing " Branch" if it exists (case-insensitive)
						If LCase(Right(branchName, 7)) = " branch" Then
							branchName = Left(branchName, Len(branchName) - 7)
						End If
						branchName = Trim(branchName)

                        Call InsertStats(conn, stats, DateCol, folderMonth, branchName)
						'msgbox "After InsertStats"
                        wb.Close False
                        Set ws = Nothing
                        Set wb = Nothing
                    End If
                End If
            Next
        Else
            If Err.Number <> 0 Then
                LogError "Error checking folder: " & subFolderPath & " | " & Err.Description & " (Err #" & Err.Number & ")"
                Err.Clear
            Else
                LogInfo "Folder not found or inaccessible: " & subFolderPath
            End If
            On Error GoTo 0
        End If
    Next


    xl.Quit
    conn.Close
    Set xl = Nothing
    Set conn = Nothing
    Set fso = Nothing
End Sub


Function ExtractFileDate(fileName)
    Dim baseName, parts, strDate, dt, lastDay
    On Error Resume Next

    ' Remove extension
    baseName = Left(fileName, InStrRev(fileName, ".") - 1)

    ' Split by space and get last part
    parts = Split(baseName, " ")
    If UBound(parts) >= 0 Then
        strDate = Trim(parts(UBound(parts)))
        If IsDate(strDate) Then
            dt = CDate(strDate)

            ' Calculate last day of the previous month
            lastDay = DateSerial(Year(dt), Month(dt), 0)

            ' Format as YYYYMMDD
            ExtractFileDate = Year(lastDay) & _
                              Right("0" & Month(lastDay), 2) & _
                              Right("0" & Day(lastDay), 2)
        Else
            ExtractFileDate = ""
        End If
    Else
        ExtractFileDate = ""
    End If

    On Error GoTo 0
End Function


Function ParseYYYYMMDD(dateStr)
    Dim yyyy, mm, dd
    yyyy = Left(dateStr, 4)
    mm = Mid(dateStr, 5, 2)
    dd = Right(dateStr, 2)
    ParseYYYYMMDD = DateSerial(yyyy, mm, dd)
End Function


Function FormatDateForSQL(d)
    FormatDateForSQL = Year(d) & Right("0" & Month(DateAdd("m", 1, d)), 2) & _
                       Right("0" & Day(DateSerial(Year(d), Month(d)+1, 0)), 2)
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
    sql = "IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '" & tablename & "') BEGIN CREATE TABLE " & tablename & _
          "(DateCol NVARCHAR(8), Month NVARCHAR(20), Branch NVARCHAR(100), EmpID NVARCHAR(50)," & _
          "TotalBookedLoans NVARCHAR(50), SumTotalBookedLoans NVARCHAR(50), TotalLoanRecap NVARCHAR(50), SumTotalLoanRecap NVARCHAR(50)," & _
          "CDLoansEligible NVARCHAR(50), SumCDLoansEligible NVARCHAR(50), CDLoansProtected NVARCHAR(50), SumCDLoansProtected NVARCHAR(50)," & _
          "CLLoansEligible NVARCHAR(50), SumCLLoansEligible NVARCHAR(50), CLLoansProtected NVARCHAR(50), SumCLLoansProtected NVARCHAR(50)," & _
          "MMP NVARCHAR(50), GAP NVARCHAR(50)) END"
    On Error Resume Next
	c.Execute sql
    If Err.Number <> 0 Then LogError "CreateTable failed: " & Err.Description
    On Error GoTo 0
End Sub



Function Summarize(ws, startRow, endRow)
    Dim dict: Set dict = CreateObject("Scripting.Dictionary")
    Dim r, empID, ins, purp, bal, mmp, gap

    For r = startRow To endRow
        empID = CleanString(ws.Cells(r, "E").Value)
        ins = CleanString(ws.Cells(r, "F").Value)
        purp = CleanString(ws.Cells(r, "J").Value)
        bal = CleanNumber(ws.Cells(r, "K").Value)
        mmp = CleanNumber(ws.Cells(r, "J").Value)  ' MMP is in column J
        gap = CleanNumber(ws.Cells(r, "H").Value)

        If IsEmpty(empID) Or empID = "" Or InStr(empID, "-----") > 0 Then
        Else
            empID = Trim(CStr(empID))
            If Not dict.Exists(empID) Then
                Dim s: Set s = CreateObject("Scripting.Dictionary")
                s.Add "TotalBookedLoans", 0
                s.Add "SumTotalBookedLoans", 0
                s.Add "TotalLoanRecap", 0
                s.Add "SumTotalLoanRecap", 0
                s.Add "CDLoansEligible", 0
                s.Add "SumCDLoansEligible", 0
                s.Add "CDLoansProtected", 0
                s.Add "SumCDLoansProtected", 0
                s.Add "CLLoansEligible", 0
                s.Add "SumCLLoansEligible", 0
                s.Add "CLLoansProtected", 0
                s.Add "SumCLLoansProtected", 0
                s.Add "MMP", 0
                s.Add "GAP", 0
                dict.Add empID, s
            End If

            Set s = dict(empID)

            ' Booked
            s("TotalBookedLoans") = s("TotalBookedLoans") + 1
            If IsNumeric(bal) Then s("SumTotalBookedLoans") = s("SumTotalBookedLoans") + bal

            ' Recap
            If purp = "770" Then s("TotalLoanRecap") = s("TotalLoanRecap") + 1: If IsNumeric(bal) Then s("SumTotalLoanRecap") = s("SumTotalLoanRecap") + bal

            ' CD Eligibility (Credit Disability + Life)
            If ins <> "0019" Then
                s("CDLoansEligible") = s("CDLoansEligible") + 1: If IsNumeric(bal) Then s("SumCDLoansEligible") = s("SumCDLoansEligible") + bal
                s("CLLoansEligible") = s("CLLoansEligible") + 1: If IsNumeric(bal) Then s("SumCLLoansEligible") = s("SumCLLoansEligible") + bal
            End If

            ' CD Protected (Disability)
            If ins <> "003" And ins <> "005" And ins <> "000" Then
                s("CDLoansProtected") = s("CDLoansProtected") + 1
                If IsNumeric(bal) Then s("SumCDLoansProtected") = s("SumCDLoansProtected") + bal
            End If

            ' CL Protected (Life)
            If ins <> "001" And ins <> "008" And ins <> "000" Then
                s("CLLoansProtected") = s("CLLoansProtected") + 1
                If IsNumeric(bal) Then s("SumCLLoansProtected") = s("SumCLLoansProtected") + bal
            End If

            If IsNumeric(mmp) And mmp > 0 Then s("MMP") = s("MMP") + 1
            If IsNumeric(gap) And gap > 0 Then s("GAP") = s("GAP") + 1
        End If
    Next

    Set Summarize = dict
End Function


Function CleanString(val)
    Dim s
    s = Trim(CStr(val))
    If Left(s, 1) = "'" Then s = Mid(s, 2)
    
    ' If value contains hyphens in a row, treat as empty
    If InStr(s, "----") > 0 Then
        s = ""
    End If

    CleanString = s
End Function


Function CleanNumber(val)
    Dim s
    s = Trim(CStr(val))
    If Left(s, 1) = "'" Then s = Trim(Mid(s, 2))
    
    ' Treat as 0 if contains "-----"
    If InStr(s, "----") > 0 Then
        CleanNumber = 0
    ElseIf IsNumeric(s) Then
        CleanNumber = CDbl(s)
    Else
        CleanNumber = 0
    End If
End Function


Sub InsertStats(c, statsDict, dateVal, moName, br)
    Dim empID, s, sql, checkSQL, rs

    For Each empID In statsDict.Keys
        Set s = statsDict(empID)

        fout.WriteLine PadRight(dateVal, 10) & _
               PadRight(moName, 10) & _
               PadRight(br, 24) & _
               PadRight(empID, 10) & _
               PadRight(Nz(s("TotalBookedLoans"), "0"), 8) & _
               PadRight(Nz(s("SumTotalBookedLoans"), "0"), 14) & _
               PadRight(Nz(s("TotalLoanRecap"), "0"), 8) & _
               PadRight(Nz(s("SumTotalLoanRecap"), "0"), 14) & _
               PadRight(Nz(s("CDLoansEligible"), "0"), 8) & _
               PadRight(Nz(s("SumCDLoansEligible"), "0"), 14) & _
               PadRight(Nz(s("CDLoansProtected"), "0"), 8) & _
               PadRight(Nz(s("SumCDLoansProtected"), "0"), 14) & _
               PadRight(Nz(s("CLLoansEligible"), "0"), 8) & _
               PadRight(Nz(s("SumCLLoansEligible"), "0"), 14) & _
               PadRight(Nz(s("CLLoansProtected"), "0"), 8) & _
               PadRight(Nz(s("SumCLLoansProtected"), "0"), 14) & _
               PadRight(Nz(s("MMP"), "0"), 6) & _
               PadRight(Nz(s("GAP"), "0"), 6)


        checkSQL = "SELECT 1 FROM " & tablename & _
                   " WHERE DateCol = '" & dateVal & "'" & _
                   " AND Month = '" & Replace(moName, "'", "''") & "'" & _
                   " AND Branch = '" & Replace(br, "'", "''") & "'" & _
                   " AND EmpID = '" & Replace(empID, "'", "''") & "'"

        On Error Resume Next
        Set rs = c.Execute(checkSQL)
        If Err.Number <> 0 Then
            LogError "Check error (" & empID & "): " & Err.Description
            Err.Clear
        ElseIf rs.EOF Then
            sql = "INSERT INTO " & tablename & " (DateCol, Month, Branch, EmpID, " & _
				  "TotalBookedLoans, SumTotalBookedLoans, TotalLoanRecap, SumTotalLoanRecap, " & _
				  "CDLoansEligible, SumCDLoansEligible, CDLoansProtected, SumCDLoansProtected, " & _
				  "CLLoansEligible, SumCLLoansEligible, CLLoansProtected, SumCLLoansProtected, MMP, GAP) VALUES (" & _
				  Quote(dateVal) & ", " & Quote(moName) & ", " & Quote(br) & ", " & Quote(empID) & ", " & _
				  Quote(Nz(s("TotalBookedLoans"), "0")) & ", " & Quote(Nz(s("SumTotalBookedLoans"), "0")) & ", " & _
				  Quote(Nz(s("TotalLoanRecap"), "0")) & ", " & Quote(Nz(s("SumTotalLoanRecap"), "0")) & ", " & _
				  Quote(Nz(s("CDLoansEligible"), "0")) & ", " & Quote(Nz(s("SumCDLoansEligible"), "0")) & ", " & _
				  Quote(Nz(s("CDLoansProtected"), "0")) & ", " & Quote(Nz(s("SumCDLoansProtected"), "0")) & ", " & _
				  Quote(Nz(s("CLLoansEligible"), "0")) & ", " & Quote(Nz(s("SumCLLoansEligible"), "0")) & ", " & _
				  Quote(Nz(s("CLLoansProtected"), "0")) & ", " & Quote(Nz(s("SumCLLoansProtected"), "0")) & ", " & _
				  Quote(Nz(s("MMP"), "0")) & ", " & Quote(Nz(s("GAP"), "0")) & ")"


			'fout.WriteLine sql
			
            c.Execute sql
            If Err.Number <> 0 Then LogError "Insert error (" & empID & "): " & Err.Description
        End If
        On Error GoTo 0
    Next
End Sub

Function Nz(val, defaultVal)
    If IsNull(val) Or Trim(val & "") = "" Then
        Nz = defaultVal
    Else
        Nz = val
    End If
End Function




Function PadRight(val, width)
    If IsNull(val) Then val = ""
    val = CStr(val)
    PadRight = Left(val & Space(width), width)
End Function


Sub LogError(msg)
    fout.WriteLine Now & " | ERROR: " & msg
End Sub


Function Quote(val)
    If IsNull(val) Or IsEmpty(val) Or Trim(CStr(val)) = "" Then
        Quote = "''"  ' or "'0'" or "'NULL'" if you want those instead
    Else
        Quote = "'" & Replace(Trim(CStr(val)), "'", "''") & "'"
    End If
End Function
