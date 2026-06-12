Option Explicit

Dim folderPath, connString, tablename
folderPath = "R:\Paycom"
tablename = "FSR_TimeDetailReport"
connString = "Provider=SQLOLEDB;" & _
             "Data Source=GOLDMINE-SQL1;" & _
             "Initial Catalog=byREQUEST;" & _
             "Integrated Security=SSPI;"

Sub StartDoc()
    Dim fso, xl, file, conn, wb, ws, lastRow
    Dim stats, fileDate, folderMonth, DateCol, branchName, monthNameStr, yearStr
    Dim skipFile

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    Set conn = OpenConnection()
    If conn Is Nothing Then Exit Sub

    Call CreateTableIfNeeded(conn)

	    ' === Write Header ===
    fout.WriteLine PadRight("DateCol", 10) & _
                   PadRight("Month", 10) & _
                   PadRight("Branch", 24) & _
                   PadRight("EmpID", 10) & _
                   PadRight("EmpName", 24) & _
                   PadRight("TotalHours", 10)
				   
				   
				   
    For Each file In fso.GetFolder(folderPath).Files
        skipFile = False

		Dim s: s = LCase(file.Name)

		If (InStr(s, "fsr sick") > 0 _
			Or InStr(s, "time detail") > 0) _
		   And (Right(s, 5) = ".xlsx") _
		   And InStr(s, "~") = 0 Then

            monthNameStr = ExtractMonthFromFilename(file.Name)
            yearStr = ExtractYearFromFilename(file.Name)
            DateCol = GetMonthEndDate(monthNameStr, yearStr)
            folderMonth = monthNameStr

			Set wb = xl.Workbooks.Open(file.Path)

			Set ws = GetTotalHoursSummarySheet(wb)
			If ws Is Nothing Then
				LogError "Sheet 'Total Hours Summary' not found in: " & file.Name
			Else
				lastRow = ws.Cells(ws.Rows.Count, "A").End(-4162).Row  ' xlUp = -4162
				Set stats = SummarizeTimeDetail(ws, 5, lastRow)
				branchName = ExtractBranchName(file.Name)
				Call InsertStats(conn, stats, DateCol, folderMonth, branchName)
			End If

            wb.Close False
            Set ws = Nothing
            Set wb = Nothing
        End If
    Next

    xl.Quit
    conn.Close
    Set xl = Nothing
    Set conn = Nothing
    Set fso = Nothing
End Sub

Function ExtractMonthFromFilename(fileName)
    Dim words: words = Split(fileName, " ")
    ExtractMonthFromFilename = words(0)
End Function

Function ExtractYearFromFilename(fileName)
    Dim words: words = Split(fileName, " ")
    ExtractYearFromFilename = words(1)
End Function

Function ExtractBranchName(fileName)
    Dim parts: parts = Split(fileName, " - ")
    If UBound(parts) >= 1 Then
        ExtractBranchName = Replace(Replace(parts(1), ".xlsx", ""), vbCrLf, "")
    Else
        ExtractBranchName = "Unknown"
    End If
End Function

Function GetMonthEndDate(monthName, yearVal)
    Dim m: m = Month(DateValue("1-" & monthName & "-" & yearVal))
    GetMonthEndDate = yearVal & Right("0" & m, 2) & Right("0" & Day(DateSerial(yearVal, m + 1, 0)), 2)
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
          "TotalHours NVARCHAR(50)) END"
    On Error Resume Next
    c.Execute sql
    If Err.Number <> 0 Then LogError "CreateTable failed: " & Err.Description
    On Error GoTo 0
End Sub

Function SummarizeTimeDetail(ws, startRow, endRow)
    Dim dict: Set dict = CreateObject("Scripting.Dictionary")
    Dim colEmpID, colLast, colFirst, colHours, colHomeDesc, colDistDesc, colBranch
    Dim r, empID, firstName, lastName, empName, branch, hours, s

    ' Resolve columns from row 1 headers
    colEmpID    = FindHeaderCol(ws, "EECode")
    colLast     = FindHeaderCol(ws, "Lastname")
    colFirst    = FindHeaderCol(ws, "Firstname")
    colHours    = FindHeaderCol(ws, "EarnHours")
    colHomeDesc = FindHeaderCol(ws, "Home Department Desc")
    colDistDesc = FindHeaderCol(ws, "Dist Department Desc")

    ' Prefer Home Department Desc, else Dist Department Desc
    If colHomeDesc <> 0 Then
        colBranch = colHomeDesc
    Else
        colBranch = colDistDesc
    End If

    ' If critical headers are missing, bail gracefully
    If colEmpID = 0 Or colHours = 0 Then
        Set SummarizeTimeDetail = dict
        Exit Function
    End If

    ' Data starts on row 2 since row 1 is headers
    For r = 2 To endRow
        empID = CleanString(ws.Cells(r, colEmpID).Value)
        If empID <> "" Then
            If colFirst <> 0 Then firstName = CleanString(ws.Cells(r, colFirst).Value) Else firstName = ""
            If colLast  <> 0 Then lastName  = CleanString(ws.Cells(r, colLast).Value)  Else lastName  = ""
            If colHours <> 0 Then hours     = CleanNumber(ws.Cells(r, colHours).Value) Else hours     = 0
            If colBranch<> 0 Then branch    = CleanString(ws.Cells(r, colBranch).Value) Else branch   = ""

            ' Name as "Firstname Lastname" (no extra space if one is missing)
            If firstName <> "" And lastName <> "" Then
                empName = firstName & " " & lastName
            Else
                empName = firstName & lastName
            End If

            If Not dict.Exists(empID) Then
                Set s = CreateObject("Scripting.Dictionary")
                s.Add "EmpName", empName
                s.Add "Branch", branch
                s.Add "TotalHours", 0
                dict.Add empID, s
            End If

            Set s = dict(empID)
            s("TotalHours") = s("TotalHours") + hours

            ' Backfill missing name/branch if later rows provide it
            If s("EmpName") = "" And empName <> "" Then s("EmpName") = empName
            If s("Branch")  = "" And branch  <> "" Then s("Branch")  = branch
        End If
    Next

    Set SummarizeTimeDetail = dict
End Function




' Find the column index of a header in row 1 (case-insensitive). Returns 0 if not found.
Function FindHeaderCol(ws, headerName)
    Dim lastCol, c
    lastCol = ws.Cells(1, ws.Columns.Count).End(-4159).Column  ' xlToLeft
    For c = 1 To lastCol
        If StrComp(Trim(CStr(ws.Cells(1, c).Value)), headerName, vbTextCompare) = 0 Then
            FindHeaderCol = c
            Exit Function
        End If
    Next
    FindHeaderCol = 0
End Function



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

Sub InsertStats(c, statsDict, dateVal, moName, br)
    Dim empID, s, sql, checkSQL, rs


    For Each empID In statsDict.Keys
        Set s = statsDict(empID)

        fout.WriteLine PadRight(dateVal, 10) & _
                           PadRight(moName, 10) & _
                           PadRight(s("Branch"), 24) & _
                           PadRight(empID, 10) & _
                           PadRight(s("EmpName"), 24) & _
                           PadRight(s("TotalHours"), 10)
		
        ' === Check for existing record ===
        checkSQL = "SELECT 1 FROM " & tablename & _
                   " WHERE DateCol = '" & dateVal & "'" & _
                   " AND Month = '" & moName & "'" & _
                   " AND Branch = '" & Replace(s("br"), "'", "''") & "'" & _
                   " AND EmpID = '" & empID & "'"

        On Error Resume Next
        Set rs = c.Execute(checkSQL)
        If Err.Number <> 0 Then
            LogError "Check error (" & empID & "): " & Err.Description
            Err.Clear
        ElseIf rs.EOF Then
            ' === Insert if not exists ===
            sql = "INSERT INTO " & tablename & " (DateCol, Month, Branch, EmpID, EmpName, TotalHours) VALUES (" & _
                  "'" & dateVal & "', '" & moName & "', '" & Replace(s("br"), "'", "''") & "', '" & empID & "', '" & Replace(s("EmpName"), "'", "''") & "', '" & s("TotalHours") & "')"

            c.Execute sql
            If Err.Number <> 0 Then LogError "Insert error (" & empID & "): " & Err.Description
        Else
            'fout.WriteLine "Skipped existing row for: " & empID & " | " & s("EmpName")
        End If
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


Function PadRight(val, width)
    PadRight = Left(val & Space(width), width)
End Function




Function GetTotalHoursSummarySheet(wb)
    Dim sh, nm

    ' If there's only one sheet, use it
    If wb.Sheets.Count = 1 Then
        Set GetTotalHoursSummarySheet = wb.Sheets(1)
        Exit Function
    End If

    ' Otherwise, look for "Total Hours Summary" (case-insensitive)
    For Each sh In wb.Sheets
        nm = Trim(CStr(sh.Name))
        If StrComp(nm, "Total Hours Summary", vbTextCompare) = 0 Then
            Set GetTotalHoursSummarySheet = sh
            Exit Function
        End If
    Next

    ' Not found
    Set GetTotalHoursSummarySheet = Nothing
End Function
