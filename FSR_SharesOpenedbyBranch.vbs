Option Explicit

Dim folderPath, connString, tablename
	folderPath = "R:\Branch Manager Hillary Reports\"  ' Root folder for all branches
	tablename = "FSR_SharesOpenedbyBranch"
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


    ' === Write Header ===
    fout.WriteLine PadRight("DateCol", 10) & _
                   PadRight("Month", 10) & _
                   PadRight("Branch", 24) & _
                   PadRight("EmpID", 10) & _
                   PadRight("NewSavings", 12) & _
                   PadRight("SumNewSavings", 14)
				   
				   
    For Each branchFolder In fso.GetFolder(folderPath).SubFolders

        branchName = branchFolder.Name
		
		' Remove trailing " Branch" if it exists (case-insensitive)
	    If LCase(Right(branchName, 7)) = " branch" Then
			branchName = Left(branchName, Len(branchName) - 7)
		End If
		branchName = Trim(branchName)
		
		
        Dim subFolderPath
        subFolderPath = branchFolder.Path & "\Hillary Reports\"

        On Error Resume Next
        If fso.FolderExists(subFolderPath) Then
            On Error GoTo 0
            For Each file In fso.GetFolder(subFolderPath).Files
                skipFile = False
 
				If InStr(LCase(file.Name), "shares_opened_by_branch-") > 0 And _
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
					
                        Set stats = SummarizeSharesOpened(ws, 5, lastRow)
						 
                        Call InsertStats(conn, stats, DateCol, folderMonth, branchName)
					 
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

    ' Remove extension safely
    baseName = Left(fileName, InStrRev(fileName, ".") - 1)

    ' Split by space and get last part
    parts = Split(baseName, " ")
    If UBound(parts) >= 0 Then
        strDate = Trim(parts(UBound(parts)))
        If IsDate(strDate) Then
            dt = CDate(strDate)
            ' Get last day of the month
            lastDay = DateSerial(Year(dt), Month(dt) + 1, 0)
            ' Format as string YYYYMMDD
            ExtractFileDate = Year(lastDay) & Right("0" & Month(lastDay), 2) & Right("0" & Day(lastDay), 2)
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
    sql = "IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '" & tablename & "') BEGIN " & _
          "CREATE TABLE " & tablename & " (" & _
          "DateCol NVARCHAR(8), " & _
          "Month NVARCHAR(20), " & _
          "Branch NVARCHAR(100), " & _
          "EmpID NVARCHAR(50), " & _
          "NewSavings NVARCHAR(50), " & _
          "SumNewSavings NVARCHAR(50)" & _
          ") END"

    On Error Resume Next
    c.Execute sql
    If Err.Number <> 0 Then LogError "CreateTable failed: " & Err.Description
    On Error GoTo 0
End Sub



Function SummarizeSharesOpened(ws, startRow, endRow)
    Dim dict: Set dict = CreateObject("Scripting.Dictionary")
    Dim r, empID, balance

    For r = startRow To endRow
        empID = CleanString(ws.Cells(r, "F").Value) ' "Created By User"
        balance = CleanNumber(ws.Cells(r, "H").Value)

        ' Skip invalid users
       If IsEmpty(empID) Or empID = "" Or InStr(empID, "-----") > 0 Or Not IsNumeric(empID) Then
            ' skip
        Else
            If Not dict.Exists(empID) Then
                Dim s: Set s = CreateObject("Scripting.Dictionary")
                s.Add "NewSavings", 0
                s.Add "SumNewSavings", 0
                dict.Add empID, s
            End If

            Set s = dict(empID)
            s("NewSavings") = s("NewSavings") + 1
            s("SumNewSavings") = s("SumNewSavings") + balance
        End If
    Next

    Set SummarizeSharesOpened = dict
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
                       PadRight(s("NewSavings"), 12) & _
                       PadRight(s("SumNewSavings"), 14)		           



        ' === Check for existing record ===
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
            ' === Insert if not found ===
            sql = "INSERT INTO " & tablename & " (DateCol, Month, Branch, EmpID, NewSavings, SumNewSavings) VALUES (" & _
                  "'" & dateVal & "', '" & moName & "', '" & Replace(br, "'", "''") & "', '" & Replace(empID, "'", "''") & "', '" & _
                  s("NewSavings") & "', '" & s("SumNewSavings") & "')"


            c.Execute sql
            If Err.Number <> 0 Then LogError "Insert error (" & empID & "): " & Err.Description
        Else
            'fout.WriteLine "Skipped existing row for: " & empID
        End If
        On Error GoTo 0
    Next
End Sub




Function PadRight(val, width)
    PadRight = Left(val & Space(width), width)
End Function
