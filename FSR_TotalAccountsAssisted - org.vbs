Option Explicit

' === CONFIG ===
    Dim folder, file, xl, wb, ws, conn, fout
    Dim folderPath, connString, logPath, logFile, tablename
    Dim entryDate, branch, empId, empName, totalEntries, excelPath
	Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
'connString = "Provider=SQLOLEDB;Data Source=BYRequest\SQLEXPRESS;Initial Catalog=byREQUEST;Integrated Security=SSPI;"


    folderPath = "R:\byrequest\reports\Lobby Tracking\"
	tablename = "FSR_TotalAccountsAssisted"
	connString = "Provider=SQLOLEDB;" & _
				 "Data Source=GOLDMINE-SQL1;" & _
				 "Initial Catalog=byREQUEST;" & _
				 "Integrated Security=SSPI;"



Sub StartDoc()
    Dim months, m, folderMonth, monthPath, monthName, entryDate

    ' Initialize
    entryDate = Year(Date) & Right("0" & Month(Date), 2) & Right("0" & Day(Date), 2)

    months = Array("January", "February", "March", "April", "May", "June", _
                   "July", "August", "September", "October", "November", "December")

    ' Open DB connection
    On Error Resume Next
    Set conn = CreateObject("ADODB.Connection")
    conn.Open connString
    If conn.State <> 1 Then
        Call LogError("Failed to open database connection." & vbCrLf & "Error: " & Err.Description)
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    Call LogInfo("Database connection established.")

    ' Create table if needed
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
                    "DateCol NVARCHAR(8)," & _
                    "Month NVARCHAR(20)," & _
                    "Branch NVARCHAR(100)," & _
                    "FSR_ID NVARCHAR(10)," & _
                    "FSR_Name NVARCHAR(100)," & _
                    "TotalPSProvided NVARCHAR(20)," & _
                    "NoSales NVARCHAR(20)," & _
                    "ProductsSold NVARCHAR(20)" & _
                    ")"
        On Error Resume Next
        conn.Execute createSQL
        If Err.Number <> 0 Then
            Call LogError("Failed to create table '" & tablename & "': " & Err.Description)
            conn.Close : Set conn = Nothing
            Exit Sub
        End If
        On Error GoTo 0
        Call LogInfo("Table '" & tablename & "' created successfully.")
    End If

    ' Iterate months
    For m = 0 To UBound(months)
        folderMonth = months(m) & "\"
        monthPath = folderPath
        If Right(monthPath, 1) <> "\" Then monthPath = monthPath & "\"
        monthPath = monthPath & folderMonth

        If Not fso.FolderExists(monthPath) Then
            Call LogError("Folder not found: " & monthPath)
        Else
            Call LogInfo("Processing folder: " & monthPath)

            For Each file In fso.GetFolder(monthPath).Files
                If LCase(fso.GetExtensionName(file.Name)) = "xlsx" And _
                   InStr(LCase(file.Name), "~") = 0 And _
                   InStr(LCase(file.Name), "total accounts assisted by user report") > 0 Then

                    excelPath = file.Path
                    Call LogInfo("Opening file: " & excelPath)

                    On Error Resume Next
                    Dim skipFile : skipFile = False
                    Set xl = CreateObject("Excel.Application")
                    xl.Visible = False
                    Set wb = xl.Workbooks.Open(excelPath)
                    If Err.Number <> 0 Or wb Is Nothing Then
                        Call LogError("Failed to open Excel file: " & excelPath & vbCrLf & "Error: " & Err.Description)
                        Err.Clear : skipFile = True
                    End If
                    On Error GoTo 0

                    If Not skipFile Then
                        Set ws = wb.Sheets(1)
                        If ws Is Nothing Then
                            Call LogError("No worksheet found in: " & excelPath)
                            skipFile = True
                        End If
                    End If

                    If Not skipFile Then
                        Dim row, lastRow
                        row = 6
                        lastRow = ws.UsedRange.Rows.Count

                        Dim idVal, nameVal, FSR_ID, FSR_Name, TotalPSProvided
                        Dim noSales, productsSold, r2, c, val
                        Dim branchName, checkSQL, insertSQL, rsCheck, monthEndDate

                        branchName = Replace(file.Name, "Total Accounts Assisted by User Report - ", "")
                        branchName = Replace(branchName, ".xlsx", "")
                        branchName = Trim(branchName)

                        monthName = months(m)
						Dim dtLastDay
						dtLastDay = DateSerial(Year(Date), m + 2, 0) ' Last day of month m+1
					
						monthEndDate = Year(dtLastDay) & Right("0" & Month(dtLastDay), 2) & Right("0" & Day(dtLastDay), 2)


                        Do While True
                            If Trim(ws.Cells(row, 2).Value) = "Total Accounts Assisted" Then Exit Do

                            idVal = Trim(ws.Cells(row, 2).Value)
                            nameVal = Trim(ws.Cells(row, 3).Value)

                            If InStr(idVal, "-") > 0 And IsNumeric(Left(idVal, 3)) And nameVal <> "" Then
                                FSR_ID = Trim(Replace(idVal, "-", ""))
                                FSR_Name = nameVal
                                TotalPSProvided = Trim(ws.Cells(row, 40).Value)
                                If TotalPSProvided = "" Then TotalPSProvided = "0"

                                noSales = 0
                                r2 = row + 49
                                If Trim(ws.Cells(r2, 3).Value) = "No Sales" Then
                                    For c = 5 To 37 ' E to AK
                                        val = ws.Cells(r2, c).Value
                                        If IsNumeric(val) Then noSales = noSales + CDbl(val)
                                    Next
                                    row = r2
                                End If

                                If IsNumeric(TotalPSProvided) And IsNumeric(noSales) Then
                                    productsSold = CInt(TotalPSProvided) - CInt(noSales)
                                Else
                                    productsSold = "0"
                                End If

                                Call LogInfo("Parsed: " & FSR_ID & " - " & FSR_Name & " | P/S: " & TotalPSProvided & " | NoSales: " & noSales & " | Sold: " & productsSold)

                                checkSQL = "SELECT 1 FROM " & tablename & _
                                           " WHERE DateCol = '" & monthEndDate & "'" & _
                                           " AND Month = '" & Replace(monthName, "'", "''") & "'" & _
                                           " AND Branch = '" & Replace(branchName, "'", "''") & "'" & _
                                           " AND FSR_ID = '" & Replace(FSR_ID, "'", "''") & "'"

                                Set rsCheck = conn.Execute(checkSQL)
                                If rsCheck.EOF Then
                                    insertSQL = "INSERT INTO " & tablename & " (DateCol, Month, Branch, FSR_ID, FSR_Name, TotalPSProvided, NoSales, ProductsSold) VALUES (" & _
                                                "'" & monthEndDate & "'," & _
                                                "'" & Replace(monthName, "'", "''") & "'," & _
                                                "'" & Replace(branchName, "'", "''") & "'," & _
                                                "'" & Replace(FSR_ID, "'", "''") & "'," & _
                                                "'" & Replace(FSR_Name, "'", "''") & "'," & _
                                                "'" & TotalPSProvided & "'," & _
                                                "'" & noSales & "'," & _
                                                "'" & productsSold & "')"
                                    conn.Execute insertSQL
                                    Call LogInfo("Inserted row for: " & FSR_ID & " - " & FSR_Name)
                                Else
                                    Call LogInfo("Skipped existing row for: " & FSR_ID & " - " & FSR_Name)
                                End If
                                rsCheck.Close : Set rsCheck = Nothing
                            End If

                            row = row + 1
                        Loop
                    End If

                    ' Cleanup Excel
                    If Not wb Is Nothing Then wb.Close False
                    If Not xl Is Nothing Then xl.Quit
                    Set ws = Nothing : Set wb = Nothing : Set xl = Nothing
                End If
            Next

            Call LogInfo("Done processing folder: " & monthPath)
        End If
    Next

    ' Close DB connection
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
        Call LogInfo("Database connection closed.")
    End If

    Call LogInfo("All months and files processed. Script completed.")
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

