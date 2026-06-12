Option Explicit

' === CONFIG ===
Dim folderPath, connString, tablename, fso, conn
folderPath = "R:\byrequest\reports\Lobby Tracking\"
tablename = "FSR_TotalAccountsAssisted"
connString = "Provider=SQLOLEDB;" & _
             "Data Source=GOLDMINE-SQL1;" & _
             "Initial Catalog=byREQUEST;" & _
             "Integrated Security=SSPI;"


Sub StartDoc()
    Dim xl, file, wb, ws
    Dim monthNames, thisYear, thisMonth, m
    Dim monthName, folderMonth
    Dim dtLastDay, dateCol
    Dim skipFile

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set xl  = CreateObject("Excel.Application")
    xl.Visible = False

    thisYear  = Year(Date)
    thisMonth = Month(Date)

    monthNames = Array("January","February","March","April","May","June", _
                       "July","August","September","October","November","December")

    ' === Connect to database once ===
    Set conn = OpenConnection()
    If conn Is Nothing Then Exit Sub
    Call CreateTableIfNeeded()

    ' === Write header once ===
    fout.WriteLine PadRight("DateCol", 10) & _
                   PadRight("Month", 10) & _
                   PadRight("Branch", 24) & _
                   PadRight("FSR_ID", 10) & _
                   PadRight("FSR_Name", 24) & _
                   PadRight("TotalPS", 10) & _
                   PadRight("NoSales", 10) & _
                   PadRight("Sold", 10)

    ' === LOOP: Jan -> current month ===
    For m = 1 To thisMonth
        monthName  = monthNames(m - 1)
        folderMonth = "R:\byrequest\reports\Lobby Tracking\" & monthName & "\"

        ' Last day of this month (YYYYMMDD) for DateCol
        dtLastDay = DateSerial(thisYear, m + 1, 0)
        dateCol   = Year(dtLastDay) & Right("0" & Month(dtLastDay), 2) & Right("0" & Day(dtLastDay), 2)

        If Not fso.FolderExists(folderMonth) Then
            Call LogError("Folder not found: " & folderMonth)
            ' keep going to next month
        Else
            For Each file In fso.GetFolder(folderMonth).Files
                If InStr(LCase(file.Name), "total accounts assisted by user report") > 0 And _
                   InStr(file.Name, "~") = 0 Then

                    Set wb = xl.Workbooks.Open(file.Path)
                    Set ws = wb.Sheets(1)
                    skipFile = False

                    ' === Extract from B6 ===
                    Dim b6Val, parts, rawMonth, rawBranch, mnth
                    b6Val = ws.Range("B6").Value
                    parts = Split(b6Val, "|")
                    If UBound(parts) >= 2 Then
                        rawMonth  = Trim(Split(parts(2), ":")(1)) ' "March 2025"
                        mnth      = Split(rawMonth, " ")
                        rawMonth  = mnth(0)                       ' "March"
                        rawBranch = Trim(Split(parts(3), ":")(1))  ' e.g., Daingerfield
                    Else
                        Call LogError("Invalid B6 format: " & b6Val & " | File: " & file.Name)
                        skipFile = True
                    End If

                    If Not skipFile Then
                        Dim row, idVal, nameVal, TotalPSProvided, noSales, productsSold
                        Dim r2, val, c, FSR_ID, FSR_Name
                        Dim checkSQL, insertSQL, rs

                        row = 6
                        Do While True
                            If InStr(LCase(Trim(ws.Cells(row, 2).Value)), "total accounts assisted") > 0 Then Exit Do

                            idVal   = Trim(ws.Cells(row, 2).Value)
                            nameVal = Trim(ws.Cells(row, 3).Value)

                            If InStr(idVal, "-") > 0 And IsNumeric(Left(idVal, 3)) And nameVal <> "" Then
                                FSR_ID   = Trim(Replace(idVal, "-", ""))
                                FSR_Name = nameVal

                                TotalPSProvided = Trim(ws.Cells(row, 40).Value) : If TotalPSProvided = "" Then TotalPSProvided = "0"

                                noSales = 0
                                r2 = row + 49
                                If Trim(ws.Cells(r2, 3).Value) = "No Sales" Then
                                    For c = 5 To 37 ' E to AK
                                        val = ws.Cells(r2, c).Value
                                        If IsNumeric(val) Then noSales = noSales + CDbl(val)
                                    Next
                                    row = r2 ' jump to "No Sales" line
                                End If

                                If IsNumeric(TotalPSProvided) And IsNumeric(noSales) Then
                                    productsSold = CInt(TotalPSProvided) - CInt(noSales)
									' Prevent negative ProductsSold
									If productsSold < 0 Then productsSold = 0
                                Else
                                    productsSold = "0"
                                End If

                                ' === Log row ===
                                fout.WriteLine PadRight(dateCol, 10) & _
                                               PadRight(rawMonth, 10) & _
                                               PadRight(rawBranch, 24) & _
                                               PadRight(FSR_ID, 10) & _
                                               PadRight(FSR_Name, 24) & _
                                               PadRight(TotalPSProvided, 10) & _
                                               PadRight(noSales, 10) & _
                                               PadRight(productsSold, 10)

                                ' === Upsert: skip if exists ===
                                checkSQL = "SELECT 1 FROM " & tablename & _
                                           " WHERE DateCol = '" & dateCol & "'" & _
                                           " AND Month = '" & Replace(rawMonth, "'", "''") & "'" & _
                                           " AND Branch = '" & Replace(rawBranch, "'", "''") & "'" & _
                                           " AND FSR_ID = '" & Replace(FSR_ID, "'", "''") & "'"

                                On Error Resume Next
                                Set rs = conn.Execute(checkSQL)
                                If Err.Number <> 0 Then
                                    fout.WriteLine "Check failed: " & Err.Description & vbCrLf & checkSQL
                                    Err.Clear
                                ElseIf rs.EOF Then
                                    insertSQL = "INSERT INTO " & tablename & _
                                                " (DateCol, Month, Branch, FSR_ID, FSR_Name, TotalPSProvided, NoSales, ProductsSold) VALUES (" & _
                                                "'" & dateCol & "', " & _
                                                "'" & rawMonth & "', " & _
                                                "'" & Replace(rawBranch, "'", "''") & "', " & _
                                                "'" & FSR_ID & "', " & _
                                                "'" & Replace(FSR_Name, "'", "''") & "', " & _
                                                "'" & TotalPSProvided & "', " & _
                                                "'" & noSales & "', " & _
                                                "'" & productsSold & "')"
                                    conn.Execute insertSQL
                                    If Err.Number <> 0 Then
                                        fout.WriteLine "Insert failed: " & Err.Description & vbCrLf & insertSQL
                                        Err.Clear
                                    End If
                                End If
                                On Error GoTo 0
                            End If

                            row = row + 1
                        Loop
                    End If

                    ' Save a uniquely named copy
                    Dim baseName, ext, nameOnly, version, saveCopyPath
                    baseName = fso.GetFileName(file.Path)
                    ext      = fso.GetExtensionName(baseName)
                    nameOnly = Left(baseName, Len(baseName) - Len(ext) - 1)
                    version  = 0
                    Do
                        If version = 0 Then
                            saveCopyPath = "S:\Reports\" & nameOnly & ".xlsx"
                        Else
                            saveCopyPath = "S:\Reports\" & nameOnly & "~" & version & ".xlsx"
                        End If
                        version = version + 1
                    Loop While fso.FileExists(saveCopyPath)

                    On Error Resume Next
                    wb.SaveCopyAs saveCopyPath
                    If Err.Number <> 0 Then
                        Call LogError("Failed to save copy to S:\Reports\: " & Err.Description)
                        Err.Clear
                    End If
                    On Error GoTo 0

                    wb.Close False
                    Set ws = Nothing
                    Set wb = Nothing
                End If
            Next
        End If
    Next ' month loop

    xl.Quit
    If Not conn Is Nothing Then conn.Close
    Set conn = Nothing
    Set xl = Nothing
    Set fso = Nothing
End Sub



' === Helpers ===
Sub CreateTableIfNeeded()
    
    Dim sql
    sql = "IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '" & tablename & "') " & _
          "BEGIN CREATE TABLE " & tablename & " (" & _
          "DateCol NVARCHAR(8), Month NVARCHAR(20), Branch NVARCHAR(100), FSR_ID NVARCHAR(10), " & _
          "FSR_Name NVARCHAR(100), TotalPSProvided NVARCHAR(20), NoSales NVARCHAR(20), ProductsSold NVARCHAR(20)) END"
    
    On Error Resume Next
    conn.Execute sql
    If Err.Number <> 0 Then LogError "CreateTable failed: " & Err.Description
    On Error GoTo 0
End Sub


Sub LogInfo(msg)
    fout.WriteLine Now & " | INFO: " & msg
End Sub

Sub LogError(msg)
    fout.WriteLine Now & " | ERROR: " & msg
End Sub


Function OpenConnection()
    On Error Resume Next

    Dim c
    Set c = CreateObject("ADODB.Connection")
    c.Open connString

    If Err.Number <> 0 Then
        LogError "DB error: " & Err.Description
        Set c = Nothing
        Err.Clear
    End If

    On Error GoTo 0
    Set OpenConnection = c
End Function



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
