Option Explicit

Dim dictUsers
Dim iniPath
Dim connStr
Dim tableName
Dim sqlProductCode
Dim addProductCategoryCode
Dim addProductClassCode
Dim addProductCode
Dim loadCtlNbr

Sub StartDoc()

    Set dictUsers = CreateObject("Scripting.Dictionary")

    iniPath = "S:\CarolinaTrust\byREQUEST\Templates\ESTATEMENTS.INI"

    connStr = ReadIniValue(iniPath, "Database", "ConnStr", "")
    tableName = ReadIniValue(iniPath, "Database", "TableName", "HMBERSHIP.MEMBERSHIPSERVICE")

    sqlProductCode = ReadIniValue(iniPath, "EStatements", "SqlProductCode", "2081")

    addProductCategoryCode = ReadIniValue(iniPath, "EStatements", "AddProductCategoryCode", "3")
    addProductClassCode = ReadIniValue(iniPath, "EStatements", "AddProductClassCode", "1")
    addProductCode = ReadIniValue(iniPath, "EStatements", "AddProductCode", "20281")

    loadCtlNbr = GetLastDayOfCurrentMonth()

    fout.WriteLine "# DEBUG ConnStr=" & connStr
    fout.WriteLine "# DEBUG TableName=" & tableName
    fout.WriteLine "# DEBUG LoadCtlNbr=" & loadCtlNbr

    WriteSqlRows

End Sub

Sub ProcessLine()

    Dim line
    Dim parts
    Dim userId

    line = Trim(srec)

    If line = "" Then Exit Sub

    parts = Split(line, vbTab)

    If UBound(parts) < 0 Then Exit Sub

    userId = Trim(parts(0))

    ' Skip header, totals, blank rows, and non-numeric rows.
    If IsNumeric(userId) Then

        userId = CStr(CLng(userId))

        If CLng(userId) > 0 Then
            If Not dictUsers.Exists(userId) Then
                dictUsers.Add userId, CLng(userId)
            End If
        End If

    End If

End Sub

Sub CloseDoc()

    Dim keys
    Dim i, j
    Dim temp
    Dim outLine

    keys = dictUsers.Keys

    If dictUsers.Count = 0 Then
        fout.WriteLine "# NO SPOOL USER ID ROWS FOUND"
        Set dictUsers = Nothing
        Exit Sub
    End If

    ' Sort ascending numerically.
    If dictUsers.Count > 1 Then

        For i = 0 To UBound(keys) - 1
            For j = i + 1 To UBound(keys)

                If CLng(keys(i)) > CLng(keys(j)) Then
                    temp = keys(i)
                    keys(i) = keys(j)
                    keys(j) = temp
                End If

            Next
        Next

    End If

    ' Append sorted spool output after SQL output.
    ' Layout:
    ' member_nbr<TAB>Action<TAB>product_category_code<TAB>product_class_code<TAB>product_code<TAB>load_ctl_nbr
    For i = 0 To UBound(keys)

        outLine = keys(i) & vbTab & _
                  "A" & vbTab & _
                  addProductCategoryCode & vbTab & _
                  addProductClassCode & vbTab & _
                  addProductCode & vbTab & _
                  loadCtlNbr

        fout.WriteLine outLine

    Next

    Set dictUsers = Nothing

End Sub

Sub WriteSqlRows()

    On Error Resume Next

    Dim cn
    Dim rs
    Dim sql
    Dim outLine
    Dim rowCount
    Dim errMsg

    rowCount = 0

    If Trim(connStr) = "" Then
        fout.WriteLine "# SQL SKIPPED: ConnStr is blank or missing in ESTATEMENTS.INI"
        Exit Sub
    End If

    sql = ""
    sql = sql & "SELECT " & vbCrLf
    sql = sql & "    member_nbr, " & vbCrLf
    sql = sql & "    'R' AS Action, " & vbCrLf
    sql = sql & "    product_category_code, " & vbCrLf
    sql = sql & "    product_class_code, " & vbCrLf
    sql = sql & "    product_code, " & vbCrLf
    sql = sql & "    load_ctl_nbr " & vbCrLf
    sql = sql & "FROM " & tableName & " " & vbCrLf
    sql = sql & "WHERE product_code IN (" & sqlProductCode & ") " & vbCrLf
    sql = sql & "ORDER BY 1"

    Set cn = CreateObject("ADODB.Connection")
    cn.Open connStr

    If Err.Number <> 0 Then
        errMsg = Err.Description
        fout.WriteLine "# SQL ERROR: Could not open database connection - " & errMsg
        Err.Clear
        On Error GoTo 0
        Set cn = Nothing
        Exit Sub
    End If

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, cn, 0, 1

    If Err.Number <> 0 Then
        errMsg = Err.Description
        fout.WriteLine "# SQL ERROR: Query failed - " & errMsg
        Err.Clear

        If Not rs Is Nothing Then
            If rs.State = 1 Then rs.Close
        End If

        If Not cn Is Nothing Then
            If cn.State = 1 Then cn.Close
        End If

        Set rs = Nothing
        Set cn = Nothing

        On Error GoTo 0
        Exit Sub
    End If

    If rs.EOF Then

        fout.WriteLine "# SQL NO OUTPUT: No existing R rows found for product_code " & sqlProductCode

    Else

        Do Until rs.EOF

            outLine = CleanField(rs.Fields("member_nbr").Value) & vbTab & _
                      CleanField(rs.Fields("Action").Value) & vbTab & _
                      CleanField(rs.Fields("product_category_code").Value) & vbTab & _
                      CleanField(rs.Fields("product_class_code").Value) & vbTab & _
                      CleanField(rs.Fields("product_code").Value) & vbTab & _
                      CleanField(rs.Fields("load_ctl_nbr").Value)

            fout.WriteLine outLine

            rowCount = rowCount + 1
            rs.MoveNext

        Loop

        fout.WriteLine "# SQL OUTPUT COMPLETE: " & rowCount & " R rows written"

    End If

    If rs.State = 1 Then rs.Close
    If cn.State = 1 Then cn.Close

    Set rs = Nothing
    Set cn = Nothing

    On Error GoTo 0

End Sub

Function ExtractLoadCtlNbrFromFileName(fileName)

    Dim fso
    Dim baseName
    Dim parts
    Dim i
    Dim piece

    ExtractLoadCtlNbrFromFileName = ""

    Set fso = CreateObject("Scripting.FileSystemObject")

    baseName = fso.GetFileName(fileName)

    Set fso = Nothing

    parts = Split(baseName, "-")

    For i = 0 To UBound(parts)

        piece = Trim(parts(i))

        If Len(piece) = 8 Then
            If IsNumeric(piece) Then
                ExtractLoadCtlNbrFromFileName = piece
                Exit Function
            End If
        End If

    Next

End Function

Function CleanField(value)

    If IsNull(value) Then
        CleanField = ""
    Else
        CleanField = Trim(CStr(value))
    End If

End Function

Function ReadIniValue(filePath, sectionName, keyName, defaultValue)

    Dim fso
    Dim ts
    Dim line
    Dim currentSection
    Dim pos
    Dim k
    Dim v

    ReadIniValue = defaultValue

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(filePath) Then
        Set fso = Nothing
        Exit Function
    End If

    Set ts = fso.OpenTextFile(filePath, 1, False)

    currentSection = ""

    Do Until ts.AtEndOfStream

        line = Trim(ts.ReadLine)

        If line <> "" Then

            If Left(line, 1) <> ";" Then

                If Left(line, 1) = "[" And Right(line, 1) = "]" Then

                    currentSection = Mid(line, 2, Len(line) - 2)

                ElseIf LCase(currentSection) = LCase(sectionName) Then

                    pos = InStr(line, "=")

                    If pos > 0 Then

                        k = Trim(Left(line, pos - 1))
                        v = Trim(Mid(line, pos + 1))

                        If LCase(k) = LCase(keyName) Then
                            ReadIniValue = v
                            ts.Close
                            Set ts = Nothing
                            Set fso = Nothing
                            Exit Function
                        End If

                    End If

                End If

            End If

        End If

    Loop

    ts.Close

    Set ts = Nothing
    Set fso = Nothing

End Function


Function GetLastDayOfCurrentMonth()

    Dim d
    Dim lastDay

    ' Get the last day of the current month.
    ' Using DateAdd to handle December correctly (month overflow)
    
    lastDay = DateAdd("d", -1, DateSerial(Year(Date), Month(Date) + 1, 1))

    ' Return as YYYYMMDD.
    GetLastDayOfCurrentMonth = Year(lastDay) & _
                               Right("0" & Month(lastDay), 2) & _
                               Right("0" & Day(lastDay), 2)

End Function