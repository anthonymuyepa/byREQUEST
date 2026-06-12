Option Explicit

Const adCmdText      = 1
Const adParamInput   = 1
Const adVarWChar     = 202
Const adDBDate       = 133
Const adInteger      = 3
Const adNumeric      = 131

Dim INI_PATH
INI_PATH = "S:\CarolinaTrust\byREQUEST\Templates\ACH_ControlTotals_Import.ini"

Dim gCfg, gCn
Dim gConnStr, gFullTableName, gTable2Part, gObjIdName
Dim gReport, gRptDate, gRptTime, gRptSeq
Dim gSourceFile, gAbortRun, gRowsInserted, gLineCounter
Dim gHaveColumnHeader
Dim gPosPosted, gPosPostedAmt, gPosUnposted, gPosUnpostedAmt, gPosUncollected, gPosUncollectedAmt

Sub StartDoc()
    On Error Resume Next

    gReport = ""
    gRptDate = Null
    gRptTime = ""
    gRptSeq = ""
    gAbortRun = False
    gRowsInserted = 0
    gLineCounter = 0
    gHaveColumnHeader = False

    gPosPosted = 0
    gPosPostedAmt = 0
    gPosUnposted = 0
    gPosUnpostedAmt = 0
    gPosUncollected = 0
    gPosUncollectedAmt = 0


	fout.WriteLine _
    PadRight("Report", 35) & _
    PadRight("RptDate", 12) & _
    PadRight("RptTime", 10) & _
    PadRight("RptSeq", 10) & _
    PadRight("Category", 35) & _
    PadRight("PostedCnt", 10) & _
    PadRight("PostedAmt", 15) & _
    PadRight("UnpostCnt", 12) & _
    PadRight("UnpostAmt", 15) & _
    PadRight("UncollCnt", 14) & _
    PadRight("UncollAmt", 15) & _
    PadRight("LineOrder", 10) & _
    PadRight("SourceFile", 60)

    gSourceFile = ""
    gSourceFile = SpoolFile.Name

    LogLine "ACH Posting Journal Control Import Started: " & Now()
    LogLine "INI_PATH = " & INI_PATH
    LogLine "SourceFile = " & gSourceFile

    Set gCfg = ReadIniSafe(INI_PATH)

    If gCfg Is Nothing Then
        LogLine "ERROR: INI not found/unreadable: " & INI_PATH
        gAbortRun = True
        EndRun
        Exit Sub
    End If



    gConnStr = GetRequired(gCfg, "Database.ConnStr")
    gFullTableName = GetRequired(gCfg, "Database.TableName")

    If Len(gConnStr) = 0 Or Len(gFullTableName) = 0 Then
        LogLine "ERROR: Missing required INI keys."
        gAbortRun = True
        EndRun
        Exit Sub
    End If

    BuildTableNames gFullTableName, gTable2Part, gObjIdName

    Set gCn = CreateObject("ADODB.Connection")
    Err.Clear
    gCn.Open gConnStr

    If Err.Number <> 0 Then
        LogLine "ERROR: DB open failed (" & Err.Number & ") " & Err.Description
        Err.Clear
        gAbortRun = True
        EndRun
        Exit Sub
    End If

    If Not EnsureStagingTable() Then
        LogLine "ERROR: Could not ensure staging table."
        gAbortRun = True
        EndRun
        Exit Sub
    End If

    If Len(gSourceFile) > 0 Then DeleteExistingSourceFileRows gSourceFile

    On Error GoTo 0
End Sub

Sub ProcessLine()
    On Error Resume Next

    If gAbortRun Then Exit Sub

    Dim line
    line = srec

    If Len(Trim(line)) = 0 Then Exit Sub

    If IsTitleLine(line) Then
        ParseTitleLine line
        Exit Sub
    End If

    If IsColumnHeader(line) Then
        CaptureColumnPositions line
        Exit Sub
    End If

    If Not gHaveColumnHeader Then Exit Sub

    If IsSkipLine(line) Then Exit Sub

    gLineCounter = gLineCounter + 1
    ParseAndInsertControlLine line, gLineCounter

    On Error GoTo 0
End Sub

Sub CloseDoc()
    On Error Resume Next

    LogLine "TOTAL CONTROL ROWS INSERTED: " & gRowsInserted
    LogLine "ACH Posting Journal Control Import Completed: " & Now()
    LogLine "========================================="
    EndRun

    On Error GoTo 0
End Sub

Function PadRight(ByVal s, ByVal n)
    s = Trim(s & "")
    If Len(s) >= n Then
        PadRight = Left(s, n)
    Else
        PadRight = s & Space(n - Len(s))
    End If
End Function

Sub ParseAndInsertControlLine(ByVal line, ByVal lineOrder)
    On Error Resume Next

    Dim category
    Dim postedCnt, postedAmt
    Dim unpostedCnt, unpostedAmt
    Dim uncollectedCnt, uncollectedAmt
    Dim re, m

    postedCnt = ""
    postedAmt = ""
    unpostedCnt = ""
    unpostedAmt = ""
    uncollectedCnt = ""
    uncollectedAmt = ""

    ' =====================================================
    ' Special handling for breakdown/summary rows:
    ' Category + Count + Amount
    ' =====================================================
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^\s*(.+?)\s+(-?\d[\d,]*)\s+(-?\d[\d,]*\.\d{2})\s*$"
    re.IgnoreCase = True
    re.Global = False

    If re.Test(line) Then
        Set m = re.Execute(line)(0)

        category = Trim(m.SubMatches(0))
        unpostedCnt = CleanInt(m.SubMatches(1))
        unpostedAmt = CleanMoney(m.SubMatches(2))

        If Len(category) > 0 And (Len(unpostedCnt) > 0 Or Len(unpostedAmt) > 0) Then
            InsertControlRow category, "", "", unpostedCnt, unpostedAmt, "", "", lineOrder
        End If

        Exit Sub
    End If

    ' =====================================================
    ' Normal fixed-position parsing
    ' =====================================================
    category = Trim(Left(line, gPosPosted - 1))

    If Len(category) = 0 Then Exit Sub
    If InStr(1, category, "Breakdown of Transactions", vbTextCompare) > 0 Then Exit Sub

    postedCnt = CleanInt(SafeSlice(line, gPosPosted, gPosPostedAmt - gPosPosted))
    postedAmt = CleanMoney(SafeSlice(line, gPosPostedAmt, gPosUnposted - gPosPostedAmt))

    unpostedCnt = CleanInt(SafeSlice(line, gPosUnposted, gPosUnpostedAmt - gPosUnposted))
    unpostedAmt = CleanMoney(SafeSlice(line, gPosUnpostedAmt, gPosUncollected - gPosUnpostedAmt))

    uncollectedCnt = CleanInt(SafeSlice(line, gPosUncollected, gPosUncollectedAmt - gPosUncollected))
    uncollectedAmt = CleanMoney(SafeSlice(line, gPosUncollectedAmt, 999))

    If Len(postedCnt) = 0 And Len(postedAmt) = 0 And _
       Len(unpostedCnt) = 0 And Len(unpostedAmt) = 0 And _
       Len(uncollectedCnt) = 0 And Len(uncollectedAmt) = 0 Then
        Exit Sub
    End If

    InsertControlRow category, postedCnt, postedAmt, unpostedCnt, unpostedAmt, uncollectedCnt, uncollectedAmt, lineOrder

    On Error GoTo 0
End Sub

Sub InsertControlRow(ByVal category, ByVal postedCnt, ByVal postedAmt, ByVal unpostedCnt, ByVal unpostedAmt, ByVal uncollectedCnt, ByVal uncollectedAmt, ByVal lineOrder)
    On Error Resume Next

    Dim sql, cmd, rowsAff

    sql = "INSERT INTO " & gTable2Part & " (" & _
          "[Report],[RptDate],[RptTime],[RptSeq],[Category]," & _
          "[PostedCount],[PostedAmount],[UnpostedCount],[UnpostedAmount]," & _
          "[UncollectedCount],[UncollectedAmount],[LineOrder],[SourceFile]" & _
          ") VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)"

    Set cmd = CreateObject("ADODB.Command")

    With cmd
        .ActiveConnection = gCn
        .CommandText = sql
        .CommandType = adCmdText

        AddParameterString cmd, 80, gReport
        AddParameterDate cmd, gRptDate
        AddParameterString cmd, 20, gRptTime
        AddParameterString cmd, 20, gRptSeq
        AddParameterString cmd, 100, category

        AddParameterNullableInt cmd, postedCnt
        AddParameterNullableMoney cmd, postedAmt
        AddParameterNullableInt cmd, unpostedCnt
        AddParameterNullableMoney cmd, unpostedAmt
        AddParameterNullableInt cmd, uncollectedCnt
        AddParameterNullableMoney cmd, uncollectedAmt

        AddParameterInt cmd, lineOrder
        AddParameterString cmd, 255, gSourceFile
    End With

    rowsAff = 0
    Err.Clear
    cmd.Execute rowsAff

    If Err.Number <> 0 Then
        LogLine "ERROR inserting control row (" & Err.Number & "): " & Err.Description
        LogLine "Category: " & category
        Err.Clear
    Else
        gRowsInserted = gRowsInserted + 1
		fout.WriteLine _
        PadRight(gReport, 35) & _
        PadRight(gRptDate, 12) & _
        PadRight(gRptTime, 10) & _
        PadRight(gRptSeq, 10) & _
        PadRight(category, 35) & _
        PadRight(postedCnt, 10) & _
        PadRight(postedAmt, 15) & _
        PadRight(unpostedCnt, 12) & _
        PadRight(unpostedAmt, 15) & _
        PadRight(uncollectedCnt, 14) & _
        PadRight(uncollectedAmt, 15) & _
        PadRight(lineOrder, 10) & _
        PadRight(gSourceFile, 60)
    End If

    Set cmd = Nothing
    On Error GoTo 0
End Sub

Function IsTitleLine(ByVal line)
    IsTitleLine = False

    If InStr(1, line, "ACH Posting Journal", vbTextCompare) > 0 Then
        IsTitleLine = True
    End If

    If InStr(1, line, "ACH Origination Posting Journal", vbTextCompare) > 0 Then
        IsTitleLine = True
    End If
	
	If InStr(1, line, "ACH Origination Exceptions", vbTextCompare) > 0 Then
        IsTitleLine = True
    End If
End Function

Sub ParseTitleLine(ByVal line)
    On Error Resume Next

    Dim pAt, pSeq, pPage

    If InStr(1, line, "ACH Origination Posting Journal", vbTextCompare) > 0 Then
        gReport = "ACH Origination Posting Journal"
    ElseIf InStr(1, line, "ACH Posting Journal", vbTextCompare) > 0 Then
        gReport = "ACH Posting Journal"
		
	ElseIf InStr(1, line, "ACH Origination Exceptions", vbTextCompare) > 0 Then
        gReport = "ACH Origination Exceptions"
    End If
 

    pAt = InStr(1, line, " at ", vbTextCompare)
    pSeq = InStr(1, line, " Seq ", vbTextCompare)
    pPage = InStr(1, line, " Page ", vbTextCompare)

    If pAt > 8 Then
        gRptDate = ParseDateFlexibleToVariantDate(Trim(Mid(line, pAt - 8, 8)))
        gRptTime = Trim(Mid(line, pAt + 13, 5))
    End If

    If pSeq > 0 Then
        If pPage > pSeq Then
            gRptSeq = Trim(Mid(line, pSeq + 5, pPage - (pSeq + 5)))
        Else
            gRptSeq = Trim(Mid(line, pSeq + 5))
        End If
    End If

    On Error GoTo 0
End Sub

Function IsColumnHeader(ByVal line)
    IsColumnHeader = False

    If InStr(1, line, "Posted", vbTextCompare) > 0 And _
       InStr(1, line, "Unposted", vbTextCompare) > 0 And _
       InStr(1, line, "Uncollected", vbTextCompare) > 0 Then
        IsColumnHeader = True
    End If
End Function

Sub CaptureColumnPositions(ByVal line)
    gPosPosted = InStr(1, line, "Posted", vbTextCompare)
    gPosPostedAmt = InStr(gPosPosted + 1, line, "Amount", vbTextCompare)
    gPosUnposted = InStr(1, line, "Unposted", vbTextCompare)
    gPosUnpostedAmt = InStr(gPosUnposted + 1, line, "Amount", vbTextCompare)
    gPosUncollected = InStr(1, line, "Uncollected", vbTextCompare)
    gPosUncollectedAmt = InStr(gPosUncollected + 1, line, "Amount", vbTextCompare)

    If gPosPosted > 0 And gPosPostedAmt > 0 And gPosUnposted > 0 And gPosUnpostedAmt > 0 And gPosUncollected > 0 And gPosUncollectedAmt > 0 Then
        gHaveColumnHeader = True
    End If
End Sub

Function IsSkipLine(ByVal line)
    Dim s
    s = Trim(line & "")

    IsSkipLine = False

    If Len(s) = 0 Then IsSkipLine = True
    If InStr(s, "-----") > 0 Then IsSkipLine = True
    If InStr(1, s, "Transaction", vbTextCompare) > 0 Then IsSkipLine = True
    If InStr(1, s, "Count", vbTextCompare) > 0 Then IsSkipLine = True
    If InStr(1, s, "Amount", vbTextCompare) = 1 Then IsSkipLine = True
    If InStr(1, s, "Principal", vbTextCompare) = 1 Then IsSkipLine = True
    If InStr(1, s, "Interest", vbTextCompare) = 1 Then IsSkipLine = True
    If InStr(1, s, "Fees", vbTextCompare) = 1 Then IsSkipLine = True
    If InStr(1, s, "Unapplied Pmt", vbTextCompare) = 1 Then IsSkipLine = True
    If InStr(1, s, "Sales Tax", vbTextCompare) = 1 Then IsSkipLine = True
    If InStr(1, s, "Billed Fees", vbTextCompare) = 1 Then IsSkipLine = True
    If InStr(1, s, "Misc Expense", vbTextCompare) = 1 Then IsSkipLine = True
    If InStr(1, s, "Breakdown of Transactions", vbTextCompare) > 0 Then IsSkipLine = True
End Function

Function SafeSlice(ByVal s, ByVal startPos, ByVal lengthVal)
    If startPos <= 0 Then
        SafeSlice = ""
    ElseIf startPos > Len(s) Then
        SafeSlice = ""
    Else
        SafeSlice = Trim(Mid(s, startPos, lengthVal))
    End If
End Function

Function CleanInt(ByVal s)
    s = Trim(CStr(s))
    s = Replace(s, ",", "")
    If IsNumeric(s) Then CleanInt = CStr(CLng(s)) Else CleanInt = ""
End Function

Function CleanMoney(ByVal s)
    s = Trim(CStr(s))
    s = Replace(s, ",", "")
    s = Replace(s, "$", "")
    If IsNumeric(s) Then CleanMoney = CStr(CDbl(s)) Else CleanMoney = ""
End Function

Function EnsureStagingTable()
    On Error Resume Next
    EnsureStagingTable = False

    Dim sql

    sql = "IF OBJECT_ID(N'" & SqlQuote(gObjIdName) & "', N'U') IS NULL " & _
          "BEGIN " & _
          " CREATE TABLE " & gTable2Part & " (" & _
          "  Id INT IDENTITY(1,1) NOT NULL PRIMARY KEY," & _
          "  Report NVARCHAR(80) NULL," & _
          "  RptDate DATE NULL," & _
          "  RptTime NVARCHAR(20) NULL," & _
          "  RptSeq NVARCHAR(20) NULL," & _
          "  Category NVARCHAR(100) NULL," & _
          "  PostedCount INT NULL," & _
          "  PostedAmount DECIMAL(18,2) NULL," & _
          "  UnpostedCount INT NULL," & _
          "  UnpostedAmount DECIMAL(18,2) NULL," & _
          "  UncollectedCount INT NULL," & _
          "  UncollectedAmount DECIMAL(18,2) NULL," & _
          "  LineOrder INT NULL," & _
          "  SourceFile NVARCHAR(255) NULL," & _
          "  ImportTimestamp DATETIME NOT NULL DEFAULT(GETDATE())" & _
          " ) " & _
          "END"

    If Not ExecNonQuery(sql) Then Exit Function

    If Not EnsureIndex("IX_ACH_POSTCTRL_RptDate", "RptDate") Then Exit Function
    If Not EnsureIndex("IX_ACH_POSTCTRL_Report", "Report") Then Exit Function
    If Not EnsureIndex("IX_ACH_POSTCTRL_Category", "Category") Then Exit Function
    If Not EnsureIndex("IX_ACH_POSTCTRL_SourceFile", "SourceFile") Then Exit Function

    EnsureStagingTable = True
    On Error GoTo 0
End Function

Function EnsureIndex(ByVal indexName, ByVal colName)
    On Error Resume Next
    EnsureIndex = False

    Dim sql

    sql = "IF NOT EXISTS (" & _
          "SELECT 1 FROM sys.indexes " & _
          "WHERE name = N'" & SqlQuote(indexName) & "' " & _
          "AND object_id = OBJECT_ID(N'" & SqlQuote(gObjIdName) & "')" & _
          ") BEGIN " & _
          "CREATE INDEX [" & Replace(indexName, "]", "]]") & "] ON " & gTable2Part & "([" & Replace(colName, "]", "]]") & "]) " & _
          "END"

    EnsureIndex = ExecNonQuery(sql)
    On Error GoTo 0
End Function

Function ExecNonQuery(ByVal sql)
    On Error Resume Next
    ExecNonQuery = False

    Dim cmd
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = gCn
    cmd.CommandType = adCmdText
    cmd.CommandText = sql

    Err.Clear
    cmd.Execute

    If Err.Number <> 0 Then
        LogLine "SQL ERROR (" & Err.Number & "): " & Err.Description
        LogLine "SQL: " & Left(sql, 500)
        Err.Clear
    Else
        ExecNonQuery = True
    End If

    Set cmd = Nothing
    On Error GoTo 0
End Function

Function SqlQuote(ByVal s)
    SqlQuote = Replace(s, "'", "''")
End Function


Function ReadIniSafe(path)
    On Error Resume Next

    Dim fso, ts, d, sect, line, p

    Set ReadIniSafe = Nothing

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then Exit Function

    Set ts = fso.OpenTextFile(path, 1)
    Set d = CreateObject("Scripting.Dictionary")

    sect = ""

    Do Until ts.AtEndOfStream
        line = Trim(ts.ReadLine)

        If Len(line) = 0 Or Left(line, 1) = ";" Then
            ' skip blank/comment lines
        ElseIf Left(line, 1) = "[" And Right(line, 1) = "]" Then
            sect = Mid(line, 2, Len(line) - 2)
        Else
            p = InStr(line, "=")
            If p > 0 Then
                d(sect & "." & Trim(Left(line, p - 1))) = Trim(Mid(line, p + 1))
            End If
        End If
    Loop

    ts.Close

    Set ReadIniSafe = d

    On Error GoTo 0
End Function

Function GetRequired(cfg, key)
    If Not (cfg Is Nothing) And cfg.Exists(key) And Len(cfg(key)) > 0 Then
        GetRequired = cfg(key)
    Else
        LogLine "ERROR: Missing INI key: " & key
        GetRequired = ""
    End If
End Function

Sub BuildTableNames(ByVal fullName, ByRef out2Part, ByRef outObjIdName)
    Dim s, parts, schemaName, tableName

    s = RemoveBrackets(Trim(fullName))
    parts = Split(s, ".")

    schemaName = "dbo"
    tableName = ""

    If UBound(parts) = 0 Then
        tableName = Trim(parts(0))
    ElseIf UBound(parts) = 1 Then
        schemaName = Trim(parts(0))
        tableName = Trim(parts(1))
    Else
        schemaName = Trim(parts(UBound(parts) - 1))
        tableName = Trim(parts(UBound(parts)))
    End If

    If Len(tableName) = 0 Then
        out2Part = ""
        outObjIdName = ""
        Exit Sub
    End If

    out2Part = "[" & schemaName & "].[" & tableName & "]"
    outObjIdName = schemaName & "." & tableName
End Sub



Function DeleteExistingSourceFileRows(ByVal sourceFile)
    On Error Resume Next
    DeleteExistingSourceFileRows = False

    Dim cmd, rowsAff, sql

    sql = "DELETE FROM " & gTable2Part & " WHERE SourceFile = ?"

    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = gCn
    cmd.CommandText = sql
    cmd.CommandType = adCmdText

    AddParameterString cmd, 255, sourceFile

    rowsAff = 0
    Err.Clear
    cmd.Execute rowsAff

    If Err.Number <> 0 Then
        LogLine "ERROR deleting existing SourceFile rows (" & Err.Number & "): " & Err.Description
        Err.Clear
    Else
        LogLine "Deleted existing rows for SourceFile. Rows affected = " & rowsAff
        DeleteExistingSourceFileRows = True
    End If

    Set cmd = Nothing

    On Error GoTo 0
End Function


Sub AddParameterString(cmdObj, paramSize, paramValue)
    Dim param, val
    On Error Resume Next

    If paramSize <= 0 Then paramSize = 50

    If IsNull(paramValue) Then
        val = ""
    Else
        val = CStr(paramValue)
    End If

    Set param = cmdObj.CreateParameter("", adVarWChar, adParamInput, paramSize, val)
    cmdObj.Parameters.Append param

    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

Sub AddParameterDate(cmdObj, dateVal)
    Dim param
    On Error Resume Next

    Set param = cmdObj.CreateParameter("", adDBDate, adParamInput, , dateVal)
    cmdObj.Parameters.Append param

    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

Sub AddParameterInt(cmdObj, intVal)
    Dim param
    On Error Resume Next

    Set param = cmdObj.CreateParameter("", adInteger, adParamInput, , intVal)
    cmdObj.Parameters.Append param

    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

Sub LogLine(msg)
    fout.WriteLine msg
End Sub

Sub EndRun()
    On Error Resume Next

    If Not gCn Is Nothing Then
        gCn.Close
        Set gCn = Nothing
        LogLine "DB connection closed"
    End If

    On Error GoTo 0
End Sub


Function ParseDateFlexibleToVariantDate(ByVal s)
    On Error Resume Next

    ParseDateFlexibleToVariantDate = Null

    s = Trim(CStr(s))
    If Len(s) = 0 Then Exit Function

    s = Replace(s, "-", "/")

    Dim p, mm, dd, yy
    p = Split(s, "/")

    If UBound(p) <> 2 Then Exit Function

    mm = CLng(p(0))
    dd = CLng(p(1))
    yy = CLng(p(2))

    If yy < 100 Then yy = 2000 + yy

    ParseDateFlexibleToVariantDate = DateSerial(yy, mm, dd)

    If Err.Number <> 0 Then
        Err.Clear
        ParseDateFlexibleToVariantDate = Null
    End If

    On Error GoTo 0
End Function



Function RemoveBrackets(ByVal s)
    s = Replace(s, "[", "")
    s = Replace(s, "]", "")
    RemoveBrackets = s
End Function


Sub AddParameterNullableInt(cmdObj, val)
    If Len(Trim(CStr(val))) = 0 Then
        cmdObj.Parameters.Append cmdObj.CreateParameter("", adInteger, adParamInput, , Null)
    Else
        cmdObj.Parameters.Append cmdObj.CreateParameter("", adInteger, adParamInput, , CLng(val))
    End If
End Sub

Sub AddParameterNullableMoney(cmdObj, val)
    Dim p

    Set p = cmdObj.CreateParameter("", adNumeric, adParamInput, , Null)
    p.Precision = 18
    p.NumericScale = 2

    If Len(Trim(CStr(val))) > 0 Then
        p.Value = CDbl(val)
    Else
        p.Value = Null
    End If

    cmdObj.Parameters.Append p
End Sub