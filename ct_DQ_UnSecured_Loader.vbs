Option Explicit

' ---------- ADO CONSTANTS ----------
Const adCmdText      = 1
Const adParamInput   = 1
Const adVarWChar     = 202
Const adDBDate       = 133
Const adInteger      = 3

' ---------- CONFIG ----------
Dim INI_PATH
INI_PATH = "S:\CarolinaTrust\byREQUEST\Templates\DQ_UnSecured.ini"

' ---------- GLOBALS ----------
Dim gCfg, gCn
Dim gConnStr, gFullTableName
Dim gTable2Part, gObjIdName
Dim gHasProcessedHeader
Dim gLineCounter
Dim gSourceFile
Dim gAbortRun
Dim gRowsInserted
Dim gCurrentPeriodDate
Dim gPeriodInitialized

Sub StartDoc()
    On Error Resume Next

    gHasProcessedHeader = False
    gLineCounter = 0
    gRowsInserted = 0
    gAbortRun = False
    gSourceFile = ""
    gCurrentPeriodDate = Null
    gPeriodInitialized = False

    gSourceFile = Spoolfile.Name

    LogLine "Delinquent Consumer Unsecured Import Started: " & Now()
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
    If Len(gTable2Part) = 0 Or Len(gObjIdName) = 0 Then
        LogLine "ERROR: Could not parse TableName: " & gFullTableName
        gAbortRun = True
        EndRun
        Exit Sub
    End If

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
    On Error GoTo 0

    LogLine "SUCCESS: Connected"
    LogLine "Target table = " & gTable2Part

    If Not EnsureStagingTable() Then
        LogLine "ERROR: Could not ensure staging table."
        gAbortRun = True
        EndRun
        Exit Sub
    End If
End Sub

Sub ProcessLine()
    On Error Resume Next
    If gAbortRun Then Exit Sub

    Dim line
    line = srec

    If Trim(line) = "" Then Exit Sub

    If Not gHasProcessedHeader Then
        gHasProcessedHeader = True
        LogLine "Header detected and skipped"
        Exit Sub
    End If

    gLineCounter = gLineCounter + 1
    InsertMappedRow line, gLineCounter

    On Error GoTo 0
End Sub

Sub InsertMappedRow(ByVal srcLine, ByVal lineOrder)
    On Error Resume Next

    Dim periodDate, accountNo, nameVal, typeVal, descVal
    Dim lastPaid, dueDate, daysDQ, balanceRaw, codeVal
    Dim sql, cmd, rowsAff

    periodDate = Null
    accountNo = ""
    nameVal = ""
    typeVal = ""
    descVal = ""
    lastPaid = ""
    dueDate = ""
    daysDQ = ""
    balanceRaw = ""
    codeVal = ""

    If Not ParseDelinquentRow(srcLine, periodDate, accountNo, nameVal, typeVal, descVal, lastPaid, dueDate, daysDQ, balanceRaw, codeVal) Then
        LogLine "SKIP: Could not parse line"
        LogLine "Line: " & Left(srcLine, 300)
        Exit Sub
    End If

    If Not gPeriodInitialized Then
        gCurrentPeriodDate = periodDate

        If Not DeleteOldPeriodRows(gCurrentPeriodDate) Then
            LogLine "ERROR: Could not delete old period rows"
            gAbortRun = True
            Exit Sub
        End If

        If Not DeleteExistingPeriodRows(gCurrentPeriodDate) Then
            LogLine "ERROR: Could not delete existing rows for PeriodDate"
            gAbortRun = True
            Exit Sub
        End If

        gPeriodInitialized = True
        LogLine "Reloading PeriodDate = " & CStr(gCurrentPeriodDate)
    Else
        If CDate(periodDate) <> CDate(gCurrentPeriodDate) Then
            LogLine "ERROR: Multiple PeriodDate values found in same file."
            LogLine "Expected: " & CStr(gCurrentPeriodDate) & " Found: " & CStr(periodDate)
            gAbortRun = True
            Exit Sub
        End If
    End If

    sql = "INSERT INTO " & gTable2Part & " (" & _
          "PeriodDate, AccountNo, FullName, LoanType, Description, LastPaid, DueDate, DaysDQ, BalanceRaw, CodeValue, LineOrder, SourceFile" & _
          ") VALUES (?,?,?,?,?,?,?,?,?,?,?,?)"

    Set cmd = CreateObject("ADODB.Command")

    With cmd
        .ActiveConnection = gCn
        .CommandText = sql
        .CommandType = adCmdText

        AddParameterDate   cmd, periodDate
        AddParameterString cmd, 30,  accountNo
        AddParameterString cmd, 200, nameVal
        AddParameterString cmd, 30,  typeVal
        AddParameterString cmd, 200, descVal
        AddParameterString cmd, 30,  lastPaid
        AddParameterString cmd, 30,  dueDate
        AddParameterString cmd, 30,  daysDQ
        AddParameterString cmd, 30,  balanceRaw
        AddParameterString cmd, 20,  codeVal
        AddParameterInt    cmd,      lineOrder
        AddParameterString cmd, 255, gSourceFile
    End With

    rowsAff = 0
    Err.Clear
    cmd.Execute rowsAff

    If Err.Number <> 0 Then
        LogLine "ERROR: Insert failed (" & Err.Number & ") " & Err.Description
        LogLine "Line: " & Left(srcLine, 300)
        Err.Clear
    Else
        gRowsInserted = gRowsInserted + 1
    End If

    Set cmd = Nothing
    On Error GoTo 0
End Sub

Function ParseDelinquentRow(ByVal srcLine, _
    ByRef outPeriodDate, ByRef outAccountNo, ByRef outName, ByRef outType, _
    ByRef outDesc, ByRef outLastPaid, ByRef outDueDate, ByRef outDaysDQ, _
    ByRef outBalanceRaw, ByRef outCode)

    On Error Resume Next

    Dim arr

    ParseDelinquentRow = False

    outPeriodDate = Null
    outAccountNo = ""
    outName = ""
    outType = ""
    outDesc = ""
    outLastPaid = ""
    outDueDate = ""
    outDaysDQ = ""
    outBalanceRaw = ""
    outCode = ""

    arr = SplitCSV_Delimited(srcLine, ",")

    If Not IsArray(arr) Then Exit Function
    If UBound(arr) < 9 Then Exit Function

    outPeriodDate = ParseDateFlexibleToVariantDate(IntelligentTrim(arr(0)))
    If IsNull(outPeriodDate) Then Exit Function

    outAccountNo  = IntelligentTrim(arr(1))
    outName       = IntelligentTrim(arr(2))
    outType       = IntelligentTrim(arr(3))
    outDesc       = IntelligentTrim(arr(4))
    outLastPaid   = IntelligentTrim(arr(5))
    outDueDate    = IntelligentTrim(arr(6))
    outDaysDQ     = IntelligentTrim(arr(7))
    outBalanceRaw = IntelligentTrim(arr(8))
    outCode       = IntelligentTrim(arr(9))

    If Len(outAccountNo) = 0 Then Exit Function

    ParseDelinquentRow = True
    On Error GoTo 0
End Function

Function EnsureStagingTable()
    On Error Resume Next
    EnsureStagingTable = False

    Dim sql
    sql = "IF OBJECT_ID(N'" & SqlQuote(gObjIdName) & "', N'U') IS NULL " & _
          "BEGIN " & _
          " CREATE TABLE " & gTable2Part & " (" & _
          "  Id INT IDENTITY(1,1) NOT NULL PRIMARY KEY," & _
          "  PeriodDate DATE NOT NULL," & _
          "  AccountNo NVARCHAR(30) NULL," & _
          "  FullName NVARCHAR(200) NULL," & _
          "  LoanType NVARCHAR(30) NULL," & _
          "  Description NVARCHAR(200) NULL," & _
          "  LastPaid NVARCHAR(30) NULL," & _
          "  DueDate NVARCHAR(30) NULL," & _
          "  DaysDQ NVARCHAR(30) NULL," & _
          "  BalanceRaw NVARCHAR(30) NULL," & _
          "  CodeValue NVARCHAR(20) NULL," & _
          "  LineOrder INT NULL," & _
          "  SourceFile NVARCHAR(255) NULL," & _
          "  ImportTimestamp DATETIME NOT NULL DEFAULT(GETDATE())" & _
          " ) " & _
          "END"

    If Not ExecNonQuery(sql) Then Exit Function

    If Not EnsureIndex("IX_DQU_PeriodDate", "PeriodDate") Then Exit Function
    If Not EnsureIndex("IX_DQU_AccountNo", "AccountNo") Then Exit Function
    If Not EnsureIndex("IX_DQU_SourceFile", "SourceFile") Then Exit Function

    EnsureStagingTable = True
    On Error GoTo 0
End Function

Function EnsureIndex(ByVal indexName, ByVal colName)
    On Error Resume Next
    EnsureIndex = False

    Dim sql
    sql = "IF NOT EXISTS (" & _
          "SELECT 1 FROM sys.indexes " & _
          "WHERE name = '" & SqlQuote(indexName) & "' " & _
          "AND object_id = OBJECT_ID(N'" & SqlQuote(gObjIdName) & "')" & _
          ") BEGIN " & _
          "CREATE INDEX [" & indexName & "] ON " & gTable2Part & "([" & colName & "]) " & _
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
        LogLine "SQL: " & Left(sql, 320)
        Err.Clear
        ExecNonQuery = False
    Else
        ExecNonQuery = True
    End If

    Set cmd = Nothing
    On Error GoTo 0
End Function

Function DeleteExistingPeriodRows(ByVal periodDate)
    On Error Resume Next
    DeleteExistingPeriodRows = False

    Dim cmd, rowsAff, sql
    sql = "DELETE FROM " & gTable2Part & " WHERE PeriodDate = ?"

    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = gCn
    cmd.CommandText = sql
    cmd.CommandType = adCmdText

    AddParameterDate cmd, periodDate

    rowsAff = 0
    Err.Clear
    cmd.Execute rowsAff

    If Err.Number <> 0 Then
        LogLine "ERROR deleting existing PeriodDate rows (" & Err.Number & "): " & Err.Description
        Err.Clear
    Else
        LogLine "Deleted existing rows for PeriodDate. Rows affected = " & rowsAff
        DeleteExistingPeriodRows = True
    End If

    Set cmd = Nothing
    On Error GoTo 0
End Function

Function DeleteOldPeriodRows(ByVal anchorPeriodDate)
    On Error Resume Next
    DeleteOldPeriodRows = False

    Dim cutoffDate, cmd, sql, rowsAff

    cutoffDate = DateSerial(Year(anchorPeriodDate), Month(anchorPeriodDate) - 2, 1)

    sql = "DELETE FROM " & gTable2Part & " WHERE PeriodDate < ?"

    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = gCn
    cmd.CommandText = sql
    cmd.CommandType = adCmdText

    AddParameterDate cmd, cutoffDate

    rowsAff = 0
    Err.Clear
    cmd.Execute rowsAff

    If Err.Number <> 0 Then
        LogLine "ERROR deleting old period rows (" & Err.Number & "): " & Err.Description
        Err.Clear
    Else
        LogLine "Deleted old rows where PeriodDate < " & CStr(cutoffDate) & ". Rows affected = " & rowsAff
        DeleteOldPeriodRows = True
    End If

    Set cmd = Nothing
    On Error GoTo 0
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

Function RemoveBrackets(ByVal s)
    s = Replace(s, "[", "")
    s = Replace(s, "]", "")
    RemoveBrackets = s
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

Function IntelligentTrim(value)
    If IsNull(value) Then
        IntelligentTrim = ""
        Exit Function
    End If

    Dim result
    result = CStr(value)
    result = Trim(result)

    If Len(result) >= 2 Then
        If Left(result, 1) = """" And Right(result, 1) = """" Then
            result = Mid(result, 2, Len(result) - 2)
        End If
    End If

    IntelligentTrim = Trim(result)
End Function

Function SplitCSV_Delimited(ByVal line, ByVal delim)
    On Error Resume Next

    Dim fields(), fieldCount, i, inQuotes, fieldValue, ch
    ReDim fields(200)

    fieldCount = 0
    inQuotes = False
    fieldValue = ""

    If Len(delim) = 0 Then delim = ","

    For i = 1 To Len(line)
        ch = Mid(line, i, 1)
        If ch = """" Then
            inQuotes = Not inQuotes
        ElseIf ch = delim And Not inQuotes Then
            fields(fieldCount) = Trim(fieldValue)
            fieldValue = ""
            fieldCount = fieldCount + 1
        Else
            fieldValue = fieldValue & ch
        End If
    Next

    fields(fieldCount) = Trim(fieldValue)
    fieldCount = fieldCount + 1

    ReDim Preserve fields(fieldCount - 1)
    SplitCSV_Delimited = fields
    On Error GoTo 0
End Function

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

Sub AddParameterString(cmdObj, paramSize, paramValue)
    Dim param, val
    On Error Resume Next

    If paramSize <= 0 Then paramSize = 50
    If IsNull(paramValue) Then val = "" Else val = CStr(paramValue)

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

Sub CloseDoc()
    On Error Resume Next

    LogLine "TOTAL ROWS INSERTED: " & gRowsInserted
    LogLine "Delinquent Consumer Unsecured Import Completed: " & Now()
    LogLine "========================================="

    EndRun
    On Error GoTo 0
End Sub