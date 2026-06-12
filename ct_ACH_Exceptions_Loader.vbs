Option Explicit

' ---------- ADO CONSTANTS ----------
Const adCmdText      = 1
Const adParamInput   = 1
Const adVarWChar     = 202
Const adDBDate       = 133
Const adInteger      = 3

' ---------- CONFIG ----------
Dim INI_PATH
INI_PATH = "S:\CarolinaTrust\byREQUEST\Templates\ACH_Exceptions.ini"

' ---------- GLOBALS ----------
Dim gCfg, gCn
Dim gConnStr, gFullTableName
Dim gTable2Part, gObjIdName
Dim gHasProcessedHeader
Dim gLineCounter
Dim gSourceFile
Dim gAbortRun
Dim gRowsInserted

Sub StartDoc()
    On Error Resume Next

    gHasProcessedHeader = False
    gLineCounter = 0
    gRowsInserted = 0
    gAbortRun = False
    gSourceFile = ""

    On Error Resume Next
    gSourceFile = SpoolFile.Name
    On Error GoTo 0

    LogLine "ACH_EXCEPTIONS_TRANCODE Import Started: " & Now()
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

    LogLine "SUCCESS: Connected"
    LogLine "Target table = " & gTable2Part

    If Not EnsureStagingTable() Then
        LogLine "ERROR: Could not ensure staging table."
        gAbortRun = True
        EndRun
        Exit Sub
    End If

    If Len(gSourceFile) > 0 Then
        If Not DeleteExistingSourceFileRows(gSourceFile) Then
            LogLine "ERROR: Could not delete existing rows for SourceFile"
            gAbortRun = True
            EndRun
            Exit Sub
        End If
    End If

    On Error GoTo 0
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

    Dim report, rptDate, rptTime, rptSeq
    Dim acct, slx, nameValue, tranDesc, companyName
    Dim amount, available, balance, description
	Dim companyID, settlement, transmission, ecc, disc
    Dim entryValue, origAcct, entryID, entryName
    Dim trace, exceptionText, raw1
    Dim parsedLineOrder, parsedSourceFile
    Dim sql, cmd, rowsAff

    report = ""
    rptDate = Null
    rptTime = ""
    rptSeq = ""
    acct = ""
    slx = ""
    nameValue = ""
    tranDesc = ""
    companyName = ""
    amount = ""
    available = ""
    balance = ""
    description = ""
    companyID = ""
    settlement = ""
    transmission = ""
    ecc = ""
	disc = ""
    entryValue = ""
    origAcct = ""
    entryID = ""
    entryName = ""
    trace = ""
    exceptionText = ""
    raw1 = ""
    parsedLineOrder = ""
    parsedSourceFile = ""

    If Not ParseTranCodeRow(srcLine, _
                            report, rptDate, rptTime, rptSeq, _
                            acct, slx, nameValue, tranDesc, companyName, _
                            amount, available, balance, description, _
                            companyID, settlement, transmission, ecc, _
                            disc, entryValue, origAcct, entryID, entryName, _
                            trace, exceptionText, raw1, _
                            parsedLineOrder, parsedSourceFile) Then

        LogLine "SKIP: Could not parse TranCode line " & lineOrder
        Exit Sub
    End If

    If Len(parsedSourceFile) > 0 Then
        gSourceFile = parsedSourceFile
    End If

    sql = "INSERT INTO " & gTable2Part & " (" & _
          "[Report], [RptDate], [RptTime], [RptSeq], [Acct], [SLX], [Name], [TranDesc], [CompanyName], " & _
          "[Amount], [Available], [Balance], [Description], [CompanyID], [SettlementDate], [TransmissionDate], " & _
          "[ECC],[Disc],[Entry], [OrigAcct], [EntryID], [EntryName], [TraceNumber], [ExceptionText], [Raw1], " & _
          "[LineOrder], [SourceFile]" & _
		  ") VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"

    Set cmd = CreateObject("ADODB.Command")

    With cmd
        .ActiveConnection = gCn
        .CommandText = sql
        .CommandType = adCmdText

        AddParameterString cmd, 50,  report
        AddParameterDate   cmd,      rptDate
        AddParameterString cmd, 20,  rptTime
        AddParameterString cmd, 20,  rptSeq

        AddParameterString cmd, 30,  acct
        AddParameterString cmd, 10,  slx
        AddParameterString cmd, 100, nameValue
        AddParameterString cmd, 50,  tranDesc
        AddParameterString cmd, 100, companyName

        AddParameterString cmd, 30,  amount
        AddParameterString cmd, 30,  available
        AddParameterString cmd, 30,  balance
        AddParameterString cmd, 100, description

        AddParameterString cmd, 50,  companyID
        AddParameterString cmd, 20,  settlement
        AddParameterString cmd, 20,  transmission
        AddParameterString cmd, 10,  ecc

        AddParameterString cmd, 20, disc
		AddParameterString cmd, 50,  entryValue
        AddParameterString cmd, 50,  origAcct
        AddParameterString cmd, 50,  entryID
        AddParameterString cmd, 100, entryName
        AddParameterString cmd, 50,  trace
        AddParameterString cmd, 255, exceptionText
        AddParameterString cmd, 500, raw1

        If IsNumeric(parsedLineOrder) Then
            AddParameterInt cmd, CLng(parsedLineOrder)
        Else
            AddParameterInt cmd, lineOrder
        End If

        AddParameterString cmd, 255, gSourceFile
    End With

    rowsAff = 0
    Err.Clear
    cmd.Execute rowsAff

    If Err.Number <> 0 Then
        LogLine "ERROR: TranCode insert failed (" & Err.Number & ") " & Err.Description
        LogLine "Line: " & Left(srcLine, 500)
        Err.Clear
    Else
        gRowsInserted = gRowsInserted + 1
        If gRowsInserted Mod 1000 = 0 Then
            LogLine "Progress: " & gRowsInserted & " TranCode rows inserted"
        End If
    End If

    Set cmd = Nothing
    On Error GoTo 0
End Sub

Function ParseTranCodeRow(ByVal srcLine, _
    ByRef outReport, ByRef outRptDate, ByRef outRptTime, ByRef outRptSeq, _
    ByRef outAcct, ByRef outSLX, ByRef outName, ByRef outTranDesc, _
    ByRef outCompanyName, ByRef outAmount, ByRef outAvailable, ByRef outBalance, _
    ByRef outDescription, ByRef outCompanyID, ByRef outSettlement, _
    ByRef outTransmission, ByRef outECC, ByRef outDisc, ByRef outEntry, ByRef outOrigAcct, _
    ByRef outEntryID, ByRef outEntryName, ByRef outTrace, _
    ByRef outExceptionText, ByRef outRaw1, ByRef outLineOrder, ByRef outSourceFile)

    On Error Resume Next

    Dim arr
    ParseTranCodeRow = False

    outReport = ""
    outRptDate = Null
    outRptTime = ""
    outRptSeq = ""
    outAcct = ""
    outSLX = ""
    outName = ""
    outTranDesc = ""
    outCompanyName = ""
    outAmount = ""
    outAvailable = ""
    outBalance = ""
    outDescription = ""
    outCompanyID = ""
    outSettlement = ""
    outTransmission = ""
    outECC = ""
    outDisc = ""
    outEntry = ""
    outOrigAcct = ""
    outEntryID = ""
    outEntryName = ""
    outTrace = ""
    outExceptionText = ""
    outRaw1 = ""
    outLineOrder = ""
    outSourceFile = ""

    arr = SplitCSV_Delimited(srcLine, ",")

    If Not IsArray(arr) Then Exit Function

 
	If UBound(arr) < 26 Then
		LogLine "WARNING: Expected 27 fields, got " & (UBound(arr) + 1)
        Exit Function
    End If

    outReport        = IntelligentTrim(arr(0))
    outRptDate       = ParseDateFlexibleToVariantDate(IntelligentTrim(arr(1)))
    outRptTime       = IntelligentTrim(arr(2))
    outRptSeq        = IntelligentTrim(arr(3))
    outAcct          = IntelligentTrim(arr(4))
    outSLX           = IntelligentTrim(arr(5))
    outName          = IntelligentTrim(arr(6))
    outTranDesc      = IntelligentTrim(arr(7))
    outCompanyName   = IntelligentTrim(arr(8))
    outAmount        = IntelligentTrim(arr(9))
    outAvailable     = IntelligentTrim(arr(10))
    outBalance       = IntelligentTrim(arr(11))
    outDescription   = IntelligentTrim(arr(12))
    outCompanyID     = IntelligentTrim(arr(13))
    outSettlement    = IntelligentTrim(arr(14))
    outTransmission  = IntelligentTrim(arr(15))
	outECC           = IntelligentTrim(arr(16))
	outDisc          = IntelligentTrim(arr(17))
	outEntry         = IntelligentTrim(arr(18))
	outOrigAcct      = IntelligentTrim(arr(19))
	outEntryID       = IntelligentTrim(arr(20))
	outEntryName     = IntelligentTrim(arr(21))
	outTrace         = IntelligentTrim(arr(22))
	outExceptionText = IntelligentTrim(arr(23))
	outRaw1          = IntelligentTrim(arr(24))
	outLineOrder     = IntelligentTrim(arr(25))
	outSourceFile    = IntelligentTrim(arr(26))

    If IsNull(outRptDate) Then
        LogLine "WARNING: Invalid RptDate: " & IntelligentTrim(arr(1))
        Exit Function
    End If

    If Len(outAcct) = 0 And Len(outTrace) = 0 And Len(outOrigAcct) = 0 Then
        LogLine "WARNING: No identifying TranCode information"
        Exit Function
    End If

    ParseTranCodeRow = True

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
          "  Report NVARCHAR(50) NULL," & _
          "  RptDate DATE NULL," & _
          "  RptTime NVARCHAR(20) NULL," & _
          "  RptSeq NVARCHAR(20) NULL," & _
          "  Acct NVARCHAR(30) NULL," & _
          "  SLX NVARCHAR(10) NULL," & _
          "  Name NVARCHAR(100) NULL," & _
          "  TranDesc NVARCHAR(50) NULL," & _
          "  CompanyName NVARCHAR(100) NULL," & _
          "  Amount NVARCHAR(30) NULL," & _
          "  Available NVARCHAR(30) NULL," & _
          "  Balance NVARCHAR(30) NULL," & _
          "  Description NVARCHAR(100) NULL," & _
          "  CompanyID NVARCHAR(50) NULL," & _
          "  SettlementDate NVARCHAR(20) NULL," & _
          "  TransmissionDate NVARCHAR(20) NULL," & _
          "  ECC NVARCHAR(10) NULL," & _
          "  Disc NVARCHAR(20) NULL," & _
          "  Entry NVARCHAR(50) NULL," & _
          "  OrigAcct NVARCHAR(50) NULL," & _
          "  EntryID NVARCHAR(50) NULL," & _
          "  EntryName NVARCHAR(100) NULL," & _
          "  TraceNumber NVARCHAR(50) NULL," & _
          "  ExceptionText NVARCHAR(255) NULL," & _
          "  Raw1 NVARCHAR(500) NULL," & _
          "  LineOrder INT NULL," & _
          "  SourceFile NVARCHAR(255) NULL," & _
          "  ImportTimestamp DATETIME NOT NULL DEFAULT(GETDATE())" & _
          " ) " & _
          "END"

    If Not ExecNonQuery(sql) Then Exit Function

    If Not EnsureIndex("IX_ACH_TRANCODE_RptDate", "RptDate") Then Exit Function
    If Not EnsureIndex("IX_ACH_TRANCODE_Acct", "Acct") Then Exit Function
    If Not EnsureIndex("IX_ACH_TRANCODE_Trace", "TraceNumber") Then Exit Function
    If Not EnsureIndex("IX_ACH_TRANCODE_OrigAcct", "OrigAcct") Then Exit Function
    If Not EnsureIndex("IX_ACH_TRANCODE_SourceFile", "SourceFile") Then Exit Function

    EnsureStagingTable = True
    On Error GoTo 0
End Function

Function EnsureIndex(ByVal indexName, ByVal colName)
    On Error Resume Next
    EnsureIndex = False

    Dim safeIndexName, safeColName, sql

    safeIndexName = Replace(indexName, "]", "]]")
    safeColName = Replace(colName, "]", "]]")

    sql = "IF NOT EXISTS (" & _
          "SELECT 1 FROM sys.indexes " & _
          "WHERE name = N'" & SqlQuote(indexName) & "' " & _
          "AND object_id = OBJECT_ID(N'" & SqlQuote(gObjIdName) & "')" & _
          ") BEGIN " & _
          "CREATE INDEX [" & safeIndexName & "] ON " & gTable2Part & "([" & safeColName & "]) " & _
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
        ExecNonQuery = False
    Else
        ExecNonQuery = True
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

    Dim fields(), fieldCount, i, inQuotes, fieldValue, ch, nextCh
    ReDim fields(300)

    fieldCount = 0
    inQuotes = False
    fieldValue = ""

    If Len(delim) = 0 Then delim = ","

    For i = 1 To Len(line)
        ch = Mid(line, i, 1)

        If ch = """" Then
            If inQuotes And i < Len(line) Then
                nextCh = Mid(line, i + 1, 1)
                If nextCh = """" Then
                    fieldValue = fieldValue & """"
                    i = i + 1
                Else
                    inQuotes = False
                End If
            ElseIf Not inQuotes Then
                inQuotes = True
            Else
                inQuotes = False
            End If

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

Sub CloseDoc()
    On Error Resume Next

    LogLine "TOTAL ROWS INSERTED: " & gRowsInserted
    LogLine "ACH_EXCEPTIONS_TRANCODE Import Completed: " & Now()
    LogLine "========================================="

    EndRun

    On Error GoTo 0
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