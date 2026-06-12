Option Explicit

' ---------- ADO CONSTANTS ----------
Const adCmdText      = 1
Const adParamInput   = 1
Const adVarWChar     = 202
Const adDBDate       = 133
Const adInteger      = 3

' ---------- CONFIG ----------
Dim INI_PATH
INI_PATH = "S:\CarolinaTrust\byREQUEST\Templates\ACH_List.ini"

' ---------- GLOBALS ----------
Dim gCfg, gCn
Dim gConnStr, gFullTableName
Dim gTable2Part, gObjIdName
Dim gHasProcessedHeader
Dim gLineCounter
Dim gSourceFile
Dim gAbortRun
Dim gRowsInserted
Dim gAnchorDate
Dim gAnchorDateInitialized
Dim gRetentionMonths

Sub StartDoc()
    On Error Resume Next

    gHasProcessedHeader = False
    gLineCounter = 0
    gRowsInserted = 0
    gAbortRun = False
    gSourceFile = ""
    gAnchorDate = Null
    gAnchorDateInitialized = False

    On Error Resume Next
    gSourceFile = Spoolfile.Name
    On Error GoTo 0

    LogLine "ACH_LIST Import Started: " & Now()
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
    gRetentionMonths = GetOptionalInt(gCfg, "Database.RetentionMonths", 3)

    If gRetentionMonths < 0 Then
        LogLine "WARNING: RetentionMonths was less than 0. Defaulting to 3."
        gRetentionMonths = 3
    End If

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

    If Len(gSourceFile) > 0 Then
        If Not DeleteExistingSourceFileRows(gSourceFile) Then
            LogLine "ERROR: Could not delete existing rows for SourceFile"
            gAbortRun = True
            EndRun
            Exit Sub
        End If
    End If
End Sub

Sub ProcessLine()
    On Error Resume Next
    If gAbortRun Then Exit Sub

    Dim line
    line = srec

    If Trim(line) = "" Then Exit Sub

    ' Skip CSV header
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
    Dim fileHdr, pri, recvRT, sendRT, fileDate, fileTime, fileMod
    Dim recsize, blk, fileFormat, receivingName, sendingName, reference
    Dim scc, companyName, companyDisc, companyID, ecc, description
    Dim descDate, entryDate, settleDays, status, originRT, batch
    Dim detail, tran, destRT, individualAcct, amount, individualID
    Dim individualName, disc, addenda, traceNumber
    Dim addendaFlag, addendaType, paymentInfo, addSeq, dtlSeq
    Dim sql, cmd, rowsAff

    report = ""
    rptDate = Null
    rptTime = ""
    rptSeq = ""
    fileHdr = ""
    pri = ""
    recvRT = ""
    sendRT = ""
    fileDate = ""
    fileTime = ""
    fileMod = ""
    recsize = ""
    blk = ""
    fileFormat = ""
    receivingName = ""
    sendingName = ""
    reference = ""
    scc = ""
    companyName = ""
    companyDisc = ""
    companyID = ""
    ecc = ""
    description = ""
    descDate = ""
    entryDate = ""
    settleDays = ""
    status = ""
    originRT = ""
    batch = ""
    detail = ""
    tran = ""
    destRT = ""
    individualAcct = ""
    amount = ""
    individualID = ""
    individualName = ""
    disc = ""
    addendaFlag = ""
    traceNumber = ""
    addenda = ""
    addendaType = ""
    paymentInfo = ""
    addSeq = ""
    dtlSeq = ""

	If Not ParseACHRow(srcLine, _
					   report, rptDate, rptTime, rptSeq, _
					   fileHdr, pri, recvRT, sendRT, _
					   fileDate, fileTime, fileMod, recsize, _
					   blk, fileFormat, receivingName, sendingName, _
					   reference, scc, companyName, companyDisc, _
					   companyID, ecc, description, descDate, _
					   entryDate, settleDays, status, originRT, _
					   batch, detail, tran, destRT, _
					   individualAcct, amount, individualID, _
					   individualName, disc, addendaFlag, _
					   traceNumber, addenda, addendaType, paymentInfo, _
					   addSeq, dtlSeq) Then

		LogLine "SKIP: Could not parse ACH line " & lineOrder
		Exit Sub
	End If

	sql = "INSERT INTO " & gTable2Part & " (" & _
		  "[Report], [RptDate], [RptTime], [RptSeq], [FileHdr], [Pri], [RecvRT], [SendRT], " & _
		  "[FileDate], [FileTime], [Mod], [Recsize], [Blk], [Format], [ReceivingName], [SendingName], [Reference], " & _
		  "[SCC], [CompanyName], [CompanyDisc], [CompanyID], [ECC], [Description], [DescDate], [EntryDate], " & _
		  "[SettleDays], [Status], [OriginRT], [Batch], [Detail], [TranCode], [DestRT], [IndividualAcct], [Amount], " & _
		  "[IndividualID], [IndividualName], [DiscPymtCode], [AddendaFlag], [TraceNumber], [Addenda], [AddendaType], [PaymentInfo], " & _
		  "[AddSeq], [DtlSeq], [LineOrder], [SourceFile]" & _
		  ") VALUES (" & _
		  "?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?" & _
		  ")"

    Set cmd = CreateObject("ADODB.Command")

    With cmd
        .ActiveConnection = gCn
        .CommandText = sql
        .CommandType = adCmdText

        AddParameterString cmd, 50,  report
        AddParameterDate   cmd,      rptDate
        AddParameterString cmd, 20,  rptTime
        AddParameterString cmd, 20,  rptSeq
        AddParameterString cmd, 20,  fileHdr
        AddParameterString cmd, 10,  pri
        AddParameterString cmd, 30,  recvRT
        AddParameterString cmd, 30,  sendRT

        AddParameterString cmd, 20,  fileDate
        AddParameterString cmd, 20,  fileTime
        AddParameterString cmd, 10,  fileMod
        AddParameterString cmd, 20,  recsize
        AddParameterString cmd, 10,  blk
        AddParameterString cmd, 20,  fileFormat
        AddParameterString cmd, 100, receivingName
        AddParameterString cmd, 100, sendingName
        AddParameterString cmd, 50,  reference

        AddParameterString cmd, 10,  scc
        AddParameterString cmd, 100, companyName
        AddParameterString cmd, 100, companyDisc
        AddParameterString cmd, 50,  companyID
        AddParameterString cmd, 10,  ecc
        AddParameterString cmd, 50,  description
        AddParameterString cmd, 20,  descDate
        AddParameterString cmd, 20,  entryDate

        AddParameterString cmd, 20,  settleDays
        AddParameterString cmd, 20,  status
        AddParameterString cmd, 30,  originRT
        AddParameterString cmd, 20,  batch
        AddParameterString cmd, 20,  detail
        AddParameterString cmd, 20,  tran
        AddParameterString cmd, 30,  destRT
        AddParameterString cmd, 50,  individualAcct
        AddParameterString cmd, 30,  amount

        AddParameterString cmd, 50,  individualID
        AddParameterString cmd, 100, individualName
        AddParameterString cmd, 20,  disc
        AddParameterString cmd, 20,  addendaFlag
        AddParameterString cmd, 50,  traceNumber
        AddParameterString cmd, 30,  addenda
        AddParameterString cmd, 30,  addendaType
        AddParameterString cmd, 255, paymentInfo
        AddParameterString cmd, 30,  addSeq
        AddParameterString cmd, 30,  dtlSeq

        AddParameterInt    cmd,      lineOrder
        AddParameterString cmd, 255, gSourceFile
    End With

    rowsAff = 0
    Err.Clear
    cmd.Execute rowsAff

    If Err.Number <> 0 Then
        LogLine "ERROR: ACH insert failed (" & Err.Number & ") " & Err.Description
        LogLine "Line: " & Left(srcLine, 500)
        Err.Clear
    Else
        gRowsInserted = gRowsInserted + 1
        If gRowsInserted Mod 1000 = 0 Then
            LogLine "Progress: " & gRowsInserted & " ACH rows inserted"
        End If
    End If

    Set cmd = Nothing
    On Error GoTo 0
End Sub

Function ParseACHRow(ByVal srcLine, _
    ByRef outReport, ByRef outRptDate, ByRef outRptTime, ByRef outRptSeq, _
    ByRef outFileHdr, ByRef outPri, ByRef outRecvRT, ByRef outSendRT, _
    ByRef outFileDate, ByRef outFileTime, ByRef outMod, ByRef outRecsize, _
    ByRef outBlk, ByRef outFormat, ByRef outReceivingName, ByRef outSendingName, _
    ByRef outReference, ByRef outSCC, ByRef outCompanyName, ByRef outCompanyDisc, _
    ByRef outCompanyID, ByRef outECC, ByRef outDescription, ByRef outDescDate, _
    ByRef outEntryDate, ByRef outSettleDays, ByRef outStatus, ByRef outOriginRT, _
    ByRef outBatch, ByRef outDetail, ByRef outTran, ByRef outDestRT, _
    ByRef outIndividualAcct, ByRef outAmount, ByRef outIndividualID, _
    ByRef outIndividualName, ByRef outDisc, ByRef outAddendaFlag, _
    ByRef outTraceNumber, ByRef outAddenda, ByRef outAddendaType, _
    ByRef outPaymentInfo, _
    ByRef outAddSeq, ByRef outDtlSeq)

    On Error Resume Next

    Dim arr

    ParseACHRow = False

    outReport = ""
    outRptDate = Null
    outRptTime = ""
    outRptSeq = ""
    outFileHdr = ""
    outPri = ""
    outRecvRT = ""
    outSendRT = ""
    outFileDate = ""
    outFileTime = ""
    outMod = ""
    outRecsize = ""
    outBlk = ""
    outFormat = ""
    outReceivingName = ""
    outSendingName = ""
    outReference = ""
    outSCC = ""
    outCompanyName = ""
    outCompanyDisc = ""
    outCompanyID = ""
    outECC = ""
    outDescription = ""
    outDescDate = ""
    outEntryDate = ""
    outSettleDays = ""
    outStatus = ""
    outOriginRT = ""
    outBatch = ""
    outDetail = ""
    outTran = ""
    outDestRT = ""
    outIndividualAcct = ""
    outAmount = ""
    outIndividualID = ""
    outIndividualName = ""
    outDisc = ""
    outAddendaFlag = ""
    outTraceNumber = ""
    outAddenda = ""
    outAddendaType = ""
    outPaymentInfo = ""
    outAddSeq = ""
    outDtlSeq = ""

    arr = SplitCSV_Delimited(srcLine, ",")

    If Not IsArray(arr) Then Exit Function

    If UBound(arr) < 43 Then
        LogLine "WARNING: Expected 44 fields, got " & (UBound(arr) + 1) & " fields"
        Exit Function
    End If

    outReport         = IntelligentTrim(arr(0))
    outRptDate        = ParseDateFlexibleToVariantDate(IntelligentTrim(arr(1)))
    outRptTime        = IntelligentTrim(arr(2))
    outRptSeq         = IntelligentTrim(arr(3))
    outFileHdr        = IntelligentTrim(arr(4))
    outPri            = IntelligentTrim(arr(5))
    outRecvRT         = IntelligentTrim(arr(6))
    outSendRT         = IntelligentTrim(arr(7))
    outFileDate       = IntelligentTrim(arr(8))
    outFileTime       = IntelligentTrim(arr(9))
    outMod            = IntelligentTrim(arr(10))
    outRecsize        = IntelligentTrim(arr(11))
    outBlk            = IntelligentTrim(arr(12))
    outFormat         = IntelligentTrim(arr(13))
    outReceivingName  = IntelligentTrim(arr(14))
    outSendingName    = IntelligentTrim(arr(15))
    outReference      = IntelligentTrim(arr(16))
    outSCC            = IntelligentTrim(arr(17))
    outCompanyName    = IntelligentTrim(arr(18))
    outCompanyDisc    = IntelligentTrim(arr(19))
    outCompanyID      = IntelligentTrim(arr(20))
    outECC            = IntelligentTrim(arr(21))
    outDescription    = IntelligentTrim(arr(22))
    outDescDate       = IntelligentTrim(arr(23))
    outEntryDate      = IntelligentTrim(arr(24))
    outSettleDays     = IntelligentTrim(arr(25))
    outStatus         = IntelligentTrim(arr(26))
    outOriginRT       = IntelligentTrim(arr(27))
    outBatch          = IntelligentTrim(arr(28))
    outDetail         = IntelligentTrim(arr(29))
    outTran           = IntelligentTrim(arr(30))
    outDestRT         = IntelligentTrim(arr(31))
    outIndividualAcct = IntelligentTrim(arr(32))
    outAmount         = IntelligentTrim(arr(33))
    outIndividualID   = IntelligentTrim(arr(34))
    outIndividualName = IntelligentTrim(arr(35))
    outDisc           = IntelligentTrim(arr(36))
    outAddendaFlag    = IntelligentTrim(arr(37))
    outTraceNumber    = IntelligentTrim(arr(38))
    outAddenda        = IntelligentTrim(arr(39))
    outAddendaType    = IntelligentTrim(arr(40))
    outPaymentInfo    = IntelligentTrim(arr(41))
    outAddSeq         = IntelligentTrim(arr(42))
    outDtlSeq         = IntelligentTrim(arr(43))

    If IsNull(outRptDate) Then
        LogLine "WARNING: Invalid RptDate: " & IntelligentTrim(arr(1))
        Exit Function
    End If

    If Len(outTraceNumber) = 0 And Len(outIndividualAcct) = 0 Then
        LogLine "WARNING: No identifying ACH information"
        Exit Function
    End If

    ParseACHRow = True

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
		  "  FileHdr NVARCHAR(20) NULL," & _
		  "  Pri NVARCHAR(10) NULL," & _
		  "  RecvRT NVARCHAR(30) NULL," & _
		  "  SendRT NVARCHAR(30) NULL," & _
		  "  FileDate NVARCHAR(20) NULL," & _
		  "  FileTime NVARCHAR(20) NULL," & _
		  "  Mod NVARCHAR(10) NULL," & _
		  "  Recsize NVARCHAR(20) NULL," & _
		  "  Blk NVARCHAR(10) NULL," & _
		  "  Format NVARCHAR(20) NULL," & _
		  "  ReceivingName NVARCHAR(100) NULL," & _
		  "  SendingName NVARCHAR(100) NULL," & _
		  "  Reference NVARCHAR(50) NULL," & _
		  "  SCC NVARCHAR(10) NULL," & _
		  "  CompanyName NVARCHAR(100) NULL," & _
		  "  CompanyDisc NVARCHAR(100) NULL," & _
		  "  CompanyID NVARCHAR(50) NULL," & _
		  "  ECC NVARCHAR(10) NULL," & _
		  "  Description NVARCHAR(50) NULL," & _
		  "  DescDate NVARCHAR(20) NULL," & _
		  "  EntryDate NVARCHAR(20) NULL," & _
		  "  SettleDays NVARCHAR(20) NULL," & _
		  "  Status NVARCHAR(20) NULL," & _
		  "  OriginRT NVARCHAR(30) NULL," & _
		  "  Batch NVARCHAR(20) NULL," & _
		  "  Detail NVARCHAR(20) NULL," & _
		  "  TranCode NVARCHAR(20) NULL," & _
		  "  DestRT NVARCHAR(30) NULL," & _
		  "  IndividualAcct NVARCHAR(50) NULL," & _
		  "  Amount NVARCHAR(30) NULL," & _
		  "  IndividualID NVARCHAR(50) NULL," & _
		  "  IndividualName NVARCHAR(100) NULL," & _
		  "  DiscPymtCode NVARCHAR(20) NULL," & _
		  "  AddendaFlag NVARCHAR(20) NULL," & _
		  "  TraceNumber NVARCHAR(50) NULL," & _
		  "  Addenda NVARCHAR(30) NULL," & _
		  "  AddendaType NVARCHAR(30) NULL," & _
		  "  PaymentInfo NVARCHAR(255) NULL," & _
		  "  AddSeq NVARCHAR(30) NULL," & _
		  "  DtlSeq NVARCHAR(30) NULL," & _
		  "  LineOrder INT NULL," & _
		  "  SourceFile NVARCHAR(255) NULL," & _
		  "  ImportTimestamp DATETIME NOT NULL DEFAULT(GETDATE())" & _
		  " ) " & _
		  "END"

    If Not ExecNonQuery(sql) Then Exit Function

    If Not EnsureIndex("IX_ACH_RptDate", "RptDate") Then Exit Function
    If Not EnsureIndex("IX_ACH_CompanyName", "CompanyName") Then Exit Function
    If Not EnsureIndex("IX_ACH_IndividualAcct", "IndividualAcct") Then Exit Function
    If Not EnsureIndex("IX_ACH_TraceNumber", "TraceNumber") Then Exit Function
    If Not EnsureIndex("IX_ACH_SourceFile", "SourceFile") Then Exit Function

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

' ============================================================================
' All helper functions below remain the same as your existing code
' ============================================================================

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

Function GetOptionalInt(cfg, key, defaultValue)
    Dim v
    GetOptionalInt = defaultValue

    If cfg Is Nothing Then Exit Function
    If Not cfg.Exists(key) Then Exit Function

    v = Trim(CStr(cfg(key)))
    If Len(v) = 0 Then Exit Function

    If IsNumeric(v) Then
        GetOptionalInt = CLng(v)
    Else
        LogLine "WARNING: INI key '" & key & "' is not numeric. Using default = " & defaultValue
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
    ReDim fields(200)

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
    LogLine "ACH_LIST Import Completed: " & Now()
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



Function NormalizeCardNumber(ByVal s)
    s = Trim(CStr(s))

    If Len(s) >= 2 Then
        If Left(s, 1) = """" And Right(s, 1) = """" Then
            s = Mid(s, 2, Len(s) - 2)
        End If
    End If

    If Left(s, 1) = "'" Then
        s = Mid(s, 2)
    End If

    s = Replace(s, " ", "")
    s = Replace(s, vbTab, "")

    NormalizeCardNumber = s
End Function

Function NormalizeAmountText(ByVal s)
    s = Trim(CStr(s))
    s = Replace(s, ",", "")
    s = Replace(s, "$", "")
    s = Replace(s, "(", "-")
    s = Replace(s, ")", "")
    NormalizeAmountText = Trim(s)
End Function