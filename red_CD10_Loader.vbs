Option Explicit

' ---------- ADO CONSTANTS ----------
Const adCmdText      = 1
Const adParamInput   = 1
Const adVarWChar     = 202
Const adDBDate       = 133
Const adInteger      = 3

' ---------- CONFIG ----------
Dim INI_PATH
INI_PATH = "S:\CarolinaTrust\byREQUEST\Templates\EFT_CD10.ini"

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

    LogLine "EFT_CD10 Import Started: " & Now()
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

    Dim lDate, cardNumber, lTime, transID, terminal, logo, seqNo, tranDesc, msgType, respCode
    Dim amountRaw, amountVal, networkName, settlementDate, chFee, surcharge
    Dim fromAccount, termAddr, rpi, ed, cv, cv2, cv3, av, cavv, eci, posEMode, cardPresent
    Dim mobl, nscr, nrsn, authID, merchID
    Dim sql, cmd, rowsAff

    ' Initialize all variables
    lDate = Null
    cardNumber = ""
    lTime = ""
    transID = ""
    terminal = ""
    logo = ""
    seqNo = ""
    tranDesc = ""
    msgType = ""
    respCode = ""
    amountRaw = ""
    amountVal = ""
    networkName = ""
    settlementDate = ""
    chFee = ""
    surcharge = ""
    fromAccount = ""
    termAddr = ""
    rpi = ""
    ed = ""
    cv = ""
    cv2 = ""
    cv3 = ""
    av = ""
    cavv = ""
    eci = ""
    posEMode = ""
    cardPresent = ""
    mobl = ""
    nscr = ""
    nrsn = ""
    authID = ""
    merchID = ""

    If Not ParseEFTCD10Row(srcLine, lDate, cardNumber, lTime, transID, terminal, logo, seqNo, _
                           tranDesc, msgType, respCode, amountRaw, amountVal, networkName, _
                           settlementDate, chFee, surcharge, fromAccount, termAddr, rpi, _
                           ed, cv, cv2, cv3, av, cavv, eci, posEMode, cardPresent, _
                           mobl, nscr, nrsn, authID, merchID) Then
        LogLine "SKIP: Could not parse line " & lineOrder
        Exit Sub
    End If

    ' First valid row establishes cleanup anchor for retention
    If Not gAnchorDateInitialized Then
        gAnchorDate = lDate
        If Not DeleteOldTranRows(gAnchorDate) Then
            LogLine "ERROR: Could not delete old transaction rows"
            gAbortRun = True
            Exit Sub
        End If
        gAnchorDateInitialized = True
        LogLine "Retention anchor LDate = " & CStr(gAnchorDate)
    End If

    ' Updated INSERT with all new fields
    sql = "INSERT INTO " & gTable2Part & " (" & _
          "LDate, CardNumber, LTime, transID, Terminal, Logo, SeqNo, TranDesc, MsgType, RespCode, " & _
          "AmountRaw, AmountValue, NetworkName, SettlementDate, chFee, Surcharge, " & _
          "FromAccount, TermAddr, RPI, " & _
          "ED, CV, CV2, CV3, AV, CAVV, ECI, PosEMode, CardPresent, " & _
          "MOBL, NSCR, NRSN, AuthID, MerchID, " & _
          "LineOrder, SourceFile" & _
          ") VALUES (" & _
          "?,?,?,?,?,?,?,?,?,?, " & _
          "?,?,?,?,?,?, " & _
          "?,?,?, " & _
          "?,?,?,?,?,?,?,?,?, " & _
          "?,?,?,?,?, " & _
          "?,?)"

    Set cmd = CreateObject("ADODB.Command")

    With cmd
        .ActiveConnection = gCn
        .CommandText = sql
        .CommandType = adCmdText

        ' Date fields
        AddParameterDate   cmd, lDate
        
        ' String fields - Core transaction data
        AddParameterString cmd, 50,  cardNumber
        AddParameterString cmd, 20,  lTime
        AddParameterString cmd, 10,  transID
        AddParameterString cmd, 10,  terminal
        AddParameterString cmd, 10,  logo
        AddParameterString cmd, 50,  seqNo
        AddParameterString cmd, 100, tranDesc
        AddParameterString cmd, 8,   msgType
        AddParameterString cmd, 8,   respCode
        AddParameterString cmd, 50,  amountRaw
        AddParameterString cmd, 50,  amountVal
        AddParameterString cmd, 100, networkName
        AddParameterString cmd, 12,  settlementDate
        AddParameterString cmd, 10,  chFee
        AddParameterString cmd, 10,  surcharge
        
        ' String fields - Account & Terminal info
        AddParameterString cmd, 100, fromAccount
        AddParameterString cmd, 255, termAddr
        AddParameterString cmd, 5,   rpi
        
        ' String fields - Extended security fields
        AddParameterString cmd, 3,   ed
        AddParameterString cmd, 3,   cv
        AddParameterString cmd, 3,   cv2
        AddParameterString cmd, 3,   cv3
        AddParameterString cmd, 3,   av
        AddParameterString cmd, 5,   cavv
        AddParameterString cmd, 5,   eci
        AddParameterString cmd, 8,   posEMode
        AddParameterString cmd, 12,  cardPresent
        
        ' String fields - Status fields
        AddParameterString cmd, 5,   mobl
        AddParameterString cmd, 5,   nscr
        AddParameterString cmd, 5,   nrsn
        AddParameterString cmd, 8,   authID
        AddParameterString cmd, 100, merchID
        
        ' Integer and tracking fields
        AddParameterInt    cmd, lineOrder
        AddParameterString cmd, 255, gSourceFile
    End With

    rowsAff = 0
    Err.Clear
    cmd.Execute rowsAff

    If Err.Number <> 0 Then
        LogLine "ERROR: Insert failed (" & Err.Number & ") " & Err.Description
        LogLine "Line: " & Left(srcLine, 500)
        Err.Clear
    Else
        gRowsInserted = gRowsInserted + 1
        If gRowsInserted Mod 1000 = 0 Then
            LogLine "Progress: " & gRowsInserted & " rows inserted"
        End If
    End If

    Set cmd = Nothing
    On Error GoTo 0
End Sub

Function ParseEFTCD10Row(ByVal srcLine, _
    ByRef outLDate, ByRef outCardNumber, ByRef outLTime, ByRef outID, ByRef outTerminal, _
    ByRef outLogo, ByRef outSeqNo, ByRef outTranDesc, ByRef outMsgType, ByRef outRespCode, _
    ByRef outAmountRaw, ByRef outAmountVal, ByRef outNetworkName, ByRef outSettlementDate, _
    ByRef outChFee, ByRef outSurcharge, ByRef outFromAccount, ByRef outTermAddr, ByRef outRPI, _
    ByRef outED, ByRef outCV, ByRef outCV2, ByRef outCV3, ByRef outAV, ByRef outCAVV, _
    ByRef outECI, ByRef outPosEMode, ByRef outCardPresent, ByRef outMOBL, ByRef outNSCR, _
    ByRef outNRSN, ByRef outAuthID, ByRef outMerchID)

    On Error Resume Next

    Dim arr

    ParseEFTCD10Row = False

    ' Initialize all outputs
    outLDate = Null
    outCardNumber = ""
    outLTime = ""
    outID = ""
    outTerminal = ""
    outLogo = ""
    outSeqNo = ""
    outTranDesc = ""
    outMsgType = ""
    outRespCode = ""
    outAmountRaw = ""
    outAmountVal = ""
    outNetworkName = ""
    outSettlementDate = ""
    outChFee = ""
    outSurcharge = ""
    outFromAccount = ""
    outTermAddr = ""
    outRPI = ""
    outED = ""
    outCV = ""
    outCV2 = ""
    outCV3 = ""
    outAV = ""
    outCAVV = ""
    outECI = ""
    outPosEMode = ""
    outCardPresent = ""
    outMOBL = ""
    outNSCR = ""
    outNRSN = ""
    outAuthID = ""
    outMerchID = ""

    arr = SplitCSV_Delimited(srcLine, ",")

    If Not IsArray(arr) Then Exit Function
    If UBound(arr) < 31 Then
        LogLine "WARNING: Expected 35 fields, got " & (UBound(arr) + 1) & " fields"
        Exit Function
    End If

    ' Field mapping based on CSV header (35 fields total)
    ' Index: 0-based
    ' 0  CardNumber
    ' 1  LDate
    ' 2  LTime
    ' 3  transID
    ' 4  Terminal
    ' 5  Logo
    ' 6  SeqNo
    ' 7  TranDesc
    ' 8  MsgType
    ' 9  RespCode
    ' 10 Amount
    ' 11 NetworkName
    ' 12 SettlementDate
    ' 13 chFee
    ' 14 Surcharge
    ' 15 FromAccount
    ' 16 TermAddr
    ' 17 RPI
    ' 18 ED
    ' 19 CV
    ' 20 CV2
    ' 21 CV3
    ' 22 AV
    ' 23 CAVV
    ' 24 ECI
    ' 25 PosEMode
    ' 26 CardPresent
    ' 27 MOBL
    ' 28 NSCR
    ' 29 NRSN
    ' 30 AuthID
    ' 31 MerchID

    outCardNumber   = NormalizeCardNumber(IntelligentTrim(arr(0)))
    outLDate        = ParseDateFlexibleToVariantDate(IntelligentTrim(arr(1)))
    outLTime        = IntelligentTrim(arr(2))
    outID           = IntelligentTrim(arr(3))
    outTerminal     = IntelligentTrim(arr(4))
    outLogo         = IntelligentTrim(arr(5))
    outSeqNo        = IntelligentTrim(arr(6))
    outTranDesc     = IntelligentTrim(arr(7))
    outMsgType      = IntelligentTrim(arr(8))
    outRespCode     = IntelligentTrim(arr(9))
    outAmountRaw    = IntelligentTrim(arr(10))
    outAmountVal    = NormalizeAmountText(outAmountRaw)
    outNetworkName  = IntelligentTrim(arr(11))
    outSettlementDate = IntelligentTrim(arr(12))
    outChFee        = IntelligentTrim(arr(13))
    outSurcharge    = IntelligentTrim(arr(14))
    outFromAccount  = IntelligentTrim(arr(15))
    outTermAddr     = IntelligentTrim(arr(16))
    outRPI          = IntelligentTrim(arr(17))
    outED           = IntelligentTrim(arr(18))
    outCV           = IntelligentTrim(arr(19))
    outCV2          = IntelligentTrim(arr(20))
    outCV3          = IntelligentTrim(arr(21))
    outAV           = IntelligentTrim(arr(22))
    outCAVV         = IntelligentTrim(arr(23))
    outECI          = IntelligentTrim(arr(24))
    outPosEMode     = IntelligentTrim(arr(25))
    outCardPresent  = IntelligentTrim(arr(26))
    outMOBL         = IntelligentTrim(arr(27))
    outNSCR         = IntelligentTrim(arr(28))
    outNRSN         = IntelligentTrim(arr(29))
    outAuthID       = IntelligentTrim(arr(30))
    outMerchID      = IntelligentTrim(arr(31))

    ' Validation
    If IsNull(outLDate) Then
        LogLine "WARNING: Invalid date format for LDate: " & IntelligentTrim(arr(1))
        Exit Function
    End If
    
    If Len(outCardNumber) = 0 And Len(outSeqNo) = 0 And Len(outMerchID) = 0 Then
        LogLine "WARNING: No identifying information (CardNumber, SeqNo, or MerchID)"
        Exit Function
    End If

    ParseEFTCD10Row = True
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
          "  LDate DATE NOT NULL," & _
          "  CardNumber NVARCHAR(50) NULL," & _
          "  LTime NVARCHAR(20) NULL," & _
          "  TransID NVARCHAR(10) NULL," & _
          "  Terminal NVARCHAR(10) NULL," & _
          "  Logo NVARCHAR(10) NULL," & _
          "  SeqNo NVARCHAR(50) NULL," & _
          "  TranDesc NVARCHAR(100) NULL," & _
          "  MsgType NVARCHAR(8) NULL," & _
          "  RespCode NVARCHAR(8) NULL," & _
          "  AmountRaw NVARCHAR(50) NULL," & _
          "  AmountValue NVARCHAR(50) NULL," & _
          "  NetworkName NVARCHAR(100) NULL," & _
          "  SettlementDate NVARCHAR(12) NULL," & _
          "  chFee NVARCHAR(10) NULL," & _
          "  Surcharge NVARCHAR(10) NULL," & _
          "  FromAccount NVARCHAR(100) NULL," & _
          "  TermAddr NVARCHAR(255) NULL," & _
          "  RPI NVARCHAR(5) NULL," & _
          "  ED NVARCHAR(3) NULL," & _
          "  CV NVARCHAR(3) NULL," & _
          "  CV2 NVARCHAR(3) NULL," & _
          "  CV3 NVARCHAR(3) NULL," & _
          "  AV NVARCHAR(3) NULL," & _
          "  CAVV NVARCHAR(5) NULL," & _
          "  ECI NVARCHAR(5) NULL," & _
          "  PosEMode NVARCHAR(8) NULL," & _
          "  CardPresent NVARCHAR(12) NULL," & _
          "  MOBL NVARCHAR(5) NULL," & _
          "  NSCR NVARCHAR(5) NULL," & _
          "  NRSN NVARCHAR(5) NULL," & _
          "  AuthID NVARCHAR(8) NULL," & _
          "  MerchID NVARCHAR(100) NULL," & _
          "  LineOrder INT NULL," & _
          "  SourceFile NVARCHAR(255) NULL," & _
          "  ImportTimestamp DATETIME NOT NULL DEFAULT(GETDATE())" & _
          " ) " & _
          "END"

    If Not ExecNonQuery(sql) Then Exit Function

    ' Create indexes for performance - single column only
    If Not EnsureIndex("IX_EFTCD10_LDate", "LDate") Then Exit Function
    If Not EnsureIndex("IX_EFTCD10_CardNumber", "CardNumber") Then Exit Function
    If Not EnsureIndex("IX_EFTCD10_SeqNo", "SeqNo") Then Exit Function
    If Not EnsureIndex("IX_EFTCD10_SourceFile", "SourceFile") Then Exit Function

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
    LogLine "EFT_CD10 Import Completed: " & Now()
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

Function DeleteOldTranRows(ByVal anchorDate)
    On Error Resume Next
    DeleteOldTranRows = False

    Dim cutoffDate, cmd, sql, rowsAff

    cutoffDate = DateSerial(Year(anchorDate), Month(anchorDate) - gRetentionMonths, 1)

    sql = "DELETE FROM " & gTable2Part & " WHERE LDate < ?"

    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = gCn
    cmd.CommandText = sql
    cmd.CommandType = adCmdText

    AddParameterDate cmd, cutoffDate

    rowsAff = 0
    Err.Clear
    cmd.Execute rowsAff

    If Err.Number <> 0 Then
        LogLine "ERROR deleting old transaction rows (" & Err.Number & "): " & Err.Description
        Err.Clear
    Else
        LogLine "RetentionMonths = " & gRetentionMonths
        LogLine "Deleted old rows where LDate < " & CStr(cutoffDate) & ". Rows affected = " & rowsAff
        DeleteOldTranRows = True
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