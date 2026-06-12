Option Explicit

' ---------- ADO CONSTANTS ----------
Const adCmdText      = 1
Const adParamInput   = 1
Const adVarWChar     = 202
Const adDBTimeStamp  = 135
Const adInteger      = 3

' ---------- CONFIG ----------
Dim INI_PATH
INI_PATH = "S:\CarolinaTrust\byREQUEST\Templates\ATM_RECON.ini"

' ---------- GLOBALS ----------
Dim gCfg, gCn
Dim gConnStr, gFullTableName
Dim gTable2Part, gObjIdName
Dim gLineCounter
Dim gSourceFile
Dim gAbortRun
Dim gRowsInserted
Dim gAnchorDate
Dim gAnchorDateInitialized
Dim gRetentionMonths
Dim gFirstLineProcessed  ' Flag to track first line

Sub StartDoc()
    On Error Resume Next

    gLineCounter = 0
    gRowsInserted = 0
    gAbortRun = False
    gSourceFile = ""
    gAnchorDate = Null
    gAnchorDateInitialized = False
    gRetentionMonths = 3
    gFirstLineProcessed = False

    On Error Resume Next
    gSourceFile = Spoolfile.Name
    On Error GoTo 0

    LogLine "ATM Reconciliation Import Started: " & Now()
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
        LogLine "WARNING: RetentionMonths < 0. Defaulting to 3."
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
    
    ' Clean the line - remove BOM and other special characters from first line
    If Not gFirstLineProcessed Then
        line = CleanFirstLine(line)
        gFirstLineProcessed = True
    End If

    If Trim(line) = "" Then Exit Sub

    gLineCounter = gLineCounter + 1
    InsertMappedRow line, gLineCounter

    On Error GoTo 0
End Sub

Function CleanFirstLine(ByVal line)
    Dim result, i, ch, asciiVal
    
    result = ""
    
    ' Remove BOM (Byte Order Mark) and other control characters from start
    For i = 1 To Len(line)
        ch = Mid(line, i, 1)
        asciiVal = AscW(ch)
        
        ' Keep only printable characters (ASCII 32-126 or valid Unicode)
        ' Skip BOM (65279) and other control characters at start
        If i = 1 And (asciiVal = 65279 Or asciiVal = 65535 Or asciiVal < 32) Then
            ' Skip BOM and control characters at the beginning
        Else
            result = result & ch
        End If
    Next
    
    ' If result is empty, return original line
    If Len(result) = 0 Then
        CleanFirstLine = line
    Else
        CleanFirstLine = result
    End If
End Function

Function CleanSpecialCharacters(ByVal s)
    Dim result, i, ch, asciiVal
    
    result = ""
    
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        asciiVal = AscW(ch)
        
        ' Keep ONLY numeric digits 0-9
        If (asciiVal >= 48 And asciiVal <= 57) Then
            result = result & ch
        End If
    Next
    
    CleanSpecialCharacters = result
End Function

Sub InsertMappedRow(ByVal srcLine, ByVal lineOrder)
    On Error Resume Next

    Dim cardNumber, tranDateTime, tranDate, tranTime, tranCode, amountRaw, amountVal
    Dim traceNo, terminalID, drCr, referenceNo
    Dim sql, cmd, rowsAff

    ' Initialize all variables
    cardNumber = ""
    tranDateTime = Null
    tranDate = Null
    tranTime = Null
    tranCode = ""
    amountRaw = ""
    amountVal = ""
    traceNo = ""
    terminalID = ""
    drCr = ""
    referenceNo = ""

    If Not ParseATMReconRow(srcLine, cardNumber, tranDateTime, tranCode, amountRaw, amountVal, traceNo, terminalID, drCr, referenceNo) Then
        LogLine "SKIP: Could not parse line " & lineOrder
        LogLine "Line: " & Left(srcLine, 500)
        Exit Sub
    End If
    
    ' Extract date and time from TranDateTime
    If Not IsNull(tranDateTime) And IsDate(tranDateTime) Then
        tranDate = CDate(tranDateTime)  ' This gives just the date part
        tranTime = tranDateTime         ' This will be converted to TIME by SQL
    End If
    
    ' Clean the card number to remove any remaining special characters
    cardNumber = CleanSpecialCharacters(cardNumber)

    ' COMMENTED OUT - Skip deletion for testing
    ' If Not gAnchorDateInitialized Then
    '     gAnchorDate = DateSerial(Year(tranDateTime), Month(tranDateTime), 1)
    '
    '     If Not DeleteOldTranRows(gAnchorDate) Then
    '         LogLine "ERROR: Could not delete old transaction rows"
    '         gAbortRun = True
    '         Exit Sub
    '     End If
    '
    '     gAnchorDateInitialized = True
    '     LogLine "Retention anchor date = " & CStr(gAnchorDate)
    ' End If

    ' Updated INSERT with only TranDate and TranTime (no TranDateTime)
    sql = "INSERT INTO " & gTable2Part & " (" & _
          "CardNumber, TranDate, TranTime, TranCode, AmountRaw, AmountValue, " & _
          "TraceNo, TerminalID, DrCr, ReferenceNo, LineOrder, SourceFile" & _
          ") VALUES (?,?,?,?,?,?,?,?,?,?,?,?)"

    Set cmd = CreateObject("ADODB.Command")

    With cmd
        .ActiveConnection = gCn
        .CommandText = sql
        .CommandType = adCmdText

        AddParameterString     cmd, 50,  cardNumber
        AddParameterDate       cmd, tranDate
        AddParameterTime       cmd, tranTime
        AddParameterString     cmd, 20,  tranCode
        AddParameterString     cmd, 50,  amountRaw
        AddParameterString     cmd, 50,  amountVal
        AddParameterString     cmd, 50,  traceNo
        AddParameterString     cmd, 50,  terminalID
        AddParameterString     cmd, 10,  drCr
        AddParameterString     cmd, 50,  referenceNo
        AddParameterInt        cmd, lineOrder
        AddParameterString     cmd, 255, gSourceFile
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
    End If

    Set cmd = Nothing
    On Error GoTo 0
End Sub

Function ParseATMReconRow(ByVal srcLine, _
    ByRef outCardNumber, ByRef outTranDateTime, ByRef outTranCode, ByRef outAmountRaw, ByRef outAmountVal, _
    ByRef outTraceNo, ByRef outTerminalID, ByRef outDrCr, ByRef outReferenceNo)

    On Error Resume Next

    Dim arr

    ParseATMReconRow = False

    outCardNumber = ""
    outTranDateTime = Null
    outTranCode = ""
    outAmountRaw = ""
    outAmountVal = ""
    outTraceNo = ""
    outTerminalID = ""
    outDrCr = ""
    outReferenceNo = ""

    arr = SplitCSV_Delimited(srcLine, ",")

    If Not IsArray(arr) Then Exit Function
    If UBound(arr) < 7 Then Exit Function

    outCardNumber   = IntelligentTrim(arr(0))
    outTranDateTime = ParseDateTimeFlexible(IntelligentTrim(arr(1)))
    outTranCode     = IntelligentTrim(arr(2))
    outAmountRaw    = IntelligentTrim(arr(3))
    
    ' Store amount as string (preserve original formatting)
    outAmountVal    = CleanAmountString(outAmountRaw)
    
    outTraceNo      = IntelligentTrim(arr(4))
    outTerminalID   = IntelligentTrim(arr(5))
    outDrCr         = IntelligentTrim(arr(6))
    outReferenceNo  = IntelligentTrim(arr(7))

    If Len(outCardNumber) = 0 Then Exit Function
    If IsNull(outTranDateTime) Then Exit Function

    ParseATMReconRow = True
    On Error GoTo 0
End Function

Function CleanAmountString(ByVal s)
    On Error Resume Next
    
    Dim result
    result = Trim(CStr(s))
    
    ' Remove currency symbols and convert parentheses for negatives
    result = Replace(result, "$", "")
    result = Replace(result, ",", "")
    
    ' Handle negative amounts in parentheses
    If Left(result, 1) = "(" And Right(result, 1) = ")" Then
        result = "-" & Trim(Mid(result, 2, Len(result) - 2))
    End If
    
    ' Remove any remaining non-numeric characters except decimal and minus
    Dim cleanResult, i, ch
    cleanResult = ""
    For i = 1 To Len(result)
        ch = Mid(result, i, 1)
        If IsNumeric(ch) Or ch = "." Or ch = "-" Then
            cleanResult = cleanResult & ch
        End If
    Next
    
    CleanAmountString = cleanResult
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
          "  CardNumber NVARCHAR(50) NULL," & _
          "  TranDate DATE NULL," & _
          "  TranTime TIME NULL," & _
          "  TranCode NVARCHAR(20) NULL," & _
          "  AmountRaw NVARCHAR(50) NULL," & _
          "  AmountValue NVARCHAR(50) NULL," & _
          "  TraceNo NVARCHAR(50) NULL," & _
          "  TerminalID NVARCHAR(50) NULL," & _
          "  DrCr NVARCHAR(10) NULL," & _
          "  ReferenceNo NVARCHAR(50) NULL," & _
          "  LineOrder INT NULL," & _
          "  SourceFile NVARCHAR(255) NULL," & _
          "  ImportTimestamp DATETIME NOT NULL DEFAULT(GETDATE())" & _
          " ) " & _
          "END"

    If Not ExecNonQuery(sql) Then Exit Function

    If Not EnsureIndex("IX_ATMRecon_TranDate", "TranDate") Then Exit Function
    If Not EnsureIndex("IX_ATMRecon_CardNumber", "CardNumber") Then Exit Function
    If Not EnsureIndex("IX_ATMRecon_TraceNo", "TraceNo") Then Exit Function
    If Not EnsureIndex("IX_ATMRecon_SourceFile", "SourceFile") Then Exit Function

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

    sql = "DELETE FROM " & gTable2Part & " WHERE TranDate < ?"

    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = gCn
    cmd.CommandText = sql
    cmd.CommandType = adCmdText

    AddParameterDate cmd, cutoffDate

    rowsAff = 0
    Err.Clear
    cmd.Execute rowsAff

    If Err.Number <> 0 Then
        LogLine "ERROR deleting old rows (" & Err.Number & "): " & Err.Description
        Err.Clear
    Else
        LogLine "RetentionMonths = " & gRetentionMonths
        LogLine "Deleted old rows where TranDate < " & CStr(cutoffDate) & ". Rows affected = " & rowsAff
        DeleteOldTranRows = True
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
    ReDim fields(100)

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

Function ParseDateTimeFlexible(ByVal s)
    On Error Resume Next
    ParseDateTimeFlexible = Null

    s = Trim(CStr(s))
    If Len(s) = 0 Then Exit Function

    ' Convert MM-DD-YYYY to MM/DD/YYYY for SQL compatibility
    If InStr(s, "-") > 0 Then
        s = Replace(s, "-", "/")
    End If

    If IsDate(s) Then
        ParseDateTimeFlexible = CDate(s)
        Exit Function
    End If

    Err.Clear
    ParseDateTimeFlexible = Null
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

Sub AddParameterTimeStamp(cmdObj, dateVal)
    Dim param
    On Error Resume Next
    Set param = cmdObj.CreateParameter("", adDBTimeStamp, adParamInput, , dateVal)
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
    LogLine "ATM Reconciliation Import Completed: " & Now()
    LogLine "========================================="

    EndRun
    On Error GoTo 0
End Sub

Sub AddParameterDate(cmdObj, dateVal)
    Const adDBDate = 133
    Dim param
    On Error Resume Next
    Set param = cmdObj.CreateParameter("", adDBDate, adParamInput, , dateVal)
    cmdObj.Parameters.Append param
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

Sub AddParameterTime(cmdObj, timeVal)
    Const adDBTime = 133
    Dim param
    On Error Resume Next
    Set param = cmdObj.CreateParameter("", adDBTime, adParamInput, , timeVal)
    cmdObj.Parameters.Append param
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub