' ============================================================
' byREQUEST ACH EPSA CSV -> SQL Staging Import (INI-driven)
' - Creates staging table with all required columns
' - CSV header is always first row
' - CSV has 31 columns with input format headers
' ============================================================
Option Explicit

' ---------- CONFIG ----------
Dim INI_PATH: INI_PATH = "S:\CarolinaTrust\byREQUEST\Templates\ACH_Trans.ini"

' ---------- GLOBALS ----------
Dim gCfg, gCn
Dim gConnStr, gFullTableName
Dim gTable2Part, gObjIdName
Dim gDateColFromSpool
Dim gHasProcessedHeader

Dim srec

' ============================================================
' StartDoc
' ============================================================
Sub StartDoc
    On Error Resume Next

    gHasProcessedHeader = False

    LogLine "ACH EPSA Import Started: " & Now()
    LogLine "INI_PATH = " & INI_PATH

    gDateColFromSpool = ExtractDateColFromSpoolName(Spoolfile.Name)
    LogLine "Spoolfile.Name = " & Spoolfile.Name
    LogLine "DateCol = " & gDateColFromSpool

    ' --- Load INI ---
    Set gCfg = ReadIniSafe(INI_PATH)
    If gCfg Is Nothing Then
        LogLine "ERROR: INI not found/unreadable: " & INI_PATH
        EndRun
        Exit Sub
    End If

    gConnStr = GetRequired(gCfg, "Database.ConnStr")
    gFullTableName = GetRequired(gCfg, "Database.TableName")
    If Len(gConnStr) = 0 Or Len(gFullTableName) = 0 Then
        EndRun
        Exit Sub
    End If

    BuildTableNames gFullTableName, gTable2Part, gObjIdName
    If Len(gTable2Part) = 0 Or Len(gObjIdName) = 0 Then
        LogLine "ERROR: Could not parse TableName: " & gFullTableName
        EndRun
        Exit Sub
    End If

    ' --- DB connect ---
    Set gCn = CreateObject("ADODB.Connection")
    Err.Clear
    gCn.Open gConnStr
    If Err.Number <> 0 Then
        LogLine "ERROR: DB open failed (" & Err.Number & ") " & Err.Description
        Err.Clear
        EndRun
        Exit Sub
    End If
    On Error GoTo 0

    LogLine "SUCCESS: Connected"
    LogLine "Target table = " & gTable2Part

    ' --- Ensure table exists with all required columns ---
    If Not EnsureStagingTable() Then
        LogLine "ERROR: Could not ensure staging table."
        EndRun
        Exit Sub
    End If

End Sub

' ============================================================
' ProcessLine - Process each line
' ============================================================
Sub ProcessLine
    On Error Resume Next

    Dim line
    line = srec
    If Trim(line) = "" Then Exit Sub

    ' First line is always the header - skip it
    If Not gHasProcessedHeader Then
        gHasProcessedHeader = True
        LogLine "CSV header detected and skipped"
        LogLine "DEBUG Header sample: " & Left(line, 150)
        Exit Sub
    End If

    ' Process data rows
    InsertMappedRow line

    On Error GoTo 0
End Sub

' ============================================================
' Insert mapped row - INPUT FORMAT (31 columns)
' CSV columns (0-30):
' 0: Report, 1: Post_Date, 2: Post_Time, 3: Settlement_Date, 
' 4: Company_Name, 5: ACH_Batch_ID, 6: PCTL, 7: LCTL, 
' 8: Post_Seq, 9: File_Post_Seq, 10: TransCode, 11: Description, 
' 12: ST, 13: TYPE, 14: EFF DATE, 15: AMOUNT, 16: CHECK#, 
' 17: SDC#, 18: MEMBER#, 19: S/L, 20: NAME, 21: FEE, 
' 22: RR, 23: TC, 24: RDFI IDENT, 25: DFI ACCOUNT#, 
' 26: INDIVIDUAL ID#, 27: INDIVIDUAL NAME, 28: DD, 
' 29: TRACE#, 30: MESSAGE
' ============================================================
Sub InsertMappedRow(ByVal csvLine)
    On Error Resume Next

    Dim f, i
    f = SplitCSV(csvLine)
    
    ' DEBUG: Show column count
    Dim colCount
    colCount = UBound(f) + 1
    LogLine "DEBUG: CSV columns found: " & colCount
    
    If colCount < 31 Then
        LogLine "SKIP: Not enough CSV fields (need 31). Got " & colCount
        If colCount > 0 Then
            LogLine "DEBUG First few columns:"
            For i = 0 To UBound(f)
                If i <= 5 Then
                    LogLine "  [" & i & "]: " & f(i)
                End If
            Next
        End If
        Exit Sub
    End If

    For i = 0 To 30
        f(i) = IntelligentTrim(f(i))
    Next

    ' HashKey for duplicate detection
    Dim hashKey
	hashKey = LCase(f(0) & "|" & f(18) & "|" & f(10) & "|" & f(14) & "|" & f(29) & "|" & f(27) & "|" & f(12))

    ' If IsDuplicate(hashKey) Then 
        ' LogLine "DEBUG: Duplicate skipped: " & hashKey
        ' Exit Sub
    ' End If

    Dim sql, cmd
    sql = "INSERT INTO " & gTable2Part & " (" & _
          "[Report],[Post_Date],[Post_Time],[Settlement_Date],[Company_Name],[ACH_Batch_ID],[PCTL],[LCTL]," & _
          "[Status],[Tran_Type],[Tran_Code],[Effective_Date],[Amount],[Member_Nbr],[Acct],[Mbr_Name],[Fee]," & _
          "[ACH_Ind_Name],[DD],[Trace_Nbr],[Description],[Type],[Post_Seq],[File_Post_Seq]," & _
          "[Source_Name],[Batch],[Check_Number],[SDC_Number],[S_L],[RR],[RDFI_ID],[File_Acc_Number],[Individual_ID],[MESSAGE]," & _
          "[HashKey],[DateCol],[Source_Message]" & _
          ") VALUES (" & _
          "?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?" & _
          ")"

    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = gCn
        .CommandText = sql
        .CommandType = 1

        ' Map CSV columns to database columns
        ' Core table columns
        AddParameterString cmd, 10,  f(0)   ' Report -> Report
        AddParameterString cmd, 20,  f(1)   ' Post_Date -> Post_Date
        AddParameterString cmd, 20,  f(2)   ' Post_Time -> Post_Time
        AddParameterString cmd, 20,  f(3)   ' Settlement_Date -> Settlement_Date
        AddParameterString cmd, 100, f(4)   ' Company_Name -> Company_Name
        AddParameterString cmd, 20,  f(5)   ' ACH_Batch_ID -> ACH_Batch_ID
        AddParameterString cmd, 20,  f(6)   ' PCTL -> PCTL
        AddParameterString cmd, 20,  f(7)   ' LCTL -> LCTL
        AddParameterString cmd, 5,   f(12)  ' ST -> Status
        AddParameterString cmd, 30,  f(13)  ' TYPE -> Tran_Type
        AddParameterString cmd, 10,  f(10)  ' TransCode -> Tran_Code
        AddParameterString cmd, 20,  f(14)  ' EFF DATE -> Effective_Date
        AddParameterString cmd, 30,  f(15)  ' AMOUNT -> Amount
        AddParameterString cmd, 30,  f(18)  ' MEMBER# -> Member_Nbr
        AddParameterString cmd, 10,  f(19)  ' S/L -> Acct
        AddParameterString cmd, 100, f(20)  ' NAME -> Mbr_Name
        AddParameterString cmd, 30,  f(21)  ' FEE -> Fee
        AddParameterString cmd, 100, f(27)  ' INDIVIDUAL NAME -> ACH_Ind_Name
        AddParameterString cmd, 5,   f(28)  ' DD -> DD
        AddParameterString cmd, 40,  f(29)  ' TRACE# -> Trace_Nbr
        AddParameterString cmd, 255, f(11)  ' Description -> Description
        AddParameterString cmd, 5,   f(23)  ' TC -> Type (use TC column for Type)
        AddParameterString cmd, 5,   f(8)   ' Post_Seq -> Post_Seq
        AddParameterString cmd, 5,   f(9)   ' File_Post_Seq -> File_Post_Seq

        ' Additional columns (store all CSV data for future use)
        AddParameterString cmd, 100, f(4)   ' Company_Name -> Source_Name (same as Company_Name)
        AddParameterString cmd, 20,  f(5)   ' ACH_Batch_ID -> Batch (same as ACH_Batch_ID)
        AddParameterString cmd, 20,  f(16)  ' CHECK# -> Check_Number
        AddParameterString cmd, 20,  f(17)  ' SDC# -> SDC_Number
        AddParameterString cmd, 10,  f(19)  ' S/L -> S_L
        AddParameterString cmd, 10,  f(22)  ' RR -> RR
        AddParameterString cmd, 20,  f(24)  ' RDFI IDENT -> RDFI_ID
        AddParameterString cmd, 20,  f(25)  ' DFI ACCOUNT# -> File_Acc_Number
        AddParameterString cmd, 30,  f(26)  ' INDIVIDUAL ID# -> Individual_ID
        AddParameterString cmd, 255, f(30)  ' MESSAGE -> MESSAGE

        ' Metadata columns
        AddParameterString cmd, 255, hashKey
        AddParameterString cmd, 30,  gDateColFromSpool
        AddParameterString cmd, 255, f(30)  ' MESSAGE -> Source_Message
    End With

    Err.Clear
    cmd.Execute
    If Err.Number <> 0 Then
        LogLine "ERROR: Insert failed (" & Err.Number & ") " & Err.Description
        LogLine "Line: " & Left(csvLine, 150)
        LogLine "HashKey: " & hashKey
        LogLine "SQL has 37 parameters, CSV has " & colCount & " columns"
        Err.Clear
    Else
        LogLine "DEBUG: Successfully inserted: " & hashKey
    End If

    Set cmd = Nothing
    On Error GoTo 0
End Sub

' ============================================================
' EnsureStagingTable - Create table with all columns
' ============================================================
Function EnsureStagingTable()
    On Error Resume Next
    EnsureStagingTable = False

    Dim sql

    sql = "IF OBJECT_ID(N'" & SqlQuote(gObjIdName) & "', N'U') IS NULL " & _
          "BEGIN " & _
          "  CREATE TABLE " & gTable2Part & " (" & _
          "    [Report]           NVARCHAR(10)  NULL," & _
          "    [Post_Date]        NVARCHAR(20)  NULL," & _
          "    [Post_Time]        NVARCHAR(20)  NULL," & _
          "    [Settlement_Date]  NVARCHAR(20)  NULL," & _
          "    [Company_Name]     NVARCHAR(100) NULL," & _
          "    [ACH_Batch_ID]     NVARCHAR(20)  NULL," & _
          "    [PCTL]             NVARCHAR(20)  NULL," & _
          "    [LCTL]             NVARCHAR(20)  NULL," & _
          "    [Status]           NVARCHAR(5)   NULL," & _
          "    [Tran_Type]        NVARCHAR(30)  NULL," & _
          "    [Tran_Code]        NVARCHAR(10)  NULL," & _
          "    [Effective_Date]   NVARCHAR(20)  NULL," & _
          "    [Amount]           NVARCHAR(30)  NULL," & _
          "    [Member_Nbr]       NVARCHAR(30)  NULL," & _
          "    [Acct]             NVARCHAR(10)  NULL," & _
          "    [Mbr_Name]         NVARCHAR(100) NULL," & _
          "    [Fee]              NVARCHAR(30)  NULL," & _
          "    [ACH_Ind_Name]     NVARCHAR(100) NULL," & _
          "    [DD]               NVARCHAR(5)   NULL," & _
          "    [Trace_Nbr]        NVARCHAR(40)  NULL," & _
          "    [Description]      NVARCHAR(255) NULL," & _
          "    [Type]             NVARCHAR(5)   NULL," & _
          "    [Post_Seq]         NVARCHAR(5)   NULL," & _
          "    [File_Post_Seq]    NVARCHAR(5)   NULL," & _
          "    [Source_Name]      NVARCHAR(100) NULL," & _
          "    [Batch]            NVARCHAR(20)  NULL," & _
          "    [Check_Number]     NVARCHAR(20)  NULL," & _
          "    [SDC_Number]       NVARCHAR(20)  NULL," & _
          "    [S_L]              NVARCHAR(10)  NULL," & _
          "    [RR]               NVARCHAR(10)  NULL," & _
          "    [RDFI_ID]          NVARCHAR(20)  NULL," & _
          "    [File_Acc_Number]  NVARCHAR(20)  NULL," & _
          "    [Individual_ID]    NVARCHAR(30)  NULL," & _
          "    [MESSAGE]          NVARCHAR(255) NULL," & _
          "    [HashKey]          NVARCHAR(255) NULL," & _
          "    [DateCol]          NVARCHAR(30)  NULL," & _
          "    [Source_Message]   NVARCHAR(255) NULL," & _
          "    [Import_Timestamp] DATETIME DEFAULT GETDATE()" & _
          "  ) " & _
          "END"

    If Not ExecNonQuery(sql) Then Exit Function

    ' Index (if missing)
    sql = "IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_ACH_EPSA_Staging_HashKey' AND object_id = OBJECT_ID(N'" & SqlQuote(gObjIdName) & "')) " & _
          "BEGIN CREATE INDEX IX_ACH_EPSA_Staging_HashKey ON " & gTable2Part & "([HashKey]) END"

    If Not ExecNonQuery(sql) Then Exit Function

    EnsureStagingTable = True
    On Error GoTo 0
End Function

' ============================================================
' Helper Functions
' ============================================================

Function IsDuplicate(ByVal hashKey)
    On Error Resume Next
    IsDuplicate = False

    Dim sql, cmd, rs, cnt
    sql = "SELECT COUNT(*) AS Cnt FROM " & gTable2Part & " WHERE [HashKey] = ?"

    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = gCn
        .CommandText = sql
        .CommandType = 1
        .Parameters.Append .CreateParameter("", 202, 1, 255, CStr(hashKey))
    End With

    Err.Clear
    Set rs = cmd.Execute
    If Err.Number <> 0 Then
        LogLine "IsDuplicate ERROR (" & Err.Number & "): " & Err.Description
        Err.Clear
        IsDuplicate = False
    Else
        cnt = 0
        If Not rs.EOF Then cnt = rs("Cnt")
        IsDuplicate = (cnt > 0)
    End If

    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set cmd = Nothing
    On Error GoTo 0
End Function

Function ExecNonQuery(ByVal sql)
    On Error Resume Next
    ExecNonQuery = False

    Dim cmd
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = gCn
    cmd.CommandType = 1
    cmd.CommandText = sql

    Err.Clear
    cmd.Execute

    If Err.Number <> 0 Then
        LogLine "SQL ERROR (" & Err.Number & "): " & Err.Description
        LogLine "SQL: " & Left(sql, 260)
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
        tableName  = Trim(parts(1))
    Else
        schemaName = Trim(parts(UBound(parts)-1))
        tableName  = Trim(parts(UBound(parts)))
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
        If Len(line)=0 Or Left(line,1)=";" Then
            ' skip
        ElseIf Left(line,1)="[" And Right(line,1)="]" Then
            sect = Mid(line,2,Len(line)-2)
        Else
            p = InStr(line,"=")
            If p>0 Then d(sect & "." & Trim(Left(line,p-1))) = Trim(Mid(line,p+1))
        End If
    Loop

    ts.Close
    Set ReadIniSafe = d
    On Error GoTo 0
End Function

Function GetRequired(cfg, key)
    If Not (cfg Is Nothing) And cfg.Exists(key) And Len(cfg(key))>0 Then
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
    
    ' Remove surrounding quotes and trim
    result = Trim(result)
    If Len(result) >= 2 Then
        If Left(result, 1) = """" And Right(result, 1) = """" Then
            result = Mid(result, 2, Len(result) - 2)
        End If
    End If
    
    IntelligentTrim = Trim(result)
End Function

Function SplitCSV(line)
    On Error Resume Next
    Dim fields(), fieldCount, i, inQuotes, fieldValue, ch
    ReDim fields(50) ' Allow up to 50 columns
    fieldCount = 0
    inQuotes = False
    fieldValue = ""
    
    For i = 1 To Len(line)
        ch = Mid(line, i, 1)
        If ch = """" Then
            inQuotes = Not inQuotes
        ElseIf ch = "," And Not inQuotes Then
            fields(fieldCount) = Trim(fieldValue)
            fieldValue = ""
            fieldCount = fieldCount + 1
        Else
            fieldValue = fieldValue & ch
        End If
    Next
    
    ' Don't forget the last field
    fields(fieldCount) = Trim(fieldValue)
    fieldCount = fieldCount + 1
    
    ReDim Preserve fields(fieldCount - 1)
    SplitCSV = fields
    On Error GoTo 0
End Function

Sub AddParameterString(cmd, paramSize, paramValue)
    Dim param, val
    On Error Resume Next
    If paramSize <= 0 Then paramSize = 50
    If IsNull(paramValue) Then 
        val = "" 
    Else 
        val = CStr(paramValue)
    End If
    Set param = cmd.CreateParameter("", 202, 1, paramSize, val) ' NVARCHAR
    cmd.Parameters.Append param
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

Function ExtractDateColFromSpoolName(fname)
    Dim pos, tmp, dotPos
    ExtractDateColFromSpoolName = ""
    If IsNull(fname) Or fname = "" Then Exit Function

    pos = InStrRev(fname, "_")
    If pos > 0 Then tmp = Mid(fname, pos + 1) Else tmp = fname

    dotPos = InStr(tmp, ".")
    If dotPos > 0 Then tmp = Left(tmp, dotPos - 1)

    ExtractDateColFromSpoolName = Trim(tmp)
End Function

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

' ============================================================
' CloseDoc
' ============================================================
Sub CloseDoc
    On Error Resume Next
    LogLine "ACH EPSA Import Completed: " & Now()
    LogLine "========================================="
    EndRun
    On Error GoTo 0
End Sub