' ============================================================
' byREQUEST COB7_GL CSV -> SQL Staging Import (INI-driven)
' - Creates or replaces staging table with COB7_GL columns
' - CSV header is always first row
' - CSV has 18 columns with COB7_GL format headers
' ============================================================
Option Explicit

' ---------- CONFIG ----------
Dim INI_PATH: INI_PATH = "S:\CarolinaTrust\byREQUEST\Templates\COB7_GL.ini"

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

    LogLine "COB7_GL Import Started: " & Now()
    LogLine "INI_PATH = " & INI_PATH

    gDateColFromSpool = ExtractDateColFromSpoolName(Spoolfile.Name)
    LogLine "Spoolfile.Name = " & Spoolfile.Name
    MsgBox "DateCol = " & gDateColFromSpool

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

    ' --- Create or replace table ---
    If Not EnsureCOB7GLTable() Then
        LogLine "ERROR: Could not create/replace COB7_GL table."
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
    InsertCOB7GLRow line

    On Error GoTo 0
End Sub

' ============================================================
' Insert COB7_GL row - INPUT FORMAT (18 columns based on header)
' CSV columns (0-17) based on header:
' 0: Report, 1: Run_Date, 2: Run_Time, 3: GL_Number, 4: Branch, 
' 5: Debit, 6: Credit, 7: Sequence, 8: Voucher, 9: Teller, 
' 10: NBR_TYPE1, 11: Mbr_Nbr, 12: BR, 13: SFX, 14: PTyp, 
' 15: NBR_TYPE2, 16: MI_Code, 17: Addl_Detail
'
' MAPPING:
' CSV -> DB
' Report -> Report
' Run_Date -> Run_Date
' Run_Time -> Run_Time
' GL_Number -> GL_Number
' Branch -> Branch
' Debit -> Debit
' Credit -> Credit
' Sequence -> Sequence
' Voucher -> Voucher
' Teller -> Teller
' NBR_TYPE1 -> Tran_Nbr_Type1
' Mbr_Nbr -> Mbr_Nbr
' BR -> BR (NEW COLUMN)
' SFX -> Suffix
' PTyp -> Product_Type_Code
' NBR_TYPE2 -> Tran_Nbr_Type2
' MI_Code -> MI_Code
' Addl_Detail -> Addl_Detail
' ============================================================
' ============================================================
' Insert COB7_GL row - INPUT FORMAT (18 columns based on header)
' CSV columns (0-17) based on header:
' 0: Report, 1: Run_Date, 2: Run_Time, 3: GL_Number, 4: Branch, 
' 5: Debit, 6: Credit, 7: Sequence, 8: Voucher, 9: Teller, 
' 10: NBR_TYPE1, 11: Mbr_Nbr, 12: BR, 13: SFX, 14: PTyp, 
' 15: NBR_TYPE2, 16: MI_Code, 17: Addl_Detail
' ============================================================
Sub InsertCOB7GLRow(ByVal csvLine)
    On Error Resume Next

    Dim f, i
    f = SplitCSV(csvLine)
    
    ' DEBUG: Show column count
    Dim colCount
    colCount = UBound(f) + 1
    
    ' Expecting exactly 18 columns based on header
    If colCount < 18 Then
        LogLine "WARNING: Expected 18 columns, got " & colCount & ". Line: " & Left(csvLine, 100)
        ' Continue processing anyway, but log warning
    ElseIf colCount > 18 Then
        LogLine "WARNING: Expected 18 columns, got " & colCount & ". Line: " & Left(csvLine, 100)
    End If

    ' Trim all columns - even if they're just spaces, they'll become empty strings
    For i = 0 To colCount - 1
        f(i) = IntelligentTrim(f(i))
    Next

    ' HashKey for tracking (not for duplicate checking)
    Dim hashKey
    hashKey = LCase(gDateColFromSpool & "|" & f(3) & "|" & f(5) & "|" & f(6) & "|" & f(7))

    If IsDuplicate(hashKey) Then 
        fout.WriteLine "DEBUG: Duplicate skipped: " & hashKey
        Exit Sub
    End If

    Dim sql, cmd
    ' Remove Import_Timestamp from INSERT since it has DEFAULT GETDATE()
    sql = "INSERT INTO " & gTable2Part & " (" & _
          "[Report],[Run_Date],[Run_Time],[GL_Number],[Branch],[Debit],[Credit]," & _
          "[Sequence],[Voucher],[Teller],[Tran_Nbr_Type1],[Mbr_Nbr],[BR],[Suffix],[Product_Type_Code]," & _
          "[Tran_Nbr_Type2],[MI_Code],[Addl_Detail]," & _
          "[HashKey],[DateCol]" & _ 
          ") VALUES (" & _
          "?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?" & _ 
          ")"

    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = gCn
        .CommandText = sql
        .CommandType = 1

        ' Map CSV columns to COB7_GL database columns
        ' All 18 CSV columns are guaranteed to exist (may be empty/spaces)
        
        AddParameterString cmd, 50,  f(0)   ' Report -> Report
        AddParameterString cmd, 20,  f(1)   ' Run_Date -> Run_Date
        AddParameterString cmd, 20,  f(2)   ' Run_Time -> Run_Time
        AddParameterString cmd, 50,  f(3)   ' GL_Number -> GL_Number
        AddParameterString cmd, 50,  f(4)   ' Branch -> Branch
        AddParameterString cmd, 30,  f(5)   ' Debit -> Debit
        AddParameterString cmd, 30,  f(6)   ' Credit -> Credit
        AddParameterString cmd, 20,  f(7)   ' Sequence -> Sequence
        AddParameterString cmd, 50,  f(8)   ' Voucher -> Voucher
        AddParameterString cmd, 50,  f(9)   ' Teller -> Teller
        AddParameterString cmd, 50,  f(10)  ' NBR_TYPE1 -> Tran_Nbr_Type1
        AddParameterString cmd, 50,  f(11)  ' Mbr_Nbr -> Mbr_Nbr
        AddParameterString cmd, 20,  f(12)  ' BR -> BR
        AddParameterString cmd, 20,  f(13)  ' SFX -> Suffix
        AddParameterString cmd, 20,  f(14)  ' PTyp -> Product_Type_Code
        AddParameterString cmd, 50,  f(15)  ' NBR_TYPE2 -> Tran_Nbr_Type2
        AddParameterString cmd, 50,  f(16)  ' MI_Code -> MI_Code
        AddParameterString cmd, 255, f(17)  ' Addl_Detail -> Addl_Detail

        ' Metadata columns
        AddParameterString cmd, 255, hashKey  ' HashKey for tracking
        AddParameterString cmd, 30,  gDateColFromSpool  ' Date from spool filename
        ' Import_Timestamp is not included - will use DEFAULT GETDATE() from SQL
    End With

    Err.Clear
    cmd.Execute
    If Err.Number <> 0 Then
        LogLine "ERROR: Insert failed (" & Err.Number & ") " & Err.Description
        LogLine "Line: " & Left(csvLine, 150)
        LogLine "HashKey: " & hashKey
        LogLine "SQL has 20 parameters, CSV has " & colCount & " columns"
        ' Log the actual SQL for debugging
        LogLine "SQL: " & Left(sql, 200)
        Err.Clear
    Else

            LogLine "DEBUG: Inserted " & counter & " records"

    End If

    Set cmd = Nothing
    On Error GoTo 0
End Sub

' ============================================================
' EnsureCOB7GLTable - Create or replace table for COB7_GL data
' ============================================================
Function EnsureCOB7GLTable()
    On Error Resume Next
    EnsureCOB7GLTable = False

    Dim sql

    ' ------------------------------------------------------------
    ' CREATE TABLE only if it does not exist  (APPEND mode forever)
    ' ------------------------------------------------------------
    sql = "IF OBJECT_ID(N'" & SqlQuote(gObjIdName) & "', N'U') IS NULL " & _
          "BEGIN " & _
          "  CREATE TABLE " & gTable2Part & " (" & _
          "    [Report]              NVARCHAR(50)  NULL," & _
          "    [Run_Date]            NVARCHAR(20)  NULL," & _
          "    [Run_Time]            NVARCHAR(20)  NULL," & _
          "    [GL_Number]           NVARCHAR(50)  NULL," & _
          "    [Branch]              NVARCHAR(50)  NULL," & _
          "    [Debit]               NVARCHAR(30)  NULL," & _
          "    [Credit]              NVARCHAR(30)  NULL," & _
          "    [Sequence]            NVARCHAR(20)  NULL," & _
          "    [Voucher]             NVARCHAR(50)  NULL," & _
          "    [Teller]              NVARCHAR(50)  NULL," & _
          "    [Tran_Nbr_Type1]      NVARCHAR(50)  NULL," & _
          "    [Mbr_Nbr]             NVARCHAR(50)  NULL," & _
          "    [BR]                  NVARCHAR(20)  NULL," & _
          "    [Suffix]              NVARCHAR(20)  NULL," & _
          "    [Product_Type_Code]   NVARCHAR(20)  NULL," & _
          "    [Tran_Nbr_Type2]      NVARCHAR(20)  NULL," & _
          "    [MI_Code]             NVARCHAR(50)  NULL," & _
          "    [Addl_Detail]         NVARCHAR(255) NULL," & _
          "    [HashKey]             NVARCHAR(255) NULL," & _
          "    [DateCol]             NVARCHAR(30)  NULL," & _
          "    [Import_Timestamp]    DATETIME DEFAULT GETDATE()" & _
          "  ) " & _
          "END"

    If Not ExecNonQuery(sql) Then Exit Function

    ' ------------------------------------------------------------
    ' Create indexes ONLY IF missing
    ' ------------------------------------------------------------
    sql = "IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_COB7_GL_HashKey' AND object_id = OBJECT_ID(N'" & SqlQuote(gObjIdName) & "')) " & _
          "CREATE INDEX IX_COB7_GL_HashKey ON " & gTable2Part & "([HashKey])"
    ExecNonQuery sql

    sql = "IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_COB7_GL_DateCol' AND object_id = OBJECT_ID(N'" & SqlQuote(gObjIdName) & "')) " & _
          "CREATE INDEX IX_COB7_GL_DateCol ON " & gTable2Part & "([DateCol])"
    ExecNonQuery sql

    sql = "IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_COB7_GL_GL_Number' AND object_id = OBJECT_ID(N'" & SqlQuote(gObjIdName) & "')) " & _
          "CREATE INDEX IX_COB7_GL_GL_Number ON " & gTable2Part & "([GL_Number])"
    ExecNonQuery sql

    sql = "IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_COB7_GL_Run_Date' AND object_id = OBJECT_ID(N'" & SqlQuote(gObjIdName) & "')) " & _
          "CREATE INDEX IX_COB7_GL_Run_Date ON " & gTable2Part & "([Run_Date])"
    ExecNonQuery sql

    sql = "IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_COB7_GL_Voucher' AND object_id = OBJECT_ID(N'" & SqlQuote(gObjIdName) & "')) " & _
          "CREATE INDEX IX_COB7_GL_Voucher ON " & gTable2Part & "([Voucher])"
    ExecNonQuery sql

    EnsureCOB7GLTable = True
    On Error GoTo 0
    LogLine "SUCCESS: COB7_GL table ensured (append mode) + indexes ensured"
End Function



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
	MsgBox  ExtractDateColFromSpoolName
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
    LogLine "COB7_GL Import Completed: " & Now()
    LogLine "========================================="
    EndRun
    On Error GoTo 0
End Sub