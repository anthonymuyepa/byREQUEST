' ============================================================
' byREQUEST COB7_GL CSV -> SQL Staging Import (INI-driven)
'   * HEAP TABLE - no clustered index, no additional indexes
'   * Daily full replace: DELETE WHERE Run_Date = ... then INSERT
'   * Run_Date is real DATE type NOT NULL
'   * Date comes from filename (normalized: 12-17-25 → 2025-12-17 etc.)
'   * Simple auto-increment ID
' ============================================================
Option Explicit

' ---------- CONFIG ----------
Dim INI_PATH: INI_PATH = "S:\CarolinaTrust\byREQUEST\Templates\COB7_GL.ini"

' ---------- GLOBALS ----------
Dim gCfg, gCn
Dim gConnStr, gFullTableName
Dim gTable2Part, gObjIdName
Dim gRunDate                   ' Date Serial
Dim gHasProcessedHeader
Dim srec
Dim counter : counter = 0
Dim gAbortRun : gAbortRun = False

' ============================================================
' StartDoc
' ============================================================
Sub StartDoc
    On Error Resume Next
    gHasProcessedHeader = False
    counter = 0
    
    LogLine "COB7_GL Daily Full-Replace Import Started: " & Now()
    LogLine "INI_PATH = " & INI_PATH
    
    gRunDate = ExtractRunDate(Spoolfile.Name)
    LogLine "Spoolfile.Name = " & Spoolfile.Name
    LogLine "Raw Run_Date from filename = " & gRunDate
    

	If IsNull(gRunDate) Then
		LogLine "CRITICAL ERROR: Invalid spool filename format."
		LogLine "Expected: COB7_MMDDYY_<suffix>"
		LogLine "Found: " & Spoolfile.Name
		gAbortRun = True
		EndRun
		Exit Sub
	End If
    LogLine "Normalized Run_Date (for DB) = " & gRunDate
    
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
    If Len(gTable2Part) = 0 Then
        LogLine "ERROR: Could not parse TableName: " & gFullTableName
        EndRun
        Exit Sub
    End If
    
    ' --- DB connect ---
    Set gCn = CreateObject("ADODB.Connection")
    gCn.Open gConnStr
    If Err.Number <> 0 Then
        LogLine "ERROR: DB open failed (" & Err.Number & ") " & Err.Description
        Err.Clear
        EndRun
        Exit Sub
    End If
    On Error GoTo 0
    
    LogLine "SUCCESS: Connected to database"
    LogLine "Target table = " & gTable2Part
    
    ' --- Create/ensure heap table ---
    If Not EnsureCOB7GLHeapTable() Then
        LogLine "ERROR: Could not create/verify COB7_GL heap table."
        EndRun
        Exit Sub
    End If
    
    ' --- Delete previous data for this Run_Date ---
    If Not DeleteExistingRunDate(gRunDate) Then
        LogLine "WARNING: Delete for Run_Date failed - continuing anyway..."
    End If
    
    On Error Resume Next
End Sub

' ============================================================
' ProcessLine
' ============================================================
Sub ProcessLine
    On Error Resume Next
	If gAbortRun Then Exit Sub

    Dim line : line = srec
    If Trim(line) = "" Then Exit Sub
    
    ' First line = header → skip
    If Not gHasProcessedHeader Then
        gHasProcessedHeader = True
        LogLine "CSV header detected and skipped"
        Exit Sub
    End If
    
    InsertCOB7GLRow line
    counter = counter + 1
    
    On Error GoTo 0
End Sub

' ============================================================
' Insert one row
' ============================================================
Sub InsertCOB7GLRow(ByVal csvLine)
    On Error Resume Next
    Dim f : f = SplitCSV(csvLine)
    Dim colCount : colCount = UBound(f) + 1
    
    If colCount < 18 Then
        LogLine "WARNING: Expected at least 18 columns, got " & colCount & " - skipping"
        Exit Sub
    End If
    
    Dim i
    For i = 0 To colCount - 1
        f(i) = IntelligentTrim(f(i))
    Next
	

	' Skip spacer/blank rows: require first 3 fields to contain something
	If Len(f(0)) = 0 And Len(f(1)) = 0 And Len(f(2)) = 0 Then Exit Sub

	' Parse CSV date in MM-DD-YY
	Dim rowDate
	rowDate = ParseRowDate_MMDDYY(Replace(f(1),"-",""))

	If IsNull(rowDate) Or DateDiff("d", rowDate, gRunDate) <> 0 Then
		LogLine "WARNING: CSV date does not match filename date. CSV=" & f(1) & _
				" FILE=" & Spoolfile.Name & " - skipping row."
		LogLine "DEBUG rowDate=" & CStr(rowDate) & " gRunDate=" & CStr(gRunDate)
		Exit Sub
	End If



	
    
    Dim sql
    sql = "INSERT INTO " & gTable2Part & " (" & _
          "[Run_Date],[Report],[Run_Time],[GL_Number],[Branch]," & _
          "[Debit],[Credit],[Sequence],[Voucher],[Teller]," & _
          "[Tran_Nbr_Type1],[Mbr_Nbr],[BR],[Suffix],[Product_Type_Code]," & _
          "[Tran_Nbr_Type2],[MI_Code],[Addl_Detail]" & _
          ") VALUES (" & _
          "?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?" & _
          ")"

    Dim cmd : Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = gCn
        .CommandText = sql
        .CommandType = 1
        
        ' Real DATE parameter
        .Parameters.Append .CreateParameter("RunDate", 133, 1, , gRunDate) 
        
        AddParameterString cmd, 50,  f(0)   ' Report
        AddParameterString cmd, 20,  f(2)   ' Run_Time
        AddParameterString cmd, 50,  f(3)   ' GL_Number
        AddParameterString cmd, 50,  f(4)   ' Branch
        AddParameterString cmd, 30,  f(5)   ' Debit
        AddParameterString cmd, 30,  f(6)   ' Credit
        AddParameterString cmd, 20,  f(7)   ' Sequence
        AddParameterString cmd, 50,  f(8)   ' Voucher
        AddParameterString cmd, 50,  f(9)   ' Teller
        AddParameterString cmd, 50,  f(10)  ' NBR_TYPE1/Tran_Nbr_Type1
        AddParameterString cmd, 50,  f(11)  ' Mbr_Nbr
        AddParameterString cmd, 20,  f(12)  ' BR
        AddParameterString cmd, 20,  f(13)  ' SFX / Suffix
        AddParameterString cmd, 20,  f(14)  ' PTyp / Product_Type_Code
        AddParameterString cmd, 50,  f(15)  ' NBR_TYPE2 / Tran_Nbr_Type2
        AddParameterString cmd, 50,  f(16)  ' MI_Code
        AddParameterString cmd, 255, f(17)  ' Addl_Detail
    End With
    
    cmd.Execute
    
    If Err.Number <> 0 Then
        LogLine "INSERT ERROR: " & Err.Number & " - " & Err.Description
        LogLine "Line sample: " & Left(csvLine, 180)
        Err.Clear
    End If
    
    Set cmd = Nothing
    On Error GoTo 0
End Sub

' ============================================================
' Create minimal HEAP table - Run_Date is real DATE NOT NULL
' ============================================================
Function EnsureCOB7GLHeapTable()
    On Error Resume Next
    EnsureCOB7GLHeapTable = False
    
    Dim sql
    sql = "IF OBJECT_ID(N'" & SqlQuote(gObjIdName) & "', N'U') IS NULL " & _
          "BEGIN " & _
          " CREATE TABLE " & gTable2Part & " (" & _
          "   [ID]              BIGINT          IDENTITY(1,1) NOT NULL," & _
          "   [Run_Date]        DATE            NOT NULL," & _
          "   [Report]          NVARCHAR(50)        NULL," & _
          "   [Run_Time]        NVARCHAR(20)        NULL," & _
          "   [GL_Number]       NVARCHAR(50)        NULL," & _
          "   [Branch]          NVARCHAR(50)        NULL," & _
          "   [Debit]           NVARCHAR(30)        NULL," & _
          "   [Credit]          NVARCHAR(30)        NULL," & _
          "   [Sequence]        NVARCHAR(20)        NULL," & _
          "   [Voucher]         NVARCHAR(50)        NULL," & _
          "   [Teller]          NVARCHAR(50)        NULL," & _
          "   [Tran_Nbr_Type1]  NVARCHAR(50)        NULL," & _
          "   [Mbr_Nbr]         NVARCHAR(50)        NULL," & _
          "   [BR]              NVARCHAR(20)        NULL," & _
          "   [Suffix]          NVARCHAR(20)        NULL," & _
          "   [Product_Type_Code] NVARCHAR(20)     NULL," & _
          "   [Tran_Nbr_Type2]  NVARCHAR(50)        NULL," & _
          "   [MI_Code]         NVARCHAR(50)        NULL," & _
          "   [Addl_Detail]     NVARCHAR(255)       NULL," & _
          "   [Import_Timestamp] DATETIME         DEFAULT GETDATE()" & _
          " ) " & _
          "END"
    
    If Not ExecNonQuery(sql) Then Exit Function
    
    LogLine "SUCCESS: COB7_GL heap table ensured (Run_Date is DATE NOT NULL, no indexes)"
    EnsureCOB7GLHeapTable = True
    On Error GoTo 0
End Function

' ============================================================
' Delete all rows for this Run_Date before importing
' ============================================================
Function DeleteExistingRunDate(normDate)
    On Error Resume Next
    DeleteExistingRunDate = False
    
    If Len(Trim(normDate)) = 0 Then
        LogLine "WARNING: Cannot delete - Run_Date is empty!"
        Exit Function
    End If
    
    Dim sql
    sql = "DELETE FROM " & gTable2Part & " WHERE Run_Date = ?"
    
    Dim cmd : Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = gCn
    cmd.CommandText = sql
    cmd.CommandType = 1
    cmd.Parameters.Append cmd.CreateParameter("RunDate", 133, 1, , gRunDate)
    
    cmd.Execute
    
    If Err.Number = 0 Then
        LogLine "Deleted previous data for Run_Date = " & normDate
        DeleteExistingRunDate = True
    Else
        LogLine "WARNING: Delete failed (" & Err.Number & ") " & Err.Description
        Err.Clear
    End If
    
    Set cmd = Nothing
    On Error GoTo 0
End Function




' ============================================================
' Existing helper functions (kept unchanged)
' ============================================================

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
        schemaName = Trim(parts(UBound(parts)-1))
        tableName = Trim(parts(UBound(parts)))
    End If
    If Len(tableName) = 0 Then Exit Sub
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
    Dim result : result = CStr(value)
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
    ReDim fields(50)
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
    If IsNull(paramValue) Then val = "" Else val = CStr(paramValue)
    Set param = cmd.CreateParameter("", 202, 1, paramSize, val) ' NVARCHAR
    cmd.Parameters.Append param
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

' Extracts MMDDYY from: COB7_121725_89374
' Returns: VB Date or Null
Function ExtractRunDate(ByVal fname)
    Dim firstPos, secondPos, token
    Dim mm, dd, yy, yyyy
	
    ExtractRunDate = Null

    If IsNull(fname) Or Trim(fname) = "" Then Exit Function


    firstPos = InStr(fname, "_")
    If firstPos = 0 Then Exit Function

    secondPos = InStr(firstPos + 1, fname, "_")
    If secondPos = 0 Then Exit Function

    token = Mid(fname, firstPos + 1, secondPos - firstPos - 1)
    token = Trim(token)

    ' Must be exactly MMDDYY
    If Len(token) <> 6 Then Exit Function
    If Not IsNumeric(token) Then Exit Function

    mm = CInt(Left(token, 2))
    dd = CInt(Mid(token, 3, 2))
    yy = CInt(Right(token, 2))

    If mm < 1 Or mm > 12 Then Exit Function
    If dd < 1 Or dd > 31 Then Exit Function

    ' Year rule: 00–49 → 2000s, 50–99 → 1900s
    If yy >= 50 Then
        yyyy = 1900 + yy
    Else
        yyyy = 2000 + yy
    End If

    ExtractRunDate = DateSerial(yyyy, mm, dd)
End Function

Function ParseRowDate_MMDDYY(ByVal token)
    ParseRowDate_MMDDYY = Null
    token = Trim(CStr(token))
    If Len(token) <> 6 Or Not IsNumeric(token) Then Exit Function

    Dim mm, dd, yy, yyyy
    mm = CInt(Left(token, 2))
    dd = CInt(Mid(token, 3, 2))
    yy = CInt(Right(token, 2))

    If mm < 1 Or mm > 12 Then Exit Function
    If dd < 1 Or dd > 31 Then Exit Function

    If yy >= 50 Then yyyy = 1900 + yy Else yyyy = 2000 + yy
    ParseRowDate_MMDDYY = DateSerial(yyyy, mm, dd)
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
	If gAbortRun Then Exit Sub

    LogLine "Inserted " & counter & " records for Run_Date " & gRunDate
    LogLine "COB7_GL Daily Import Completed: " & Now()
    LogLine "========================================="
    EndRun
    On Error GoTo 0
End Sub