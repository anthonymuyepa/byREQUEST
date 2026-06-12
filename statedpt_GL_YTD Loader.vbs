Option Explicit

' ============================================================
' GL Report Staging Import 
' - Same logic as your ERTR loader:
'   * Read INI (ConnStr, TableName, ReportType)
'   * Create table if missing
'   * Duplicate protection: abort if SourceFile already loaded
'   * Skip header row
'   * Insert rows via parameterized ADO Command
'
' REPORT TYPES (INI: GL.ReportType):
'   GLBudgetReport_MTD
'   GLBudgetReport_YTD
'   GLIncomeStatement
'   GLBalanceSheet
'
' EXPECTED CSV COLUMN ORDER (minimum)
'   0 EffectiveDate     (MM/DD/YYYY or YYYY-MM-DD)
'   1 GLAccount
'   2 GLAccountName
'   3 MajorSection
'   4 Amount            (MTD or YTD depending on report)
'   5 Actual
'   6 Budget
'   7 Variance
'   8 VarPct
'   9 LineType          (optional; IncomeStatement/BalanceSheet only)
'  10 LineOrder         (optional; IncomeStatement/BalanceSheet only)
'
' Notes:
' - If your CSV does NOT include LineType/LineOrder, it still loads (NULLs).
' - SourceFile is always Spoolfile.Name (not from CSV).
' ============================================================

' ---------- ADO CONSTANTS ----------
Const adCmdText      = 1
Const adParamInput   = 1
Const adVarWChar     = 202
Const adDBDate       = 133
Const adInteger      = 3
Const adOpenKeyset   = 1
Const adLockReadOnly = 1

' ---------- CONFIG ----------
Dim INI_PATH
INI_PATH = "S:\CarolinaTrust\byREQUEST\Templates\GL_YTD.ini"

' ---------- GLOBALS ----------
Dim gCfg, gCn
Dim gConnStr, gFullTableName
Dim gTable2Part, gObjIdName
Dim gHasProcessedHeader
Dim gLineCounter

Dim gSourceFile
Dim gAbortRun : gAbortRun = False
Dim gRowsInserted : gRowsInserted = 0

Dim gReportType   ' GLBudgetReport_MTD / GLBudgetReport_YTD / GLIncomeStatement / GLBalanceSheet
Dim gHasHeader    ' default True (INI override)
Dim gDelimiter    ' default comma (INI override)

' ============================================================
' StartDoc
' ============================================================
Sub StartDoc
    On Error Resume Next

    gHasProcessedHeader = False
    gRowsInserted = 0
    gAbortRun = False
	gLineCounter = 0
	
    gSourceFile = ""
    On Error Resume Next
    gSourceFile = Spoolfile.Name
    On Error GoTo 0

    LogLine "GL Staging Import Started: " & Now()
    LogLine "INI_PATH = " & INI_PATH
    LogLine "Spoolfile.Name (SourceFile) = " & gSourceFile

    ' --- Load INI ---
    Set gCfg = ReadIniSafe(INI_PATH)
    If gCfg Is Nothing Then
        LogLine "ERROR: INI not found/unreadable: " & INI_PATH
        gAbortRun = True
        EndRun
        Exit Sub
    End If

    gConnStr = GetRequired(gCfg, "Database.ConnStr")
    gFullTableName = GetRequired(gCfg, "Database.TableName")


    gHasHeader = True
    If Not gCfg Is Nothing Then
        If gCfg.Exists("GL.HasHeader") Then
            gHasHeader = (LCase(Trim(gCfg("GL.HasHeader"))) <> "false")
        End If
    End If

    gDelimiter = ","
    If Not gCfg Is Nothing Then
        If gCfg.Exists("GL.Delimiter") Then
            If Len(Trim(gCfg("GL.Delimiter"))) > 0 Then gDelimiter = Trim(gCfg("GL.Delimiter"))
        End If
    End If
	
	Dim expectedToken
	expectedToken = ""
	
	If gCfg.Exists("GL.ExpectedToken") Then
		expectedToken = UCase(Trim(gCfg("GL.ExpectedToken")))
	End If

	If expectedToken <> "" Then
		If InStr(1, UCase(gSourceFile), expectedToken) = 0 Then
			LogLine "ABORTING - File name does not match expected token '" & expectedToken & "'"
			LogLine "File: " & gSourceFile
			gAbortRun = True
			EndRun
			Exit Sub
		End If
	End If
	
	

	If Len(gConnStr) = 0 Or Len(gFullTableName) = 0 Then
		LogLine "ERROR: Missing required INI keys. Need Database.ConnStr and Database.TableName"
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

    ' --- DB connect ---
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
    LogLine "ReportType = " & gReportType
    LogLine "Target table = " & gTable2Part
    LogLine "HasHeader = " & CStr(gHasHeader)
    LogLine "Delimiter = [" & gDelimiter & "]"

    ' --- Ensure table exists (based on report type) ---
    If Not EnsureStagingTable_GL() Then
        LogLine "ERROR: Could not ensure staging table."
        gAbortRun = True
        EndRun
        Exit Sub
    End If

    ' --- Duplicate protection by SourceFile ---
    If SourceFileAlreadyImported(gSourceFile) Then
        LogLine "ABORTING - This SourceFile was already imported: " & gSourceFile
        LogLine "DELETE FROM " & gTable2Part & " WHERE SourceFile = '" & SqlQuote(gSourceFile) & "';"
        gAbortRun = True
        EndRun
        Exit Sub
    End If
End Sub

' ============================================================
' ProcessLine
' ============================================================
Sub ProcessLine
    On Error Resume Next
    If gAbortRun Then Exit Sub

    Dim line
    line = srec
    If Trim(line) = "" Then Exit Sub

    If gHasHeader And (Not gHasProcessedHeader) Then
        gHasProcessedHeader = True
        LogLine "CSV header detected and skipped"
        Exit Sub
    End If

    gLineCounter = gLineCounter + 1
    InsertMappedRow_GL line, gLineCounter

    On Error GoTo 0
End Sub
' ============================================================
' Insert mapped row - GL staging (parameterized)
' ============================================================
Sub InsertMappedRow_GL(ByVal csvLine, ByVal lineOrder)
    On Error Resume Next

    Dim f, colCount, i
    f = SplitCSV_Delimited(csvLine, gDelimiter)
    colCount = UBound(f) + 1

    If colCount < 4 Then
        LogLine "SKIP: Not enough CSV fields. Got " & colCount
        Exit Sub
    End If

    For i = 0 To UBound(f)
        f(i) = IntelligentTrim(f(i))
    Next

    Dim effDateVal
    effDateVal = ParseDateFlexibleToVariantDate(f(0))
    If IsNull(effDateVal) Then
        LogLine "SKIP: Invalid EffectiveDate: " & f(0)
        Exit Sub
    End If

    Dim sql, cmd, rowsAff

    sql = "INSERT INTO " & gTable2Part & " (" & _
          "EffectiveDate, GLAccount, GLAccountName, Amount, Actual, Budget, Variance, VarPct, LineOrder, SourceFile" & _
          ") VALUES (?,?,?,?,?,?,?,?,?,?)"

    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = gCn
        .CommandText = sql
        .CommandType = adCmdText

		AddParameterDate   cmd, effDateVal
		AddParameterString cmd, 20,  f(1)  
		AddParameterString cmd, 120, f(2)  
		AddParameterString cmd, 30,  f(3) 
		AddParameterString cmd, 30,  SafeField(f, colCount, 4)  
		AddParameterString cmd, 30,  SafeField(f, colCount, 5)  
		AddParameterString cmd, 30,  SafeField(f, colCount, 6)  
		AddParameterString cmd, 30,  SafeField(f, colCount, 7)  
		AddParameterInt    cmd,      lineOrder
		AddParameterString cmd, 255, gSourceFile
    End With

    rowsAff = 0
    Err.Clear
    cmd.Execute rowsAff

    If Err.Number <> 0 Then
        LogLine "ERROR: Insert failed (" & Err.Number & ") " & Err.Description
        Err.Clear
    Else
        gRowsInserted = gRowsInserted + 1
    End If

    Set cmd = Nothing
    On Error GoTo 0
End Sub


Function SafeField(ByRef arr, ByVal colCount, ByVal idx)
    SafeField = ""
    If idx <= colCount - 1 Then SafeField = arr(idx)
End Function


' ============================================================
' EnsureStagingTable - GL (creates the right schema for report type)
' ============================================================
Function EnsureStagingTable_GL()
    On Error Resume Next
    EnsureStagingTable_GL = False

    Dim sql

    ' One universal staging shape (Amount is generic)
    ' Uses gTable2Part + gObjIdName from BuildTableNames(Database.TableName)

    sql = "IF OBJECT_ID(N'" & SqlQuote(gObjIdName) & "', N'U') IS NULL " & _
          "BEGIN " & _
          " CREATE TABLE " & gTable2Part & " (" & _
          "  Id INT IDENTITY(1,1) NOT NULL PRIMARY KEY," & _
          "  EffectiveDate DATE NOT NULL," & _
          "  GLAccount NVARCHAR(20) NULL," & _
          "  GLAccountName NVARCHAR(120) NULL," & _
          "  Amount NVARCHAR(30) NULL," & _
          "  Actual NVARCHAR(30) NULL," & _
          "  Budget NVARCHAR(30) NULL," & _
          "  Variance NVARCHAR(30) NULL," & _
          "  VarPct NVARCHAR(30) NULL," & _
          "  LineOrder INT NULL," & _
          "  SourceFile NVARCHAR(255) NULL," & _
          "  ImportTimestamp DATETIME NOT NULL DEFAULT(GETDATE())" & _
          " ) " & _
          "END"

    If Not ExecNonQuery(sql) Then Exit Function

    ' --- Indexes: EffectiveDate, GLAccount, SourceFile (dup protection) ---
    If Not EnsureIndex("IX_GL_EffectiveDate", "EffectiveDate") Then Exit Function
    If Not EnsureIndex("IX_GL_GLAccount", "GLAccount") Then Exit Function
    If Not EnsureIndex("IX_GL_SourceFile", "SourceFile") Then Exit Function

    EnsureStagingTable_GL = True
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

Function SourceFileAlreadyImported(ByVal fileName)
    On Error Resume Next
    SourceFileAlreadyImported = False

    Dim sql, cmd, rs
    sql = "SELECT COUNT(*) AS Cnt FROM " & gTable2Part & " WHERE SourceFile = ?"

    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = gCn
    cmd.CommandText = sql
    cmd.CommandType = adCmdText
    cmd.Parameters.Append cmd.CreateParameter("", adVarWChar, adParamInput, 255, fileName)

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open cmd, , adOpenKeyset, adLockReadOnly
    If Not rs.EOF Then
        SourceFileAlreadyImported = (CLng(rs("Cnt")) > 0)
    End If
    rs.Close

    Set rs = Nothing
    Set cmd = Nothing
    On Error GoTo 0
End Function

' ============================================================
' Helpers (same style as your base script)
' ============================================================
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
            ' skip
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
    ' Same as your SplitCSV, but supports delimiter override
    On Error Resume Next

    Dim fields(), fieldCount, i, inQuotes, fieldValue, ch
    ReDim fields(200)

    fieldCount = 0
    inQuotes = False
    fieldValue = ""

    If Len(delim) = 0 Then delim = ","
    Dim d : d = delim

    For i = 1 To Len(line)
        ch = Mid(line, i, 1)
        If ch = """" Then
            inQuotes = Not inQuotes
        ElseIf ch = d And Not inQuotes Then
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

Function ParseDateFlexibleToVariantDate(ByVal s)
    ' Accepts:
    '   MM/DD/YY, MM/DD/YYYY
    '   YYYY-MM-DD
    '   MM-DD-YYYY
    ' Returns Date variant, or Null
    On Error Resume Next

    ParseDateFlexibleToVariantDate = Null
    s = Trim(CStr(s))
    If Len(s) = 0 Then Exit Function

    ' Try ISO first
    If InStr(s, "-") > 0 And Len(s) >= 10 Then
        Dim a : a = Split(s, "-")
        If UBound(a) = 2 Then
            If Len(a(0)) = 4 And IsNumeric(a(0)) And IsNumeric(a(1)) And IsNumeric(a(2)) Then
                ParseDateFlexibleToVariantDate = DateSerial(CLng(a(0)), CLng(a(1)), CLng(a(2)))
                If Err.Number <> 0 Then
                    Err.Clear
                    ParseDateFlexibleToVariantDate = Null
                End If
                Exit Function
            End If
        End If
    End If

    ' Normalize - to /
    s = Replace(s, "-", "/")
    Dim p : p = Split(s, "/")
    If UBound(p) <> 2 Then Exit Function

    Dim mm, dd, yy
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

    LogLine "TOTAL ROWS INSERTED: " & gRowsInserted
    LogLine "GL Staging Import Completed: " & Now()
    LogLine "========================================="
    EndRun
    On Error GoTo 0
End Sub