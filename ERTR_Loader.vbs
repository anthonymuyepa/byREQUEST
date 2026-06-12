Option Explicit

' ============================================================
' byREQUEST ERTR Flattened CSV -> SQL Staging Import (INI-driven)
'
' Input CSV columns (21):
'  0  Post_Date        (MM/DD/YYYY)  -> stored as DATE
'  1  TC
'  2  RDFI R&T#
'  3  DFI ACCOUNT
'  4  RTN AMOUNT
'  5  INDIVIDUAL ID
'  6  INDIVIDUAL NAME
'  7  DD
'  8  A
'  9  TRACE NUMBER
' 10  RRC
' 11  ORC
' 12  DRC
' 13  SRC BH#
' 14  RTN TRACE#
' 15  RTN RDFI#
' 16  DEATHDATE
' 17  INFO
' 18  COMPANY NAME
' 19  COMPANY ID
' 20  ERTR_File
'
' Notes:
' - CSV header is always the first row (skipped)
' - Duplicate protection: blocks if ERTR_File already exists in table
' - DB ConnStr and TableName come from INI:
'     [Database]
'     ConnStr=...
'     TableName=dbo.ERTR_RTN_Staging
' ============================================================

' ---------- ADO CONSTANTS ----------
Const adCmdText      = 1
Const adParamInput   = 1
Const adVarWChar     = 202
Const adDBDate       = 133
Const adOpenKeyset   = 1
Const adLockReadOnly = 1

' ---------- CONFIG ----------
Dim INI_PATH
INI_PATH = "C:\byREQUEST\Templates\ERTR.ini"   ' reuse your existing INI
INI_PATH = "S:\CarolinaTrust\byREQUEST\Templates\ERTR.ini"

' ---------- GLOBALS ----------
Dim gCfg, gCn
Dim gConnStr, gFullTableName
Dim gTable2Part, gObjIdName
Dim gHasProcessedHeader
Dim srec

Dim gERTR_File
Dim gAbortRun : gAbortRun = False
Dim gRowsInserted : gRowsInserted = 0

' ============================================================
' StartDoc
' ============================================================
Sub StartDoc
    On Error Resume Next

    gHasProcessedHeader = False
    gERTR_File = ""
    On Error Resume Next
    gERTR_File = Spoolfile.Name
    On Error GoTo 0

    LogLine "ERTR Flattened CSV Import Started: " & Now()
    LogLine "INI_PATH = " & INI_PATH
    LogLine "Spoolfile.Name (ERTR_File) = " & gERTR_File

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
    If Len(gConnStr) = 0 Or Len(gFullTableName) = 0 Then
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
    LogLine "Target table = " & gTable2Part

    ' --- Ensure table exists ---
    If Not EnsureStagingTable_ERTR() Then
        LogLine "ERROR: Could not ensure staging table."
        gAbortRun = True
        EndRun
        Exit Sub
    End If

    ' --- Duplicate protection by ERTR_File ---
    If ERTRFileAlreadyImported(gERTR_File) Then
        LogLine "ABORTING - This ERTR file was already imported: " & gERTR_File
        LogLine "DELETE FROM " & gTable2Part & " WHERE ERTR_File = '" & SqlQuote(gERTR_File) & "';"
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

    ' First row is header - skip
    If Not gHasProcessedHeader Then
        gHasProcessedHeader = True
        LogLine "CSV header detected and skipped"
        Exit Sub
    End If

    InsertMappedRow_ERTR line

    On Error GoTo 0
End Sub

' ============================================================
' Insert mapped row - 21 columns
' ============================================================
Sub InsertMappedRow_ERTR(ByVal csvLine)
    On Error Resume Next

    Dim f, i, colCount
    f = SplitCSV(csvLine)
    colCount = UBound(f) + 1

    If colCount < 21 Then
        LogLine "SKIP: Not enough CSV fields (need 21). Got " & colCount
        LogLine "Line: " & Left(csvLine, 200)
        Exit Sub
    End If

    For i = 0 To 20
        f(i) = IntelligentTrim(f(i))
    Next

    ' Parse Post_Date (MM/DD/YYYY) -> DATE
    Dim postDateVal
    postDateVal = ParseUSDateToVariantDate(f(0))
    If IsNull(postDateVal) Then
        LogLine "SKIP: Post_Date not parseable: " & f(0)
        Exit Sub
    End If

    ' Critical field
    If f(1) = "" Then
        LogLine "SKIP: TC empty"
        Exit Sub
    End If

    Dim sql, cmd, rowsAff
	sql = "INSERT INTO " & gTable2Part & " (" & _
		  "[Post_Date],[TC],[RDFI_RT#],[DFI_Account],[RTN_Amount],[Individual_ID],[Individual_Name],[DD],[A_Flag]," & _
		  "[Trace_Number],[RRC],[ORC],[DRC],[SRC_BH#],[RTN_Trace#],[RTN_RDFI#],[DeathDate],[Info]," & _
		  "[Company_Name],[Company_ID],[ERTR_File]" & _
		  ") VALUES (" & _
		  "?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?" & _
		  ")"


    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = gCn
        .CommandText = sql
        .CommandType = adCmdText

        AddParameterDate   cmd,      postDateVal  ' Post_Date (DATE)
        AddParameterString cmd, 10,  f(1)         ' TC
        AddParameterString cmd, 30,  f(2)         ' RDFI R&T#
        AddParameterString cmd, 50,  f(3)         ' DFI ACCOUNT
        AddParameterString cmd, 30,  f(4)         ' RTN AMOUNT
        AddParameterString cmd, 30,  f(5)         ' INDIVIDUAL ID
        AddParameterString cmd, 120, f(6)         ' INDIVIDUAL NAME
        AddParameterString cmd, 5,   f(7)         ' DD
        AddParameterString cmd, 5,   f(8)         ' A
        AddParameterString cmd, 60,  f(9)         ' TRACE NUMBER
        AddParameterString cmd, 10,  f(10)        ' RRC
        AddParameterString cmd, 10,  f(11)        ' ORC
        AddParameterString cmd, 10,  f(12)        ' DRC
        AddParameterString cmd, 30,  f(13)        ' SRC BH#
        AddParameterString cmd, 60,  f(14)        ' RTN TRACE#
        AddParameterString cmd, 30,  f(15)        ' RTN RDFI#
        AddParameterString cmd, 20,  f(16)        ' DEATHDATE (keep as text for now)
        AddParameterString cmd, 255, f(17)        ' INFO
        AddParameterString cmd, 150, f(18)        ' COMPANY NAME
        AddParameterString cmd, 40,  f(19)        ' COMPANY ID
        AddParameterString cmd, 255, f(20)        ' ERTR_File (from CSV)

    End With

    rowsAff = 0
    Err.Clear
    cmd.Execute rowsAff

    If Err.Number <> 0 Then
        LogLine "ERROR: Insert failed (" & Err.Number & ") " & Err.Description
        LogLine "Line: " & Left(csvLine, 220)
        Err.Clear
    Else
        If IsNumeric(rowsAff) Then
            gRowsInserted = gRowsInserted + CLng(rowsAff)
        Else
            gRowsInserted = gRowsInserted + 1
        End If
    End If

    Set cmd = Nothing
    On Error GoTo 0
End Sub

' ============================================================
' EnsureStagingTable - ERTR
' ============================================================
Function EnsureStagingTable_ERTR()
    On Error Resume Next
    EnsureStagingTable_ERTR = False

    Dim sql
    sql = "IF OBJECT_ID(N'" & SqlQuote(gObjIdName) & "', N'U') IS NULL " & _
          "BEGIN " & _
          " CREATE TABLE " & gTable2Part & " (" & _
          "  [id] INT IDENTITY(1,1) PRIMARY KEY," & _
          "  [Post_Date] DATE NULL," & _
          "  [TC] NVARCHAR(10) NULL," & _
          "  [RDFI_RT#] NVARCHAR(30) NULL," & _
          "  [DFI_Account] NVARCHAR(50) NULL," & _
          "  [RTN_Amount] NVARCHAR(30) NULL," & _
          "  [Individual_ID] NVARCHAR(30) NULL," & _
          "  [Individual_Name] NVARCHAR(120) NULL," & _
          "  [DD] NVARCHAR(5) NULL," & _
          "  [A_Flag] NVARCHAR(5) NULL," & _
          "  [Trace_Number] NVARCHAR(60) NULL," & _
          "  [RRC] NVARCHAR(10) NULL," & _
          "  [ORC] NVARCHAR(10) NULL," & _
          "  [DRC] NVARCHAR(10) NULL," & _
          "  [SRC_BH#] NVARCHAR(30) NULL," & _
          "  [RTN_Trace#] NVARCHAR(60) NULL," & _
          "  [RTN_RDFI#] NVARCHAR(30) NULL," & _
          "  [DeathDate] NVARCHAR(20) NULL," & _
          "  [Info] NVARCHAR(255) NULL," & _
          "  [Company_Name] NVARCHAR(150) NULL," & _
          "  [Company_ID] NVARCHAR(40) NULL," & _
          "  [ERTR_File] NVARCHAR(255) NULL," & _
          "  [Import_Timestamp] DATETIME DEFAULT GETDATE()" & _
          " ) " & _
          "END"

    If Not ExecNonQuery(sql) Then Exit Function

    sql = "IF NOT EXISTS (" & _
          "SELECT 1 FROM sys.indexes " & _
          "WHERE name = 'IX_ERTR_Staging_ERTR_File' " & _
          "AND object_id = OBJECT_ID(N'" & SqlQuote(gObjIdName) & "')" & _
          ") BEGIN " & _
          "CREATE INDEX IX_ERTR_Staging_ERTR_File ON " & gTable2Part & "([ERTR_File]) " & _
          "END"

    If Not ExecNonQuery(sql) Then Exit Function

    EnsureStagingTable_ERTR = True
    On Error GoTo 0
End Function

Function ERTRFileAlreadyImported(ByVal fileName)
    On Error Resume Next
    ERTRFileAlreadyImported = False

    Dim sql, cmd, rs
    sql = "SELECT COUNT(*) AS Cnt FROM " & gTable2Part & " WHERE [ERTR_File] = ?"

    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = gCn
    cmd.CommandText = sql
    cmd.CommandType = adCmdText
    cmd.Parameters.Append cmd.CreateParameter("", adVarWChar, adParamInput, 255, fileName)

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open cmd, , adOpenKeyset, adLockReadOnly
    If Not rs.EOF Then
        ERTRFileAlreadyImported = (CLng(rs("Cnt")) > 0)
    End If
    rs.Close

    Set rs = Nothing
    Set cmd = Nothing
    On Error GoTo 0
End Function

' ============================================================
' Helpers (adapted from your base script)
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

Function SplitCSV(line)
    On Error Resume Next

    Dim fields(), fieldCount, i, inQuotes, fieldValue, ch
    ReDim fields(120)

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

Function ParseUSDateToVariantDate(ByVal s)
    ' Accepts: MM/DD/YY or MM/DD/YYYY (also allows '-')
    ' Returns: Date variant, or Null if not parseable
    On Error Resume Next

    ParseUSDateToVariantDate = Null
    s = Trim(CStr(s))
    If Len(s) = 0 Then Exit Function

    s = Replace(s, "-", "/")
    Dim p : p = Split(s, "/")
    If UBound(p) <> 2 Then Exit Function

    Dim mm, dd, yy
    mm = CLng(p(0))
    dd = CLng(p(1))
    yy = CLng(p(2))
    If yy < 100 Then yy = 2000 + yy

    ParseUSDateToVariantDate = DateSerial(yy, mm, dd)

    If Err.Number <> 0 Then
        Err.Clear
        ParseUSDateToVariantDate = Null
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
    LogLine "ERTR Flattened CSV Import Completed: " & Now()
    LogLine "========================================="
    EndRun
    On Error GoTo 0
End Sub
