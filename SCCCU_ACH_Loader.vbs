
' ============================================================
' byREQUEST ACH EPSA CSV -> SQL Staging Import (INI-driven)
' - Creates staging table with all required columns
' - CSV header is always first row
' - CSV has 31 columns with input format headers
' ============================================================
Option Explicit

' ---------- ADO CONSTANTS (numeric; VBScript doesn't have enums) ----------
Const adCmdText     = 1
Const adParamInput  = 1
Const adVarWChar    = 202
Const adDBDate      = 133
Const adOpenKeyset  = 1
Const adLockReadOnly= 1

' ---------- CONFIG ----------

	Dim INI_PATH
	INI_PATH = "C:\byREQUEST\Templates\EPSA_ACH.ini"
    INI_PATH = "S:\CarolinaTrust\byREQUEST\Templates\ACH_Trans.ini"

' ---------- GLOBALS ----------
Dim gCfg, gCn
Dim gConnStr, gFullTableName
Dim gTable2Part, gObjIdName
Dim gDateColFromSpool
Dim gHasProcessedHeader
Dim srec
Dim gEPSA_File, gFileDateMMDDYY, gFileSequence
Dim gAbortRun : gAbortRun = False
Dim gRowsInserted
gRowsInserted = 0



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

    ' --- Parse filename EPSA_mmddyy_sequence ---
    Dim fileBase, parts
    fileBase = Spoolfile.Name
    If InStrRev(fileBase, ".") > 0 Then
        fileBase = Left(fileBase, InStrRev(fileBase, ".") - 1)
    End If

    ' Expect: EPSA_011026_89779  (or similar)
    parts = Split(fileBase, "_")
	If UBound(parts) >= 2 Then
		gFileDateMMDDYY = Trim(parts(1))   ' e.g. "011026"
		
		Dim seqRaw 
		seqRaw = Trim(parts(2))            ' full part after second _
		
		' Option A: Strict left-6 (your current approach - simple & fast)
		gFileSequence = Left(seqRaw, 6)
		gEPSA_File = gFileDateMMDDYY & "_" & gFileSequence
		
		' Optional: Validate lengths (helps catch bad parses)
		If Len(gFileDateMMDDYY) <> 6 Or Not IsNumeric(gFileDateMMDDYY) _
		   Or Len(gFileSequence) <> 6 Or Not IsNumeric(gFileSequence) Then
			LogLine "WARNING: Parsed values look invalid - Date:" & gFileDateMMDDYY & " Seq:" & gFileSequence
			' Decide if you want to abort or continue
		End If
		
	Else
		gEPSA_File = ""
		LogLine "ERROR: Cannot parse mmddyy_sequence from filename: " & Spoolfile.Name
		LogLine "Processing aborted."
		gAbortRun = True
		EndRun
		Exit Sub
	End If

    LogLine "EPSA File identifier: " & gEPSA_File
    LogLine "  -> Date part: " & gFileDateMMDDYY
    LogLine "  -> Sequence:  " & gFileSequence

    ' --- Load INI ---
    Set gCfg = ReadIniSafe(INI_PATH)
    If gCfg Is Nothing Then
        LogLine "ERROR: INI not found/unreadable: " & INI_PATH
        EndRun
		gAbortRun = True

        Exit Sub
    End If

    gConnStr = GetRequired(gCfg, "Database.ConnStr")
    gFullTableName = GetRequired(gCfg, "Database.TableName")
    If Len(gConnStr) = 0 Or Len(gFullTableName) = 0 Then
        EndRun
		gAbortRun = True

        Exit Sub
    End If

    BuildTableNames gFullTableName, gTable2Part, gObjIdName
    If Len(gTable2Part) = 0 Or Len(gObjIdName) = 0 Then
        LogLine "ERROR: Could not parse TableName: " & gFullTableName
        EndRun
		gAbortRun = True

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
		gAbortRun = True
        Exit Sub
    End If

    ' --- Early check: has this file already been processed? ---
    If EPSAFileAlreadyImported(gEPSA_File) Then
		LogLine "ABORTING - This EPSA file was already imported: " & gEPSA_File
		LogLine "Reason: Duplicate protection is enabled to prevent re-importing the same EPSA file."
		LogLine "Action required to re-import:"
		LogLine "  1) Manually delete ALL rows for this EPSA file from the staging table."
		LogLine "  2) Re-run the job after deletion is complete."
		LogLine " "
		LogLine "Use the EXACT SQL below in SSMS / SQL command:"
		LogLine "------------------------------------------------------------"
		LogLine "DELETE FROM " & gTable2Part
		LogLine "WHERE EPSA_File = '" & gEPSA_File & "';"
		LogLine "------------------------------------------------------------"
		LogLine "IMPORTANT:"
		LogLine " - Partial deletes will NOT allow the job to rerun."
		LogLine " - If ANY row remains for this EPSA_File, the job will abort again."
		LogLine " - ALL rows for this EPSA_File must be deleted before re-importing."
		LogLine "------------------------------------------------------------"
		LogLine " "
        EndRun
		gAbortRun = True
        Exit Sub
    End If
End Sub

' ============================================================
' ProcessLine - Process each line
' ============================================================
Sub ProcessLine
    On Error Resume Next
	If gAbortRun Then Exit Sub
	
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
' ============================================================
Sub InsertMappedRow(ByVal csvLine)
    On Error Resume Next

    Dim f, i
    f = SplitCSV(csvLine)

    Dim colCount
    colCount = UBound(f) + 1
    LogLine "DEBUG: CSV columns found: " & colCount

    If colCount < 31 Then
        LogLine "SKIP: Not enough CSV fields (need 31). Got " & colCount
        If colCount > 0 Then
            LogLine "DEBUG First few columns:"
            For i = 0 To UBound(f)
                If i <= 5 Then LogLine " [" & i & "]: " & f(i)
            Next
        End If
        Exit Sub
    End If

    For i = 0 To 30
        f(i) = IntelligentTrim(f(i))
    Next

    If f(1) = "" Then
        LogLine "SKIP: Post_Date is empty"
        Exit Sub
    End If

    ' Validate Post_Date roughly matches filename date
    Dim rowDateClean
    rowDateClean = Replace(Replace(Replace(f(1), "/", ""), "-", ""), " ", "")
    If Len(rowDateClean) >= 6 Then
        rowDateClean = Left(rowDateClean, 6)
    Else
        rowDateClean = ""
    End If

    ' If rowDateClean <> gFileDateMMDDYY Then
        ' LogLine "SKIP row - Post_Date mismatch (file:" & gFileDateMMDDYY & " <> row:" & rowDateClean & ")"
        ' Exit Sub
    ' End If

    ' Skip entire row if critical fields are blank
    If f(1) = "" Or f(2) = "" Or f(14) = "" Then
        LogLine "SKIP: Missing required fields (Post_Date, Post_Time, or EFF DATE)"
        Exit Sub
    End If

    Dim postDateVal
    postDateVal = ParseUSDateToVariantDate(f(1))
    If IsNull(postDateVal) Then
        LogLine "SKIP: Post_Date not parseable as US date: " & f(1)
        Exit Sub
    End If

    Dim sql, cmd
	sql = "INSERT INTO " & gTable2Part & " (" & _
		  "[Report],[Post_Date],[Post_Time],[Settlement_Date],[Company_Name],[ACH_Batch_ID],[PCTL],[LCTL]," & _
		  "[Status],[Tran_Type],[Tran_Code],[Effective_Date],[Amount],[Member_Nbr],[Acct],[Mbr_Name],[Fee]," & _
		  "[ACH_Ind_Name],[DD],[Trace_Nbr],[Description],[Type],[Post_Seq],[File_Post_Seq]," & _
		  "[Source_Name],[Batch],[Check_Number],[SDC_Number],[S_L],[RR],[RDFI_ID]," & _
		  "[DFI_Account_Number]," & _
		  "[Individual_ID],[MESSAGE]," & _
		  "[EPSA_File]" & _
		  ") VALUES (" & _
		  "?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?" & _
		  ")"


    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = gCn
        .CommandText = sql
        .CommandType = adCmdText

' 36 parameters in the exact same order as the INSERT list above:

		AddParameterString cmd, 10,  f(0)        ' Report
		AddParameterDate   cmd,      postDateVal ' Post_Date (DATE)
		AddParameterString cmd, 20,  f(2)        ' Post_Time
		AddParameterString cmd, 20,  f(3)        ' Settlement_Date
		AddParameterString cmd, 100, f(4)        ' Company_Name
		AddParameterString cmd, 20,  f(5)        ' ACH_Batch_ID
		AddParameterString cmd, 20,  f(6)        ' PCTL
		AddParameterString cmd, 20,  f(7)        ' LCTL
		AddParameterString cmd, 5,   f(12)       ' Status (ST)
		AddParameterString cmd, 30,  f(13)       ' Tran_Type (TYPE)
		AddParameterString cmd, 10,  f(10)       ' Tran_Code
		AddParameterString cmd, 20,  f(14)       ' Effective_Date
		AddParameterString cmd, 30,  f(15)       ' Amount
		AddParameterString cmd, 30,  f(18)       ' Member_Nbr
		AddParameterString cmd, 10,  f(19)       ' Acct
		AddParameterString cmd, 100, f(20)       ' Mbr_Name
		AddParameterString cmd, 30,  f(21)       ' Fee
		AddParameterString cmd, 100, f(27)       ' ACH_Ind_Name
		AddParameterString cmd, 5,   f(28)       ' DD
		AddParameterString cmd, 40,  f(29)       ' Trace_Nbr
		AddParameterString cmd, 255, f(11)       ' Description
		AddParameterString cmd, 5,   f(23)       ' Type
		AddParameterString cmd, 5,   f(8)        ' Post_Seq
		AddParameterString cmd, 5,   f(9)        ' File_Post_Seq

		AddParameterString cmd, 100, f(4)        ' Source_Name
		AddParameterString cmd, 20,  f(5)        ' Batch
		AddParameterString cmd, 20,  f(16)       ' Check_Number
		AddParameterString cmd, 20,  f(17)       ' SDC_Number
		AddParameterString cmd, 10,  f(19)       ' S_L
		AddParameterString cmd, 10,  f(22)       ' RR
		AddParameterString cmd, 20,  f(24)       ' RDFI_ID
		AddParameterString cmd, 50,  f(25)       ' DFI_Account_Number
		AddParameterString cmd, 30,  f(26)       ' Individual_ID
		AddParameterString cmd, 255, f(30)       ' MESSAGE

		AddParameterString cmd, 100, gEPSA_File  ' EPSA_File

    End With

	Dim rowsAff
	rowsAff = 0

	Err.Clear
	cmd.Execute rowsAff

	If Err.Number <> 0 Then
		LogLine "ERROR: Insert failed (" & Err.Number & ") " & Err.Description
		LogLine "Line: " & Left(csvLine, 150)
		LogLine "EPSA_File: " & gEPSA_File
		Err.Clear
	Else
		' For a normal INSERT, rowsAff = 1
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
' EnsureStagingTable - Create table with all columns
' ============================================================
Function EnsureStagingTable()
    On Error Resume Next
    EnsureStagingTable = False

 Dim sql

	sql = "IF OBJECT_ID(N'" & SqlQuote(gObjIdName) & "', N'U') IS NULL " & _
		  "BEGIN " & _
		  " CREATE TABLE " & gTable2Part & " (" & _
		  " [id] INT IDENTITY(1,1) PRIMARY KEY," & _
		  " [Report] NVARCHAR(10) NULL," & _
		  " [Post_Date] DATE NULL," & _
		  " [Post_Time] NVARCHAR(20) NULL," & _
		  " [Settlement_Date] NVARCHAR(20) NULL," & _
		  " [Company_Name] NVARCHAR(100) NULL," & _
		  " [ACH_Batch_ID] NVARCHAR(20) NULL," & _
		  " [PCTL] NVARCHAR(20) NULL," & _
		  " [LCTL] NVARCHAR(20) NULL," & _
		  " [Status] NVARCHAR(5) NULL," & _
		  " [Tran_Type] NVARCHAR(30) NULL," & _
		  " [Tran_Code] NVARCHAR(10) NULL," & _
		  " [Effective_Date] NVARCHAR(20) NULL," & _
		  " [Amount] NVARCHAR(30) NULL," & _
		  " [Member_Nbr] NVARCHAR(30) NULL," & _
		  " [Acct] NVARCHAR(10) NULL," & _
		  " [Mbr_Name] NVARCHAR(100) NULL," & _
		  " [Fee] NVARCHAR(30) NULL," & _
		  " [ACH_Ind_Name] NVARCHAR(100) NULL," & _
		  " [DD] NVARCHAR(5) NULL," & _
		  " [Trace_Nbr] NVARCHAR(40) NULL," & _
		  " [Description] NVARCHAR(255) NULL," & _
		  " [Type] NVARCHAR(5) NULL," & _
		  " [Post_Seq] NVARCHAR(5) NULL," & _
		  " [File_Post_Seq] NVARCHAR(5) NULL," & _
		  " [Source_Name] NVARCHAR(100) NULL," & _
		  " [Batch] NVARCHAR(20) NULL," & _
		  " [Check_Number] NVARCHAR(20) NULL," & _
		  " [SDC_Number] NVARCHAR(20) NULL," & _
		  " [S_L] NVARCHAR(10) NULL," & _
		  " [RR] NVARCHAR(10) NULL," & _
		  " [RDFI_ID] NVARCHAR(20) NULL," & _
		  " [DFI_Account_Number] NVARCHAR(50) NULL," & _   
		  " [Individual_ID] NVARCHAR(30) NULL," & _
		  " [MESSAGE] NVARCHAR(255) NULL," & _
		  " [EPSA_File] NVARCHAR(100) NULL," & _
		  " [Import_Timestamp] DATETIME DEFAULT GETDATE()" & _
		  " ) " & _
		  "END"


    If Not ExecNonQuery(sql) Then Exit Function

    sql = "IF NOT EXISTS (" & _
          "SELECT 1 FROM sys.indexes " & _
          "WHERE name = 'IX_ACH_EPSA_Staging_EPSA_File' " & _
          "AND object_id = OBJECT_ID(N'" & SqlQuote(gObjIdName) & "')" & _
          ") BEGIN " & _
          "CREATE INDEX IX_ACH_EPSA_Staging_EPSA_File ON " & gTable2Part & "([EPSA_File]) " & _
          "END"

    If Not ExecNonQuery(sql) Then Exit Function

    EnsureStagingTable = True
    On Error GoTo 0
End Function

' ============================================================
' Helper Functions
' ============================================================
Function EPSAFileAlreadyImported(ByVal fileId)
    On Error Resume Next
    EPSAFileAlreadyImported = False

    Dim sql, cmd, rs
    sql = "SELECT COUNT(*) AS Cnt FROM " & gTable2Part & " WHERE [EPSA_File] = ?"

    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = gCn
    cmd.CommandText = sql
    cmd.CommandType = adCmdText
    cmd.Parameters.Append cmd.CreateParameter("", adVarWChar, adParamInput, 100, fileId)

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open cmd, , adOpenKeyset, adLockReadOnly

    If Not rs.EOF Then
        EPSAFileAlreadyImported = (CLng(rs("Cnt")) > 0)
    End If

    rs.Close
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

    ' dateVal should be a VBScript Date variant
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

    Dim p
    p = Split(s, "/")
    If UBound(p) <> 2 Then Exit Function

    Dim mm, dd, yy
    mm = CLng(p(0))
    dd = CLng(p(1))
    yy = CLng(p(2))

    If yy < 100 Then
        ' Adjust as needed; common ACH files are 20xx
        yy = 2000 + yy
    End If

    ParseUSDateToVariantDate = DateSerial(yy, mm, dd)

    If Err.Number <> 0 Then
        Err.Clear
        ParseUSDateToVariantDate = Null
    End If

    On Error GoTo 0
End Function

Function ExtractDateColFromSpoolName(fname)
    Dim pos, tmp, dotPos
    ExtractDateColFromSpoolName = ""

    If IsNull(fname) Or fname = "" Then Exit Function

    pos = InStrRev(fname, "_")
    If pos > 0 Then
        tmp = Mid(fname, pos + 1)
    Else
        tmp = fname
    End If

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
	If gAbortRun Then Exit Sub
	
	LogLine "TOTAL ROWS INSERTED: " & gRowsInserted

    LogLine "ACH EPSA Import Completed: " & Now()
    LogLine "========================================="
    EndRun
    On Error GoTo 0
End Sub
