Option Explicit

' ============================================================
' byREQUEST - Ensure Reconciliation View + Indexes
' - Creates/updates dbo.vw_ACH_EPSA_Reconciled
' - Ensures supporting indexes exist on EPSA + ERTR
' - Safe to run repeatedly (idempotent)
' ============================================================

Const adCmdText = 1

' ---- CONFIG ----
Dim INI_PATH
INI_PATH = "S:\CarolinaTrust\byREQUEST\Templates\ERTR.ini"

Const VIEW_FULLNAME = "dbo.vw_ACH_EPSA_Reconciled"   ' change if desired

' ---- GLOBALS ----
Dim gCn, gCfg, gConnStr
Dim gAbortRun : gAbortRun = False

Sub StartDoc()
    On Error Resume Next

    LogLine "Ensure View/Indexes Started: " & Now()
    LogLine "INI_PATH = " & INI_PATH
    LogLine "View     = " & VIEW_FULLNAME

    Set gCfg = ReadIniSafe(INI_PATH)
    If gCfg Is Nothing Then
        LogLine "ERROR: INI not found/unreadable: " & INI_PATH
        gAbortRun = True
        Exit Sub
    End If

    gConnStr = GetRequired(gCfg, "Database.ConnStr")
    If Len(gConnStr) = 0 Then
        LogLine "ERROR: Missing Database.ConnStr in INI"
        gAbortRun = True
        Exit Sub
    End If

    Set gCn = CreateObject("ADODB.Connection")
    gCn.Open gConnStr
    If Err.Number <> 0 Then
        LogLine "ERROR: DB open failed (" & Err.Number & ") " & Err.Description
        Err.Clear
        gAbortRun = True
        Exit Sub
    End If

    LogLine "SUCCESS: Connected"

    EnsureViewAndIndexes

    On Error GoTo 0
End Sub

Sub ProcessLine()
    ' This is a maintenance job; ignore input lines
End Sub

Sub CloseDoc()
    On Error Resume Next
    If gAbortRun Then
        LogLine "ABORTED"
    Else
        LogLine "Ensure View/Indexes Completed: " & Now()
    End If

    If Not gCn Is Nothing Then
        gCn.Close
        Set gCn = Nothing
        LogLine "DB connection closed"
    End If
    On Error GoTo 0
End Sub

' ============================================================
' Core: create view + indexes
' ============================================================
Sub EnsureViewAndIndexes()
    On Error Resume Next

    ' 1) Create or alter view (safe even if it exists)
    Dim viewSql
    viewSql = BuildViewSql(VIEW_FULLNAME)

    If ExecNonQuery(viewSql) Then
        LogLine "OK: View ensured (CREATE OR ALTER succeeded): " & VIEW_FULLNAME
    Else
        LogLine "ERROR: Failed to create/alter view: " & VIEW_FULLNAME
        gAbortRun = True
        Exit Sub
    End If

    ' 2) Ensure indexes (create only if missing)
    EnsureIndex "dbo.ACH_EPSA", "IX_EPSA_PostDate", "Post_Date"
    EnsureIndex "dbo.ACH_EPSA", "IX_EPSA_Trace_RDFI", "Trace_Nbr, RDFI_ID"

    EnsureIndex "dbo.ACH_ERTR", "IX_ERTR_PostDate", "Post_Date"
    EnsureIndex "dbo.ACH_ERTR", "IX_ERTR_RTNTrace_RDFI", "[RTN_Trace#], [RTN_RDFI#]"

    On Error GoTo 0
End Sub

Function BuildViewSql(ByVal viewFullName)
    ' Build the CREATE OR ALTER VIEW statement as one string
    Dim sql
    sql = ""
    sql = sql & "CREATE OR ALTER VIEW " & viewFullName & vbCrLf
    sql = sql & "AS" & vbCrLf
    sql = sql & "WITH ERTR_Match AS (" & vbCrLf
    sql = sql & "    SELECT" & vbCrLf
    sql = sql & "        e.id AS ERTR_id," & vbCrLf
    sql = sql & "        LTRIM(RTRIM(e.[RTN_Trace#])) AS Orig_Trace," & vbCrLf
    sql = sql & "        LEFT(LTRIM(RTRIM(e.[RTN_RDFI#])), 8) AS Orig_RDFI8," & vbCrLf
    sql = sql & "        e.RRC       AS Return_Code," & vbCrLf
    sql = sql & "        e.Post_Date AS Return_Post_Date," & vbCrLf
    sql = sql & "        e.ERTR_File" & vbCrLf
    sql = sql & "    FROM dbo.ACH_ERTR e" & vbCrLf
    sql = sql & "    WHERE NULLIF(LTRIM(RTRIM(e.[RTN_Trace#])), '') IS NOT NULL" & vbCrLf
    sql = sql & "      AND NULLIF(LTRIM(RTRIM(e.[RTN_RDFI#])), '')  IS NOT NULL" & vbCrLf
    sql = sql & ")" & vbCrLf
    sql = sql & "SELECT" & vbCrLf
    sql = sql & "    a.*," & vbCrLf
    sql = sql & "    CASE" & vbCrLf
    sql = sql & "        WHEN a.Status = 'N' THEN 'N'" & vbCrLf
    sql = sql & "        WHEN a.Status = 'NP' AND m.ERTR_id IS NOT NULL THEN 'RT'" & vbCrLf
    sql = sql & "        ELSE NULL" & vbCrLf
    sql = sql & "    END AS Final_Status," & vbCrLf
    sql = sql & "    m.ERTR_id," & vbCrLf
    sql = sql & "    m.Return_Code," & vbCrLf
    sql = sql & "    m.Return_Post_Date," & vbCrLf
    sql = sql & "    m.ERTR_File AS Matched_ERTR_File" & vbCrLf
    sql = sql & "FROM dbo.ACH_EPSA a" & vbCrLf
    sql = sql & "LEFT JOIN ERTR_Match m" & vbCrLf
    sql = sql & "  ON LTRIM(RTRIM(a.Trace_Nbr)) = m.Orig_Trace" & vbCrLf
    sql = sql & " AND LEFT(LTRIM(RTRIM(a.RDFI_ID)), 8) = m.Orig_RDFI8" & vbCrLf
    sql = sql & "WHERE" & vbCrLf
    sql = sql & "    a.Status = 'N'" & vbCrLf
    sql = sql & " OR (a.Status = 'NP' AND m.ERTR_id IS NOT NULL);" & vbCrLf

    BuildViewSql = sql
End Function

Sub EnsureIndex(ByVal table2Part, ByVal idxName, ByVal colsList)
    On Error Resume Next

    If IndexExists(table2Part, idxName) Then
        LogLine "OK: Index exists: " & idxName & " on " & table2Part
        Exit Sub
    End If

    Dim sql
    sql = "CREATE INDEX " & idxName & " ON " & table2Part & " (" & colsList & ");"

    If ExecNonQuery(sql) Then
        LogLine "OK: Index created: " & idxName & " on " & table2Part
    Else
        LogLine "ERROR: Failed to create index: " & idxName & " on " & table2Part
    End If

    On Error GoTo 0
End Sub

Function IndexExists(ByVal table2Part, ByVal idxName)
    On Error Resume Next
    IndexExists = False

    Dim schemaName, tableName, dotPos
    dotPos = InStr(table2Part, ".")
    If dotPos > 0 Then
        schemaName = Replace(Left(table2Part, dotPos - 1), "[", "")
        schemaName = Replace(schemaName, "]", "")
        tableName = Replace(Mid(table2Part, dotPos + 1), "[", "")
        tableName = Replace(tableName, "]", "")
    Else
        schemaName = "dbo"
        tableName = table2Part
    End If

    Dim sql, rs
    sql = "SELECT 1 " & _
          "FROM sys.indexes i " & _
          "JOIN sys.objects o ON i.object_id = o.object_id " & _
          "JOIN sys.schemas s ON o.schema_id = s.schema_id " & _
          "WHERE i.name = '" & SqlQuote(idxName) & "' " & _
          "AND o.name = '" & SqlQuote(tableName) & "' " & _
          "AND s.name = '" & SqlQuote(schemaName) & "';"

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, gCn, 1, 1
    If Not rs.EOF Then IndexExists = True
    rs.Close
    Set rs = Nothing

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
        LogLine "SQL: " & Left(sql, 300)
        Err.Clear
        ExecNonQuery = False
    Else
        ExecNonQuery = True
    End If

    Set cmd = Nothing
    On Error GoTo 0
End Function

' ============================================================
' INI + small helpers
' ============================================================
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

Function SqlQuote(ByVal s)
    SqlQuote = Replace(s, "'", "''")
End Function

Sub LogLine(msg)
    fout.WriteLine msg
End Sub
