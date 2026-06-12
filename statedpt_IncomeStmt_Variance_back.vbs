Option Explicit

' =========================================================
' GL Variance Report
' Uses these three staging tables:
'   dbo.GLBalanceSheet
'   dbo.GLBudgetReport
'   dbo.GLIncomeStatement
'
' Notes:
' - Dates are real SQL date columns already.
' - Other fields remain raw text in staging.
' - Numeric conversion happens only inside SQL.
' - Handles raw numeric text formats such as:
'       <838.32>
'       915.8
'       0
'       <38,125,681.10>
'       454488.95
'       <90.70>
'       <195,431.62>
' =========================================================

' ---------------- GLOBALS ----------------
Dim gMode, gRelOffset, gYear, gMonth, gQuarter
Dim gStartDateStr, gEndDateStr
Dim gConnStr, gFullTableName, gTable2Part, gObjIdName

Dim conn, rs
Dim INI_PATH : INI_PATH = "S:\StateDepartment\GL_IncomeStmt_Variance.ini"

' ---------------------------------------------------------
' Simple log helper
' ---------------------------------------------------------
Sub LogLine(msg)
    On Error Resume Next
    fout.WriteLine msg
    On Error GoTo 0
End Sub

' ---------------------------------------------------------
' StartDoc
' ---------------------------------------------------------
Sub StartDoc
    On Error Resume Next

    Dim startKey, endKey
    Dim sqlIS, sqlBS, cmdIS, cmdBS, rsIS, rsBS
    Dim rowCount, bsRowCount
    Dim periodDisplay, startDateObj, endDateObj, priorDateObj

    If Not FileExists(INI_PATH) Then
        LogLine "ERROR: INI file not found: " & INI_PATH
        Exit Sub
    End If

    LoadParameters INI_PATH
	
	If Trim(gMode) = "" Then
		LogLine "ERROR: Aborting due to invalid INI parameters."
		Exit Sub
	End If

    If Trim(gConnStr) = "" Then
        LogLine "ERROR: ConnStr not defined in [Database] section."
        Exit Sub
    End If

	ComputeDateRange startKey, endKey
	If startKey = "" Or endKey = "" Then
		LogLine "ERROR: Could not compute date range from parameters."
		Exit Sub
	End If

	startDateObj = FormatKeyToDate(startKey)
	endDateObj   = FormatKeyToDate(endKey)

	' Open DB early because MONTHLY may need actual loaded date from SQL
	Set conn = CreateObject("ADODB.Connection")
	Err.Clear
	conn.Open gConnStr
	If Err.Number <> 0 Then
		LogLine "ERROR opening DB connection: " & Err.Description & " (Err=" & Err.Number & ")"
		Exit Sub
	End If

	If UCase(Trim(gMode)) = "MONTHLY" Then
		endDateObj = ResolveMonthlyAsOfDate(conn, startDateObj, endDateObj)
	End If

	priorDateObj = DateAdd("d", -1, DateSerial(Year(endDateObj), Month(endDateObj), 1))
	
	priorDateObj = endDateObj
    ' =====================================================
    ' SECTION 1: INCOME STATEMENT BUDGET VARIANCE
    '   Current MTD/YTD from dbo.GLBudgetReport
    '   Prior actual from dbo.GLIncomeStatement
    ' =====================================================
    sqlIS = BuildIncomeStatementVarianceSql()

    Set cmdIS = CreateObject("ADODB.Command")
    With cmdIS
        .ActiveConnection = conn
        .CommandText = sqlIS
        .CommandType = 1 ' adCmdText
        .Parameters.Append .CreateParameter("", 133, 1, , endDateObj)    ' @AsOf
        .Parameters.Append .CreateParameter("", 133, 1, , priorDateObj)  ' @Prior
    End With

    Err.Clear
    Set rsIS = cmdIS.Execute
    If Err.Number <> 0 Then
        LogLine "ERROR executing Income Statement variance query: " & Err.Description & " (Err=" & Err.Number & ")"
        CleanupObjects Nothing, rsIS, cmdIS
        CleanupObjects conn, rs, Nothing
        Exit Sub
    End If

    fout.WriteLine "Income Statement Budget Variance "  & Fmt(endDateObj, "mmm d, yyyy") & " | Prior " & Fmt(priorDateObj, "mmm d, yyyy")
    WriteHeaderRow

    If rsIS.EOF Then
        fout.WriteLine "No Income Statement rows returned for =" & FmtDate8(endDateObj) & " Prior=" & FmtDate8(priorDateObj)
    Else
        rowCount = 0
        Do While Not rsIS.EOF
            rowCount = rowCount + 1
            WriteDataRow rsIS
            rsIS.MoveNext
        Loop

        fout.WriteLine ""
        'fout.WriteLine "Total Income Statement Rows: " & rowCount
    End If

    CleanupObjects Nothing, rsIS, cmdIS

    ' =====================================================
    ' SECTION 2: BALANCE SHEET BUDGET VARIANCE
    '   Current MTD/YTD from dbo.GLBudgetReport
    '   Prior actual from dbo.GLBalanceSheet
    ' =====================================================

    fout.WriteLine "BALANCE SHEET    BUDGET VARIANCE " & Fmt(endDateObj, "mmm d, yyyy") & _
                   " | Prior " & Fmt(priorDateObj, "mmm d, yyyy")

    sqlBS = BuildBalanceSheetVarianceSql()

    Set cmdBS = CreateObject("ADODB.Command")
    With cmdBS
        .ActiveConnection = conn
        .CommandText = sqlBS
        .CommandType = 1 ' adCmdText
        .Parameters.Append .CreateParameter("", 133, 1, , endDateObj)    ' @AsOf
        .Parameters.Append .CreateParameter("", 133, 1, , priorDateObj)  ' @Prior
    End With

    Err.Clear
    Set rsBS = cmdBS.Execute
    If Err.Number <> 0 Then
        LogLine "ERROR executing Balance Sheet variance query: " & Err.Description & " (Err=" & Err.Number & ")"
        CleanupObjects Nothing, rsBS, cmdBS
        CleanupObjects conn, rs, Nothing
        Exit Sub
    End If

    WriteHeaderRow

    If rsBS.EOF Then
        fout.WriteLine "No Balance Sheet rows returned for AsOf=" & FmtDate8(endDateObj) & " Prior=" & FmtDate8(priorDateObj)
    Else
        bsRowCount = 0
        Do While Not rsBS.EOF
            bsRowCount = bsRowCount + 1
            WriteDataRow rsBS
            rsBS.MoveNext
        Loop

        fout.WriteLine ""
        fout.WriteLine "Total Balance Sheet Rows: " & bsRowCount
    End If

    CleanupObjects Nothing, rsBS, cmdBS
    CleanupObjects conn, rs, Nothing

    On Error GoTo 0
End Sub

Sub ProcessLine
End Sub

Sub CloseDoc
End Sub

' =========================================================
' SQL builders
' =========================================================

Function BuildIncomeStatementVarianceSql()
    Dim sql, nActual, nBudget, nAmount

    nActual = SqlNormalizeNumber("Actual")
    nBudget = SqlNormalizeNumber("Budget")
    nAmount = SqlNormalizeNumber("Amount")

    sql = ""
    sql = sql & "DECLARE @AsOf date = ?; " & vbCrLf
    sql = sql & "DECLARE @Prior date = ?; " & vbCrLf
    sql = sql & "WITH " & vbCrLf

    sql = sql & "mtd AS ( " & vbCrLf
    sql = sql & "  SELECT " & vbCrLf
    sql = sql & "    GLAccount, " & vbCrLf
    sql = sql & "    MAX(GLAccountName) AS GLAccountName, " & vbCrLf
    sql = sql & "    ActualMTD = MAX(" & nActual & "), " & vbCrLf
    sql = sql & "    BudgetMTD = MAX(" & nBudget & ") " & vbCrLf
    sql = sql & "  FROM dbo.GLBudgetReport " & vbCrLf
    sql = sql & "  WHERE EffectiveDate = @Prior  AND UPPER(ISNULL(Prd,'')) = 'MTD' " & vbCrLf
    sql = sql & "  GROUP BY GLAccount " & vbCrLf
    sql = sql & "), " & vbCrLf

    sql = sql & "ytd AS ( " & vbCrLf
    sql = sql & "  SELECT " & vbCrLf
    sql = sql & "    GLAccount, " & vbCrLf
    sql = sql & "    MAX(GLAccountName) AS GLAccountName, " & vbCrLf
    sql = sql & "    ActualYTD = MAX(" & nActual & "), " & vbCrLf
    sql = sql & "    BudgetYTD = MAX(" & nBudget & ") " & vbCrLf
    sql = sql & "  FROM dbo.GLBudgetReport " & vbCrLf
    sql = sql & "  WHERE EffectiveDate = @Prior  AND UPPER(ISNULL(Prd,'')) = 'YTD' " & vbCrLf
    sql = sql & "  GROUP BY GLAccount " & vbCrLf
    sql = sql & "), " & vbCrLf

    sql = sql & "prior AS ( " & vbCrLf
    sql = sql & "  SELECT " & vbCrLf
    sql = sql & "    GLAccount, " & vbCrLf
    sql = sql & "    MAX(GLAccountName) AS GLAccountName, " & vbCrLf
    sql = sql & "    ActualPriorMTD = MAX(" & nAmount & ") " & vbCrLf
    sql = sql & "  FROM dbo.GLIncomeStatement " & vbCrLf
    sql = sql & "  WHERE EffectiveDate = @Prior " & vbCrLf
    sql = sql & "  GROUP BY GLAccount " & vbCrLf
    sql = sql & "), " & vbCrLf

    sql = sql & "keys AS ( " & vbCrLf
    sql = sql & "  SELECT GLAccount, MAX(GLAccountName) AS GLAccountName " & vbCrLf
    sql = sql & "  FROM ( " & vbCrLf
    sql = sql & "    SELECT GLAccount, GLAccountName FROM mtd " & vbCrLf
    sql = sql & "    UNION ALL SELECT GLAccount, GLAccountName FROM ytd " & vbCrLf
    sql = sql & "    UNION ALL SELECT GLAccount, GLAccountName FROM prior " & vbCrLf
    sql = sql & "  ) k " & vbCrLf
    sql = sql & "  GROUP BY GLAccount " & vbCrLf
    sql = sql & ") " & vbCrLf

    sql = sql & "SELECT " & vbCrLf
    sql = sql & "  k.GLAccount AS [GLAccount], " & vbCrLf
    sql = sql & "  k.GLAccountName AS [GLAccountName], " & vbCrLf
    sql = sql & "  m.ActualMTD, " & vbCrLf
    sql = sql & "  m.BudgetMTD, " & vbCrLf
    sql = sql & "  (ISNULL(m.ActualMTD,0) - ISNULL(m.BudgetMTD,0)) AS DiffMTD, " & vbCrLf
    sql = sql & "  CASE WHEN NULLIF(m.BudgetMTD,0) IS NULL THEN NULL ELSE (ISNULL(m.ActualMTD,0) - ISNULL(m.BudgetMTD,0))/NULLIF(m.BudgetMTD,0) END AS PctDiffMTD, " & vbCrLf
    sql = sql & "  y.ActualYTD, " & vbCrLf
    sql = sql & "  y.BudgetYTD, " & vbCrLf
    sql = sql & "  (ISNULL(y.ActualYTD,0) - ISNULL(y.BudgetYTD,0)) AS DiffYTD, " & vbCrLf
    sql = sql & "  CASE WHEN NULLIF(y.BudgetYTD,0) IS NULL THEN NULL ELSE (ISNULL(y.ActualYTD,0) - ISNULL(y.BudgetYTD,0))/NULLIF(y.BudgetYTD,0) END AS PctDiffYTD, " & vbCrLf
    sql = sql & "  p.ActualPriorMTD, " & vbCrLf
    sql = sql & "  (ISNULL(m.ActualMTD,0) - ISNULL(p.ActualPriorMTD,0)) AS DiffPriorMTD, " & vbCrLf
    sql = sql & "  CASE WHEN NULLIF(p.ActualPriorMTD,0) IS NULL THEN NULL ELSE (ISNULL(m.ActualMTD,0) - ISNULL(p.ActualPriorMTD,0))/NULLIF(p.ActualPriorMTD,0) END AS PctDiffPriorMTD " & vbCrLf
    sql = sql & "FROM keys k " & vbCrLf
    sql = sql & "LEFT JOIN mtd m ON m.GLAccount = k.GLAccount " & vbCrLf
    sql = sql & "LEFT JOIN ytd y ON y.GLAccount = k.GLAccount " & vbCrLf
    sql = sql & "LEFT JOIN prior p ON p.GLAccount = k.GLAccount " & vbCrLf
	sql = sql & "ORDER BY TRY_CONVERT(bigint, CASE WHEN k.GLAccount NOT LIKE '%[^0-9-]%' THEN REPLACE(k.GLAccount,'-','') ELSE NULL END), k.GLAccount;"

    BuildIncomeStatementVarianceSql = sql
End Function

Function BuildBalanceSheetVarianceSql()
    Dim sql, nActual, nBudget, nAmount

    nActual = SqlNormalizeNumber("Actual")
    nBudget = SqlNormalizeNumber("Budget")
    nAmount = SqlNormalizeNumber("Amount")

    sql = ""
    sql = sql & "DECLARE @AsOf date = ?; " & vbCrLf
    sql = sql & "DECLARE @Prior date = ?; " & vbCrLf
    sql = sql & "WITH " & vbCrLf

    sql = sql & "mtd AS ( " & vbCrLf
    sql = sql & "  SELECT " & vbCrLf
    sql = sql & "    GLAccount, " & vbCrLf
    sql = sql & "    MAX(GLAccountName) AS GLAccountName, " & vbCrLf
    sql = sql & "    ActualMTD = MAX(" & nActual & "), " & vbCrLf
    sql = sql & "    BudgetMTD = MAX(" & nBudget & ") " & vbCrLf
    sql = sql & "  FROM dbo.GLBudgetReport " & vbCrLf
    sql = sql & "  WHERE EffectiveDate = @Prior AND UPPER(ISNULL(Prd,'')) = 'MTD' " & vbCrLf
    sql = sql & "  GROUP BY GLAccount " & vbCrLf
    sql = sql & "), " & vbCrLf

    sql = sql & "ytd AS ( " & vbCrLf
    sql = sql & "  SELECT " & vbCrLf
    sql = sql & "    GLAccount, " & vbCrLf
    sql = sql & "    MAX(GLAccountName) AS GLAccountName, " & vbCrLf
    sql = sql & "    ActualYTD = MAX(" & nActual & "), " & vbCrLf
    sql = sql & "    BudgetYTD = MAX(" & nBudget & ") " & vbCrLf
    sql = sql & "  FROM dbo.GLBudgetReport " & vbCrLf
    sql = sql & "  WHERE EffectiveDate = @Prior AND UPPER(ISNULL(Prd,'')) = 'YTD' " & vbCrLf
    sql = sql & "  GROUP BY GLAccount " & vbCrLf
    sql = sql & "), " & vbCrLf

    sql = sql & "prior AS ( " & vbCrLf
    sql = sql & "  SELECT " & vbCrLf
    sql = sql & "    GLAccount, " & vbCrLf
    sql = sql & "    MAX(GLAccountName) AS GLAccountName, " & vbCrLf
    sql = sql & "    ActualPriorMTD = MAX(" & nAmount & ") " & vbCrLf
    sql = sql & "  FROM dbo.GLBalanceSheet " & vbCrLf
    sql = sql & "  WHERE RunDate = @AsOf " & vbCrLf
    sql = sql & "  GROUP BY GLAccount " & vbCrLf
    sql = sql & "), " & vbCrLf

    sql = sql & "keys AS ( " & vbCrLf
    sql = sql & "  SELECT GLAccount, MAX(GLAccountName) AS GLAccountName " & vbCrLf
    sql = sql & "  FROM ( " & vbCrLf
    sql = sql & "    SELECT GLAccount, GLAccountName FROM mtd " & vbCrLf
    sql = sql & "    UNION ALL SELECT GLAccount, GLAccountName FROM ytd " & vbCrLf
    sql = sql & "    UNION ALL SELECT GLAccount, GLAccountName FROM prior " & vbCrLf
    sql = sql & "  ) k " & vbCrLf
    sql = sql & "  GROUP BY GLAccount " & vbCrLf
    sql = sql & ") " & vbCrLf

    sql = sql & "SELECT " & vbCrLf
    sql = sql & "  k.GLAccount AS [GLAccount], " & vbCrLf
    sql = sql & "  k.GLAccountName AS [GLAccountName], " & vbCrLf
    sql = sql & "  m.ActualMTD, " & vbCrLf
    sql = sql & "  m.BudgetMTD, " & vbCrLf
    sql = sql & "  (ISNULL(m.ActualMTD,0) - ISNULL(m.BudgetMTD,0)) AS DiffMTD, " & vbCrLf
    sql = sql & "  CASE WHEN NULLIF(m.BudgetMTD,0) IS NULL THEN NULL ELSE (ISNULL(m.ActualMTD,0) - ISNULL(m.BudgetMTD,0))/NULLIF(m.BudgetMTD,0) END AS PctDiffMTD, " & vbCrLf
    sql = sql & "  y.ActualYTD, " & vbCrLf
    sql = sql & "  y.BudgetYTD, " & vbCrLf
    sql = sql & "  (ISNULL(y.ActualYTD,0) - ISNULL(y.BudgetYTD,0)) AS DiffYTD, " & vbCrLf
    sql = sql & "  CASE WHEN NULLIF(y.BudgetYTD,0) IS NULL THEN NULL ELSE (ISNULL(y.ActualYTD,0) - ISNULL(y.BudgetYTD,0))/NULLIF(y.BudgetYTD,0) END AS PctDiffYTD, " & vbCrLf
    sql = sql & "  p.ActualPriorMTD, " & vbCrLf
    sql = sql & "  (ISNULL(m.ActualMTD,0) - ISNULL(p.ActualPriorMTD,0)) AS DiffPriorMTD, " & vbCrLf
    sql = sql & "  CASE WHEN NULLIF(p.ActualPriorMTD,0) IS NULL THEN NULL ELSE (ISNULL(m.ActualMTD,0) - ISNULL(p.ActualPriorMTD,0))/NULLIF(p.ActualPriorMTD,0) END AS PctDiffPriorMTD " & vbCrLf
    sql = sql & "FROM keys k " & vbCrLf
    sql = sql & "LEFT JOIN mtd m ON m.GLAccount = k.GLAccount " & vbCrLf
    sql = sql & "LEFT JOIN ytd y ON y.GLAccount = k.GLAccount " & vbCrLf
    sql = sql & "LEFT JOIN prior p ON p.GLAccount = k.GLAccount " & vbCrLf
    sql = sql & "ORDER BY TRY_CONVERT(bigint, CASE WHEN k.GLAccount NOT LIKE '%[^0-9-]%' THEN REPLACE(k.GLAccount,'-','') ELSE NULL END), k.GLAccount;"

    BuildBalanceSheetVarianceSql = sql
End Function

Function SqlNormalizeNumber(ByVal colName)
    Dim s
    s = ""
    s = s & "TRY_CONVERT(decimal(18,2), " & _
            "CASE " & _
            "WHEN LTRIM(RTRIM(ISNULL(" & colName & ",''))) = '' THEN NULL " & _
            "WHEN LTRIM(RTRIM(ISNULL(" & colName & ",''))) LIKE '<%>' THEN " & _
            "  '-' + REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(LTRIM(RTRIM(ISNULL(" & colName & ",''))), '<',''), '>',''), '$',''), ',',''), ' ','') " & _
            "ELSE " & _
            "  REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(LTRIM(RTRIM(ISNULL(" & colName & ",''))), '$',''), ',',''), '(', '-'), ')',''), ' ','') , '>', '') " & _
            "END" & _
            ")"
    SqlNormalizeNumber = s
End Function

' =========================================================
' Output helpers
' =========================================================

Sub WriteHeaderRow()
    fout.WriteLine _
        Pad("GLAccount", 16) & _
        Pad("GL Account Name", 54) & _
        Pad("Actual MTD", 20) & _
        Pad("Budget MTD", 16) & _
        Pad("Diff", 16) & _
        Pad("%Diff", 10) & _
        Pad("Actual YTD", 20) & _
        Pad("Budget YTD", 20) & _
        Pad("Diff", 16) & _
        Pad("%Diff", 10) & _
        Pad("Actual Prior", 20) & _
        Pad("Diff", 16) & _
        Pad("%Diff", 10)
End Sub

Sub WriteDataRow(ByRef r)
    fout.WriteLine _
        Pad(r("GLAccount"), 16) & _
        Pad(r("GLAccountName"), 54) & _
        Pad(FmtMoney(r("ActualMTD")), 20) & _
        Pad(FmtMoney(r("BudgetMTD")), 16) & _
        Pad(FmtMoney(r("DiffMTD")), 16) & _
        Pad(FmtPct100(r("PctDiffMTD")), 10) & _
        Pad(FmtMoney(r("ActualYTD")), 20) & _
        Pad(FmtMoney(r("BudgetYTD")), 20) & _
        Pad(FmtMoney(r("DiffYTD")), 16) & _
        Pad(FmtPct100(r("PctDiffYTD")), 10) & _
        Pad(FmtMoney(r("ActualPriorMTD")), 20) & _
        Pad(FmtMoney(r("DiffPriorMTD")), 16) & _
        Pad(FmtPct100(r("PctDiffPriorMTD")), 10)
End Sub
' =========================================================
' Generic helpers
' =========================================================

Sub CleanupObjects(ByRef c, ByRef r, ByRef cm)
    On Error Resume Next
    If Not r Is Nothing Then
        If r.State = 1 Then r.Close
        Set r = Nothing
    End If
    If Not cm Is Nothing Then Set cm = Nothing
    If Not c Is Nothing Then
        If c.State = 1 Then c.Close
        Set c = Nothing
    End If
    On Error GoTo 0
End Sub

Sub LoadParameters(iniPath)
    Dim cfg, modeUpper
    Set cfg = ReadIniSafe(iniPath)
    If cfg Is Nothing Then
        LogLine "ERROR: Could not read INI file: " & iniPath
        Exit Sub
    End If

    gMode         = GetIniValue(cfg, "Report.Mode", "MONTHLY")
    gRelOffset    = GetIniValue(cfg, "Report.RelativeOffset", "0")
    gYear         = GetIniValue(cfg, "Report.Year", "")
    gMonth        = GetIniValue(cfg, "Report.Month", "")
    gQuarter      = GetIniValue(cfg, "Report.Quarter", "")
    gStartDateStr = GetIniValue(cfg, "Report.StartDate", "")
    gEndDateStr   = GetIniValue(cfg, "Report.EndDate", "")

    gConnStr       = GetIniValue(cfg, "Database.ConnStr", "")
    gFullTableName = GetIniValue(cfg, "Database.TableName", "")

    If IsNumeric(gRelOffset) Then
        gRelOffset = CInt(gRelOffset)
    Else
        LogLine "ERROR: Report.RelativeOffset must be numeric. Value=[" & gRelOffset & "]"
        gRelOffset = 0
    End If

    modeUpper = UCase(Trim(gMode))
    Select Case modeUpper
        Case "CUSTOM", "DAILY", "WEEKLY", "QUARTERLY", "YTD", "MTD", "MONTHLY"
            ' valid
        Case Else
            LogLine "ERROR: Invalid Report.Mode value: [" & gMode & "]"
            gMode = ""
            Exit Sub
    End Select

    If Trim(gConnStr) = "" Then
        LogLine "ERROR: Database.ConnStr is blank."
        Exit Sub
    End If

    If Trim(gFullTableName) = "" Then
        LogLine "ERROR: Database.TableName is blank."
        Exit Sub
    End If
End Sub
Function GetIniValue(cfg, key, defaultValue)
    If Not cfg Is Nothing And cfg.Exists(key) And Len(cfg(key)) > 0 Then
        GetIniValue = Trim(cfg(key))
    Else
        GetIniValue = defaultValue
    End If
End Function

Sub ComputeDateRange(ByRef startKey, ByRef endKey)
    Dim mode, today, dStart, dEnd
    Dim y, m, q, tmp, dow, weekStart, baseDate

    mode  = UCase(Trim(gMode))
    today = Date()

    Select Case mode
        Case "CUSTOM"
            If gStartDateStr <> "" Then dStart = CDate(gStartDateStr) Else dStart = today
            If gEndDateStr <> "" Then dEnd = CDate(gEndDateStr) Else dEnd = today

        Case "DAILY"
            dStart = DateAdd("d", gRelOffset, today)
            dEnd = dStart

        Case "WEEKLY"
            dow = Weekday(today, vbMonday)
            weekStart = DateAdd("d", 1 - dow, today)
            If gRelOffset <> 0 Then weekStart = DateAdd("ww", gRelOffset, weekStart)
            dStart = weekStart
            dEnd = DateAdd("d", 6, dStart)

        Case "QUARTERLY"
            If gYear <> "" And IsNumeric(gYear) Then y = CInt(gYear) Else y = Year(today)
            If gQuarter <> "" And IsNumeric(gQuarter) Then q = CInt(gQuarter) Else q = DatePart("q", today)

            If gRelOffset <> 0 Then
                q = q + gRelOffset
                While q > 4
                    q = q - 4
                    y = y + 1
                Wend
                While q < 1
                    q = q + 4
                    y = y - 1
                Wend
            End If

            m = (q - 1) * 3 + 1
            dStart = DateSerial(y, m, 1)
            dEnd = DateAdd("d", -1, DateAdd("m", 3, dStart))

        Case "YTD"
            If gYear <> "" And IsNumeric(gYear) Then y = CInt(gYear) Else y = Year(today)
            dStart = DateSerial(y, 1, 1)
            dEnd = today

        Case "MTD"
            baseDate = today
            If gRelOffset <> 0 Then baseDate = DateAdd("m", gRelOffset, baseDate)

            dStart = DateSerial(Year(baseDate), Month(baseDate), 1)

            If Year(baseDate) = Year(today) And Month(baseDate) = Month(today) Then
                dEnd = today
            Else
                dEnd = DateAdd("d", -1, DateAdd("m", 1, dStart))
            End If

        Case Else ' MONTHLY
            If gYear <> "" And IsNumeric(gYear) Then y = CInt(gYear) Else y = Year(today)
            If gMonth <> "" And IsNumeric(gMonth) Then m = CInt(gMonth) Else m = Month(today)

            If gRelOffset <> 0 Then
                tmp = DateAdd("m", gRelOffset, DateSerial(y, m, 1))
                y = Year(tmp)
                m = Month(tmp)
            End If

            dStart = DateSerial(y, m, 1)
            dEnd = DateAdd("d", -1, DateAdd("m", 1, dStart))
    End Select

    If mode <> "CUSTOM" And mode <> "DAILY" And dEnd > today Then dEnd = today

    startKey = FormatDateKey(dStart)
    endKey   = FormatDateKey(dEnd)
End Sub

Function FormatDateKey(d)
    FormatDateKey = CStr(Year(d)) & Right("0" & Month(d), 2) & Right("0" & Day(d), 2)
End Function

Function FormatKeyToDate(ByVal yyyymmdd)
    Dim y, m, d
    If Len(yyyymmdd) <> 8 Then
        FormatKeyToDate = Date()
        Exit Function
    End If
    y = CInt(Left(yyyymmdd, 4))
    m = CInt(Mid(yyyymmdd, 5, 2))
    d = CInt(Right(yyyymmdd, 2))
    FormatKeyToDate = DateSerial(y, m, d)
End Function

Function Pad(val, width)
    Dim s
    If IsNull(val) Then s = "" Else s = CStr(val)
    s = Trim(s)
    Pad = Left(s & Space(width), width)
End Function

Function FmtDate8(val)
    On Error Resume Next
    If IsNull(val) Or Len(Trim(CStr(val))) = 0 Then
        FmtDate8 = ""
    ElseIf IsDate(val) Then
        FmtDate8 = Right("0" & Month(CDate(val)), 2) & "/" & _
                   Right("0" & Day(CDate(val)), 2) & "/" & _
                   CStr(Year(CDate(val)))
    Else
        FmtDate8 = Trim(CStr(val))
    End If
    On Error GoTo 0
End Function

Function FmtTimeShort(val)
    On Error Resume Next
    If IsNull(val) Then
        FmtTimeShort = ""
    Else
        FmtTimeShort = Trim(CStr(val))
    End If
    On Error GoTo 0
End Function

Function FileExists(path)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(path)
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
        schemaName = Trim(parts(UBound(parts) - 1))
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

Function Fmt(d, mask)
    On Error Resume Next

    Dim s
    s = ""

    s = Format(d, mask)
    If Err.Number = 0 And Len(Trim(CStr(s))) > 0 Then
        Fmt = CStr(s)
        Exit Function
    End If

    Err.Clear
    If Not IsDate(d) Then
        Fmt = ""
        Exit Function
    End If

    Select Case LCase(mask)
        Case "dddd, mmmm d, yyyy"
            Fmt = WeekdayName(Weekday(d), False) & ", " & MonthName(Month(d), False) & " " & Day(d) & ", " & Year(d)
        Case "mmm d"
            Fmt = MonthName(Month(d), True) & " " & Day(d)
        Case "mmm d, yyyy"
            Fmt = MonthName(Month(d), True) & " " & Day(d) & ", " & Year(d)
        Case "mmmm yyyy"
            Fmt = MonthName(Month(d), False) & " " & Year(d)
        Case Else
            Fmt = MonthName(Month(d), True) & " " & Day(d) & ", " & Year(d)
    End Select

    On Error GoTo 0
End Function

Function FmtMoney(val)
    On Error Resume Next
    Dim n
    If IsNull(val) Or Len(Trim(CStr(val))) = 0 Then
        FmtMoney = ""
        Exit Function
    End If

    n = CDbl(val)
    If n < 0 Then
        FmtMoney = "(" & Replace(FormatNumber(Abs(n), 2), "-", "") & ")"
    Else
        FmtMoney = FormatNumber(n, 2)
    End If
    On Error GoTo 0
End Function

Function FmtPct100(val)
    On Error Resume Next
    If IsNull(val) Or Len(Trim(CStr(val))) = 0 Then
        FmtPct100 = ""
        Exit Function
    End If
    FmtPct100 = FormatNumber(CDbl(val) * 100, 2)
    On Error GoTo 0
End Function


Function ResolveMonthlyAsOfDate(ByRef c, ByVal monthStart, ByVal monthEnd)
    On Error Resume Next

    Dim cmd, rs, sql, resolvedDate
    resolvedDate = monthEnd

    sql = ""
    sql = sql & "DECLARE @MonthStart date = ?; "
    sql = sql & "DECLARE @MonthEnd   date = ?; "
    sql = sql & "WITH d AS ( "
    sql = sql & "    SELECT EffectiveDate AS Dt "
    sql = sql & "    FROM dbo.GLIncomeStatement "
    sql = sql & "    WHERE EffectiveDate >= @MonthStart AND EffectiveDate <= @MonthEnd "
    sql = sql & "    UNION ALL "
    sql = sql & "    SELECT RunDate AS Dt "
    sql = sql & "    FROM dbo.GLBalanceSheet "
    sql = sql & "    WHERE RunDate >= @MonthStart AND RunDate <= @MonthEnd "
    sql = sql & ") "
    sql = sql & "SELECT MAX(Dt) AS AsOfDate FROM d;"

    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = c
        .CommandText = sql
        .CommandType = 1
        .Parameters.Append .CreateParameter("", 133, 1, , monthStart)
        .Parameters.Append .CreateParameter("", 133, 1, , monthEnd)
    End With

    Err.Clear
    Set rs = cmd.Execute

    If Err.Number <> 0 Then
        LogLine "ERROR resolving monthly AsOf date: " & Err.Description & " (Err=" & Err.Number & ")"
        Err.Clear
    Else
        If Not rs.EOF Then
            If Not IsNull(rs("AsOfDate")) Then
                resolvedDate = CDate(rs("AsOfDate"))
            End If
        End If
    End If

    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Set cmd = Nothing

    ResolveMonthlyAsOfDate = resolvedDate
    On Error GoTo 0
End Function