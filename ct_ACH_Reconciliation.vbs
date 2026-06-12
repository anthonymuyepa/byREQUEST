Option Explicit

Const adCmdText    = 1
Const adParamInput = 1
Const adDBDate     = 133

Dim gMode, gRelOffset, gYear, gMonth, gQuarter
Dim gStartDateStr, gEndDateStr
Dim gConnStr, gFullTableName, gTable2Part, gObjIdName

Dim conn
Dim INI_PATH
INI_PATH = "S:\CarolinaTrust\byREQUEST\Templates\ACH_Exceptions_Reconciliationini.ini"

Sub StartDoc()
    On Error Resume Next

    Dim startKey, endKey
    Dim startDateObj, endDateObj, periodDisplay
    Dim cmdSum, rsSum
    Dim cmdMatched, rsMatched
    Dim cmdUnmatchedEx, rsUnmatchedEx
    Dim cmdUnmatchedList, rsUnmatchedList
    Dim sqlSum, sqlMatched, sqlUnmatchedEx, sqlUnmatchedList
	Dim cmdControlSummary, rsControlSummary
	Dim sqlControlSummary

    If Not FileExists(INI_PATH) Then
        LogLine "ERROR: INI file not found: " & INI_PATH
        Exit Sub
    End If

    LoadParameters INI_PATH

    If Trim(gConnStr) = "" Then
        LogLine "ERROR: ConnStr not defined in [Database] section."
        Exit Sub
    End If

    BuildTableNames gFullTableName, gTable2Part, gObjIdName

    ComputeDateRange startKey, endKey
    If startKey = "" Or endKey = "" Then
        LogLine "ERROR: Could not compute date range."
        Exit Sub
    End If

    startDateObj = FormatKeyToDate(startKey)
    endDateObj   = FormatKeyToDate(endKey)

    periodDisplay = FmtDate8(startDateObj) & " to " & FmtDate8(endDateObj)

    SpoolFile.User1 = "ACHRECON-" & _
                      Right("0" & Month(endDateObj), 2) & _
                      Right("0" & Day(endDateObj), 2) & _
                      Year(endDateObj)

    Set conn = CreateObject("ADODB.Connection")
    Err.Clear
    conn.Open gConnStr

    If Err.Number <> 0 Then
        LogLine "ERROR opening DB connection: " & Err.Description & " (Err=" & Err.Number & ")"
        Err.Clear
        Exit Sub
    End If

    fout.WriteLine "ACH EXCEPTIONS RECONCILIATION REPORT"
    fout.WriteLine "Period: " & periodDisplay
    fout.WriteLine ""

    ' =====================================================
    ' 1) SUMMARY
    ' =====================================================
    sqlSum = _
    "WITH list_norm AS ( " & _
    "   SELECT LTRIM(RTRIM(TraceNumber)) AS TraceKey " & _
    "   FROM dbo.ACH_LIST " & _
    "   WHERE TraceNumber IS NOT NULL " & _
    "     AND RptDate BETWEEN ? AND ? " & _
    "   GROUP BY LTRIM(RTRIM(TraceNumber)) " & _
    "), ex_norm AS ( " & _
    "   SELECT 'NSF' AS ExceptionType, RptDate, LTRIM(RTRIM(TraceNumber)) AS TraceKey, Amount " & _
    "   FROM dbo.ACH_Exceptions_NSF WHERE TraceNumber IS NOT NULL AND RptDate BETWEEN ? AND ? " & _
    "   UNION ALL " & _
    "   SELECT 'TranCode', RptDate, LTRIM(RTRIM(TraceNumber)), Amount " & _
    "   FROM dbo.ACH_Exceptions_TranCode WHERE TraceNumber IS NOT NULL AND RptDate BETWEEN ? AND ? " & _
    "   UNION ALL " & _
    "   SELECT 'Exceptions', RptDate, LTRIM(RTRIM(TraceNumber)), Amount " & _
    "   FROM dbo.ACH_Exceptions WHERE TraceNumber IS NOT NULL AND RptDate BETWEEN ? AND ? " & _
    ") " & _
    "SELECT " & _
    "   e.ExceptionType, " & _
    "   COUNT(*) AS ExceptionCount, " & _
    "   SUM(CASE WHEN l.TraceKey IS NOT NULL THEN 1 ELSE 0 END) AS MatchedToACHList, " & _
    "   SUM(CASE WHEN l.TraceKey IS NULL THEN 1 ELSE 0 END) AS NotFoundInACHList, " & _
    "   SUM(TRY_CONVERT(decimal(18,2), REPLACE(REPLACE(LTRIM(RTRIM(e.Amount)), ',', ''), '$', ''))) AS ExceptionAmount " & _
    "FROM ex_norm e " & _
    "LEFT JOIN list_norm l ON l.TraceKey = e.TraceKey " & _
    "GROUP BY e.ExceptionType " & _
    "ORDER BY e.ExceptionType"

    Set cmdSum = CreateCommandWithDates(sqlSum, startDateObj, endDateObj, 4)
    Set rsSum = cmdSum.Execute

    fout.WriteLine "REPORT SUMMARY BY EXCEPTION TYPE"
    fout.WriteLine Pad("ExcepType", 22) & _
                   Pad("Count", 15) & _
                   Pad("Matched", 15) & _
                   Pad("External", 15) & _
                   Pad("Amount", 24)

    If rsSum.EOF Then
        fout.WriteLine "No exception records found for this period."
    Else
        Do While Not rsSum.EOF
            fout.WriteLine _
                Pad(rsSum("ExceptionType"), 22) & _
                Pad(rsSum("ExceptionCount"), 15) & _
                Pad(rsSum("MatchedToACHList"), 15) & _
                Pad(rsSum("NotFoundInACHList"), 15) & _
                Pad(FormatAmount(rsSum("ExceptionAmount")), 24)

            rsSum.MoveNext
        Loop
    End If

    fout.WriteLine ""
	
	
	fout.WriteLine ""

	' =====================================================
	' 1B) CONTROL SUMMARY / VARIANCE
	' =====================================================
	sqlControlSummary = _
	"SELECT " & _
	"   'NSF' AS Category, " & _
	"   ISNULL(c.ControlCount,0) AS ControlCount, " & _
	"   ISNULL(c.ControlAmount,0) AS ControlAmount, " & _
	"   ISNULL(r.ReconCount,0) AS ReconCount, " & _
	"   ISNULL(r.ReconAmount,0) AS ReconAmount, " & _
	"   ISNULL(c.ControlCount,0) - ISNULL(r.ReconCount,0) AS CountVariance, " & _
	"   ISNULL(c.ControlAmount,0) - ISNULL(r.ReconAmount,0) AS AmountVariance " & _
	"FROM " & _
	"(SELECT " & _
	"   SUM(ISNULL(UnpostedCount,0)) AS ControlCount, " & _
	"   SUM(ISNULL(UnpostedAmount,0)) AS ControlAmount " & _
	" FROM dbo.ACH_ControlTotals " & _
	" WHERE RptDate BETWEEN ? AND ? " & _
	"   AND Category = 'NSF') c " & _
	"CROSS JOIN " & _
	"(SELECT " & _
	"   COUNT(*) AS ReconCount, " & _
	"   SUM(TRY_CONVERT(decimal(18,2), REPLACE(REPLACE(LTRIM(RTRIM(Amount)), ',', ''), '$', ''))) AS ReconAmount " & _
	" FROM dbo.ACH_Exceptions_NSF " & _
	" WHERE RptDate BETWEEN ? AND ?) r"

	Set cmdControlSummary = CreateCommandWithDates(sqlControlSummary, startDateObj, endDateObj, 2)
	Set rsControlSummary = cmdControlSummary.Execute
	If Err.Number <> 0 Then
		fout.WriteLine "ERROR executing control summary: " & Err.Description
		Err.Clear
	End If

	fout.WriteLine "REPORT CONTROL SUMMARY / VARIANCE"
	fout.WriteLine Pad("Category", 22) & _
				   Pad("ControlCnt", 15) & _
				   Pad("ContrAmt", 15) & _
				   Pad("ReconCnt", 15) & _
				   Pad("ReconAmt", 15) & _
				   Pad("CountVar", 15) & _
				   Pad("AmountVar", 15)

	If rsControlSummary.EOF Then
		fout.WriteLine "No control summary records found for this period."
	Else
		Do While Not rsControlSummary.EOF
			fout.WriteLine _
				Pad(rsControlSummary("Category"), 22) & _
				Pad(rsControlSummary("ControlCount"), 15) & _
				Pad(FormatAmount(rsControlSummary("ControlAmount")), 15) & _
				Pad(rsControlSummary("ReconCount"), 15) & _
				Pad(FormatAmount(rsControlSummary("ReconAmount")), 15) & _
				Pad(rsControlSummary("CountVariance"), 15) & _
				Pad(FormatAmount(rsControlSummary("AmountVariance")), 15)

			rsControlSummary.MoveNext
		Loop
	End If
	
	
	fout.WriteLine ""
	
	
	fout.WriteLine ""

    ' =====================================================
    ' 2) MATCHED EXCEPTIONS TO ACH_LIST
    ' =====================================================
    sqlMatched = _
    "WITH list_norm AS ( " & _
    "   SELECT " & _
    "       LTRIM(RTRIM(TraceNumber)) AS TraceKey, " & _
    "       MIN(RptDate) AS RptDate, " & _
    "       MAX(CompanyName) AS ACHCompanyName, " & _
    "       MAX(CompanyID) AS ACHCompanyID, " & _
    "       MAX(ECC) AS ACHECC, " & _
    "       MAX(TranCode) AS ACHTranCode, " & _
    "       MAX(IndividualAcct) AS ACHAcct, " & _
    "       MAX(IndividualName) AS ACHName, " & _
    "       MAX(Description) AS ACHDescription, " & _
    "       MAX(TRY_CONVERT(decimal(18,2), REPLACE(REPLACE(LTRIM(RTRIM(Amount)), ',', ''), '$', ''))) AS ACHAmount " & _
    "   FROM dbo.ACH_LIST " & _
    "   WHERE TraceNumber IS NOT NULL " & _
    "     AND RptDate BETWEEN ? AND ? " & _
    "   GROUP BY LTRIM(RTRIM(TraceNumber)) " & _
    "), ex_norm AS ( " & _
    "   SELECT 'NSF' AS ExceptionType, RowID AS ExceptionId, RptDate, RptTime, RptSeq, Acct, SLX, [ID], [Name], TranDesc, CompanyName, Amount, Available, Balance, Description, CompanyID, SettlementDate, TransmissionDate, ECC, Disc, [Entry], OrigAcct, EntryID, EntryName, LTRIM(RTRIM(TraceNumber)) AS TraceKey, TraceNumber, Description AS ExceptionText, SourceFile " & _
    "   FROM dbo.ACH_Exceptions_NSF WHERE TraceNumber IS NOT NULL AND RptDate BETWEEN ? AND ? " & _
    "   UNION ALL " & _
    "   SELECT 'TranCode', Id, RptDate, RptTime, RptSeq, Acct, SLX, '' AS [ID], [Name], TranDesc, CompanyName, Amount, Available, Balance, Description, CompanyID, SettlementDate, TransmissionDate, ECC, Disc, [Entry], OrigAcct, EntryID, EntryName, LTRIM(RTRIM(TraceNumber)), TraceNumber, ExceptionText, SourceFile " & _
    "   FROM dbo.ACH_Exceptions_TranCode WHERE TraceNumber IS NOT NULL AND RptDate BETWEEN ? AND ? " & _
    "   UNION ALL " & _
    "   SELECT 'Exceptions', Id, RptDate, RptTime, RptSeq, Acct, SLX, '' AS [ID], [Name], TranDesc, CompanyName, Amount, Available, Balance, Description, CompanyID, SettlementDate, TransmissionDate, ECC, Disc, [Entry], OrigAcct, EntryID, EntryName, LTRIM(RTRIM(TraceNumber)), TraceNumber, ExceptionText, SourceFile " & _
    "   FROM dbo.ACH_Exceptions WHERE TraceNumber IS NOT NULL AND RptDate BETWEEN ? AND ? " & _
    ") " & _
    "SELECT e.*, l.ACHCompanyName, l.ACHCompanyID, l.ACHECC, l.ACHTranCode, l.ACHAcct, l.ACHName, l.ACHAmount, l.ACHDescription " & _
    "FROM ex_norm e " & _
    "INNER JOIN list_norm l ON l.TraceKey = e.TraceKey " & _
    "ORDER BY e.RptDate, e.ExceptionType, e.TraceKey"

    Set cmdMatched = CreateCommandWithDates(sqlMatched, startDateObj, endDateObj, 4)
    Set rsMatched = cmdMatched.Execute

    fout.WriteLine "REPORT MATCHED EXCEPTIONS TO ACH_LIST"
    WriteFullHeader True

    If rsMatched.EOF Then
        fout.WriteLine "No matched exception records found for this period."
    Else
        Do While Not rsMatched.EOF
            WriteExceptionRow rsMatched, True
            rsMatched.MoveNext
        Loop
    End If

    fout.WriteLine ""

    ' =====================================================
    ' 3) EXTERNAL EXCEPTIONS
    ' =====================================================
    sqlUnmatchedEx = _
    "WITH list_norm AS ( " & _
    "   SELECT LTRIM(RTRIM(TraceNumber)) AS TraceKey " & _
    "   FROM dbo.ACH_LIST " & _
    "   WHERE TraceNumber IS NOT NULL " & _
    "     AND RptDate BETWEEN ? AND ? " & _
    "   GROUP BY LTRIM(RTRIM(TraceNumber)) " & _
    "), ex_norm AS ( " & _
    "   SELECT 'NSF' AS ExceptionType, RowID AS ExceptionId, RptDate, RptTime, RptSeq, Acct, SLX, [ID], [Name], TranDesc, CompanyName, Amount, Available, Balance, Description, CompanyID, SettlementDate, TransmissionDate, ECC, Disc, [Entry], OrigAcct, EntryID, EntryName, LTRIM(RTRIM(TraceNumber)) AS TraceKey, TraceNumber, Description AS ExceptionText, SourceFile " & _
    "   FROM dbo.ACH_Exceptions_NSF WHERE TraceNumber IS NOT NULL AND RptDate BETWEEN ? AND ? " & _
    "   UNION ALL " & _
    "   SELECT 'TranCode', Id, RptDate, RptTime, RptSeq, Acct, SLX, '' AS [ID], [Name], TranDesc, CompanyName, Amount, Available, Balance, Description, CompanyID, SettlementDate, TransmissionDate, ECC, Disc, [Entry], OrigAcct, EntryID, EntryName, LTRIM(RTRIM(TraceNumber)), TraceNumber, ExceptionText, SourceFile " & _
    "   FROM dbo.ACH_Exceptions_TranCode WHERE TraceNumber IS NOT NULL AND RptDate BETWEEN ? AND ? " & _
    "   UNION ALL " & _
    "   SELECT 'Exceptions', Id, RptDate, RptTime, RptSeq, Acct, SLX, '' AS [ID], [Name], TranDesc, CompanyName, Amount, Available, Balance, Description, CompanyID, SettlementDate, TransmissionDate, ECC, Disc, [Entry], OrigAcct, EntryID, EntryName, LTRIM(RTRIM(TraceNumber)), TraceNumber, ExceptionText, SourceFile " & _
    "   FROM dbo.ACH_Exceptions WHERE TraceNumber IS NOT NULL AND RptDate BETWEEN ? AND ? " & _
    ") " & _
    "SELECT e.* " & _
    "FROM ex_norm e " & _
    "LEFT JOIN list_norm l ON l.TraceKey = e.TraceKey " & _
    "WHERE l.TraceKey IS NULL " & _
    "ORDER BY e.RptDate, e.ExceptionType, e.TraceKey"

    Set cmdUnmatchedEx = CreateCommandWithDates(sqlUnmatchedEx, startDateObj, endDateObj, 4)
    Set rsUnmatchedEx = cmdUnmatchedEx.Execute

    fout.WriteLine "REPORT EXTERNAL EXCEPTIONS"
    WriteFullHeader False

    If rsUnmatchedEx.EOF Then
        fout.WriteLine "No external exception records found for this period."
    Else
        Do While Not rsUnmatchedEx.EOF
            WriteExceptionRow rsUnmatchedEx, False
            rsUnmatchedEx.MoveNext
        Loop
    End If

    fout.WriteLine ""

    ' =====================================================
    ' 4) ACH_LIST ITEMS WITH NO EXCEPTION
    ' =====================================================
    sqlUnmatchedList = _
    "WITH list_norm AS ( " & _
    "   SELECT " & _
    "       LTRIM(RTRIM(TraceNumber)) AS TraceKey, " & _
    "       MIN(RptDate) AS RptDate, " & _
    "       MAX(CompanyName) AS ACHCompanyName, " & _
    "       MAX(CompanyID) AS ACHCompanyID, " & _
    "       MAX(ECC) AS ACHECC, " & _
    "       MAX(TranCode) AS ACHTranCode, " & _
    "       MAX(IndividualAcct) AS ACHAcct, " & _
    "       MAX(IndividualName) AS ACHName, " & _
    "       MAX(Description) AS ACHDescription, " & _
    "       MAX(TRY_CONVERT(decimal(18,2), REPLACE(REPLACE(LTRIM(RTRIM(Amount)), ',', ''), '$', ''))) AS ACHAmount " & _
    "   FROM dbo.ACH_LIST " & _
    "   WHERE TraceNumber IS NOT NULL " & _
    "     AND RptDate BETWEEN ? AND ? " & _
    "   GROUP BY LTRIM(RTRIM(TraceNumber)) " & _
    "), ex_trace AS ( " & _
    "   SELECT LTRIM(RTRIM(TraceNumber)) AS TraceKey FROM dbo.ACH_Exceptions_NSF WHERE TraceNumber IS NOT NULL AND RptDate BETWEEN ? AND ? " & _
    "   UNION " & _
    "   SELECT LTRIM(RTRIM(TraceNumber)) FROM dbo.ACH_Exceptions_TranCode WHERE TraceNumber IS NOT NULL AND RptDate BETWEEN ? AND ? " & _
    "   UNION " & _
    "   SELECT LTRIM(RTRIM(TraceNumber)) FROM dbo.ACH_Exceptions WHERE TraceNumber IS NOT NULL AND RptDate BETWEEN ? AND ? " & _
    ") " & _
    "SELECT l.* " & _
    "FROM list_norm l " & _
    "LEFT JOIN ex_trace e ON e.TraceKey = l.TraceKey " & _
    "WHERE e.TraceKey IS NULL " & _
    "ORDER BY l.RptDate, l.TraceKey"

    Set cmdUnmatchedList = CreateCommandWithDates(sqlUnmatchedList, startDateObj, endDateObj, 4)
    Set rsUnmatchedList = cmdUnmatchedList.Execute

    fout.WriteLine "REPORT ACH_LIST RECORDS WITH NO EXCEPTION"
    WriteFullHeader True

    If rsUnmatchedList.EOF Then
        fout.WriteLine "No ACH_LIST-only records found for this period."
    Else
        Do While Not rsUnmatchedList.EOF
            WriteACHListOnlyRow rsUnmatchedList
            rsUnmatchedList.MoveNext
        Loop
    End If

 
	 ' =====================================================
	' 5) CONTROL TOTALS RAW
	' =====================================================
	Dim cmdControlRaw, rsControlRaw
	Dim sqlControlRaw

	sqlControlRaw = _
	"SELECT Report, RptDate, RptTime, RptSeq, Category, " & _
	"       PostedCount, PostedAmount, UnpostedCount, UnpostedAmount, " & _
	"       UncollectedCount, UncollectedAmount, SourceFile " & _
	"FROM dbo.ACH_ControlTotals " & _
	"WHERE RptDate BETWEEN ? AND ? " & _
	"ORDER BY RptDate, Report, LineOrder"

	Set cmdControlRaw = CreateCommandWithDates(sqlControlRaw, startDateObj, endDateObj, 1)
	Set rsControlRaw = cmdControlRaw.Execute

	fout.WriteLine ""
	fout.WriteLine "REPORT CONTROL TOTALS RAW"
	fout.WriteLine _
		Pad("Report", 22) & _
		Pad("RptDate", 15) & _
		Pad("RptTime", 15) & _
		Pad("RptSeq", 15) & _
		Pad("Category", 15) & _
		Pad("PostedCnt", 15) & _
		Pad("PostedAmt", 15) & _
		Pad("UnpostCnt", 24) & _
		Pad("UnpostAmt", 18) & _
		Pad("UncollCnt", 24) & _
		Pad("UncollAmt", 15) & _
		Pad("SourceFile", 15)

	If rsControlRaw.EOF Then
		fout.WriteLine "No control total rows found for this period."
	Else
		Do While Not rsControlRaw.EOF
			fout.WriteLine _
				Pad(rsControlRaw("Report"), 22) & _
				Pad(FmtDate8(rsControlRaw("RptDate")), 15) & _
				Pad(rsControlRaw("RptTime"), 15) & _
				Pad(rsControlRaw("RptSeq"), 15) & _
				Pad(rsControlRaw("Category"), 15) & _
				Pad(rsControlRaw("PostedCount"), 15) & _
				Pad(FormatAmount(rsControlRaw("PostedAmount")), 15) & _
				Pad(rsControlRaw("UnpostedCount"), 24) & _
				Pad(FormatAmount(rsControlRaw("UnpostedAmount")), 18) & _
				Pad(rsControlRaw("UncollectedCount"), 24) & _
				Pad(FormatAmount(rsControlRaw("UncollectedAmount")), 15) & _
				Pad(rsControlRaw("SourceFile"), 27)

			rsControlRaw.MoveNext
		Loop
	End If








    CleanupObjects conn, rsSum, cmdSum
    CleanupObjects Nothing, rsMatched, cmdMatched
    CleanupObjects Nothing, rsUnmatchedEx, cmdUnmatchedEx
    CleanupObjects Nothing, rsUnmatchedList, cmdUnmatchedList

    On Error GoTo 0
End Sub

Sub ProcessLine()
End Sub

Sub CloseDoc()
End Sub

Sub WriteFullHeader(ByVal includeACH)
    Dim line

    line = _
        Pad("Type", 22) & _
        Pad("RptDate", 15) & _
        Pad("RptTime", 15) & _
        Pad("RptSeq", 15) & _
        Pad("Acct", 15) & _
        Pad("SLX", 15) & _
        Pad("ID", 15) & _
        Pad("Name", 24) & _
        Pad("TranDesc", 18) & _
        Pad("CompanyName", 24) & _
        Pad("Amount", 15) & _
        Pad("Available", 15) & _
        Pad("Balance", 15) & _
        Pad("Description", 28) & _
        Pad("CompanyID", 16) & _
        Pad("SettlementDate", 16) & _
        Pad("TransmissionDate", 18) & _
        Pad("ECC",15) & _
        Pad("Disc", 22) & _
        Pad("Entry", 16) & _
        Pad("OrigAcct", 16) & _
        Pad("EntryID", 20) & _
        Pad("EntryName", 24) & _
        Pad("TraceNumber", 20) & _
        Pad("ExceptionText", 32) & _
        Pad("SourceFile", 40)

    If includeACH Then
        line = line & _
            Pad("ACHCompany", 24) & _
            Pad("ACHAmount", 15) & _
            Pad("ACHCompanyID", 16) & _
            Pad("ACHECC", 15) & _
            Pad("ACHTranCode", 15) & _
            Pad("ACHAcct", 18) & _
            Pad("ACHName", 24) & _
            Pad("ACHDescription", 28)
    End If

    fout.WriteLine line
End Sub

Sub WriteExceptionRow(ByRef rs, ByVal includeACH)
    Dim line

    line = _
        Pad(rs("ExceptionType"), 22) & _
        Pad(FmtDate8(rs("RptDate")), 15) & _
        Pad(rs("RptTime"), 15) & _
        Pad(rs("RptSeq"), 15) & _
        Pad(rs("Acct"), 15) & _
        Pad(rs("SLX"), 15) & _
        Pad(rs("ID"), 15) & _
        Pad(rs("Name"), 24) & _
        Pad(rs("TranDesc"), 18) & _
        Pad(rs("CompanyName"), 24) & _
        Pad(rs("Amount"), 15) & _
        Pad(rs("Available"), 15) & _
        Pad(rs("Balance"), 15) & _
        Pad(rs("Description"), 28) & _
        Pad(rs("CompanyID"), 16) & _
        Pad(rs("SettlementDate"), 16) & _
        Pad(rs("TransmissionDate"), 18) & _
        Pad(rs("ECC"), 15) & _
        Pad(rs("Disc"), 22) & _
        Pad(rs("Entry"), 16) & _
        Pad(rs("OrigAcct"), 16) & _
        Pad(rs("EntryID"), 20) & _
        Pad(rs("EntryName"), 24) & _
        Pad(rs("TraceNumber"), 20) & _
        Pad(rs("ExceptionText"), 32) & _
        Pad(rs("SourceFile"), 40)

    If includeACH Then
        line = line & _
            Pad(rs("ACHCompanyName"), 24) & _
            Pad(FormatAmount(rs("ACHAmount")), 14) & _
            Pad(rs("ACHCompanyID"), 16) & _
            Pad(rs("ACHECC"), 15) & _
            Pad(rs("ACHTranCode"), 15) & _
            Pad(rs("ACHAcct"), 18) & _
            Pad(rs("ACHName"), 24) & _
            Pad(rs("ACHDescription"), 28)
    End If

    fout.WriteLine line
End Sub

Sub WriteACHListOnlyRow(ByRef rs)
    fout.WriteLine _
        Pad("ACH_LIST", 22) & _
        Pad(FmtDate8(rs("RptDate")), 15) & _
        Pad("", 15) & _
        Pad("", 15) & _
        Pad(rs("ACHAcct"), 15) & _
        Pad("",15) & _
        Pad("", 15) & _
        Pad(rs("ACHName"), 24) & _
        Pad("", 18) & _
        Pad(rs("ACHCompanyName"), 24) & _
        Pad(FormatAmount(rs("ACHAmount")), 15) & _
        Pad("", 15) & _
        Pad("", 15) & _
        Pad(rs("ACHDescription"), 28) & _
        Pad(rs("ACHCompanyID"), 16) & _
        Pad("", 16) & _
        Pad("", 18) & _
        Pad(rs("ACHECC"), 15) & _
        Pad("", 22) & _
        Pad("", 16) & _
        Pad("", 16) & _
        Pad("", 15) & _
        Pad("", 24) & _
        Pad(rs("TraceKey"), 20) & _
        Pad("", 32) & _
        Pad("", 40) & _
        Pad(rs("ACHCompanyName"), 24) & _
        Pad(FormatAmount(rs("ACHAmount")), 15) & _
        Pad(rs("ACHCompanyID"), 16) & _
        Pad(rs("ACHECC"), 15) & _
        Pad(rs("ACHTranCode"), 15) & _
        Pad(rs("ACHAcct"), 18) & _
        Pad(rs("ACHName"), 24) & _
        Pad(rs("ACHDescription"), 28)
End Sub

Function CreateCommandWithDates(ByVal sqlText, ByVal startDateObj, ByVal endDateObj, ByVal pairCount)
    Dim cmd, i

    Set cmd = CreateObject("ADODB.Command")

    With cmd
        .ActiveConnection = conn
        .CommandText = sqlText
        .CommandType = adCmdText

        For i = 1 To pairCount
            .Parameters.Append .CreateParameter("", adDBDate, adParamInput, , startDateObj)
            .Parameters.Append .CreateParameter("", adDBDate, adParamInput, , endDateObj)
        Next
    End With

    Set CreateCommandWithDates = cmd
End Function

Sub LogLine(msg)
    On Error Resume Next
    fout.WriteLine msg
    On Error GoTo 0
End Sub

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
    Dim cfg
    Set cfg = ReadIniSafe(iniPath)

    If cfg Is Nothing Then
        LogLine "ERROR: Could not read INI file: " & iniPath
        Exit Sub
    End If

    gMode         = GetIniValue(cfg, "Report.Mode", "DAILY")
    gRelOffset    = GetIniValue(cfg, "Report.RelativeOffset", "0")
    gYear         = GetIniValue(cfg, "Report.Year", "")
    gMonth        = GetIniValue(cfg, "Report.Month", "")
    gQuarter      = GetIniValue(cfg, "Report.Quarter", "")
    gStartDateStr = GetIniValue(cfg, "Report.StartDate", "")
    gEndDateStr   = GetIniValue(cfg, "Report.EndDate", "")

    gConnStr       = GetIniValue(cfg, "Database.ConnStr", "")
    gFullTableName = GetIniValue(cfg, "Database.TableName", "[dbo].[ACH_LIST]")

    If IsNumeric(gRelOffset) Then
        gRelOffset = CInt(gRelOffset)
    Else
        gRelOffset = 0
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
    Dim y, m, q, tmp, dow, weekStart
    Dim baseDate

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

            If gRelOffset <> 0 Then
                baseDate = DateAdd("m", gRelOffset, baseDate)
            End If

            dStart = DateSerial(Year(baseDate), Month(baseDate), 1)

            If Year(baseDate) = Year(today) And Month(baseDate) = Month(today) Then
                dEnd = today
            Else
                dEnd = DateAdd("d", -1, DateAdd("m", 1, dStart))
            End If

        Case Else
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

    If IsNull(val) Then
        s = ""
    Else
        s = CStr(val)
    End If

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

Function FormatAmount(val)
    On Error Resume Next

    If IsNull(val) Or Len(Trim(CStr(val))) = 0 Then
        FormatAmount = ""
        Exit Function
    End If

    If IsNumeric(val) Then
        FormatAmount = FormatNumber(CDbl(val), 2)
    Else
        FormatAmount = Trim(CStr(val))
    End If

    If Err.Number <> 0 Then
        FormatAmount = Trim(CStr(val))
        Err.Clear
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