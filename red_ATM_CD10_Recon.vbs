Option Explicit

' ---------------- GLOBALS ----------------
Dim gMode, gRelOffset, gYear, gMonth, gQuarter
Dim gStartDateStr, gEndDateStr

Dim gConnStr
Dim gCD10FullTableName, gATMFullTableName
Dim gCD10Table2Part, gCD10ObjIdName
Dim gATMTable2Part, gATMObjIdName

Dim gWeekdayLookbackDays, gWeekdayLookaheadDays
Dim gWeekendLookbackDays, gWeekendLookaheadDays
Dim gExcludePreAuth, gRequireApprovedResponse
 
Dim gUseFullCardNumber

' Rule enable flags
Dim gEnableRule1, gEnableRule2, gEnableRule3

' Custom base scores
Dim gRule1BaseScore, gRule2BaseScore, gRule3BaseScore

' Custom penalties
Dim gDatePenaltyPerDay, gTimePenaltyWithin1Hour, gTimePenaltyWithin24Hours, gTimePenaltyOver24Hours

Dim gSortOrder

Dim conn, rs
Dim INI_PATH: INI_PATH = "S:\CarolinaTrust\byREQUEST\Templates\ATM_CD10_Recon.ini"

' ---------------------------------------------------------
'  Simple log helper
' ---------------------------------------------------------
Sub LogLine(msg)
    On Error Resume Next
    fout.WriteLine msg
    On Error GoTo 0
End Sub

' ---------------------------------------------------------
'  StartDoc

Sub StartDoc
    On Error Resume Next

    Dim startKey, endKey
    Dim sql, cmd
    Dim startDateObj, endDateObj
    Dim atmStartDateObj, atmEndDateObj
    Dim rowCount
    Dim matchedCount, unmatchedCD10Count, unmatchedATMCount
    
    ' Arrays for date analysis
    Dim matchedByDate(), unmatchedCD10ByDate(), unmatchedATMByDate()
    Dim dateDict, currentDate, dateKey
    Dim i

    If Not FileExists(INI_PATH) Then
        LogLine "ERROR: INI file not found: " & INI_PATH
        Exit Sub
    End If

    LoadParameters INI_PATH

    If Trim(gConnStr) = "" Then
        LogLine "ERROR: ConnStr not defined in [Database] section."
        Exit Sub
    End If

    BuildTableNames gCD10FullTableName, gCD10Table2Part, gCD10ObjIdName
    BuildTableNames gATMFullTableName, gATMTable2Part, gATMObjIdName
    
    ' LogLine "Resolved CD10 table: " & gCD10Table2Part
    ' LogLine "Resolved ATM table: " & gATMTable2Part

    If Len(gCD10Table2Part) = 0 Then
        LogLine "ERROR: Could not parse CD10Table: " & gCD10FullTableName
        Exit Sub
    End If

    If Len(gATMTable2Part) = 0 Then
        LogLine "ERROR: Could not parse ATMTable: " & gATMFullTableName
        Exit Sub
    End If

    ComputeDateRange startKey, endKey

    If startKey = "" Or endKey = "" Then
        LogLine "ERROR: Could not compute date range from parameters."
        Exit Sub
    End If

    startDateObj = FormatKeyToDate(startKey)
    endDateObj   = FormatKeyToDate(endKey)
    
    ' LogLine "Converted startKey: " & startKey & " to date: " & CStr(startDateObj)
    ' LogLine "Converted endKey: " & endKey & " to date: " & CStr(endDateObj)
    
    If Not IsDate(startDateObj) Then
        LogLine "ERROR: Invalid start date from key: " & startKey
        Exit Sub
    End If
    If Not IsDate(endDateObj) Then
        LogLine "ERROR: Invalid end date from key: " & endKey
        Exit Sub
    End If

    atmStartDateObj = DateAdd("d", -gWeekendLookbackDays, startDateObj)
    atmEndDateObj   = DateAdd("d", gWeekendLookaheadDays, endDateObj)

    Set conn = CreateObject("ADODB.Connection")
    Err.Clear
    conn.Open gConnStr
    If Err.Number <> 0 Then
        LogLine "ERROR opening DB connection: " & Err.Description & " (Err=" & Err.Number & ")"
        Exit Sub
    End If
    On Error GoTo 0

    sql = BuildReconSQL()

    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = sql
        .CommandType = 1

        .Parameters.Append .CreateParameter("@CD10Start", 135, 1, , startDateObj)
        .Parameters.Append .CreateParameter("@CD10End", 135, 1, , endDateObj)
        .Parameters.Append .CreateParameter("@ExcludePreAuth", 3, 1, , gExcludePreAuth)
        .Parameters.Append .CreateParameter("@RequireApproved", 3, 1, , gRequireApprovedResponse)
        .Parameters.Append .CreateParameter("@ATMStart", 135, 1, , atmStartDateObj)
        .Parameters.Append .CreateParameter("@ATMEnd", 135, 1, , atmEndDateObj)
        .Parameters.Append .CreateParameter("@WeekendLookback", 3, 1, , gWeekendLookbackDays)
        .Parameters.Append .CreateParameter("@WeekdayLookback", 3, 1, , gWeekdayLookbackDays)
        .Parameters.Append .CreateParameter("@WeekendLookahead", 3, 1, , gWeekendLookaheadDays)
        .Parameters.Append .CreateParameter("@WeekdayLookahead", 3, 1, , gWeekdayLookaheadDays)
    End With

    ' LogLine "Executing reconciliation query..."
    ' LogLine "CD10 StartDate = " & CStr(startDateObj)
    ' LogLine "CD10 EndDate   = " & CStr(endDateObj)
    ' LogLine "ATM StartDate  = " & CStr(atmStartDateObj)
    ' LogLine "ATM EndDate    = " & CStr(atmEndDateObj)
    ' LogLine "ExcludePreAuth = " & CStr(gExcludePreAuth)
    ' LogLine "RequireApprovedResponse = " & CStr(gRequireApprovedResponse)
    ' LogLine "WeekendLookbackDays = " & CStr(gWeekendLookbackDays)
    ' LogLine "WeekdayLookbackDays = " & CStr(gWeekdayLookbackDays)
    ' LogLine "WeekendLookaheadDays = " & CStr(gWeekendLookaheadDays)
    ' LogLine "WeekdayLookaheadDays = " & CStr(gWeekdayLookaheadDays)
    
    ' LogLine "Total parameters added: " & cmd.Parameters.Count

    On Error Resume Next
    Set rs = cmd.Execute
    If Err.Number <> 0 Then
        LogLine "ERROR executing reconciliation query: " & Err.Description
        LogLine "Error Number: " & Err.Number
        LogLine "Error Source: " & Err.Source
        CleanupObjects conn, rs, cmd
        Exit Sub
    End If
    On Error GoTo 0

    Call PrintRulesLegend(False)

    ' Initialize counters
    rowCount = 0
    matchedCount = 0
    unmatchedCD10Count = 0
    unmatchedATMCount = 0
    
    ' Create dictionary for date aggregation
    Set dateDict = CreateObject("Scripting.Dictionary")

    fout.WriteLine "ATM vs CD10 RECONCILIATION"
    fout.WriteLine "Date Range: " & CStr(startDateObj) & " to " & CStr(endDateObj)
    fout.WriteLine "Report Generated: " & Now()

	fout.WriteLine ""
	
    fout.WriteLine _
        Pad("STATUS", 25) & _
        Pad("RULE", 10) & _
        Pad("SCORE", 6) & _
        Pad("CD10 DATE/TIME", 25) & _
        Pad("ATM DATE/TIME", 25) & _
        Pad("LAST4", 10) & _
        Pad("AMOUNT", 12) & _
        Pad("CD10 SEQ", 12) & _
        Pad("ATM TRACE", 12) & _
        Pad("ATM REF", 12) & _
        Pad("CD10 DESC", 16) & _
        Pad("ATM CODE", 10)

    ' Process each row and count by status
    Do While Not rs.EOF
        Dim status, cd10Date, atmDate
        
        status = SafeString(rs("ReconStatus"))
        cd10Date = rs("CD10_TranDateTime")
        atmDate = rs("ATM_TranDateTime")
        
        ' Determine the date key for aggregation
        If status = "MATCHED" Then
            If Not IsNull(cd10Date) Then
                dateKey = FormatDateKey(CDate(cd10Date))
            ElseIf Not IsNull(atmDate) Then
                dateKey = FormatDateKey(CDate(atmDate))
            Else
                dateKey = "UNKNOWN"
            End If
        ElseIf status = "UNMATCHED_CD10" Then
            If Not IsNull(cd10Date) Then
                dateKey = FormatDateKey(CDate(cd10Date))
            Else
                dateKey = "UNKNOWN"
            End If
        ElseIf status = "UNMATCHED_ATM" Then
            If Not IsNull(atmDate) Then
                dateKey = FormatDateKey(CDate(atmDate))
            Else
                dateKey = "UNKNOWN"
            End If
        Else
            dateKey = "UNKNOWN"
        End If
        
        ' Initialize dictionary entry if not exists
        If Not dateDict.Exists(dateKey) Then
            dateDict.Add dateKey, Array(0, 0, 0)  ' (Matched, UnmatchedCD10, UnmatchedATM)
        End If
        
        ' Update counts in dictionary
        Dim counts
        counts = dateDict(dateKey)
        
        Select Case status
            Case "MATCHED"
                matchedCount = matchedCount + 1
                counts(0) = counts(0) + 1
            Case "UNMATCHED_CD10"
                unmatchedCD10Count = unmatchedCD10Count + 1
                counts(1) = counts(1) + 1
            Case "UNMATCHED_ATM"
                unmatchedATMCount = unmatchedATMCount + 1
                counts(2) = counts(2) + 1
        End Select
        
        dateDict(dateKey) = counts
        
        ' Output the row
        fout.WriteLine _
            Pad(status, 25) & _
            Pad(SafeString(rs("MatchRule")), 10) & _
            Pad(SafeString(rs("MatchScore")), 6) & _
            Pad(SafeString(rs("CD10_TranDateTime")), 25) & _
            Pad(SafeString(rs("ATM_TranDateTime")), 25) & _
            Pad(ChooseCard(rs("CD10_CardLast4"), rs("ATM_CardLast4")), 10) & _
            Pad(ChooseAmount(rs("CD10_Amount"), rs("ATM_Amount")), 12) & _
            Pad(SafeString(rs("CD10_SeqNo")), 12) & _
            Pad(SafeString(rs("ATM_TraceNo")), 12) & _
            Pad(SafeString(rs("ATM_ReferenceNo")), 12) & _
            Pad(SafeString(rs("CD10_TranDes")), 16) & _
            Pad(SafeString(rs("ATM_TranCode")), 10)
        
        rowCount = rowCount + 1
        rs.MoveNext
    Loop
    
    ' ============================================================
    ' ENHANCED SUMMARY WITH DATE BREAKDOWN
    ' ============================================================
    fout.WriteLine ""
    fout.WriteLine "=" & String(78, "=")
    fout.WriteLine "SUMMARY"
    fout.WriteLine "=" & String(78, "=")
    fout.WriteLine ""
    
    ' Overall totals
    fout.WriteLine "OVERALL TOTALS:"
    fout.WriteLine Pad("Total Records:", 25) &   Pad(rowCount, 10)
	fout.WriteLine Pad("Matched Records:", 25) &   Pad(matchedCount, 10)
    fout.WriteLine Pad("Unmatched CD10 Records:", 25) &   Pad(unmatchedCD10Count, 10)
    fout.WriteLine Pad("Unmatched ATM Records:", 25) &   Pad(unmatchedATMCount, 10)

    
    If matchedCount > 0 Then
        Dim matchRate
        matchRate = Round((matchedCount / rowCount) * 100, 2)
		fout.WriteLine Pad("Match Rate:", 25) &   Pad(matchRate & "%", 10)
    End If
    fout.WriteLine ""
    
    ' Date breakdown header
    fout.WriteLine "DATE BREAKDOWN:"
    fout.WriteLine "-------------------------------------------------------------------------------"
    fout.WriteLine Pad("Date", 25) & _
                   Pad("Matched", 16) & _
                   Pad("Unmatched CD10", 25) & _
                   Pad("Unmatched ATM", 25) & _
                   Pad("DayTotal", 10) & _
                   Pad("Daily%", 12)
    fout.WriteLine "-------------------------------------------------------------------------------"
    
    ' Sort dates and output
    Dim sortedKeys, j, key
    sortedKeys = dateDict.Keys
    
    ' Simple bubble sort for dates
    For i = 0 To UBound(sortedKeys) - 1
        For j = i + 1 To UBound(sortedKeys)
            If sortedKeys(i) > sortedKeys(j) Then
                key = sortedKeys(i)
                sortedKeys(i) = sortedKeys(j)
                sortedKeys(j) = key
            End If
        Next
    Next
    
    ' Output each date's breakdown
    For i = 0 To UBound(sortedKeys)
        key = sortedKeys(i)
        counts = dateDict(key)
        
        Dim dailyMatched, dailyUnmatchedCD10, dailyUnmatchedATM, dailyTotal, dailyMatchRate
        dailyMatched = counts(0)
        dailyUnmatchedCD10 = counts(1)
        dailyUnmatchedATM = counts(2)
        dailyTotal = dailyMatched + dailyUnmatchedCD10 + dailyUnmatchedATM
        If dailyTotal > 0 Then
            dailyMatchRate = Round((dailyMatched / dailyTotal) * 100, 2)
        Else
            dailyMatchRate = 0
        End If
        
        ' Format date for display
        Dim displayDate
        If key = "UNKNOWN" Then
            displayDate = "UNKNOWN"
        Else
            displayDate = key
        End If
        
        fout.WriteLine Pad(displayDate, 25) & _
                       Pad(dailyMatched, 16) & _
                       Pad(dailyUnmatchedCD10, 25) & _
                       Pad(dailyUnmatchedATM, 25) & _
                       Pad(dailyTotal, 10) & _
                       Pad(dailyMatchRate & "%", 12)
    Next
    
    ' fout.WriteLine "-------------------------------------------------------------------------------"
    ' fout.WriteLine ""
    
    ' ' Add settlement delay analysis
    ' fout.WriteLine "SETTLEMENT DELAY ANALYSIS:"
    ' fout.WriteLine "-------------------------------------------------------------------------------"
    
    ' ' Calculate pending settlements (unmatched CD10 from recent dates)
    ' Dim pendingCount, recentDate
    ' pendingCount = 0
    ' recentDate = DateAdd("d", -3, Date())
    
    ' For i = 0 To UBound(sortedKeys)
        ' key = sortedKeys(i)
        ' counts = dateDict(key)
        
        ' If key <> "UNKNOWN" Then
            ' Dim keyDate
            ' keyDate = FormatKeyToDate(key)
            ' If keyDate >= recentDate And keyDate <= endDateObj Then
                ' pendingCount = pendingCount + counts(1)  ' Unmatched CD10 from last 3 days
            ' End If
        ' End If
    ' Next
    
    ' fout.WriteLine "  Pending Settlements (last 3 days): " & pendingCount
    ' fout.WriteLine "  Note: Recent transactions may take 1-5 days to appear in ATM_RECON"
    ' fout.WriteLine "  Expected match rate increases to 75-85% after 7-14 days"
    ' fout.WriteLine ""

    ' LogLine "Reconciliation completed successfully. Total rows: " & rowCount
    ' LogLine "Summary - Matched: " & matchedCount & ", Unmatched CD10: " & unmatchedCD10Count & ", Unmatched ATM: " & unmatchedATMCount

    CleanupObjects conn, rs, cmd
    On Error GoTo 0
End Sub

' Helper function to format date key to display date
Function FormatKeyToDisplayDate(ByVal yyyymmdd)
    Dim y, m, d
    
    ' Handle invalid input
    If IsNull(yyyymmdd) Or Len(CStr(yyyymmdd)) <> 8 Then
        FormatKeyToDisplayDate = CStr(yyyymmdd)
        Exit Function
    End If
    
    y = CInt(Mid(yyyymmdd, 1, 4))
    m = CInt(Mid(yyyymmdd, 5, 2))
    d = CInt(Mid(yyyymmdd, 7, 2))
    
    ' Validate year is reasonable (2000-2099)
    If y < 2000 Or y > 2099 Then
        FormatKeyToDisplayDate = CStr(yyyymmdd)
        Exit Function
    End If
    
    FormatKeyToDisplayDate = MonthName(m) & " " & d & ", " & y
End Function

Sub CloseDoc
End Sub

 
' =========================================================
'  SQL builder 
' =========================================================
Function BuildReconSQL()
    Dim sql
    Dim ruleLogic
    Dim hasAnyRule
    Dim matchCondition

    ' Check if any rules are enabled
    hasAnyRule = (gEnableRule1 = 1 Or gEnableRule2 = 1 Or gEnableRule3 = 1)
    
    If Not hasAnyRule Then
        LogLine "ERROR: No rules enabled! At least one rule must be enabled."
        BuildReconSQL = ""
        Exit Function
    End If

    ' Build match condition based on configuration
    If gUseFullCardNumber = 1 Then
        matchCondition = "c.CardNumber_Norm = a.CardNumber_Norm"
    Else
        matchCondition = "c.CardLast4 = a.CardLast4"
    End If

    ' Build dynamic CASE statement for MatchRule and MatchScore
    ruleLogic = "   CASE" & vbCrLf
    
    ' Rule 1: Reference Match
    If gEnableRule1 = 1 Then
        ruleLogic = ruleLogic & "     WHEN c.SeqNo_Norm <> ''" & vbCrLf
        ruleLogic = ruleLogic & "      AND (" & vbCrLf
        ruleLogic = ruleLogic & "           c.SeqNo_NoLeadingZeros = a.TerminalID_NoLeadingZeros" & vbCrLf
        ruleLogic = ruleLogic & "        OR c.SeqNo_Norm = a.TraceNo_Norm" & vbCrLf
        ruleLogic = ruleLogic & "        OR c.SeqNo_Norm = a.ReferenceNo_Norm" & vbCrLf
        ruleLogic = ruleLogic & "      )" & vbCrLf
        ruleLogic = ruleLogic & "      AND " & matchCondition & vbCrLf
        ruleLogic = ruleLogic & "      AND c.AmountValue_Norm = a.AmountValue_Norm" & vbCrLf
        ruleLogic = ruleLogic & "     THEN 'RULE_1'" & vbCrLf
    End If
    
    ' Rule 2: Same Day Match
    If gEnableRule2 = 1 Then
        ruleLogic = ruleLogic & "     WHEN " & matchCondition & vbCrLf
        ruleLogic = ruleLogic & "      AND c.AmountValue_Norm = a.AmountValue_Norm" & vbCrLf
        ruleLogic = ruleLogic & "      AND c.TranDate = a.TranDate" & vbCrLf
        If gEnableRule1 = 1 Then
            ruleLogic = ruleLogic & "      AND NOT (c.SeqNo_Norm <> '' AND (" & vbCrLf
            ruleLogic = ruleLogic & "           c.SeqNo_NoLeadingZeros = a.TerminalID_NoLeadingZeros" & vbCrLf
            ruleLogic = ruleLogic & "        OR c.SeqNo_Norm = a.TraceNo_Norm" & vbCrLf
            ruleLogic = ruleLogic & "        OR c.SeqNo_Norm = a.ReferenceNo_Norm" & vbCrLf
            ruleLogic = ruleLogic & "      ))" & vbCrLf
        End If
        ruleLogic = ruleLogic & "     THEN 'RULE_2'" & vbCrLf
    End If
    
    ' Rule 3: Date Window Match
    If gEnableRule3 = 1 Then
        ruleLogic = ruleLogic & "     ELSE 'RULE_3'" & vbCrLf
    End If
    
    ruleLogic = ruleLogic & "   END AS MatchRule," & vbCrLf
    
    ' Build dynamic CASE statement for MatchScore
    ruleLogic = ruleLogic & "   (CASE" & vbCrLf
    
    ' Rule 1 Score
    If gEnableRule1 = 1 Then
        ruleLogic = ruleLogic & "      WHEN c.SeqNo_Norm <> ''" & vbCrLf
        ruleLogic = ruleLogic & "       AND (" & vbCrLf
        ruleLogic = ruleLogic & "            c.SeqNo_NoLeadingZeros = a.TerminalID_NoLeadingZeros" & vbCrLf
        ruleLogic = ruleLogic & "         OR c.SeqNo_Norm = a.TraceNo_Norm" & vbCrLf
        ruleLogic = ruleLogic & "         OR c.SeqNo_Norm = a.ReferenceNo_Norm" & vbCrLf
        ruleLogic = ruleLogic & "       )" & vbCrLf
        ruleLogic = ruleLogic & "       AND " & matchCondition & vbCrLf
        ruleLogic = ruleLogic & "       AND c.AmountValue_Norm = a.AmountValue_Norm THEN " & gRule1BaseScore & vbCrLf
    End If
    
    ' Rule 2 Score
    If gEnableRule2 = 1 Then
        ruleLogic = ruleLogic & "      WHEN " & matchCondition & vbCrLf
        ruleLogic = ruleLogic & "       AND c.AmountValue_Norm = a.AmountValue_Norm" & vbCrLf
        ruleLogic = ruleLogic & "       AND c.TranDate = a.TranDate THEN " & gRule2BaseScore & vbCrLf
    End If
    
    ' Rule 3 Score
    If gEnableRule3 = 1 Then
        ruleLogic = ruleLogic & "      ELSE " & gRule3BaseScore & vbCrLf
    End If
    
    ruleLogic = ruleLogic & "    END" & vbCrLf
    ruleLogic = ruleLogic & "    - ABS(DATEDIFF(DAY, c.TranDate, a.TranDate)) * " & gDatePenaltyPerDay & vbCrLf
    ruleLogic = ruleLogic & "    - CASE" & vbCrLf
    ruleLogic = ruleLogic & "        WHEN ABS(DATEDIFF(MINUTE, c.TranDateTime, a.TranDateTime)) <= 60 THEN " & gTimePenaltyWithin1Hour & vbCrLf
    ruleLogic = ruleLogic & "        WHEN ABS(DATEDIFF(MINUTE, c.TranDateTime, a.TranDateTime)) <= 1440 THEN " & gTimePenaltyWithin24Hours & vbCrLf
    ruleLogic = ruleLogic & "        ELSE " & gTimePenaltyOver24Hours & vbCrLf
    ruleLogic = ruleLogic & "      END) AS MatchScore" & vbCrLf

    ' Build the complete SQL
    sql = ""
    sql = sql & "WITH CD10 AS ("
    sql = sql & " SELECT "
    sql = sql & "   c.Id, "
    sql = sql & "   c.LDate, "
    sql = sql & "   LTRIM(RTRIM(ISNULL(c.CardNumber,''))) AS CardNumber_Norm, "
    sql = sql & "   RIGHT(LTRIM(RTRIM(ISNULL(c.CardNumber,''))), 4) AS CardLast4, "
    sql = sql & "   LTRIM(RTRIM(ISNULL(c.SeqNo,''))) AS SeqNo_Norm, "
    sql = sql & "   LTRIM(RTRIM(ISNULL(TRY_CAST(TRY_CAST(c.SeqNo AS BIGINT) AS NVARCHAR(50)),''))) AS SeqNo_NoLeadingZeros, "
    sql = sql & "   UPPER(LTRIM(RTRIM(ISNULL(c.TranDesc,'')))) AS TranDes_Norm, "
    sql = sql & "   UPPER(LTRIM(RTRIM(ISNULL(c.RespCode,'')))) AS Rsp_Norm, "
    sql = sql & "   TRY_CONVERT(decimal(18,2), c.AmountValue) AS AmountValue_Norm, "
    sql = sql & "   c.SourceFile, "
    sql = sql & "   CAST(c.LDate AS datetime) + CAST(CAST(c.LTime AS time) AS datetime) AS TranDateTime, "
    sql = sql & "   CAST(c.LDate AS date) AS TranDate, "
    sql = sql & "   CASE "
    sql = sql & "     WHEN UPPER(LTRIM(RTRIM(ISNULL(c.TranDesc,'')))) LIKE '%PRE%' "
    sql = sql & "       OR UPPER(LTRIM(RTRIM(ISNULL(c.TranDesc,'')))) LIKE '%AUTH%' "
    sql = sql & "       OR UPPER(LTRIM(RTRIM(ISNULL(c.TranDesc,'')))) LIKE '%HOLD%' "
    sql = sql & "     THEN 1 ELSE 0 END AS IsLikelyPreAuth, "
    sql = sql & "   CASE "
    sql = sql & "     WHEN UPPER(LTRIM(RTRIM(ISNULL(c.RespCode,'')))) IN ('A','APP','APPROVED','00','C') "
    sql = sql & "       OR LTRIM(RTRIM(ISNULL(c.RespCode,''))) = '' "
    sql = sql & "     THEN 1 ELSE 0 END AS IsApproved "
    sql = sql & " FROM " & gCD10Table2Part & " c "
    sql = sql & " WHERE c.LDate BETWEEN ? AND ? "
    sql = sql & "), "
    sql = sql & "CD10F AS ( "
    sql = sql & " SELECT * "
    sql = sql & " FROM CD10 "
    sql = sql & " WHERE AmountValue_Norm IS NOT NULL "
    sql = sql & "   AND AmountValue_Norm <> 0 "
    sql = sql & "   AND CardNumber_Norm <> '' "
    sql = sql & "   AND (? = 0 OR IsLikelyPreAuth = 0) "
    sql = sql & "   AND (? = 0 OR IsApproved = 1) "
	sql = sql & "   AND UPPER(TranDes_Norm) IN ('PUR-CHK', 'WDL', 'MTF', 'PCR', 'DEP', 'TFR', 'MTC') "
    sql = sql & "), "
    sql = sql & "ATMR AS ( "
    sql = sql & " SELECT "
    sql = sql & "   a.Id, "
    sql = sql & "   LTRIM(RTRIM(ISNULL(a.CardNumber,''))) AS CardNumber_Norm, "
    sql = sql & "   RIGHT(LTRIM(RTRIM(ISNULL(a.CardNumber,''))), 4) AS CardLast4, "
    sql = sql & "   a.TranDateTime, "
    sql = sql & "   CAST(a.TranDateTime AS date) AS TranDate, "
    sql = sql & "   UPPER(LTRIM(RTRIM(ISNULL(a.TranCode,'')))) AS TranCode_Norm, "
    sql = sql & "   ABS(TRY_CONVERT(decimal(18,2), a.AmountValue)) AS AmountValue_Norm, "
    sql = sql & "   LTRIM(RTRIM(ISNULL(a.TraceNo,''))) AS TraceNo_Norm, "
    sql = sql & "   LTRIM(RTRIM(ISNULL(a.ReferenceNo,''))) AS ReferenceNo_Norm, "
    sql = sql & "   LTRIM(RTRIM(ISNULL(a.TerminalID,''))) AS TerminalID_Norm, "
    sql = sql & "   LTRIM(RTRIM(ISNULL(TRY_CAST(TRY_CAST(a.TerminalID AS BIGINT) AS NVARCHAR(50)),''))) AS TerminalID_NoLeadingZeros, "
    sql = sql & "   a.SourceFile "
    sql = sql & " FROM " & gATMTable2Part & " a "
    sql = sql & " WHERE CAST(a.TranDateTime AS date) BETWEEN ? AND ? "
    sql = sql & "   AND TRY_CONVERT(decimal(18,2), a.AmountValue) IS NOT NULL "
    sql = sql & "   AND TRY_CONVERT(decimal(18,2), a.AmountValue) <> 0 "
    sql = sql & "   AND LTRIM(RTRIM(ISNULL(a.CardNumber,''))) <> '' "
    sql = sql & "), "
    sql = sql & "CandidateMatches AS ( "
    sql = sql & " SELECT "
    sql = sql & "   c.Id AS CD10_Id, "
    sql = sql & "   a.Id AS ATMRecon_Id, "
    sql = sql & "   c.TranDateTime AS CD10_TranDateTime, "
    sql = sql & "   a.TranDateTime AS ATM_TranDateTime, "
    sql = sql & "   c.CardNumber_Norm AS CD10_CardNumber, "
    sql = sql & "   a.CardNumber_Norm AS ATM_CardNumber, "
    sql = sql & "   c.CardLast4 AS CD10_CardLast4, "
    sql = sql & "   a.CardLast4 AS ATM_CardLast4, "
    sql = sql & "   c.AmountValue_Norm AS CD10_Amount, "
    sql = sql & "   a.AmountValue_Norm AS ATM_Amount, "
    sql = sql & "   c.SeqNo_Norm AS CD10_SeqNo, "
    sql = sql & "   a.TraceNo_Norm AS ATM_TraceNo, "
    sql = sql & "   a.ReferenceNo_Norm AS ATM_ReferenceNo, "
    sql = sql & "   a.TerminalID_Norm AS ATM_TerminalID, "
    sql = sql & "   c.TranDes_Norm AS CD10_TranDes, "
    sql = sql & "   a.TranCode_Norm AS ATM_TranCode, "
    sql = sql & "   c.SourceFile AS CD10_SourceFile, "
    sql = sql & "   a.SourceFile AS ATM_SourceFile, "
    sql = sql & "   ABS(DATEDIFF(DAY, c.TranDate, a.TranDate)) AS DateDiffDays, "
    sql = sql & "   ABS(DATEDIFF(MINUTE, c.TranDateTime, a.TranDateTime)) AS DateDiffMinutes, "
    
    ' Insert the dynamic rule logic here
    sql = sql & ruleLogic
    
    sql = sql & " FROM CD10F c "
    sql = sql & " INNER JOIN ATMR a "
    sql = sql & "   ON " & matchCondition & " "
    sql = sql & "  AND c.AmountValue_Norm = a.AmountValue_Norm "
    
    ' Only add date window condition if Rule 3 is enabled
    If gEnableRule3 = 1 Then
        sql = sql & "  AND a.TranDate >= DATEADD(DAY, "
        sql = sql & "       CASE WHEN DATENAME(WEEKDAY, c.TranDate) = 'Monday' THEN -? ELSE -? END, c.TranDate) "
        sql = sql & "  AND a.TranDate <= DATEADD(DAY, "
        sql = sql & "       CASE WHEN DATENAME(WEEKDAY, c.TranDate) IN ('Friday','Saturday','Sunday') THEN ? ELSE ? END, c.TranDate) "
    End If
    
    sql = sql & "), "
    sql = sql & "Ranked AS ( "
    sql = sql & " SELECT "
    sql = sql & "   cm.*, "
    sql = sql & "   ROW_NUMBER() OVER (PARTITION BY cm.CD10_Id ORDER BY cm.MatchScore DESC, cm.DateDiffDays ASC, cm.DateDiffMinutes ASC, cm.ATMRecon_Id) AS rn_cd10, "
    sql = sql & "   ROW_NUMBER() OVER (PARTITION BY cm.ATMRecon_Id ORDER BY cm.MatchScore DESC, cm.DateDiffDays ASC, cm.DateDiffMinutes ASC, cm.CD10_Id) AS rn_atm "
    sql = sql & " FROM CandidateMatches cm "
    sql = sql & "), "
    sql = sql & "FinalMatches AS ( "
    sql = sql & " SELECT * "
    sql = sql & " FROM Ranked "
    sql = sql & " WHERE rn_cd10 = 1 AND rn_atm = 1 "
    sql = sql & "), "
    sql = sql & "AllResults AS ( "
    sql = sql & "SELECT "
    sql = sql & "  'MATCHED' AS ReconStatus, "
    sql = sql & "  fm.MatchRule, "
    sql = sql & "  fm.MatchScore, "
    sql = sql & "  fm.CD10_TranDateTime, "
    sql = sql & "  fm.ATM_TranDateTime, "
    sql = sql & "  fm.CD10_CardLast4, "
    sql = sql & "  fm.ATM_CardLast4, "
    sql = sql & "  fm.CD10_Amount, "
    sql = sql & "  fm.ATM_Amount, "
    sql = sql & "  fm.CD10_SeqNo, "
    sql = sql & "  fm.ATM_TraceNo, "
    sql = sql & "  fm.ATM_ReferenceNo, "
    sql = sql & "  fm.CD10_TranDes, "
    sql = sql & "  fm.ATM_TranCode, "
    sql = sql & "  fm.CD10_TranDateTime AS SortDate, "
    sql = sql & "  fm.CD10_CardLast4 AS SortCard "
    sql = sql & "FROM FinalMatches fm "
    sql = sql & "UNION ALL "
    sql = sql & "SELECT "
    sql = sql & "  'UNMATCHED_CD10' AS ReconStatus, "
    sql = sql & "  'NO_MATCH' AS MatchRule, "
    sql = sql & "  CAST(NULL AS int) AS MatchScore, "
    sql = sql & "  c.TranDateTime AS CD10_TranDateTime, "
    sql = sql & "  CAST(NULL AS datetime) AS ATM_TranDateTime, "
    sql = sql & "  c.CardLast4 AS CD10_CardLast4, "
    sql = sql & "  CAST(NULL AS nvarchar(4)) AS ATM_CardLast4, "
    sql = sql & "  c.AmountValue_Norm AS CD10_Amount, "
    sql = sql & "  CAST(NULL AS decimal(18,2)) AS ATM_Amount, "
    sql = sql & "  c.SeqNo_Norm AS CD10_SeqNo, "
    sql = sql & "  CAST(NULL AS nvarchar(50)) AS ATM_TraceNo, "
    sql = sql & "  CAST(NULL AS nvarchar(50)) AS ATM_ReferenceNo, "
    sql = sql & "  c.TranDes_Norm AS CD10_TranDes, "
    sql = sql & "  CAST(NULL AS nvarchar(50)) AS ATM_TranCode, "
    sql = sql & "  c.TranDateTime AS SortDate, "
    sql = sql & "  c.CardLast4 AS SortCard "
    sql = sql & "FROM CD10F c "
    sql = sql & "LEFT JOIN FinalMatches fm ON c.Id = fm.CD10_Id "
    sql = sql & "WHERE fm.CD10_Id IS NULL "
    sql = sql & "UNION ALL "
    sql = sql & "SELECT "
    sql = sql & "  'UNMATCHED_ATM' AS ReconStatus, "
    sql = sql & "  'NO_MATCH' AS MatchRule, "
    sql = sql & "  CAST(NULL AS int) AS MatchScore, "
    sql = sql & "  CAST(NULL AS datetime) AS CD10_TranDateTime, "
    sql = sql & "  a.TranDateTime AS ATM_TranDateTime, "
    sql = sql & "  CAST(NULL AS nvarchar(4)) AS CD10_CardLast4, "
    sql = sql & "  a.CardLast4 AS ATM_CardLast4, "
    sql = sql & "  CAST(NULL AS decimal(18,2)) AS CD10_Amount, "
    sql = sql & "  a.AmountValue_Norm AS ATM_Amount, "
    sql = sql & "  CAST(NULL AS nvarchar(50)) AS CD10_SeqNo, "
    sql = sql & "  a.TraceNo_Norm AS ATM_TraceNo, "
    sql = sql & "  a.ReferenceNo_Norm AS ATM_ReferenceNo, "
    sql = sql & "  CAST(NULL AS nvarchar(100)) AS CD10_TranDes, "
    sql = sql & "  a.TranCode_Norm AS ATM_TranCode, "
    sql = sql & "  a.TranDateTime AS SortDate, "
    sql = sql & "  a.CardLast4 AS SortCard "
    sql = sql & "FROM ATMR a "
    sql = sql & "LEFT JOIN FinalMatches fm ON a.Id = fm.ATMRecon_Id "
    sql = sql & "WHERE fm.ATMRecon_Id IS NULL "
    sql = sql & ") "
    sql = sql & "SELECT "
    sql = sql & "  ReconStatus, "
    sql = sql & "  MatchRule, "
    sql = sql & "  MatchScore, "
    sql = sql & "  CD10_TranDateTime, "
    sql = sql & "  ATM_TranDateTime, "
    sql = sql & "  CD10_CardLast4, "
    sql = sql & "  ATM_CardLast4, "
    sql = sql & "  CD10_Amount, "
    sql = sql & "  ATM_Amount, "
    sql = sql & "  CD10_SeqNo, "
    sql = sql & "  ATM_TraceNo, "
    sql = sql & "  ATM_ReferenceNo, "
    sql = sql & "  CD10_TranDes, "
    sql = sql & "  ATM_TranCode "
    sql = sql & "FROM AllResults "
    
    ' Build dynamic ORDER BY
    Dim orderByClause
    Select Case LCase(Trim(gSortOrder))
        Case "datedesc"
            orderByClause = "ORDER BY ReconStatus, SortDate DESC, SortCard"
        Case "dateasc"
            orderByClause = "ORDER BY ReconStatus, SortDate ASC, SortCard"
        Case "cardthendate"
            orderByClause = "ORDER BY ReconStatus, SortCard, SortDate"
        Case "rulepriority"
            orderByClause = "ORDER BY ReconStatus, " & _
                            "CASE MatchRule " & _
                            "    WHEN 'RULE_1_LAST4_AMOUNT_REF' THEN 1 " & _
                            "    WHEN 'RULE_2_LAST4_AMOUNT_SAME_DAY' THEN 2 " & _
                            "    WHEN 'RULE_3_LAST4_AMOUNT_DATE_WINDOW' THEN 3 " & _
                            "    ELSE 99 " & _
                            "END, SortDate, SortCard"
        Case "rulethendate"
            orderByClause = "ORDER BY ReconStatus, MatchRule, SortDate, SortCard"
        Case "scoredesc"
            orderByClause = "ORDER BY ReconStatus, MatchScore DESC, SortDate, SortCard"
        Case "scoreasc"
            orderByClause = "ORDER BY ReconStatus, MatchScore ASC, SortDate, SortCard"
        Case "statusonly"
            orderByClause = "ORDER BY ReconStatus"
        Case "statuscard"
            orderByClause = "ORDER BY ReconStatus, SortCard, SortDate"
        Case "ruleandscore"
            orderByClause = "ORDER BY ReconStatus, " & _
                            "CASE MatchRule " & _
                            "    WHEN 'RULE_1_LAST4_AMOUNT_REF' THEN 1 " & _
                            "    WHEN 'RULE_2_LAST4_AMOUNT_SAME_DAY' THEN 2 " & _
                            "    WHEN 'RULE_3_LAST4_AMOUNT_DATE_WINDOW' THEN 3 " & _
                            "    ELSE 99 " & _
                            "END, MatchScore DESC, SortDate, SortCard"
        Case "amountdesc"
            orderByClause = "ORDER BY ReconStatus, " & _
                            "CASE WHEN CD10_Amount IS NOT NULL THEN CD10_Amount ELSE ATM_Amount END DESC, " & _
                            "SortDate, SortCard"
        Case "amountasc"
            orderByClause = "ORDER BY ReconStatus, " & _
                            "CASE WHEN CD10_Amount IS NOT NULL THEN CD10_Amount ELSE ATM_Amount END ASC, " & _
                            "SortDate, SortCard"
        Case Else
            LogLine "Using default sort order: ScoreDesc (highest confidence first)"
            orderByClause = "ORDER BY ReconStatus, MatchScore DESC, SortDate, SortCard"
    End Select

    sql = sql & orderByClause
    BuildReconSQL = sql
End Function
' =========================================================
'  Helpers
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
    Dim cfg
    Set cfg = ReadIniSafe(iniPath)
    If cfg Is Nothing Then
        LogLine "ERROR: Could not read INI file: " & iniPath
        Exit Sub
    End If

 
    gMode         = GetIniValue(cfg, "Report.Mode", "MTD")
    gRelOffset    = GetIniValue(cfg, "Report.RelativeOffset", "0")
    gYear         = GetIniValue(cfg, "Report.Year", "")
    gMonth        = GetIniValue(cfg, "Report.Month", "")
    gQuarter      = GetIniValue(cfg, "Report.Quarter", "")
    gStartDateStr = GetIniValue(cfg, "Report.StartDate", "")
    gEndDateStr   = GetIniValue(cfg, "Report.EndDate", "")

    ' Database settings
    gConnStr           = GetIniValue(cfg, "Database.ConnStr", "")
    gCD10FullTableName = GetIniValue(cfg, "Database.CD10Table", "[dbo].[EFT_CD10]")
    gATMFullTableName  = GetIniValue(cfg, "Database.ATMTable", "[dbo].[ATM_Recon]")

    ' Rule enable flags (default to 1 = enabled)
    gEnableRule1 = CLng(GetIniValue(cfg, "Reconciliation.EnableRule1", "1"))
    gEnableRule2 = CLng(GetIniValue(cfg, "Reconciliation.EnableRule2", "1"))
    gEnableRule3 = CLng(GetIniValue(cfg, "Reconciliation.EnableRule3", "1"))
    
    ' Custom base scores (defaults: 100, 80, 60)
    gRule1BaseScore = CLng(GetIniValue(cfg, "Reconciliation.Rule1BaseScore", "100"))
    gRule2BaseScore = CLng(GetIniValue(cfg, "Reconciliation.Rule2BaseScore", "80"))
    gRule3BaseScore = CLng(GetIniValue(cfg, "Reconciliation.Rule3BaseScore", "60"))
    
    ' Custom penalties (defaults: 5, 0, 2, 5)
    gDatePenaltyPerDay = CLng(GetIniValue(cfg, "Reconciliation.DatePenaltyPerDay", "5"))
    gTimePenaltyWithin1Hour = CLng(GetIniValue(cfg, "Reconciliation.TimePenaltyWithin1Hour", "0"))
    gTimePenaltyWithin24Hours = CLng(GetIniValue(cfg, "Reconciliation.TimePenaltyWithin24Hours", "2"))
    gTimePenaltyOver24Hours = CLng(GetIniValue(cfg, "Reconciliation.TimePenaltyOver24Hours", "5"))

    ' Date window settings
    gWeekdayLookbackDays     = CLng(GetIniValue(cfg, "Reconciliation.WeekdayLookbackDays", "1"))
    gWeekdayLookaheadDays    = CLng(GetIniValue(cfg, "Reconciliation.WeekdayLookaheadDays", "1"))
    gWeekendLookbackDays     = CLng(GetIniValue(cfg, "Reconciliation.WeekendLookbackDays", "3"))
    gWeekendLookaheadDays    = CLng(GetIniValue(cfg, "Reconciliation.WeekendLookaheadDays", "1"))
    gExcludePreAuth          = CLng(GetIniValue(cfg, "Reconciliation.ExcludePreAuth", "0"))
    gRequireApprovedResponse = CLng(GetIniValue(cfg, "Reconciliation.RequireApprovedResponse", "0"))
	
	
 
	gSortOrder = GetIniValue(cfg, "Report.SortOrder", "StatusThenDate")
	LogLine "Sort Order = " & gSortOrder
 
	gUseFullCardNumber = CLng(GetIniValue(cfg, "Reconciliation.UseFullCardNumber", "0"))
	LogLine "Use Full Card Number: " & IIf(gUseFullCardNumber = 1, "YES", "NO (using last 4 digits)")

    If IsNumeric(gRelOffset) Then
        gRelOffset = CInt(gRelOffset)
    Else
        gRelOffset = 0
    End If
    
    ' Log the rule configuration
    LogLine "Rule Configuration:"
    LogLine "  Rule 1 (Reference Match): " & IIf(gEnableRule1 = 1, "ENABLED (Score=" & gRule1BaseScore & ")", "DISABLED")
    LogLine "  Rule 2 (Same Day): " & IIf(gEnableRule2 = 1, "ENABLED (Score=" & gRule2BaseScore & ")", "DISABLED")
    LogLine "  Rule 3 (Date Window): " & IIf(gEnableRule3 = 1, "ENABLED (Score=" & gRule3BaseScore & ")", "DISABLED")
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
            If gRelOffset <> 0 Then baseDate = DateAdd("m", gRelOffset, baseDate)

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
            If p > 0 Then d(sect & "." & Trim(Left(line, p - 1))) = Trim(Mid(line, p + 1))
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

Function SafeString(ByVal v)
    If IsNull(v) Then
        SafeString = ""
    Else
        SafeString = Trim(CStr(v))
    End If
End Function

Function ChooseCard(ByVal v1, ByVal v2)
    If SafeString(v1) <> "" Then
        ChooseCard = SafeString(v1)
    Else
        ChooseCard = SafeString(v2)
    End If
End Function

Function ChooseAmount(ByVal v1, ByVal v2)
    If SafeString(v1) <> "" Then
        ChooseAmount = SafeString(v1)
    Else
        ChooseAmount = SafeString(v2)
    End If
End Function

Sub AddParameterInt(cmdObj, intVal)
    Const adInteger = 3
    Dim param
    On Error Resume Next
    Set param = cmdObj.CreateParameter("", adInteger, 1, , CLng(intVal))
    cmdObj.Parameters.Append param
    On Error GoTo 0
End Sub

' ============================================================================
' Print Reconciliation Rules Legend
' Set printLegend to True to print, False to skip
' ============================================================================
Sub PrintRulesLegend(printLegend)
    If Not printLegend Then Exit Sub
    
    fout.WriteLine ""
    fout.WriteLine "RECONCILIATION RULES"
    fout.WriteLine "-------------------"
    
    ' Only show enabled rules
    If gEnableRule1 = 1 Then
        fout.WriteLine "RULE_1 (Score " & gRule1BaseScore & "): LAST4 + AMOUNT + REFERENCE MATCH"
        fout.WriteLine "  - Same last 4 digits of card number"
        fout.WriteLine "  - Same transaction amount"
        fout.WriteLine "  - CD10 SeqNo matches ATM TraceNo OR ATM ReferenceNo"
        fout.WriteLine "  - Highest confidence - automatic reconciliation"
        fout.WriteLine ""
    End If
    
    If gEnableRule2 = 1 Then
        fout.WriteLine "RULE_2 (Score " & gRule2BaseScore & "): LAST4 + AMOUNT + SAME DAY"
        fout.WriteLine "  - Same last 4 digits of card number"
        fout.WriteLine "  - Same transaction amount"
        fout.WriteLine "  - Same transaction date"
        fout.WriteLine "  - High confidence - recommend auto-reconciliation"
        fout.WriteLine ""
    End If
    
    If gEnableRule3 = 1 Then
        fout.WriteLine "RULE_3 (Score " & gRule3BaseScore & "): LAST4 + AMOUNT + DATE WINDOW"
        fout.WriteLine "  - Same last 4 digits of card number"
        fout.WriteLine "  - Same transaction amount"
        fout.WriteLine "  - Dates within window (Weekend: " & gWeekendLookbackDays & "/" & gWeekendLookaheadDays & " days, Weekday: " & gWeekdayLookbackDays & "/" & gWeekdayLookaheadDays & " days)"
        fout.WriteLine "  - Medium confidence - may require review"
        fout.WriteLine ""
    End If
    
    fout.WriteLine "SCORE ADJUSTMENTS:"
    fout.WriteLine "  - Base score reduced by " & gDatePenaltyPerDay & " points per day difference"
    fout.WriteLine "  - Additional reduction for time differences:"
    fout.WriteLine "    * 0-60 minutes: " & gTimePenaltyWithin1Hour & " points"
    fout.WriteLine "    * 1-24 hours: " & gTimePenaltyWithin24Hours & " points"
    fout.WriteLine "    * >24 hours: " & gTimePenaltyOver24Hours & " points"
    fout.WriteLine ""
    
    ' Print current score interpretation guide
    fout.WriteLine "SCORE INTERPRETATION:"
    fout.WriteLine "  90-100: Exceptional match - Auto-reconcile"
    fout.WriteLine "  75-89:  Strong match - Low risk"
    fout.WriteLine "  60-74:  Good match - Review large amounts"
    fout.WriteLine "  50-59:  Fair match - Investigate"
    fout.WriteLine "  Below 50: Weak match - Manual review required"
    fout.WriteLine ""
End Sub