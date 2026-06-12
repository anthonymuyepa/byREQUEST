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
' ---------------------------------------------------------
' ---------------------------------------------------------
'  StartDoc - UPDATED with proper record counting
' ---------------------------------------------------------
Sub StartDoc
    On Error Resume Next

    Dim startKey, endKey
    Dim sql, cmd
    Dim startDateObj, endDateObj
    Dim atmStartDateObj, atmEndDateObj
    Dim rowCount
    Dim matchedCount, unmatchedCD10Count, unmatchedATMCount

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
    
    LogLine "Resolved CD10 table: " & gCD10Table2Part
    LogLine "Resolved ATM table: " & gATMTable2Part

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
    
    ' Log date conversion for debugging
    LogLine "Converted startKey: " & startKey & " to date: " & CStr(startDateObj)
    LogLine "Converted endKey: " & endKey & " to date: " & CStr(endDateObj)
    
    ' Validate dates
    If Not IsDate(startDateObj) Then
        LogLine "ERROR: Invalid start date from key: " & startKey
        Exit Sub
    End If
    If Not IsDate(endDateObj) Then
        LogLine "ERROR: Invalid end date from key: " & endKey
        Exit Sub
    End If

    ' Calculate ATM date range
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

    ' ========== TEMPORARY DEBUG QUERY - PLACE HERE ==========
    ' This will test if full card number + amount matching works
    On Error Resume Next
    Dim debugSql, debugRs
    Dim debugStartDate, debugEndDate, debugATMStart, debugATMEnd
    
    ' Use your actual date values from the calculated dates
    debugStartDate = Year(startDateObj) & "-" & Right("0" & Month(startDateObj), 2) & "-" & Right("0" & Day(startDateObj), 2)
    debugEndDate = Year(endDateObj) & "-" & Right("0" & Month(endDateObj), 2) & "-" & Right("0" & Day(endDateObj), 2)
    debugATMStart = Year(atmStartDateObj) & "-" & Right("0" & Month(atmStartDateObj), 2) & "-" & Right("0" & Day(atmStartDateObj), 2)
    debugATMEnd = Year(atmEndDateObj) & "-" & Right("0" & Month(atmEndDateObj), 2) & "-" & Right("0" & Day(atmEndDateObj), 2)
    
    debugSql = "SELECT " & _
               "    c.CardNumber AS CD10_CardNumber, " & _
               "    a.CardNumber AS ATM_CardNumber, " & _
               "    c.AmountValue_Norm AS CD10_Amount, " & _
               "    a.AmountValue_Norm AS ATM_Amount, " & _
               "    c.LDate AS CD10_Date, " & _
               "    a.TranDateTime AS ATM_Date, " & _
               "    c.SeqNo, " & _
               "    a.TraceNo, " & _
               "    a.ReferenceNo " & _
               "FROM " & _
               "    (SELECT " & _
               "        LTRIM(RTRIM(CardNumber)) AS CardNumber, " & _
               "        TRY_CONVERT(decimal(18,2), AmountValue) AS AmountValue_Norm, " & _
               "        LDate, " & _
               "        SeqNo " & _
               "     FROM " & gCD10Table2Part & " " & _
               "     WHERE LDate BETWEEN '" & debugStartDate & "' AND '" & debugEndDate & "' " & _
               "       AND TRY_CONVERT(decimal(18,2), AmountValue) IS NOT NULL " & _
               "    ) c " & _
               "INNER JOIN " & _
               "    (SELECT " & _
               "        LTRIM(RTRIM(CardNumber)) AS CardNumber, " & _
               "        ABS(TRY_CONVERT(decimal(18,2), AmountValue)) AS AmountValue_Norm, " & _
               "        TranDateTime, " & _
               "        TraceNo, " & _
               "        ReferenceNo " & _
               "     FROM " & gATMTable2Part & " " & _
               "     WHERE CAST(TranDateTime AS date) BETWEEN '" & debugATMStart & "' AND '" & debugATMEnd & "' " & _
               "       AND TRY_CONVERT(decimal(18,2), AmountValue) IS NOT NULL " & _
               "    ) a " & _
               "ON c.CardNumber = a.CardNumber " & _
               "AND c.AmountValue_Norm = a.AmountValue_Norm " & _
               "ORDER BY c.LDate, a.TranDateTime"
    
    LogLine "DEBUG: Running simple full card + amount match query..."
    Set debugRs = conn.Execute(debugSql)
    
    If Err.Number <> 0 Then
        LogLine "DEBUG ERROR: " & Err.Description
        Err.Clear
    Else
        If Not debugRs.EOF Then
            LogLine "========== SIMPLE QUERY MATCHES =========="
            LogLine "CD10_CardNumber | ATM_CardNumber | CD10_Amount | ATM_Amount | CD10_Date | ATM_Date"
            Dim matchCount
            matchCount = 0
            Do While Not debugRs.EOF And matchCount < 20
                LogLine "  " & debugRs("CD10_CardNumber") & " | " & _
                               debugRs("ATM_CardNumber") & " | " & _
                               debugRs("CD10_Amount") & " | " & _
                               debugRs("ATM_Amount") & " | " & _
                               debugRs("CD10_Date") & " | " & _
                               debugRs("ATM_Date")
                matchCount = matchCount + 1
                debugRs.MoveNext
            Loop
            LogLine "Total simple matches found (first " & matchCount & " shown)"
            LogLine "=========================================="
        Else
            LogLine "DEBUG: No matches found with full card number + amount"
            LogLine "DEBUG: Check if CardNumber formats match between tables"
        End If
        debugRs.Close
    End If
    Set debugRs = Nothing
    On Error GoTo 0
    ' ========== END TEMPORARY DEBUG QUERY ==========

    ' Build the main reconciliation SQL
    sql = BuildReconSQL()
    
    ' Check if SQL was built successfully
    If sql = "" Then
        LogLine "ERROR: Failed to build reconciliation SQL (no rules enabled?)"
        CleanupObjects conn, rs, cmd
        Exit Sub
    End If

    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = sql
        .CommandType = 1   ' adCmdText

        ' 1-2: CD10 date range
        .Parameters.Append .CreateParameter("@CD10Start", 135, 1, , startDateObj)
        .Parameters.Append .CreateParameter("@CD10End", 135, 1, , endDateObj)

        ' 3-4: CD10 filters
        .Parameters.Append .CreateParameter("@ExcludePreAuth", 3, 1, , gExcludePreAuth)
        .Parameters.Append .CreateParameter("@RequireApproved", 3, 1, , gRequireApprovedResponse)

        ' 5-6: ATM date range
        .Parameters.Append .CreateParameter("@ATMStart", 135, 1, , atmStartDateObj)
        .Parameters.Append .CreateParameter("@ATMEnd", 135, 1, , atmEndDateObj)

        ' 7-10: matching windows
        .Parameters.Append .CreateParameter("@WeekendLookback", 3, 1, , gWeekendLookbackDays)
        .Parameters.Append .CreateParameter("@WeekdayLookback", 3, 1, , gWeekdayLookbackDays)
        .Parameters.Append .CreateParameter("@WeekendLookahead", 3, 1, , gWeekendLookaheadDays)
        .Parameters.Append .CreateParameter("@WeekdayLookahead", 3, 1, , gWeekdayLookaheadDays)
    End With

    LogLine "Executing reconciliation query..."
    LogLine "CD10 StartDate = " & CStr(startDateObj)
    LogLine "CD10 EndDate   = " & CStr(endDateObj)
    LogLine "ATM StartDate  = " & CStr(atmStartDateObj)
    LogLine "ATM EndDate    = " & CStr(atmEndDateObj)
    LogLine "ExcludePreAuth = " & CStr(gExcludePreAuth)
    LogLine "RequireApprovedResponse = " & CStr(gRequireApprovedResponse)
    LogLine "WeekendLookbackDays = " & CStr(gWeekendLookbackDays)
    LogLine "WeekdayLookbackDays = " & CStr(gWeekdayLookbackDays)
    LogLine "WeekendLookaheadDays = " & CStr(gWeekendLookaheadDays)
    LogLine "WeekdayLookaheadDays = " & CStr(gWeekdayLookaheadDays)
    
    LogLine "Total parameters added: " & cmd.Parameters.Count

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

    ' Print rules legend (set to True to show)
    Call PrintRulesLegend(False)

    ' Initialize counters
    rowCount = 0
    matchedCount = 0
    unmatchedCD10Count = 0
    unmatchedATMCount = 0

    fout.WriteLine "ATM vs CD10 RECONCILIATION"
    fout.WriteLine "Date Range: " & CStr(startDateObj) & " to " & CStr(endDateObj)
    fout.WriteLine "Report Generated: " & Now()
    fout.WriteLine ""

    fout.WriteLine _
        Pad("STATUS", 16) & _
        Pad("RULE", 28) & _
        Pad("SCORE", 8) & _
        Pad("CD10 DATE/TIME", 20) & _
        Pad("ATM DATE/TIME", 20) & _
        Pad("LAST4", 8) & _
        Pad("AMOUNT", 12) & _
        Pad("CD10 SEQ", 12) & _
        Pad("ATM TRACE", 12) & _
        Pad("ATM REF", 12) & _
        Pad("CD10 DESC", 16) & _
        Pad("ATM CODE", 10)

    ' Process each row and count by status
    Do While Not rs.EOF
        ' Count by status
        Select Case SafeString(rs("ReconStatus"))
            Case "MATCHED"
                matchedCount = matchedCount + 1
            Case "UNMATCHED_CD10"
                unmatchedCD10Count = unmatchedCD10Count + 1
            Case "UNMATCHED_ATM"
                unmatchedATMCount = unmatchedATMCount + 1
        End Select
        
        ' Output the row
        fout.WriteLine _
            Pad(SafeString(rs("ReconStatus")), 16) & _
            Pad(SafeString(rs("MatchRule")), 28) & _
            Pad(SafeString(rs("MatchScore")), 8) & _
            Pad(SafeString(rs("CD10_TranDateTime")), 20) & _
            Pad(SafeString(rs("ATM_TranDateTime")), 20) & _
            Pad(ChooseCard(rs("CD10_CardLast4"), rs("ATM_CardLast4")), 8) & _
            Pad(ChooseAmount(rs("CD10_Amount"), rs("ATM_Amount")), 12) & _
            Pad(SafeString(rs("CD10_SeqNo")), 12) & _
            Pad(SafeString(rs("ATM_TraceNo")), 12) & _
            Pad(SafeString(rs("ATM_ReferenceNo")), 12) & _
            Pad(SafeString(rs("CD10_TranDes")), 16) & _
            Pad(SafeString(rs("ATM_TranCode")), 10)
        
        rowCount = rowCount + 1
        rs.MoveNext
    Loop
    
    ' Output summary statistics
    fout.WriteLine ""
    fout.WriteLine "SUMMARY"
    fout.WriteLine "-------"
    fout.WriteLine "Total Records: " & rowCount
    fout.WriteLine "Matched Records: " & matchedCount
    fout.WriteLine "Unmatched CD10 Records: " & unmatchedCD10Count
    fout.WriteLine "Unmatched ATM Records: " & unmatchedATMCount
    
    ' Calculate match rate if there are matched records
    If matchedCount > 0 Then
        Dim matchRate
        matchRate = Round((matchedCount / rowCount) * 100, 2)
        fout.WriteLine "Match Rate: " & matchRate & "%"
    End If

    LogLine "Reconciliation completed successfully. Total rows: " & rowCount
    LogLine "Summary - Matched: " & matchedCount & ", Unmatched CD10: " & unmatchedCD10Count & ", Unmatched ATM: " & unmatchedATMCount

    CleanupObjects conn, rs, cmd
    On Error GoTo 0
End Sub

Sub ProcessLine
End Sub

Sub CloseDoc
End Sub

' =========================================================
'  SQL builder - FIXED ORDER BY ISSUE
' =========================================================
Function BuildReconSQL()
    Dim sql

    sql = ""
    sql = sql & "WITH CD10 AS ("
    sql = sql & " SELECT "
    sql = sql & "   c.Id, "
    sql = sql & "   c.LDate, "
    sql = sql & "   LTRIM(RTRIM(ISNULL(c.CardNumber,''))) AS CardNumber_Norm, "
    sql = sql & "   RIGHT(LTRIM(RTRIM(ISNULL(c.CardNumber,''))), 4) AS CardLast4, "
    sql = sql & "   LTRIM(RTRIM(ISNULL(c.SeqNo,''))) AS SeqNo_Norm, "
    sql = sql & "   UPPER(LTRIM(RTRIM(ISNULL(c.TranDes,'')))) AS TranDes_Norm, "
    sql = sql & "   UPPER(LTRIM(RTRIM(ISNULL(c.Rsp,'')))) AS Rsp_Norm, "
    sql = sql & "   TRY_CONVERT(decimal(18,2), c.AmountValue) AS AmountValue_Norm, "
    sql = sql & "   c.SourceFile, "
    sql = sql & "   CAST(c.LDate AS datetime) AS TranDateTime, "
    sql = sql & "   CAST(c.LDate AS date) AS TranDate, "
    sql = sql & "   CASE "
    sql = sql & "     WHEN UPPER(LTRIM(RTRIM(ISNULL(c.TranDes,'')))) LIKE '%PRE%' "
    sql = sql & "       OR UPPER(LTRIM(RTRIM(ISNULL(c.TranDes,'')))) LIKE '%AUTH%' "
    sql = sql & "       OR UPPER(LTRIM(RTRIM(ISNULL(c.TranDes,'')))) LIKE '%HOLD%' "
    sql = sql & "     THEN 1 ELSE 0 END AS IsLikelyPreAuth, "
    sql = sql & "   CASE "
    sql = sql & "     WHEN UPPER(LTRIM(RTRIM(ISNULL(c.Rsp,'')))) IN ('A','APP','APPROVED','00','C') "
    sql = sql & "       OR LTRIM(RTRIM(ISNULL(c.Rsp,''))) = '' "
    sql = sql & "     THEN 1 ELSE 0 END AS IsApproved "
    sql = sql & " FROM " & gCD10Table2Part & " c "
    sql = sql & " WHERE c.LDate BETWEEN ? AND ? "
    sql = sql & "), "
    sql = sql & "CD10F AS ( "
    sql = sql & " SELECT * "
    sql = sql & " FROM CD10 "
    sql = sql & " WHERE AmountValue_Norm IS NOT NULL "
    sql = sql & "   AND AmountValue_Norm <> 0 "
    sql = sql & "   AND CardLast4 <> '' "
    sql = sql & "   AND (? = 0 OR IsLikelyPreAuth = 0) "
    sql = sql & "   AND (? = 0 OR IsApproved = 1) "
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
    sql = sql & "   a.SourceFile "
    sql = sql & " FROM " & gATMTable2Part & " a "
    sql = sql & " WHERE CAST(a.TranDateTime AS date) BETWEEN ? AND ? "
    sql = sql & "   AND TRY_CONVERT(decimal(18,2), a.AmountValue) IS NOT NULL "
    sql = sql & "   AND TRY_CONVERT(decimal(18,2), a.AmountValue) <> 0 "
    sql = sql & "   AND RIGHT(LTRIM(RTRIM(ISNULL(a.CardNumber,''))), 4) <> '' "
    sql = sql & "), "
    sql = sql & "CandidateMatches AS ( "
    sql = sql & " SELECT "
    sql = sql & "   c.Id AS CD10_Id, "
    sql = sql & "   a.Id AS ATMRecon_Id, "
    sql = sql & "   c.TranDateTime AS CD10_TranDateTime, "
    sql = sql & "   a.TranDateTime AS ATM_TranDateTime, "
    sql = sql & "   c.CardLast4 AS CD10_CardLast4, "
    sql = sql & "   a.CardLast4 AS ATM_CardLast4, "
    sql = sql & "   c.AmountValue_Norm AS CD10_Amount, "
    sql = sql & "   a.AmountValue_Norm AS ATM_Amount, "
    sql = sql & "   c.SeqNo_Norm AS CD10_SeqNo, "
    sql = sql & "   a.TraceNo_Norm AS ATM_TraceNo, "
    sql = sql & "   a.ReferenceNo_Norm AS ATM_ReferenceNo, "
    sql = sql & "   c.TranDes_Norm AS CD10_TranDes, "
    sql = sql & "   a.TranCode_Norm AS ATM_TranCode, "
    sql = sql & "   c.SourceFile AS CD10_SourceFile, "
    sql = sql & "   a.SourceFile AS ATM_SourceFile, "
    sql = sql & "   ABS(DATEDIFF(DAY, c.TranDate, a.TranDate)) AS DateDiffDays, "
    sql = sql & "   ABS(DATEDIFF(MINUTE, c.TranDateTime, a.TranDateTime)) AS DateDiffMinutes, "
    sql = sql & "   CASE "
    sql = sql & "     WHEN c.SeqNo_Norm <> '' "
    sql = sql & "      AND (c.SeqNo_Norm = a.TraceNo_Norm OR c.SeqNo_Norm = a.ReferenceNo_Norm) "
    sql = sql & "      AND c.CardLast4 = a.CardLast4 "
    sql = sql & "      AND c.AmountValue_Norm = a.AmountValue_Norm "
    sql = sql & "     THEN 'RULE_1_LAST4_AMOUNT_REF' "
    sql = sql & "     WHEN c.CardLast4 = a.CardLast4 "
    sql = sql & "      AND c.AmountValue_Norm = a.AmountValue_Norm "
    sql = sql & "      AND c.TranDate = a.TranDate "
    sql = sql & "     THEN 'RULE_2_LAST4_AMOUNT_SAME_DAY' "
    sql = sql & "     ELSE 'RULE_3_LAST4_AMOUNT_DATE_WINDOW' "
    sql = sql & "   END AS MatchRule, "
    sql = sql & "   (CASE "
    sql = sql & "      WHEN c.SeqNo_Norm <> '' "
    sql = sql & "       AND (c.SeqNo_Norm = a.TraceNo_Norm OR c.SeqNo_Norm = a.ReferenceNo_Norm) "
    sql = sql & "       AND c.CardLast4 = a.CardLast4 "
    sql = sql & "       AND c.AmountValue_Norm = a.AmountValue_Norm THEN 100 "
    sql = sql & "      WHEN c.CardLast4 = a.CardLast4 "
    sql = sql & "       AND c.AmountValue_Norm = a.AmountValue_Norm "
    sql = sql & "       AND c.TranDate = a.TranDate THEN 80 "
    sql = sql & "      ELSE 60 "
    sql = sql & "    END "
    sql = sql & "    - ABS(DATEDIFF(DAY, c.TranDate, a.TranDate)) * 5 "
    sql = sql & "    - CASE "
    sql = sql & "        WHEN ABS(DATEDIFF(MINUTE, c.TranDateTime, a.TranDateTime)) <= 60 THEN 0 "
    sql = sql & "        WHEN ABS(DATEDIFF(MINUTE, c.TranDateTime, a.TranDateTime)) <= 1440 THEN 2 "
    sql = sql & "        ELSE 5 "
    sql = sql & "      END) AS MatchScore "
    sql = sql & " FROM CD10F c "
    sql = sql & " INNER JOIN ATMR a "
    sql = sql & "   ON c.CardLast4 = a.CardLast4 "
    sql = sql & "  AND c.CardLast4 <> '' "
    sql = sql & "  AND c.AmountValue_Norm = a.AmountValue_Norm "
    sql = sql & "  AND a.TranDate >= DATEADD(DAY, "
    sql = sql & "       CASE WHEN DATENAME(WEEKDAY, c.TranDate) = 'Monday' THEN -? ELSE -? END, c.TranDate) "
    sql = sql & "  AND a.TranDate <= DATEADD(DAY, "
    sql = sql & "       CASE WHEN DATENAME(WEEKDAY, c.TranDate) IN ('Friday','Saturday','Sunday') THEN ? ELSE ? END, c.TranDate) "
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
    sql = sql & "ORDER BY ReconStatus, SortDate, SortCard"

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

    gConnStr           = GetIniValue(cfg, "Database.ConnStr", "")
    gCD10FullTableName = GetIniValue(cfg, "Database.CD10Table", "[dbo].[EFT_CD10]")
    gATMFullTableName  = GetIniValue(cfg, "Database.ATMTable", "[dbo].[ATM_Recon]")

    gWeekdayLookbackDays     = CLng(GetIniValue(cfg, "Reconciliation.WeekdayLookbackDays", "1"))
    gWeekdayLookaheadDays    = CLng(GetIniValue(cfg, "Reconciliation.WeekdayLookaheadDays", "1"))
    gWeekendLookbackDays     = CLng(GetIniValue(cfg, "Reconciliation.WeekendLookbackDays", "3"))
    gWeekendLookaheadDays    = CLng(GetIniValue(cfg, "Reconciliation.WeekendLookaheadDays", "1"))
    gExcludePreAuth          = CLng(GetIniValue(cfg, "Reconciliation.ExcludePreAuth", "0"))
    gRequireApprovedResponse = CLng(GetIniValue(cfg, "Reconciliation.RequireApprovedResponse", "0"))

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