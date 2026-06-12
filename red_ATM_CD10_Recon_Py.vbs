Option Explicit

' ---------------- GLOBALS ----------------
Dim gMode, gRelOffset, gYear, gMonth, gQuarter
Dim gStartDateStr, gEndDateStr
Dim gConnStr
Dim gCD10FullTableName, gATMFullTableName
Dim gCD10Table2Part, gCD10ObjIdName
Dim gATMTable2Part, gATMObjIdName
Dim gIncludeCardPrefixes
Dim gCD10LookbackDays, gCD10LookaheadDays

Dim gIncludeTranDesc


Dim gExcludeMsgTypes     
Dim gIncludeRespCodes



 
Dim gATMLookbackDays, gATMLookaheadDays
Dim gTimeWindowMinutes

Dim gDebugMode
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
Sub StartDoc
    On Error Resume Next

    Dim startKey, endKey
    Dim sql, cmd
    Dim startDateObj, endDateObj
    Dim atmStartDateObj, atmEndDateObj
    Dim rowCount, matchedCount, unmatchedCD10Count, unmatchedATMCount
    Dim dateDict
    Set dateDict = CreateObject("Scripting.Dictionary")

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
    BuildTableNames gATMFullTableName,  gATMTable2Part,  gATMObjIdName

    If gDebugMode = 1 Then
        LogLine Pad("Resolved CD10 table:", 25) & gCD10Table2Part
        LogLine Pad("Resolved ATM table:",  25) & gATMTable2Part
    End If

    If Len(gCD10Table2Part) = 0 Or Len(gATMTable2Part) = 0 Then
        LogLine "ERROR: Could not parse table names."
        Exit Sub
    End If

	ComputeDateRange startKey, endKey
		startDateObj = FormatKeyToDate(startKey)
		endDateObj   = FormatKeyToDate(endKey)

		' Validate dates first
		If Not IsDate(startDateObj) Or Not IsDate(endDateObj) Then
			LogLine "ERROR: Invalid date range."
			Exit Sub
		End If

		' Calculate all expanded windows after validation
		Dim cd10StartDateObj, cd10EndDateObj
		cd10StartDateObj = DateAdd("d", -gCD10LookbackDays,  startDateObj)
		cd10EndDateObj   = DateAdd("d",  gCD10LookaheadDays, endDateObj)

		atmStartDateObj = DateAdd("d", -gATMLookbackDays,  startDateObj)
		atmEndDateObj   = DateAdd("d",  gATMLookaheadDays, endDateObj)

		If gDebugMode = 1 Then
			LogLine Pad("Report Mode:",      25) & gMode
			LogLine Pad("CD10 Date Range:",  25) & CStr(startDateObj)     & " to " & CStr(endDateObj)
			LogLine Pad("CD10 Expanded:",    25) & CStr(cd10StartDateObj) & " to " & CStr(cd10EndDateObj)
			LogLine Pad("ATM Date Range:",   25) & CStr(atmStartDateObj)  & " to " & CStr(atmEndDateObj)
			If gMode = "CUSTOM" Then
				LogLine Pad("Custom Start Date:", 25) & gStartDateStr
				LogLine Pad("Custom End Date:",   25) & gEndDateStr
			ElseIf gMode = "MONTHLY" Or gMode = "MTD" Then
				LogLine Pad("Year:",  25) & gYear
				LogLine Pad("Month:", 25) & gMonth
			ElseIf gMode = "WEEKLY" Then
				LogLine Pad("Relative Offset:", 25) & gRelOffset
			End If
		End If



    ' -------------------------------------------------------
    '  Open connection
    ' -------------------------------------------------------
    Set conn = CreateObject("ADODB.Connection")
    conn.Open gConnStr
    If Err.Number <> 0 Then
        LogLine "ERROR opening DB connection: " & Err.Description
        CleanupObjects conn, rs, cmd
        Exit Sub
    End If
    Err.Clear

    ' -------------------------------------------------------
    '  STEP 1 - Create indexes (fire and forget, no recordset)
    ' -------------------------------------------------------
    If gDebugMode = 1 Then LogLine "STEP 1: Creating indexes..."
    conn.Execute BuildIndexSQL()
    If Err.Number <> 0 Then
        LogLine "WARNING: Index creation error (non-fatal): " & Err.Description
        Err.Clear
    ElseIf gDebugMode = 1 Then
        LogLine "STEP 1: Indexes ready."
    End If

    ' -------------------------------------------------------
    '  STEP 2 - Run recon query
    ' -------------------------------------------------------
    If gDebugMode = 1 Then LogLine "STEP 2: Running recon query..."

    sql = BuildReconSQL()

	Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText      = sql
        .CommandType      = 1
        .CommandTimeout   = 300        ' 5 minutes - increase if still timing out
        .Parameters.Append .CreateParameter("@CD10Start", 135, 1, , cd10StartDateObj)
        .Parameters.Append .CreateParameter("@CD10End",   135, 1, , cd10EndDateObj)
        .Parameters.Append .CreateParameter("@ATMStart",  135, 1, , atmStartDateObj)
        .Parameters.Append .CreateParameter("@ATMEnd",    135, 1, , atmEndDateObj)
    End With
	
	
    Set rs = cmd.Execute
    If Err.Number <> 0 Then
        LogLine "ERROR executing recon query: " & Err.Description
        CleanupObjects conn, rs, cmd
        Exit Sub
    End If

    If gDebugMode = 1 Then
        LogLine "Available fields: " & GetAvailableFields(rs)
        If Not rs.EOF Then
            LogLine "First row - Card Number: " & Nz(rs("Card Number").Value)
            LogLine "First row - DateTime: "    & Nz(rs("DateTime").Value)
            LogLine "First row - TranCode: "    & Nz(rs("TranCode").Value)
            LogLine "First row - File: "        & Nz(rs("File").Value)
        End If
    End If
    Err.Clear

    ' -------------------------------------------------------
    '  Write header
    ' -------------------------------------------------------
    fout.WriteLine _
        Pad("Card Number",   25) & _
        Pad("ATM_DateTime",  25) & _
        Pad("ATM_TranCode",  15) & _
        Pad("TranAmt",       15) & _
        Pad("CD10_Acctnbr",  15) & _
        Pad("ATM_TermSeq",   15) & _
        Pad("CD10_SeqNo",    15) & _
        Pad("Status",        15) & _
        Pad("ATM_TraceNo",   15) & _
        Pad("CD10_RespCd",   15) & _
        Pad("CD10_TranDesc", 15) & _
        Pad("CD10_Date",     15) & _
        Pad("CD10_Time",     15) & _
        Pad("File",          15)

    ' -------------------------------------------------------
    '  Process rows
    ' -------------------------------------------------------
    rowCount          = 0
    matchedCount      = 0
    unmatchedCD10Count = 0
    unmatchedATMCount  = 0

    Do While Not rs.EOF
        Dim fileFlag, cd10Date, cd10Time
        fileFlag = Nz(rs("File").Value)

        If fileFlag = "MATCHED" Then
            matchedCount = matchedCount + 1
        ElseIf fileFlag = "NMOSI" Then
            unmatchedATMCount = unmatchedATMCount + 1
        Else
            unmatchedCD10Count = unmatchedCD10Count + 1
        End If

        Dim dt
        dt = rs("DateTime").Value
        If IsDate(dt) Then
            Dim dateKey
            dateKey = FormatDateTime(dt, 2)
            If Not dateDict.Exists(dateKey) Then
                dateDict.Add dateKey, Array(0, 0, 0)
            End If
            Dim cnt
            cnt = dateDict(dateKey)
            If fileFlag = "MATCHED" Then
                cnt(0) = cnt(0) + 1
            ElseIf fileFlag = "NMOSI" Then
                cnt(2) = cnt(2) + 1
            Else
                cnt(1) = cnt(1) + 1
            End If
            dateDict(dateKey) = cnt
        End If

        fout.WriteLine _
            Pad(Nz(rs("Card Number").Value), 25) & _
            Pad(Nz(rs("DateTime").Value),    25) & _
            Pad(Nz(rs("TranCode").Value),    15) & _
            Pad(Nz(rs("TranAmt").Value),     15) & _
            Pad(Nz(rs("Acctnbr").Value),     15) & _
            Pad(Nz(rs("TermSeq").Value),     15) & _
            Pad(Nz(rs("CD10Seq").Value),     15) & _
            Pad(Nz(rs("Status").Value),      15) & _
            Pad(Nz(rs("Misc1").Value),       15) & _
            Pad(Nz(rs("RespCd").Value),      15) & _
            Pad(Nz(rs("TranDesc").Value),    15) & _
            Pad(Nz(rs("Date").Value),        15) & _
            Pad(Nz(rs("Time").Value),        15) & _
            Pad(fileFlag,                    15)

        rowCount = rowCount + 1
        rs.MoveNext
    Loop

    ' -------------------------------------------------------
    '  Summary
    ' -------------------------------------------------------
    fout.WriteLine ""
    fout.WriteLine "SUMMARY"
    fout.WriteLine ""

    fout.WriteLine "REPORT CONFIGURATION:"
    fout.WriteLine Pad("Mode:", 25) & gMode
    If gMode = "CUSTOM" Then
        fout.WriteLine Pad("Date Range:", 25) & CStr(startDateObj) & " to " & CStr(endDateObj)
    ElseIf gMode = "MONTHLY" Then
        fout.WriteLine Pad("Year:",  25) & gYear
        fout.WriteLine Pad("Month:", 25) & gMonth
    ElseIf gMode = "WEEKLY" Then
        fout.WriteLine Pad("Relative Offset:", 25) & gRelOffset
    End If
    fout.WriteLine Pad("Time Window:",       25) & gTimeWindowMinutes & " minutes (" & Round(gTimeWindowMinutes / 1440, 1) & " days)"

	
	fout.WriteLine Pad("ATM Lookback Days:",  25) & gATMLookbackDays
    fout.WriteLine Pad("ATM Lookahead Days:", 25) & gATMLookaheadDays
    fout.WriteLine Pad("CD10 Lookback Days:", 25) & gCD10LookbackDays
    fout.WriteLine Pad("CD10 Lookahead Days:",25) & gCD10LookaheadDays
	
	
    LogLine Pad("Include TranDesc:",    25) & gIncludeTranDesc
	LogLine Pad("Exclude MsgTypes:",   25) & gExcludeMsgTypes
	LogLine Pad("Include RespCodes:",  25) & gIncludeRespCodes
    LogLine Pad("Include CardPrefixes:",25) & gIncludeCardPrefixes
    fout.WriteLine ""

    fout.WriteLine "OVERALL TOTALS:"
    fout.WriteLine Pad("Total Records:",          25) & Pad(rowCount,            25)
    fout.WriteLine Pad("Matched Records:",         25) & Pad(matchedCount,        25)
    fout.WriteLine Pad("Unmatched CD10 Records:",  25) & Pad(unmatchedCD10Count,  25)
    fout.WriteLine Pad("Unmatched ATM Records:",   25) & Pad(unmatchedATMCount,   25)

    If rowCount > 0 Then
        Dim matchRate
        matchRate = Round((matchedCount / rowCount) * 100, 2)
        fout.WriteLine Pad("Match Rate:", 25) & Pad(matchRate & "%", 25)
    End If
    fout.WriteLine ""

    fout.WriteLine "DATE BREAKDOWN:"
    fout.WriteLine Pad("Date", 25) & Pad("Matched", 25) & Pad("Unmatched CD10", 20) & Pad("Unmatched ATM", 20) & Pad("DayTotal", 15)

    If dateDict.Count > 0 Then
        Dim sortedDateKeys, dateIdx, dateIdx2, tempKey
        sortedDateKeys = dateDict.Keys

        For dateIdx = 0 To UBound(sortedDateKeys) - 1
            For dateIdx2 = dateIdx + 1 To UBound(sortedDateKeys)
                If sortedDateKeys(dateIdx) > sortedDateKeys(dateIdx2) Then
                    tempKey = sortedDateKeys(dateIdx)
                    sortedDateKeys(dateIdx) = sortedDateKeys(dateIdx2)
                    sortedDateKeys(dateIdx2) = tempKey
                End If
            Next
        Next

        For dateIdx = 0 To UBound(sortedDateKeys)
            Dim currentKey
            currentKey = sortedDateKeys(dateIdx)
            Dim currentCounts
            currentCounts = dateDict(currentKey)

            Dim dailyMatched, dailyUnmatchedCD10, dailyUnmatchedATM, dailyTotal, dailyMatchRate
            dailyMatched       = currentCounts(0)
            dailyUnmatchedCD10 = currentCounts(1)
            dailyUnmatchedATM  = currentCounts(2)
            dailyTotal         = dailyMatched + dailyUnmatchedCD10 + dailyUnmatchedATM
            If dailyTotal > 0 Then
                dailyMatchRate = Round((dailyMatched / dailyTotal) * 100, 2)
            Else
                dailyMatchRate = 0
            End If

            fout.WriteLine Pad(currentKey,         25) & _
                           Pad(dailyMatched,        25) & _
                           Pad(dailyUnmatchedCD10,  20) & _
                           Pad(dailyUnmatchedATM,   20) & _
                           Pad(dailyTotal,          15)
        Next
    Else
        fout.WriteLine "No date breakdown available"
    End If

    CleanupObjects conn, rs, cmd
End Sub

Sub CloseDoc
End Sub

' ============================================================
'  STEP 1 - Build index creation SQL (INI-driven table names)
'  Uses IF NOT EXISTS so safe to run every time
' ============================================================
Function BuildIndexSQL()
    Dim sql
    Dim cd10Obj, atmObj

    ' Use the ObjIdName (schema.table without brackets) for OBJECT_ID lookup
    cd10Obj = "'" & gCD10ObjIdName & "'"
    atmObj  = "'" & gATMObjIdName  & "'"

    sql = ""

    ' CD10 indexes
    sql = sql & "IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_CD10_LDate' AND object_id=OBJECT_ID(" & cd10Obj & ")) "
    sql = sql & "  CREATE INDEX IX_CD10_LDate     ON " & gCD10Table2Part & " (LDate); "

    sql = sql & "IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_CD10_CardAmt' AND object_id=OBJECT_ID(" & cd10Obj & ")) "
    sql = sql & "  CREATE INDEX IX_CD10_CardAmt   ON " & gCD10Table2Part & " (CardNumber, AmountValue); "

    ' ATM indexes
    sql = sql & "IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_ATM_TranDate' AND object_id=OBJECT_ID(" & atmObj & ")) "
    sql = sql & "  CREATE INDEX IX_ATM_TranDate   ON " & gATMTable2Part  & " (TranDate); "

    sql = sql & "IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_ATM_CardAmt' AND object_id=OBJECT_ID(" & atmObj & ")) "
    sql = sql & "  CREATE INDEX IX_ATM_CardAmt    ON " & gATMTable2Part  & " (CardNumber, AmountValue); "

    BuildIndexSQL = sql
End Function

' ============================================================
'  STEP 2 - Build recon SQL (fully INI-driven filters)
' ============================================================
Function BuildReconSQL()
    Dim sql
    Dim cardFilter, tranDescFilter, excludeMsgFilter, includeRespFilter
    Dim prefixes, tranDescs, respCodes, excludeCodes
    Dim i, item, parts

    ' -------------------------------------------------------
    '  Card prefix filter  (INI: IncludeCardPrefixes)
    ' -------------------------------------------------------
    cardFilter = ""
    If Trim(gIncludeCardPrefixes) <> "" And UCase(Trim(gIncludeCardPrefixes)) <> "ALL" Then
        prefixes = Split(gIncludeCardPrefixes, ",")
        Dim cardClauses()
        ReDim cardClauses(UBound(prefixes))
        For i = 0 To UBound(prefixes)
            cardClauses(i) = "CardNumber LIKE '" & Trim(prefixes(i)) & "%'"
        Next
        cardFilter = "AND (" & Join(cardClauses, " OR ") & ") "
    End If

    ' -------------------------------------------------------
    '  TranDesc filter  (INI: IncludeTranDesc)
    ' -------------------------------------------------------
    tranDescFilter = ""
    If Trim(gIncludeTranDesc) <> "" Then
        tranDescs = Split(gIncludeTranDesc, ",")
        Dim descClauses()
        ReDim descClauses(UBound(tranDescs))
        For i = 0 To UBound(tranDescs)
            descClauses(i) = "TranDesc LIKE '" & Trim(tranDescs(i)) & "%'"
        Next
        tranDescFilter = "AND (" & Join(descClauses, " OR ") & ") "
    End If


	' Exclude MsgType  (INI: ExcludeMsgTypes)
		excludeMsgFilter = ""
		If Trim(gExcludeMsgTypes) <> "" Then
			excludeCodes = Split(gExcludeMsgTypes, ",")
			Dim excClauses()
			ReDim excClauses(UBound(excludeCodes))
			For i = 0 To UBound(excludeCodes)
				excClauses(i) = "'" & Trim(excludeCodes(i)) & "'"
			Next
			excludeMsgFilter = "AND MsgType NOT IN (" & Join(excClauses, ",") & ") "
		End If

    ' -------------------------------------------------------
    '  Include RespCode last 2 chars  (INI: IncludeRespCodes)
    ' -------------------------------------------------------
    includeRespFilter = ""
    If Trim(gIncludeRespCodes) <> "" Then
        respCodes = Split(gIncludeRespCodes, ",")
        Dim respClauses()
        ReDim respClauses(UBound(respCodes))
        For i = 0 To UBound(respCodes)
            respClauses(i) = "'" & Trim(respCodes(i)) & "'"
        Next
        includeRespFilter = "AND RIGHT(LTRIM(RTRIM(RespCode)),2) IN (" & Join(respClauses, ",") & ") "
    End If

    ' -------------------------------------------------------
    '  Build SQL
    ' -------------------------------------------------------
    sql = ""
    sql = sql & "WITH CD10_Norm AS ("
    sql = sql & "  SELECT"
    sql = sql & "    LTRIM(RTRIM(ISNULL(CardNumber,'')))                      AS CardNumber,"
    sql = sql & "    TRY_CONVERT(DECIMAL(18,2), REPLACE(AmountValue,',',''))  AS Amount,"
    sql = sql & "    LDate, TranDesc, RespCode, MsgType, FromAccount, SeqNo"
    sql = sql & "  FROM " & gCD10Table2Part
    sql = sql & "  WHERE LDate BETWEEN ? AND ?"
    sql = sql & "    AND LTRIM(RTRIM(ISNULL(CardNumber,''))) <> '' "
    sql = sql & "    " & cardFilter
    sql = sql & "    " & tranDescFilter
    sql = sql & "    " & excludeMsgFilter
    sql = sql & "    " & includeRespFilter
    sql = sql & "    AND TRY_CONVERT(DECIMAL(18,2), REPLACE(AmountValue,',','')) IS NOT NULL"
    sql = sql & "    AND TRY_CONVERT(DECIMAL(18,2), REPLACE(AmountValue,',','')) <> 0"
    sql = sql & "),"

    sql = sql & "ATM_Norm AS ("
    sql = sql & "  SELECT"
    sql = sql & "    LTRIM(RTRIM(ISNULL(CardNumber,'')))                      AS CardNumber,"
    sql = sql & "    TRY_CONVERT(DECIMAL(18,2), REPLACE(AmountValue,',',''))  AS Amount,"
    sql = sql & "    TranDate, TranCode, TraceNo, TerminalID"
    sql = sql & "  FROM " & gATMTable2Part
    sql = sql & "  WHERE TranDate BETWEEN ? AND ?"
    sql = sql & "    AND LTRIM(RTRIM(ISNULL(CardNumber,''))) <> '' "
    sql = sql & "    " & cardFilter
    sql = sql & "    AND TRY_CONVERT(DECIMAL(18,2), REPLACE(AmountValue,',','')) IS NOT NULL"
    sql = sql & "    AND TRY_CONVERT(DECIMAL(18,2), REPLACE(AmountValue,',','')) <> 0"
    sql = sql & "),"

    sql = sql & "CD10_Deduped AS ("
    sql = sql & "  SELECT CardNumber, Amount, LDate, TranDesc, RespCode,"
    sql = sql & "         FromAccount, SeqNo,"
    sql = sql & "    ROW_NUMBER() OVER (PARTITION BY CardNumber, Amount"
    sql = sql & "                       ORDER BY LDate, SeqNo) AS rn"
    sql = sql & "  FROM CD10_Norm"
    sql = sql & "),"

    sql = sql & "ATM_Deduped AS ("
    sql = sql & "  SELECT CardNumber, Amount, TranDate, TranCode, TraceNo, TerminalID,"
    sql = sql & "    ROW_NUMBER() OVER (PARTITION BY CardNumber, Amount"
    sql = sql & "                       ORDER BY TranDate, TraceNo) AS rn"
    sql = sql & "  FROM ATM_Norm"
    sql = sql & "),"

    sql = sql & "Matched AS ("
    sql = sql & "  SELECT"
    sql = sql & "    c.CardNumber AS CD10CardNumber,"
    sql = sql & "    a.CardNumber AS ATMCardNumber,"
    sql = sql & "    c.Amount     AS CD10Amount,"
    sql = sql & "    a.Amount     AS ATMAmount,"
    sql = sql & "    c.FromAccount, c.SeqNo, c.RespCode, c.TranDesc,"
    sql = sql & "    a.TranCode,    a.TraceNo, a.TerminalID,"
    sql = sql & "    c.LDate    AS CD10Date,"
    sql = sql & "    a.TranDate AS ATMDate"
    sql = sql & "  FROM CD10_Deduped c"
    sql = sql & "  INNER JOIN ATM_Deduped a"
    sql = sql & "    ON  c.CardNumber = a.CardNumber"
    sql = sql & "    AND c.Amount     = a.Amount"
    sql = sql & "    AND c.rn         = 1"
    sql = sql & "    AND a.rn         = 1"
    sql = sql & ")"

    ' MATCHED
    sql = sql & " SELECT"
    sql = sql & "  1                                    AS SortOrder,"
    sql = sql & "  ATMCardNumber                        AS [Card Number],"
    sql = sql & "  ATMDate                              AS [DateTime],"
    sql = sql & "  TranCode                             AS [TranCode],"
    sql = sql & "  ATMAmount                            AS [TranAmt],"
    sql = sql & "  FromAccount                          AS [Acctnbr],"
    sql = sql & "  TerminalID                           AS [TermSeq],"
    sql = sql & "  SeqNo                                AS [CD10Seq],"
    sql = sql & "  'MATCHED'                            AS [Status],"
    sql = sql & "  TraceNo                              AS [Misc1],"
    sql = sql & "  RespCode                             AS [RespCd],"
    sql = sql & "  TranDesc                             AS [TranDesc],"
    sql = sql & "  CONVERT(VARCHAR(10), CD10Date, 101)  AS [Date],"
    sql = sql & "  ''                                   AS [Time],"
    sql = sql & "  'MATCHED'                            AS [File]"
    sql = sql & " FROM Matched"

    sql = sql & " UNION ALL"

    ' ATM only - NMOSI
    sql = sql & " SELECT"
    sql = sql & "  2                                    AS SortOrder,"
    sql = sql & "  a.CardNumber                         AS [Card Number],"
    sql = sql & "  a.TranDate                           AS [DateTime],"
    sql = sql & "  a.TranCode                           AS [TranCode],"
    sql = sql & "  a.Amount                             AS [TranAmt],"
    sql = sql & "  ''                                   AS [Acctnbr],"
    sql = sql & "  a.TerminalID                         AS [TermSeq],"
    sql = sql & "  ''                                   AS [CD10Seq],"
    sql = sql & "  'UNMATCHED_ATM'                      AS [Status],"
    sql = sql & "  a.TraceNo                            AS [Misc1],"
    sql = sql & "  ''                                   AS [RespCd],"
    sql = sql & "  ''                                   AS [TranDesc],"
    sql = sql & "  CONVERT(VARCHAR(10), a.TranDate, 101) AS [Date],"
    sql = sql & "  ''                                   AS [Time],"
    sql = sql & "  'NMOSI'                              AS [File]"
    sql = sql & " FROM ATM_Norm a"
    sql = sql & " LEFT JOIN Matched m"
    sql = sql & "   ON  a.CardNumber = m.ATMCardNumber"
    sql = sql & "   AND a.Amount     = m.ATMAmount"
    sql = sql & " WHERE m.ATMCardNumber IS NULL"

    sql = sql & " UNION ALL"

    ' CD10 only - NMSW
    sql = sql & " SELECT"
    sql = sql & "  3                                    AS SortOrder,"
    sql = sql & "  c.CardNumber                         AS [Card Number],"
    sql = sql & "  c.LDate                              AS [DateTime],"
    sql = sql & "  ''                                   AS [TranCode],"
    sql = sql & "  c.Amount                             AS [TranAmt],"
    sql = sql & "  c.FromAccount                        AS [Acctnbr],"
    sql = sql & "  ''                                   AS [TermSeq],"
    sql = sql & "  c.SeqNo                              AS [CD10Seq],"
    sql = sql & "  'UNMATCHED_CD10'                     AS [Status],"
    sql = sql & "  ''                                   AS [Misc1],"
    sql = sql & "  c.RespCode                           AS [RespCd],"
    sql = sql & "  c.TranDesc                           AS [TranDesc],"
    sql = sql & "  CONVERT(VARCHAR(10), c.LDate, 101)   AS [Date],"
    sql = sql & "  ''                                   AS [Time],"
    sql = sql & "  'NMSW'                               AS [File]"
    sql = sql & " FROM CD10_Norm c"
    sql = sql & " LEFT JOIN Matched m"
    sql = sql & "   ON  c.CardNumber = m.CD10CardNumber"
    sql = sql & "   AND c.Amount     = m.CD10Amount"
    sql = sql & " WHERE m.CD10CardNumber IS NULL"

    sql = sql & " ORDER BY SortOrder, [DateTime], [Card Number]"

    BuildReconSQL = sql
End Function

' ============================================================
'  Helper functions
' ============================================================
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

    gMode         = GetIniValue(cfg, "Report.Mode",          "MTD")
    gRelOffset    = GetIniValue(cfg, "Report.RelativeOffset", "0")
    gYear         = GetIniValue(cfg, "Report.Year",           "")
    gMonth        = GetIniValue(cfg, "Report.Month",          "")
    gQuarter      = GetIniValue(cfg, "Report.Quarter",        "")
    gStartDateStr = GetIniValue(cfg, "Report.StartDate",      "")
    gEndDateStr   = GetIniValue(cfg, "Report.EndDate",        "")

    gConnStr           = GetIniValue(cfg, "Database.ConnStr",   "")
    gCD10FullTableName = GetIniValue(cfg, "Database.CD10Table", "[dbo].[EFT_CD10]")
    gATMFullTableName  = GetIniValue(cfg, "Database.ATMTable",  "[dbo].[ATM_RECON]")

    gATMLookbackDays   = CLng(GetIniValue(cfg, "Reconciliation.ATMLookbackDays",   "0"))
    gATMLookaheadDays  = CLng(GetIniValue(cfg, "Reconciliation.ATMLookaheadDays",  "0"))
	gCD10LookbackDays  = CLng(GetIniValue(cfg, "Reconciliation.CD10LookbackDays",  "0"))
	gCD10LookaheadDays = CLng(GetIniValue(cfg, "Reconciliation.CD10LookaheadDays", "0"))
    gTimeWindowMinutes = CLng(GetIniValue(cfg, "Reconciliation.TimeWindowMinutes", "60"))

    gIncludeCardPrefixes = GetIniValue(cfg, "Reconciliation.IncludeCardPrefixes", "")
    gIncludeTranDesc     = GetIniValue(cfg, "Reconciliation.IncludeTranDesc",     "")
	gExcludeMsgTypes = GetIniValue(cfg, "Reconciliation.ExcludeMsgTypes", "")
	gIncludeRespCodes = GetIniValue(cfg, "Reconciliation.IncludeRespCodes", "")

    gDebugMode = CLng(GetIniValue(cfg, "Debug.Enabled", "0"))

	If gDebugMode = 1 Then
        LogLine "Debug Mode: ON"
        LogLine Pad("Time Window Minutes:", 25) & gTimeWindowMinutes
        LogLine Pad("ATM Lookback Days:",   25) & gATMLookbackDays
        LogLine Pad("ATM Lookahead Days:",  25) & gATMLookaheadDays
        LogLine Pad("Card Prefixes:",       25) & gIncludeCardPrefixes
        LogLine Pad("Include TranDesc:",    25) & gIncludeTranDesc
        LogLine Pad("Exclude MsgTypes:",    25) & gExcludeMsgTypes
        LogLine Pad("Include RespCodes:",   25) & gIncludeRespCodes
    End If

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
    Dim mode, today, dStart, dEnd, y, m, q, tmp, dow, weekStart, baseDate
    mode = UCase(Trim(gMode))
    today = Date()

    Select Case mode
        Case "CUSTOM"
            If gStartDateStr <> "" Then dStart = CDate(gStartDateStr) Else dStart = today
            If gEndDateStr   <> "" Then dEnd   = CDate(gEndDateStr)   Else dEnd   = today
        Case "DAILY"
            dStart = DateAdd("d", gRelOffset, today)
            dEnd   = dStart
        Case "WEEKLY"
            dow       = Weekday(today, vbMonday)
            weekStart = DateAdd("d", 1 - dow, today)
            If gRelOffset <> 0 Then weekStart = DateAdd("ww", gRelOffset, weekStart)
            dStart = weekStart
            dEnd   = DateAdd("d", 6, dStart)
        Case "QUARTERLY"
            If gYear    <> "" And IsNumeric(gYear)    Then y = CInt(gYear)    Else y = Year(today)
            If gQuarter <> "" And IsNumeric(gQuarter) Then q = CInt(gQuarter) Else q = DatePart("q", today)
            If gRelOffset <> 0 Then
                q = q + gRelOffset
                Do While q > 4: q = q - 4: y = y + 1: Loop
                Do While q < 1: q = q + 4: y = y - 1: Loop
            End If
            m      = (q - 1) * 3 + 1
            dStart = DateSerial(y, m, 1)
            dEnd   = DateAdd("d", -1, DateAdd("m", 3, dStart))
        Case "YTD"
            If gYear <> "" And IsNumeric(gYear) Then y = CInt(gYear) Else y = Year(today)
            dStart = DateSerial(y, 1, 1)
            dEnd   = today
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
            If gYear  <> "" And IsNumeric(gYear)  Then y = CInt(gYear)  Else y = Year(today)
            If gMonth <> "" And IsNumeric(gMonth) Then m = CInt(gMonth) Else m = Month(today)
            If gRelOffset <> 0 Then
                tmp = DateAdd("m", gRelOffset, DateSerial(y, m, 1))
                y   = Year(tmp)
                m   = Month(tmp)
            End If
            dStart = DateSerial(y, m, 1)
            dEnd   = DateAdd("d", -1, DateAdd("m", 1, dStart))
    End Select

    If mode <> "CUSTOM" And mode <> "DAILY" And dEnd > today Then dEnd = today

    startKey = FormatDateKey(dStart)
    endKey   = FormatDateKey(dEnd)
End Sub

Function FormatDateKey(d)
    FormatDateKey = Year(d) & Right("0" & Month(d), 2) & Right("0" & Day(d), 2)
End Function

Function FormatKeyToDate(yyyymmdd)
    If Len(yyyymmdd) <> 8 Then
        FormatKeyToDate = Date()
        Exit Function
    End If
    FormatKeyToDate = DateSerial(CInt(Left(yyyymmdd,4)), CInt(Mid(yyyymmdd,5,2)), CInt(Right(yyyymmdd,2)))
End Function

Function Pad(val, width)
    Dim s
    If IsNull(val) Then s = "" Else s = CStr(val)
    s = Trim(s)
    If Len(s) >= width Then
        Pad = Left(s, width)
    Else
        Pad = s & Space(width - Len(s))
    End If
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
    Set d  = CreateObject("Scripting.Dictionary")
    sect   = ""
    Do Until ts.AtEndOfStream
        line = Trim(ts.ReadLine)
        If Len(line) = 0 Or Left(line,1) = ";" Then
        ElseIf Left(line,1) = "[" And Right(line,1) = "]" Then
            sect = Mid(line, 2, Len(line) - 2)
        Else
            p = InStr(line, "=")
            If p > 0 Then d(sect & "." & Trim(Left(line,p-1))) = Trim(Mid(line,p+1))
        End If
    Loop
    ts.Close
    Set ReadIniSafe = d
    On Error GoTo 0
End Function

Sub BuildTableNames(fullName, out2Part, outObjIdName)
    Dim s, parts, schemaName, tableName
    s = Replace(Replace(Trim(fullName), "[", ""), "]", "")
    parts = Split(s, ".")
    schemaName = "dbo"
    tableName  = ""
    If UBound(parts) = 0 Then
        tableName = parts(0)
    ElseIf UBound(parts) = 1 Then
        schemaName = parts(0)
        tableName  = parts(1)
    Else
        schemaName = parts(UBound(parts) - 1)
        tableName  = parts(UBound(parts))
    End If
    If Len(tableName) = 0 Then
        out2Part    = ""
        outObjIdName = ""
        Exit Sub
    End If
    out2Part     = "[" & schemaName & "].[" & tableName & "]"
    outObjIdName = schemaName & "." & tableName
End Sub

Function RemoveBrackets(s)
    RemoveBrackets = Replace(Replace(s, "[", ""), "]", "")
End Function

Function GetAvailableFields(rs)
    Dim i, list
    list = ""
    On Error Resume Next
    For i = 0 To rs.Fields.Count - 1
        If list <> "" Then list = list & ", "
        list = list & rs.Fields(i).Name
    Next
    On Error GoTo 0
    GetAvailableFields = list
End Function

Function Nz(value)
    If IsNull(value) Then Nz = "" Else Nz = value
End Function