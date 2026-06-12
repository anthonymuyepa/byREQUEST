Option Explicit

' =========================================================
'  ACH Period Report (INI-driven) - Status='NP'
'  - Reads INI for Report params + Database.ConnStr + Database.TableName
'  - Computes date range from Mode / RelativeOffset
'  - Queries table by Post_Date BETWEEN start/end AND Status='NP'
'  - SELECT list matches ACH_EPSA columns (excluding id, Import_Timestamp)
' =========================================================

' ---------------- GLOBALS ----------------
Dim gMode, gRelOffset, gYear, gMonth, gQuarter
Dim gStartDateStr, gEndDateStr
Dim gConnStr, gFullTableName, gTable2Part, gObjIdName

Dim conn, rs
Dim INI_PATH: INI_PATH = "C:\byREQUEST\Templates\EPSA_DAILY.ini"
    INI_PATH = "S:\CarolinaTrust\byREQUEST\Templates\EPSA_MTD.ini"

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
    Dim errMsg
    Dim sql, cmd
    Dim rowCount, totalAmount, amount
    Dim periodDisplay, startDateObj, endDateObj

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
    If Len(gTable2Part) = 0 Then
        LogLine "ERROR: Could not parse TableName: " & gFullTableName
        Exit Sub
    End If

    ComputeDateRange startKey, endKey
    If startKey = "" Or endKey = "" Then
        LogLine "ERROR: Could not compute date range from parameters."
        Exit Sub
    End If

    startDateObj = FormatKeyToDate(startKey)
    endDateObj   = FormatKeyToDate(endKey)

    Select Case UCase(Trim(gMode))
        Case "DAILY"
            periodDisplay = Fmt(startDateObj, "dddd, mmmm d, yyyy")
        Case "WEEKLY"
            periodDisplay = "Week of " & Fmt(startDateObj, "mmm d") & " to " & Fmt(endDateObj, "mmm d, yyyy")
        Case "MONTHLY"
            periodDisplay = Fmt(startDateObj, "mmmm yyyy")
        Case "QUARTERLY"
            Dim quarterNum
            quarterNum = DatePart("q", startDateObj)
            periodDisplay = "Q" & quarterNum & " " & Year(startDateObj)
        Case "YTD"
            periodDisplay = "Year to Date " & Year(startDateObj)
        Case "CUSTOM"
            periodDisplay = Fmt(startDateObj, "mmm d, yyyy") & " to " & Fmt(endDateObj, "mmm d, yyyy")
        Case "MTD"
            periodDisplay = "MTD " & Fmt(startDateObj, "mmmm yyyy") & " (" & _
                            Fmt(startDateObj, "mmm d, yyyy") & " to " & Fmt(endDateObj, "mmm d, yyyy") & ")"
        Case Else
            periodDisplay = Fmt(startDateObj, "mmm d, yyyy") & " to " & Fmt(endDateObj, "mmm d, yyyy")
    End Select

    Spoolfile.User1 = UCase(gMode) & "-" & _
                      Right("0" & Month(endDateObj), 2) & _
                      Right("0" & Day(endDateObj), 2) & _
                      Year(endDateObj)

    'MsgBox Spoolfile.User1

    ' Open DB
    Set conn = CreateObject("ADODB.Connection")
    Err.Clear
    conn.Open gConnStr
    If Err.Number <> 0 Then
        LogLine "ERROR opening DB connection: " & Err.Description & " (Err=" & Err.Number & ")"
        Exit Sub
    End If

    ' ' =====================================================
    ' ' MAIN QUERY: Status='NP' (NO id, NO Import_Timestamp)
    ' ' =====================================================
    ' sql = "SELECT " & _
          ' "[Report],[Post_Date],[Post_Time],[Settlement_Date],[Company_Name]," & _
          ' "[ACH_Batch_ID],[PCTL],[LCTL],[Status],[Tran_Type],[Tran_Code],[Effective_Date]," & _
          ' "[Amount],[Member_Nbr],[Mbr_Name],[Fee],[ACH_Ind_Name],[DD],[Trace_Nbr],[Description]," & _
          ' "[Type],[Post_Seq],[File_Post_Seq],[Source_Name],[Check_Number],[SDC_Number],[S_L],[RR]," & _
          ' "[RDFI_ID],[DFI_ACCOUNT],[Individual_ID],[MESSAGE],[EPSA_File] " & _
          ' "FROM " & gTable2Part & " " & _
          ' "WHERE [Post_Date] BETWEEN ? AND ? " & _
          ' "AND ( [Status] <> 'NP' OR [Status] IS NULL ) " & _
          ' "ORDER BY [Post_Date], [Post_Time]"

    ' Set cmd = CreateObject("ADODB.Command")
    ' With cmd
        ' .ActiveConnection = conn
        ' .CommandText = sql
        ' .CommandType = 1  ' adCmdText
        ' .Parameters.Append .CreateParameter("", 133, 1, , startDateObj) ' adDBDate
        ' .Parameters.Append .CreateParameter("", 133, 1, , endDateObj)
    ' End With

    ' Err.Clear
    ' Set rs = cmd.Execute
    ' If Err.Number <> 0 Then
        ' LogLine "ERROR executing query: " & Err.Description & " (Err=" & Err.Number & ")"
        ' CleanupObjects conn, rs, cmd
        ' Exit Sub
    ' End If

    ' If rs.EOF Then
        ' fout.WriteLine "No records found (Status <> 'NP' or Status IS NULL) for " & periodDisplay & "."
        ' CleanupObjects conn, rs, cmd
        ' Exit Sub
    ' End If

    ' ' Title
    ' 'fout.WriteLine "ACH EPSA Report (Status <> NP) - " & periodDisplay

    ' ' Headers (fixed width)
    ' fout.WriteLine _
        ' Pad("Report", 10) & _
        ' Pad("Post_Date", 12) & _
        ' Pad("Post_Time", 12) & _
        ' Pad("Settlement", 16) & _
        ' Pad("Company_Name", 25) & _
        ' Pad("Batch_ID", 16) & _
        ' Pad("PCTL", 8) & _
        ' Pad("LCTL", 8) & _
        ' Pad("Status", 8) & _
        ' Pad("Tran_Type", 12) & _
        ' Pad("Tran_Code", 12) & _
        ' Pad("Effective", 16) & _
        ' Pad("Amount", 12) & _
        ' Pad("Member_Nbr", 15) & _
        ' Pad("Mbr_Name", 25) & _
        ' Pad("Fee", 10) & _
        ' Pad("ACH_Ind_Name", 25) & _
        ' Pad("DD", 4) & _
        ' Pad("Trace_Nbr", 20) & _
        ' Pad("Description", 30) & _
        ' Pad("Type", 8) & _
        ' Pad("Post_Seq", 10) & _
        ' Pad("File_Post_Seq", 15) & _
        ' Pad("Source_Name", 14) & _
        ' Pad("Check_Nbr", 12) & _
        ' Pad("SDC_Number", 12) & _
        ' Pad("S_L", 4) & _
        ' Pad("RR", 4) & _
        ' Pad("RDFI_ID", 10) & _
        ' Pad("DFI_ACCOUNT", 14) & _
        ' Pad("Individual_ID", 14) & _
        ' Pad("MESSAGE", 35) & _
        ' Pad("EPSA_File", 14)

    ' rowCount = 0
    ' totalAmount = 0

    ' Do While Not rs.EOF
        ' rowCount = rowCount + 1

        ' amount = 0
        ' If Not IsNull(rs("Amount")) And IsNumeric(rs("Amount")) Then
            ' amount = CDbl(rs("Amount"))
        ' End If
        ' totalAmount = totalAmount + amount

        ' fout.WriteLine _
            ' Pad(rs("Report"), 10) & _
            ' Pad(FmtDate8(rs("Post_Date")), 12) & _
            ' Pad(FmtTimeShort(rs("Post_Time")), 12) & _
            ' Pad(FmtDate8(rs("Settlement_Date")), 16) & _
            ' Pad(rs("Company_Name"), 25) & _
            ' Pad(rs("ACH_Batch_ID"), 16) & _
            ' Pad(rs("PCTL"), 8) & _
            ' Pad(rs("LCTL"), 8) & _
            ' Pad(rs("Status"), 8) & _
            ' Pad(rs("Tran_Type"), 12) & _
            ' Pad(rs("Tran_Code"), 12) & _
            ' Pad(FmtDate8(rs("Effective_Date")), 16) & _
            ' Pad(FormatNumber(amount, 2), 12) & _
            ' Pad(rs("Member_Nbr"), 15) & _
            ' Pad(rs("Mbr_Name"), 25) & _
            ' Pad(rs("Fee"), 10) & _
            ' Pad(rs("ACH_Ind_Name"), 25) & _
            ' Pad(rs("DD"), 4) & _
            ' Pad(rs("Trace_Nbr"), 20) & _
            ' Pad(rs("Description"), 30) & _
            ' Pad(rs("Type"), 8) & _
            ' Pad(rs("Post_Seq"), 10) & _
            ' Pad(rs("File_Post_Seq"), 15) & _
            ' Pad(rs("Source_Name"), 14) & _
            ' Pad(rs("Check_Number"), 12) & _
            ' Pad(rs("SDC_Number"), 12) & _
            ' Pad(rs("S_L"), 4) & _
            ' Pad(rs("RR"), 4) & _
            ' Pad(rs("RDFI_ID"), 10) & _
            ' Pad(rs("DFI_ACCOUNT"), 14) & _
            ' Pad(rs("Individual_ID"), 14) & _
            ' Pad(rs("MESSAGE"), 35) & _
            ' Pad(rs("EPSA_File"), 14)

        ' rs.MoveNext
    ' Loop

    ' fout.WriteLine ""
    ' 'fout.WriteLine "Total Records: " & rowCount & " | Total Amount: " & FormatNumber(totalAmount, 2)
	
	


	' ------------------------------------------------------------
	' 2B) DETAIL - CLEARED (Investigation)
	' ------------------------------------------------------------
	Dim sqlClr, cmdClr, rsClr

	sqlClr = _
	"WITH epsa_ranked AS ( " & _
	"  SELECT " & _
	"    'EPSA' AS Report, " & _
	"    e.Post_Date, e.Post_Time, " & _
	"    COALESCE(TRY_CONVERT(date, e.Settlement_Date), TRY_CONVERT(date, e.Settlement_Date, 10), TRY_CONVERT(date, e.Settlement_Date, 1), TRY_CONVERT(date, e.Settlement_Date, 101)) AS Settlement_Date, " & _
	"    e.Company_Name, e.ACH_Batch_ID, e.PCTL, e.LCTL, e.Status, e.Tran_Type, e.Tran_Code, " & _
	"    COALESCE(TRY_CONVERT(date, e.Effective_Date), TRY_CONVERT(date, e.Effective_Date, 10), TRY_CONVERT(date, e.Effective_Date, 1), TRY_CONVERT(date, e.Effective_Date, 101)) AS Effective_Date, " & _
	"    LTRIM(RTRIM(e.Amount)) AS Amount, " & _
	"    e.Member_Nbr, e.Mbr_Name, e.Fee, e.ACH_Ind_Name, e.DD, " & _
	"    LTRIM(RTRIM(e.Trace_Nbr)) AS Trace_Nbr, " & _
	"    e.Description, e.Type, e.Post_Seq, e.File_Post_Seq, e.Source_Name, " & _
	"    e.Check_Number, e.SDC_Number, e.S_L, e.RR, e.RDFI_ID, e.DFI_ACCOUNT, e.Individual_ID, e.[MESSAGE], e.EPSA_File, " & _
	"    ROW_NUMBER() OVER ( " & _
	"      PARTITION BY LTRIM(RTRIM(e.Trace_Nbr)) " & _
	"      ORDER BY e.Post_Date DESC, e.id DESC " & _
	"    ) AS rn " & _
	"  FROM dbo.ACH_EPSA e " & _
	"  WHERE e.Trace_Nbr IS NOT NULL " & _
	"), epsa_winner AS ( " & _
	"  SELECT * FROM epsa_ranked WHERE rn = 1 AND Settlement_Date IS NOT NULL " & _
	"), ertr_norm AS ( " & _
	"  SELECT LTRIM(RTRIM(r.Trace_Number)) AS TraceKey " & _
	"  FROM dbo.ACH_ERTR r WHERE r.Trace_Number IS NOT NULL " & _
	") " & _
	"SELECT " & _
	"  e.Report, e.Post_Date, e.Post_Time, e.Settlement_Date, e.Company_Name, e.ACH_Batch_ID, e.PCTL, e.LCTL, e.Status, e.Tran_Type, e.Tran_Code, e.Effective_Date, " & _
	"  e.Amount, e.Member_Nbr, e.Mbr_Name, e.Fee, e.ACH_Ind_Name, e.DD, e.Trace_Nbr, e.Description, e.Type, e.Post_Seq, e.File_Post_Seq, e.Source_Name, " & _
	"  e.Check_Number, e.SDC_Number, e.S_L, e.RR, e.RDFI_ID, e.DFI_ACCOUNT, e.Individual_ID, e.[MESSAGE], e.EPSA_File " & _
	"FROM epsa_winner e " & _
	"LEFT JOIN ertr_norm r ON r.TraceKey = e.Trace_Nbr " & _
	"WHERE e.Settlement_Date BETWEEN ? AND ? " & _
	"  AND r.TraceKey IS NULL " & _
	"  AND e.Status = 'P' " & _
	"ORDER BY e.Settlement_Date, e.Trace_Nbr"

	Set cmdClr = CreateObject("ADODB.Command")
	With cmdClr
		.ActiveConnection = conn
		.CommandText = sqlClr
		.CommandType = 1
		.Parameters.Append .CreateParameter("", 133, 1, , startDateObj)
		.Parameters.Append .CreateParameter("", 133, 1, , endDateObj)
	End With

	Set rsClr = cmdClr.Execute


'	fout.WriteLine "REPORT    ACH TRANSACTIONS THAT CLEARED (NO RETURN FOUND)"
    fout.WriteLine _
        Pad("Report", 10) & _
        Pad("Post_Date", 12) & _
        Pad("Post_Time", 12) & _
        Pad("Settlement", 16) & _
        Pad("Company_Name", 25) & _
        Pad("Batch_ID", 16) & _
        Pad("PCTL", 8) & _
        Pad("LCTL", 8) & _
        Pad("Status", 8) & _
        Pad("Tran_Type", 12) & _
        Pad("Tran_Code", 12) & _
        Pad("Effective", 16) & _
        Pad("Amount", 12) & _
        Pad("Member_Nbr", 15) & _
        Pad("Mbr_Name", 25) & _
        Pad("Fee", 10) & _
        Pad("ACH_Ind_Name", 25) & _
        Pad("DD", 4) & _
        Pad("Trace_Nbr", 20) & _
        Pad("Description", 30) & _
        Pad("Type", 8) & _
        Pad("Post_Seq", 10) & _
        Pad("File_Post_Seq", 15) & _
        Pad("Source_Name", 14) & _
        Pad("Check_Nbr", 12) & _
        Pad("SDC_Number", 12) & _
        Pad("S_L", 4) & _
        Pad("RR", 4) & _
        Pad("RDFI_ID", 10) & _
        Pad("DFI_ACCOUNT", 14) & _
        Pad("Individual_ID", 14) & _
        Pad("MESSAGE", 35) & _
        Pad("EPSA_File", 14)


	If rsClr.EOF Then
		fout.WriteLine "CLEARED   No rows returned for this period."
	Else
		Do While Not rsClr.EOF
			Dim amountText, amountNum

			amountText = ""
			If Not IsNull(rsClr("Amount")) Then amountText = CStr(rsClr("Amount"))

			amountNum = 0
			If Len(Trim(amountText)) > 0 Then
				On Error Resume Next
				amountNum = CDbl(Replace(Replace(Replace(amountText, ",", ""), "$", ""), " ", ""))
				If Err.Number <> 0 Then
					amountNum = 0
					Err.Clear
				End If
				On Error GoTo 0
			End If

			fout.WriteLine _
				Pad(rsClr("Report"), 10) & _
				Pad(FmtDate8(rsClr("Post_Date")), 12) & _
				Pad(FmtTimeShort(rsClr("Post_Time")), 12) & _
				Pad(FmtDate8(rsClr("Settlement_Date")), 16) & _
				Pad(rsClr("Company_Name"), 25) & _
				Pad(rsClr("ACH_Batch_ID"), 16) & _
				Pad(rsClr("PCTL"), 8) & _
				Pad(rsClr("LCTL"), 8) & _
				Pad(rsClr("Status"), 8) & _
				Pad(rsClr("Tran_Type"), 12) & _
				Pad(rsClr("Tran_Code"), 12) & _
				Pad(FmtDate8(rsClr("Effective_Date")), 16) & _
				Pad(FormatNumber(amountNum, 2), 12) & _
				Pad(rsClr("Member_Nbr"), 15) & _
				Pad(rsClr("Mbr_Name"), 25) & _
				Pad(rsClr("Fee"), 10) & _
				Pad(rsClr("ACH_Ind_Name"), 25) & _
				Pad(rsClr("DD"), 4) & _
				Pad(rsClr("Trace_Nbr"), 20) & _
				Pad(rsClr("Description"), 30) & _
				Pad(rsClr("Type"), 8) & _
				Pad(rsClr("Post_Seq"), 10) & _
				Pad(rsClr("File_Post_Seq"), 15) & _
				Pad(rsClr("Source_Name"), 14) & _
				Pad(rsClr("Check_Number"), 12) & _
				Pad(rsClr("SDC_Number"), 12) & _
				Pad(rsClr("S_L"), 4) & _
				Pad(rsClr("RR"), 4) & _
				Pad(rsClr("RDFI_ID"), 10) & _
				Pad(rsClr("DFI_ACCOUNT"), 14) & _
				Pad(rsClr("Individual_ID"), 14) & _
				Pad(rsClr("MESSAGE"), 35) & _
				Pad(rsClr("EPSA_File"), 14)
			rsClr.MoveNext
		Loop
	End If

	fout.WriteLine ""



	Dim viewName, sqlViewDDL

	viewName = "dbo.vw_ACH_Recon_Summary"

	' ------------------------------------------------------------
	' 0) Ensure SUMMARY view exists (safe date parsing)
	' ------------------------------------------------------------
	sqlViewDDL = _
	"CREATE OR ALTER VIEW " & viewName & " AS " & vbCrLf & _
	"WITH epsa_ranked AS ( " & vbCrLf & _
	"    SELECT " & vbCrLf & _
	"        COALESCE( " & vbCrLf & _
	"          TRY_CONVERT(date, e.Settlement_Date), " & vbCrLf & _
	"          TRY_CONVERT(date, e.Settlement_Date, 10), " & vbCrLf & _
	"          TRY_CONVERT(date, e.Settlement_Date, 1), " & vbCrLf & _
	"          TRY_CONVERT(date, e.Settlement_Date, 101) " & vbCrLf & _
	"        ) AS Settlement_Date, " & vbCrLf & _
	"        LTRIM(RTRIM(e.Trace_Nbr)) AS TraceKey, " & vbCrLf & _
	"        e.Post_Date, e.id, " & vbCrLf & _
	"        ABS(TRY_CONVERT(decimal(18,2), REPLACE(REPLACE(LTRIM(RTRIM(e.Amount)), ',', ''), '$', ''))) AS AmtKey, " & vbCrLf & _
	"        ROW_NUMBER() OVER ( " & vbCrLf & _
	"            PARTITION BY LTRIM(RTRIM(e.Trace_Nbr)) " & vbCrLf & _
	"            ORDER BY e.Post_Date DESC, e.id DESC " & vbCrLf & _
	"        ) AS rn " & vbCrLf & _
	"    FROM dbo.ACH_EPSA e " & vbCrLf & _
	"    WHERE e.Trace_Nbr IS NOT NULL " & vbCrLf & _
	"), " & vbCrLf & _
	"epsa_winner AS ( " & vbCrLf & _
	"    SELECT Settlement_Date, TraceKey, AmtKey " & vbCrLf & _
	"    FROM epsa_ranked " & vbCrLf & _
	"    WHERE rn = 1 AND Settlement_Date IS NOT NULL " & vbCrLf & _
	"), " & vbCrLf & _
	"ertr_norm AS ( " & vbCrLf & _
	"    SELECT " & vbCrLf & _
	"        LTRIM(RTRIM(r.Trace_Number)) AS TraceKey, " & vbCrLf & _
	"        ABS(TRY_CONVERT(decimal(18,2), REPLACE(REPLACE(LTRIM(RTRIM(r.RTN_Amount)), ',', ''), '$', ''))) AS AmtKey " & vbCrLf & _
	"    FROM dbo.ACH_ERTR r " & vbCrLf & _
	"    WHERE r.Trace_Number IS NOT NULL " & vbCrLf & _
	") " & vbCrLf & _
	"SELECT " & vbCrLf & _
	"    e.Settlement_Date, " & vbCrLf & _
	"    COUNT(*) AS TotalPosted, " & vbCrLf & _
	"    COUNT(r.TraceKey) AS Returned, " & vbCrLf & _
	"    COUNT(*) - COUNT(r.TraceKey) AS Cleared, " & vbCrLf & _
	"    CAST(CASE WHEN COUNT(*) = 0 THEN 0 " & vbCrLf & _
	"         ELSE (COUNT(r.TraceKey) * 100.0 / COUNT(*)) END AS decimal(9,4)) AS ReturnRatePct " & vbCrLf & _
	"FROM epsa_winner e " & vbCrLf & _
	"LEFT JOIN ertr_norm r " & vbCrLf & _
	"  ON r.TraceKey = e.TraceKey " & vbCrLf & _
	" AND ( (e.AmtKey IS NOT NULL AND r.AmtKey IS NOT NULL AND e.AmtKey = r.AmtKey) " & vbCrLf & _
	"       OR (e.AmtKey IS NULL OR r.AmtKey IS NULL) ) " & vbCrLf & _
	"GROUP BY e.Settlement_Date;"

	conn.Execute sqlViewDDL

	' ------------------------------------------------------------
	' 1) SUMMARY (Monitoring)
	' ------------------------------------------------------------
	Dim sqlSum, cmdSum, rsSum

	sqlSum = _
	"SELECT Settlement_Date, TotalPosted, Returned, Cleared, ReturnRatePct " & _
	"FROM " & viewName & " " & _
	"WHERE Settlement_Date BETWEEN ? AND ? " & _
	"ORDER BY Settlement_Date"

	Set cmdSum = CreateObject("ADODB.Command")
	With cmdSum
		.ActiveConnection = conn
		.CommandText = sqlSum
		.CommandType = 1
		.Parameters.Append .CreateParameter("", 133, 1, , startDateObj)
		.Parameters.Append .CreateParameter("", 133, 1, , endDateObj)
	End With

	Set rsSum = cmdSum.Execute

	fout.WriteLine "REPORT    ACH RECONCILIATION SUMMARY"
	fout.WriteLine Pad("Settlement", 11) & Pad("TotalPosted", 13) & Pad("Returned", 11) & Pad("Cleared", 17) & Pad("ReturnRate%", 13)

	If rsSum.EOF Then
		fout.WriteLine "No reconciliation rows returned for this period."
	Else
		Do While Not rsSum.EOF
			fout.WriteLine _
				Pad(FmtDate8(rsSum("Settlement_Date")), 12) & _
				Pad(CStr(rsSum("TotalPosted")), 12) & _
				Pad(CStr(rsSum("Returned")), 10) & _
				Pad(CStr(rsSum("Cleared")), 16) & _
				Pad(CStr(rsSum("ReturnRatePct")), 12)
			rsSum.MoveNext
		Loop
	End If

	fout.WriteLine ""


	' ------------------------------------------------------------
	' 2A) DETAIL - RETURNED (Investigation)
	' ------------------------------------------------------------
	Dim sqlRet, cmdRet, rsRet

	sqlRet = _
	"WITH epsa_ranked AS ( " & _
	"  SELECT " & _
	"    COALESCE(TRY_CONVERT(date, e.Settlement_Date), TRY_CONVERT(date, e.Settlement_Date, 10), TRY_CONVERT(date, e.Settlement_Date, 1), TRY_CONVERT(date, e.Settlement_Date, 101)) AS Settlement_Date, " & _
	"    LTRIM(RTRIM(e.Trace_Nbr)) AS TraceKey, " & _
	"    e.Post_Date, " & _
	"    e.Post_Time, " & _
	"    e.Company_Name, " & _
	"    e.ACH_Batch_ID, " & _
	"    e.[Status], " & _
	"    e.[Tran_Type], " & _
	"    e.[Tran_Code], " & _
	"    e.[Type], " & _
	"    e.Amount, " & _
	"    e.id, " & _
	"    ROW_NUMBER() OVER (PARTITION BY LTRIM(RTRIM(e.Trace_Nbr)) ORDER BY e.Post_Date DESC, e.id DESC) AS rn " & _
	"  FROM dbo.ACH_EPSA e " & _
	"  WHERE e.Trace_Nbr IS NOT NULL " & _
	"), epsa_winner AS ( " & _
	"  SELECT * FROM epsa_ranked WHERE rn = 1 AND Settlement_Date IS NOT NULL " & _
	"), ertr_norm AS ( " & _
	"  SELECT " & _
	"    LTRIM(RTRIM(r.Trace_Number)) AS TraceKey, " & _
	"    r.Post_Date AS ERTR_Post_Date, " & _
	"    LTRIM(RTRIM(r.RTN_Amount)) AS RTN_Amount, " & _
	"    r.RRC, r.ORC, r.DRC, " & _
	"    r.Company_ID, " & _
	"    r.Company_Name AS ERTR_Company_Name, " & _
	"    r.Individual_Name, " & _
	"    r.ERTR_File, " & _
	"    r.[RDFI_RT#], " & _
	"    r.[SRC_BH#], " & _
	"    r.[RTN_Trace#], " & _
	"    r.[RTN_RDFI#], " & _
	"    r.[DFI_Account] " & _
	"  FROM dbo.ACH_ERTR r " & _
	"  WHERE r.Trace_Number IS NOT NULL " & _
	") " & _
	"SELECT " & _
	"  e.Settlement_Date, " & _
	"  e.TraceKey, " & _
	"  e.Post_Date, " & _
	"  e.Post_Time, " & _
	"  e.Company_Name, " & _
	"  e.ACH_Batch_ID, " & _
	"  e.[Status], " & _
	"  e.[Tran_Type], " & _
	"  e.[Tran_Code], " & _
	"  e.[Type], " & _
	"  LTRIM(RTRIM(e.Amount)) AS EPSA_Amount, " & _
	"  r.ERTR_Post_Date, " & _
	"  r.RTN_Amount AS ERTR_Amount, " & _
	"  r.RRC, r.ORC, r.DRC, r.Company_ID, r.ERTR_Company_Name, r.Individual_Name, r.ERTR_File, " & _
	"  r.[RDFI_RT#], r.[SRC_BH#], r.[RTN_Trace#], r.[RTN_RDFI#], r.[DFI_Account] " & _
	"FROM epsa_winner e " & _
	"INNER JOIN ertr_norm r ON r.TraceKey = e.TraceKey " & _
	"WHERE e.Settlement_Date BETWEEN ? AND ? " & _
	"ORDER BY e.Settlement_Date, e.TraceKey"
	
	Set cmdRet = CreateObject("ADODB.Command")
	With cmdRet
		.ActiveConnection = conn
		.CommandText = sqlRet
		.CommandType = 1
		.Parameters.Append .CreateParameter("", 133, 1, , startDateObj)
		.Parameters.Append .CreateParameter("", 133, 1, , endDateObj)
	End With

	Set rsRet = cmdRet.Execute

'	fout.WriteLine "REPORT    ACH TRANSACTIONS THAT WERE LATER RETURNED"

	fout.WriteLine ""
	fout.WriteLine ""
	fout.WriteLine ""
	fout.WriteLine _
		Pad("ERTR", 10) & _
		Pad("Post_Date", 12) & _
		Pad("Post_Time", 12) & _
		Pad("Settlement_Date", 16) & _
		Pad("ERTR_Company_Name", 25) & _
		Pad("TraceKey", 16) & _
		Pad("", 8) & _
		Pad("", 8) & _
		Pad("Status", 8) & _
		Pad("Tran_Type", 12) & _
		Pad("Tran_Code", 12) & _
		Pad("", 16) & _
		Pad("EPSA_Amount", 12) & _
		Pad("", 15) & _
		Pad("Individual_Name", 25) & _
		Pad("", 10) & _
		Pad("", 25) & _
		Pad("", 4) & _
		Pad("DFI_Account", 20) & _
		Pad("", 30) & _
		Pad("Type", 8) & _
		Pad("", 10) & _
		Pad("", 15) & _
		Pad("", 14) & _
		Pad("", 12) & _
		Pad("", 12) & _
		Pad("", 4) & _
		Pad("RRC", 4) & _
		Pad("RDFI_RT#", 10) & _
		Pad("", 14) & _
		Pad("Company_ID", 14) & _
		Pad("MESSAGE", 35) & _
		Pad("ERTR_File", 14)
		
		
	If rsRet.EOF Then
		fout.WriteLine "ERTR   No rows returned for this period."
	Else
		Do While Not rsRet.EOF
			fout.WriteLine _
				Pad("ERTR", 10) & _
				Pad(FmtDate8(rsRet("Post_Date")), 12) & _
				Pad(FmtTimeShort(rsRet("Post_Time")), 12) & _
				Pad(FmtDate8(rsRet("Settlement_Date")), 16) & _
				Pad(CStr(rsRet("ERTR_Company_Name")), 25) & _
				Pad(CStr(rsRet("TraceKey")), 16) & _
				Pad("", 8) & _
				Pad("", 8) & _
				Pad(CStr(rsRet("Status")), 8) & _
				Pad(CStr(rsRet("Tran_Type")), 12) & _
				Pad(CStr(rsRet("Tran_Code")), 12) & _
				Pad("", 16) & _
				Pad(CStr(rsRet("EPSA_Amount")), 12) & _
				Pad("", 15) & _
				Pad(CStr(rsRet("Individual_Name")), 25) & _
				Pad("", 10) & _
				Pad("", 25) & _
				Pad("", 4) & _
				Pad(CStr(rsRet("DFI_Account")), 20) & _
				Pad("", 30) & _
				Pad(CStr(rsRet("Type")), 8) & _
				Pad("", 10) & _
				Pad("", 15) & _
				Pad("", 14) & _
				Pad("", 12) & _
				Pad("", 12) & _
				Pad("", 4) & _
				Pad(CStr(rsRet("RRC")), 4) & _
				Pad(CStr(rsRet("RDFI_RT#")), 14) & _
				Pad("", 10) & _
				Pad(CStr(rsRet("Company_ID")), 14) & _
				Pad( _
					"SRC_BH#=" & CStr(rsRet("SRC_BH#")) & _
					" RTN_TRACE#=" & CStr(rsRet("RTN_Trace#")) & _
					" RTN_RDFI#=" & CStr(rsRet("RTN_RDFI#")) _
				, 35) & _
				Pad(CStr(rsRet("ERTR_File")), 14)

			rsRet.MoveNext
		Loop
	End If



    CleanupObjects conn, rs, cmd
    CleanupObjects Nothing, rs2, cmd2

    On Error GoTo 0
End Sub

Sub ProcessLine
End Sub

Sub CloseDoc
End Sub

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

    gMode         = GetIniValue(cfg, "Report.Mode", "MONTHLY")
    gRelOffset    = GetIniValue(cfg, "Report.RelativeOffset", "0")
    gYear         = GetIniValue(cfg, "Report.Year", "")
    gMonth        = GetIniValue(cfg, "Report.Month", "")
    gQuarter      = GetIniValue(cfg, "Report.Quarter", "")
    gStartDateStr = GetIniValue(cfg, "Report.StartDate", "")
    gEndDateStr   = GetIniValue(cfg, "Report.EndDate", "")

    gConnStr       = GetIniValue(cfg, "Database.ConnStr", "")
    gFullTableName = GetIniValue(cfg, "Database.TableName", "[dbo].[ACH_EPSA]")

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
                While q > 4: q = q - 4: y = y + 1: Wend
                While q < 1: q = q + 4: y = y - 1: Wend
            End If

            m = (q - 1) * 3 + 1
            dStart = DateSerial(y, m, 1)
            dEnd = DateAdd("d", -1, DateAdd("m", 3, dStart))

        Case "YTD"
            If gYear <> "" And IsNumeric(gYear) Then y = CInt(gYear) Else y = Year(today)
            dStart = DateSerial(y, 1, 1)
            dEnd = today

		Case "MTD"
			Dim baseDate
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


Function Fmt(d, mask)
    ' Tries VBA-style Format(); if not available, falls back to VBScript formatting
    On Error Resume Next

    Dim s
    s = ""

    ' Try VBA/host-provided Format first
    s = Format(d, mask)
    If Err.Number = 0 And Len(Trim(CStr(s))) > 0 Then
        Fmt = CStr(s)
        Exit Function
    End If

    ' Fallback
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
            ' Generic safe default
            Fmt = MonthName(Month(d), True) & " " & Day(d) & ", " & Year(d)
    End Select

    On Error GoTo 0
End Function
