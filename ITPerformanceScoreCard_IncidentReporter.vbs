'==========================================================
' ExcelIncidentConsumer_Ini.vbs
' Data Consumption Script — reads KPIs from IncidentReporter
' (ExcludeVendors fully applied across all sections)
'==========================================================
Option Explicit

' --- Config path ---
Dim INI_PATH: INI_PATH = "S:\Reports\ITScoreCard\ITPerformanceScoreCard.ini"

' --- Globals ---
Dim gCfg, gPeriodStart, gPeriodLabel
Dim mode, rel, yr, mo, qStart, qEnd, anchor

'========================
' byREQUEST entry points
'========================
Sub StartDoc()
    On Error Resume Next
    Fout.WriteLine "IT Performance Scorecard"

    ' --- Load INI configuration ---
    Set gCfg = ReadIniSafe(INI_PATH)
    If gCfg Is Nothing Then
        Fout.WriteLine "ERROR: INI not found or unreadable: " & INI_PATH
        Exit Sub
    End If

    ' --- Determine reporting period label ---
    gPeriodStart = ResolvePeriodStart(gCfg)
    gPeriodLabel = MonthName(Month(gPeriodStart)) & " " & Year(gPeriodStart)
    Fout.WriteLine "Period: " & gPeriodLabel

    ' --- Open DB connection ---
    Dim connStr: connStr = GetIni(gCfg, "Report.Database.ConnStr")
    If Len(connStr) = 0 Then connStr = GetIni(gCfg, "Database.ConnStr")
    If Len(connStr) = 0 Then
        Fout.WriteLine "ERROR: Missing Database.ConnStr in INI"
        Exit Sub
    End If

    Dim cn: Set cn = CreateObject("ADODB.Connection")
    cn.Open connStr
    If Err.Number <> 0 Then
        Fout.WriteLine "ERROR: Cannot open database connection: " & Err.Description
        Exit Sub
    End If

    ' --- Build query range (supports YTD, MONTHLY, CUSTOM) ---
    mode = UCase(NullToEmpty(GetIni(gCfg, "Report.Mode")))
    rel  = 0
    If Len(NullToEmpty(GetIni(gCfg, "Report.RelativeOffset"))) > 0 Then rel = CLng(GetIni(gCfg, "Report.RelativeOffset"))

    yr = NullToEmpty(GetIni(gCfg, "Report.Year"))
    mo = NullToEmpty(GetIni(gCfg, "Report.Month"))

    If mode = "CUSTOM" Then
        Dim sIn, eIn
        sIn = NullToEmpty(GetIni(gCfg, "Report.StartDate"))
        eIn = NullToEmpty(GetIni(gCfg, "Report.EndDate"))
        If Not IsDate(sIn) Or Not IsDate(eIn) Then
            Fout.Writeline "ERROR: CUSTOM mode requires valid StartDate and EndDate."
            cn.Close: Exit Sub
        End If
        qStart = CDate(sIn)
        ' make qEnd exclusive for SQL BETWEEN style
        qEnd   = DateAdd("d", 1, CDate(eIn))
    ElseIf mode = "YTD" Then
        ' Anchor is today shifted by RelativeOffset months, unless a Year is specified
        anchor = DateAdd("m", rel, Date)
        If Len(yr) > 0 Then anchor = DateSerial(CLng(yr), Month(anchor), Day(anchor))
        qStart = DateSerial(Year(anchor), 1, 1)
        ' End at end of the anchor month (exclusive)
        qEnd   = DateSerial(Year(anchor), Month(DateAdd("m", 1, anchor)), 1)
    ElseIf mode = "MONTHLY" Then
        If Len(yr) > 0 And Len(mo) > 0 Then
            anchor = DateSerial(CLng(yr), CLng(mo), 1)
        Else
            anchor = DateAdd("m", rel, DateSerial(Year(Date), Month(Date), 1))
        End If
        qStart = anchor
        qEnd   = DateAdd("m", 1, anchor)
    Else
        ' default: prior month (MONTHLY) if unspecified
        anchor = DateAdd("m", -1, DateSerial(Year(Date), Month(Date), 1))
        qStart = anchor
        qEnd   = DateAdd("m", 1, anchor)
    End If

    ' =======================
    ' Output sections in order
    ' =======================
    DisplaySystemPerformanceKPIs cn, qStart, qEnd

    Fout.WriteLine ""

    DisplayVendorKPIs cn, qStart, qEnd
    'DisplayHelpDeskYTD gPeriodStart
    'DisplayHelpDeskMTD gPeriodStart

    DisplayYTDTrend cn, qStart, qEnd

	DisplayVendorIncidentTracking cn, qStart, qEnd



    SummarizeCalcs
    cn.Close
    Set cn = Nothing
    On Error GoTo 0
End Sub

Sub ProcessLine()
    ' no-op for consumer
End Sub

Sub CloseDoc()
    ' no-op
End Sub

'========================
'  Exclude helpers (NEW)
'========================
Private Function SqlQuote(s)
    SqlQuote = "'" & Replace(s, "'", "''") & "'"
End Function

Private Function ParseVendorExcludes()
    Dim raw: raw = GetIni(gCfg, "Filters.ExcludeVendors")
    Dim s:  s  = Replace(NullToEmpty(raw), ";", ",") ' allow ; or ,
    Dim parts: parts = Split(s, ",")
    Dim i, t, out(): ReDim out(0) : out(0) = ""
    Dim n: n = 0
    For i = 0 To UBound(parts)
        t = Trim(parts(i))
        If Len(t) > 0 Then
            If n = 0 Then
                ReDim out(0)
            Else
                ReDim Preserve out(n)
            End If
            out(n) = t
            n = n + 1
        End If
    Next
    If n = 0 Then
        ParseVendorExcludes = Array() ' empty
    Else
        ParseVendorExcludes = out
    End If
End Function

' Build " AND NOT (...)" for a given column name (e.g., "Vendor" or "v.VendorName")
' Rules:
'   - If token contains '%' or '_' -> treat as SQL LIKE pattern
'   - If token ends with '*'       -> treat as prefix (LIKE 'token%')
'   - Else                         -> default to prefix (LIKE 'token%') for short tokens like "Network"
Private Function BuildVendorExcludePredicate(ByVal colName)
    Dim toks: toks = ParseVendorExcludes()
    If (IsArray(toks) = False) Or (UBound(toks) < 0) Then
        BuildVendorExcludePredicate = ""
        Exit Function
    End If

    Dim i, t, clauses(), c: c = -1
    For i = 0 To UBound(toks)
        t = toks(i)
        If InStr(t, "%") > 0 Or InStr(t, "_") > 0 Then
            c = c + 1 : ReDim Preserve clauses(c)
            clauses(c) = colName & " LIKE " & SqlQuote(t)
        ElseIf Right(t,1) = "*" Then
            c = c + 1 : ReDim Preserve clauses(c)
            clauses(c) = colName & " LIKE " & SqlQuote(Left(t, Len(t)-1) & "%")
        Else
            c = c + 1 : ReDim Preserve clauses(c)
            clauses(c) = colName & " LIKE " & SqlQuote(t & "%")
            ' For exact-only behavior, replace the previous line with:
            'clauses(c) = "LTRIM(RTRIM(" & colName & ")) = " & SqlQuote(t)
        End If
    Next

    If c >= 0 Then
        BuildVendorExcludePredicate = " AND NOT (" & Join(clauses, " OR ") & ")"
    Else
        BuildVendorExcludePredicate = ""
    End If
End Function

' Defensive skip when iterating vendor lists
Private Function ShouldExcludeVendor(ByVal vendorName)
    Dim toks: toks = ParseVendorExcludes()
    Dim i, t
    ShouldExcludeVendor = False
    If (IsArray(toks) = False) Or (UBound(toks) < 0) Then Exit Function
    For i = 0 To UBound(toks)
        t = toks(i)
        If InStr(t, "%") > 0 Or InStr(t, "_") > 0 Then
            If LCase(vendorName) Like LCase(Replace(t, "%", "*")) Then
                ShouldExcludeVendor = True: Exit Function
            End If
        ElseIf Right(t,1) = "*" Then
            If StrComp(Left(vendorName, Len(t)-1), Left(t, Len(t)-1), vbTextCompare) = 0 Then
                ShouldExcludeVendor = True: Exit Function
            End If
        Else
            If StrComp(Left(vendorName, Len(t)), t, vbTextCompare) = 0 Then
                ShouldExcludeVendor = True: Exit Function
            End If
        End If
    Next
End Function

'========================
'  Sections
'========================

' Display only Core System Availability (with fixed 0.000% current/YTD and proper variance formatting)
Sub DisplaySystemPerformanceKPIs(cn, qStart, qEnd)
    ' ---- Header line back in ----
    Fout.Writeline ""
    Fout.Writeline "CORE SYSTEM PERFORMANCE KPIs"

    ' Goal from INI (default 99.995)
    Dim goalAvailability: goalAvailability = GetIni(gCfg, "Goals.CoreSystemAvailability")
    If Len(goalAvailability) = 0 Then goalAvailability = "99.995"

    ' Compute real period & YTD (weighted across vendors)
    Dim periodPct, ytdPct
    periodPct = 0: ytdPct = 0
    ComputeCoreAvailability cn, qStart, qEnd, periodPct, ytdPct

    ' Format fields
    Dim goalStr:    goalStr    = FormatNumber(CDbl(goalAvailability), 3) & "%"
    Dim curPeriod:  curPeriod  = FormatNumber(periodPct, 3) & "%"
    Dim curYTD:     curYTD     = FormatNumber(ytdPct, 3) & "%"
    Dim varPeriod:  varPeriod  = periodPct - CDbl(goalAvailability)
    Dim varYTD:     varYTD     = ytdPct    - CDbl(goalAvailability)
    Dim varPeriodStr: varPeriodStr = FormatVariance3(varPeriod)
    Dim varYTDStr:    varYTDStr    = FormatVariance3(varYTD)

    ' Table header + one KPI row
    Fout.Writeline PadRight("KPI Description", 30) & _
                   PadRight("Goal", 15) & _
                   PadRight("Current Period", 15) & _
                   PadRight("Goal Variance", 15) & _
                   PadRight("Current YTD", 15) & _
                   PadRight("Goal Variance", 15)

    Fout.Writeline PadRight("Core System Availability", 30) & _
                   PadRight(goalStr, 15) & _
                   PadRight(curPeriod, 15) & _
                   PadRight(varPeriodStr, 15) & _
                   PadRight(curYTD, 15) & _
                   PadRight(varYTDStr, 15)
End Sub

' Wrapper (availability now; extend later if you re-add incident KPIs)
Sub DisplayVendorKPIs(cn, qStart, qEnd)
    DisplayVendorAvailabilityKPIs cn, qStart, qEnd
End Sub

Sub DisplayVendorAvailabilityKPIs(cn, qStart, qEnd)
    On Error Resume Next

    Dim sql, cmd, rs
    Dim targetAvailability
    targetAvailability = GetIni(gCfg, "VendorAvailabilityGoals.TargetAvailability")
    If Len(targetAvailability) = 0 Then targetAvailability = 99.995


	Dim yellowAvailability
		yellowAvailability = GetIni(gCfg, "VendorAvailabilityGoals.YellowAvailability")

		If Len(yellowAvailability) = 0 Then
			yellowAvailability = 99
	End If



    Fout.WriteLine ""
    Dim spanLabel
    spanLabel = MonthName(Month(qStart)) & " " & Year(qStart)
    If DateDiff("m", qStart, qEnd) > 1 Then
        ' Show a compact label for multi-month spans (e.g., YTD)
        Dim thru: thru = DateAdd("d", -1, qEnd)
        spanLabel = "YTD through " & MonthName(Month(thru)) & " " & Year(thru)
    End If
    Fout.WriteLine "VENDOR AVAILABILITY KPIs"

    ' First get ALL monitored vendors from YTD data (CURRENT YEAR ONLY)
    Dim allVendorsDict
    Set allVendorsDict = CreateObject("Scripting.Dictionary")

    Dim ytdStart: ytdStart = DateSerial(Year(qStart), 1, 1)
    Dim ytdEnd: ytdEnd = qEnd
    Dim currentYear: currentYear = Year(qStart)

    sql = "SELECT DISTINCT Vendor FROM vw_VendorAvailability_CardWizard " & _
          "WHERE ReportDate >= ? AND ReportDate < ? " & _
          "AND YEAR(ReportDate) = ? " & _                       
          BuildVendorExcludePredicate("Vendor") & _
          " ORDER BY Vendor"

    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = 1
    cmd.CommandText = sql
    cmd.Parameters.Append cmd.CreateParameter("@YTDStart", 133, 1, , ytdStart)
    cmd.Parameters.Append cmd.CreateParameter("@YTDEnd", 133, 1, , ytdEnd)
    cmd.Parameters.Append cmd.CreateParameter("@Year", 3, 1, , currentYear)  ' Add year filter

    Set rs = cmd.Execute
    Do Until rs.EOF
        Dim vendorKey
        vendorKey = NzStr(rs("Vendor"))
        If Len(vendorKey) > 0 Then
            allVendorsDict(vendorKey) = True
        End If
        rs.MoveNext
    Loop
    rs.Close

    ' Now get current period data for vendors
    sql = "SELECT " & _
          "  Vendor, " & _
          "  SUM(TotalDowntimeMinutes)        AS TotalDownMin, " & _
          "  SUM(TotalMonitoringMinutes)      AS TotalMonMin, " & _
          "  SUM(COALESCE(IncidentCount, 0))  AS IncidentCount " & _
          "FROM vw_VendorAvailability_CardWizard " & _
          "WHERE ReportDate >= ? AND ReportDate < ? " & _
          BuildVendorExcludePredicate("Vendor") & _
          " GROUP BY Vendor"

    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = 1
    cmd.CommandText = sql
    cmd.Parameters.Append cmd.CreateParameter("@Start", 133, 1, , qStart)
    cmd.Parameters.Append cmd.CreateParameter("@End", 133, 1, , qEnd)

    Set rs = cmd.Execute

    If Err.Number <> 0 Then
        Fout.WriteLine "ERROR: Query failed - " & Err.Description
        Exit Sub
    End If

    If rs Is Nothing Or rs.EOF Then
        Fout.WriteLine "No vendor availability data found for the period."
        If Not rs Is Nothing Then rs.Close
        Set cmd = Nothing
        Exit Sub
    End If

    Fout.WriteLine PadRight("Vendor", 30) & PadRight("Availability", 15) & _
                   PadRight("Goal", 15) & PadRight("Variance", 15) & _
                   PadRight("Downtime", 15) & PadRight("Incidents", 15) & PadRight("Status", 15)

    Dim totalDownMin, totalMonMin, totalIncidents
    Dim vendorCount, vendorsAtYellow, vendorsMeetingGoal, perfectAvailabilityVendors
    Dim maxDowntimeHours, criticalIncidents, highIncidents
    
    totalDownMin = 0: totalMonMin = 0: totalIncidents = 0
    maxDowntimeHours = 0
    criticalIncidents = 0
    highIncidents = 0
	vendorsAtYellow = 0

    vendorCount = allVendorsDict.Count  ' Total vendors being monitored (excludes filtered vendors)

	vendorsMeetingGoal = 0
	vendorsAtYellow = 0
    perfectAvailabilityVendors = vendorCount ' Start with all perfect, subtract those with incidents
    
    Do Until rs.EOF
        Dim vendorName, dmin, mmin, inc, availability, variance, status
        Dim downtimeHoursDisplay, downtimeHoursRaw

        vendorName = NzStr(rs("Vendor")) : If Len(vendorName) = 0 Then vendorName = "Internal"
        dmin = NzD(rs("TotalDownMin"))
        mmin = NzD(rs("TotalMonMin"))
        inc  = CLng(NzD(rs("IncidentCount")))

        If mmin > 0 Then
            availability = (1 - (dmin / mmin)) * 100
        Else
            availability = 100
        End If

        variance = availability - CDbl(targetAvailability)
        


		' after you compute availability and status for THIS vendor (single row)
		status = GetVendorStatus3(availability, targetAvailability, yellowAvailability)

		Select Case status
			Case "PASS"
				vendorsMeetingGoal = vendorsMeetingGoal + 1
				vendorsAtYellow    = vendorsAtYellow + 1
			Case "MONITOR"
				vendorsAtYellow    = vendorsAtYellow + 1
			Case "FAIL"
				' nothing
		End Select
								
        ' Count perfect availability (100% exactly)
        If availability = 100 Then
            ' Already counted in perfectAvailabilityVendors (starts with all)
        Else
            perfectAvailabilityVendors = perfectAvailabilityVendors - 1  ' Subtract non-perfect vendors
        End If

        ' Calculate hours for display AND tracking
        downtimeHoursRaw = dmin / 60
        downtimeHoursDisplay = FormatNumber(Round(downtimeHoursRaw, 2), 2, -1, 0, 0) & "h"

        ' Track max downtime using the raw value
        If downtimeHoursRaw > maxDowntimeHours Then maxDowntimeHours = downtimeHoursRaw

        ' Track severity distribution
        If downtimeHoursRaw >= 10 Then
            criticalIncidents = criticalIncidents + inc
        ElseIf downtimeHoursRaw >= 2 Then
            highIncidents = highIncidents + inc
        End If

        Fout.WriteLine PadRight(vendorName, 30) & _
                       PadRight(FormatNumber(availability, 3) & "%", 15) & _
                       PadRight(CStr(targetAvailability) & "%", 15) & _
                       PadRight(FormatVariance3(variance), 15) & _
                       PadRight(downtimeHoursDisplay, 15) & _
                       PadRight(CStr(inc), 15) & _
                       PadRight(status, 15)

        totalDownMin = totalDownMin + dmin
        totalMonMin  = totalMonMin  + mmin
        totalIncidents = totalIncidents + inc

        rs.MoveNext
    Loop

    Dim overallAvailability
    If totalMonMin > 0 Then
        overallAvailability = (1 - (totalDownMin / totalMonMin)) * 100
    Else
        overallAvailability = 100
    End If

    ' Get YTD incidents count
    Dim totalIncidentsYTD
    totalIncidentsYTD = 0

    sql = "SELECT SUM(COALESCE(IncidentCount, 0)) AS TotalIncidentsYTD " & _
          "FROM vw_VendorAvailability_CardWizard " & _
          "WHERE ReportDate >= ? AND ReportDate < ? " & _
          "AND YEAR(ReportDate) = ? " & _
          BuildVendorExcludePredicate("Vendor")

    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = 1
    cmd.CommandText = sql
    cmd.Parameters.Append cmd.CreateParameter("@YTDStart", 133, 1, , ytdStart)
    cmd.Parameters.Append cmd.CreateParameter("@YTDEnd", 133, 1, , ytdEnd)
    cmd.Parameters.Append cmd.CreateParameter("@Year", 3, 1, , currentYear)

    Set rs = cmd.Execute
    If Not rs.EOF Then
        totalIncidentsYTD = CLng(NzD(rs("TotalIncidentsYTD")))
    End If
    rs.Close

    ' Calculate MoM change (placeholder)
    Dim momChangePct, currentMonth
    currentMonth = Month(DateAdd("d", -1, qEnd)) ' Last month in the range
    momChangePct = 0
    
Dim helpYTDLines, helpMTDLines


Dim reportMonthDate
reportMonthDate = DateAdd("d", -1, qEnd)

helpYTDLines = GetHelpDeskYTDLines(reportMonthDate)
helpMTDLines = GetHelpDeskMTDLines(reportMonthDate)


Dim leftWidth, mtdWidth
leftWidth = 75
mtdWidth = 44

	Fout.WriteLine ""
	Fout.WriteLine PadRight("OVERALL PERFORMANCE SUMMARY", leftWidth) & _
               PadRight("HELP DESK (MTD)", mtdWidth) & _
               "HELP DESK (YTD)"

	Fout.WriteLine PadRight( _
				PadRight("Vendors Meeting Goal", 30) & _
				PadRight(vendorsMeetingGoal & "/" & vendorCount, 15) & _
				PadRight(FormatNumber((vendorsMeetingGoal / vendorCount) * 100, 1) & "%", 15), _
				leftWidth) & _
				PadRight(helpMTDLines(0), mtdWidth) & _
				helpYTDLines(0)

	Fout.WriteLine PadRight( _
				PadRight("Vendors at/above Yellow", 30) & _
				PadRight(vendorsAtYellow & "/" & vendorCount, 15) & _
				PadRight(FormatNumber((vendorsAtYellow / vendorCount) * 100, 1) & "%", 15), _
				leftWidth) & _
				PadRight(helpMTDLines(1), mtdWidth) & _
				helpYTDLines(1)

	Fout.WriteLine PadRight( _
				PadRight("Perfect Availability Vendors", 30) & _
				PadRight(perfectAvailabilityVendors & "/" & vendorCount, 15) & _
				PadRight(FormatNumber((perfectAvailabilityVendors / vendorCount) * 100, 1) & "%", 15), _
				leftWidth) & _
				PadRight(helpMTDLines(2), mtdWidth) & _
				helpYTDLines(2)

	Fout.WriteLine PadRight( _
				PadRight("Max Single Vendor Downtime", 30) & _
				FormatNumber(maxDowntimeHours, 2) & " hours", _
				leftWidth) & _
				PadRight(helpMTDLines(3), mtdWidth) & _
				helpYTDLines(3)

	Fout.WriteLine PadRight( _
				PadRight("Severity Distribution", 30) & _
				PadRight("Critical: " & criticalIncidents, 15) & _
				PadRight("High: " & highIncidents, 15), _
				leftWidth)

	Fout.WriteLine PadRight( _
				PadRight("Total Downtime - Current", 30) & _
				FormatNumber(totalDownMin / 60, 2) & " hours", _
				leftWidth)

	Fout.WriteLine PadRight( _
				PadRight("Total Incidents -Current", 30) & _
				totalIncidents, _
				leftWidth)
    
    DisplayIncidentSummary cn, ytdStart, qEnd, currentMonth

    Set cmd = Nothing
    On Error GoTo 0
End Sub

'========================
'  Helpers
'========================
Function PadRight(text, length)
    If Len(text) >= length Then
        PadRight = Left(text, length)
    Else
        PadRight = text & Space(length - Len(text))
    End If
End Function

Function NzD(v)
    If IsNull(v) Or IsEmpty(v) Then
        NzD = 0
    Else
        On Error Resume Next
        NzD = CDbl(v)
        If Err.Number <> 0 Then NzD = 0 : Err.Clear
        On Error GoTo 0
    End If
End Function

Function NzStr(v)
    If IsNull(v) Or IsEmpty(v) Then
        NzStr = ""
    Else
        NzStr = CStr(v)
    End If
End Function

Function NullToEmpty(v)
    If IsNull(v) Or IsEmpty(v) Then NullToEmpty = "" Else NullToEmpty = CStr(v)
End Function

Function FormatVariance3(v)
    ' 3-decimals, negatives in parentheses, % inside the parens, positives with leading +
    If v < 0 Then
        FormatVariance3 = "(" & Replace(FormatNumber(Abs(v), 3, -1, -1, -1), "-", "") & "%)"
    Else
        FormatVariance3 = "+" & FormatNumber(v, 3, -1, -1, -1) & "%"
    End If
End Function

' Resolve month start using INI; honors Report.RelativeOffset if Year/Month missing
Function ResolvePeriodStart(cfg)
    Dim yr, mo, rel
    yr = NullToEmpty(GetIni(cfg, "Report.Year"))
    mo = NullToEmpty(GetIni(cfg, "Report.Month"))
    rel = 0
    On Error Resume Next
    If Len(NullToEmpty(GetIni(cfg, "Report.RelativeOffset"))) > 0 Then _
        rel = CLng(GetIni(cfg, "Report.RelativeOffset"))
    On Error GoTo 0

    If Len(yr) = 0 Or Len(mo) = 0 Then
        Dim dt: dt = DateAdd("m", rel, Date)
        ResolvePeriodStart = DateSerial(Year(dt), Month(dt), 1)
    Else
        ResolvePeriodStart = DateSerial(CLng(yr), CLng(mo), 1)
    End If
End Function

' Read INI into nested dictionary: cfg(section)(key) = value
Function ReadIniSafe(path)
    On Error Resume Next
    Dim fso, ts, d, sect, line, p, sec
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then Exit Function

    Set ts = fso.OpenTextFile(path, 1)
    Set d  = CreateObject("Scripting.Dictionary")

    Do Until ts.AtEndOfStream
        line = Trim(ts.ReadLine)
        If Len(line) = 0 Or Left(line,1) = ";" Then
            ' skip
        ElseIf Left(line,1) = "[" And Right(line,1) = "]" Then
            sect = Mid(line,2,Len(line)-2)
            If Not d.Exists(sect) Then d.Add sect, CreateObject("Scripting.Dictionary")
        Else
            p = InStr(line, "=")
            If p > 0 And Len(sect) > 0 Then
                Set sec = d(sect)
                sec(Trim(Left(line, p - 1))) = Trim(Mid(line, p + 1))
            End If
        End If
    Loop

    ts.Close
    Set ReadIniSafe = d
    On Error GoTo 0
End Function

' Get INI value with dotted path "Section.Key"
Function GetIni(cfg, keyPath)
    On Error Resume Next
    Dim parts, section, key
    parts = Split(keyPath, ".")
    If UBound(parts) = 1 Then
        section = parts(0): key = parts(1)
        If Not cfg Is Nothing And cfg.Exists(section) Then
            Dim sec: Set sec = cfg(section)
            If sec.Exists(key) Then GetIni = sec(key) Else GetIni = ""
        Else
            GetIni = ""
        End If
    Else
        ' fallback if stored flat
        If Not cfg Is Nothing And cfg.Exists(keyPath) Then GetIni = cfg(keyPath) Else GetIni = ""
    End If
    On Error GoTo 0
End Function

' Returns two Doubles: periodPct, ytdPct (both 0..100)
Sub ComputeCoreAvailability(cn, qStart, qEnd, ByRef periodPct, ByRef ytdPct)
    On Error Resume Next
    Dim cmd, rs, sql

    ' --- Current Period (single month at qStart) ---
    sql = "SELECT SUM(TotalDowntimeMinutes) AS dmin, SUM(TotalMonitoringMinutes) AS mmin " & _
          "FROM vw_VendorAvailability_CardWizard " & _
          "WHERE ReportDate = ? " & _
          BuildVendorExcludePredicate("Vendor")

    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = 1
    cmd.CommandText = sql
    cmd.Parameters.Append cmd.CreateParameter("@ReportDate", 133, 1, , qStart)

    Set rs = cmd.Execute
    If Not rs Is Nothing And Not rs.EOF Then
        Dim dmin, mmin
        dmin = NzD(rs("dmin")) : mmin = NzD(rs("mmin"))
        If mmin > 0 Then
            periodPct = (1 - (dmin / mmin)) * 100
        Else
            periodPct = 100
        End If
    Else
        periodPct = 0
    End If
    If Not rs Is Nothing Then rs.Close
    Set cmd = Nothing

    ' --- YTD (Jan 1 -> qEnd exclusive) ---
    Dim ytdStart: ytdStart = DateSerial(Year(qStart), 1, 1)

    sql = "SELECT SUM(TotalDowntimeMinutes) AS dmin, SUM(TotalMonitoringMinutes) AS mmin " & _
          "FROM vw_VendorAvailability_CardWizard " & _
          "WHERE ReportDate >= ? AND ReportDate < ? " & _
          BuildVendorExcludePredicate("Vendor")

    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = 1
    cmd.CommandText = sql
    cmd.Parameters.Append cmd.CreateParameter("@YTDStart", 133, 1, , ytdStart)
    cmd.Parameters.Append cmd.CreateParameter("@YTDEnd",   133, 1, , qEnd)

    Set rs = cmd.Execute
    If Not rs Is Nothing And Not rs.EOF Then
        Dim dminY, mminY
        dminY = NzD(rs("dmin")) : mminY = NzD(rs("mmin"))
        If mminY > 0 Then
            ytdPct = (1 - (dminY / mminY)) * 100
        Else
            ytdPct = 100
        End If
    Else
        ytdPct = 0
    End If
    If Not rs Is Nothing Then rs.Close
    Set cmd = Nothing
    On Error GoTo 0
End Sub

Sub SummarizeCalcs()
    Fout.WriteLine ""
    Fout.WriteLine Space(280) & "Calculation Summary"

    ' Data scope
    Fout.WriteLine Space(280) & "MonitoringMinutes:"
    Fout.WriteLine Space(280) & "--CardWizard: 9h/day x business days (US federal holidays excluded)."
    Fout.WriteLine Space(280) & "--Others: daysInMonth x 1,440."
	Fout.WriteLine ""
    Fout.WriteLine Space(280) & "Availability% = (1 - DowntimeMinutes / MonitoringMinutes) x 100."
	Fout.WriteLine ""
    Fout.WriteLine Space(280) & "PASS / MONITOR / FAIL vs availability thresholds (Target %, Yellow %)."
	Fout.WriteLine Space(280) & "PASS is at or above Target, MONITOR is below Target but at or above Yellow, FAIL is below Yellow."
	Fout.WriteLine Space(280) & "If TargetAvailability = YellowAvailability then PASS/FAIL only."
	

    ' Core system availability (weighted)
    Fout.WriteLine Space(280) & "Current Period = (1 - SUM(DowntimeMinutes_all) / SUM(MonitoringMinutes_all)) x 100 for month at Start."
	Fout.WriteLine Space(280) & "YTD availability is calculated from total downtime and total monitored minutes from Jan 1 through periodEnd."
	Fout.WriteLine Space(280) & "YTD availability = (1 - Total Downtime Minutes / Total Monitored Minutes) x 100 for the period Jan 1 through periodEnd."
    Fout.WriteLine Space(280) & "YTD = (1 - SUM(DT_YTD) / SUM(Mon_YTD)) x 100 for [Jan 1, periodEnd)."
    Fout.WriteLine Space(280) & "Variance: +x.xxx% or (x.xxx%)."
    Fout.WriteLine ""

    ' -------------------------------
    ' INCIDENT SEVERITY (HOW & CHANGE)
    ' -------------------------------
    Fout.WriteLine Space(280) & "Incident Severity (Critical / High / Low):"
    Fout.WriteLine Space(280) & "-- Severity is assigned PER VENDOR PER MONTH based on that month's TOTAL downtime minutes."
    Fout.WriteLine Space(280) & "-- All incidents in the month fall into ONE bucket using downtime thresholds:"
    Fout.WriteLine Space(280) & "   * Critical:  DowntimeMinutes >= CriticalThreshold"
    Fout.WriteLine Space(280) & "   * High:      DowntimeMinutes >= HighThreshold and < CriticalThreshold"
    Fout.WriteLine Space(280) & "   * Low:       DowntimeMinutes > 0 and < HighThreshold"
    Fout.WriteLine Space(280) & ""
    Fout.WriteLine Space(280) & "How to change thresholds:"
    Fout.WriteLine Space(280) & "-- Edit INI [Filters].SeverityThresholds as: CriticalMinutes,HighMinutes (both in minutes)."
    Fout.WriteLine Space(280) & "-- Example:  SeverityThresholds=600,120   (Critical=600 min, High=120 min; Low is >0 and <120)."
    Fout.WriteLine Space(280) & "-- If omitted/invalid, defaults are Critical=600, High=120."
    Fout.WriteLine Space(280) & "-- This report reads those values at runtime (GetSeverityThresholds)."

    Fout.WriteLine ""

    ' -------------------------------
    ' EXCLUDING CERTAIN VENDORS
    ' -------------------------------
    Fout.WriteLine Space(280) & "Exclude Vendors from all KPIs and counts:"
    Fout.WriteLine Space(280) & "-- Edit INI [Filters].ExcludeVendors with a comma seperated list of vendor names."
    Fout.WriteLine Space(280) & "-- Example:  ExcludeVendors=Facility Power, Network"
    Fout.WriteLine Space(280) & "-- Matching is exact (case-insensitive)."
    Fout.WriteLine Space(280) & "-- Excluded vendors are omitted from availability, incident totals, and all severity buckets."

    Fout.WriteLine ""
    Fout.WriteLine Space(280) & "Why RED status with MoM = +0.000%:"
    Fout.WriteLine Space(280) & "-- GREEN requires: (1) YTD >= Target AND (2) MoM >= 0"
    Fout.WriteLine Space(280) & "-- MoM only compares recent two months"
    Fout.WriteLine Space(280) & "-- Status uses YTD average (all months since January)"
    Fout.WriteLine Space(280) & "-- Early outages keep YTD below target until diluted by clean months"
    Fout.WriteLine ""
    Fout.WriteLine Space(280) & "Detailed Example: January outage affecting YTD through July"
    Fout.WriteLine Space(280) & "-- January: 60 minutes downtime"
    Fout.WriteLine Space(280) & "-- Monthly monitoring: 9h/day x 22 days = 11,880 minutes"
    Fout.WriteLine Space(280) & "-- Total monitoring minutes (Jan-Jul): 7 months x 11,880 = 83,160 minutes"
    Fout.WriteLine Space(280) & "-- YTD Availability Calculation:"
    Fout.WriteLine Space(280) & "---- Available minutes = Total minutes - Downtime"
    Fout.WriteLine Space(280) & "---- Available minutes = 83,160 - 60 = 83,100 minutes"
    Fout.WriteLine Space(280) & "---- YTD % = (83,100 / 83,160) x 100 = 99.928%"
    Fout.WriteLine Space(280) & "-- Result: YTD (99.928%) < Target (99.995%) --> Status = RED"
    Fout.WriteLine ""
    Fout.WriteLine Space(280) & "  Meanwhile, if recent months are clean:"
    Fout.WriteLine Space(280) & "---- June = 100.000%, July = 100.000%"
    Fout.WriteLine Space(280) & "---- MoM = 100.000% - 100.000% = +0.000% (no change)"
    Fout.WriteLine ""
    Fout.WriteLine Space(280) & "  Bottom line: Early outages have long-lasting impact on YTD"
    Fout.WriteLine Space(280) & "  despite good recent performance (neutral MoM)"
End Sub


' Enhanced YTD summary with monthly breakdown
Sub DisplayYTDTrend(cn, qStart, qEnd)
    On Error Resume Next

    Dim targetAvailability
    targetAvailability = GetIni(gCfg, "VendorAvailabilityGoals.TargetAvailability")
    If Len(targetAvailability) = 0 Then targetAvailability = 99.995

    ' Determine YTD window
    Dim anchorYear, ytdStart, ytdEnd, currentMonth
    anchorYear = Year(qStart)
    ytdStart = DateSerial(anchorYear, 1, 1)
    ytdEnd = qEnd
    currentMonth = Month(DateAdd("d", -1, qEnd)) ' Last month in the range

    Fout.WriteLine ""
    Fout.WriteLine "YEAR-TO-DATE SUMMARY"
    
    ' Build dynamic header with only YTD months
    Dim header, monthHeaders, i, monthDate
    header = PadRight("Vendor", 30)
    
    ' Generate month headers only for YTD months
    For i = 1 to currentMonth
        monthDate = DateSerial(anchorYear, i, 1)
        header = header & PadRight(Left(MonthName(i), 3), 15)
    Next
    header = header & PadRight("YTD Total", 15) & PadRight("Status", 15)
    
    Fout.WriteLine header

    ' Get all vendors first (exclude applied)
    Dim vendorsDict
    Set vendorsDict = CreateObject("Scripting.Dictionary")
    
    Dim sql, cmd, rs
    sql = "SELECT DISTINCT Vendor FROM vw_VendorAvailability_CardWizard " & _
          "WHERE ReportDate >= ? AND ReportDate < ? " & _
          BuildVendorExcludePredicate("Vendor") & _
          " ORDER BY Vendor"
    
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = 1
    cmd.CommandText = sql
    cmd.Parameters.Append cmd.CreateParameter("@Start", 133, 1, , ytdStart)
    cmd.Parameters.Append cmd.CreateParameter("@End", 133, 1, , ytdEnd)
    
    Set rs = cmd.Execute
    Do Until rs.EOF
        Dim vendorKey
        vendorKey = NzStr(rs("Vendor"))
        If Len(vendorKey) > 0 Then
            vendorsDict.Add vendorKey, vendorKey
        End If
        rs.MoveNext
    Loop
    rs.Close

    ' Get monthly data for each vendor
    Dim vendorArr
    vendorArr = vendorsDict.Keys
    
    For Each vendorKey In vendorArr
        If Not ShouldExcludeVendor(vendorKey) Then
            PrintVendorMonthlyYTD cn, vendorKey, ytdStart, ytdEnd, currentMonth, targetAvailability
        End If
    Next

    Set cmd = Nothing
    On Error GoTo 0
End Sub

Private Sub PrintVendorMonthlyYTD(cn, vendorName, ytdStart, ytdEnd, currentMonth, targetAvailability)
    Dim sql, cmd, rs
    Dim monthlyData(12) ' Array to store monthly availability
    Dim i, totalDmin, totalMmin
    
    ' Initialize arrays
    For i = 1 To 12
        monthlyData(i) = 100 ' Default to 100% if no data
    Next
    totalDmin = 0
    totalMmin = 0

    ' Get monthly breakdown
    sql = "SELECT MONTH(ReportDate) as MonthNum, " & _
          "       SUM(TotalDowntimeMinutes) AS dmin, " & _
          "       SUM(TotalMonitoringMinutes) AS mmin " & _
          "FROM vw_VendorAvailability_CardWizard " & _
          "WHERE Vendor = ? AND ReportDate >= ? AND ReportDate < ? " & _
          "GROUP BY MONTH(ReportDate) " & _
          "ORDER BY MONTH(ReportDate)"
    
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = 1
    cmd.CommandText = sql
    cmd.Parameters.Append cmd.CreateParameter("@Vendor", 200, 1, 100, vendorName)
    cmd.Parameters.Append cmd.CreateParameter("@Start", 133, 1, , ytdStart)
    cmd.Parameters.Append cmd.CreateParameter("@End", 133, 1, , ytdEnd)
    
    Set rs = cmd.Execute
    
    Do Until rs.EOF
        Dim monthNum, dmin, mmin, availability
        monthNum = CLng(rs("MonthNum"))
        dmin = NzD(rs("dmin"))
        mmin = NzD(rs("mmin"))
        
        If mmin > 0 Then
            availability = (1 - (dmin / mmin)) * 100
        Else
            availability = 100
        End If
        
        monthlyData(monthNum) = availability
        totalDmin = totalDmin + dmin
        totalMmin = totalMmin + mmin
        
        rs.MoveNext
    Loop
    rs.Close

    ' Calculate YTD total
    Dim ytdAvailability, status
    If totalMmin > 0 Then
        ytdAvailability = (1 - (totalDmin / totalMmin)) * 100
    Else
        ytdAvailability = 100
    End If
    
    ' ' Determine status
    ' If ytdAvailability >= CDbl(targetAvailability) Then
        ' status = "PASS"
    ' Else
        ' status = "FAIL"
    ' End If
	
	
	Dim yellowAvailability
		yellowAvailability = GetIni(gCfg, "VendorAvailabilityGoals.YellowAvailability")

	If Len(yellowAvailability) = 0 Then
			yellowAvailability = 99
	End If
	
	status = GetVendorStatus3(ytdAvailability, targetAvailability, yellowAvailability)

		
	

    ' Print vendor row
    Dim output
    output = PadRight(vendorName, 30)
    
    ' Only show months that have occurred
    Dim i2
    For i2 = 1 To currentMonth
        output = output & PadRight(FormatNumber(monthlyData(i2), 3) & "%", 15)
    Next
    
    output = output & PadRight(FormatNumber(ytdAvailability, 3) & "%", 15) & _
             PadRight(status, 15)
    
    Fout.WriteLine output
    
    Set cmd = Nothing
End Sub

Private Sub PrintVendorYTD(vendorName, dsum, msum, lastAvail, prevAvail, targetAvailability)
    Dim ytdAvail, trend, status, ytdStr, trendStr

    If msum > 0 Then
        ytdAvail = (1 - (dsum / msum)) * 100
    Else
        ytdAvail = 100
    End If

    ' MoM trend: last month minus previous month (0 if only one month)
    trend = lastAvail - prevAvail
    If prevAvail = 0 Then trend = 0

    If (ytdAvail >= CDbl(targetAvailability)) And (trend >= 0) Then
        status = "GREEN"
    Else
        status = "RED"
    End If

    ytdStr   = FormatNumber(ytdAvail, 3, -1, -1, -1) & "%"
    If trend < 0 Then
        trendStr = "(" & FormatNumber(Abs(trend), 3, -1, -1, -1) & "%)"
    Else
        trendStr = "+" & FormatNumber(trend, 3, -1, -1, -1) & "%"
    End If

    Fout.WriteLine PadRight(vendorName, 30) & _
                   PadRight(ytdStr, 15) & _
                   PadRight(trendStr, 15) & _
                   PadRight(status, 15)
End Sub

Function MonthEnd(dt)
    MonthEnd = DateSerial(Year(dt), Month(DateAdd("m", 1, dt)), 0)
End Function

Sub DisplayVendorIncidentTracking(cn, qStart, qEnd)
    On Error Resume Next
    
    Dim ytdStart, currentYear, currentMonth
    ytdStart = DateSerial(Year(qStart), 1, 1)
    currentYear = Year(qStart)
    currentMonth = Month(DateAdd("d", -1, qEnd)) ' Last month with data
    
    Fout.WriteLine ""
    Fout.WriteLine "VENDOR INCIDENT TRACKING"
    
    ' Build dynamic header
    Dim header, i
    header = PadRight("Vendor", 30)
    
    ' Add month columns (Jan through current month)
    For i = 1 To currentMonth
        header = header & PadRight(Left(MonthName(i), 3), 15)
    Next
    
    ' Add summary columns
    header = header & PadRight("Total", 15) & _
                   PadRight("Critical", 15) & _
                   PadRight("High", 15) & _
                   PadRight("Low", 15) & _
                   PadRight("Last", 15) & _
                   PadRight("Trend", 15)
    
    Fout.WriteLine header

    ' Get all vendors with incidents in YTD (exclude applied)
    Dim vendorsDict
    Set vendorsDict = CreateObject("Scripting.Dictionary")
    
    Dim sql, cmd, rs
    sql = "SELECT DISTINCT Vendor " & _
          "FROM vw_VendorAvailability_CardWizard " & _
          "WHERE ReportDate >= ? AND ReportDate < ? " & _
          BuildVendorExcludePredicate("Vendor") & _
          " ORDER BY Vendor"
    
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = 1
    cmd.CommandText = sql
    cmd.Parameters.Append cmd.CreateParameter("@Start", 133, 1, , ytdStart)
    cmd.Parameters.Append cmd.CreateParameter("@End", 133, 1, , qEnd)
    
    Set rs = cmd.Execute
    
    Do Until rs.EOF
        Dim vendorKey
        vendorKey = NzStr(rs("Vendor"))
        If Len(vendorKey) > 0 Then
            vendorsDict.Add vendorKey, vendorKey
        End If
        rs.MoveNext
    Loop
    rs.Close

    ' Process each vendor
    Dim vendorArr
    vendorArr = vendorsDict.Keys
    
    For Each vendorKey In vendorArr
        If Not ShouldExcludeVendor(vendorKey) Then
            PrintVendorIncidentDetails cn, vendorKey, ytdStart, qEnd, currentMonth
        End If
    Next

    ' Add summary section with monthly totals under each column
    DisplayIncidentTotals cn, ytdStart, qEnd, currentMonth
    
    Set cmd = Nothing
    On Error GoTo 0
End Sub



Private Function IsVendorExcluded(vendorName)
    Dim ex, v, key, i
    ex = ParseVendorExcludes()
    key = LCase(Trim(vendorName))
    IsVendorExcluded = False
    If IsArray(ex) Then
        For i = LBound(ex) To UBound(ex)
            v = LCase(Trim(ex(i)))
            If Len(v) > 0 And key = v Then
                IsVendorExcluded = True
                Exit Function
            End If
        Next
    End If
End Function




Private Sub PrintVendorIncidentDetails(cn, vendorName, startDate, endDate, currentMonth)
    On Error Resume Next
    
    ' Skip excluded vendors defensively (in case caller list wasn't filtered)
    If IsVendorExcluded(vendorName) Then Exit Sub
    
    Dim sql, cmd, rs
    Dim monthlyIncidents(12) ' Array to store monthly incident counts
    Dim totalIncidents, criticalCount, highCount, lowCount
    Dim lastIncidentDate, trend
    Dim i, hasData

    ' Load thresholds from INI (defaults 600/120)
    Dim tCrit, tHigh
    tCrit = GetSeverityThresholds()(0)
    tHigh = GetSeverityThresholds()(1)
    
    ' Initialize arrays
    For i = 1 To 12
        monthlyIncidents(i) = 0
    Next
    totalIncidents = 0
    criticalCount = 0
    highCount = 0
    lowCount = 0
    hasData = False
    
    ' Get monthly incident breakdown with severity classification
    sql = "SELECT " & _
          "  MONTH(ReportDate) as MonthNum, " & _
          "  SUM(COALESCE(IncidentCount, 0)) AS MonthIncidents, " & _
          "  SUM(TotalDowntimeMinutes) AS DowntimeMinutes, " & _
          "  MAX(ReportDate) AS LastDate " & _
          "FROM vw_VendorAvailability_CardWizard " & _
          "WHERE Vendor = ? AND ReportDate >= ? AND ReportDate < ? " & _
          "GROUP BY MONTH(ReportDate) " & _
          "ORDER BY MONTH(ReportDate)"
    
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = 1
    cmd.CommandText = sql
    cmd.Parameters.Append cmd.CreateParameter("@Vendor", 200, 1, 200, vendorName)
    cmd.Parameters.Append cmd.CreateParameter("@Start", 133, 1, , startDate)
    cmd.Parameters.Append cmd.CreateParameter("@End", 133, 1, , endDate)
    
    Set rs = cmd.Execute
    If Err.Number <> 0 Or rs Is Nothing Then Exit Sub
    
lastIncidentDate = Empty   ' Add this before the query

Do Until rs.EOF
		Dim monthNum, monthIncidents, downtimeMinutes, currentMonthDate
		monthNum = CLng(rs("MonthNum"))
		monthIncidents = CLng(NzD(rs("MonthIncidents")))
		downtimeMinutes = NzD(rs("DowntimeMinutes"))
		currentMonthDate = rs("LastDate")

		If monthIncidents > 0 Then
			If IsEmpty(lastIncidentDate) Or currentMonthDate > lastIncidentDate Then
				lastIncidentDate = currentMonthDate
			End If
		End If

			
        monthlyIncidents(monthNum) = monthIncidents
        totalIncidents = totalIncidents + monthIncidents
        
        ' Classify incidents by severity based on monthly total downtime (INI thresholds)
        If downtimeMinutes >= tCrit Then
            criticalCount = criticalCount + monthIncidents
        ElseIf downtimeMinutes >= tHigh Then
            highCount = highCount + monthIncidents
        ElseIf downtimeMinutes > 0 Then
            lowCount = lowCount + monthIncidents
        End If
        
        hasData = True
        rs.MoveNext
    Loop
    rs.Close
    
    If Not hasData Then Exit Sub
    
    ' Calculate trend
    trend = CalculateIncidentTrend(cn, vendorName, startDate, endDate, currentMonth)
    
    ' Format last incident date
    Dim lastDateStr
    If IsNull(lastIncidentDate) Or IsEmpty(lastIncidentDate) Then
        lastDateStr = "Never"
    Else
        lastDateStr = FormatDateTime(lastIncidentDate, 2) ' Short date
    End If
    
    ' Build output line
    Dim output
    output = PadRight(vendorName, 30)
    
    ' Add monthly incident counts
    For i = 1 To currentMonth
        If monthlyIncidents(i) > 0 Then
            output = output & PadRight(CStr(monthlyIncidents(i)), 15)
        Else
            output = output & PadRight("0", 15)
        End If
    Next
    
    ' Add summary columns
    output = output & PadRight(CStr(totalIncidents), 15) & _
                   PadRight(CStr(criticalCount), 15) & _
                   PadRight(CStr(highCount), 15) & _
                   PadRight(CStr(lowCount), 15) & _
                   PadRight(lastDateStr, 15) & _
                   PadRight(trend, 15)
    
    Fout.WriteLine output
    
    Set cmd = Nothing
    On Error GoTo 0
End Sub



Private Sub DisplayIncidentTotals(cn, startDate, endDate, currentMonth)
    On Error Resume Next
    
    Dim sql, cmd, rs
    Dim monthlyTotals(12)
    Dim grandTotal, criticalTotal, highTotal, lowTotal
    Dim i
    
    ' Load thresholds once
    Dim thresholds : thresholds = GetSeverityThresholds()
    Dim tCrit : tCrit = thresholds(0)
    Dim tHigh : tHigh = thresholds(1)
    
    For i = 1 To 12: monthlyTotals(i) = 0: Next
    grandTotal = 0: criticalTotal = 0: highTotal = 0: lowTotal = 0
    
    ' IMPORTANT: Remove the HAVING clause that filtered out 0-downtime incidents
    sql = "SELECT " & _
          " MONTH(ReportDate) AS MonthNum, " & _
          " SUM(COALESCE(IncidentCount,0))            AS MonthIncidents, " & _
          " SUM(TotalDowntimeMinutes)                 AS MonthDowntime " & _
          "FROM vw_VendorAvailability_CardWizard " & _
          "WHERE ReportDate >= ? AND ReportDate < ? " & _
          BuildVendorExcludePredicate("Vendor") & " " & _
          "GROUP BY Vendor, YEAR(ReportDate), MONTH(ReportDate) " & _
          "ORDER BY MONTH(ReportDate)"
    
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = cn
    cmd.CommandType = 1
    cmd.CommandText = sql
    cmd.Parameters.Append cmd.CreateParameter("@Start", 133, 1, , startDate)
    cmd.Parameters.Append cmd.CreateParameter("@End",   133, 1, , endDate)
    
    Set rs = cmd.Execute
    
    Do Until rs.EOF
        Dim monthNum        : monthNum       = CLng(rs("MonthNum"))
        Dim incidents       : incidents      = CLng(rs("MonthIncidents"))
        Dim downtimeMinutes : downtimeMinutes = NzD(rs("MonthDowntime"))
        
        ' Monthly total (now includes incidents with 0 downtime)
        monthlyTotals(monthNum) = monthlyTotals(monthNum) + incidents
        grandTotal = grandTotal + incidents
        
        ' Severity still based only on actual downtime (correct business rule)
        If downtimeMinutes >= tCrit Then
            criticalTotal = criticalTotal + incidents
        ElseIf downtimeMinutes >= tHigh Then
            highTotal = highTotal + incidents
        ElseIf downtimeMinutes > 0 Then
            lowTotal = lowTotal + incidents
        End If
        
        rs.MoveNext
    Loop
    rs.Close
    Set cmd = Nothing
    
    ' Build footer line
    Dim line : line = ""
    For i = 1 To currentMonth
        line = line & PadRight(CStr(monthlyTotals(i)), 15)
    Next
    
    line = line & PadRight(CStr(grandTotal), 15) & _
                  PadRight(CStr(criticalTotal), 15) & _
                  PadRight(CStr(highTotal), 15) & _
                  PadRight(CStr(lowTotal), 15)
    
    Fout.WriteLine ""
    Fout.WriteLine PadRight("Incident Tracking Totals", 30) & line
    
    On Error GoTo 0
End Sub


Private Function CalculateIncidentTrend(cn, vendorName, startDate, endDate, currentMonth)
    ' Compare current month vs previous month
    On Error Resume Next
    
    If currentMonth < 2 Then
        CalculateIncidentTrend = "N/A"
        Exit Function
    End If
    
    Dim currentMonthCount, previousMonthCount
    Dim currentMonthStart, currentMonthEnd, previousMonthStart, previousMonthEnd
    
    ' Current month (most recent)
    currentMonthStart = DateSerial(Year(startDate), currentMonth, 1)
    currentMonthEnd = DateAdd("m", 1, currentMonthStart)
    
    ' Previous month
    previousMonthStart = DateAdd("m", -1, currentMonthStart)
    previousMonthEnd = currentMonthStart
    
    currentMonthCount = GetVendorIncidentCount(cn, vendorName, currentMonthStart, currentMonthEnd)
    previousMonthCount = GetVendorIncidentCount(cn, vendorName, previousMonthStart, previousMonthEnd)
    
    If previousMonthCount = 0 Then
        If currentMonthCount = 0 Then
            CalculateIncidentTrend = "0"       ' No incidents both months
        Else
            CalculateIncidentTrend = "+2"      ' New incidents (treat as significant increase)
        End If
    Else
        Dim changePct
        changePct = ((currentMonthCount - previousMonthCount) / previousMonthCount) * 100
        
        If changePct > 25 Then
            CalculateIncidentTrend = "+2"      ' Significant increase
        ElseIf changePct > 10 Then
            CalculateIncidentTrend = "+1"      ' Moderate increase
        ElseIf changePct < -25 Then
            CalculateIncidentTrend = "-2"      ' Significant decrease
        ElseIf changePct < -10 Then
            CalculateIncidentTrend = "-1"      ' Moderate decrease
        Else
            CalculateIncidentTrend = "0"       ' No change
        End If
    End If
End Function

Private Function GetVendorIncidentCount(cn, vendorName, startDate, endDate)
    On Error Resume Next
    
    Dim sql, cmd, rs, count
    count = 0
    
    sql = "SELECT SUM(COALESCE(IncidentCount, 0)) AS IncCount " & _
          "FROM vw_VendorAvailability_CardWizard " & _
          "WHERE Vendor = ? AND ReportDate >= ? AND ReportDate < ?"
    
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = 1
    cmd.CommandText = sql
    cmd.Parameters.Append cmd.CreateParameter("@Vendor", 200, 1, 100, vendorName)
    cmd.Parameters.Append cmd.CreateParameter("@Start", 133, 1, , startDate)
    cmd.Parameters.Append cmd.CreateParameter("@End", 133, 1, , endDate)
    
    Set rs = cmd.Execute
    If Not rs.EOF Then
        count = CLng(NzD(rs("IncCount")))
    End If
    
    If rs.State = 1 Then rs.Close
    Set cmd = Nothing
    
    GetVendorIncidentCount = count
End Function

Private Sub DisplayIncidentSummary(cn, startDate, endDate, currentMonth)
    On Error Resume Next
    
    Dim sql, cmd, rs
    Dim totalIncidents, vendorsWithIncidents, maxIncidents
    Dim monthlyTotals(12)
    Dim i
    
    ' Initialize monthly totals
    For i = 1 To 12
        monthlyTotals(i) = 0
    Next
    
    ' Get overall statistics (exclude applied)
    sql = "SELECT " & _
          "  COUNT(DISTINCT Vendor) AS VendorCount, " & _
          "  SUM(COALESCE(IncidentCount, 0)) AS TotalIncidents, " & _
          "  MAX(COALESCE(IncidentCount, 0)) AS MaxIncidents " & _
          "FROM vw_VendorAvailability_CardWizard " & _
          "WHERE ReportDate >= ? AND ReportDate < ? " & _
          "AND COALESCE(IncidentCount, 0) > 0 " & _
          BuildVendorExcludePredicate("Vendor")
    
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = 1
    cmd.CommandText = sql
    cmd.Parameters.Append cmd.CreateParameter("@Start", 133, 1, , startDate)
    cmd.Parameters.Append cmd.CreateParameter("@End", 133, 1, , endDate)
    
    Set rs = cmd.Execute
    
    If Not rs.EOF Then
        vendorsWithIncidents = CLng(NzD(rs("VendorCount")))
        totalIncidents = CLng(NzD(rs("TotalIncidents")))
        maxIncidents = CLng(NzD(rs("MaxIncidents")))
    End If
    rs.Close
    
    ' Get monthly totals (exclude applied)
    sql = "SELECT " & _
          "  MONTH(ReportDate) as MonthNum, " & _
          "  SUM(COALESCE(IncidentCount, 0)) AS MonthTotal " & _
          "FROM vw_VendorAvailability_CardWizard " & _
          "WHERE ReportDate >= ? AND ReportDate < ? " & _
          BuildVendorExcludePredicate("Vendor") & _
          " GROUP BY MONTH(ReportDate) " & _
          "ORDER BY MONTH(ReportDate)"
    
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = 1
    cmd.CommandText = sql
    cmd.Parameters.Append cmd.CreateParameter("@Start", 133, 1, , startDate)
    cmd.Parameters.Append cmd.CreateParameter("@End", 133, 1, , endDate)
    
    Set rs = cmd.Execute
    
    Do Until rs.EOF
        Dim monthNum, monthTotal
        monthNum = CLng(rs("MonthNum"))
        monthTotal = CLng(NzD(rs("MonthTotal")))
        monthlyTotals(monthNum) = monthTotal
        rs.MoveNext
    Loop
    rs.Close
    
    ' Display summary
    Fout.WriteLine PadRight("Total Vendors with Incidents", 30) & vendorsWithIncidents
    Fout.WriteLine PadRight("Total Incidents YTD", 30) & totalIncidents
    Fout.WriteLine PadRight("Average per Month", 30) & FormatNumber(totalIncidents / currentMonth, 1)
    
    Set cmd = Nothing
End Sub


Function GetSeverityThresholds()
    Dim raw, parts, tCrit, tHigh
    tCrit = 600 : tHigh = 120  ' defaults
    raw = GetIni(gCfg, "Filters.SeverityThresholds")
    If Len(Trim(raw)) > 0 Then
        parts = Split(raw, ",")
        If UBound(parts) >= 1 Then
            ' Treat first value as Critical, second as High
            If IsNumeric(Trim(parts(0))) Then tCrit = CLng(Trim(parts(0)))
            If IsNumeric(Trim(parts(1))) Then tHigh = CLng(Trim(parts(1)))
        End If
    End If
    GetSeverityThresholds = Array(tCrit, tHigh)
End Function




Private Function GetVendorStatus3(ByVal availability, ByVal targetAvailability, ByVal yellowAvailability)

    availability = CDbl(availability)
    targetAvailability = CDbl(targetAvailability)
    yellowAvailability = CDbl(yellowAvailability)

    If availability >= targetAvailability Then
        GetVendorStatus3 = "PASS"
    ElseIf availability >= yellowAvailability Then
        GetVendorStatus3 = "MONITOR"
    Else
        GetVendorStatus3 = "FAIL"
    End If

End Function


'========================
'  HELP DESK REPORTS
'========================

Private Function GetIniFolder()
    Dim p, pos
    p = INI_PATH
    pos = InStrRev(p, "\")
    If pos > 0 Then
        GetIniFolder = Left(p, pos)
    Else
        GetIniFolder = ""
    End If
End Function

Private Function GetHelpDeskFolder()
    Dim raw, folder
    raw = NullToEmpty(GetIni(gCfg, "HelpDesk.ReportsFolder"))
    If Len(raw) = 0 Then raw = "HelpDeskReports"

    If InStr(raw, ":\") > 0 Or Left(raw, 2) = "\\" Then
        folder = raw
    Else
        folder = GetIniFolder() & raw
    End If

    If Right(folder, 1) <> "\" Then folder = folder & "\"
    GetHelpDeskFolder = folder
End Function

Private Function ParseStatusSet(rawValue)
    Dim d, parts, i, t
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1   ' vbTextCompare (case-insensitive)

    If Len(Trim(NullToEmpty(rawValue))) = 0 Then
        Set ParseStatusSet = d
        Exit Function
    End If

    parts = Split(rawValue, ",")
    For i = 0 To UBound(parts)
        t = Trim(parts(i))
        If Len(t) > 0 And Not d.Exists(t) Then
            d.Add t, True
        End If
    Next
    Set ParseStatusSet = d
End Function

Private Function ParseCsvRow(ByVal rowText)
    Dim i, ch, inQuotes, field, out(), idx
    ReDim out(0)
    idx = 0
    field = ""
    inQuotes = False

    For i = 1 To Len(rowText)
        ch = Mid(rowText, i, 1)
        If ch = """" Then
            inQuotes = Not inQuotes
        ElseIf ch = "," And Not inQuotes Then
            out(idx) = field
            idx = idx + 1
            ReDim Preserve out(idx)
            field = ""
        Else
            field = field & ch
        End If
    Next
    out(idx) = field
    ParseCsvRow = out
End Function

' Parse a Help Desk CSV using INI-driven status classification.
'   Solved = status appears in HelpDesk.SolvedStatuses
'   Open   = everything else (anything not yet solved)
'
' Returns:
'   (0) Total
'   (1) Open
'   (2) Resolved
'   (3) PctSolved   -- Resolved / Total * 100
'   (4) FoundFlag
Private Function ReadHelpDeskCSV(filePath)
    Dim fso, ts, line, parts
    Dim statusCol, lineNum, status, i
    Dim total, openCnt, resolvedCnt
    Dim solvedSet
    Dim result(4)

    result(0) = 0 : result(1) = 0 : result(2) = 0 : result(3) = 0 : result(4) = False

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        ReadHelpDeskCSV = result
        Exit Function
    End If

    Set solvedSet = ParseStatusSet(GetIni(gCfg, "HelpDesk.SolvedStatuses"))

    ' Default if INI key is missing
    If solvedSet.Count = 0 Then
        solvedSet.Add "Resolved", True
        solvedSet.Add "Closed",   True
        solvedSet.Add "Approved", True
    End If

    Set ts = fso.OpenTextFile(filePath, 1, False)
    total = 0 : openCnt = 0 : resolvedCnt = 0
    lineNum = 0
    statusCol = -1

    Do Until ts.AtEndOfStream
        line = ts.ReadLine
        lineNum = lineNum + 1
        parts = ParseCsvRow(line)

        If lineNum = 1 Then
            For i = 0 To UBound(parts)
                If StrComp(Trim(parts(i)), "Status", vbTextCompare) = 0 Then
                    statusCol = i
                    Exit For
                End If
            Next
            If statusCol = -1 Then statusCol = 2   ' fallback: column C
        Else
            If UBound(parts) >= statusCol Then
                status = Trim(parts(statusCol))
                If Len(status) > 0 Then
                    total = total + 1
                    If solvedSet.Exists(status) Then
                        resolvedCnt = resolvedCnt + 1
                    Else
                        openCnt = openCnt + 1
                    End If
                End If
            End If
        End If
    Loop
    ts.Close

    result(0) = total
    result(1) = openCnt
    result(2) = resolvedCnt
    If total > 0 Then
        result(3) = (resolvedCnt / total) * 100
    Else
        result(3) = 0
    End If
    result(4) = True

    ReadHelpDeskCSV = result
End Function

Sub DisplayHelpDeskYTD(periodStart)
    Dim folder, fname, fpath, r
    folder = GetHelpDeskFolder()
    fname = "Yearly Report - " & Month(periodStart) & "-1-" & Year(periodStart) & ".csv"
    fpath = folder & fname

    Fout.WriteLine ""
    Fout.WriteLine "HELP DESK (YTD)"

    r = ReadHelpDeskCSV(fpath)
    If Not r(4) Then
        Fout.WriteLine "  Source file not found: " & fpath
        Exit Sub
    End If

    'Fout.WriteLine "  Source: " & fname
    Fout.WriteLine PadRight("Open Tickets",     30) & r(1)
    Fout.WriteLine PadRight("Total Tickets",    30) & r(0)
    Fout.WriteLine PadRight("Resolved Tickets", 30) & r(2)
    Fout.WriteLine PadRight("% Solved",         30) & FormatNumber(r(3), 1) & "%"
End Sub

Sub DisplayHelpDeskMTD(periodStart)
    Dim folder, fname, fpath, r
    folder = GetHelpDeskFolder()
    fname = "Monthly Report - " & MonthName(Month(periodStart)) & ".csv"
    fpath = folder & fname

	Dim reportMonthDate
	reportMonthDate = DateAdd("d", -1, qEnd)
    Fout.WriteLine ""
    Fout.WriteLine  "HELP DESK (MTD)"
 
              

    r = ReadHelpDeskCSV(fpath)
    If Not r(4) Then
        Fout.WriteLine "  Source file not found: " & fpath
        Exit Sub
    End If

    'Fout.WriteLine "  Source: " & fname
    Fout.WriteLine PadRight("Open Tickets",     30) & r(1)
    Fout.WriteLine PadRight("Total Tickets",    30) & r(0)
    Fout.WriteLine PadRight("Resolved Tickets", 30) & r(2)
    Fout.WriteLine PadRight("% Solved",         30) & FormatNumber(r(3), 1) & "%"
End Sub


Function GetHelpDeskYTDLines1(periodStart)
    Dim folder, fname, fpath, r
    Dim lines(3)

    folder = GetHelpDeskFolder()
    fname = "Yearly Report - " & Month(periodStart) & "-1-" & Year(periodStart) & ".csv"
    fpath = folder & fname

    r = ReadHelpDeskCSV(fpath)

    If Not r(4) Then
        lines(0) = "Source file not found: " & fpath
        lines(1) = ""
        lines(2) = ""
        lines(3) = ""
    Else
        lines(0) = PadRight("Open Tickets",     25) & r(1)
        lines(1) = PadRight("Total Tickets",    25) & r(0)
        lines(2) = PadRight("Resolved Tickets", 25) & r(2)
        lines(3) = PadRight("% Solved",         25) & FormatNumber(r(3), 1) & "%"
    End If

    GetHelpDeskYTDLines = lines
End Function


Function GetHelpDeskMTDLines(periodStart)
    Dim folder, fname, fpath, r
    Dim lines(3)

    folder = GetHelpDeskFolder()
    fname = "Monthly Report - " & MonthName(Month(periodStart)) & ".csv"
    fpath = folder & fname

    r = ReadHelpDeskCSV(fpath)

    If Not r(4) Then
        lines(0) = "Source file not found: " & fpath
        lines(1) = ""
        lines(2) = ""
        lines(3) = ""
    Else
        lines(0) = PadRight("Open Tickets",     20) & r(1)
        lines(1) = PadRight("Total Tickets",    20) & r(0)
        lines(2) = PadRight("Resolved", 20) & r(2)
        lines(3) = PadRight("% Solved",20) & FormatNumber(r(3), 1) & "%"
    End If

    GetHelpDeskMTDLines = lines
End Function

Function GetHelpDeskYTDLines(periodStart)
    Dim folder, fname, fpath, r
    Dim lines(3)

    folder = GetHelpDeskFolder()
    fname = "Yearly Report - " & Month(periodStart) & "-1-" & Year(periodStart) & ".csv"
    fpath = folder & fname

    r = ReadHelpDeskCSV(fpath)

    If Not r(4) Then
        lines(0) = PadRight("File not found", 20)
        lines(1) = PadRight("Expected file", 20) & fname
        lines(2) = ""
        lines(3) = ""
    Else
        lines(0) = PadRight("Open Tickets",     20) & r(1)
        lines(1) = PadRight("Total Tickets",    20) & r(0)
        lines(2) = PadRight("Resolved", 20) & r(2)
        lines(3) = PadRight("% Solved",20) & FormatNumber(r(3), 1) & "%"
    End If

    GetHelpDeskYTDLines = lines
End Function