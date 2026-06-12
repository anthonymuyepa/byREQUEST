'==========================================================
' ExcelIncidentConsumer_Ini.vbs
' Data Consumption Script — reads KPIs from IncidentReporter
' (Cleaned: unused queries/blocks removed)
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
	
	
    Fout.WriteLine "Reporting period: " & gPeriodLabel 

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

	DisplayYTDTrend cn, qStart, qEnd

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
          "AND YEAR(ReportDate) = ?"  ' Only current year vendors

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
          "GROUP BY Vendor "

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
    Dim vendorCount, vendorsMeetingGoal, perfectAvailabilityVendors
    Dim maxDowntimeHours, criticalIncidents, highIncidents
    
    totalDownMin = 0: totalMonMin = 0: totalIncidents = 0
    maxDowntimeHours = 0
    criticalIncidents = 0
    highIncidents = 0

    vendorCount = allVendorsDict.Count  ' Total vendors being monitored
    vendorsMeetingGoal = vendorCount    ' Start with all meeting goal, subtract those that fail
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
        
        If availability >= CDbl(targetAvailability) Then
            status = "PASS"
            ' vendorsMeetingGoal already counts all vendors as meeting goal
        Else
            status = "FAIL"
            vendorsMeetingGoal = vendorsMeetingGoal - 1  ' Subtract vendors that failed
        End If
                        
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
          "AND YEAR(ReportDate) = ?"  ' Only current year

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

    ' Calculate MoM change (placeholder - you'll need to implement this)
    Dim momChangePct
    momChangePct = 0  ' Placeholder - need to query previous month data

    Fout.WriteLine ""
    Fout.WriteLine "OVERALL PERFORMANCE SUMMARY"
    Fout.WriteLine PadRight("Vendors Meeting Goal", 30) & vendorsMeetingGoal & "/" & vendorCount & _
                   " (" & FormatNumber((vendorsMeetingGoal / vendorCount) * 100, 1) & "%)"
    Fout.WriteLine PadRight("Perfect Availability Vendors", 30) & perfectAvailabilityVendors & "/" & vendorCount & _
                   " (" & FormatNumber((perfectAvailabilityVendors / vendorCount) * 100, 1) & "%)"

    Fout.WriteLine PadRight("Max Single Vendor Downtime", 30) & FormatNumber(maxDowntimeHours, 2) & " hours"
    Fout.WriteLine PadRight("Severity Distribution", 30) & "Critical: " & criticalIncidents & ", High: " & highIncidents
 

    Fout.WriteLine PadRight("Total Downtime", 30) & FormatNumber(totalDownMin / 60, 2) & " hours"
    Fout.WriteLine PadRight("Total Incidents", 30) & totalIncidents
    Fout.WriteLine PadRight("Total Incidents YTD", 30) & totalIncidentsYTD

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
          "WHERE ReportDate = ?"

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
          "WHERE ReportDate >= ? AND ReportDate < ?"

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
    Fout.WriteLine Space(120) & "Calculation Summary"

    ' Data scope
    Fout.WriteLine Space(120) & "MonitoringMinutes:"
    Fout.WriteLine Space(120) & "--CardWizard: 9h/day x business days (US federal holidays excluded)."
    Fout.WriteLine Space(120) & "--Others: daysInMonth x 1,440."
    Fout.WriteLine Space(120) & "Availability% = (1 - DowntimeMinutes / MonitoringMinutes) x 100."
    Fout.WriteLine Space(120) & "PASS/FAIL vs goal (default 99.995%)."

    ' Core system availability (weighted)
    Fout.WriteLine Space(120) & "Current Period = (1 - SUM(DowntimeMinutes_all) / SUM(MonitoringMinutes_all)) x 100 for month at Start."
    Fout.WriteLine Space(120) & "YTD = (1 - SUM(DT_YTD) / SUM(Mon_YTD)) x 100 for [Jan 1, periodEnd)."
    Fout.WriteLine Space(120) & "Variance: +x.xxx% or (x.xxx%)."
	Fout.WriteLine ""
	Fout.WriteLine ""
	Fout.WriteLine Space(120) & "Why RED status with MoM = +0.000%:"
	Fout.WriteLine Space(120) & "-- GREEN requires: (1) YTD >= Target AND (2) MoM >= 0"
	Fout.WriteLine Space(120) & "-- MoM only compares recent two months"
	Fout.WriteLine Space(120) & "-- Status uses YTD average (all months since January)"
	Fout.WriteLine Space(120) & "-- Early outages keep YTD below target until diluted by clean months"
	Fout.WriteLine ""
	Fout.WriteLine Space(120) & "Detailed Example: January outage affecting YTD through July"
	Fout.WriteLine Space(120) & "-- January: 60 minutes downtime"
	Fout.WriteLine Space(120) & "-- Monthly monitoring: 9h/day x 22 days = 11,880 minutes"
	Fout.WriteLine Space(120) & "-- Total monitoring minutes (Jan-Jul): 7 months x 11,880 = 83,160 minutes"
	Fout.WriteLine Space(120) & "-- YTD Availability Calculation:"
	Fout.WriteLine Space(120) & "---- Available minutes = Total minutes - Downtime"
	Fout.WriteLine Space(120) & "---- Available minutes = 83,160 - 60 = 83,100 minutes"
	Fout.WriteLine Space(120) & "---- YTD % = (83,100 / 83,160) x 100 = 99.928%"
	Fout.WriteLine Space(120) & "-- Result: YTD (99.928%) < Target (99.995%) --> Status = RED"
	Fout.WriteLine ""
	Fout.WriteLine Space(120) & "  Meanwhile, if recent months are clean:"
	Fout.WriteLine Space(120) & "---- June = 100.000%, July = 100.000%"
	Fout.WriteLine Space(120) & "---- MoM = 100.000% - 100.000% = +0.000% (no change)"
	Fout.WriteLine ""
	Fout.WriteLine Space(120) & "  Bottom line: Early outages have long-lasting impact on YTD"
	Fout.WriteLine Space(120) & "  despite good recent performance (neutral MoM)"
	End Sub

' Year-to-date summary with simple green/red trending
Sub DisplayYTDTrend(cn, qStart, qEnd)
    On Error Resume Next

    Dim targetAvailability
    targetAvailability = GetIni(gCfg, "VendorAvailabilityGoals.TargetAvailability")
    If Len(targetAvailability) = 0 Then targetAvailability = 99.995

    ' Determine YTD window for the anchor year shown in the report
    Dim anchorYear, ytdStart, ytdEnd
    anchorYear = Year(DateAdd("d", -1, qEnd))  ' year of the period being shown
    ytdStart   = DateSerial(anchorYear, 1, 1)
    ytdEnd     = qEnd                           ' first day of month after anchor (exclusive)

    Fout.WriteLine ""
    Fout.WriteLine "YEAR-TO-DATE SUMMARY"
    Fout.WriteLine PadRight("Vendor", 30) & _
                   PadRight("Availability", 15) & _
                   PadRight("MoM Trend", 15) & _
                   PadRight("Status", 15)

    ' Pull monthly buckets for YTD
    Dim sql, cmd, rs
    sql = "SELECT Vendor, ReportDate, " & _
          "       SUM(TotalDowntimeMinutes) AS dmin, " & _
          "       SUM(TotalMonitoringMinutes) AS mmin " & _
          "FROM vw_VendorAvailability_CardWizard " & _
          "WHERE ReportDate >= ? AND ReportDate < ? " & _
          "GROUP BY Vendor, ReportDate " & _
          "ORDER BY Vendor, ReportDate"

    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = 1
    cmd.CommandText = sql
    cmd.Parameters.Append cmd.CreateParameter("@Start", 133, 1, , ytdStart)
    cmd.Parameters.Append cmd.CreateParameter("@End",   133, 1, , ytdEnd)

    Set rs = cmd.Execute
    If Err.Number <> 0 Then
        Fout.WriteLine "ERROR: YTD trend query failed: " & Err.Description
        Exit Sub
    End If

    If rs Is Nothing Or rs.EOF Then
        Fout.WriteLine "No YTD data available."
        If Not rs Is Nothing Then rs.Close
        Set cmd = Nothing
        Exit Sub
    End If

    ' Iterate vendor-by-vendor
    Dim curVendor, prevVendor, dsum, msum, prevAvail, lastAvail, havePrev
    prevVendor = ""
    dsum = 0 : msum = 0 : prevAvail = 0 : lastAvail = 0 : havePrev = False

    Do Until rs.EOF
        curVendor = NzStr(rs("Vendor"))
        If Len(curVendor) = 0 Then curVendor = "Internal"

        ' New vendor group? flush previous
        If prevVendor <> "" And curVendor <> prevVendor Then
            PrintVendorYTD prevVendor, dsum, msum, lastAvail, prevAvail, targetAvailability
            dsum = 0 : msum = 0 : prevAvail = 0 : lastAvail = 0 : havePrev = False
        End If

        ' Monthly availability for trend
        Dim dmin, mmin, monthAvail
        dmin = NzD(rs("dmin")) : mmin = NzD(rs("mmin"))
        If mmin > 0 Then monthAvail = (1 - (dmin / mmin)) * 100 Else monthAvail = 100

        ' Track MoM (last two months only needed for delta)
        If havePrev Then
            prevAvail = lastAvail
        End If
        lastAvail = monthAvail
        havePrev = True

        ' Accumulate YTD weights
        dsum = dsum + dmin
        msum = msum + mmin

        prevVendor = curVendor
        rs.MoveNext
    Loop

    ' flush last vendor
    If prevVendor <> "" Then
        PrintVendorYTD prevVendor, dsum, msum, lastAvail, prevAvail, targetAvailability
    End If

    rs.Close
    Set cmd = Nothing
    On Error GoTo 0
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