Option Explicit

' ---------- CONFIGURE THESE POSITIONS TO MATCH YOUR REPORT LAYOUT ---------
' Column positions (1-based) and lengths taken from the detail lines
Dim POS_KeyDateTime, POS_RespPers, POS_Actv, POS_TableId, POS_ColumnId, POS_OldData, POS_NewData
Dim LEN_KeyDateTime, LEN_RespPers, LEN_Actv, LEN_TableId, LEN_ColumnId, LEN_OldData

POS_KeyDateTime = 1
LEN_KeyDateTime = 34          ' e.g. "11/10/2025 09:09:01 PM"

POS_RespPers = 36
LEN_RespPers = 20             ' adjust as needed

POS_Actv = 57
LEN_Actv = 5                  ' "ADD", "CHG", etc.

POS_TableId = 63
LEN_TableId = 16

POS_ColumnId = 79
LEN_ColumnId = 20

POS_OldData = 99
LEN_OldData = 18

POS_NewData = 118             ' where "New Data" text starts

' Output widths for the flattened line
Dim W_ActivityCategory, W_ActivityType, W_KeyDateTime, W_RespPers, W_Actv
Dim W_TableId, W_ColumnId, W_OldData, W_NewData

W_ActivityCategory = 12
W_ActivityType     = 10
W_KeyDateTime      = 34
W_RespPers         = 20
W_Actv             = 5
W_TableId          = 16
W_ColumnId         = 20
W_OldData          = 18
W_NewData          = 80

' Global variables to track current record
Dim gActivityCategory, gActivityType, gKeyDateTime, gRespPers
Dim gActv, gTableId, gColumnId, gOldData, gNewData
Dim gRowPending
Dim gPendingDate, gPendingTime

'----------------- Helper: left-pad to fixed width -------------------------
Function PadField(val, width)
    PadField = Left(val & String(width, " "), width)
End Function

'========================= StartDoc ========================================
Sub StartDoc()
    Dim i, srec
    For i = 1 To 31
        srec = getLine
        If i < 7 Then
            Fout.WriteLine srec
        End If
    Next

    gActivityCategory = ""
    gActivityType     = ""
    gKeyDateTime      = ""
    gRespPers         = ""
    gActv             = ""
    gTableId          = ""
    gColumnId         = ""
    gOldData          = ""
    gNewData          = ""
    gRowPending       = False
    gPendingDate      = ""
    gPendingTime      = ""

    ' Optional header row for Excel
    Dim hdr
    hdr = ""
    hdr = hdr & PadField("ActivityCat", W_ActivityCategory)
    hdr = hdr & PadField("Type",        W_ActivityType)
    hdr = hdr & PadField("KeyDateTime", W_KeyDateTime)
    hdr = hdr & PadField("RespPers",    W_RespPers)
    hdr = hdr & PadField("Actv",        W_Actv)
    hdr = hdr & PadField("TableId",     W_TableId)
    hdr = hdr & PadField("ColumnId",    W_ColumnId)
    hdr = hdr & PadField("OldData",     W_OldData)
    hdr = hdr & PadField("NewData",     W_NewData)

    Fout.WriteLine hdr
End Sub

'========================= ProcessLine =====================================
Sub ProcessLine()
    Dim line
    line = srec
    line = Replace(Replace(Replace(line, Chr(12), ""), Chr(13), ""), Chr(10), "")

    '---------------------------------------------------------
    ' Skip report header/boilerplate lines
    '---------------------------------------------------------
    If InStr(line, "Run Date:") > 0 _
       Or InStr(line, "Post Date:") > 0 _
       Or InStr(line, "Key Date Time   RespPer") > 0 _
       Or InStr(line, "Page:") > 0 Then
        Exit Sub
    End If

    ' 1) Detect new Activity header
    If InStr(line, "Activity Category =") > 0 Then
        FlushPendingRow
        gPendingDate = ""
        gPendingTime = ""

        Dim p1, p2, tmp
        p1 = InStr(line, "Activity Category =")
        p2 = InStr(line, "Activity Type =")

        If p1 > 0 And p2 > 0 Then
            tmp = Mid(line, p1 + Len("Activity Category ="), _
                      p2 - (p1 + Len("Activity Category =")))
            gActivityCategory = CleanValue(Trim(tmp))
            gActivityType = CleanValue(Trim(Mid(line, p2 + Len("Activity Type ="))))
        End If

        Exit Sub
    End If

    ' Skip empty lines and non-detail fluff
    If Trim(line) = "" Then Exit Sub

    ' Check for date line (Key Date only row)
    If IsDateLine(line) Then
        ' If we already have a pending date and a pending record, flush it
        If gPendingDate <> "" And gRowPending Then
            FlushPendingRow
        End If
        gPendingDate = CleanValue(Trim(line))
        Exit Sub
    End If

    ' Check for time line (Key Time only row)
    If IsTimeLine(line) Then
        gPendingTime = CleanValue(Trim(line))
        ' Combine date and time when we have both and no detail on this line
        If gPendingDate <> "" And gPendingTime <> "" Then
            If gRowPending Then
                FlushPendingRow
            End If
            gKeyDateTime = gPendingDate & " " & gPendingTime
            gPendingDate = ""
            gPendingTime = ""
        End If
        Exit Sub
    End If

    ' 2) Parse columns from the detail line using fixed positions
    Dim kdt, rsp, actv, tbl, colId, oldD, newD

    kdt   = Trim(Mid(line, POS_KeyDateTime, LEN_KeyDateTime))
    rsp   = Trim(Mid(line, POS_RespPers,    LEN_RespPers))
    actv  = Trim(Mid(line, POS_Actv,        LEN_Actv))
    tbl   = Trim(Mid(line, POS_TableId,     LEN_TableId))
    colId = Trim(Mid(line, POS_ColumnId,    LEN_ColumnId))
    oldD  = Trim(Mid(line, POS_OldData,     LEN_OldData))
    newD  = RTrim(Mid(line, POS_NewData))   ' rest of line is New Data

    ' Clean the values
    kdt   = CleanValue(kdt)
    rsp   = CleanValue(rsp)
    actv  = CleanValue(actv)
    tbl   = CleanValue(tbl)
    colId = CleanValue(colId)
    oldD  = CleanValue(oldD)
    newD  = CleanValue(newD)

    ' 2a) Check if this line starts with a new date/time (new record group)
    Dim hasNewDateTime
    hasNewDateTime = (kdt <> "" And IsDateTime(kdt))

    If hasNewDateTime Then
        Dim fullKDT

        ' If kdt is only a time and we have a pending date, combine them
        If (InStr(kdt, "/") = 0 And InStr(kdt, "-") = 0) And gPendingDate <> "" Then
            fullKDT = gPendingDate & " " & kdt
        Else
            fullKDT = kdt
        End If

        FlushPendingRow

        gKeyDateTime = fullKDT
        gPendingDate = ""      ' consumed the pending date

        If rsp <> "" Then gRespPers = rsp

        ' If this line also has data fields, treat as a new record
        If actv <> "" Or tbl <> "" Or colId <> "" Or oldD <> "" Then
            gActv     = actv
            gTableId  = tbl
            gColumnId = colId
            gOldData  = oldD
            gNewData  = newD
            gRowPending = True
        End If
        Exit Sub
    ElseIf gPendingDate <> "" And gPendingTime <> "" Then
        ' Use the pending date/time if we have both
        gKeyDateTime = gPendingDate & " " & gPendingTime
        gPendingDate = ""
        gPendingTime = ""
    End If

    If rsp <> "" Then gRespPers = rsp

    ' 3) Check if this is a CONTINUATION line (empty key fields but has content)
    Dim isContinuation
    isContinuation = (actv = "" And tbl = "" And colId = "" And oldD = "" And Len(Trim(line)) > 0)

    If isContinuation And gRowPending Then
        ' This is a continuation of the current record's NewData
        Dim contData
        If Len(line) >= POS_NewData Then
            contData = Trim(RTrim(Mid(line, POS_NewData)))
        Else
            contData = Trim(line)
        End If

        If contData <> "" Then
            gNewData = gNewData & " " & contData
        End If
        Exit Sub
    End If

    ' 4) If not a continuation, check if this starts a NEW logical row
    If actv <> "" Or tbl <> "" Or colId <> "" Or oldD <> "" Then
        ' Start of a new logical record -> flush previous row
        FlushPendingRow

        gActv     = actv
        gTableId  = tbl
        gColumnId = colId
        gOldData  = oldD
        gNewData  = newD
        gRowPending = True
    End If
End Sub

'========================= CloseDoc ========================================
Sub CloseDoc()
    FlushPendingRow
End Sub

'========================= FlushPendingRow =================================
Sub FlushPendingRow()
    If Not gRowPending Then Exit Sub

    Dim outLine
    outLine = ""
    outLine = outLine & PadField(gActivityCategory, W_ActivityCategory)
    outLine = outLine & PadField(gActivityType,     W_ActivityType)
    outLine = outLine & PadField(gKeyDateTime,      W_KeyDateTime)
    outLine = outLine & PadField(gRespPers,         W_RespPers)
    outLine = outLine & PadField(gActv,             W_Actv)
    outLine = outLine & PadField(gTableId,          W_TableId)
    outLine = outLine & PadField(gColumnId,         W_ColumnId)
    outLine = outLine & PadField(gOldData,          W_OldData)
 
	'Remove 
	if Instr(gNewData, ":\") > 0 Then
		gNewData = replace( gNewData, " ", "") 
		
	End if
	
    outLine = outLine & PadField(gNewData,          W_NewData)

    Fout.WriteLine outLine

    ' Reset row-level fields
    gRowPending = False
    gActv = ""
    gTableId = ""
    gColumnId = ""
    gOldData = ""
    gNewData = ""
End Sub

'----------------- Helper: Check if line is *Key Date* only ----------------
Function IsDateLine(line)
    Dim segment, trimmed, rest
    ' Look only at the Key Date/Time field area
    segment = Mid(line, POS_KeyDateTime, LEN_KeyDateTime)
    trimmed = Trim(segment)
    ' Everything after the Key Date/Time area
    rest = Trim(Mid(line, POS_RespPers))

    ' It's a pure date header only if the rest of the line is blank
    If rest <> "" Then
        IsDateLine = False
        Exit Function
    End If

    ' Check for date pattern (MM/DD/YYYY or MM-DD-YYYY)
    If Len(trimmed) = 10 And (InStr(trimmed, "/") > 0 Or InStr(trimmed, "-") > 0) Then
        IsDateLine = True
    Else
        IsDateLine = False
    End If
End Function

'----------------- Helper: Check if line is *Key Time* only ----------------
Function IsTimeLine(line)
    Dim segment, trimmed, rest
    ' Look only at the Key Date/Time field area
    segment = Mid(line, POS_KeyDateTime, LEN_KeyDateTime)
    trimmed = Trim(segment)
    ' Everything after the Key Date/Time area
    rest = Trim(Mid(line, POS_RespPers))

    ' It's a pure time header only if the rest of the line is blank
    If rest <> "" Then
        IsTimeLine = False
        Exit Function
    End If

    ' Check for time pattern (HH:MM:SS AM/PM)
    If (InStr(trimmed, ":") > 0 And (InStr(trimmed, "AM") > 0 Or InStr(trimmed, "PM") > 0)) Then
        IsTimeLine = True
    Else
        IsTimeLine = False
    End If
End Function

'----------------- Helper: Check if string is a date/time ------------------
Function IsDateTime(value)
    Dim trimmed
    trimmed = Trim(value)
    ' Check for date-only or date+time or time-only pattern
    If (Len(trimmed) = 10 And (InStr(trimmed, "/") > 0 Or InStr(trimmed, "-") > 0)) Or _
       (InStr(trimmed, ":") > 0 And (InStr(trimmed, "AM") > 0 Or InStr(trimmed, "PM") > 0)) Then
        IsDateTime = True
    Else
        IsDateTime = False
    End If
End Function

'----------------- Helper: Clean value by removing unwanted text -----------
Function CleanValue(value)
    If IsNull(value) Or value = "" Then
        CleanValue = ""
    Else
        CleanValue = Trim(Replace(Replace(value, "GROUP", ""), "ENDURB", ""))
    End If
End Function


