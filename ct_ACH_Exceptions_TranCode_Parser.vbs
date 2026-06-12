Option Explicit

Dim gBlockLines, gLineCounter, gSourceFile
Dim gReport, gRptDate, gRptTime, gRptSeq 
Dim gStopProcessing

Sub StartDoc()

	gStopProcessing = False
    Set gBlockLines = CreateObject("System.Collections.ArrayList")
    gLineCounter = 0

    On Error Resume Next
    gSourceFile = SpoolFile.Name
    On Error GoTo 0

	Dim firstLine
	
	srec = getline()
	firstLine = srec   

	ParseReportTitleLine firstLine, "ACH Exceptions - Tran Code"

	' Skip header lines AFTER extracting metadata

'	srec = getline()

    ' Header (fixed width)
	fout.WriteLine _
		PadRight("Report", 32) & _
		PadRight("RptDate", 12) & _
		PadRight("RptTime", 10) & _
		PadRight("RptSeq", 10) & _
		PadRight("Acct", 12) & _
        PadRight("SLX", 6) & _
        PadRight("Name", 24) & _
        PadRight("TranDesc", 18) & _
        PadRight("CompanyName", 22) & _
        PadRight("Amount", 12) & _
        PadRight("Available", 12) & _
        PadRight("Balance", 12) & _
        PadRight("Description", 28) & _
        PadRight("CompanyID", 14) & _
        PadRight("Settlement", 14) & _
        PadRight("Transmission", 14) & _
        PadRight("ECC", 6) & _
		PadRight("Disc", 20) & _
        PadRight("Entry", 16) & _
        PadRight("OrigAcct", 16) & _
        PadRight("EntryID", 14) & _
        PadRight("EntryName", 24) & _
        PadRight("Trace", 18) & _
        PadRight("ExceptionText", 30) & _
        PadRight("Raw1", 100) & _
        PadRight("LineOrder", 10) & _
        PadRight("SourceFile", 60)

End Sub

Sub ProcessLine()
	If gStopProcessing Then Exit Sub
    
	Dim line
    line = srec
	
	If IsTotalsLine(line) Then

		If gBlockLines.Count > 0 Then
			ProcessBlock
			gBlockLines.Clear
		End If

		gStopProcessing = True
		Exit Sub
	End If
 
	If IsSkipLine(line) Then Exit Sub


    If InStr(line, "-----") > 0 Then

        If gBlockLines.Count > 0 Then
            ProcessBlock
            gBlockLines.Clear
        End If

        Exit Sub
    End If

    If Trim(line) <> "" Then
        gBlockLines.Add line
    End If

End Sub

Sub CloseDoc()

    If gBlockLines.Count > 0 Then
        ProcessBlock
    End If

End Sub

' ============================
' PROCESS BLOCK
' ============================
Sub ProcessBlock()

    Dim acct, slx, name, tranDesc, companyName
    Dim amount, available, balance, description
    Dim entry, origAcct, entryID, entryName, trace
    Dim exceptionText, raw1, raw2
	Dim companyID, settlement, transmission, ecc, disc


    acct = ""
    slx = ""
    name = ""
    tranDesc = ""
    companyName = ""
    amount = ""
    available = ""
    balance = ""
    description = ""
    companyID = ""
    settlement = ""
    transmission = ""
    ecc = ""
	disc = ""
    entry = ""
    origAcct = ""
    entryID = ""
    entryName = ""
    trace = ""
    exceptionText = ""
    raw1 = ""
    raw2 = ""

    ' ---- LINE 1 (fixed width)
    If gBlockLines.Count >= 1 Then
        Dim l1
        l1 = gBlockLines(0)

        acct        = SafeMidTrim(l1, 1, 10)
        slx         = SafeMidTrim(l1, 11, 1)
        name        = SafeMidTrim(l1, 16, 17)
        tranDesc    = SafeMidTrim(l1, 33, 16)
        companyName = SafeMidTrim(l1, 49, 18)
        amount      = SafeMidTrim(l1, 69, 11)
        available   = SafeMidTrim(l1, 83, 10)
        balance     = SafeMidTrim(l1, 97, 10)
        description = SafeMidTrim(l1, 107, 30)
    End If

    ' ---- LINE 2 (Company)
    If gBlockLines.Count >= 2 Then
        Dim l2
        l2 = gBlockLines(1)

        companyID   = ExtractBetween(l2, "Company ID", "Settlement")
        settlement  = ExtractBetween(l2, "Settlement", "Transmission")
        transmission= ExtractBetween(l2, "Transmission", "ECC")
        ecc         = ExtractAfter(l2, "ECC")
    End If

    ' ---- LINE 3 (Entry / Trace)
	If gBlockLines.Count >= 3 Then
		Dim l3
		l3 = gBlockLines(2)

		disc      = ExtractBetween(l3, "Disc", "Ent")
		entry     = ExtractBetween(l3, "Ent", "Acct")
		origAcct  = ExtractBetween(l3, "Acct", "ID")
		entryID   = ExtractBetween(l3, "ID", "Nme")
		entryName = ExtractBetween(l3, "Nme", "Trc")
		trace     = ExtractAfter(l3, "Trc")
	End If

    ' ---- LINE 4 (Exception)
    If gBlockLines.Count >= 4 Then
        exceptionText = Trim(gBlockLines(3))
    End If

    ' ---- LINE 5 (Raw)
    If gBlockLines.Count >= 5 Then
        raw1 = Trim(gBlockLines(4))
    End If

    If gBlockLines.Count >= 6 Then
        raw2 = Trim(gBlockLines(5))
    End If
    
	gLineCounter = gLineCounter + 1
    ' ---- OUTPUT
    fout.WriteLine _
		PadRight(gReport, 32) & _
		PadRight(gRptDate, 12) & _
		PadRight(gRptTime, 10) & _
		PadRight(gRptSeq, 10) & _
		PadRight(acct, 12) & _
		PadRight(slx, 6) & _
        PadRight(name, 24) & _
        PadRight(tranDesc, 18) & _
        PadRight(companyName, 22) & _
        PadRight(amount, 12) & _
        PadRight(available, 12) & _
        PadRight(balance, 12) & _
        PadRight(description, 28) & _
        PadRight(companyID, 14) & _
        PadRight(settlement, 14) & _
        PadRight(transmission, 14) & _
        PadRight(ecc, 6) & _
		PadRight(disc, 20) & _
        PadRight(entry, 16) & _
        PadRight(origAcct, 16) & _
        PadRight(entryID, 14) & _
        PadRight(entryName, 24) & _
        PadRight(trace, 18) & _
        PadRight(exceptionText, 30) & _
        PadRight(raw1, 100) & _
        PadRight(gLineCounter, 10) & _
        PadRight(gSourceFile, 60)

End Sub

' ============================
' HELPERS
' ============================
Function SafeMidTrim(ByVal s, ByVal startPos, ByVal lengthVal)
    If startPos > Len(s) Then
        SafeMidTrim = ""
    Else
        SafeMidTrim = Trim(Mid(s, startPos, lengthVal))
    End If
End Function

Function ExtractBetween(ByVal s, ByVal a, ByVal b)
    Dim p1, p2
    ExtractBetween = ""

    p1 = InStr(1, s, a, vbTextCompare)
    If p1 = 0 Then Exit Function
    p1 = p1 + Len(a)

    p2 = InStr(p1, s, b, vbTextCompare)
    If p2 = 0 Then
        ExtractBetween = Trim(Mid(s, p1))
    Else
        ExtractBetween = Trim(Mid(s, p1, p2 - p1))
    End If
End Function

Function ExtractAfter(ByVal s, ByVal token)
    Dim p
    p = InStr(1, s, token, vbTextCompare)

    If p = 0 Then
        ExtractAfter = ""
    Else
        ExtractAfter = Trim(Mid(s, p + Len(token)))
    End If
End Function

Function PadRight(ByVal s, ByVal n)
    s = Trim(s & "")
    If Len(s) >= n Then
        PadRight = Left(s, n)
    Else
        PadRight = s & Space(n - Len(s))
    End If
End Function


Sub ParseReportTitleLine(ByVal line, ByVal reportName)

    Dim pAt, pSeq, pPage

    gReport = reportName
    gRptDate = ""
    gRptTime = ""
    gRptSeq = ""

    pAt = InStr(1, line, " at ", vbTextCompare)
    pSeq = InStr(1, line, " Seq ", vbTextCompare)
    pPage = InStr(1, line, " Page ", vbTextCompare)

	If pAt > 0 Then
		gRptDate = Trim(Mid(line, pAt - 8, 8))
		gRptTime = Trim(Mid(line, pAt + 13, 5))
	End If
	
    If pSeq > 0 Then
        If pPage > pSeq Then
            gRptSeq = Trim(Mid(line, pSeq + 5, pPage - (pSeq + 5)))
        Else
            gRptSeq = Trim(Mid(line, pSeq + 5))
        End If
    End If

End Sub

Function IsSkipLine(ByVal line)

    Dim s
    s = Trim(line & "")

    IsSkipLine = False

    If Len(s) = 0 Then
        IsSkipLine = True
        Exit Function
    End If

    ' ----- dashed separator handled separately
    If InStr(s, "-----") > 0 Then Exit Function

    ' Header line
    If InStr(s, "Acct") > 0 And _
       InStr(s, "S/L/X") > 0 And _
       InStr(s, "Tran Desc") > 0 Then
        IsSkipLine = True
        Exit Function
    End If

    ' Totals footer
    If InStr(s, "Total Exceptions") > 0 Then
        IsSkipLine = True
        Exit Function
    End If

    ' Page header repeat
    If InStr(s, "ACH Exceptions") > 0 And InStr(s, "Page") > 0 Then
        IsSkipLine = True
        Exit Function
    End If

    ' Form feed / junk
    If InStr(s, Chr(12)) > 0 Then
        IsSkipLine = True
        Exit Function
    End If
	

End Function


Function IsTotalsLine(ByVal line)

    Dim s
    s = Trim(line & "")

    IsTotalsLine = False

    If InStr(s, "Total Exceptions") > 0 And _
       InStr(s, "Unposted Amount") > 0 And _
       InStr(s, "Credit Amount") > 0 Then

        IsTotalsLine = True
    End If

End Function