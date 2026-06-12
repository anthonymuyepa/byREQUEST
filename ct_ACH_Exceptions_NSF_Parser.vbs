Option Explicit

Dim gReport, gRptDate, gRptTime, gRptSeq
Dim gSourceFile, gLineCounter, gRowsWritten
Dim gIgnoreNextDash

Dim gBlockLines
Set gBlockLines = CreateObject("System.Collections.ArrayList")

Sub StartDoc()

    gLineCounter = 0
    gRowsWritten = 0
	gIgnoreNextDash = False
    gSourceFile = ""

    On Error Resume Next
    gSourceFile = SpoolFile.Name
    On Error GoTo 0

    srec = getline()
    ParseReportTitleLine srec, "ACH Exceptions - NSF"

    ' HEADER
    fout.WriteLine _
    PadRight("Report", 28) & _
    PadRight("RptDate", 12) & _
    PadRight("RptTime", 10) & _
    PadRight("RptSeq", 10) & _
    PadRight("Acct", 14) & _
    PadRight("SLX", 6) & _
    PadRight("ID", 6) & _
    PadRight("Name", 24) & _
    PadRight("TranDesc", 18) & _
    PadRight("CompanyName", 24) & _
    PadRight("Amount", 14) & _
    PadRight("Available", 14) & _
    PadRight("Balance", 14) & _
    PadRight("Description", 12) & _
    PadRight("OD1_Acct", 14) & _
    PadRight("OD1_Type", 10) & _
    PadRight("OD1_ID", 10) & _
    PadRight("OD1_Available", 14) & _
    PadRight("OD1_Balance", 14) & _
    PadRight("OD2_Acct", 14) & _
    PadRight("OD2_Type", 10) & _
    PadRight("OD2_ID", 10) & _
    PadRight("OD2_Available", 14) & _
    PadRight("OD2_Balance", 14) & _
    PadRight("CompanyID", 18) & _
    PadRight("SettlementDate", 16) & _
    PadRight("TransmissionDate", 18) & _
    PadRight("ECC", 6) & _
    PadRight("Disc", 24) & _
    PadRight("Entry", 18) & _
    PadRight("OrigAcct", 20) & _
    PadRight("EntryID", 18) & _
    PadRight("EntryName", 24) & _
    PadRight("TraceNumber", 18) & _
	PadRight("Raw1", 100) & _
    PadRight("LineOrder", 10) & _
    PadRight("SourceFile", 80)
End Sub

Sub ProcessLine()

    Dim line
    line = srec
    gLineCounter = gLineCounter + 1

    If Trim(line) = "" Then Exit Sub

    ' Repeating page header inside a record.
    ' The next dashed line is page decoration, not block end.
    If IsRepeatedColumnHeader(line) Then
        gIgnoreNextDash = True
        Exit Sub
    End If

    ' BLOCK SEPARATOR
    If StartsWithText(line, "-----") Then

        If gIgnoreNextDash Then
            gIgnoreNextDash = False
            Exit Sub
        End If

        If gBlockLines.Count > 0 Then
            ParseBlock gBlockLines
            gBlockLines.Clear
        End If

        Exit Sub
    End If

    If IsSkipLine(line) Then Exit Sub

    gBlockLines.Add line

End Sub

Sub CloseDoc()

    If gBlockLines.Count > 0 Then
        ParseBlock gBlockLines
        gBlockLines.Clear
    End If

End Sub

' =========================
'  CORE BLOCK PARSER
' =========================
Sub ParseBlock(ByVal blk)

    Dim i, line
    Dim acct, slx, idVal, nameVal, tranDesc, companyName
    Dim amount, available, balance, description
    Dim odCount
    Dim od1Acct, od1Type, od1ID, od1Available, od1Balance
    Dim od2Acct, od2Type, od2ID, od2Available, od2Balance
    Dim companyID, settlementDate, transmissionDate, ecc
    Dim entry, origAcct, entryID, entryName, traceNumber, disc
	Dim raw1
 

    acct = "" : slx = "" : idVal = "" : nameVal = "" : tranDesc = "" : companyName = ""
    amount = "" : available = "" : balance = "" : description = ""
    odCount = 0
    od1Acct = "" : od1Type = "" : od1ID = "" : od1Available = "" : od1Balance = ""
    od2Acct = "" : od2Type = "" : od2ID = "" : od2Available = "" : od2Balance = ""
    companyID = "" : settlementDate = "" : transmissionDate = "" : ecc = ""
    entry = "" : origAcct = "" : entryID = "" : entryName = "" : traceNumber = "" : disc = ""
	raw1 =""

    If blk.Count = 0 Then Exit Sub

    For i = 0 To blk.Count - 1
        line = blk(i)

        If IsNSFMainLine(line) Then
            acct        = SafeMidTrim(line, 1, 10)
            slx         = SafeMidTrim(line, 12, 1)
            idVal       = SafeMidTrim(line, 14, 2)
            nameVal     = SafeMidTrim(line, 17, 16)
            tranDesc    = SafeMidTrim(line, 34, 15)
            companyName = SafeMidTrim(line, 50, 20)
            amount      = SafeMidTrim(line, 69, 12)
            available   = SafeMidTrim(line, 84, 10)
            balance     = SafeMidTrim(line, 97, 11)
            description = SafeMidTrim(line, 108, 11)

        ElseIf InStr(1, line, "Overdraw Protection", vbTextCompare) > 0 Then
            odCount = odCount + 1
			ParseOverdrawLineFlat line, odCount, _
				od1Acct, od1Type, od1ID, od1Available, od1Balance, _
				od2Acct, od2Type, od2ID, od2Available, od2Balance

        ElseIf StartsWithText(line, "Company ID") Then
            companyID        = ExtractBetween(line, "Company ID", "Settlement")
            settlementDate   = ExtractBetween(line, "Settlement", "Transmission")
            transmissionDate = ExtractBetween(line, "Transmission", "ECC")
            ecc              = ExtractAfter(line, "ECC")

        ElseIf StartsWithText(line, "Disc") Then
            disc        = SafeMidTrim(line, 6, 20)
            entry       = ExtractBetween(line, "Ent", "Acct")
            origAcct    = ExtractBetween(line, "Acct", "ID")
            entryID     = ExtractBetween(line, "ID", "Nme")
            entryName   = ExtractBetween(line, "Nme", "Trc")
            traceNumber = ExtractAfter(line, "Trc")
		Else
            If Len(raw1) = 0 Then raw1 = Trim(line)
        End If

    Next

    If Len(acct) = 0 Or Len(description) = 0 Then Exit Sub

    WriteCurrentRecord acct, slx, idVal, nameVal, tranDesc, companyName, _
        amount, available, balance, description, _
        od1Acct, od1Type, od1ID, od1Available, od1Balance, _
        od2Acct, od2Type, od2ID, od2Available, od2Balance, _
        companyID, settlementDate, transmissionDate, ecc, _
        disc, entry, origAcct, entryID, entryName, traceNumber, raw1

End Sub

Sub ParseOverdrawLineFlat(ByVal line, ByVal odCount, _
    ByRef od1Acct, ByRef od1Type, ByRef od1ID, ByRef od1Available, ByRef od1Balance, _
    ByRef od2Acct, ByRef od2Type, ByRef od2ID, ByRef od2Available, ByRef od2Balance)

    Dim rest, parts

    rest = Trim(Replace(line, "Overdraw Protection", "", 1, -1, vbTextCompare))
    parts = SplitBySpaces(rest)

    If Not IsArray(parts) Then Exit Sub
    If UBound(parts) < 2 Then Exit Sub

    If odCount = 1 Then
        od1Acct = parts(0)
        od1Type = parts(1)
        od1ID   = parts(2)

        If UBound(parts) >= 3 Then
            If IsNumeric(Left(parts(3), 1)) Then
                od1Available = parts(3)
                If UBound(parts) >= 4 Then od1Balance = parts(4)
            Else
                od1Available = JoinFromIndex(parts, 3)
                od1Balance = ""
            End If
        End If

    ElseIf odCount = 2 Then
        od2Acct = parts(0)
        od2Type = parts(1)
        od2ID   = parts(2)

        If UBound(parts) >= 3 Then
            If IsNumeric(Left(parts(3), 1)) Then
                od2Available = parts(3)
                If UBound(parts) >= 4 Then od2Balance = parts(4)
            Else
                od2Available = JoinFromIndex(parts, 3)
                od2Balance = ""
            End If
        End If
    End If

End Sub

Function JoinFromIndex(ByVal arr, ByVal startIndex)

    Dim i, result
    result = ""

    For i = startIndex To UBound(arr)
        If Len(result) > 0 Then result = result & " "
        result = result & arr(i)
    Next

    JoinFromIndex = result

End Function


' =========================
' WRITE OUTPUT
' =========================
Sub WriteCurrentRecord(ByVal acct, ByVal slx, ByVal idVal, ByVal nameVal, _
    ByVal tranDesc, ByVal companyName, ByVal amount, _
    ByVal available, ByVal balance, ByVal description, _
    ByVal od1Acct, ByVal od1Type, ByVal od1ID, ByVal od1Available, ByVal od1Balance, _
    ByVal od2Acct, ByVal od2Type, ByVal od2ID, ByVal od2Available, ByVal od2Balance, _
    ByVal companyID, ByVal settlementDate, ByVal transmissionDate, _
    ByVal ecc, ByVal disc, ByVal entry, ByVal origAcct, ByVal entryID, _
    ByVal entryName, ByVal traceNumber, ByVal raw1)

    fout.WriteLine _
        PadRight(gReport, 28) & _
        PadRight(gRptDate, 12) & _
        PadRight(gRptTime, 10) & _
        PadRight(gRptSeq, 10) & _
        PadRight(acct, 14) & _
        PadRight(slx, 6) & _
        PadRight(idVal, 6) & _
        PadRight(nameVal, 24) & _
        PadRight(tranDesc, 18) & _
        PadRight(companyName, 24) & _
        PadRight(amount, 14) & _
        PadRight(available, 14) & _
        PadRight(balance, 14) & _
        PadRight(description, 12) & _
        PadRight(od1Acct, 14) & _
        PadRight(od1Type, 10) & _
        PadRight(od1ID, 10) & _
        PadRight(od1Available, 14) & _
        PadRight(od1Balance, 14) & _
        PadRight(od2Acct, 14) & _
        PadRight(od2Type, 10) & _
        PadRight(od2ID, 10) & _
        PadRight(od2Available, 14) & _
        PadRight(od2Balance, 14) & _
        PadRight(companyID, 18) & _
        PadRight(settlementDate, 16) & _
        PadRight(transmissionDate, 18) & _
        PadRight(ecc, 6) & _
        PadRight(disc, 24) & _
        PadRight(entry, 18) & _
        PadRight(origAcct, 20) & _
        PadRight(entryID, 18) & _
        PadRight(entryName, 24) & _
        PadRight(traceNumber, 18) & _
		PadRight(raw1, 100) & _
		PadRight(gLineCounter, 10) & _
        PadRight(gSourceFile, 80)

    gRowsWritten = gRowsWritten + 1

End Sub


Function SplitBySpaces(ByVal s)

    Dim re, matches, arr(), i

    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "\S+"
    re.Global = True

    Set matches = re.Execute(s)

    If matches.Count = 0 Then
        SplitBySpaces = Array()
        Exit Function
    End If

    ReDim arr(matches.Count - 1)

    For i = 0 To matches.Count - 1
        arr(i) = matches(i).Value
    Next

    SplitBySpaces = arr

End Function

' =========================
' HELPERS
' =========================
Function IsSkipLine(ByVal line)
    Dim s
    s = Trim(line)

    If s = "" Then IsSkipLine = True: Exit Function
    If Left(s, 5) = "Acct " Then IsSkipLine = True: Exit Function
    If InStr(1, s, "Total Exceptions", vbTextCompare) > 0 Then IsSkipLine = True: Exit Function
    If InStr(1, s, "Page", vbTextCompare) > 0 And InStr(1, s, "ACH Exceptions", vbTextCompare) > 0 Then IsSkipLine = True: Exit Function

    IsSkipLine = False
End Function

Function SafeMidTrim(ByVal s, ByVal startPos, ByVal lengthVal)
    If startPos > Len(s) Then
        SafeMidTrim = ""
    Else
        SafeMidTrim = Trim(Mid(s, startPos, lengthVal))
    End If
End Function

Function StartsWithText(ByVal line, ByVal prefix)
    StartsWithText = (LCase(Left(Trim(line), Len(prefix))) = LCase(prefix))
End Function

Function ExtractBetween(ByVal s, ByVal startToken, ByVal endToken)
    Dim p1, p2
    ExtractBetween = ""

    p1 = InStr(1, s, startToken, vbTextCompare)
    If p1 = 0 Then Exit Function

    p1 = p1 + Len(startToken)
    p2 = InStr(p1, s, endToken, vbTextCompare)

    If p2 = 0 Then
        ExtractBetween = Trim(Mid(s, p1))
    Else
        ExtractBetween = Trim(Mid(s, p1, p2 - p1))
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

Function IsRepeatedColumnHeader(ByVal line)

    Dim s
    s = Trim(line & "")

    IsRepeatedColumnHeader = False

    If InStr(1, s, "Acct", vbTextCompare) > 0 And _
       InStr(1, s, "S/L/X", vbTextCompare) > 0 And _
       InStr(1, s, "Tran Desc", vbTextCompare) > 0 And _
       InStr(1, s, "Company Name", vbTextCompare) > 0 Then

        IsRepeatedColumnHeader = True
    End If

End Function

Function IsNSFMainLine(ByVal line)
    Dim s
    s = Trim(line & "")

    IsNSFMainLine = False

    If Len(s) < 20 Then Exit Function
    If InStr(1, s, "NSF", vbTextCompare) = 0 Then Exit Function

    ' First visible character should be numeric or ? followed by account digits
    If IsNumeric(Left(s, 1)) Or Left(s, 1) = "?" Then
        IsNSFMainLine = True
    End If
End Function

Function ExtractAfter(ByVal s, ByVal token)
    Dim p
    ExtractAfter = ""

    p = InStr(1, s, token, vbTextCompare)
    If p = 0 Then Exit Function

    ExtractAfter = Trim(Mid(s, p + Len(token)))
End Function