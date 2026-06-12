Option Explicit

Dim inRecord
Dim currRecord

Sub StartDoc
    inRecord = False
    currRecord = ""

    ' Skip header lines, but write the first 7 to output
    Dim i
    For i = 1 To 34
        srec = getLine   
        If i < 7 Then
			'fout.WriteLine srec
        End If
    Next
	dim header
 
	header = "Account Number  Parcel/Policy Number Customer Name                       Property  Address                                          Organization Name                    Org Number   Major Stat  Escrow Balance   Disb Due Date   Due Amount   Calendar Period  Map  Lot     Minor  Sold Escrow Payment   Pay Due Date   Last Amount   Type      Block          Description        Property"
	
	
	
	fout.WriteLine header


End Sub

Sub ProcessLine
    Dim line, first16, AcctKey, padLen

    line = srec

    '---------------------------------------------------------
    ' Skip report header/boilerplate lines
    '---------------------------------------------------------
    If InStr(line, "Run Date:") > 0 _
       Or InStr(line, "Post Date:") > 0 _
       Or InStr(line, "Page:") > 0 _
       Or InStr(line, "Maturity Date:") > 0 _
       Or InStr(line, "Due Amount   Calendar Period") > 0 _
       Or InStr(line, "Last Amount   Type") > 0 _
       Or InStr(line, "Org Number         Description") > 0 _
	   Or (InStr(line, "Property") > 0 and Len(line) <=10) _
       Or InStr(line, "Parcel/Policy Numberr") > 0 Then
        Exit Sub
    End If

    '---------------------------------------------------------
    ' Blank line = end of record
    '---------------------------------------------------------
	Dim logicalBlank
	logicalBlank = Trim(Replace(Replace(Replace(line, Chr(12), ""), Chr(13), ""), Chr(10), ""))

	If Len(logicalBlank) < 2 Then
		' Blank line or FF = end of current record
		If inRecord And Len(currRecord) > 0 Then
			' ----------- SHIFT BLOCK -----------
			Dim AcctNo, CustName, BlockB, Addr1, BlockC, Addr2, OrgName, PolicyNo, PropertyDesc, description

			' Core slices (adjust only if your spec says otherwise)
			AcctNo       = Mid(currRecord, 1, 16)     ' 1–16   (16 chars)

			' Fix overlap: use 36 chars so custName ends at 52 (17–52)
			CustName     = Mid(currRecord, 17, 36)    ' 17–52  (36 chars)

			' Match the comment (53–140 = 88 chars) instead of the old 67
			BlockB       = Mid(currRecord, 53, 88)    ' 53–140 (88 chars)

			Addr1        = Mid(currRecord, 142, 35)   ' 142–176 (35 chars)
			BlockC       = Mid(currRecord, 178, 83)   ' 178–260 (83 chars)
			Addr2        = Mid(currRecord, 262, 35)   ' 262–296 (35 chars)
			OrgName      = Mid(currRecord, 298, 50)   ' 298–347 (50 chars)
			Description  = Mid(currRecord, 348, 18)    ' 348–365 (18 chars)
			PropertyDesc = Mid(currRecord, 366, 59)   ' 366–424 (59 chars)
			PolicyNo     = Mid(currRecord, 426, 21)   ' 426–446 (21 chars)

			' Rebuild using *real* variable names (no A, E, F)
			currRecord = AcctNo & PolicyNo & CustName & PropertyDesc & OrgName & BlockB &  BlockC & Description & Addr1 & Addr2

		' ----------------------------------
			
			fout.WriteLine currRecord
			currRecord = ""
			inRecord = False
		End If
		Exit Sub
	End If

    '---------------------------------------------------------
    ' Detect start of a new record:
    ' a record is all numeric in Trim(first 16)
    '---------------------------------------------------------
	first16 = Left(line,16)

    AcctKey = Trim(first16)

    If IsAllDigits(AcctKey) and len(srec) > 110 Then
        ' Start of a new record

        ' If we were already in a record, flush the previous one
		If inRecord And Len(currRecord) > 0 Then
		' ----------- SHIFT BLOCK -----------
			

			' Core slices (adjust only if your spec says otherwise)
			AcctNo       = Mid(currRecord, 1, 16)     ' 1–16   (16 chars)

			' Fix overlap: use 36 chars so custName ends at 52 (17–52)
			CustName     = Mid(currRecord, 17, 36)    ' 17–52  (36 chars)

			' Match the comment (53–140 = 88 chars) instead of the old 67
			BlockB       = Mid(currRecord, 53, 88)    ' 53–140 (88 chars)

			Addr1        = Mid(currRecord, 142, 35)   ' 142–176 (35 chars)
			BlockC       = Mid(currRecord, 178, 83)   ' 178–260 (83 chars)
			Addr2        = Mid(currRecord, 262, 35)   ' 262–296 (35 chars)
			OrgName      = Mid(currRecord, 298, 50)   ' 298–347 (50 chars)
		    Description  = Mid(currRecord, 348, 18)    ' 348–365 (18 chars)
			PropertyDesc = Mid(currRecord, 366, 59)   ' 366–424 (59 chars)
			PolicyNo     = Mid(currRecord, 426, 21)   ' 426–446 (21 chars)

			' Rebuild using *real* variable names (no A, E, F)
			currRecord = AcctNo & PolicyNo & CustName & PropertyDesc & OrgName & BlockB &  BlockC & Description & Addr1 & Addr2
		' ----------------------------------

			fout.WriteLine currRecord
		End If


        ' Build new record text starting from this line
        padLen = 125 - Len(line)
        If padLen < 1 Then padLen = 1

        currRecord = line & String(padLen, " ")
        inRecord = True

    Else
        ' Continuation line: append to current record
		If inRecord Then
			Dim tLen
			tLen = Len(line)

			If tLen > 60 Then
				padLen = 120 - tLen
			Else

				padLen = 60 - tLen
			End If

			If padLen < 1 Then padLen = 1

			currRecord = currRecord & line & String(padLen, " ")
		End If
	End If

End Sub

Sub CloseDoc
    ' Flush last record if file doesn't end with a blank line
    If inRecord And Len(currRecord) > 0 Then
	
	
			Dim AcctNo, CustName, BlockB, Addr1, BlockC, Addr2, OrgName, PolicyNo, PropertyDesc, description

			' Core slices (adjust only if your spec says otherwise)
			AcctNo       = Mid(currRecord, 1, 16)     ' 1–16   (16 chars)

			' Fix overlap: use 36 chars so custName ends at 52 (17–52)
			CustName     = Mid(currRecord, 17, 36)    ' 17–52  (36 chars)

			' Match the comment (53–140 = 88 chars) instead of the old 67
			BlockB       = Mid(currRecord, 53, 88)    ' 53–140 (88 chars)

			Addr1        = Mid(currRecord, 142, 35)   ' 142–176 (35 chars)
			BlockC       = Mid(currRecord, 178, 83)   ' 178–260 (83 chars)
			Addr2        = Mid(currRecord, 262, 35)   ' 262–296 (35 chars)
			OrgName      = Mid(currRecord, 298, 50)   ' 298–347 (50 chars)
			Description  = Mid(currRecord, 348, 18)    ' 348–365 (18 chars)
			PropertyDesc = Mid(currRecord, 366, 59)   ' 366–424 (59 chars)
			PolicyNo     = Mid(currRecord, 426, 21)   ' 426–446 (21 chars)

			' Rebuild using *real* variable names (no A, E, F)
			currRecord = AcctNo & PolicyNo & CustName & PropertyDesc & OrgName & BlockB &  BlockC & Description & Addr1 & Addr2

        fout.WriteLine currRecord
    End If
End Sub

'==================== Helpers ====================

Function IsAllDigits(txt)
    Dim cleaned, re

    cleaned = Trim(txt)
    If Len(cleaned) = 0 Then
        IsAllDigits = False
        Exit Function
    End If

    Set re = New RegExp
    re.Pattern = "^\d+$"
    re.Global  = False
    re.IgnoreCase = True

    IsAllDigits = re.Test(cleaned)
End Function
