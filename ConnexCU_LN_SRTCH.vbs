Option Explicit

Dim inRecord
Dim currRecord

Sub StartDoc
    inRecord = False
    currRecord = ""

    ' Skip header lines, but write the first 7 to output
    Dim i
    For i = 1 To 30
        srec = getLine   
        If i < 7 Then
            fout.WriteLine srec
        End If
    Next
	dim header
 
	header = "Account Number  Customer Name                Major Minor    Old Rt  Mgn Fxd   Min Rt    New Rt  PTyp    Note Balance     Old Pmt Amt  Nxt Pmt Chg Pmt Rounding Meth  Index Rt  Mgn Pct   Max Rt       Term  PCal    Credit Limit     New Pmt Amt  Next Rt Chg  Rt Rounding Meth                                       Rt Eff Dt  PSch     Next Due Dt    Calc Pmt Amt                    Pmt Chg Tol"	
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
       Or InStr(line, "Rounding Meth") > 0 _
       Or InStr(line, "Pmt Chg Tol") > 0 _
       Or InStr(line, "Tax ID") > 0 _
       Or InStr(line, "Account Number") > 0 Then
        Exit Sub
    End If

    '---------------------------------------------------------
    ' Blank line = end of record
    '---------------------------------------------------------
	Dim logicalBlank
	logicalBlank = Trim(Replace(line, Chr(12), ""))

	If Len(logicalBlank) < 2 Then
		' Blank line or FF = end of current record
		If inRecord And Len(currRecord) > 0 Then
			' (If you added the shift logic, keep it here before writing)
					' ----------- SHIFT BLOCK -----------
			Dim A, B, C, D
			A = Mid(currRecord, 1, 16)
			B = Mid(currRecord, 137, 29)
			C = Mid(currRecord, 17, 119)
			D = Mid(currRecord, 166)

			currRecord = A & B & C & D
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
    first16 = Left(line, 16)
    AcctKey = Trim(first16)

    If IsAllDigits(AcctKey) Then
        ' Start of a new record

        ' If we were already in a record, flush the previous one
		If inRecord And Len(currRecord) > 0 Then
		' ----------- SHIFT BLOCK -----------

			A = Mid(currRecord, 1, 16)
			B = Mid(currRecord, 137, 29)
			C = Mid(currRecord, 17, 119)
			D = Mid(currRecord, 166)

			currRecord = A & B & C & D
			' ----------------------------------

			fout.WriteLine currRecord
		End If


        ' Build new record text starting from this line
        padLen = 136 - Len(line)
        If padLen < 1 Then padLen = 1

        currRecord = line & String(padLen, " ")
        inRecord = True

    Else
        ' Continuation line: append to current record
        If inRecord Then
            padLen = 136 - Len(line)
            If padLen < 1 Then padLen = 1

            currRecord = currRecord & line & String(padLen, " ")
        End If
    End If
End Sub

Sub CloseDoc
    ' Flush last record if file doesn't end with a blank line
    If inRecord And Len(currRecord) > 0 Then
	
	
				' ----------- SHIFT BLOCK -----------
			Dim A, B, C, D
			A = Mid(currRecord, 1, 16)
			B = Mid(currRecord, 137, 29)
			C = Mid(currRecord, 17, 119)
			D = Mid(currRecord, 166)

			currRecord = A & B & C & D
			' ----------------------------------

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
