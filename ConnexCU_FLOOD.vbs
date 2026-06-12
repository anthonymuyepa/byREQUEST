Option Explicit

Dim inRecord
Dim currRecord

Sub StartDoc
    inRecord = False
    currRecord = ""

    ' Skip header lines, but write the first 7 to output
    Dim i
    For i = 1 To 28
        srec = getLine   
        If i < 7 Then
            fout.WriteLine srec
        End If
    Next
	
	fout.WriteLine " Account Number  Name                      Major  Esc  Zone            Ins Org Name           Coverage Amount    Expiration Date Minor       Panel           Policy Number              Prop Value     Maturity Date                     Property Address                              Property Address2                     Community       Premium Frequency      Premium Amount     Loan Next Due Date"
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
       Or InStr(line, "Loan Next Due Date") > 0 _
       Or InStr(line, "Expiration Date") > 0 _
       Or InStr(line, "Tax ID") > 0 _
       Or InStr(line, "Account Number") > 0 Then
        Exit Sub
    End If

    '---------------------------------------------------------
    ' Blank line = end of record
    '---------------------------------------------------------
    If Len(Trim(line)) = 0 Then
        If inRecord And Len(currRecord) > 0 Then
    
			' ----------- SHIFT BLOCK -----------
			Dim A, B, C, D, E
			A = Mid(currRecord, 1, 129)
			B = Mid(currRecord, 130, 45)
			C = Mid(currRecord, 175, 85)
			D = Mid(currRecord, 260, 20)
			E = Mid(currRecord, 280)
			currRecord = A & C & B & D & E
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

			A = Mid(currRecord, 1, 129)
			B = Mid(currRecord, 130, 45)
			C = Mid(currRecord, 175, 85)
			D = Mid(currRecord, 260, 20)
			E = Mid(currRecord, 280)
			currRecord = A & C & B & D & E
			' ----------------------------------

			fout.WriteLine currRecord
		End If


        ' Build new record text starting from this line
        padLen = 130 - Len(line)
        If padLen < 1 Then padLen = 1

        currRecord = line & String(padLen, " ")
        inRecord = True

    Else
        ' Continuation line: append to current record
        If inRecord Then
            padLen = 130 - Len(line)
            If padLen < 1 Then padLen = 1

            currRecord = currRecord & " " & line & String(padLen, " ")
        End If
    End If
End Sub

Sub CloseDoc
    ' Flush last record if file doesn't end with a blank line
    If inRecord And Len(currRecord) > 0 Then
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
