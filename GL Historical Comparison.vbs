Option Explicit

Dim inRecord
Dim currRecord

Sub StartDoc
    inRecord = False
    currRecord = ""

	fout.WriteLine "GL Account  Name                                                      Current                 Prior              Variance       Var%"
End Sub

Sub ProcessLine
    Dim line, first10, AcctKey, padLen

    line = srec

    '---------------------------------------------------------
    ' Skip report header/boilerplate lines
    '---------------------------------------------------------
    If InStr(line, "Run Date:") > 0 _
       Or InStr(line, "Post Date:") > 0 _
       Or InStr(line, "------") > 0 _
	   Or InStr(line, "GL Account  Name") > 0 _
	   Or InStr(line, "waiting wiating :") > 0 _
	   Or InStr(line, "=======") > 0 _
       Or InStr(line, "Carolina Trust Federal Credit Union") > 0 _
	   Or Instr(line, "GL Accounting Month:") > 0 _
	   Or Instr(line, "Branches:") > 0 _
       Or InStr(line, "Current Vs Prior Month") > 0 Then
        Exit Sub
    End If

	If Mid(line, 13, 5) = "TOTAL" Then
		' This is a TOTAL line
		Exit Sub
	End If
    '---------------------------------------------------------
    ' Blank line = end of record
    '---------------------------------------------------------
    If Len(Trim(line)) = 0 Then
        If inRecord And Len(currRecord) > 0 Then
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
    first10 = Left(line, 6)
    AcctKey = Trim(first10)

    If IsAllDigits(AcctKey) Then
        ' Start of a new record

        ' If we were already in a record, flush the previous one
		If inRecord And Len(currRecord) > 0 Then
			fout.WriteLine currRecord
		End If


        currRecord = line
        inRecord = True

    Else
        ' Continuation line: append to current record
        If inRecord Then
          currRecord = currRecord & line
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
