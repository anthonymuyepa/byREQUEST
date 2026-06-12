Option Explicit

Dim inRecord
Dim currRecord

Sub StartDoc
    inRecord = False
    currRecord = ""
	Dim i
	For i = 1 to 3 
		srec = GetLine
		fout.WriteLine srec
	Next


End Sub

Sub ProcessLine
    Dim line, first10, AcctKey, padLen, logicalBlank

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
	   Or Instr(line, "New Share Accounts This Period ") > 0 _
	   Or Instr(line, "Branches:") > 0 _
       Or InStr(line, "Account    Name                ID") > 0 Then
        Exit Sub
    End If


	' Blank line or FF  or carriagereturn
    logicalBlank = Trim(Replace(line, Chr(12), ""))

    If Len(logicalBlank) < 2 Then
        Exit Sub
    End If


    '---------------------------------------------------------
    ' Detect start of a new record:
    ' a record is all numeric in Trim(first 16)
    '---------------------------------------------------------
    first10 = Left(line, 10)
    AcctKey = Trim(first10)

	If IsAllDigits(AcctKey) Then
		' Start of a new record
		fout.WriteLine line
		
		currRecord = line
		inRecord = True

	Else
		' Continuation line: append to current record
		If inRecord Then
			' Check if current line is full record
			If Left(line, 30) = Space(30) Then
					line = Left(currRecord, 29) & Mid(line, 30)
					fout.WriteLine line
			End If
		End If
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
