Option Explicit

Dim inRecord
Dim currRecord

Sub StartDoc
    inRecord = False
    currRecord = ""
	Dim i
	For i = 1 to 1
		srec = GetLine
		fout.WriteLine srec
	Next
	
	fout.WriteLine "Account #  Name                  Shr          Withdraw  Street           City             ST Zip        Lst Tr     Closed     Death Date"

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
	   Or Instr(line, "New Share Accounts This Period ") > 0 _
	   Or Instr(line, "Branches:") > 0 _
       Or InStr(line, "Account    Name                ID") > 0 Then
        Exit Sub
    End If

    '---------------------------------------------------------
    ' Blank line = end of record
    '---------------------------------------------------------
    If Len(Trim(line)) = 0 Then
        If inRecord And Len(currRecord) > 0 Then
		
				If Len(currRecord) < 104 Then
					currRecord = currRecord & Space(104 - Len(currRecord))
				End If

				If InStr(Mid(currRecord, 99, 1), "-") > 0 Then
					' Insert 1 space at position 104
					' → keep chars 1–103, add " ", then the rest starting at 104
					currRecord = Left(currRecord, 103) & " " & Mid(currRecord, 104)
				Else
					' Insert 3 spaces at position 99
					' → keep chars 1–98, add "   ", then the rest starting at 99
					currRecord = Left(currRecord, 98) & "   " & Mid(currRecord, 99)
				End If
				'currRecord = Replace(currRecord,"$"," ")
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
    first10 = Left(line, 10)
    AcctKey = Trim(first10)

    If IsAllDigits(AcctKey) Then
        ' Start of a new record

        ' If we were already in a record, flush the previous one
		If inRecord And Len(currRecord) > 0 Then
		
				If Len(currRecord) < 104 Then
					currRecord = currRecord & Space(104 - Len(currRecord))
				End If

				If InStr(Mid(currRecord, 99, 1), "-") > 0 Then
					' Insert 1 space at position 104
					' → keep chars 1–103, add " ", then the rest starting at 104
					currRecord = Left(currRecord, 103) & " " & Mid(currRecord, 104)
				Else
					' Insert 3 spaces at position 99
					' → keep chars 1–98, add "   ", then the rest starting at 99
					currRecord = Left(currRecord, 98) & "   " & Mid(currRecord, 99)
				End If
		
			    'currRecord = Replace(currRecord,"$","")
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
	
				If Len(currRecord) < 104 Then
					currRecord = currRecord & Space(104 - Len(currRecord))
				End If

				If InStr(Mid(currRecord, 99, 1), "-") > 0 Then
					' Insert 1 space at position 104
					' → keep chars 1–103, add " ", then the rest starting at 104
					currRecord = Left(currRecord, 103) & " " & Mid(currRecord, 104)
				Else
					' Insert 3 spaces at position 99
					' → keep chars 1–98, add "   ", then the rest starting at 99
					currRecord = Left(currRecord, 98) & "   " & Mid(currRecord, 99)
				End If

	
		'currRecord = Replace(currRecord,"$","")
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
