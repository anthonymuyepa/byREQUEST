Sub ProcessLine
    ' Skip line if it contains "Federal Credit Union"
    If InStr(sRec, "Federal Credit Union") > 0 Then Exit Sub

    ' Skip line if it contains dashes ("----------")
    If InStr(sRec, "----------") > 0 Then Exit Sub
    If InStr(sRec, "                                     Maturity") > 0 Then Exit Sub
    ' Skip line if it contains "From:"
    If InStr(sRec, "From:") > 0 Then Exit Sub
	If InStr(sRec, "Totals for Entire") > 0 Then Exit Sub
	If InStr(sRec, "Selected Record Count") > 0 Then Exit Sub
	If Len(Trim(sRec)) = 0 Then Exit Sub
	If InStr(sRec, "Acct. No.  ID    Name                Date") > 0 Then 
		sRec ="Acct. No.  ID    Name                 Maturity Date"
	End if 
	
	Dim account, ID, parts
	If InStr(Left(sRec, 13), "-") > 0 Then 
		parts = Split(Left(sRec, 13), "-")
		Account = parts(0)
		ID = parts(1)
		srec = Account & " " & ID & mid(sRec,14)
	End if
	
    fout.WriteLine Trim(sRec)
End Sub