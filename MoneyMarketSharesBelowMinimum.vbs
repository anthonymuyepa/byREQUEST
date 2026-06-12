Sub ProcessLine
    ' Skip line if it contains "Federal Credit Union"
    If InStr(sRec, "Federal Credit Union") > 0 Then Exit Sub

    ' Skip line if it contains dashes ("----------")
    If InStr(sRec, "----------") > 0 Then Exit Sub
    If InStr(sRec, "                                     Maturity") > 0 Then Exit Sub
    ' Skip line if it contains "From:"
    If InStr(sRec, "************************************") > 0 Then Exit Sub
	If InStr(sRec, "Totals for Entire") > 0 Then Exit Sub
	If InStr(sRec, "Selected Record Count") > 0 Then Exit Sub
	If Len(Trim(sRec)) = 0 Then Exit Sub
	
	
	
    fout.WriteLine Trim(sRec)
End Sub