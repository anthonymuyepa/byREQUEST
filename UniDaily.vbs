
Dim reachedEnd
sub startDoc
    Dim i
    For i = 1 To 3
		srec = getLine
		If len(srec)>3 Then
			fout.WriteLine srec
		End if
    Next
	

End Sub

Sub ProcessLine
    If reachedEnd Then Exit Sub


        If InStr(srec, "Totals for Entire Report") = 1 Then
            reachedEnd = True
            Exit Sub
        End If

        If InStr(srec, "====") > 0 Then Exit Sub

        If InStr(srec, "UniWyo Credit") > 0 Then Exit Sub
        If InStr(srec, "Account Name") > 0 Then Exit Sub			

		If len(srec)>3 Then
			fout.WriteLine srec
		End if
 
End Sub

