

Sub ProcessLine
    If InStr(sRec, "Old Value              New Value") > 0 Then
        fout.WriteLine "Account     ID    Post Date   User   Record        Field               Old Value                         New Value"
    
    ElseIf Len(Trim(sRec)) > 2 And InStr(sRec, "---------------------") = 0 Then
        fout.WriteLine Trim(sRec)
    End If
End Sub
