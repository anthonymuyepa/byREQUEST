Dim header, endshares
header = 0
endshares = 0

Sub ProcessLine()
    ' srec already contains the current line

    If InStr(srec, "SHARE ANALYSIS BY AGE") > 0 Then
        header = 1
    End If

    ' If we are inside the SHARE ANALYSIS section
    If header = 1 And endshares = 0 Then
        ' Skip lines with only dashes
        If InStr(srec, "-----") = 0 Then
            ' Don't remove % from header line, only data lines
            If InStr(srec, "MEMBER AGE GROUP") = 0 Then
                fout.WriteLine AlignColumns(RemovePercents(srec))
            Else
                fout.WriteLine AlignColumns(srec)
            End If
        End If

        ' Check for TOTALS to stop capturing
        If InStr(srec, "TOTALS") > 0 Then
            endshares = 1
        End If
    End If

    ' (Optional) Default processing for other lines can go here
End Sub

' Helper function to align columns (optional to expand later)
Function AlignColumns(line)
    AlignColumns = line  '
End Function



' Helper function to remove percentage symbols from the line
Function RemovePercents(line)
    RemovePercents = Replace(line, "%", "")
End Function