Option Explicit

Dim currentGroup(), memberCount, accountNumber

Sub StartDoc
    Dim i
    
    For i = 1 To 28
        srec = getLine
		if i < 9 Then
			fout.WriteLine srec
		End if
	Next
    
	Fout.Writeline "Name                              Tax ID         Address                         City                State  Zip       # of Dups PurgeYN   Account Number(s)"

    ReDim currentGroup(0)
    memberCount = 0
    accountNumber = ""
End Sub

Sub ProcessLine
    Dim currentLine
    currentLine = srec

    '---------------------------------------------------------
    ' Skip report header / footer / column-heading lines
    '---------------------------------------------------------
    If InStr(currentLine, "Run Date:") > 0 _
       Or InStr(currentLine, "Post Date:") > 0 _
       Or InStr(currentLine, "Page:") > 0 _
       Or (InStr(currentLine, "Name") > 0 And InStr(currentLine, "Tax ID") > 0) _
       Or InStr(currentLine, "Account Numbe") > 0 Then
        Exit Sub
    End If

    '---------------------------------------------------------
    ' Normal processing below
    '---------------------------------------------------------
    ' Skip empty lines during processing (they separate groupAccts)
    If Len(Trim(currentLine)) < 2 Then
        ' If we were in a group, process it before moving to next
        If memberCount > 0 Then
            Call foutGroup(currentGroup, accountNumber, memberCount)
            memberCount = 0
			ReDim currentGroup(0)
            accountNumber = ""
        End If

    Else
        ' Check for account number line (but not the "Account Number" header)
        If IsAccountLine(currentLine) Then

            accountNumber = Trim(currentLine)

        Else
            ' This is a Group member name
            If memberCount > 0 Then
                ReDim Preserve currentGroup(memberCount)
            End If
            currentGroup(memberCount) = currentLine
            memberCount = memberCount + 1
        End If
    End If
End Sub


Sub ProcessLastGroup()
    ' Process the last Group if file doesn't end with blank line
    If memberCount > 0 Then
        Call foutGroup(currentGroup, accountNumber, memberCount)
    End If
End Sub
 
' Subroutine to write a Group to the output file
Sub foutGroup(currentGroup, accountNumber, memberCount)
    Dim j, paddedName, nameLen
    
    For j = 0 To memberCount - 1
        If currentGroup(j) <> "" Then
            nameLen = Len(currentGroup(j))
            If nameLen < 138 Then
                paddedName = currentGroup(j) & String(138 - nameLen, " ")
            Else
                paddedName = currentGroup(j)
            End If
            
            fout.WriteLine paddedName & accountNumber
        End If
    Next
End Sub

Function IsAccountLine(line)
    Dim re
    Set re = New RegExp
    re.Pattern = "^\s*\d{2,}\s*(,\s*\d{3,}\s*)*$"
    re.IgnoreCase = True
    re.Global = False

    IsAccountLine = re.Test(line)
End Function
