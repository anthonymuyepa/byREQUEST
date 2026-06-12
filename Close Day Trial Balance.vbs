Option Explicit

Dim inRecord
Dim currRecord
Dim srec
Dim fout

Sub StartDoc
    inRecord = False
    currRecord = ""
    
    Dim i
    For i = 1 To 1
        srec = getLine
        If srec <> "" Then
           ' fout.WriteLine srec
        End If
    Next
	fout.WriteLine "Type        Nbr  Name                             Open Count   Close Count               Balance"
End Sub

Sub ProcessLine
    Dim line, len99, AcctKey, padLen
    Dim logicalBlank, AcctType, AcctNo, formattedAcctType

    line = srec

    '---------------------------------------------------------
    ' Skip report header/boilerplate lines
    '---------------------------------------------------------
    If InStr(line, "Run Date:") > 0 _
       Or InStr(line, "Post Date:") > 0 _
       Or InStr(line, "Page:") > 0 _
       Or InStr(line, "Date:") > 0 _
       Or InStr(line, " Totals ") > 0 _
       Or InStr(line, "Close Day Trial Balance") > 0 _
       Or InStr(line, "Type             Name") > 0 Then
        Exit Sub
    End If

    '---------------------------------------------------------
    ' Blank line = end of record
    '---------------------------------------------------------
    logicalBlank = Trim(Replace(line, Chr(12), ""))

    If Len(logicalBlank) < 2 Then
        ' Blank line or FF = end of current record
        If inRecord And Len(currRecord) > 0 Then
            ProcessCurrentRecord
            currRecord = ""
            inRecord = False
        End If
        Exit Sub
    End If

    '---------------------------------------------------------
    ' Detect start of a new record:
    '---------------------------------------------------------
    len99 = Len(Trim(line))

    If len99 > 95 Then
        ' Start of a new record

        ' If we were already in a record, flush the previous one
        If inRecord Then
            ProcessCurrentRecord
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

Sub ProcessCurrentRecord
    Dim AcctType, AcctNo, formattedAcctType
    Dim totalWidth, paddingNeeded
    
    If Len(currRecord) >= 16 Then
        AcctType = Left(currRecord, 16)
		AcctType = Trim(AcctType)
        ' Extract last 4 digits from the first 16 characters
        AcctNo = Right(AcctType, 4)
		
		'MsgBox "AcctType: " & AcctType & " " & AcctNo
        AcctType = Left(AcctType, Len(AcctType) - 4)
        AcctType = Trim(AcctType)
        
        ' Fixed column alignment 
        totalWidth = 16  ' Total width for account type + number
        paddingNeeded = totalWidth - Len(AcctType) - Len(AcctNo)
        
        If paddingNeeded > 0 Then
            formattedAcctType = AcctType & Space(paddingNeeded) & AcctNo
        Else
            formattedAcctType = AcctType & " " & AcctNo
        End If
        
        ' Output the formatted line
        If Len(currRecord) > 16 Then
            fout.WriteLine formattedAcctType & Mid(currRecord, 17)
        Else
            fout.WriteLine formattedAcctType
        End If
    Else
        ' Handle records shorter than 16 characters
        fout.WriteLine currRecord
    End If
End Sub

Sub CloseDoc
    ' Flush last record if file doesn't end with a blank line
    If inRecord And Len(currRecord) > 0 Then
        ProcessCurrentRecord
    End If
End Sub
