Option Explicit

Dim inRecord
Dim currRecord

Sub StartDoc()
    inRecord = False
    currRecord = ""
    
    fout.WriteLine "LoanType         Rate      NbrLoans  PctTotalLoans   Dollar Value PctgofDollarValue     Weighted Avg     Average Balance"
End Sub

Sub ProcessLine()
    Dim line, first8, AcctKey, padLen
    
    line = srec
    
    '---------------------------------------------------------
    ' Skip report header/boilerplate lines
    '---------------------------------------------------------
    If InStr(line, "Run Date:") > 0 _
       Or InStr(line, "Type                   of Loans") > 0 _
       Or InStr(line, "---") > 0 _
	   Or InStr(line, "Carolina Trust Federal Credit Union ") > 0 _
       Or InStr(line, "      Loan      Rate           Number  Percent of") > 0 Then
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
    ' Look for account number pattern in first 7 characters
    '---------------------------------------------------------
    first8 = Left(line, 8)
    
    ' Check if first 8 chars contain mostly spaces followed by digits
    ' This indicates an account number field
    If IsAccountNumberLine(first8) Then
        ' Start of a new record
        
        ' If we were already in a record, flush the previous one
        If inRecord And Len(currRecord) > 0 Then
            fout.WriteLine currRecord
        End If
        
        currRecord = line
        inRecord = True
        
    Else
        ' Continuation line: append to current record
        If inRecord and len(line) < 100 Then
            currRecord = currRecord & line
        End If
    End If
End Sub

Sub CloseDoc()
    ' Flush last record if file doesn't end with a blank line
    If inRecord And Len(currRecord) > 0 Then
        fout.WriteLine currRecord
    End If
End Sub

'==================== Helpers ====================

Function IsAccountNumberLine(txt)
    Dim i, spaceCount, digitCount, char
    
    ' Reset counts
    spaceCount = 0
    digitCount = 0
    
    ' Analyze first 12 characters
    For i = 1 To Len(txt)
        char = Mid(txt, i, 1)
        If char = " " Then
            spaceCount = spaceCount + 1
        ElseIf IsNumeric(char) Then
            digitCount = digitCount + 1
        End If
    Next 
    
    ' Consider it an account number line if:
    ' - Has at least 1 digit
    ' - Has significant leading spaces (at least 6 spaces in first 12 chars)
    ' - Digits appear after spaces
    If digitCount >= 1 And spaceCount >= 5 Then
        IsAccountNumberLine = True
    Else
        IsAccountNumberLine = False
    End If
End Function

Function IsAllDigits(txt)
    Dim cleaned, re
    
    cleaned = Trim(txt)
    If Len(cleaned) = 0 Then
        IsAllDigits = False
        Exit Function
    End If
    
    Set re = New RegExp
    re.Pattern = "^\d+$"
    re.Global = False
    re.IgnoreCase = True
    
    IsAllDigits = re.Test(cleaned)
End Function