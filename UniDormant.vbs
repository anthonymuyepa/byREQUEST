Dim curAcctLine, haveAcct, expectingFirstTxn, reachedEnd
Dim depositsSection()  ' dynamic array to store the end section
Dim depositsCount

Sub startDoc
	srec = getLine
	fout.WriteLine srec
	srec = getLine
	fout.WriteLine srec
    curAcctLine = "": haveAcct = False: expectingFirstTxn = False
    reachedEnd = False
    ReDim depositsSection(0)
    depositsCount = 0
    fout.WriteLine "Account    Name                S/L/X            S ID   Tran     Amount    Principal  Interest     Fees  New Balance  Last Trn   Effect  Usr  Ovr  Sec    Seq"
End Sub

Sub ProcessLine
    Dim line: line = srec

    ' If we are already past "Transaction   Deposits", store lines for later
    If reachedEnd Then
        StoreDepositsLine line
        Exit Sub
    End If

    ' Check for start of deposits section
    If InStr(line, "Transaction   Deposits") = 1 Then
        reachedEnd = True
        StoreDepositsLine line
        Exit Sub
    End If

    If Trim(line) = "" Or InStr(line, "----") > 0 Then Exit Sub

    ' ===== If we are in header-only mode, handle continuations =====
    If haveAcct Then
        If expectingFirstTxn Then
            fout.WriteLine curAcctLine & line
            expectingFirstTxn = False
            Exit Sub
        Else
            If Not IsStartWithDigits(line) Then
                fout.WriteLine curAcctLine & line
                Exit Sub
            Else
                haveAcct = False
            End If
        End If
    End If

    ' ===== New account or one-line record =====
    If IsStartWithDigits(line) Then
        Dim acctName25: acctName25 = Left(line & Space(24), 24)
        If Len(RTrim(line)) <= 25 Then
            curAcctLine = acctName25
            haveAcct = True
            expectingFirstTxn = True
        Else
            Dim txnPart: txnPart = Mid(line, 25)
            fout.WriteLine acctName25 & Space(24) & txnPart
            haveAcct = False
            expectingFirstTxn = False
        End If
    End If
End Sub

Sub CloseDoc
    Dim i
    For i = 1 To 5
        fout.WriteLine 
    Next

    For i = 1 To depositsCount
        fout.WriteLine depositsSection(i)
    Next
End Sub

' ===== Helpers =====

Sub StoreDepositsLine(line)
    depositsCount = depositsCount + 1
    ReDim Preserve depositsSection(depositsCount)
    depositsSection(depositsCount) = line
End Sub

Function IsStartWithDigits(line)
    Dim i, ch
    i = 1
    Do While i <= Len(line) And Mid(line, i, 1) = " "
        i = i + 1
    Loop
    If i > Len(line) Then
        IsStartWithDigits = False
    Else
        ch = Mid(line, i, 1)
        IsStartWithDigits = (ch >= "0" And ch <= "9")
    End If
End Function
