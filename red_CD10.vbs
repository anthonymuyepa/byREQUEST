Option Explicit

Dim gBuffer
Dim gContCount
Dim gReportYear  ' Global variable to store the extracted year

Sub StartDoc()
    gBuffer = ""
    gContCount = 0
    gReportYear = ""  ' Initialize as empty

    Fout.WriteLine Pad("CardNumber",18) & _
                   Pad("LDate",12) & _
                   Pad("LTime",10) & _
                   Pad("Logo",10) & _
                   Pad("SeqNo",10) & _
                   Pad("TranDes",10) & _
                   Pad("Rsp",10) & _
                   Pad("Amount",15) & _
                   Pad("NetName",10) & _
                   Pad("FromAccount",15) & _
                   Pad("TermAddr",53) & _
                   Pad("MerchantName",20) & _
                   Pad("RP",5) & _
                   "MerchID"
End Sub

Sub ProcessLine()
    Dim s, firstChar, first5
    Dim lineForYear

    s = srec

    If Trim(s) = "" Then Exit Sub
    If InStr(s, "CLIENT:") > 0 Then Exit Sub
    If InStr(s, " TRANSACTION DATA") > 0 Then Exit Sub
    If InStr(s, "--------------") > 0 Then Exit Sub
    If InStr(s, "CARDHOLDER N") > 0 Then Exit Sub
    
    ' Extract year from line containing "FOR" (e.g., "FOR AUG 06, 2025")
    If gReportYear = "" Then
        If InStr(s, "FOR") > 0 Then
            lineForYear = s
            gReportYear = ExtractYearFromLine(lineForYear)
            If gReportYear <> "" Then
                ' Optional: Log the extracted year
                ' fout.WriteLine "-- Extracted Year: " & gReportYear & " --"
            End If
        End If
    End If

    ' Skip lines beginning with date pattern like 08/06
    first5 = Left(Trim(s), 5)
    If InStr(first5, "/") > 0 Then Exit Sub

    ' Pad every physical line to 135 before concatenating
    s = PadLine(s, 135)

    firstChar = Left(LTrim(s), 1)

    ' New logical record starts when first non-space character is numeric
    If IsNumeric(firstChar) Then

        ' Flush prior buffered record before starting a new one
        If gBuffer <> "" Then
            WriteCurrentRecord
        End If

        gBuffer = s
        gContCount = 0

    Else
        ' Continuation lines
        If gBuffer <> "" Then
            gBuffer = gBuffer & s
            gContCount = gContCount + 1

            ' CD10 record expected to have 3 continuation lines
            If gContCount >= 3 Then
                WriteCurrentRecord
            End If
        End If
    End If
End Sub

Sub CloseDoc()
    If gBuffer <> "" Then
        WriteCurrentRecord
    End If
End Sub

Function ExtractYearFromLine(ByVal line)
    Dim parts, i, word, possibleYear
    
    ExtractYearFromLine = ""
    
    ' Split the line into words
    parts = Split(line, " ")
    
    ' Look for a 4-digit number that could be a year
    For i = 0 To UBound(parts)
        word = Trim(parts(i))
        If IsNumeric(word) And Len(word) = 4 Then
            possibleYear = CInt(word)
            ' Validate year is reasonable (1900-2099)
            If possibleYear >= 1900 And possibleYear <= 2099 Then
                ExtractYearFromLine = CStr(possibleYear)
                Exit Function
            End If
        End If
    Next
End Function

Sub WriteCurrentRecord()
    Dim CardNumber, LocalDate, LocalTime, Logo, SequenceNo, TranDesc
    Dim RespCode, Amount, NetworkName, FromAccount, TermAddr
    Dim MerchantName, RPI, MerchID
    Dim p

    CardNumber   = Trim(Mid(gBuffer, 1, 16))
    LocalDate    = Trim(Mid(gBuffer, 24, 5))
    LocalDate    = ExpandLocalDate(LocalDate)
    LocalTime    = Trim(Mid(gBuffer, 30, 5))
    Logo         = Trim(Mid(gBuffer, 54, 4))
    SequenceNo   = Trim(Mid(gBuffer, 68, 6))
    TranDesc     = Trim(Mid(gBuffer, 75, 7))
    RespCode     = Trim(Mid(gBuffer, 87, 3))
    Amount       = Trim(Mid(gBuffer, 95, 11))
    NetworkName  = Trim(Mid(gBuffer, 107, 7))
    MerchantName = Trim(Mid(gBuffer, 226, 17))

    FromAccount = ""
    p = InStr(1, gBuffer, "ACCT:", vbTextCompare)
    If p > 0 Then
        FromAccount = Trim(Mid(gBuffer, p + Len("ACCT:"), 14))
    End If

    TermAddr = ""
    p = InStr(1, gBuffer, "TERM ADDR:", vbTextCompare)
    If p > 0 Then
        TermAddr = Trim(Mid(gBuffer, p + Len("TERM ADDR:"), 53))
    End If

    RPI = ""
    p = InStr(1, gBuffer, "RPI=", vbTextCompare)
    If p > 0 Then
        RPI = Trim(Mid(gBuffer, p + 4, 2))
    End If

    MerchID = ""
    p = InStr(1, gBuffer, "MERCH ID:", vbTextCompare)
    If p > 0 Then
        MerchID = Trim(Mid(gBuffer, p + Len("MERCH ID:"), 18))
    End If

    Fout.WriteLine _
        Pad(CardNumber,18) & _
        Pad(LocalDate,12) & _
        Pad(LocalTime,10) & _
        Pad(Logo,10) & _
        Pad(SequenceNo,10) & _
        Pad(TranDesc,10) & _
        Pad(RespCode,10) & _
        Pad(Amount,15) & _
        Pad(NetworkName,10) & _
        Pad(FromAccount,15) & _
        Pad(TermAddr,53) & _
        Pad(MerchantName,20) & _
        Pad(RPI,5) & _
        Trim(MerchID)

    gBuffer = ""
    gContCount = 0
End Sub

Function ExpandLocalDate(ByVal s)
    Dim parts, mm, dd, yyyy

    s = Trim(CStr(s))
    ExpandLocalDate = s

    If Len(s) = 0 Then Exit Function

    parts = Split(s, "/")
    If UBound(parts) <> 1 Then Exit Function

    If IsNumeric(parts(0)) And IsNumeric(parts(1)) Then
        mm = Right("0" & CStr(CInt(parts(0))), 2)
        dd = Right("0" & CStr(CInt(parts(1))), 2)
        
        ' Use the extracted year from the file, or fall back to current year
        If gReportYear <> "" And IsNumeric(gReportYear) Then
            yyyy = gReportYear
        Else
            yyyy = CStr(Year(Date))
        End If
        
        ExpandLocalDate = mm & "/" & dd & "/" & yyyy
    End If
End Function

Function Pad(ByVal s, ByVal n)
    s = CStr(s)
    If Len(s) < n Then
        Pad = s & Space(n - Len(s))
    Else
        Pad = Left(s, n)
    End If
End Function

Function PadLine(ByVal s, ByVal n)
    s = RTrim(CStr(s))
    If Len(s) < n Then
        PadLine = s & Space(n - Len(s))
    Else
        PadLine = Left(s, n)
    End If
End Function