Option Explicit

Dim gBuffer
Dim gContCount
Dim gReportYear

Sub StartDoc()
    gBuffer = ""
    gContCount = 0
    gReportYear = ""

    Fout.WriteLine Pad("CardNumber",18)  & _
                   Pad("LDate",12)  & _
                   Pad("LTime",10)  & _
                   Pad("TransID",10)  & _
                   Pad("Terminal",10)  & _
                   Pad("Logo",10)  & _
                   Pad("SeqNo",12)  & _
                   Pad("TranDesc",12)  & _
                   Pad("MsgType",10)  & _
                   Pad("RespCode",10)  & _
                   Pad("Amount",15)  & _
                   Pad("NetworkName",12)  & _
                   Pad("SettlementDate",15)  & _
                   Pad("chFee",10)  & _
                   Pad("Surcharge",10)  & _
                   Pad("FromAccount",15)  & _
                   Pad("TermAddr",45)  & _
                   Pad("RPI",5)  & _
                   Pad("ED",5)  & _
                   Pad("CV",5)  & _
                   Pad("CV2",5)  & _
                   Pad("CV3",5)  & _
                   Pad("AV",5)  & _
                   Pad("CAVV",5)  & _
                   Pad("ECI",5)  & _
                   Pad("PosEMode",9)  & _
                   Pad("CardPresent",12)  & _
                   Pad("MOBL",8)  & _
                   Pad("NSCR",8)  & _
                   Pad("NRSN",8)  & _
                   Pad("AuthID",10)  & _
                   "MerchID"
End Sub

Sub ProcessLine()
    Dim s, firstChar, first5
    Dim lineForYear

    s = srec

    If Trim(s) = "" Or Len(s) < 2 Then Exit Sub
    If InStr(s, "CLIENT:") > 0 Then Exit Sub
    If InStr(s, " TRANSACTION DATA") > 0 Then Exit Sub
    If InStr(s, "--------------") > 0 Then Exit Sub
    If InStr(s, "CARDHOLDER N") > 0 Then Exit Sub
    If InStr(s, "                                  EFT") > 0 Then Exit Sub
    If InStr(s, "   PAGE ") > 0 Then Exit Sub
    
    ' Extract year from line containing "FOR"
    If gReportYear = "" Then
        If InStr(s, "FOR") > 0 Then
            lineForYear = s
            gReportYear = ExtractYearFromLine(lineForYear)
        End If
    End If

    ' Skip lines beginning with date pattern like 08/06
    first5 = Left(Trim(s), 5)
    If InStr(first5, "/") > 0 Then Exit Sub

    ' Pad every physical line to 200 before concatenating
    s = PadLine(s, 200)

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
        ' Continuation lines - just append
        If gBuffer <> "" Then
            gBuffer = gBuffer & s
            gContCount = gContCount + 1
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
    
    parts = Split(line, " ")
    
    For i = 0 To UBound(parts)
        word = Trim(parts(i))
        If IsNumeric(word) And Len(word) = 4 Then
            possibleYear = CInt(word)
            If possibleYear >= 1900 And possibleYear <= 2099 Then
                ExtractYearFromLine = CStr(possibleYear)
                Exit Function
            End If
        End If
    Next
End Function

Sub WriteCurrentRecord()
    Dim CardNumber, LocalDate, LocalTime, TransID, Terminal, Logo, Zone, SequenceNo
    Dim TranDesc, MsgType, RespCode, Amount, NetworkName
    Dim SettlementDate, chFee, Surcharge, FromAccount, TermAddr, RPI
    Dim ED, CV, CV2, CV3, AV, CAVV, ECI, PosEMode, CardPresent
    Dim MOBL, NSCR, NRSN, AuthID, MerchID
    Dim p, temp, lineContent

    ' ========== MAIN LINE FIELDS (Fixed positions) ==========
    CardNumber   = Trim(Mid(gBuffer, 1, 16))
    LocalDate    = Trim(Mid(gBuffer, 24, 5))
    LocalDate    = ExpandLocalDate(LocalDate)
    LocalTime    = Trim(Mid(gBuffer, 30, 5))
    TransID      = Trim(Mid(gBuffer, 36, 8))
    Terminal     = Trim(Mid(gBuffer, 45, 8))
    Logo         = Trim(Mid(gBuffer, 54, 4))
    Zone         = Trim(Mid(gBuffer, 59, 7))		
    SequenceNo   = Trim(Mid(gBuffer, 68, 6))
    TranDesc     = Trim(Mid(gBuffer, 75, 7))
    MsgType      = Trim(Mid(gBuffer, 83, 2))
    RespCode     = Trim(Mid(gBuffer, 87, 3))
    Amount       = Trim(Mid(gBuffer, 91, 14))
    NetworkName  = Trim(Mid(gBuffer, 107, 6))
    SettlementDate = ExpandLocalDate(Trim(Mid(gBuffer, 114, 5)))
    chFee        = Trim(Mid(gBuffer, 120, 5))
    Surcharge    = Trim(Mid(gBuffer, 126, 6))

    ' ========== LINE 1: FROM ACCOUNT and TERM ADDR ==========
    FromAccount = ""
    p = InStr(1, gBuffer, " ACCT: ", vbTextCompare)
    If p > 0 Then
        temp = Mid(gBuffer, p + Len(" ACCT: "), 30)
        FromAccount = Trim(temp)
    End If

    TermAddr = ""
    p = InStr(1, gBuffer, "TERM ADDR:", vbTextCompare)
    If p > 0 Then
        temp = Mid(gBuffer, p + Len("TERM ADDR:"))
        If InStr(temp, "RPI=") > 0 Then
            temp = Left(temp, InStr(temp, "RPI=") - 1)
        End If
        TermAddr = Trim(temp)
    End If

    RPI = ""
    p = InStr(1, gBuffer, "RPI=", vbTextCompare)
    If p > 0 Then
        RPI = Trim(Mid(gBuffer, p + 4, 2))
    End If

    ' ========== MERCH ID EXTRACTION ==========
    MerchID = ""
    p = InStr(1, gBuffer, "MERCH ID:", vbTextCompare)
    If p > 0 Then
        MerchID = Trim(Mid(gBuffer, p + Len("MERCH ID:"), 18))
    End If

    ' ========== AUTH ID EXTRACTION ==========
    AuthID = ""
    p = InStr(1, gBuffer, "AUTH ID:", vbTextCompare)
    If p > 0 Then
        AuthID = Trim(Mid(gBuffer, p + Len("AUTH ID:"), 8))
    End If

    ' ========== EXTENDED FIELDS (ED, CV, CV2, CV3, AV, CAVV, ECI, PosEMode, CardPresent, MOBL, NSCR, NRSN) ==========
    lineContent = gBuffer
    
    ' Initialize all fields
    ED = "" : CV = "" : CV2 = "" : CV3 = "" : AV = "" : CAVV = ""
    ECI = "" : PosEMode = "" : CardPresent = "" : MOBL = "" : NSCR = "" : NRSN = ""
    
    ' Look for patterns like: "ED=xxx CV=xxx CV2=xxx CV3=xxx AV=xxx CAVV=xxx"
    p = InStr(1, lineContent, "ED=", vbTextCompare)
    If p > 0 Then
        temp = Mid(lineContent, p)
        
        ED = ExtractValueAdvanced(temp, "ED=", " ", "CV=")
        CV = ExtractValueAdvanced(temp, "CV=", " ", "CV2=")
        CV2 = ExtractValueAdvanced(temp, "CV2=", " ", "CV3=")
        CV3 = ExtractValueAdvanced(temp, "CV3=", " ", "AV=")
        AV = ExtractValueAdvanced(temp, "AV=", " ", "CAVV=")
        CAVV = ExtractValueAdvanced(temp, "CAVV=", " ", "ECI=")
        ECI = ExtractValueAdvanced(temp, "ECI=", " ", "POS")
    End If
    
    ' POS EMODE EXTRACTION
    p = InStr(1, lineContent, "POS EMODE=", vbTextCompare)
    If p > 0 Then
        temp = Mid(lineContent, p)
        PosEMode = ExtractValueAdvanced(temp, "POS EMODE=", " ", " ")
        If Len(PosEMode) > 2 Then
            PosEMode = Left(PosEMode, 4)
        End If
    End If
    
    ' CARD PRESENT STATUS
    If InStr(1, lineContent, "NOT PRESENT", vbTextCompare) > 0 Then
        CardPresent = "NOT PRESENT"
    ElseIf InStr(1, lineContent, "PRESENT", vbTextCompare) > 0 Then
        CardPresent = "PRESENT"
    End If
    
    ' MOBL EXTRACTION
    p = InStr(1, lineContent, "MOBL=", vbTextCompare)
    If p > 0 Then
        temp = Mid(lineContent, p)
        MOBL = ExtractValueAdvanced(temp, "MOBL=", " ", " ")
        MOBL = Left(MOBL,2)
    End If
    
    ' Handle MOBLSNA without equals sign
    If MOBL = "" And InStr(1, lineContent, "MOBLSNA", vbTextCompare) > 0 Then
        MOBL = "SNA"
    End If
  
    ' N-SCR EXTRACTION
    p = InStr(1, lineContent, "N-SCR=", vbTextCompare)
    If p > 0 Then
        temp = Mid(lineContent, p)
        NSCR = ExtractValueAdvanced(temp, "N-SCR=", " ", " ")
        NSCR = Left(NSCR, 2)
    End If
    
    ' N-RSN EXTRACTION
    p = InStr(1, lineContent, "N-RSN=", vbTextCompare)
    If p > 0 Then
        temp = Mid(lineContent, p)
        NRSN = ExtractValueAdvanced(temp, "N-RSN=", " ", " ")
        
        If NRSN = "" Then
            If InStr(1, lineContent, "N-RSNSLR", vbTextCompare) > 0 Then
                NRSN = "SLR"
            ElseIf InStr(1, lineContent, "N-RSNSNA", vbTextCompare) > 0 Then
                NRSN = "SNA"
            End If
        End If
    End If

    ' Clean up values
    If CardNumber = "" Then CardNumber = "0"
    If Amount = "" Then Amount = "0"
    
	
	'fout.Writeline gBuffer
    ' Write to output matching header exactly
    fout.WriteLine _
        Pad(CardNumber,18)  & _
        Pad(LocalDate,12)  & _
        Pad(LocalTime,10)  & _
        Pad(TransID,10)  & _
        Pad(Terminal,10)  & _
        Pad(Logo,10)  & _
        Pad(SequenceNo,12)  & _
        Pad(TranDesc,12)  & _
        Pad(MsgType,10)  & _
        Pad(RespCode,10)  & _
        Pad(Amount,15)  & _
        Pad(NetworkName,12)  & _
        Pad(SettlementDate,15)  & _
        Pad(chFee,10)  & _
        Pad(Surcharge,10)  & _
        Pad(FromAccount,15)  & _
        Pad(TermAddr,45)  & _
        Pad(RPI,5)  & _
        Pad(ED,5)  & _
        Pad(CV,5)  & _
        Pad(CV2,5)  & _
        Pad(CV3,5)  & _
        Pad(AV,5)  & _
        Pad(CAVV,5)  & _
        Pad(ECI,5)  & _
        Pad(PosEMode,9)  & _
        Pad(CardPresent,12)  & _
        Pad(MOBL,8)  & _
        Pad(NSCR,8)  & _
        Pad(NRSN,8)  & _
        Pad(AuthID,10)  & _
        Pad(MerchID,30)

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

Function ExtractValueAdvanced(ByVal text, ByVal startTag, ByVal endDelimiter, ByVal stopAtTag)
    Dim startPos, endPos, value, stopPos, delimPos
    
    ExtractValueAdvanced = ""
    
    startPos = InStr(1, text, startTag, vbTextCompare)
    If startPos = 0 Then Exit Function
    
    startPos = startPos + Len(startTag)
    endPos = Len(text) + 1
    
    If stopAtTag <> "" And stopAtTag <> " " Then
        stopPos = InStr(startPos, text, stopAtTag, vbTextCompare)
        If stopPos > 0 And stopPos < endPos Then
            endPos = stopPos
        End If
    End If
    
    If endDelimiter <> "" And endDelimiter <> " " Then
        delimPos = InStr(startPos, text, endDelimiter, vbTextCompare)
        If delimPos > 0 And delimPos < endPos Then
            endPos = delimPos
        End If
    End If
    
    value = Trim(Mid(text, startPos, endPos - startPos))
    value = CleanExtendedValue(value)
    
    ExtractValueAdvanced = value
End Function

Function CleanMOBLValue(ByVal value)
    Dim result
    result = Trim(UCase(value))
    
    result = Replace(result, vbCrLf, "")
    result = Replace(result, vbTab, "")
    
    Select Case result
        Case "NA", "N/A", "N-A": result = "NA"
        Case "NR", "N/R", "N-R": result = "NR"
        Case "SNA", "S-N-A": result = "SNA"
        Case "SLR", "S-L-R": result = "SLR"
    End Select
    
    CleanMOBLValue = result
End Function

Function CleanNSCRValue(ByVal value)
    Dim result
    result = Trim(UCase(value))
    
    result = Replace(result, vbCrLf, "")
    result = Replace(result, vbTab, "")
    
    Select Case result
        Case "NA", "N/A", "N-A": result = "NA"
        Case "SNA", "S-N-A": result = "SNA"
    End Select
    
    CleanNSCRValue = result
End Function

Function CleanExtendedValue(ByVal value)
    Dim result, ch
    
    result = Trim(CStr(value))
    
    Do While Len(result) > 0
        ch = Right(result, 1)
        If ch = " " Or ch = vbCr Or ch = vbLf Or ch = vbTab Or ch = ";" Or ch = "," Then
            result = Left(result, Len(result) - 1)
        Else
            Exit Do
        End If
    Loop
    
    Do While Len(result) > 0
        ch = Left(result, 1)
        If ch = " " Or ch = vbCr Or ch = vbLf Or ch = vbTab Or ch = ";" Or ch = "," Then
            result = Mid(result, 2)
        Else
            Exit Do
        End If
    Loop
    
    CleanExtendedValue = result
End Function