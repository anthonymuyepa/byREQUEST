Option Explicit

'============================================================
' ACH List Parser
'============================================================

Dim gReportName, gReportDate, gReportTime, gReportSeq

Dim gBatchSCC, gCompanyName, gCompanyDiscretion, gCompanyID
Dim gECC, gEntryDescription, gDescDate, gEntryDate
Dim gSettlementDays, gBatchStatus, gOriginRT, gBatchNumber

Dim gDetailHeaderSeen, gBatchCount, gDetailCount
Dim gPendingRecordType

Dim gFileRecordType, gFileHeaderType, gFilePriority
Dim gFileReceivingRT, gFileSendingRT, gFileDate, gFileTime
Dim gFileModifier, gFileRecordSize, gFileBlockFactor
Dim gFileFormat, gFileReceivingName, gFileSendingName, gFileReference

Dim gPendingDetail
Dim gPendingDetailRecord, gPendingTranCode, gPendingDestRT
Dim gPendingIndividualAcct, gPendingAmount, gPendingIndividualID
Dim gPendingIndividualName, gPendingDisc, gPendingAddenda
Dim gPendingTraceNumber

Dim gDetailLayoutType
Dim gAddendaLayoutType

Dim OUT_HEADER

OUT_HEADER = _
    PadRight("Report", 12) & _
    PadRight("RptDate", 12) & _
    PadRight("RptTime", 10) & _
    PadRight("RptSeq", 10) & _
    PadRight("FileHdr", 8) & _
    PadRight("Pri", 5) & _
    PadRight("RecvRT", 14) & _
    PadRight("SendRT", 12) & _
    PadRight("FileDate", 10) & _
    PadRight("FileTime", 10) & _
    PadRight("Mod", 5) & _
    PadRight("Recsize", 8) & _
    PadRight("Blk", 5) & _
    PadRight("Format", 7) & _
    PadRight("ReceivingName", 25) & _
    PadRight("SendingName", 25) & _
    PadRight("Reference", 11) & _
    PadRight("SCC", 5) & _
    PadRight("CompanyName", 26) & _
    PadRight("CompanyDisc", 22) & _
    PadRight("CompanyID", 14) & _
    PadRight("ECC", 6) & _
    PadRight("Desc", 14) & _
    PadRight("DescDate", 10) & _
    PadRight("EntryDate", 10) & _
    PadRight("SettleDays", 12) & _
    PadRight("Status", 8) & _
    PadRight("OriginRT", 10) & _
    PadRight("Batch", 10) & _
    PadRight("Detail", 8) & _
    PadRight("Tran", 6) & _
    PadRight("DestRT", 10) & _
    PadRight("IndividualAcct", 16) & _
    PadRight("Amount", 13) & _
    PadRight("IndividualID", 18) & _
    PadRight("IndividualName", 25) & _
    PadRight("Disc", 6) & _
    PadRight("Addenda", 8) & _
    PadRight("TraceNumber", 16) & _
    PadRight("Addenda Type", 15) & _
    PadRight("PaymentInfo", 125) & _
    PadRight("AddSeq", 14) & _
    PadRight("DtlSeq", 14)


Sub StartDoc()

    ClearAllContext

    gBatchCount = 0
    gDetailCount = 0
    gDetailHeaderSeen = False
    gPendingDetail = False
    gPendingRecordType = ""
    gDetailLayoutType = ""
    gAddendaLayoutType = ""

    srec = getLine

    If Len(Trim(srec & "")) > 0 Then
        ParseReportTitleLine srec
    End If

    fout.WriteLine OUT_HEADER


End Sub


Sub ProcessLine()

    Dim line
    line = srec

    If IsBlankLine(line) Then Exit Sub
    If IsPageHeaderOrNoise(line) Then Exit Sub

    If gPendingRecordType <> "" Then

        Select Case gPendingRecordType
            Case "FILE"
                ParseFileHeaderDataLine line

            Case "BATCH"
                ParseBatchHeaderDataLine line
                gBatchCount = gBatchCount + 1
                gDetailHeaderSeen = False

            Case "BATCHCTRL"
                ParseBatchControlDataLine line
        End Select

        gPendingRecordType = ""
        Exit Sub
    End If

    If gPendingDetail Then

        If StartsWithText(line, "Addenda Type") Then
            gAddendaLayoutType = GetAddendaLayoutType(line)
            Exit Sub
        End If

        ParseAndWriteAddendaLine line
        gDetailCount = gDetailCount + 1
        Exit Sub

    End If

    If StartsWithText(line, "File Hdr") Then
        gPendingRecordType = "FILE"
        Exit Sub
    End If

    If StartsWithText(line, "Batch Hdr") Then
        gPendingRecordType = "BATCH"
        Exit Sub
    End If

    If StartsWithText(line, "Detail") Then
        gDetailHeaderSeen = True
        gDetailLayoutType = GetDetailLayoutType(line)
        Exit Sub
    End If

    If StartsWithText(line, "Batch Ctrl") Then
        gPendingRecordType = "BATCHCTRL"
        gDetailHeaderSeen = False
        Exit Sub
    End If

    If gDetailHeaderSeen And IsDetailDataLine(line) Then

        ParseAndWriteDetailLine line

        If Not gPendingDetail Then
            gDetailCount = gDetailCount + 1
        End If

        Exit Sub
    End If

End Sub


Sub CloseDoc()

    If gPendingDetail Then
        WriteDetailRow gPendingDetailRecord, gPendingTranCode, gPendingDestRT, _
                       gPendingIndividualAcct, gPendingAmount, gPendingIndividualID, _
                       gPendingIndividualName, gPendingDisc, gPendingAddenda, _
                       gPendingTraceNumber, "", "", "", ""

        gPendingDetail = False
        gDetailCount = gDetailCount + 1
    End If

    ' fout.WriteLine ""
    ' fout.WriteLine "Batches Parsed: " & gBatchCount
    ' fout.WriteLine "Details Parsed: " & gDetailCount

End Sub


'============================================================
' Layout Detection
'============================================================

Function GetDetailLayoutType(ByVal line)

    If InStr(1, line, "Ident Number", vbTextCompare) > 0 _
       And InStr(1, line, "Pymt Addenda", vbTextCompare) > 0 Then
        GetDetailLayoutType = "PAYMENT"
        Exit Function
    End If

    If InStr(1, line, "Individual   ID", vbTextCompare) > 0 _
       And InStr(1, line, "Disc Addenda", vbTextCompare) > 0 Then
        GetDetailLayoutType = "INDIVIDUAL"
        Exit Function
    End If

    If InStr(1, line, "Identification", vbTextCompare) > 0 _
       And InStr(1, line, "Code Addenda", vbTextCompare) > 0 Then
        GetDetailLayoutType = "TERMINAL"
        Exit Function
    End If

    GetDetailLayoutType = "INDIVIDUAL"

End Function


Function GetAddendaLayoutType(ByVal line)

    If InStr(1, line, "Payment Related Info", vbTextCompare) > 0 Then
        GetAddendaLayoutType = "PAYMENTINFO"
        Exit Function
    End If

    If InStr(1, line, "Ref1", vbTextCompare) > 0 _
       And InStr(1, line, "Terminal", vbTextCompare) > 0 Then
        GetAddendaLayoutType = "TERMINALINFO"
        Exit Function
    End If

    GetAddendaLayoutType = "PAYMENTINFO"

End Function


'============================================================
' Main Parsers
'============================================================

Sub ParseReportTitleLine(ByVal line)

    gReportName = "ACH List"
    gReportDate = ""
    gReportTime = ""
    gReportSeq = ""

    Dim pAt, pSeq, pPage

    pAt = InStr(1, line, " at ", vbTextCompare)
    pSeq = InStr(1, line, " Seq ", vbTextCompare)
    pPage = InStr(1, line, " Page ", vbTextCompare)

    If pAt > 0 Then
        gReportDate = Trim(Mid(line, pAt - 8, 8))
    End If

    If pAt > 0 Then
        gReportTime = Trim(Mid(line, pAt + 13, 5))
    End If

    If pSeq > 0 Then
        If pPage > pSeq Then
            gReportSeq = Trim(Mid(line, pSeq + 5, pPage - (pSeq + 5)))
        Else
            gReportSeq = Trim(Mid(line, pSeq + 5))
        End If
    End If

End Sub


Sub ParseFileHeaderDataLine(ByVal line)

    gFileRecordType     = SafeMidTrim(line, 1, 4)
    gFileHeaderType     = SafeMidTrim(line, 6, 3)
    gFilePriority       = SafeMidTrim(line, 10, 3)
    gFileReceivingRT    = SafeMidTrim(line, 14, 13)
    gFileSendingRT      = SafeMidTrim(line, 28, 11)
    gFileDate           = SafeMidTrim(line, 40, 6)
    gFileTime           = SafeMidTrim(line, 47, 4)
    gFileModifier       = SafeMidTrim(line, 52, 4)
    gFileRecordSize     = SafeMidTrim(line, 56, 7)
    gFileBlockFactor    = SafeMidTrim(line, 64, 4)
    gFileFormat         = SafeMidTrim(line, 68, 6)
    gFileReceivingName  = SafeMidTrim(line, 75, 24)
    gFileSendingName    = SafeMidTrim(line, 99, 24)
    gFileReference      = SafeMidTrim(line, 123, 10)

End Sub


Sub ParseBatchHeaderDataLine(ByVal line)

    gBatchSCC          = SafeMidTrim(line, 11, 3)
    gCompanyName       = SafeMidTrim(line, 15, 16)
    gCompanyDiscretion = SafeMidTrim(line, 32, 20)
    gCompanyID         = SafeMidTrim(line, 53, 10)
    gECC               = SafeMidTrim(line, 64, 3)
    gEntryDescription  = SafeMidTrim(line, 68, 10)
    gDescDate          = SafeMidTrim(line, 79, 9)
    gEntryDate         = SafeMidTrim(line, 89, 7)
    gSettlementDays    = SafeMidTrim(line, 96, 7)
    gBatchStatus       = SafeMidTrim(line, 103, 7)
    gOriginRT          = SafeMidTrim(line, 110, 11)
    gBatchNumber       = SafeMidTrim(line, 121, 8)

End Sub


Sub ParseBatchControlDataLine(ByVal line)
    ' Reserved for validation later.
End Sub


Sub ParseAndWriteDetailLine(ByVal line)

    Select Case gDetailLayoutType
        Case "PAYMENT"
            ParsePaymentDetailLine line

        Case "INDIVIDUAL"
            ParseIndividualDetailLine line

        Case "TERMINAL"
            ParseTerminalDetailLine line

        Case Else
            ParseIndividualDetailLine line
    End Select

End Sub


Sub ParseIndividualDetailLine(ByVal line)

    Dim detailRecord, tranCode, destRT, individualAcct
    Dim amount, individualID, individualName, disc, addenda, traceNumber

    detailRecord   = SafeMidTrim(line, 1, 6)
    tranCode       = SafeMidTrim(line, 11, 2)
    destRT         = SafeMidTrim(line, 14, 9)
    individualAcct = SafeMidTrim(line, 24, 16)
    amount         = SafeMidTrim(line, 42, 11)
    individualID   = SafeMidTrim(line, 54, 15)
    individualName = SafeMidTrim(line, 70, 21)
    disc           = SafeMidTrim(line, 91, 4)
    addenda        = SafeMidTrim(line, 96, 8)
    traceNumber    = SafeMidTrim(line, 108, 15)

    HandleParsedDetail detailRecord, tranCode, destRT, individualAcct, amount, _
                       individualID, individualName, disc, addenda, traceNumber

End Sub


Sub ParsePaymentDetailLine(ByVal line)

    Dim detailRecord, tranCode, destRT, individualAcct
    Dim amount, identNumber, individualName, pymt, addenda, traceNumber

    detailRecord   = SafeMidTrim(line, 1, 6)
    tranCode       = SafeMidTrim(line, 11, 2)
    destRT         = SafeMidTrim(line, 14, 9)
    individualAcct = SafeMidTrim(line, 24, 16)
    amount         = SafeMidTrim(line, 42, 11)
    identNumber    = SafeMidTrim(line, 54, 15)
    individualName = SafeMidTrim(line, 70, 21)
    pymt           = SafeMidTrim(line, 91, 4)
    addenda        = SafeMidTrim(line, 96, 8)
    traceNumber    = SafeMidTrim(line, 108, 15)

    HandleParsedDetail detailRecord, tranCode, destRT, individualAcct, amount, _
                       identNumber, individualName, pymt, addenda, traceNumber

End Sub


Sub ParseTerminalDetailLine(ByVal line)

    Dim detailRecord, tranCode, destRT, individualAcct
    Dim amount, identification, nameValue, codeValue, addenda, traceNumber

    detailRecord   = SafeMidTrim(line, 1, 6)
    tranCode       = SafeMidTrim(line, 11, 2)
    destRT         = SafeMidTrim(line, 14, 9)
    individualAcct = SafeMidTrim(line, 24, 16)
    amount         = SafeMidTrim(line, 42, 11)
    identification = SafeMidTrim(line, 54, 15)
    nameValue      = SafeMidTrim(line, 70, 21)
    codeValue      = SafeMidTrim(line, 91, 4)
    addenda        = SafeMidTrim(line, 96, 8)
    traceNumber    = SafeMidTrim(line, 108, 15)

    HandleParsedDetail detailRecord, tranCode, destRT, individualAcct, amount, _
                       identification, nameValue, codeValue, addenda, traceNumber

End Sub


Sub HandleParsedDetail(ByVal detailRecord, ByVal tranCode, ByVal destRT, _
                       ByVal individualAcct, ByVal amount, ByVal individualID, _
                       ByVal individualName, ByVal disc, ByVal addenda, _
                       ByVal traceNumber)

    If Trim(addenda) = "1" Then

        gPendingDetail = True
        gPendingDetailRecord = detailRecord
        gPendingTranCode = tranCode
        gPendingDestRT = destRT
        gPendingIndividualAcct = individualAcct
        gPendingAmount = amount
        gPendingIndividualID = individualID
        gPendingIndividualName = individualName
        gPendingDisc = disc
        gPendingAddenda = addenda
        gPendingTraceNumber = traceNumber

    Else

        WriteDetailRow detailRecord, tranCode, destRT, individualAcct, amount, _
                       individualID, individualName, disc, addenda, traceNumber, _
                       "", "", "", ""

    End If

End Sub


Sub ParseAndWriteAddendaLine(ByVal line)

    Select Case gAddendaLayoutType
        Case "PAYMENTINFO"
            ParsePaymentInfoAddendaLine line

        Case "TERMINALINFO"
            ParseTerminalInfoAddendaLine line

        Case Else
            ParsePaymentInfoAddendaLine line
    End Select

End Sub


Sub ParsePaymentInfoAddendaLine(ByVal line)

    Dim addendaType, paymentInfo, addSeq, dtlSeq

    addendaType = SafeMidTrim(line, 1, 13)
    paymentInfo = SafeMidTrim(line, 14, 87)
    addSeq      = SafeMidTrim(line, 101, 7)
    dtlSeq      = SafeMidTrim(line, 111, 12)

    WriteDetailRow gPendingDetailRecord, gPendingTranCode, gPendingDestRT, _
                   gPendingIndividualAcct, gPendingAmount, gPendingIndividualID, _
                   gPendingIndividualName, gPendingDisc, gPendingAddenda, _
                   gPendingTraceNumber, addendaType, paymentInfo, addSeq, dtlSeq

    gPendingDetail = False

End Sub


Sub ParseTerminalInfoAddendaLine(ByVal line)

    Dim addendaType, ref1, ref2, terminal, serialNo, authDate
    Dim authCode, locationValue, cityValue, stateValue, addendaTrace
    Dim paymentInfo

    addendaType   = SafeMidTrim(line, 1, 15)
    ref1          = SafeMidTrim(line, 14, 7)
    ref2          = SafeMidTrim(line, 22, 4)
    terminal      = SafeMidTrim(line, 26, 8)
    serialNo      = SafeMidTrim(line, 35, 6)
    authDate      = SafeMidTrim(line, 42, 4)
    authCode      = SafeMidTrim(line, 47, 6)
    locationValue = SafeMidTrim(line, 54, 27)
    cityValue     = SafeMidTrim(line, 82, 16)
    stateValue    = SafeMidTrim(line, 99, 4)
    addendaTrace  = SafeMidTrim(line, 108, 15)

    paymentInfo = "Ref1=" & ref1 & "; Ref2=" & ref2 & _
                  "; TER=" & terminal & "; SN=" & serialNo & _
                  "; Date=" & authDate & "; Auth=" & authCode & _
                  "; Loc=" & locationValue & "; City=" & cityValue & _
                  "; ST=" & stateValue

    WriteDetailRow gPendingDetailRecord, gPendingTranCode, gPendingDestRT, _
                   gPendingIndividualAcct, gPendingAmount, gPendingIndividualID, _
                   gPendingIndividualName, gPendingDisc, gPendingAddenda, _
                   gPendingTraceNumber, addendaType, paymentInfo, "", addendaTrace

    gPendingDetail = False

End Sub


Sub WriteDetailRow(ByVal detailRecord, ByVal tranCode, ByVal destRT, _
                   ByVal individualAcct, ByVal amount, ByVal individualID, _
                   ByVal individualName, ByVal disc, ByVal addenda, _
                   ByVal traceNumber, ByVal addendaType, ByVal paymentInfo, _
                   ByVal addSeq, ByVal dtlSeq)

    Dim outLine

    outLine = _
        PadRight(gReportName, 12) & _
        PadRight(gReportDate, 12) & _
        PadRight(gReportTime, 10) & _
        PadRight(gReportSeq, 10) & _
        PadRight(gFileRecordType & " " & gFileHeaderType, 8) & _
        PadRight(gFilePriority, 5) & _
        PadRight(gFileReceivingRT, 14) & _
        PadRight(gFileSendingRT, 12) & _
        PadRight(gFileDate, 10) & _
        PadRight(gFileTime, 10) & _
        PadRight(gFileModifier, 5) & _
        PadRight(gFileRecordSize, 8) & _
        PadRight(gFileBlockFactor, 5) & _
        PadRight(gFileFormat, 7) & _
        PadRight(gFileReceivingName, 25) & _
        PadRight(gFileSendingName, 25) & _
        PadRight(gFileReference, 11) & _
        PadRight(gBatchSCC, 5) & _
        PadRight(gCompanyName, 26) & _
        PadRight(gCompanyDiscretion, 22) & _
        PadRight(gCompanyID, 14) & _
        PadRight(gECC, 6) & _
        PadRight(gEntryDescription, 14) & _
        PadRight(gDescDate, 10) & _
        PadRight(gEntryDate, 10) & _
        PadRight(gSettlementDays, 12) & _
        PadRight(gBatchStatus, 8) & _
        PadRight(gOriginRT, 10) & _
        PadRight(gBatchNumber, 10) & _
        PadRight(detailRecord, 8) & _
        PadRight(tranCode, 6) & _
        PadRight(destRT, 10) & _
        PadRight(individualAcct, 16) & _
        PadRight(amount, 13) & _
        PadRight(individualID, 18) & _
        PadRight(individualName, 25) & _
        PadRight(disc, 6) & _
        PadRight(addenda, 8) & _
        PadRight(traceNumber, 18) & _
        PadRight(addendaType, 15) & _
        PadRight(paymentInfo, 122) & _
        PadRight(addSeq, 14) & _
        PadRight(dtlSeq, 14)

    fout.WriteLine outLine

End Sub


'============================================================
' Detection Helpers
'============================================================

Function IsDetailDataLine(ByVal line)

    Dim s
    s = Trim(line & "")

    If Len(s) < 20 Then
        IsDetailDataLine = False
        Exit Function
    End If

    If StartsWithText(s, "Batch") Then
        IsDetailDataLine = False
        Exit Function
    End If

    If StartsWithText(s, "File") Then
        IsDetailDataLine = False
        Exit Function
    End If

    If StartsWithText(s, "Detail") Then
        IsDetailDataLine = False
        Exit Function
    End If

    If StartsWithText(s, "Addenda") Then
        IsDetailDataLine = False
        Exit Function
    End If

    IsDetailDataLine = IsNumeric(Left(s, 2))

End Function


Function IsPageHeaderOrNoise(ByVal line)

    Dim s
    s = Trim(line & "")

    If Len(s) = 0 Then
        IsPageHeaderOrNoise = True
        Exit Function
    End If

    If InStr(1, s, "Page", vbTextCompare) > 0 _
       And InStr(1, s, "ACH List", vbTextCompare) > 0 Then
        IsPageHeaderOrNoise = True
        Exit Function
    End If

    IsPageHeaderOrNoise = False

End Function


Function IsBlankLine(ByVal line)
    IsBlankLine = (Len(Trim(line & "")) = 0)
End Function


Function StartsWithText(ByVal line, ByVal prefix)
    StartsWithText = (LCase(Left(Trim(line & ""), Len(prefix))) = LCase(prefix))
End Function


'============================================================
' String Helpers
'============================================================

Function SafeMidTrim(ByVal s, ByVal startPos, ByVal lengthVal)

    If startPos > Len(s) Then
        SafeMidTrim = ""
    Else
        SafeMidTrim = Trim(Mid(s, startPos, lengthVal))
    End If

End Function


Function PadRight(ByVal s, ByVal n)

    s = Trim(s & "")

    If Len(s) >= n Then
        PadRight = Left(s, n)
    Else
        PadRight = s & Space(n - Len(s))
    End If

End Function


Function PadLeft(ByVal s, ByVal n)

    s = Trim(s & "")

    If Len(s) >= n Then
        PadLeft = Right(s, n)
    Else
        PadLeft = Space(n - Len(s)) & s
    End If

End Function


Sub ClearAllContext()

    gReportName = ""
    gReportDate = ""
    gReportTime = ""
    gReportSeq = ""

    gFileRecordType = ""
    gFileHeaderType = ""
    gFilePriority = ""
    gFileReceivingRT = ""
    gFileSendingRT = ""
    gFileDate = ""
    gFileTime = ""
    gFileModifier = ""
    gFileRecordSize = ""
    gFileBlockFactor = ""
    gFileFormat = ""
    gFileReceivingName = ""
    gFileSendingName = ""
    gFileReference = ""

    gBatchSCC = ""
    gCompanyName = ""
    gCompanyDiscretion = ""
    gCompanyID = ""
    gECC = ""
    gEntryDescription = ""
    gDescDate = ""
    gEntryDate = ""
    gSettlementDays = ""
    gBatchStatus = ""
    gOriginRT = ""
    gBatchNumber = ""

End Sub