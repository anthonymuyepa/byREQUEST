Option Explicit

'============================================================
' ATM Daily Summary Extract
'
' Purpose:
'   - Find each ATM block starting with:
'       "         -PRCH24            "
'   - Within each block, find the TERMINAL line where
'       "TERMINAL     " starts at position 23
'   - Extract PREVIOUS SUSPENSE TOTALS:
'       Terminal
'       Cash-In
'       Cash-Out
'       Net Cash
'       Withdrawals
'       Check Deposits
'       Cash Deposits
'
'   - Optionally extract TODAYS SUSPENSE TOTALS:
'       Today Cash-In
'       Today Cash-Out
'       Today Net Cash
'       Today Withdrawals
'       Today Check Deposits
'       Today Cash Deposits
'
' Notes:
'   - TODAYS SUSPENSE TOTALS are also inside the -PRCH24 block.
'   - The script separates PREVIOUS and TODAY totals so one section
'     does not overwrite the other.
'============================================================

Dim gInBlock
Dim gHasTerminal
Dim gBlockLineCount

Dim gTerminal
Dim gCashIn
Dim gCashOut
Dim gNetCash
Dim gWithdrawals
Dim gCheckDeposits
Dim gCashDeposits
Dim gTransfersFrom
Dim gTransfersTo
Dim gPaymentsFrom
Dim gPaymentsTo
Dim gTotalDebits
Dim gTotalCredits
Dim gTotalMemo
Dim gInquiries
Dim gOtherMemo
Dim gExceptions
Dim gRejects

Dim INCLUDE_TODAYS_SUSPENSE_TOTALS
Dim gSuspenseSection

Dim gTodayCashIn
Dim gTodayCashOut
Dim gTodayNetCash
Dim gTodayWithdrawals
Dim gTodayCheckDeposits
Dim gTodayCashDeposits
Dim gTodayTransfersFrom
Dim gTodayTransfersTo
Dim gTodayPaymentsFrom
Dim gTodayPaymentsTo
Dim gTodayTotalDebits
Dim gTodayTotalCredits
Dim gTodayTotalMemo
Dim gTodayInquiries
Dim gTodayOtherMemo

Dim BLOCK_START_TEXT

Sub StartDoc()

    BLOCK_START_TEXT = " -PRCH24  "

    '========================================================
    ' OPTIONAL TODAY'S SUSPENSE TOTALS CAPTURE
    '
    ' Set to True to include Today's Suspense Totals columns.
    ' Set to False if Today's Suspense Totals are not wanted.
    '========================================================
    INCLUDE_TODAYS_SUSPENSE_TOTALS = True

    gInBlock = False
    gHasTerminal = False
    gBlockLineCount = 0

    ResetBlock

    If INCLUDE_TODAYS_SUSPENSE_TOTALS = True Then

        Fout.WriteLine _
            PadRight("Terminal", 15) & _
            PadLeft("Cash-In", 15) & _
            PadLeft("Cash-Out", 15) & _
            PadLeft("Net Cash", 15) & _
            PadLeft("Withdrawals", 15) & _
            PadLeft("Check Deposits", 18) & _
            PadLeft("Cash Deposits", 17) & _
            PadLeft("Today Cash-In", 18) & _
            PadLeft("Today Cash-Out", 18) & _
            PadLeft("Today Net Cash", 18) & _
            PadLeft("Today Withdrawals", 20) & _
            PadLeft("Today Check Dep", 20) & _
            PadLeft("Today Cash Dep", 18)

    Else

        Fout.WriteLine _
            PadRight("Terminal", 15) & _
            PadLeft("Cash-In", 15) & _
            PadLeft("Cash-Out", 15) & _
            PadLeft("Net Cash", 15) & _
            PadLeft("Withdrawals", 15) & _
            PadLeft("Check Deposits", 18) & _
            PadLeft("Cash Deposits", 17)

    End If

End Sub

Sub ProcessLine()

    Dim line
    line = srec

    '--------------------------------------------------------
    ' A new block starts.
    ' If we were already in a block, output the previous one.
    '--------------------------------------------------------
    If InStr(1, line, BLOCK_START_TEXT, vbTextCompare) > 0 Then

        If gInBlock = True Then
            WriteBlock
        End If

        ResetBlock
        gInBlock = True
        gBlockLineCount = 0
        Exit Sub

    End If

    If gInBlock = False Then Exit Sub

    gBlockLineCount = gBlockLineCount + 1

    '--------------------------------------------------------
    ' Find the correct TERMINAL line.
    ' There may be another TERMINAL elsewhere, but the one
    ' we want has TERMINAL beginning at position 23.
    '--------------------------------------------------------
    If Mid(line, 23, Len("TERMINAL")) = "TERMINAL" Then
        gHasTerminal = True
        ParseTerminalLine line
        Exit Sub
    End If

    '--------------------------------------------------------
    ' Once terminal has been found, parse suspense totals lines.
    '--------------------------------------------------------
    If gHasTerminal = True Then
        ParseSuspenseLine line
    End If

End Sub

Sub CloseDoc()

    If gInBlock = True Then
        WriteBlock
    End If

End Sub

'============================================================
' RESET BLOCK
'============================================================

Sub ResetBlock()

    gTerminal = ""
    gCashIn = ""
    gCashOut = ""
    gNetCash = ""
    gWithdrawals = ""
    gCheckDeposits = ""
    gCashDeposits = ""
    gTransfersFrom = ""
    gTransfersTo = ""
    gPaymentsFrom = ""
    gPaymentsTo = ""
    gTotalDebits = ""
    gTotalCredits = ""
    gTotalMemo = ""
    gInquiries = ""
    gOtherMemo = ""
    gExceptions = ""
    gRejects = ""

    gSuspenseSection = ""

    gTodayCashIn = ""
    gTodayCashOut = ""
    gTodayNetCash = ""
    gTodayWithdrawals = ""
    gTodayCheckDeposits = ""
    gTodayCashDeposits = ""
    gTodayTransfersFrom = ""
    gTodayTransfersTo = ""
    gTodayPaymentsFrom = ""
    gTodayPaymentsTo = ""
    gTodayTotalDebits = ""
    gTodayTotalCredits = ""
    gTodayTotalMemo = ""
    gTodayInquiries = ""
    gTodayOtherMemo = ""

    gHasTerminal = False

End Sub

'============================================================
' WRITE ONE OUTPUT ROW
'============================================================

Sub WriteBlock()

    ' Skip the whole block if required data was not found
    If HasRequiredATMData() = False Then
        Exit Sub
    End If

    If INCLUDE_TODAYS_SUSPENSE_TOTALS = True Then

        Fout.WriteLine _
            PadRight(gTerminal, 15) & _
            PadLeft(gCashIn, 15) & _
            PadLeft(gCashOut, 15) & _
            PadLeft(gNetCash, 15) & _
            PadLeft(gWithdrawals, 15) & _
            PadLeft(gCheckDeposits, 18) & _
            PadLeft(gCashDeposits, 17) & _
            PadLeft(gTodayCashIn, 18) & _
            PadLeft(gTodayCashOut, 18) & _
            PadLeft(gTodayNetCash, 18) & _
            PadLeft(gTodayWithdrawals, 20) & _
            PadLeft(gTodayCheckDeposits, 20) & _
            PadLeft(gTodayCashDeposits, 18)

    Else

        Fout.WriteLine _
            PadRight(gTerminal, 15) & _
            PadLeft(gCashIn, 15) & _
            PadLeft(gCashOut, 15) & _
            PadLeft(gNetCash, 15) & _
            PadLeft(gWithdrawals, 15) & _
            PadLeft(gCheckDeposits, 18) & _
            PadLeft(gCashDeposits, 17)

    End If

End Sub

'============================================================
' PARSE TERMINAL LINE
'
' Sample:
' 1134 N. AVALON BLVD. TERMINAL  COHO11 ...
'
' Rule:
'   Use the TERMINAL at position 23.
'   Extract the token immediately after TERMINAL as Terminal ID.
'============================================================

Sub ParseTerminalLine(ByVal line)

    Dim posTerminal
    Dim afterTerminal
    Dim parts

    posTerminal = InStr(23, line, "TERMINAL", vbTextCompare)

    If posTerminal > 0 Then
        afterTerminal = Trim(Mid(line, posTerminal + Len("TERMINAL")))
        parts = SplitBySpaces(afterTerminal)

        If IsArray(parts) Then
            If UBound(parts) >= 0 Then
                gTerminal = parts(0)
            End If
        End If
    End If

End Sub

'============================================================
' PARSE SUSPENSE SECTIONS
'
' This keeps PREVIOUS SUSPENSE TOTALS and TODAYS SUSPENSE TOTALS
' separate. It also prevents counts and transaction detail lines
' from being treated as totals.
'============================================================

Sub ParseSuspenseLine(ByVal line)

    '--------------------------------------------------------
    ' Detect PREVIOUS totals section.
    ' The section header line also contains CASH-IN/CASH-OUT,
    ' so parse it immediately.
    '--------------------------------------------------------
    If InStr(1, line, "*** PREVIOUS SUSPENSE TOTALS", vbTextCompare) > 0 Then
        gSuspenseSection = "PREVIOUS_TOTALS"
        ParsePreviousSuspenseTotalsLine line
        Exit Sub
    End If

    '--------------------------------------------------------
    ' Detect TODAYS totals section.
    ' The section header line also contains CASH-IN/CASH-OUT,
    ' so parse it immediately.
    '--------------------------------------------------------
    If InStr(1, line, "*** TODAYS SUSPENSE TOTALS", vbTextCompare) > 0 Then
        gSuspenseSection = "TODAYS_TOTALS"

        '====================================================
        ' OPTIONAL BLOCK:
        ' Comment this block out, or set
        ' INCLUDE_TODAYS_SUSPENSE_TOTALS = False in StartDoc,
        ' if Today's Suspense Totals are not wanted.
        '====================================================
        If INCLUDE_TODAYS_SUSPENSE_TOTALS = True Then
            ParseTodaysSuspenseTotalsLine line
        End If

        Exit Sub
    End If

    '--------------------------------------------------------
    ' Stop parsing totals when counts section starts.
    '--------------------------------------------------------
    If InStr(1, line, "*** PREVIOUS SUSPENSE COUNTS", vbTextCompare) > 0 Then
        gSuspenseSection = ""
        Exit Sub
    End If

    If InStr(1, line, "*** TODAYS SUSPENSE COUNTS", vbTextCompare) > 0 Then
        gSuspenseSection = ""
        Exit Sub
    End If

    '--------------------------------------------------------
    ' Parse detail lines only while inside a totals section.
    '--------------------------------------------------------
    If gSuspenseSection = "PREVIOUS_TOTALS" Then
        ParsePreviousSuspenseTotalsLine line
        Exit Sub
    End If

    If gSuspenseSection = "TODAYS_TOTALS" Then

        '====================================================
        ' OPTIONAL BLOCK:
        ' Comment this block out, or set
        ' INCLUDE_TODAYS_SUSPENSE_TOTALS = False in StartDoc,
        ' if Today's Suspense Totals are not wanted.
        '====================================================
        If INCLUDE_TODAYS_SUSPENSE_TOTALS = True Then
            ParseTodaysSuspenseTotalsLine line
        End If

        Exit Sub
    End If

End Sub

'============================================================
' PARSE PREVIOUS SUSPENSE TOTALS / DETAIL LINES
'============================================================

Sub ParsePreviousSuspenseTotalsLine(ByVal line)

    If InStr(1, line, "CASH-IN", vbTextCompare) > 0 Then
        gCashIn = ExtractAmountAfterLabel(line, "CASH-IN")
    End If

    If InStr(1, line, "CASH-OUT", vbTextCompare) > 0 Then
        gCashOut = ExtractAmountAfterLabel(line, "CASH-OUT")
    End If

    If InStr(1, line, "NET CASH", vbTextCompare) > 0 Then
        gNetCash = ExtractAmountAfterLabel(line, "NET CASH")
    End If

    If InStr(1, line, "WITHDRAWALS", vbTextCompare) > 0 Then
        gWithdrawals = ExtractAmountAfterLabel(line, "WITHDRAWALS")
    End If

    If InStr(1, line, "CHECK DEPOSITS", vbTextCompare) > 0 Then
        gCheckDeposits = ExtractAmountAfterLabel(line, "CHECK DEPOSITS")
    End If

    If InStr(1, line, "CASH DEPOSITS", vbTextCompare) > 0 Then
        gCashDeposits = ExtractAmountAfterLabel(line, "CASH DEPOSITS")
    End If

    If InStr(1, line, "TOTAL DEBITS", vbTextCompare) > 0 Then
        gTotalDebits = ExtractAmountAfterLabel(line, "TOTAL DEBITS")
    End If

    If InStr(1, line, "TOTAL CREDITS", vbTextCompare) > 0 Then
        gTotalCredits = ExtractAmountAfterLabel(line, "TOTAL CREDITS")
    End If

    If InStr(1, line, "TOTAL MEMO", vbTextCompare) > 0 Then
        gTotalMemo = ExtractAmountAfterLabel(line, "TOTAL MEMO")
    End If

    If InStr(1, line, "INQUIRIES", vbTextCompare) > 0 Then
        gInquiries = ExtractAmountAfterLabel(line, "INQUIRIES")
    End If

    If InStr(1, line, "OTHER MEMO", vbTextCompare) > 0 Then
        gOtherMemo = ExtractAmountAfterLabel(line, "OTHER MEMO")
    End If

    If InStr(1, line, "TRANSFERS FROM", vbTextCompare) > 0 Then
        gTransfersFrom = ExtractAmountAfterLabel(line, "TRANSFERS FROM")
    End If

    If InStr(1, line, "TRANSFERS TO", vbTextCompare) > 0 Then
        gTransfersTo = ExtractAmountAfterLabel(line, "TRANSFERS TO")
    End If

    If InStr(1, line, "PAYMENTS FROM", vbTextCompare) > 0 Then
        gPaymentsFrom = ExtractAmountAfterLabel(line, "PAYMENTS FROM")
    End If

    If InStr(1, line, "PAYMENTS TO", vbTextCompare) > 0 Then
        gPaymentsTo = ExtractAmountAfterLabel(line, "PAYMENTS TO")
    End If

    If InStr(1, line, "EXCEPTION", vbTextCompare) > 0 Then
        gExceptions = ExtractAmountAfterLabel(line, "EXCEPTION")
    End If

    If InStr(1, line, "REJECT", vbTextCompare) > 0 Then
        gRejects = ExtractAmountAfterLabel(line, "REJECT")
    End If

End Sub

'============================================================
' PARSE TODAY'S SUSPENSE TOTALS / DETAIL LINES
'============================================================

Sub ParseTodaysSuspenseTotalsLine(ByVal line)

    If InStr(1, line, "CASH-IN", vbTextCompare) > 0 Then
        gTodayCashIn = ExtractAmountAfterLabel(line, "CASH-IN")
    End If

    If InStr(1, line, "CASH-OUT", vbTextCompare) > 0 Then
        gTodayCashOut = ExtractAmountAfterLabel(line, "CASH-OUT")
    End If

    If InStr(1, line, "NET CASH", vbTextCompare) > 0 Then
        gTodayNetCash = ExtractAmountAfterLabel(line, "NET CASH")
    End If

    If InStr(1, line, "WITHDRAWALS", vbTextCompare) > 0 Then
        gTodayWithdrawals = ExtractAmountAfterLabel(line, "WITHDRAWALS")
    End If

    If InStr(1, line, "CHECK DEPOSITS", vbTextCompare) > 0 Then
        gTodayCheckDeposits = ExtractAmountAfterLabel(line, "CHECK DEPOSITS")
    End If

    If InStr(1, line, "CASH DEPOSITS", vbTextCompare) > 0 Then
        gTodayCashDeposits = ExtractAmountAfterLabel(line, "CASH DEPOSITS")
    End If

    If InStr(1, line, "TOTAL DEBITS", vbTextCompare) > 0 Then
        gTodayTotalDebits = ExtractAmountAfterLabel(line, "TOTAL DEBITS")
    End If

    If InStr(1, line, "TOTAL CREDITS", vbTextCompare) > 0 Then
        gTodayTotalCredits = ExtractAmountAfterLabel(line, "TOTAL CREDITS")
    End If

    If InStr(1, line, "TOTAL MEMO", vbTextCompare) > 0 Then
        gTodayTotalMemo = ExtractAmountAfterLabel(line, "TOTAL MEMO")
    End If

    If InStr(1, line, "INQUIRIES", vbTextCompare) > 0 Then
        gTodayInquiries = ExtractAmountAfterLabel(line, "INQUIRIES")
    End If

    If InStr(1, line, "OTHER MEMO", vbTextCompare) > 0 Then
        gTodayOtherMemo = ExtractAmountAfterLabel(line, "OTHER MEMO")
    End If

    If InStr(1, line, "TRANSFERS FROM", vbTextCompare) > 0 Then
        gTodayTransfersFrom = ExtractAmountAfterLabel(line, "TRANSFERS FROM")
    End If

    If InStr(1, line, "TRANSFERS TO", vbTextCompare) > 0 Then
        gTodayTransfersTo = ExtractAmountAfterLabel(line, "TRANSFERS TO")
    End If

    If InStr(1, line, "PAYMENTS FROM", vbTextCompare) > 0 Then
        gTodayPaymentsFrom = ExtractAmountAfterLabel(line, "PAYMENTS FROM")
    End If

    If InStr(1, line, "PAYMENTS TO", vbTextCompare) > 0 Then
        gTodayPaymentsTo = ExtractAmountAfterLabel(line, "PAYMENTS TO")
    End If

End Sub

'============================================================
' EXTRACT AMOUNT AFTER LABEL
'
' Example:
'   CASH-IN        68,740.00
'   WITHDRAWALS    2,180.00-
'
' Returns the first amount-looking token after the label.
' Handles trailing minus values like:
'   2,180.00-
'============================================================

Function ExtractAmountAfterLabel(ByVal line, ByVal label)

    Dim pos
    Dim afterLabel
    Dim tokens
    Dim i
    Dim t

    ExtractAmountAfterLabel = ""

    pos = InStr(1, line, label, vbTextCompare)
    If pos = 0 Then Exit Function

    afterLabel = Mid(line, pos + Len(label))
    tokens = SplitBySpaces(afterLabel)

    If Not IsArray(tokens) Then Exit Function

    For i = 0 To UBound(tokens)
        t = Trim(tokens(i))

        If LooksLikeAmount(t) Then
            ExtractAmountAfterLabel = NormalizeAmount(t)
            Exit Function
        End If
    Next

End Function

'============================================================
' SPLIT BY MULTIPLE SPACES
'============================================================

Function SplitBySpaces(ByVal text)

    Dim re

    text = Trim(text)

    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "\s+"
    re.Global = True

    SplitBySpaces = Split(re.Replace(text, " "), " ")

End Function

'============================================================
' AMOUNT HELPERS
'============================================================

Function LooksLikeAmount(ByVal s)

    Dim t

    LooksLikeAmount = False

    t = Trim(s)
    If t = "" Then Exit Function

    t = Replace(t, ",", "")

    ' Handle trailing minus, e.g. 2,180.00-
    If Right(t, 1) = "-" Then
        t = "-" & Left(t, Len(t) - 1)
    End If

    If IsNumeric(t) Then
        LooksLikeAmount = True
    End If

End Function

Function NormalizeAmount(ByVal s)

    Dim t
    Dim isTrailingMinus
    Dim n

    t = Trim(s)
    isTrailingMinus = False

    If Right(t, 1) = "-" Then
        isTrailingMinus = True
        t = Left(t, Len(t) - 1)
    End If

    t = Replace(t, ",", "")

    If IsNumeric(t) Then
        n = CDbl(t)

        If isTrailingMinus Then
            n = n * -1
        End If

        NormalizeAmount = FormatNumber(n, 2, -1, 0, -1)
    Else
        NormalizeAmount = s
    End If

End Function

'============================================================
' PAD HELPERS
'============================================================

Function PadRight(ByVal s, ByVal n)

    If IsNull(s) Then s = ""
    s = CStr(s)

    PadRight = Left(s & Space(n), n)

End Function

Function PadLeft(ByVal s, ByVal n)

    If IsNull(s) Then s = ""
    s = CStr(s)

    PadLeft = Right(Space(n) & s, n)

End Function

'============================================================
' VALID BLOCK CHECK
'============================================================

Function HasRequiredATMData()

    HasRequiredATMData = False

    ' Required minimum data points for a valid block.
    ' These are based on PREVIOUS SUSPENSE TOTALS.
    If Trim(gTerminal) = "" Then Exit Function
    If Trim(gCashIn) = "" Then Exit Function
    If Trim(gCashOut) = "" Then Exit Function
    If Trim(gNetCash) = "" Then Exit Function
    If Trim(gWithdrawals) = "" Then Exit Function
    If Trim(gCheckDeposits) = "" Then Exit Function
    If Trim(gCashDeposits) = "" Then Exit Function

    HasRequiredATMData = True

End Function