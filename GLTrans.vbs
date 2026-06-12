'==============================================================
' GL Transaction Register Parser - Flattened Output Only
' byREQUEST-style script
'
' Output:
'   One fixed-width flattened dataset
'
' LineType values:
'   Transaction
'   BeginningBalance
'   EndingBalance
'   AverageEndingBalance
'
' No GL SUMMARY output
' No PARSE EXCEPTIONS output
'==============================================================

Option Explicit

'==============================================================
' Column Widths
'==============================================================
Dim W_RPTDATE, W_START, W_END, W_PAGE, W_GL, W_GLDESC, W_EFFDATE, W_POSTDATE, W_SEQ, W_USER, W_CAT, W_REF, W_COMMENT, W_AMOUNT, W_LINETYPE



'==============================================================
' Global Variables
'==============================================================
Dim g_ReportDate
Dim g_StartDate
Dim g_EndDate
Dim g_PageNo
Dim g_SourceFile

Dim g_CurrentGL
Dim g_CurrentDesc
Dim g_CurrentBeginningBalance
Dim g_CurrentEndingBalance
Dim g_CurrentDebitTotal
Dim g_CurrentCreditTotal
Dim g_CurrentTransCount

Dim g_LineNo

'==============================================================
' StartDoc
'==============================================================
Sub StartDoc()

	
	W_RPTDATE = 10
	W_START = 10
	W_END = 10
	W_PAGE = 6
	W_GL = 18
	W_GLDESC = 50
	W_EFFDATE = 10
	W_POSTDATE = 10
	W_SEQ = 10
	W_USER = 8
	W_CAT = 6
	W_REF = 12
	W_COMMENT = 55
	W_AMOUNT = 14
	W_LINETYPE = 22

 
    g_ReportDate = ""
    g_StartDate = ""
    g_EndDate = ""
    g_PageNo = ""
    g_SourceFile = ""

    g_CurrentGL = ""
    g_CurrentDesc = ""
    g_CurrentBeginningBalance = ""
    g_CurrentEndingBalance = ""
    g_CurrentDebitTotal = 0
    g_CurrentCreditTotal = 0
    g_CurrentTransCount = 0

    g_LineNo = 0

    On Error Resume Next
    g_SourceFile = SpoolFile.Name
    On Error GoTo 0

    Fout.WriteLine _
        PadRight("RptDate", W_RPTDATE) & _
        PadRight("Start", W_START) & _
        PadRight("End", W_END) & _
        PadRight("Page", W_PAGE) & _
        PadRight("GL Account", W_GL) & _
        PadRight("GL Description", W_GLDESC) & _
        PadRight("EffDate", W_EFFDATE) & _
        PadRight("PostDate", W_POSTDATE) & _
        PadRight("Seq", W_SEQ) & _
        PadRight("User", W_USER) & _
        PadRight("Cat", W_CAT) & _
        PadRight("Reference", W_REF) & _
        PadRight("Comment", W_COMMENT) & _
        PadRight("Debit", W_AMOUNT) & _
        PadRight("Credit", W_AMOUNT) & _
        PadRight("Balance", W_AMOUNT) & _
        PadRight("LineType", W_LINETYPE)

End Sub

'==============================================================
' ProcessLine
'==============================================================
Sub ProcessLine()

    Dim line

    g_LineNo = g_LineNo + 1
    line = srec

    If Trim(line) = "" Then Exit Sub

    Call CaptureHeaderInfo(line)

    If IsHeaderLine(line) Then Exit Sub
    If IsDashedLine(line) Then Exit Sub
    If IsAmountOnlyLine(line) Then Exit Sub

    If IsGLStartLine(line) Then
        Call ParseGLStartLine(line)
        Exit Sub
    End If

    If g_CurrentGL = "" Then Exit Sub

    If InStr(1, line, "Beginning Balance", vbTextCompare) > 0 Then
        Call ParseBeginningBalance(line)
        Exit Sub
    End If

    If InStr(1, line, "Average and Ending Balance", vbTextCompare) > 0 Then
        Call ParseEndingBalance(line)
        Exit Sub
    End If

    If InStr(1, line, "Ending Balance", vbTextCompare) > 0 Then
        Call ParseEndingBalance(line)
        Exit Sub
    End If

    If IsTransactionLine(line) Then
        Call ParseTransactionLine(line)
        Exit Sub
    End If
	

    'Flattened output only.
    'Unknown lines are skipped intentionally.

End Sub

'==============================================================
' CloseDoc
'==============================================================
Sub CloseDoc()

    'Flattened output only.
    'Do not print GL SUMMARY.
    'Do not print PARSE EXCEPTIONS.

End Sub

'==============================================================
' Header Parsing
'==============================================================
Sub CaptureHeaderInfo(ByVal line)

    Dim p, parts, i

    If InStr(1, line, "GL Transaction Register", vbTextCompare) > 0 Then

        parts = SplitBySpaces(line)

        For i = 0 To UBound(parts)

            If IsDateToken(parts(i)) Then
                If g_ReportDate = "" Then g_ReportDate = parts(i)
            End If

            If UCase(parts(i)) = "PAGE" Then
                If i + 1 <= UBound(parts) Then
                    g_PageNo = parts(i + 1)
                End If
            End If

        Next

    End If

    If InStr(1, line, "Start Date:", vbTextCompare) > 0 Then

        p = InStr(1, line, "Start Date:", vbTextCompare)
        If p > 0 Then
            g_StartDate = Trim(Mid(line, p + Len("Start Date:"), 9))
        End If

        p = InStr(1, line, "End Date:", vbTextCompare)
        If p > 0 Then
            g_EndDate = Trim(Mid(line, p + Len("End Date:"), 9))
        End If

    End If

End Sub

Function IsHeaderLine(ByVal line)

    Dim t
    t = Trim(CStr(line))

    IsHeaderLine = False

    If InStr(1, t, "GL Transaction Register", vbTextCompare) > 0 Then IsHeaderLine = True
    If InStr(1, t, "For All Accounts", vbTextCompare) > 0 Then IsHeaderLine = True
    If InStr(1, t, "Start Date:", vbTextCompare) > 0 Then IsHeaderLine = True

    If InStr(1, t, "Effect", vbTextCompare) > 0 _
        And InStr(1, t, "Credits", vbTextCompare) > 0 _
        And InStr(1, t, "Balance", vbTextCompare) > 0 Then

        IsHeaderLine = True

    End If

End Function

'==============================================================
' GL Section Detection
'==============================================================
Function IsGLStartLine(ByVal line)

    Dim t, token1, token2

    t = Trim(CStr(line))

    IsGLStartLine = False

    If Len(t) < 18 Then Exit Function

    token1 = GetToken(t, 1)
    token2 = GetToken(t, 2)

    'Normal format:
    '09/03/19 211100-0000.0035 SALARIES - WEST COBB Beginning Balance 3,537.72
    If IsDateToken(token1) And IsGLAccount(token2) Then
        IsGLStartLine = True
        Exit Function
    End If

    'Wrapped/variation format:
    '130800-0000.0000 WIRE TRANSFER FEES
    If IsGLAccount(token1) Then
        IsGLStartLine = True
        Exit Function
    End If

End Function

Sub ParseGLStartLine(ByVal line)

    Dim t
    Dim token1, token2
    Dim glAcct
    Dim descStart
    Dim descText
    Dim balPos
    Dim balText
    Dim balAmt

    t = Trim(CStr(line))

    token1 = GetToken(t, 1)
    token2 = GetToken(t, 2)

    If IsDateToken(token1) And IsGLAccount(token2) Then
        glAcct = token2
    ElseIf IsGLAccount(token1) Then
        glAcct = token1
    Else
        Exit Sub
    End If

    descStart = InStr(1, t, glAcct, vbTextCompare) + Len(glAcct)
    descText = Trim(Mid(t, descStart))

    g_CurrentGL = glAcct
    g_CurrentDesc = ""
    g_CurrentBeginningBalance = ""
    g_CurrentEndingBalance = ""
    g_CurrentDebitTotal = 0
    g_CurrentCreditTotal = 0
    g_CurrentTransCount = 0

    balPos = InStr(1, descText, "Beginning Balance", vbTextCompare)

    If balPos > 0 Then

        balText = Mid(descText, balPos)
        balAmt = GetLastTrailingAmount(balText)

        g_CurrentDesc = Trim(Left(descText, balPos - 1))

        If balAmt <> "" Then
            g_CurrentBeginningBalance = CleanAmount(balAmt)
            Call WriteBalanceRow("BeginningBalance", "Beginning Balance", g_CurrentBeginningBalance)
        End If

    Else
        g_CurrentDesc = descText
    End If

End Sub

'==============================================================
' Balance Parsing
'==============================================================
Sub ParseBeginningBalance(ByVal line)

    Dim amt

    amt = GetLastTrailingAmount(line)

    If amt <> "" Then
        g_CurrentBeginningBalance = CleanAmount(amt)
        Call WriteBalanceRow("BeginningBalance", "Beginning Balance", g_CurrentBeginningBalance)
    End If

End Sub

Sub ParseEndingBalance(ByVal line)

    Dim amounts
    Dim debitAmt, creditAmt, balanceAmt
    Dim commentText
    Dim lineType
    Dim amountCount

    amounts = GetTrailingAmountsFromLine(line)
    amountCount = UBoundSafe(amounts) + 1

    debitAmt = "0.00"
    creditAmt = "0.00"
    balanceAmt = ""

    If amountCount = 1 Then

        balanceAmt = CleanAmount(amounts(0))

    ElseIf amountCount = 2 Then

        creditAmt = CleanAmount(amounts(0))
        balanceAmt = CleanAmount(amounts(1))

    ElseIf amountCount >= 3 Then

        debitAmt = CleanAmount(amounts(0))
        creditAmt = CleanAmount(amounts(1))
        balanceAmt = CleanAmount(amounts(2))

    Else
        Exit Sub
    End If

    g_CurrentDebitTotal = CDblSafe(debitAmt)
    g_CurrentCreditTotal = CDblSafe(creditAmt)
    g_CurrentEndingBalance = balanceAmt

    If InStr(1, line, "Average and Ending Balance", vbTextCompare) > 0 Then
        commentText = "Average and Ending Balance"
        lineType = "AverageEndingBalance"
    Else
        commentText = "Ending Balance"
        lineType = "EndingBalance"
    End If

    Fout.WriteLine _
        PadRight(g_ReportDate, W_RPTDATE) & _
        PadRight(g_StartDate, W_START) & _
        PadRight(g_EndDate, W_END) & _
        PadRight(g_PageNo, W_PAGE) & _
        PadRight(g_CurrentGL, W_GL) & _
        PadRight(g_CurrentDesc, W_GLDESC) & _
        PadRight("", W_EFFDATE) & _
        PadRight("", W_POSTDATE) & _
        PadRight("", W_SEQ) & _
        PadRight("", W_USER) & _
        PadRight("", W_CAT) & _
        PadRight("", W_REF) & _
        PadRight(commentText, W_COMMENT) & _
        PadRight(debitAmt, W_AMOUNT) & _
        PadRight(creditAmt, W_AMOUNT) & _
        PadRight(balanceAmt, W_AMOUNT) & _
        PadRight(lineType, W_LINETYPE)

End Sub

Sub WriteBalanceRow(ByVal lineType, ByVal commentText, ByVal balanceAmt)

    Fout.WriteLine _
        PadRight(g_ReportDate, W_RPTDATE) & _
        PadRight(g_StartDate, W_START) & _
        PadRight(g_EndDate, W_END) & _
        PadRight(g_PageNo, W_PAGE) & _
        PadRight(g_CurrentGL, W_GL) & _
        PadRight(g_CurrentDesc, W_GLDESC) & _
        PadRight("", W_EFFDATE) & _
        PadRight("", W_POSTDATE) & _
        PadRight("", W_SEQ) & _
        PadRight("", W_USER) & _
        PadRight("", W_CAT) & _
        PadRight("", W_REF) & _
        PadRight(commentText, W_COMMENT) & _
        PadRight("0.00", W_AMOUNT) & _
        PadRight("0.00", W_AMOUNT) & _
        PadRight(balanceAmt, W_AMOUNT) & _
        PadRight(lineType, W_LINETYPE)

End Sub

'==============================================================
' Transaction Parsing
'==============================================================
Function IsTransactionLine(ByVal line)

    Dim t

    t = Trim(CStr(line))

    IsTransactionLine = False

    If Len(t) < 35 Then Exit Function

    If IsDateToken(GetToken(t, 1)) And IsDateToken(GetToken(t, 2)) Then
        IsTransactionLine = True
    End If

End Function

Sub ParseTransactionLine(ByVal line)

    Dim t
    Dim effectDate, postDate, seqNo, userCode, catCode
    Dim referenceNo
    Dim restText
    Dim amounts
    Dim debitAmt, creditAmt, balanceAmt
    Dim comment
    Dim amountCount
    Dim lineType
    Dim token6

    t = Trim(line)

    lineType = "Transaction"

    effectDate = GetToken(t, 1)
    postDate = GetToken(t, 2)
    seqNo = GetToken(t, 3)
    userCode = GetToken(t, 4)
    catCode = GetToken(t, 5)

    referenceNo = ""

    token6 = GetToken(t, 6)

    If IsReferenceToken(token6) Then
        referenceNo = token6
        restText = RemoveFirstNTokens(t, 6)
    Else
        restText = RemoveFirstNTokens(t, 5)
    End If

    'Use trailing amounts only so numbers inside comments are ignored.
    amounts = GetTrailingAmountsFromLine(restText)
    amountCount = UBoundSafe(amounts) + 1

    debitAmt = "0.00"
    creditAmt = "0.00"
    balanceAmt = ""

    If amountCount = 1 Then

        creditAmt = CleanAmount(amounts(0))
        balanceAmt = ""

    ElseIf amountCount = 2 Then

        creditAmt = CleanAmount(amounts(0))
        balanceAmt = CleanAmount(amounts(1))

    ElseIf amountCount >= 3 Then

        debitAmt = CleanAmount(amounts(0))
        creditAmt = CleanAmount(amounts(1))
        balanceAmt = CleanAmount(amounts(2))

    Else

        Exit Sub

    End If

    comment = RemoveTrailingAmounts(restText)

    g_CurrentTransCount = g_CurrentTransCount + 1

    Fout.WriteLine _
        PadRight(g_ReportDate, W_RPTDATE) & _
        PadRight(g_StartDate, W_START) & _
        PadRight(g_EndDate, W_END) & _
        PadRight(g_PageNo, W_PAGE) & _
        PadRight(g_CurrentGL, W_GL) & _
        PadRight(g_CurrentDesc, W_GLDESC) & _
        PadRight(effectDate, W_EFFDATE) & _
        PadRight(postDate, W_POSTDATE) & _
        PadRight(seqNo, W_SEQ) & _
        PadRight(userCode, W_USER) & _
        PadRight(catCode, W_CAT) & _
        PadRight(referenceNo, W_REF) & _
        PadRight(comment, W_COMMENT) & _
        PadRight(debitAmt, W_AMOUNT) & _
        PadRight(creditAmt, W_AMOUNT) & _
        PadRight(balanceAmt, W_AMOUNT) & _
        PadRight(lineType, W_LINETYPE)

End Sub

'==============================================================
' Utility Functions
'==============================================================
Function PadRight(ByVal s, ByVal n)

    If IsNull(s) Then s = ""

    s = CStr(s)

    PadRight = Left(s & Space(n), n)

End Function



Function GetToken(ByVal text, ByVal tokenNo)

    Dim arr

    arr = SplitBySpaces(text)

    If tokenNo - 1 <= UBound(arr) Then
        GetToken = arr(tokenNo - 1)
    Else
        GetToken = ""
    End If

End Function

Function SplitBySpaces(ByVal text)

    Dim re, matches, arr()
    Dim i

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "\S+"

    Set matches = re.Execute(CStr(text))

    If matches.Count = 0 Then
        ReDim arr(0)
        arr(0) = ""
    Else
        ReDim arr(matches.Count - 1)

        For i = 0 To matches.Count - 1
            arr(i) = matches(i).Value
        Next
    End If

    SplitBySpaces = arr

End Function

Function RemoveFirstNTokens(ByVal text, ByVal n)

    Dim arr, i, result

    arr = SplitBySpaces(text)

    result = ""

    For i = n To UBound(arr)
        If result <> "" Then result = result & " "
        result = result & arr(i)
    Next

    RemoveFirstNTokens = result

End Function

Function IsDateToken(ByVal s)

    Dim re

    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^\d{2}/\d{2}/\d{2}$"
    re.IgnoreCase = True

    IsDateToken = re.Test(Trim(CStr(s)))

End Function

Function IsGLAccount(ByVal s)

    Dim re

    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^\d{5,6}-\d{4}\.\d{4}$"
    re.IgnoreCase = True

    IsGLAccount = re.Test(Trim(CStr(s)))

End Function

Function IsReferenceToken(ByVal s)

    Dim re

    s = Trim(CStr(s))

    IsReferenceToken = False

    If s = "" Then Exit Function

    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^\d{3,12}$"

    If re.Test(s) Then
        IsReferenceToken = True
    End If

End Function

Function IsDashedLine(ByVal line)

    Dim t

    t = Trim(CStr(line))

    IsDashedLine = False

    If t = "" Then Exit Function

    t = Replace(t, " ", "")
    t = Replace(t, vbTab, "")
    t = Replace(t, "-", "")

    If t = "" Then
        IsDashedLine = True
    End If

End Function

Function IsAmountOnlyLine(ByVal line)

    Dim t
    Dim re

    t = Trim(CStr(line))

    IsAmountOnlyLine = False

    If t = "" Then Exit Function

    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^(<\d{1,3}(,\d{3})*(\.\d{2})>|<\d+(\.\d{2})>|-?\d{1,3}(,\d{3})*(\.\d{2})-?|-?\d+(\.\d{2})-?)$"

    If re.Test(t) Then
        IsAmountOnlyLine = True
    End If

End Function

'==============================================================
' Amount Helpers
'==============================================================
Function GetTrailingAmountsFromLine(ByVal line)

    Dim re, matches
    Dim amountText

    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True

    re.Pattern = "((\s+<\d{1,3}(,\d{3})*(\.\d{2})>|\s+<\d+(\.\d{2})>|\s+-?\d{1,3}(,\d{3})*(\.\d{2})-?|\s+-?\d+(\.\d{2})-?)+)\s*$"

    Set matches = re.Execute(CStr(line))

    If matches.Count = 0 Then
        GetTrailingAmountsFromLine = Split("", "|")
    Else
        amountText = Trim(matches(0).Value)
        GetTrailingAmountsFromLine = GetAmountsFromText(amountText)
    End If

End Function

Function GetAmountsFromText(ByVal text)

    Dim re, matches
    Dim arr()
    Dim i

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True

    re.Pattern = "<\d{1,3}(,\d{3})*(\.\d{2})>|<\d+(\.\d{2})>|-?\d{1,3}(,\d{3})*(\.\d{2})-?|-?\d+(\.\d{2})-?"

    Set matches = re.Execute(CStr(text))

    If matches.Count = 0 Then
        GetAmountsFromText = Split("", "|")
        Exit Function
    End If

    ReDim arr(matches.Count - 1)

    For i = 0 To matches.Count - 1
        arr(i) = matches(i).Value
    Next

    GetAmountsFromText = arr

End Function

Function GetLastTrailingAmount(ByVal line)

    Dim arr

    arr = GetTrailingAmountsFromLine(line)

    If UBoundSafe(arr) >= 0 Then
        GetLastTrailingAmount = arr(UBound(arr))
    Else
        GetLastTrailingAmount = ""
    End If

End Function

Function RemoveTrailingAmounts(ByVal text)

    Dim re
    Dim result

    result = CStr(text)

    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True

    re.Pattern = "((\s+<\d{1,3}(,\d{3})*(\.\d{2})>|\s+<\d+(\.\d{2})>|\s+-?\d{1,3}(,\d{3})*(\.\d{2})-?|\s+-?\d+(\.\d{2})-?)+)\s*$"

    result = re.Replace(result, "")

    RemoveTrailingAmounts = Trim(result)

End Function

Function CleanAmount(ByVal amt)

    Dim s
    Dim isNegative

    s = Trim(CStr(amt))
    isNegative = False

    If s = "" Then
        CleanAmount = ""
        Exit Function
    End If

    If Left(s, 1) = "<" And Right(s, 1) = ">" Then
        isNegative = True
        s = Mid(s, 2, Len(s) - 2)
    End If

    s = Replace(s, ",", "")

    If Right(s, 1) = "-" Then
        isNegative = True
        s = Left(s, Len(s) - 1)
    End If

    If Left(s, 1) = "-" Then
        isNegative = True
        s = Mid(s, 2)
    End If

    If isNegative Then
        CleanAmount = FormatNumberPlain(-CDblSafe(s))
    Else
        CleanAmount = FormatNumberPlain(CDblSafe(s))
    End If

End Function

Function CDblSafe(ByVal s)

    Dim v

    On Error Resume Next

    v = CDbl(Replace(CStr(s), ",", ""))

    If Err.Number <> 0 Then
        Err.Clear
        v = 0
    End If

    On Error GoTo 0

    CDblSafe = v

End Function

Function FormatNumberPlain(ByVal v)

    FormatNumberPlain = Replace(FormatNumber(CDblSafe(v), 2, -1, 0, 0), ",", "")

End Function

Function UBoundSafe(ByVal arr)

    Dim u

    On Error Resume Next

    u = UBound(arr)

    If Err.Number <> 0 Then
        Err.Clear
        u = -1
    End If

    On Error GoTo 0

    UBoundSafe = u

End Function

'==============================================================
' Compatibility No-op
'==============================================================
Sub AddException(ByVal rawLine, ByVal issueType)

    'No-op by design.
    'Flattened output only.
    'This prevents SUMMARY / PARSE EXCEPTIONS sections from printing.

End Sub