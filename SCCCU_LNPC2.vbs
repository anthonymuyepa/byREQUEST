' ============================================================
' byREQUEST-Style Parser: Loan Coupons "Farthest Coupons Thru"
'
' Goal:
' - Extract MEMBER_NBR, LOAN_NBR, COUPONS_THRU_DATE that is the FARTHHEST COUPONS_THRU_DATE for each member/loan.
' ============================================================

Option Explicit

' ---------- Globals ----------
Dim gRx
Dim gHasCurrent
Dim gCurMember, gCurLoan
Dim gMaxDate  ' VB Date

' ---------- byREQUEST entry points ----------
Sub StartDoc()
    gHasCurrent = False
    gCurMember = ""
    gCurLoan = ""
    gMaxDate = Empty

    ' Regex finds: MEMBER LOAN DATE (MM-DD-YY) anywhere in the line
    ' Allows spaces/tabs/commas between tokens.
    Set gRx = New RegExp
    gRx.Global = False
    gRx.IgnoreCase = True
    gRx.Pattern = "\b(\d+)[,\t ]+(\d+)[,\t ]+(\d{2}-\d{2}-\d{2})\b"


    ' Optional header:
	fout.Write "LOAN" & vbLf
	fout.Write "MEMBER_NBR,LOAN_NBR,COUPONS_THRU_DATE" & vbLf

End Sub

Sub ProcessLine
    Dim m, memberNbr, loanNbr, dateStr, d

    If gRx.Test(sRec) Then
        Set m = gRx.Execute(sRec)(0)

        memberNbr = Trim(m.SubMatches(0))
        loanNbr   = Trim(m.SubMatches(1))
        dateStr   = Trim(m.SubMatches(2)) ' MM-DD-YY

        d = ParseMMDDYY(dateStr)
        If IsEmpty(d) Then Exit Sub  ' bad date; ignore safely

        ' If this is the first match ever, initialize
        If Not gHasCurrent Then
            gHasCurrent = True
            gCurMember = memberNbr
            gCurLoan = loanNbr
            gMaxDate = d
            Exit Sub
        End If

        ' Member changed -> flush prior record and start tracking new member
        If memberNbr <> gCurMember Then
            FlushCurrent
            gHasCurrent = True
            gCurMember = memberNbr
            gCurLoan = loanNbr
            gMaxDate = d
            Exit Sub
        End If

        ' Same member: loan is expected same, but we keep the latest date regardless
        If d > gMaxDate Then gMaxDate = d
    End If
End Sub

Sub CloseDoc()
    ' flush last tracked member
    If gHasCurrent Then FlushCurrent

    On Error Resume Next
    If Not fout Is Nothing Then fout.Close
    Set fout = Nothing
    Set gRx = Nothing
End Sub

' ---------- Helpers ----------
Sub FlushCurrent()
    If Not gHasCurrent Then Exit Sub
    fout.Write gCurMember & "," & gCurLoan & "," & FormatMMDDYYYY(gMaxDate)  & vbLf
End Sub

Function ParseMMDDYY(ByVal s)
    ' Expects "MM-DD-YY"
    ' Interprets YY as 2000+YY (e.g., 26 -> 2026)
    On Error Resume Next

    Dim mm, dd, yy, yyyy
    mm = CInt(Mid(s, 1, 2))
    dd = CInt(Mid(s, 4, 2))
    yy = CInt(Mid(s, 7, 2))

    yyyy = 2000 + yy
    ParseMMDDYY = DateSerial(yyyy, mm, dd)

    If Err.Number <> 0 Then
        Err.Clear
        ParseMMDDYY = Empty
    End If
    On Error GoTo 0
End Function

Function FormatMMDDYYYY(ByVal d)
    ' Always returns MM/DD/YYYY with leading zeros
    Dim mm, dd, yyyy
    mm = Right("0" & CStr(Month(d)), 2)
    dd = Right("0" & CStr(Day(d)), 2)
    yyyy = CStr(Year(d))
    FormatMMDDYYYY = mm & "/" & dd & "/" & yyyy
End Function
