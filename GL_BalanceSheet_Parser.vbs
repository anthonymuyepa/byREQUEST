Option Explicit

' =========================================================
' GL Balance Sheet Extract
' Purpose:
'   - Read the report header from the first few lines
'   - Extract:
'       1) EffectiveDate = end of month from text like "for February 2026"
'       2) Seq number from header text like "Seq 331395"
'   - Write qualifying detail lines with:
'       Report_Date, Seq, and original account line
'
' Notes:
'   - This version uses the accounting period in the header
'     (e.g. "February 2026") and converts it to month-end date.
'   - MsgBox removed so the script can run unattended.
' =========================================================

Dim gHeaderPart
Dim gReportDate, gSeqNumber, gHeaderParsed, gTargetMetric

Sub StartDoc()

    Dim i, line

    ' -----------------------------
    ' Initialize per run
    ' -----------------------------
    gReportDate = ""
    gSeqNumber = ""
    gHeaderParsed = False
    gTargetMetric = "MTD"   ' Reserved for future use

    fout.WriteLine "Report_Date  Seq     GL Account  Name                                                               Amount"

    ' ---------------------------------------------------
    ' Read the first few lines only to locate the header
    ' ---------------------------------------------------
    For i = 1 To 3
        line = GetLine

        If Not gHeaderParsed Then
            If InStr(1, line, "GL Balance Sheet", vbTextCompare) > 0 And _
               InStr(1, line, "Seq", vbTextCompare) > 0 Then

                gHeaderPart = line

                ' Extract month-end EffectiveDate from:
                ' "GL Balance Sheet for February 2026"
                gReportDate = GetEffectiveDate(line)

                ' Extract Seq number from:
                ' "Seq 331395"
                gSeqNumber = GetSeqNumber(line)

                gHeaderParsed = True
                Exit For
            End If
        End If
    Next

End Sub

Sub ProcessLine()

    Dim trimmedLine
    trimmedLine = Trim(srec)

    If trimmedLine = "" Then Exit Sub

    ' Only process lines that appear to start with an account number
    If IsNumeric(Left(trimmedLine, 1)) Then

        ' Keep normal expected detail rows.
        ' Also keep category/title rows such as 100000-8888 even if they have no amount.
        If (Len(srec) > 70 And Len(srec) < 100) Or IsGLCategoryTitleRow(trimmedLine) Then

            fout.WriteLine _
                PadRight(gReportDate, 12) & " " & _
                PadRight(gSeqNumber, 7) & " " & _
                srec

        End If

    End If

End Sub

Function GetEffectiveDate(ByVal sLine)
    Dim re, m, monthYearText, dtFirst, dtEnd

    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "for\s+([A-Za-z]+\s+\d{4})"
    re.IgnoreCase = True
    re.Global = False

    If re.Test(sLine) Then
        Set m = re.Execute(sLine)(0)
        monthYearText = Trim(m.SubMatches(0))

        dtFirst = CDate("1 " & monthYearText)
        dtEnd = DateSerial(Year(dtFirst), Month(dtFirst) + 1, 0)

        GetEffectiveDate = Right("0" & Month(dtEnd), 2) & "/" & _
                           Right("0" & Day(dtEnd), 2) & "/" & _
                           Year(dtEnd)
    Else
        GetEffectiveDate = ""
    End If
End Function

Function GetSeqNumber(ByVal sLine)
    Dim re, matches

    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "\bSeq\s+(\d+)\b"
    re.IgnoreCase = True
    re.Global = False

    If re.Test(sLine) Then
        Set matches = re.Execute(sLine)
        GetSeqNumber = Trim(matches(0).SubMatches(0))
    Else
        GetSeqNumber = ""
    End If
End Function

Function PadRight(ByVal s, ByVal n)
    If IsNull(s) Then s = ""
    s = CStr(s)
    PadRight = Left(s & Space(n), n)
End Function


Function IsGLCategoryTitleRow(ByVal sLine)

    Dim firstToken
    Dim restOfLine

    IsGLCategoryTitleRow = False

    sLine = Trim(sLine)
    If sLine = "" Then Exit Function

    firstToken = Split(sLine, " ")(0)

    ' Category/title rows requested by client use suffix -8888.
    If Right(firstToken, 5) <> "-8888" Then Exit Function

    ' Make sure it has a description after the GL account.
    restOfLine = Trim(Mid(sLine, Len(firstToken) + 1))
    If restOfLine = "" Then Exit Function

    IsGLCategoryTitleRow = True

End Function