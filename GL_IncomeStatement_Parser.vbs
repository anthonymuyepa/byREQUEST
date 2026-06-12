Option Explicit

' =========================================================
' GL Income Statement Extract
'
' Purpose:
'   - Read the report header from the first few lines
'   - Extract:
'       1) EffectiveDate from header period text such as
'          "for February 2026", converted to month-end date
'       2) Seq number from header text like "Seq 331396"
'   - Output qualifying detail rows
' =========================================================

Dim gHeaderPart
Dim gReportDate, gSeqNumber, gHeaderParsed, gTargetMetric

Sub StartDoc()

    Dim i, line

    ' init per run
    gReportDate = ""
    gSeqNumber = ""
    gHeaderParsed = False
    gHeaderPart = ""

    fout.WriteLine "Report_Date  Seq     GL Account       Name                                            Mth to Date   Gross%   Qtr to Date   Gross%   Year to Date   Gross%"

    gTargetMetric = "MTD"   ' reserved for future use

    ' ---------------------------------------------------
    ' Read only the first few lines to find the header
    ' ---------------------------------------------------
    For i = 1 To 3
        line = GetLine
        
        If Not gHeaderParsed Then
            If InStr(1, line, "GL Income Statement", vbTextCompare) > 0 And _
               InStr(1, line, "Seq", vbTextCompare) > 0 Then

                gHeaderPart = line

                ' Extract month-end EffectiveDate from:
                ' "GL Income Statement for February 2026"
                gReportDate = GetEffectiveDate(line)

                ' Extract Seq number from:
                ' "Seq 331396"
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

    ' Skip decorative/footer lines
    If InStr(1, srec, "* * *", vbTextCompare) > 0 Then Exit Sub

    ' Process only rows that start with a numeric GL account
    If IsNumeric(Left(trimmedLine, 1)) Then
       If Len(srec) > 120 Or IsGLCategoryHeader(trimmedLine) Then
			fout.WriteLine gReportDate & "  " & gSeqNumber & "  " & srec
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

Function IsGLCategoryHeader(ByVal sLine)

    Dim firstToken
    Dim restOfLine

    IsGLCategoryHeader = False

    sLine = Trim(sLine)
    If sLine = "" Then Exit Function

    firstToken = Split(sLine, " ")(0)

    ' Category/title rows requested by client use suffix -8888.
    If Right(firstToken, 5) <> "-8888" Then Exit Function

    ' Make sure it has a description after the GL account.
    restOfLine = Trim(Mid(sLine, Len(firstToken) + 1))
    If restOfLine = "" Then Exit Function

    IsGLCategoryHeader = True

End Function