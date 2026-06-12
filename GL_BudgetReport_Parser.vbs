Option Explicit

' =========================================================
' GL Budget Report Extract
' Author: Anthony Muyepa
'
' Purpose:
'   - Read GL Budget Report header
'   - Extract EffectiveDate from period text such as:
'       "GL Budget Report for February 2026"
'     and convert it to month-end date (02/28/2026)
'   - Extract Seq number from header
'   - Output qualifying detail rows
'   - Handle continuation lines by reusing the left portion
'     of the previous full detail row
' =========================================================

Dim gHeaderPart
Dim gReportDate, gSeqNumber, gHeaderParsed, countt, currentRecord

Sub StartDoc()
    gReportDate = ""
    gSeqNumber = ""
    gHeaderParsed = False
    gHeaderPart = ""
    countt = 0
    currentRecord = ""

    fout.WriteLine "Report_Date  GL Account  Name                                     Prd               Actual                Budget              Variance       Var%"
End Sub

Sub ProcessLine()
    Dim line, trimmedLine

    line = srec
    trimmedLine = Trim(line)

    If trimmedLine = "" Then Exit Sub

    ' ---------------------------------------------------
    ' 1) Extract Report Date and Seq from report header
    ' ---------------------------------------------------
    If Not gHeaderParsed Then
        If InStr(1, line, "GL Budget Report", vbTextCompare) > 0 And _
           InStr(1, line, "Seq", vbTextCompare) > 0 Then

            gHeaderPart = line

            ' Extract month-end date from "for February 2026"
            gReportDate = GetEffectiveDate(line)

            ' Extract Seq number from "Seq 331398"
            gSeqNumber = GetSeqNumber(line)

            gHeaderParsed = True
        End If

        Exit Sub
    End If

    ' ---------------------------------------------------
    ' 2) Skip divider/header/footer lines
    ' ---------------------------------------------------
    If InStr(1, line, "-----", vbTextCompare) > 0 Then Exit Sub
    If InStr(1, line, "====", vbTextCompare) > 0 Then Exit Sub
    If InStr(1, line, "State Dept FCU", vbTextCompare) > 0 Then Exit Sub
    If InStr(1, line, "GL Account", vbTextCompare) > 0 Then Exit Sub
    If InStr(1, line, "Variance       Var", vbTextCompare) > 0 Then Exit Sub
    If InStr(1, line, "Page", vbTextCompare) > 0 Then Exit Sub

	' ---------------------------------------------------
	' 3) Full detail row or category/title row
	'    Expected to begin with numeric GL account
	' ---------------------------------------------------
	If IsNumeric(Left(trimmedLine, 1)) Then

		' Normal budget detail rows are long because they include
		' Actual, Budget, Variance, and Var%.
		If Len(line) > 120 Then
			currentRecord = Left(line, 52)
			fout.WriteLine gReportDate & "   " & line
			countt = countt + 1
			Exit Sub
		End If

		' Category/title rows such as 118500-8888 may have no values,
		' so they are shorter but still need to be preserved for layout.
		If IsGLCategoryTitleRow(trimmedLine) Then
			currentRecord = Left(line, 52)
			fout.WriteLine gReportDate & "   " & line
			countt = countt + 1
			Exit Sub
		End If

	End If
	
    ' ---------------------------------------------------
    ' 4) Continuation line
    '    First 52 chars are blank, reuse prior left side
    ' ---------------------------------------------------
    If Len(line) > 120 Then
        If Trim(Left(line, 52)) = "" And Len(currentRecord) > 0 Then
            fout.WriteLine gReportDate & "   " & currentRecord & Mid(line, 53)
            countt = countt + 1
            Exit Sub
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