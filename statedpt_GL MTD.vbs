Dim gHeaderPart 

Dim gReportDate, gSeqNumber, gHeaderParsed, gTargetMetric

Sub StartDoc()

    ' init per run
    gReportDate = ""
    gSeqNumber = ""
    gHeaderParsed = False

    fout.WriteLine "ReportDate Seq  GL Account  Name                                     Prd               Actual                Budget              Variance       Var%"

	gTargetMetric = "QTD" ' default


End Sub




Sub ProcessLine()

    Dim line, trimmed
    line = srec
    If Trim(line) = "" Then Exit Sub
    trimmed = Trim(line)

    ' ---------------------------------------------------
    ' 1) Extract Report Date and Seq from the report header line (first page)
    ' ---------------------------------------------------
    If Not gHeaderParsed Then
        
        If InStr(1, line, "GL Budget Report", vbTextCompare) > 0 And _
           InStr(1, line, "Seq", vbTextCompare) > 0 Then
            
            ' Extract date after "at"  -> 09/29/12
            Dim posAt, tempStr, parts
            posAt = InStr(1, line, " at ", vbTextCompare)
            If posAt = 0 Then posAt = InStr(1, line, "at", vbTextCompare)

            If posAt > 0 Then
                tempStr = Mid(line, posAt + 3)
                parts = Split(Trim(tempStr), " ")
                If UBound(parts) >= 0 Then gReportDate = parts(0)
            End If

            ' Extract Seq number -> 7412
            Dim posSeq, seqPart, seqParts
            posSeq = InStr(1, line, "Seq", vbTextCompare)
            If posSeq > 0 Then
                seqPart = Mid(line, posSeq + 4)
                seqParts = Split(Trim(seqPart), " ")
                If UBound(seqParts) >= 0 Then gSeqNumber = Trim(seqParts(0))
            End If

            gHeaderParsed = True
        End If

        Exit Sub
    End If


	' 2) Header line: starts with number
	If IsNumeric(Left(trimmed, 1)) Then

		gHeaderPart = ExtractHeaderPart(trimmed)

		' If the target metric is on THIS SAME header line (MTD often is),
		' output immediately.
		If InStr(1, trimmed, gTargetMetric, vbTextCompare) > 0 Then
			fout.WriteLine gReportDate & "  " & gSeqNumber & "  " & gHeaderPart & " " & ExtractMetricPart(trimmed, gTargetMetric)
		End If

		Exit Sub
	End If

    ' ---------------------------------------------------
    ' 3) gTargetMetric line: output Date + Seq + HeaderPart + gTargetMetric line
    ' ---------------------------------------------------
	If InStr(1, trimmed, gTargetMetric, vbTextCompare) > 0 Then
		If gHeaderPart <> "" Then
			fout.WriteLine gReportDate & "  " & gSeqNumber & "  " & gHeaderPart & " " & trimmed
		Else
			fout.WriteLine gReportDate & "  " & gSeqNumber & "  " & trimmed
		End If
	End If

End Sub


Function ExtractHeaderPart(ByVal headerLine)
    Dim t, p
    t = headerLine

    p = InStr(1, t, " MTD", vbTextCompare)
    If p = 0 Then p = InStr(1, t, " QTD", vbTextCompare)
    If p = 0 Then p = InStr(1, t, " YTD", vbTextCompare)
    If p = 0 Then p = InStr(1, t, " Full Year", vbTextCompare)

    If p > 0 Then
        ExtractHeaderPart = Left(t, p - 1)
    Else
        ExtractHeaderPart = t
    End If
End Function



Function ExtractMetricPart(ByVal s, ByVal metric)
    Dim p, t
    t = s

    p = InStr(1, t, metric, vbTextCompare)
    If p > 0 Then
        ExtractMetricPart = Mid(t, p)
    Else
        ExtractMetricPart = ""
    End If
End Function