Dim gHeaderPart 

Dim gReportDate, gSeqNumber, gHeaderParsed, gTargetMetric

Sub StartDoc()

    ' init per run
    gReportDate = ""
    gSeqNumber = ""
    gHeaderParsed = False

    fout.WriteLine "Run_Date  Seq     GL Account  Name                                                               Amount"

    gTargetMetric = "MTD" ' default

    For i = 1 To 3
        line = GetLine ' Fixed: Assign GetLine result to line variable
        
        ' ---------------------------------------------------
        ' 1) Extract Report Date and Seq from the report header line (first page)
        ' ---------------------------------------------------
        If Not gHeaderParsed Then
            
            If InStr(1, line, "GL Balance Sheet", vbTextCompare) > 0 And _
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
        End If
    Next 

End Sub






Sub ProcessLine()


    If Trim(srec) = "" Then Exit Sub


    If IsNumeric(Left(Trim(srec), 1)) Then
      
		If len(srec)> 70 and len(srec) < 100 Then

				fout.WriteLine gReportDate & "  " & gSeqNumber & "  " & srec

			
		End If
	End If

End Sub
