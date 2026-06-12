Option Explicit

Dim gHeaderPart
Dim gReportDate, gSeqNumber, gHeaderParsed, countt, currentRecord

Sub StartDoc()
    gReportDate = ""
    gSeqNumber = ""
    gHeaderParsed = False
    gHeaderPart = ""
	countt=0
	currentRecord = " "
    fout.WriteLine "ReportDate GL Account  Name                                     Prd               Actual                Budget              Variance       Var%"
End Sub

Sub ProcessLine()
    Dim line
    line = srec
    If Trim(line) = "" Then Exit Sub
	
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
	
	
	
	
	
 

	

    If InStr(srec,"-----") > 1 Then Exit Sub
	If InStr(srec,"====") > 1 Then Exit Sub
	If InStr(srec,"State Dept FCU") > 1 Then Exit Sub
	If InStr(srec,"GL Account") > 1 Then Exit Sub
	If InStr(srec,"Variance       Var") > 1 Then Exit Sub
 
 
			

    If IsNumeric(Left(srec, 1)) and len(srec) > 120  Then
        currentRecord = Left(srec, 52)
        fout.WriteLine gReportDate & "   " & srec
    End If
    
    If Instr(srec,"                                                    " ) > 0 and len(srec) > 120 and len(currentRecord) > 2    Then
        'continuation line
        fout.WriteLine gReportDate & "   "& currentRecord & Mid(srec, 53)
    End If
	
End Sub

