' VBScript to merge Dealer Name (4th line) with numeric values (2nd line with decimals)
' Author: AnthonyM

Dim lineBuffer(4)
Dim lineCount, dataline
Dim headerln1, headerln2  
Dim dealerLine, numericLine, LastLine

lineCount = 0
numericLine = ""
headerln1 = "Dealer                                                 Credit Score      Credit Score      Credit Score      Credit Score      Credit Score      Credit Score      Credit Score"
headerln2 = "                                                       A+=730 & Up       A=690-729         B=650-689         C=620-649         D=590-619         E=000-589         A thru F"

Sub StartDoc()
	sRec = Getline
	fout.writeline sRec
    Dim i
    For i = 2 To 7
        sRec = Getline
    Next
    fout.writeline headerln1
    fout.writeline headerln2
End Sub

Public Function processline


    ' If the line is short and not a header, treat it as dealer name
    If Len(srec) < 45 And Len(Trim(srec))> 1 and InStr(srec, "CREDITGRADELOAN.") < 1 And InStr(srec, "|") < 1 Then
        dealerLine = srec

        Exit Function
    End If


    ' If the line contains numeric data (with decimal), and it's not a header
		If (InStr(srec, ".") > 0 And InStr(srec, "|") > 0 And InStr(srec, ".CREDITGRADELOAN.") < 1) Then
			numericLine = numericLine & srec
			LastLine = sRec
			
			
			' Output merged dealer name and numeric values
			numericLine = replace(numericLine,"|"," ")
			numericLine = Replace(numericLine, vbCrLf, " ")
			numericLine = Replace(numericLine, vbCr, " ")
			numericLine = Replace(numericLine, vbLf, " ")
			fout.writeline dealerLine & space(45 - len(dealerLine)) & numericLine

			' Reset for next record
			dealerLine = ""
			numericLine = ""
		

		End If
		
	
	
End Function

Public Function closedoc
    ' Handles any cleanup, if necessary

End Function