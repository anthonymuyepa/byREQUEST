Option Explicit
' Created: 
'   3-01-2017  Winship
'      Created for PIMA GL Accrual Exceptions for FASB 91


' Universal Variables
Dim FirstTitleHeader
Dim CalcMethod
Dim FirstColumnHeader
Dim FileSummary
Dim SummaryHeaders


sub processline
    
	Dim Line2
	Dim NewLine
	
	If FileSummary <> 1 Then ' If we aren't in the summary section
		
		' Only create one overall header
		If InStr(srec,"GL Accrual Exceptions for FASB 91") > 0 Then 
			If FirstTitleHeader <> 1 Then
				Newline = srec
				FirstTitleHeader = 1
				fout.writeline NewLine
			End if
		
		'Create an array from Calc Method to add it in as a column
		Elseif InStr(srec,"Calculation Method:") > 0 Then 
			CalcMethod = split(srec, ": ")
		
		' Only print the first header, add Calc Method and join both lines
		Elseif InStr(srec,"Account    Name") > 0 AND FirstColumnHeader <> 1 Then 
			srec = replace(srec, space(68),"") ' remove the large spacing between the 2nd and 3rd columns
			Line2 = GetLine()
			NewLine = CalcMethod(0) & space(2) & srec & space(60 - len(srec)) & Line2
			FirstColumnHeader = 1
			fout.writeline NewLine
			
		' If it says loan type, jump to summary section coding	
		Elseif 	InStr(srec,"Loan Type") > 0 Then
			FileSummary = 1
			
		' Add Calc Method and join both lines of data	
		Elseif InStr(srec,"00") > 0 Then
			srec = replace(srec, space(68),"")
			Line2 = GetLine()
			NewLine = CalcMethod(1) & space(20-len(CalcMethod(1))) & srec & space(60 - len(srec)) & Line2
			fout.writeline NewLine
			
		
		End If
	End if
	
	
	' If we have reached the summary section
	If FileSummary = 1 Then
		' If you hit a page break, ignore it from the header to the "------" break
		If InStr(srec,"GL Accrual Exceptions for FASB 91") Then 
			SummaryHeaders = 1
		Elseif InStr(srec,"-------------") Then
			SummaryHeaders = 0
		Elseif SummaryHeaders <> 1 Then ' If not a page header then print
			NewLine = srec
			fout.writeline NewLine	
		End if
	End if
	
end sub


