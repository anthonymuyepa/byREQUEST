' VBScript to merge Dealer Name with numeric values (into one line)
' Author: AnthonyM
' Updated: 10-27-2025 numeric line is one after dealer line. (not before)

Dim headerln1, headerln2 ' Header lines for output
Dim dealerLine, numericLine, LastLine, thisdealer


' Initialize variables

numericLine = ""
dealerLine = ""
thisdealer = 0

' Define header lines for the output file
headerln1 = "Dealer                                                 Credit Score      Credit Score      Credit Score      Credit Score      Credit Score      Credit Score      Credit Score"
headerln2 = "                                                       A=690 & Up        B=650-689         C=620-649         D=590-619         E=001-589           F=000           A thru F    "

Sub StartDoc()
    ' Initialize document processing - skip first 8 lines and write headers
    
    ' Read and write the first line
    sRec = Getline()
    If Not IsEmpty(sRec) Then
        fout.writeline sRec
    End If
    
    ' Skip lines 2 through 8 (read but don't process)
    Dim i
    For i = 2 To 8
        sRec = Getline()
        ' Optional: Add logging here if needed
    Next
    
    ' Write header lines
    fout.writeline headerln1
    fout.writeline headerln2
    
    ' Reset dealer tracking
    thisdealer = 0
    dealerLine = ""
    numericLine = ""
End Sub

Public Sub processline()
        
    ' Identify dealer line based on specific criteria:
    ' - Length less than 45 characters
    ' - Not empty after trimming
    ' - Doesn't contain ".DLRS2", "|", or "_" characters
    If Len(srec) < 45 And Len(srec) > 1 And _
       InStr(srec, ".DLRS2") < 1 And _
       InStr(srec, "|") < 1 And _
       InStr(srec, "_") < 1 Then
        
        dealerLine = srec
        thisdealer = 1
        Exit Sub
    End If
    
    ' Process numeric data line if we have an active dealer
    If thisdealer > 0 Then
        ' Check if line contains numeric data with decimal and pipe characters
        If InStr(srec, ".") > 0 And InStr(srec, "|") > 0 And InStr(srec, ".DLRS2") < 1 Then
            ' Clean up the numeric line by replacing pipes with spaces
            numericLine = Replace(srec, "|", " ")
            
            ' Calculate padding to align columns (45 characters for dealer name)
            Dim padding
            padding = 45 - Len(dealerLine)
            If padding < 0 Then padding = 0
            
            ' Write combined dealer and numeric data
            fout.writeline dealerLine & Space(padding) & numericLine
            
            ' Reset dealer tracking
            thisdealer = 0
            dealerLine = ""
            numericLine = ""
            Exit Sub
        End If
    End If
	
	If thisdealer < 1 Then
        ' Check if line contains numeric data with decimal and pipe characters
        If InStr(srec, ".") > 0 And InStr(srec, "|") > 0 And InStr(srec, ".DLRS2") < 1 Then
			LastLine = Replace(srec, "|", " ")
		End if
		
	End if
	
    
End Sub


Sub CloseDoc
	' Uncoment line below to print the last summary line if required
	'fout.writeline space(45) & LastLine
End Sub
