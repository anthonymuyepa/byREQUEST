Option Explicit

Dim fout, saveEffect, saveGLAcct, savePost, saveSequence, saveUsr, SaveCat, saveReference, saveComment, saveDebits, saveCredits, saveBalance
Dim lineData, columnHeaders, dataLineWidths, formattedLine, Header, saveAcct, saveGLname
Dim saveBeginBal

Header = 0

' Initialize variables
saveEffect = ""
savePost = ""
saveSequence = ""
saveUsr = ""
SaveCat = ""
saveGLAcct = ""
saveReference = ""
saveComment = ""
saveDebits = ""
saveCredits = ""
saveBalance = ""

' Define column headers and widths
columnHeaders = "Effect   GL Account       Post     Sequence Usr  Cat Reference  Description/Comment                                         Debits           Credits          Balance"
dataLineWidths = Array(9,17,9, 9, 5, 4, 11, 60, 17, 17, 18)


Sub StartDoc()
	sRec = GetLine
	Srec = replace(sRec, "Page     1","")
	fout.WriteLine sRec
	sRec = GetLine
	fout.WriteLine sRec
	sRec = GetLine
	fout.WriteLine sRec
    fout.WriteLine columnHeaders
End Sub

Private Sub ProcessLine
    ' Ignore dashed lines
    If InStr(sRec, "--------") > 0 Then 
        sRec = GetLine
    End If
    
    ' Ignore unwanted summary lines
    If InStr(sRec, "Average and Ending Balance") > 0 Then 
        sRec = GetLine
    End If


    ' Check if the line is a data line
    If IsDate(Mid(sRec, 1, 8)) And Len(sRec) > 129 Then
		'MsgBox "LongLine: " & sRec
        If InStr(Mid(sRec, 10, 8), "/")>0 Then
		     'MsgBox "DataLine: " & sRec
            ' Extract data columns
            saveEffect = Trim(Mid(sRec, 1, 8))
			
            savePost = Trim(Mid(sRec, 10, 8))
            saveSequence = Trim(Mid(sRec, 19, 6))
            saveUsr = Trim(Mid(sRec, 26, 4))
			SaveCat = Trim(Mid(sRec, 31, 3))
            saveReference = Trim(Mid(sRec, 35, 9))
            saveComment = Trim(Mid(sRec, 46, 32))
            saveDebits = Trim(Mid(sRec, 78, 16))
            saveCredits = Trim(Mid(sRec, 97, 16))
            saveBalance = TRim(Mid(sRec, 116))
            
		
			' Handle cases where comments spill into debit space or saveDebits contains non-numeric text
			If IsNumeric(Left(saveDebits, 1)) = False Then
				Dim splitPos
				' Check if saveDebits contains a space to split
				splitPos = InStrRev(saveDebits, " ") ' Find last space from right in saveDebits
				If splitPos > 0 Then
					Dim potentialDebits
					potentialDebits = Trim(Mid(saveDebits, splitPos + 1))
					' Check if the portion after the last space is numeric
					If IsNumeric(potentialDebits) Then
						saveDebits = potentialDebits  ' Retain numeric part as debits
						saveComment = saveComment & " " & Left(saveDebits, splitPos) ' Append the rest to comments
					Else
						' Treat the entire saveDebits as part of saveComment
						saveComment = saveComment & " " & saveDebits
						saveDebits = ""
					End If
				Else
					' If no space exists, treat the entire saveDebits as part of saveComment
					saveComment = saveComment & " " & saveDebits
					saveDebits = ""
				End If
			End If
			
			
            ' Format the line into columns
            formattedLine = FormatLine(Array(saveEffect, "",savePost, saveSequence, saveUsr, SaveCat, saveReference, saveComment, saveDebits, saveCredits, saveBalance), dataLineWidths)
            fout.WriteLine formattedLine
           
        ElseIf InStr(Mid(sRec, 10, 8), "-")>0 Then
            ' Extract balance information
			
            saveEffect = Trim(Mid(sRec, 1, 8))
            saveGLAcct = Trim(Mid(sRec, 10, 16))
            saveComment = Trim(Mid(sRec, 27, 67)) '& Space(dataLineWidths(7) - len(Trim(Mid(sRec, 27, 67))) - len(Trim(Mid(sRec, 96, 17))) -2) & Trim(Mid(sRec, 96, 17))
            saveBalance = Trim(Mid(sRec, 116))
		
            formattedLine = FormatLine(Array(saveEffect,saveGLAcct, "", "", "", "","", saveComment, "","", saveBalance), dataLineWidths)
            fout.WriteLine formattedLine

            
        Else
            
        End If
		
    End If
	
	If InStr (sRec,"Grand Totals") > 0 Then
        saveComment = Mid(sRec,1,14)
	    saveDebits = (Mid(sRec, 78, 16))
        saveCredits =(Mid(sRec, 97, 16))
  
	    formattedLine = FormatLine(Array("", "","", "", "","", "", saveComment, saveDebits,saveCredits, ""), dataLineWidths)
            fout.WriteLine formattedLine	   
    End If
	
End Sub

Private Function FormatLine(dataArray, widths)
    Dim i, output
    output = ""
    For i = 0 To UBound(dataArray)
        output = output & PadRight(dataArray(i), widths(i))
    Next
    FormatLine = output
End Function

Private Function PadRight(input, width)
    PadRight = Left(input & Space(width), width)
End Function

Private Sub CloseDoc()
    ' Add any necessary cleanup operations here
End Sub
