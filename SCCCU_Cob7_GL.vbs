Option Explicit

Dim header, Report, Run_Date, Run_Hour
Dim additionalDetail, pendingOutputLine



header ="Report  Run_Date Run_Time GL_Number  Branch            Debit             Credit Sequence  Voucher  Teller  NBR_TYPE1   Mbr_Nbr   BR  SFX PTyp NBR_TYPE2 MI_Code Addl_Detail"

Sub StartDoc()
	fout.WriteLine header
	Dim i 
	For i = 1 to 4
			srec = getLine()
		' Single check (no loop needed if checking same string)
		If Instr(srec, "GENERAL LEDGER ACCOUNT") > 0 Then
			Report = Trim(Mid(srec, 1, 5))
			' Pad with spaces to make colums align
			Report = Report & Space(8 - Len(Report))	
			Run_Date = Trim(mid(srec, 11,8))
			Run_Date = Run_Date & Space(9 - Len(Run_Date))
			Run_Hour = Trim(mid(srec, 20,2))
			Run_Hour = Run_Hour & Space(6 - Len(Run_Hour))
			
			
		End If
		
	Next		
	
	additionalDetail = False
	pendingOutputLine = ""
	
End Sub






Sub ProcessLine()
    Dim TranNbrType1, Mbr_Nbr, Suffix, BR, Product_Type_Code, Tran_Nbr_Type2
    
    ' Check if we're expecting an additional detail line
    If additionalDetail Then
        ' Check if current line matches the additional detail pattern
        Dim isDetailLine
        isDetailLine = (Instr(1,srec, "                                                                     ") > 0 And  Mid(srec, 70, 1) <> " ")
        
        If isDetailLine Then
            ' Yes, this is the additional detail
            Dim addlDetail, paddingLength
            addlDetail = Mid(srec, 70)
            
            ' Pad to column 160
            paddingLength = 160 - Len(pendingOutputLine)
            If paddingLength < 0 Then paddingLength = 0
            
            fout.WriteLine pendingOutputLine & Space(paddingLength) & addlDetail
 
        Else
            ' No, this is not a detail line
            ' Write the pending line without detail
            fout.WriteLine pendingOutputLine
            
            ' Now process current line (which might be another pattern match)
            ' Reset flag and reprocess
            additionalDetail = False
            ' Re-enter with current line
            ' We'll handle this by setting a flag and processing in main loop
        End If
        additionalDetail = False
        'Exit Sub
    End If
	
		additionalDetail = False
		pendingOutputLine = ""
    
    ' Normal line processing
  
		If Len(srec) >= 15 Then
            If IsPatternMatch(Left(srec, 15)) Then

				if len(srec) <=95 Then
					Dim chk 
					chk = mid(srec, 88)
					srec = left(srec,87) & "  " & chk 
				end if
				
				pendingOutputLine = Report & Run_Date & Run_Hour & srec 
				additionalDetail = True
            End If
        End If
	
	
End Sub


Sub CloseDoc()
    On Error Resume Next
    fout.WriteLine pendingOutputLine
End Sub



Function IsPatternMatch(strToCheck)
    Dim i
    
    ' Check length first
    If Len(strToCheck) < 15 Then
        IsPatternMatch = False
        Exit Function
    End If
    
    ' Check positions 1-3: must be spaces
    For i = 1 To 3
        If Mid(strToCheck, i, 1) <> " " Then
            IsPatternMatch = False
            Exit Function
        End If
    Next
    
    ' Check positions 4-7: must be digits
    For i = 4 To 7
        If Not IsNumeric(Mid(strToCheck, i, 1)) Then
            IsPatternMatch = False
            Exit Function
        End If
    Next
    
    ' Check position 8: must be hyphen
    If Mid(strToCheck, 8, 1) <> "-" Then
        IsPatternMatch = False
        Exit Function
    End If
    
    ' Check positions 9-12: must be digits
    For i = 9 To 12
        If Not IsNumeric(Mid(strToCheck, i, 1)) Then
            IsPatternMatch = False
            Exit Function
        End If
    Next
    
    ' Check positions 13-15: must be spaces
    For i = 13 To 15
        If Mid(strToCheck, i, 1) <> " " Then
            IsPatternMatch = False
            Exit Function
        End If
    Next
    
    ' If all checks passed
    IsPatternMatch = True
End Function


