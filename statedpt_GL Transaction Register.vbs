Rem               +===========================================+
Rem               |                                           |
Rem               |        GL Transaction Register            |
Rem               |    includes GL name & number in line      |
Rem               |           Written July 22, 2013           |
Rem               |                                           |
Rem               | Revised Aug 4  , 2017 Miles Winship       |
Rem               |Edited PeriodSpot to prevent breakage      |
Rem               | Revised July 31, 2017 Miles Winship       |
Rem               |Adjusted spacing for totals                |
Rem               | Revised July 25, 2017 Miles Winship       |
Rem               |Altered periodSpot to not work on comment  |
Rem               | Revised July 24, 2017 Miles Winship       |
Rem               | 	Notes: Included Ending Balance        |
Rem               |Included inactive accounts, Added Beginning|
Rem               |Balance, Added account date, Fixed spacing |
Rem               |between comment and debit columns, added   |
Rem               |additional comments                        |
Rem               | Revised July 31, 2013 Unknown modifictions|
Rem               | Revised Sept 11, 2013 Comment field wider |
Rem               | Revised April 21, 2014 Comment field too  |
Rem               |  wide for column, if no DB move fields    |
Rem				  | Revised June 17, 2014 ZTS: Comment field  |
Rem 			  |  bleeding into debits fixed               |
Rem               +===========================================+


Option Explicit


Dim Header, saveAcct, saveGLname,saveBeginBal,nameLen, Newline, Firstline, Comment, periodSpot, SaveDate
Dim count
Dim FirstRec
FirstRec = 0

Const ReplaceChr = " "
saveAcct=""
saveGLname=""

Header = 0

Private Sub ProcessLine
	SetUserVariablesBasedOnFilename (FirstRec)
	FirstRec = 1
	'fout.WriteLine sRec 'Disabled
    
	Dim Month
	Dim Variable4
	Month = Array("Jan","Feb","Mar","Apr","May","June","Jul","Aug","Sep","Oct","Nov","Dec")
	Variable4 = Cint(spoolfile.User4)
	spoolfile.User8 = Month(Variable4-1)
	 
    If Instr(sRec," Ending Balance ") > 0 Then 'If line contains "Ending balance" space and print the line 
       'sRec = GetLine 'Disabled
	   srec = replace(srec,space(20),"")
	   fout.WriteLine left(srec,8) & space(123) & Right(srec, Len(srec) - 8)
    End If

    If Instr(sRec,"GL Month End Closing") > 0 Then ' Ignore this line(from previous revisions, current file did not contain such text)
       sRec = GetLine
    End If

    If InStr (sRec,"Grand Totals") > 0 Then    'ashcraft changed 109 to 139 to get totals to line up
       fout.Writeline SPACE (132) & Mid(sRec,1,16) & SPACE(32) & Mid(sRec,78,38)             ' Write Grand Total line with spaces padding to align
    End If
    
    If InStr(sRec, "----------") > 0 Then                                     ' Eliminate lines with dashes
         Header = 1                                                           ' Flag to indicate done with page headings
         'sRec = GetLine 'Disabled
    End If

    If InStr(sRec,"Beginning Balance") > 0 Then                                ' Store Date, GL Name, Account number, and Beginning Balance from this line
         SaveDate = left(sRec,8)
		 saveAcct = Mid(sRec,10,16)
         saveGLname = Trim(Mid(sRec,27,34))
         nameLen = Len(saveGLname)
         saveBeginBal = trim(Mid(sRec,113,20))
         sRec=getLine
        
		if instr(srec,"  No Activity ") then	' If account has no activity go ahead and print the line with proper spacing
			fout.writeline SaveDate & space(2) & saveAcct & SPACE(2) & saveGLname & SPACE(36 - nameLen) & space(19-len(saveBeginBal)) & saveBeginBal & SPACE(49) & trim(srec)
		end if
		
    End If 

  If IsDate (Mid(sRec,10,8)) Then          ' If the line contains a record
	'msgbox srec	
		sRec = RTrim(sRec)
		If Len(sRec) < 90 Then	' For when the data gets broken between 2 lines due to overlapping comment and debit number, retrieve all the record's data and print
			Firstline = sRec
			sRec = GetLine
			sRec = Trim(sRec)
			Comment = Mid(Firstline,18,68)
			Comment = RTrim(Comment)
			Newline = SaveDate & space(2) & saveAcct & SPACE(2) & saveGLname & SPACE(36 - nameLen) & space(19-len(saveBeginBal)) & saveBeginBal & SPACE(2) & Mid(Firstline,1,17) & SPACE(2) & Comment & Space(82 - Len(Comment)) & sRec
		Else 	' For normal records
			Comment = Mid(sRec,18,63)
			Comment = RTrim(Comment)
			If Len(Comment) > 61 Then	' Fix spacing between comment and debit number, print record
				Dim NumericPeriod
				Set NumericPeriod = New RegExp	
				NumericPeriod.pattern = "\d\.\d"
				Comment = Mid(sRec, 81, 15)
				if NumericPeriod.test(mid(Comment,10)) = "True" then
					periodSpot = InStrRev(Comment, ".")
				end if
				count = periodSpot - 1
				If periodSpot > 1 Then
					Do While Mid(Comment, count, 1) <> " "
					count = count - 1
					Loop
					Newline = SaveDate & space(2) & saveAcct & SPACE(2) & saveGLname & SPACE(36 - nameLen) & space(19-len(saveBeginBal)) & saveBeginBal & SPACE(2) & Mid(sRec,1,17) & SPACE(2) & Mid(sRec,18,63 + count) & Space(16) & Mid(sRec,81 + count,52)
				Else
					Newline = SaveDate & space(2) & saveAcct & SPACE(2) & saveGLname & SPACE(36 - nameLen) & space(19-len(saveBeginBal)) & saveBeginBal & SPACE(2) & Mid(sRec,1,17) & SPACE(2) & Mid(sRec,18,73) & Space(16) & Mid(sRec,91,42)
				End If	
			Else
				Newline = SaveDate & space(2) & saveAcct & SPACE(2) & saveGLname & SPACE(36 - nameLen) & space(19-len(saveBeginBal)) & saveBeginBal & SPACE(2) & Mid(sRec,1,17) & SPACE(2) & Mid(sRec,18,63) & Space(16) & Mid(sRec,81,52)
			End If
		End If
	
	fout.Writeline Newline
     End If 

 If Instr(sRec,"Effect") > 0 Then          ' Heading line, adjust alignment for additional fields 
    sRec = "Date" & space(6) & "GL Acct" & SPACE(11) & "GL Acct Name " & SPACE(26) & "Beginning Balance" & space(2) & Mid(sRec,1,14) & SPACE(3) & Mid(sRec,17,8)   _
              & Mid(sRec,25,55) & Space(17) & Mid(sRec,81,52)
 End If
  
 If Header < 1 Then 
    fout.Writeline sRec                    ' Write out the first page heading

 End If 

End Sub
