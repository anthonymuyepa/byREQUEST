Rem               +===========================================================+
Rem               |                                                           |
Rem               |    JAX FCU                                                |
Rem               |    ACH Debits Exceptions NSF                       	      |
Rem               |  Edited by ZTS on 9/22/15 to add header and smart spacing |    
Rem               | CDS 10-19-2016 addtl line Overdraw Protection             |  
Rem               | CDS 6-25-2020 Parse out Total line, set User1 to title    |       
Rem               +===========================================================+


'--------------------------------------------------

Dim EndReport, HeadText, First,NewLine,DiscRec,OdCt,CtOut,holdTotals
Dim OdProtect(20)          ' Array to store Overdraft Protection info


HeadText = "Account   S/L/X ID  Name             Tran Desc       Company Name           Amount     Available       Balance   Description                Company ID   Settlement Transmission ECC Disc                    Ent           Account       ID               Name                    Trc Number         OD Protect   ID            Available       Balance" 
First = 0
OdCt = 0

Private Sub ProcessLine
   	
		If First = 0 Then
		'	SetUserVariablesBasedOnFilename (First) 
                        fout.writeline sRec                                    ' Added 10-19-2016 to include first line of page heading
                  Spoolfile.user1=Trim(Mid(sRec,42,35))	     ' added 6-16-2020 to name folder s and files correctly		
                        fout.WriteLine HeadText
		 	First = 1
		End If 
                  
                 If InStr(sRec,"JAX FEDERAL Credit Union") > 0 Then             ' This is heading
                    If First > 0 Then
                       sRec=Getline
                       sRec=Getline
                       sRec=Getline                                              ' Skip headings on subsequent pages
                    End If
                 End If


 		 If InStr(sRec, "----------------------------------------------------------------------------------------------------------") > 0     Then  
                     sRec=Getline
           End If

                 If IsNumeric(Mid(sRec,1,10)) Then    ' Start of next record... 				
 	       '  msgbox sRec
                        NewLine=Mid(sRec,1,11) & Space(2) & Mid(sRec,12,1) & Space(2) & Mid(sRec,14,2) & Space(2) & Mid(sRec,17,91) & Space(2) & Mid(sRec,108,17)      ' This line has account number
                        NewLine = NewLine & Space(136 - Len(newLine))   ' Pad out line, some are longer
 	         End If

                 If  InStr(sRec,"Overdraw Protection") > 0 Then
               ' msgbox "OD " & sRec
                          ODProtect(OdCt)=Mid(sRec,57,51)   ' Note - there may be more than one
                          OdCt=OdCt + 1
                 End If

                If InStr(sRec,"Company ID") > 0 Then
                   NewLine = NewLine & Space(4) & Mid(sRec,12,11) & Space(4) & Mid(sRec,34,8) & Space(4) & Mid(sRec,56,8) & Space(2) & Mid(sRec,69,3) & Space(1)
                End If 

                If InStr(sRec,"Disc ") > 0 Then
                   FormatDisc
                End If
 		
		If InStr(sRec, "Total Exceptions") > 0 Then ' but check to see if we are are at the bottom..  
 			         EndReport = 1
                    formatTotals
                
 		End If            
 			
End Sub

Sub FormatDisc()
    DiscRec=Mid(sRec,6,20) & Space(4) & Mid(sRec,31,10) & Space(4) & Mid(sRec,47,12) & Space(2)
    If InStr(sRec," ID ") > 0 then
 ' msgbox "here1 " & sRec
       DiscRec=DiscRec & Mid(sRec,68,15) & Space(2) & Mid(sRec,88,22) & Space(2) & Mid(sRec,115,15)
      Else
 ' msgbox "Here2 " & sRec
       DiscRec=DiscRec & Space(17) & Mid(sRec,69,22) & Space(2) & Mid(sRec,96,15)                 ' This line does not have an ID so data is pushed to left.
   End If

   NewLine=NewLine & DiscRec 
   If OdCt < 1 Then 
      fout.Writeline NewLine
      Else
         OdCt = OdCt - 1 
         For CtOut = 0 to OdCt
             fout.Writeline NewLine & space(4) & OdProtect(CtOut)
         Next
   End If

   OdCt = 0

End Sub

Sub formatTotals        ' This is done at end of file

    holdTotals=sRec
    sRec=Getline

    fout.Writeline Space(20) & "Total Exceptions" & Mid(sRec,12,5)
    fout.Writeline Space(20) & Mid(holdTotals,22,15) & Mid(sRec,22,15) & Space(3) & Mid(holdTotals,42,8) & Space(1) & Mid(sRec,42,8)
    fout.Writeline Space(20) & Mid(holdTotals,56,15) & Space(2) & Mid(sRec,56,15) & Mid(holdTotals,74,8) & Space(2) & Mid(sRec,74,8)
    fout.Writeline Space(20) & Mid(holdTotals,89,15) & Space(3) & Mid(sRec,89,15) & Mid(holdTotals,106,8) & Space(1) & Mid(sRec,106,8)
    fout.Writeline Space(20) & Mid(holdTotals,118,14) & Space(2) & Mid(sRec,118,14)

End Sub
