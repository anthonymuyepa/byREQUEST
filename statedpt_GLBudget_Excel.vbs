Rem               +==========================================+
Rem               |                                          |
Rem               |    Symitar - GL Budget report            |    
Rem               |     Separate file into YTD only-         |
Rem               |       and MTD-QTD-YTD                    |
Rem               |     Written 6/2014   Hillary Software    |
Rem               |  This version adds a value to 3rd line   |
Rem               |    of heading so file can be burst       |
Rem               |  Last modified: 6/19/2014                |
Rem               +==========================================+

Option Explicit

Const OverwriteTrue = TRUE
Const ReplaceChr = " "
Dim Lnct,SpaceCnt,outRec,mFlag,FirstRec, FirstName
Dim Line3Head, saveHeading, saveHeading2
Dim saveAcct,SaveName,Full_Heading
Dim headRec,Elems
Dim ArrayPos(18)      ' Array for reformatting detail lines
Dim ArrayLen(18)
Dim ArryPos(18)       ' Array for reformatting first line of Accounts
Dim ArryLen(81)       
Dim newHeading, newHeading2
Dim rType

FirstName=0
rType="YTD"                      ' Value to search for if no 'Branch', chgs to 'MTD' on Income statement
Line3Head="Section Balance Sheet     "
FirstRec=0
SpaceCnt=0
Lnct=2   
mFlag=0                          ' 0=YTD Only; 1=MTD,QTY,YTD 
newHeading="GL Account  Name                                      Branch         ACTUAL MTD            BUDGET MTD               VAR MTD             ACTUAL QTD            BUDGET QTD               VAR QTD             ACTUAL YTD           BUDGET YTD              VAR YTD        ACTUAL Full Year    BUDGET Full Year       VAR Full Year"  
'newHeading2=newHeading & "BUDGET QTD               VAR QTD             ACTUAL YTD            BUDGET YTD              VAR YTD       ACTUAL Full Year    BUDGET Full Year         VAR Full Year"                                                 

Full_Heading=Space(71) & "                      - MTD -                                                            - QTD -                            "
Full_Heading=Full_Heading & "                               - YTD                                                             Full Year                   "                                                 

ArrayPos(0)=1                    ' Setting for Re-arrange of Heading & Detail lines
ArrayLen(0)=54
ArrayPos(1)=73
ArrayLen(1)=6
ArrayPos(2)=113
ArrayLen(2)=66
ArrayPos(3)=246         
ArrayLen(3)=66
ArrayPos(4)=378
ArrayLen(4)=66
ArrayPos(5)=512
ArrayLen(5)=66                  ' Do not include percent of variance per specs


Private Sub ProcessLine

    SetUserVariablesBasedOnFilename (FirstName)
    FirstName = 1

    If FirstRec < 1 Then          
      If InStr(sRec,"GL Budget Report for" ) > 0 Then     ' Heading line 1
         DeletePage
         saveHeading=sRec
         fout.Writeline sRec
         sRec=Getline
         saveHeading2=sRec
         fout.Writeline sRec
         sRec=Getline
         fout.Writeline sRec
         sRec=Getline
         ' fout.Writeline sRec
         fout.Writeline Full_Heading                      ' Write additional line of heading, will be modified using macro
         fout.Writeline newHeading
         'fout.Writeblanklines (2)                        ' Use Macro to create proper headings
         sRec=Getline
         'fout.Writeline sRec                             ' DO NOT  Write dashes in heading per Brian S
         fout.WriteBlankLines (1)                 
         sRec=Getline         
         fout.Writeline Line3Head                         ' Write Section header for 'Balance Sheet'
         FirstRec=1
      End If  
   End If 
         
   
   If InStr(sRec," OPERATING INCOME") > 0    Then         ' We have reached the MTD-QTD part
      Line3Head="Section Income Statement"
      mFlag=1
      rType="MTD"
      lnct=4
      SpaceCnt=135
      fout.Writeline Line3Head
   End If
             
   outRec=sRec
	
   If IsNumeric (Mid(sRec,1,6))    Then     ' Start of a new account
      saveAcct=Mid(sRec,1,6)
      saveName=Mid(sRec,13,41)
      If InStr(sRec,rType) > 0     Then     ' First major heading, Join lines as for Branch
         outRec=JoinWithNLines(Lnct)        ' lnct contains 2 for Balance, 4 for Income Stmt
         outRec=Mid(outRec,1,53) & space(54) & Mid(outRec,54,Len (outRec) - 53)   ' Pad with spaces
         outRec=rearrange(outRec,6,ArrayPos,ArrayLen)
  
      End If
   End If
   
   If InStr(srec," Branch ") > 6 Then
      outRec=JoinWithNLines(Lnct)       ' lnct contains 2 for Balance, 4 for Income       
      outRec=SaveAcct & Space(6) & SaveName & " " & outRec 
'msgbox Mid(outRec,113,66)
      outRec=rearrange(outRec,6,ArrayPos,ArrayLen)         
   End If   

  

   fout.Writeline outRec               ' Write out all records, whether formatted or not

   
End Sub



