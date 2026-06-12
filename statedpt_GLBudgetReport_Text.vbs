Rem               +==========================================+
Rem               |                                          |
Rem               |    Symitar - GL Budget report            |    
Rem               |     Separate file into YTD only-         |
Rem               |       and MTD-QTD-YTD                    |
Rem               |     Written 6/2014   Hillary Software    |
Rem               |  This version adds a value to 3rd line   |
Rem               |    of heading so file can be burst       |
Rem               +==========================================+

'Edit - 08-15-2016
'	Changed the script so that it has a dynamic process to determine when at end of header.
'	This is prevent change in the number of lines of the header from messing up the file format

Option Explicit

Const OverwriteTrue = TRUE
Const ReplaceChr = " "
Dim Lnct,SpaceCnt,outRec,mFlag,FirstRec
Dim Line3Head, saveHeading, saveHeading2
Dim saveAcct,SaveName
Dim headRec
 
Line3Head="Section Balance Sheet     "
FirstRec=0
SpaceCnt=0
Lnct=2   

'Variable for number of lines in header
Dim HeaderNumLines, headerloop
HeaderNumLines = 3

Private Sub ProcessLine

            
   If InStr(sRec,"GL Budget Report for" ) > 0 Then     ' Heading line 1
      If FirstRec < 1 Then 
         saveHeading=sRec
         fout.Writeline sRec
         sRec=Getline
         saveHeading2=sRec
         fout.Writeline sRec
       '  fout.Writeline Line3Head
         sRec=Getline
		 'Keep printing lines until dashed line is found which indicates end of header
		 Do Until InStr(sRec,"---------------") > 0 
			 fout.WriteLine sRec
			 sRec=getLine
			 HeaderNumLines = HeaderNumLines + 1
		 Loop
         FirstRec=1
      Else
		headerloop = 0
         Do Until headerloop = HeaderNumLines
			sRec=GetLine
			headerloop = headerloop + 1
		 Loop' Skip headings on subsequent pages
      End If 
   End If 

   fout.Writeline sRec
   
End Sub



