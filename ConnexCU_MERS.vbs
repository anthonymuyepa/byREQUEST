Option Explicit

Sub StartDoc
    Dim i
   
    For i = 1 To 28
        srec = getLine
        If i > 1 And i < 4 Then
            fout.WriteLine srec
        End If                      
    Next
	
    Fout.Writeline "         AccountNumber     Field Name                  MERS DATA                       DNA DATA"
	
End Sub

Sub ProcessLine
    '---------------------------------------------------------
    ' Skip report header / footer / column-heading lines
    '---------------------------------------------------------
    If InStr(srec, "Run Date:") > 0 _
	   or Len(srec) < 2 _
	   or Instr(srec, Chr(12)) > 0 _
       Or InStr(srec, "Post Date:") > 0 _
       Or InStr(srec, "Page:") > 0 _
       Or (InStr(srec, "               Account") > 0 _
           Or InStr(srec, "MERS DATA                       DNA DATA") > 0) Then 
        Exit Sub
    End If


     fout.WriteLine srec 


End Sub