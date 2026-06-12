Dim FirstRec
FirstRec = 0
Private Sub ProcessLine
	SetUserVariablesBasedOnFilename (FirstRec)
	FirstRec = 1
	
	Dim Month
	Dim Variable4
	Month = Array("Jan","Feb","Mar","Apr","May","June","Jul","Aug","Sep","Oct","Nov","Dec")
	Variable4 = Cint(spoolfile.User4)
	spoolfile.User8 = Month(Variable4-1)
	
	If instr(sRec,"==") > 0 then
	  sRec = getline
	End If
	
	fout.WriteLine sRec
	
	

End Sub