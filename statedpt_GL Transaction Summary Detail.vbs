Option Explicit
' Last modification date 
'   3-22-2017  Winship
'   Created for PIMA GL Transaction Summary Detail




Dim HeaderCount ' Used to print only the first header

Dim CategoryType

sub processline

	if first = 0 then
		SetUserVariablesBasedOnFilename (First)
		first =1
	end if
	Dim IsData
    Set IsData = New RegExp
	IsData.Pattern = "\d{6}\-\d{4}\.\d{4}"
	
	if HeaderCount <> 1 Then
		if instr(srec,"PIMA FCU") then
			srec = left(srec,InStrRev(srec,"Page    ")-1)
			fout.writeline srec
		elseif instr(srec,"GL Account       Name                Reference") then
			fout.writeline srec & space(4) & "GL Category"
		elseif instr(srec,"--------------------------------------") then
			HeaderCount = 1
			srec = srec & "------------------------------"
			fout.writeline srec
		end if
	end if
	
	if instr(srec,"GL Category:") then
		CategoryType = trim(replace(srec,"GL Category:",""))

	elseif IsData.Test(left(srec,17)) = "True" then
		Set IsData = New RegExp
		IsData.Pattern = "\d{6}\-\d{4}\.\d{4}"
		fout.writeline srec & Space(4) & CategoryType
	
	end if
	
	
	
	
	
	
	
	
end sub



