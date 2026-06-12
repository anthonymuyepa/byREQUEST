Option Explicit
' Last modification date 
' 	7/31/17 Winship
'	Added User variables
'   4-13-2017  Winship
'   Created for PIMA file GL Trial Balance under subprocess line

Dim IsBranchEntry ' Search for Branch entries using the Branch keyord followed by the 4-digit number
Set IsBranchEntry = New RegExp
IsBranchEntry.pattern = "\bBranch\b\s\d{4}"

Dim Header
Dim DoNotPrint

Dim FirstRec
FirstRec = 0

sub processline
    if FirstRec = 0 then
		SetUserVariablesBasedOnFilename (FirstRec)
		FirstRec = 1
	end if
	
	if instr(srec, "---------------") AND Header <> 1 then ' Mark the Header as read when done
		Header = 1
	end if
	
	if instr(srec, "GL Trial Balance for") AND Header = 1 then 	' If the start of a header, don not print lines
		DoNotPrint = 1
	end if
	
	
	if IsBranchEntry.Test(mid(srec,13,11)) = False AND DoNotPrint <> 1 then	' Print the line if it isnt a branch and not a header
		fout.writeline srec
	end if
	
	if instr(srec, "----------------") AND DoNotPrint = 1 then	' If the line has dashes, assume the header has ended and re-enable printing
		DoNotPrint = 0
	end if
	
	
	
end sub


