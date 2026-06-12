''''''''''''''''''''''''''''''''''''''''''''''''
' Script Name: DataExplorerInsertData.vbs   based on InsertMaster.vbs
' Renamed to be less wrong   Steve Ashcraft 12-12-2022   Previously name InsertIncentive which was a little confusing.
' Author: ZTS/SEA 
' Date: July 11, 2023
' Desc:
'	This is an adaption of the old populateCSV script that would take a CSV file through byREQUEST and append the records into a text driver driven database for BDC tool
'	The program now takes csv files and stores the information in a sql server .  Database name is set in PreDEScriptForxxxxx where xxxxx is name of DE project.
'	The table name is based on the spooler name. If there is not a table with the same name aas the spooler name, a table with be created.
'	More historical data goes here 
option explicit
dim Requiredfiles()						' list of files that we need and then we create triggerfile
dim FileWithListOfSpoolersProcessedForProject

Private Function replByPos(s, f, L, n)
   replByPos = Left(s, f-1) & n & Mid(s, f + L )
End Function


Private Sub StartDoc
   on error resume next
      'msgbox "START : Spoolfile.Name=" & Spoolfile.Name
   BeginCheckFileList
       'msgbox "back from begincheckfilelist"
' Grab the date from the Spoolfile variables to be added to the end of each line
   ' set datefield based on spoolfile name
   SetDateOfDataFromFilename()
      'msgbox "datefield=" & datefield
   fout.writeline "in start doc datefield=" & datefield
   ' Assume to skip the first line, comment out if not needed
   sRec = GetLine
       'msgbox "first line with headers="  & srec
   ' Grab first line of csv file for column headers
   Set Fields = RegSplit(SplitString, sRec)
   numFields = Fields.Count
	colCreate = "("
	cols = "("
	Dim j
	If UseHeader = 1 Then
		For i=0 to (Fields.Count - 1) Step 1
			If Len(Fields(i)) <=1 Then
				Fields(i) = "dummy" & i
			End If
			' Check if col is unique
			If i > 0 Then 
				For j=0 to Fields.Count - 1 Step 1
					If  j<>i and StrComp(Fields(j),Fields(i)) = 0 Then
						Fields(j) = Fields(i) & j
						    'msgbox "numCol=" & numcol   & " i=" & i &   " fields(i)=" &   fields(i) & "j=" & j & " fields(j)=" & fields(j)
					End If
				Next
			End If
			colCreate = colCreate & chr(34) & Trim(Fields(i)) & chr(34) & " NVARCHAR(50), "
			cols =           cols & chr(34) & Trim(Fields(i)) & chr(34) & ", "
		Next
		    'msgbox " use header colCreate = " & colCreate
	Else
		    'msgbox "not using header"
		currField = 0
		Do While currField < numFields
			colCreate = colCreate & "col" & currField & " NVARCHAR(50), "
			cols = cols & "col" & currField & ", "
			currField = currField + 1
		loop
	End If
	colCreate = colCreate & "DateCol NVARCHAR(50))"
	cols = cols & "DateCol )"
	    'msgbox "cols=" & cols
    fout.writeline cols
	fout.writeline colCreate
end sub

Private Sub ProcessLine()
	    'msgbox iln & ":" & srec
	' skip lines that need to be skipped
	If len(srec) < 5 then   
      ' skip this line
	else	
		Dim valueee
		on error resume next
		valueee = sRec
		Set Fields = RegSplit(SplitString, sRec)
		insertQuery = "INSERT INTO " & tablename & " "
		values = " VALUES ("
		Dim currField 
		currField = 0
		    'msgbox "num of fields in processline=" & fields.count
		For i=0 to (Fields.Count - 1) Step 1
			'Field = Trim(Replace(Field,Chr(34),""))
			Fields(i) = Replace(Fields(i),"'","''")
			If InStr(Fields(i),"$") = 1 or InStr(Fields(i),"$") = 2 or IsNumeric(Mid(Fields(i),2,1)) or IsNumeric(Left(Fields(i),1)) Then
				Fields(i) = Replace(Fields(i),"$","")
				Fields(i) = Replace(Fields(i),",","")
				Fields(i) = Replace(Fields(i),"(","-")
				Fields(i) = Replace(Fields(i),")","")
			End If
			values = values & "'" & Fields(i) & "', "
			currField = currField + 1
		Next
		
		values = values & "'" & DateField & "')"	
		    'msgbox "values=" & values
		    'msgbox "cols= " & cols
		insertQuery = insertQuery & cols & values
		fout.writeline iln & " " & insertquery
		tableQuery = "IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '" & tablename & "') BEGIN CREATE TABLE " & tablename & " " & colCreate & " END;"
		fout.writeline tablequery
		    'msgbox tablequery
		On Error Resume Next 
		err.clear
		Recordset.Open tableQuery,Connection
		If err.number <> 0 Then
			msgbox "after tablequery.connection err=" & err.number
			fout.WriteLine "Error on Line " & iLN & " with tablequery " & tablequery & "caused: " & err.description
			return
		End If
		err.clear
		 'msgbox "insertquery=" & insertquery
		Recordset.Open insertQuery,Connection
		If err.number <> 0 Then
			 'msgbox "after insertquery.connection err=" & err.number & Err.Description
			fout.WriteLine "Error on Line " & iLN & " with query " & inserquery & "caused: " & err.description
			return
		end If 
   end if 
End Sub

private sub CloseDoc
        'msgbox "Entered CloseDoc"
	EndCheckFileList	
	    'msgbox "CloseDoc - back from EndCheckFileList"
	Connection.Close
	Recordset.Close
	Set Connection = Nothing
	Set Recordset = Nothing
	Set FSO = nothing
end sub

'this sub finds the matches     
Private Function RegSplit(strMatchPattern, strPhrase) 
	'create variables 
	Dim objRegEx, Match, Matches, StrReturnStr,startString,endString,length,item 
	set RegSplit = CreateObject("System.Collections.ArrayList")
	'create instance of RegExp object 
	Set objRegEx = New RegExp 
	'Global makes RegExp match all occurrences in string 
	objRegEx.global = true
	'set the pattern 
	objRegEx.Pattern = strMatchPattern 
	'create the collection of matches 
	Set Matches = objRegEx.Execute(strPhrase) 
	'add all matches to the collection
	startString=1
	endString=-1
	length=0
	For Each Match in Matches 
		startString=endString+2
		endstring=match.firstindex
		length = endString-startString+1
		item  =  mid(strPhrase,startString,length)
		' strip quotation marks if string starts with quotation mark
		if mid(item,1,1) =  chr(34) then 		
			item = replace(item,chr(34),"")
		end if
		RegSplit.add( item)
	Next 
End Function 
Rem               +===========================================================+
Rem               |  Create Trigger File for Data Explorer if all files are present                                                          |
Rem               |      |
Rem               |    Steve Ashcraft  Aug. 2021                           |
Rem               |  Edited by |
Rem               +===========================================================+
Rem       This script should be the Scripts directory of byRequest
'--------------------------------------------------


Private Sub BeginCheckFileList	
	'For each final report there is a list of "things" that have to be processed so this Private Function needs to be updated for each new final report
	' that uses DataExplorerInsertData.vbS
	    msgbox "entered begincheckfilelist"
	redim requiredfiles (20)
	dim i
	for i = 1 to 19
		requiredfiles(i)=""
	next
	select case lcase(DataExplorerProjectName)
		case "loanincentives"
			msgbox "before array"
			RequiredFiles(1) =  "centralizedIncentive" 
			RequiredFiles(2) =  "debtprotection"
			
			msgbox "after array"
			        
		case "newloansbypapergrade"
			RequiredFiles(1) =  "eomloansbybranchcode"

	end select
		 msgbox "begincheckfilelist    projectname=" & dataexplorerprojectname
		 
	     
	FileWithListOfSpoolersProcessedForProject=  "SpoolersProcessedFor" & DataExplorerProjectName & ".txt"
	     msgbox "file with list of spoolers processed for project=:" & FileWithListOfSpoolersProcessedForProject
	TriggerFile =  "TriggerFileFor" & DataExplorerProjectName & ".txt"
	' Add the name of the current file to the FilesProcessed file 
	' so we can detect when the files we need are present 
	 msgbox " Triggerfile=" & triggerfile
	Call AddSpoolerToListOfSpoolersProcessed(Spooler)   
         msgbox "file with list of processed files=" &	FileWithListOfProcessedSpoolers & "  spooler="&  spooler
End Sub  

Private Sub EndCheckFileList
         msgbox "entered EndCheckFileList"
	if AllSpoolersProcessed = 1 Then	'   have all been processed
		     msgbox "endcheckfile list triggerfile=" & triggerfile
		call SaveStringToFile(TriggerFile ,"Ready to run Data Explorer")   ' This will change the time stamp on the trigger file 
		 msgbox "before clear " & FileWithListOfSpoolersProcessedForProject
		call Clearfile (FileWithListOfSpoolersProcessedForProject,"")                      ' once we have all of the files we will have to re process them all if we want to run dataexplorer again on this host
		    'msgbox " Cleared " & FileWithListOfProcessedFiles & " err.description=" & Err.Description
	end if
end Sub


Const FSOForReading = 1
Const FSOForWriting = 2

Private Function LoadStringFromFile(filename)                              ' Retrieve all of contents of 'filename' into a string
    on error resume next
	     msgbox "loadstringfromfile=" & filename &":"
    Dim f
	LoadStringFromFile = ""
	err.clear
	if err.number  = 0 then
		     msgbox "open text file:" & filename
		Set f = FSO.OpenTextFile(filename, FSOForReading, true)
		if err.number = 0 then
			LoadStringFromFile = f.ReadAll
		Else
			fout.writeline "Error getting string from " & filename & " err.description=" & err.description
		end if
	end if
    f.Close
	on error goto 0
End Function

Sub SaveStringToFile(filename, text)                                 ' Append 'text' to the current contents of 'filename'
    on error resume next											 ' so that we have a list of the files that have been processed
    Dim FSO, f
	err.clear														 ' for this iteration
    Set FSO = CreateObject("Scripting.FileSystemObject")
	if err.number = 0 then
        Set f = FSO.OpenTextFile(filename, FSOForWriting,true)
	    f.write text
		if err <> 0 then
			 msgbox  "SaveStringToFile cannot write to " & filename & " Err=" & err.description
		   fout.writeline  "SaveStringToFile cannot write to " & filename & " Err=" & err.description
		end if
	 else 
		 msgbox "SaveStringToFile cannot create file system object. Err=" & err.description
	    f.out.writeline "SaveStringToFile cannot create file system object. Err=" & err.description
	 end if
	 f.Close
	 on error goto 0
End Sub
Sub ClearFile(filename, text)                                 ' Append 'text' to the current contents of 'filename'
    on error resume next											 ' so that we have a list of the files that have been processed
    Dim FSO, f
	err.clear														 ' for this iteration
    Set FSO = CreateObject("Scripting.FileSystemObject")
	if err.number = 0 then
        Set f = FSO.OpenTextFile(filename, FSOForWriting,true)
	    f.write text
		if err <> 0 then
		   fout.writeline  "SaveStringToFile cannot write to " & filename & " Err=" & err.description
		end if
	 else 
	    f.out.writeline "SaveStringToFile cannot create file system object. Err=" & err.description
	 end if
	 f.Close
	 on error goto 0
End Sub


Private Function AddSpoolerToListOfSpoolersProcessed(tmpFileName)            ' Add the name of the current Spoolfile to the list of Spoolfiles that this 
	     msgbox "Entered AddSpoolerToListOfSpoolersProcessed tmpfilename=" & tmpfilename
	dim prevString
	prevString = LoadStringFromFile(FileWithListOfSpoolersProcessedForProject)
	     msgbox "addspoolertolistofspooersprocessed Name of FileWithListOfSpoolersProcessedForProject=" & FileWithListOfSpoolersProcessedForProject
	     msgbox "addspoolertolistofspooersprocessed list of previous spoolers processef" & tmpfilename
	SaveStringToFile FileWithListOfSpoolersProcessedForProject,LoadStringFromFile(FileWithListOfSpoolersProcessedForProject)&tmpFileName
	     msgbox "exiting AddSpoolerToListOfSpoolersProcessed"
End Function

Private Function AllSpoolersProcessed()                                   ' See if the list of files have all been processed  have all been processed                                               '   0-->not yet    1--> yes so we can do run Data Explorer
	dim strListOfFiles
	     msgbox "entered allspoolersProcessed FileWithListOfSpoolersProcessedForProject=" & FileWithListOfSpoolersProcessedForProject
	AllSpoolersProcessed=1
	strListOfFiles = ucase(LoadStringFromFile(FileWithListOfSpoolersProcessedForProject))
	dim i
	     msgbox "allspoolersprocessedstrListOfFiles= " & strListOfFiles  
	for i = 1 to 19
	        msgbox "i=" & i & " ucase(requiredFiles(i))=" & ucase(requiredFiles(i))
	   if len(requiredfiles(i)) < 1 Then
		   exit for
	    end if
	   if instr(strListOfFiles,ucase(requiredFiles(i))) = 0 Then
	       AllSpoolersProcessed=0
	   end if
	next
	     msgbox "Exit AllSpoolersProcessed=" & AllSpoolersProcessed
	 
End Function





public sub  SetDateOfDataFromFilename()
  ' Grab the date from the Spoolfile variables to be added to the end of each line
   Dim fullDate, fixedDate, month, day, year
   fullDate = Split(Spoolfile.Name,"___")
   fixedDate = Split(fullDate(1),".")
   month = Mid(fixedDate(0),1,2)
   day = Mid(fixedDate(0),3,2)
   year = Mid(fixedDate(0),5,4)
   datefield = year  & month & day
   fout.writeline "GetDateOfDataFromFilename in start doc datefield=" & datefield
      'msgbox "SetDateOfDataFromFilename=" & datefield
end sub