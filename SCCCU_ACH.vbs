Option Explicit

Dim gPendingLine   ' holds the "first line" waiting for its continuation
Dim header, EPSA, PCTL, LCTL, SETTLDTE, ACH_Batch_ID, Post_Time,Post_Date,Post_Seq, File_Post_Seq, ACH_Tran_Code, Report, companyName
header = "Report Post_Date Post_Time Settlement_Date Company_Name                  ACH_Batch_ID PCTL   LCTL   Post_Seq File_Post_Seq TransCode Description                  ST TYPE     EFF DATE         AMOUNT   CHECK#  SDC#      MEMBER# S/L NAME                                 FEE   RR   TC  RDFI IDENT  DFI ACCOUNT#      INDIVIDUAL ID#         INDIVIDUAL NAME        DD TRACE#          MESSAGE"
Dim gTranMap  ' Dictionary for transaction codes 
Dim y

	Dim INI_PATH
	INI_PATH = "S:\CarolinaTrust\byREQUEST\Templates\ACH_Trans.ini"


Sub StartDoc()
    gPendingLine = ""
    y = 0
    Dim ReportDateTime, i 
    PCTL = ""
    SETTLDTE = ""
    ACH_Batch_ID = ""
    Report = ""
    ReportDateTime = ""
	companyName = ""
	

	
    For i = 1 to 14
        srec = getline()

        ' Extract Report name and date-time when PCTL# is found
        If instr(srec, "PCTL#:") > 0 Then
            ' Get first word (report name) before first space
            Dim firstSpacePos
            firstSpacePos = instr(srec, " ")
            If firstSpacePos > 0 Then
                Report = Trim(mid(srec, 1, firstSpacePos - 1))
				' Simple safe padding for your code
				If Len(Report & "") > 0 Then
					If Len(Report) >= 6 Then
						Report = Left(Report, 6)
					Else
						Report = Report & Space(6 - Len(Report))
					End If
				Else
					Report = Space(6) ' Empty string becomes 6 spaces
				End If

            End If
            
            ' Extract date-time (everything after report name until PCTL#)
            Dim reportEndPos
            reportEndPos = Len(Report) + 1 ' Position after report name
            
            Dim pctlPos
            pctlPos = instr(srec, "PCTL#:")
            
            If pctlPos > reportEndPos Then
                ' Extract everything between report name and PCTL#
                ReportDateTime = Trim(mid(srec, reportEndPos, pctlPos - reportEndPos))
            End If
        End If
		
		' In your processing loop, add this:
		If instr(srec, "CO NAME:") > 0 Then
			Dim companyNameStart, companyNameEnd
			
			' Find the position after "CO NAME:"
			companyNameStart = instr(srec, "CO NAME:") + 8  ' 8 characters in "CO NAME:"
			
			' Find the position of "ID:" after the company name
			companyNameEnd = instr(companyNameStart, srec, " ID:")
			
			If companyNameEnd > 0 Then
				' Extract everything between "CO NAME:" and " ID:"
				companyName = Trim(mid(srec, companyNameStart, companyNameEnd - companyNameStart))
			End If
			
			
			If Len(companyName & "") > 0 Then
				If Len(companyName) >= 22 Then
					' Truncate if longer than 22 characters
					companyName = Left(companyName, 22)
				Else
					' Pad with spaces if shorter than 22 characters
					companyName = companyName & Space(22 - Len(companyName))
				End If
			Else
				' If empty, use 22 spaces
				companyName = Space(22)
			End If
		End If
		
        
        ' Extract PCTL#
        If instr(srec, "PCTL#:") > 0 Then
            Dim pctlValuePos
            pctlValuePos = instr(srec, "PCTL#:") + 6 ' Position after "PCTL#:"
            PCTL = Trim(mid(srec, pctlValuePos, 10))
			' Pad with spaces if shorter than 7 characters
			If Len(PCTL) >= 7 Then
				PCTL = Left(PCTL, 7)
			Else
				
				PCTL = PCTL & Space(7 - Len(PCTL))
			End If			
			
			
			
        End If
        
        ' Extract SETTL DTE
        If instr(srec, "SETTL DTE:") > 0 Then
            Dim settlPos
            settlPos = instr(srec, "SETTL DTE:") + 10
            SETTLDTE = Trim(mid(srec, settlPos, 10))
        End If
        
        ' Extract ACH BTCH#
        If instr(srec, "ACH BTCH#:") > 0 Then
            Dim achPos
            achPos = instr(srec, "ACH BTCH#:") + 10
            ACH_Batch_ID = Trim(mid(srec, achPos, 7))
			
			' Pad with spaces if shorter than 12 characters
			If Len(ACH_Batch_ID) >= 13 Then
				ACH_Batch_ID = Left(ACH_Batch_ID, 13)
			Else
				
				ACH_Batch_ID = ACH_Batch_ID & Space(13 - Len(ACH_Batch_ID))
			End If
			
        End If
		
		' Extract LCTL#
        If instr(srec, "*** LCTL#:") > 0 Then
            Dim lctlPos
            lctlPos = instr(srec, "*** LCTL#:") + 10
            LCTL= Trim(mid(srec, lctlPos, 7))
			
			' Pad with spaces if shorter than 7 characters
			If Len(LCTL) >= 7 Then
				LCTL = Left(LCTL, 7)
			Else
				
				LCTL = LCTL & Space(7 - Len(LCTL))
			End If
			
        End If
		
    Next
    
    ' Parse the date:
	

	Dim dateTime, timeParts

	dateTime = Split(ReportDateTime, " ")  ' ["EPSA", "12-15-25", "02:20:11"]
	    ' Write the header with Report in first position
    fout.WriteLine header	
 
	If UBound(dateTime) > 0 Then
		Dim d
		d = Split(dateTime(0), "-")   ' MM-DD-YY

		Post_Date = DateSerial(2000 + CInt(d(2)) + 1, CInt(d(0)), CInt(d(1)))
		Post_Time = dateTime(1)

		
		
		' Parse time directly (HH:MM:SS)
		timeParts = Split(Post_Time, ":")
		If UBound(timeParts) >= 0 Then
			Post_Seq = timeParts(0)  ' Get hour directly from time string
 
		Else
			Post_Seq = "00"
		End If
		
		' Determine File_Post_Seq based on hour
		Select Case Post_Seq
			Case "01", "02", "03", "04", "05"
				File_Post_Seq = "D"
			Case "06", "07"
				File_Post_Seq = "O"
			Case "08", "09", "10", "11", "12"
				File_Post_Seq = "A"
			Case "13"
				File_Post_Seq = "B"
			Case "14"
				File_Post_Seq = "C"
			Case Else
				File_Post_Seq = "U"
		End Select
	Else
		File_Post_Seq = "U"
	End If
	
	' Read INI and build transaction code map


	Dim gCfg
	Set gCfg = ReadIniSafe(INI_PATH)
	If gCfg Is Nothing Then
		Fout.WriteLine "ERROR: INI not found or unreadable: " & INI_PATH
		Exit Sub
	End If
' MsgBox "here"
	Call InitializeTransactionMap()


 
End Sub

Sub ProcessLine()
  Dim line 
  
  line = srec
  
  
  
  ' If we already have a pending first-line, decide whether this line is its continuation
  If gPendingLine <> "" Then
  
	
  
  
    If IsContinuationLine(line) Then
      ' Merge: keep first line as-is   
	    Dim TransCode
		TransCode = Trim(Mid(line, 1, 4)) 
		 ' Get enrichment from INI mapping
		Dim enrichType, enrichDesc
		Call GetTranEnrichment(TransCode, enrichType, enrichDesc)
		
	
		If Len(enrichDesc) >= 28 Then
			' Truncate if longer than 28 characters
		    enrichDesc = Left(enrichDesc, 28)
		Else
			' Pad with spaces if shorter than 28 characters
			enrichDesc = enrichDesc & Space(28 - Len(enrichDesc))
		End If
		
		If Len(enrichType) >= 9 Then
            ' Truncate if longer than 9 characters
            enrichType = Left(enrichType, 9)
        Else
            ' Pad with spaces if shorter than 9 characters
            enrichType = enrichType & Space(9 - Len(enrichType))
        End If
 
		
		fout.WriteLine Report &" " & Post_Date &"  " & Post_Time & "  " & SETTLDTE &"      " &  companyName & "        " & ACH_Batch_ID & PCTL &  LCTL   & Post_Seq & "       " & File_Post_Seq & "             " & enrichType & " " & enrichDesc & " " & gPendingLine & " " & line
	  
		gPendingLine = ""
      Exit Sub
    End If
  End If

  ' If this is a first-line that should wait for a continuation, buffer it
  If IsFirstLineToMerge(line) Then
    gPendingLine = line

    Exit Sub
  End If

End Sub


' --- Helpers ---

Private Function IsFirstLineToMerge(ByVal s)
    s = Trim(s)
    
    If Len(s) < 10 Then  ' Minimum length for a transaction line
        IsFirstLineToMerge = False
        Exit Function
    End If
    
    ' Check if starts with P or NP pattern
    Dim first3
    first3 = Left(s, 3)
    
    ' Check: starts with P/NP AND contains a slash (date)
    If (InStr(first3, "P ") > 0 Or InStr(first3, "NP ") > 0) And InStr(s, "/") > 0 Then
        IsFirstLineToMerge = True
    Else
        IsFirstLineToMerge = False
    End If
End Function



Private Function IsContinuationLine(ByVal s)
  If Len(s) < 6 Then
    IsContinuationLine = False
    Exit Function
  End If

  IsContinuationLine = (Mid(s, 1, 2) = "  " _
                        And IsNumeric(Mid(s, 3, 2)) _
                        And Mid(s, 5, 2) = "  ")
End Function


Private Function IsAlpha(ByVal ch)
  Dim c
  If Len(ch) <> 1 Then
    IsAlpha = False
    Exit Function
  End If
  c = UCase(ch)
  IsAlpha = (c >= "A" And c <= "Z")
End Function




 

' Add this function at the top of your script
Function LogLine(msg)  
    'fout.WriteLine msg
End Function


Sub InitializeTransactionMap()
    ' Read INI file and build transaction code dictionary
    Dim gCfg
    Set gCfg = ReadIniSafe(INI_PATH)
    
    If gCfg Is Nothing Then
        LogLine "ERROR: INI not found or unreadable: " & INI_PATH
        LogLine "Transaction code enrichment will not be available."
        Set gTranMap = CreateObject("Scripting.Dictionary")
        Exit Sub
    End If
    
    ' Build the transaction code map
    Set gTranMap = BuildTranCodeMap(gCfg)
    
    ' MsgBox "Loaded " & gTranMap.Count & " transaction codes from INI"
    
    ' Debug: list loaded codes
    Dim codeKey
    For Each codeKey In gTranMap.Keys
        LogLine "  Code " & codeKey & ": " & gTranMap(codeKey)
    Next
End Sub

' INI Reading Function (if you don't have ReadIniSafe)
Function ReadIniSafe(iniFile)
    On Error Resume Next
    Dim fso, ts, line, section, key, value, dict
    Set dict = CreateObject("Scripting.Dictionary")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(iniFile) Then
        LogLine "ERROR: INI file does not exist: " & iniFile
        Set ReadIniSafe = Nothing
        Exit Function
    End If
    
    Set ts = fso.OpenTextFile(iniFile, 1) ' 1 = ForReading
    section = ""
    
    Do While Not ts.AtEndOfStream
        line = Trim(ts.ReadLine)
        
        ' Skip comments and empty lines
        If Len(line) = 0 Or Left(line, 1) = ";" Then
            ' Skip
        ' Check for section
        ElseIf Left(line, 1) = "[" And Right(line, 1) = "]" Then
            section = Mid(line, 2, Len(line) - 2)
        ' Process key=value pairs
        ElseIf InStr(line, "=") > 0 Then
            key = Trim(Left(line, InStr(line, "=") - 1))
            value = Trim(Mid(line, InStr(line, "=") + 1))
            
            ' Store as "section.key"
            If Len(section) > 0 Then
                dict(section & "." & key) = value
            Else
                dict(key) = value
            End If
        End If
    Loop
    
    ts.Close
    Set ReadIniSafe = dict
    On Error GoTo 0
End Function

' Transaction Code Map Builder
Function BuildTranCodeMap(cfg)
    Dim d, k
    Set d = CreateObject("Scripting.Dictionary")
    
    If cfg Is Nothing Then
        ' MsgBox "ERROR: cfg is Nothing in BuildTranCodeMap"
        Set BuildTranCodeMap = d
        Exit Function
    End If
    
    ' MsgBox "DEBUG BuildTranCodeMap: cfg has " & cfg.Count & " keys"
    
    ' DEBUG: Show ALL keys to see their format
    ' MsgBox "=== ALL KEYS IN CFG ==="
    Dim count
    count = 0
    For Each k In cfg.Keys
        ' MsgBox "Key " & count & ": '" & k & "' = '" & cfg(k) & "'"
        count = count + 1
        
        ' Check if it matches what we expect
        If LCase(Left(k, 10)) = "TranCodes." Then
           ' MsgBox"  ^ MATCHES TranCodes.!"
        End If
    Next
   
    
    ' Look for TranCodes section
    Dim foundCount
    foundCount = 0
    
    For Each k In cfg.Keys
        ' Check for TranCodes section entries
		
        If LCase(Left(k, 10)) = "trancodes." Then
            Dim code
            code = Mid(k, 11)  ' Get code after "TranCodes."
            ' MsgBox "DEBUG: Adding code '" & code & "' = '" & cfg(k) & "'"
            d(Trim(code)) = Trim(cfg(k))
            foundCount = foundCount + 1
        End If
    Next
    
    ' MsgBox "DEBUG: Found " & foundCount & " transaction codes"
    
    ' If nothing found, try alternative section names
    If foundCount = 0 Then
        ' MsgBox "DEBUG: Trying alternative section names..."
        
        Dim altSections
        altSections = Array("TranCode.", "TransCodes.", "TransCode.", "TransactionCodes.", "Transactions.", "Codes.")
        
        Dim altSection
        For Each altSection In altSections
            For Each k In cfg.Keys
                If LCase(Left(k, Len(altSection))) = LCase(altSection) Then
                    code = Mid(k, Len(altSection) + 1)
                    WScript.Echo "DEBUG: Found in '" & altSection & "': '" & code & "' = '" & cfg(k) & "'"
                    d(Trim(code)) = Trim(cfg(k))
                    foundCount = foundCount + 1
                End If
            Next
        Next
    End If
    
    Set BuildTranCodeMap = d
End Function

' Transaction Enrichment Function
Sub GetTranEnrichment(ByVal code, ByRef outType, ByRef outDesc)
    outType = ""
    outDesc = ""
    
    If gTranMap Is Nothing Then Exit Sub
    If Len(Trim(code)) = 0 Then Exit Sub
    
    Dim codeKey
    codeKey = Trim(code)
    
    ' Try exact match
    If gTranMap.Exists(codeKey) Then
   
    ElseIf IsNumeric(codeKey) Then
        ' Try without leading zeros
        codeKey = CStr(CLng(codeKey))
        If Not gTranMap.Exists(codeKey) Then
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    ' Parse the value
    Dim parts
    parts = Split(gTranMap(codeKey), "|")
    
    If UBound(parts) >= 0 Then
        outType = Trim(parts(0))
		
		
        If UBound(parts) >= 1 Then
            outDesc = Trim(parts(1))
        End If
		
    End If
End Sub