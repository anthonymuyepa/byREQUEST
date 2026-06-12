
' ================================
' This Script performs a FULL OUTER JOIN on two CSV files 

' Date 05/13/2025     Anthonym@hillary.com 

'  How to Use:
' 1. Configure parameters in 'join_params.txt' .
' 2. This script reads two CSV files from a specified folder.
' 3. It selects the most recent files matching two keywords, joins them by specified column indexes,
'    and writes a full outer join output with aligned columns for further byREQUEST processing.
'
' 🔧 Sample join_params.txt:
' FolderPath=S:\Tmp
' File1Keyword=Loan and Address FM
' File2Keyword=Loan Next Due Date
' OutputFile=Combined_Latest.txt
' File1JoinColumn=0
' File2JoinColumn=0
' ================================

Option Explicit

' Global declarations
Dim fso, folder, file, configFile, configDict
Dim folderPath, keywords(3), joinCols(3), filePaths(3)
dim indi
Sub startDoc()
	indi = 1
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set configDict = CreateObject("Scripting.Dictionary")

    ' === Load config parameters ===
    Dim configFile, line, parts
    'Set configFile = fso.OpenTextFile("S:\CarolinaTrust\byREQUEST\Scripts\CalorinaTrust_params.txt", 1) 
	Set configFile = fso.OpenTextFile("S:\CarolinaTrust\byREQUEST\Scripts\UniWyo_params.txt", 1) 
	
	
	
    Do Until configFile.AtEndOfStream
        line = Trim(configFile.ReadLine)
        If line <> "" And InStr(line, "=") > 0 Then
            parts = Split(line, "=")
            configDict(Trim(parts(0))) = Trim(parts(1))
            'MsgBox "Loaded: " & Trim(parts(0)) & " = " & Trim(parts(1))
        End If
    Loop
    configFile.Close

    ' === Initialize keywords and join columns ===
    Dim keywords(3), joinCols(3), filePaths(3), i
    folderPath = configDict("FolderPath")

    For i = 0 To 3
        Dim keyK, keyJ
        keyK = "File" & (i + 1) & "Keyword"
        keyJ = "File" & (i + 1) & "JoinColumns"

        If configDict.Exists(keyK) And Trim(configDict(keyK)) <> "" Then
            keywords(i) = configDict(keyK)
        Else
            keywords(i) = ""
        End If

        If configDict.Exists(keyJ) And Trim(configDict(keyJ)) <> "" Then
            joinCols(i) = Split(configDict(keyJ), ",")
        Else
            joinCols(i) = Array()
        End If
    Next

    ' === Find latest matching files ===
    Dim latestDates(3)
    For i = 0 To 3
        latestDates(i) = #1/1/1900#
        filePaths(i) = ""
    Next

    Set folder = fso.GetFolder(folderPath)
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "csv" Then
            For i = 0 To 3
                If keywords(i) <> "" And InStr(file.Name, keywords(i)) > 0 Then
                    If file.DateLastModified > latestDates(i) Then
                        latestDates(i) = file.DateLastModified
                        filePaths(i) = file.Path
                    End If
                End If
            Next
        End If
    Next

    For i = 0 To 3
        If filePaths(i) <> "" Then 
			'MsgBox "Matched file for '" & keywords(i) & "':" & vbCrLf & filePaths(i)
		End if
    Next

    ' === Filter to non-blank valid files ===
    Dim validPaths(), validCols()
    ReDim validPaths(-1)
    ReDim validCols(-1)

    For i = 0 To 3
        If filePaths(i) <> "" Then
            Dim tf
            Set tf = fso.OpenTextFile(filePaths(i), 1)
            tf.SkipLine
            If Not tf.AtEndOfStream Then
                ReDim Preserve validPaths(UBound(validPaths) + 1)
                ReDim Preserve validCols(UBound(validCols) + 1)
                validPaths(UBound(validPaths)) = filePaths(i)
                validCols(UBound(validCols)) = joinCols(i)
               ' MsgBox "Using file: " & filePaths(i)
            Else
                MsgBox "Skipped blank file: " & filePaths(i)
            End If
            tf.Close
        End If
    Next

    ' === Exit early if not enough files ===
    Dim tempPath, outPath
    outPath = folderPath & "\Combined_Latest.txt"
    tempPath = folderPath & "\temp_join.txt"
  

    If UBound(validPaths) < 0 Then
        fout.WriteLine "No data to process."
        fout.Close
       ' MsgBox "No valid files found."
        Exit Sub
    ElseIf UBound(validPaths) = 0 Then
       ' MsgBox "Only one file found. Outputting directly."
        Call OutputSingleFile(validPaths(0), True)
     
        Exit Sub
    End If

    ' === Join first two files ===
    'MsgBox "Joining:" & vbCrLf & validPaths(0) & vbCrLf & "WITH" & vbCrLf & validPaths(1)
    Call FullOuterJoinToFile(validPaths(0), validCols(0), validPaths(1), validCols(1), tempPath)

    
	
		' Inspect full header of temp file
	Dim fTmp, headerRaw, headersArray, colLine, idx
	Set fTmp = fso.OpenTextFile(tempPath, 1)
	headerRaw = fTmp.ReadLine
	fTmp.Close

	headersArray = ParseCSVLine(headerRaw)
	colLine = "Columns in tempPath:" & vbCrLf
	For idx = 0 To UBound(headersArray)
		colLine = colLine & "[" & idx & "] " & headersArray(idx) & vbCrLf
	Next
	'MsgBox colLine
		
	
	
	
	' === Define correct left-side join cols for 3rd file (tempJoin col 3 is Account from File2) ===
    Dim currentJoinCols
    currentJoinCols = Array(3)  ' Adjust if structure changes
	
	
	' 🐞 Debug column index and name
	'MsgBox "CurrentJoinCols = " & Join(currentJoinCols, ",") & vbCrLf & GetColumnNamesByIndexes(tempPath, currentJoinCols)
	
	
	
    ' === Join remaining files (3rd and 4th if available) ===
    For i = 2 To UBound(validPaths)
        'MsgBox "Joining temp with: " & vbCrLf & validPaths(i) & vbCrLf & "Temp JoinCols: " & Join(currentJoinCols, ",") & vbCrLf & "Incoming JoinCols: " & Join(validCols(i), ",")

        Call FullOuterJoinToFile(tempPath, currentJoinCols, validPaths(i), validCols(i), tempPath)

       ' MsgBox "Column count in temp after join: " & GetColumnCountFromFile(tempPath)
    Next

    ' === Final copy of temp join result to output ===
    Dim tf2
	Dim rawLine, parsed
	Set tf2 = fso.OpenTextFile(tempPath, 1)
	Do Until tf2.AtEndOfStream
		rawLine = tf2.ReadLine
		parsed = ParseCSVLine(rawLine)
		fout.WriteLine FormatLine(parsed)
	Loop
	tf2.Close
	fout.Close

End Sub




Sub FullOuterJoinToFile(file1Path, joinCols1, file2Path, joinCols2, outFilePath)
    Dim dict1, dict2, processedKeys
    Dim file1, file2, foutLocal
    Dim headers1, headers2, line, fields, key
    Dim i, j, k
    Dim leftColCount, rightColCount

    Set dict1 = CreateObject("Scripting.Dictionary")
    Set dict2 = CreateObject("Scripting.Dictionary")
    Set processedKeys = CreateObject("Scripting.Dictionary")

    ' === Read File 1 ===
    Set file1 = fso.OpenTextFile(file1Path, 1)
    headers1 = ParseCSVLine(file1.ReadLine)
    leftColCount = UBound(headers1) + 1
    Do Until file1.AtEndOfStream
        line = file1.ReadLine
        If Trim(line) <> "" Then
            fields = ParseCSVLine(line)
            key = BuildJoinKey(fields, joinCols1)
            If Not dict1.Exists(key) Then Set dict1(key) = CreateObject("Scripting.Dictionary")
            dict1(key).Add dict1(key).Count, fields
        End If
    Loop
    file1.Close

    ' === Read File 2 ===
    Set file2 = fso.OpenTextFile(file2Path, 1)
    headers2 = ParseCSVLine(file2.ReadLine)
    rightColCount = UBound(headers2) + 1
    Do Until file2.AtEndOfStream
        line = file2.ReadLine
        If Trim(line) <> "" Then
            fields = ParseCSVLine(line)
            key = BuildJoinKey(fields, joinCols2)
            If Not dict2.Exists(key) Then Set dict2(key) = CreateObject("Scripting.Dictionary")
            dict2(key).Add dict2(key).Count, fields
        End If
    Loop
    file2.Close

    ' === Write Output Header ===
    Set foutLocal = fso.CreateTextFile(outFilePath, True)
    foutLocal.WriteLine Join(headers1, ",") & "," & Join(headers2, ",")

    ' === Matched and unmatched from File 1 ===
    For Each k In dict1.Keys
        If dict2.Exists(k) Then
            For i = 0 To dict1(k).Count - 1
                For j = 0 To dict2(k).Count - 1
                    foutLocal.WriteLine CleanJoin(dict1(k)(i)) & "," & CleanJoin(dict2(k)(j))
                Next
            Next
        Else
            For i = 0 To dict1(k).Count - 1
                foutLocal.WriteLine CleanJoin(dict1(k)(i)) & "," & PadBlankCsv(rightColCount)
            Next
        End If
        processedKeys.Add k, True
    Next

    ' === Unmatched from File 2 ===
    For Each k In dict2.Keys
        If Not processedKeys.Exists(k) Then
            For j = 0 To dict2(k).Count - 1
                foutLocal.WriteLine PadBlankCsv(leftColCount) & "," & CleanJoin(dict2(k)(j))
            Next
        End If
    Next

    foutLocal.Close
End Sub






Function BuildJoinKey(fields, colArray)
    Dim key, i, v
    key = ""
    For Each i In colArray
        v = Replace(Trim(fields(CInt(i))), " ", "")
        If IsNumeric(v) Then
            key = key & "|" & FormatKeyNumber(v)
        Else
            key = key & "|" & LCase(v)
        End If
    Next
    BuildJoinKey = key
End Function

Function FormatKeyNumber(value)
    FormatKeyNumber = Right("0000000000" & Replace(Trim(value), " ", ""), 10)
End Function

Sub OutputSingleFile(filePath, isFile1)
    Dim file, line, fields, headers, key
    Dim joinCol, sourceName

    Set file = fso.OpenTextFile(filePath, 1)
    headers = ParseCSVLine(file.ReadLine)
    fout.WriteLine FormatLine(headers)

    Do Until file.AtEndOfStream
        line = file.ReadLine
        If Trim(line) <> "" Then
            If isFile1 And InStr(Left(line, 12), "-") > 0 Then
                line = Replace(Left(line, 12), "-", " ") & Mid(line, 13)
            End If
            fields = ParseCSVLine(line)
            fout.WriteLine FormatLine(fields)
        End If
    Loop
    file.Close
End Sub

Function ParseCSVLine(line)
    Dim inQuotes, i, ch, current, result
    inQuotes = False
    current = ""
    result = Array()

    For i = 1 To Len(line)
        ch = Mid(line, i, 1)
        Select Case ch
            Case """"
                inQuotes = Not inQuotes
            Case ","
                If inQuotes Then
                    current = current & ch
                Else
                    ReDim Preserve result(UBound(result) + 1)
                    result(UBound(result)) = current
                    current = ""
                End If
            Case Else
                current = current & ch
        End Select
    Next

    ReDim Preserve result(UBound(result) + 1)
    result(UBound(result)) = current
    ParseCSVLine = result
End Function

Function FormatLine(fields)
    Dim result, i
    result = ""
    For i = 0 To UBound(fields)
        If i < 6 Then
            result = result & PadRight(Trim(fields(i)), 18)
        Else
            result = result & PadRight(Trim(fields(i)), 36)
        End If
    Next
    FormatLine = result
End Function

Function PadRight(value, width)
    PadRight = Left(value & Space(width), width)
End Function


Function GetColumnCountFromFile(filePath)
    Dim f, headerLine
    Set f = fso.OpenTextFile(filePath, 1)
    headerLine = f.ReadLine
    f.Close
    GetColumnCountFromFile = UBound(ParseCSVLine(headerLine)) + 1
End Function


Function GetColumnNamesByIndexes(filePath, indexes)
    Dim f, headerLine, headers, i, out
    Set f = fso.OpenTextFile(filePath, 1)
    headerLine = f.ReadLine
    f.Close

    headers = ParseCSVLine(headerLine)
    out = ""

    For i = 0 To UBound(indexes)
        If CLng(indexes(i)) <= UBound(headers) Then
            out = out & "[" & indexes(i) & "] " & headers(indexes(i)) & vbCrLf
        Else
            out = out & "[" & indexes(i) & "] (Invalid Index)" & vbCrLf
        End If
    Next

    GetColumnNamesByIndexes = out
End Function


Function PadBlankColumns(count)
    Dim i, result
    result = ""
    For i = 1 To count
        If i < 7 Then
            result = result & PadRight("", 18)
        Else
            result = result & PadRight("", 36)
        End If
    Next
    PadBlankColumns = result
End Function

Function CleanJoin(arr)
    Dim i, safe
    safe = ""
    For i = 0 To UBound(arr)
        If i > 0 Then safe = safe & ","
        safe = safe & Replace(arr(i), ",", "") ' remove inner commas
    Next
    CleanJoin = safe
End Function

Function PadBlankCsv(count)
    Dim i, blanks
    blanks = ""
    For i = 1 To count
        If i > 1 Then blanks = blanks & ","
        blanks = blanks & ""
    Next
    PadBlankCsv = blanks
End Function
