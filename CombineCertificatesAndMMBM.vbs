
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



Dim fso, folder, file, configFile, configDict
Dim folderPath, file1Keyword, file2Keyword, outputFile, file1Path, file2Path

Sub startDoc
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set configDict = CreateObject("Scripting.Dictionary")

    ' 🔧 Load parameters from config file
    Set configFile = fso.OpenTextFile("S:\CarolinaTrust\byREQUEST\Scripts\CombineCertificatesAndMMBM_params.txt", 1)
    Do Until configFile.AtEndOfStream
        Dim line, parts
        line = Trim(configFile.ReadLine)
        If line <> "" And InStr(line, "=") > 0 Then
            parts = Split(line, "=")
            configDict(Trim(parts(0))) = Trim(parts(1))
        End If
    Loop
    configFile.Close

    folderPath     = configDict("FolderPath")
    file1Keyword   = configDict("File1Keyword")
    file2Keyword   = configDict("File2Keyword")


    Set folder = fso.GetFolder(folderPath)

    file1Path = ""
    file2Path = ""

    Dim latest1Date, latest2Date
    latest1Date = #1/1/1900#
    latest2Date = #1/1/1900#

    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "csv" Then
            If InStr(file.Name, file1Keyword) > 0 Then
                If file.DateLastModified > latest1Date Then
                    latest1Date = file.DateLastModified
                    file1Path = file.Path
                End If
            ElseIf InStr(file.Name, file2Keyword) > 0 Then
                If file.DateLastModified > latest2Date Then
                    latest2Date = file.DateLastModified
                    file2Path = file.Path
                End If
            End If
        End If
    Next

    If file1Path <> "" And file2Path <> "" Then
		Call FullOuterJoinCSV(file1Path, file2Path)
	ElseIf file1Path <> "" Then
		fout.WriteLine "Only " & file1Keyword & " found. Outputting single file..."
		Call OutputSingleFile(file1Path, True)  ' True = isFile1
	ElseIf file2Path <> "" Then
		fout.WriteLine "Only " & file2Keyword & " found. Outputting single file..."
		Call OutputSingleFile(file2Path, False) ' False = isFile1
	Else
		fout.WriteLine "❌ No matching input files found. Nothing to process."
		Exit Sub
	End If
End Sub

Sub FullOuterJoinCSV(file1Path, file2Path)
    Dim file1, file2, headers1, headers2
    Dim dict1, dict2, line, fields, key
    Dim processedKeys, k, i, j
    Dim joinCol1, joinCol2

    Set dict1 = CreateObject("Scripting.Dictionary")
    Set dict2 = CreateObject("Scripting.Dictionary")
    Set processedKeys = CreateObject("Scripting.Dictionary")

    joinCol1 = CInt(configDict("File1JoinColumn"))
    joinCol2 = CInt(configDict("File2JoinColumn"))

    ' Read file 1 (multiple lines per key)
    Set file1 = fso.OpenTextFile(file1Path, 1)
    headers1 = ParseCSVLine(file1.ReadLine)
    Do Until file1.AtEndOfStream
        line = file1.ReadLine
        If Trim(line) <> "" Then
            fields = ParseCSVLine(line)
            key = LCase(Replace(Trim(fields(joinCol1)), " ", ""))
            If Not dict1.Exists(key) Then
                Set dict1(key) = CreateObject("Scripting.Dictionary")
            End If
            dict1(key).Add dict1(key).Count, fields
        End If
    Loop
    file1.Close

    ' Read file 2 (multiple lines per key)
    Set file2 = fso.OpenTextFile(file2Path, 1)
    headers2 = ParseCSVLine(file2.ReadLine)
    Do Until file2.AtEndOfStream
        line = file2.ReadLine
        If Trim(line) <> "" Then
            fields = ParseCSVLine(line)
            key = LCase(Replace(Trim(fields(joinCol2)), " ", ""))
            If Not dict2.Exists(key) Then
                Set dict2(key) = CreateObject("Scripting.Dictionary")
            End If
            dict2(key).Add dict2(key).Count, fields
        End If
    Loop
    file2.Close

    fout.WriteLine FormatLine(headers1) & FormatLine(headers2)

    ' Write matches and unmatched from dict1
    For Each k In dict1.Keys
        If dict2.Exists(k) Then
            For i = 0 To dict1(k).Count - 1
                For j = 0 To dict2(k).Count - 1
                    fout.WriteLine FormatLine(dict1(k)(i)) & FormatLine(dict2(k)(j))
                Next
            Next
        Else
            For i = 0 To dict1(k).Count - 1
                fout.WriteLine FormatLine(dict1(k)(i)) & Space(Len(FormatLine(headers2)))
            Next
        End If
        processedKeys.Add k, True
    Next

    ' Write unmatched from dict2
    For Each k In dict2.Keys
        If Not processedKeys.Exists(k) Then
            For j = 0 To dict2(k).Count - 1
                fout.WriteLine Space(Len(FormatLine(headers1))) & FormatLine(dict2(k)(j))
            Next
        End If
    Next
End Sub


Sub OutputSingleFile(filePath, isFile1)
    Dim file, line, fields, headers, key
    Dim joinCol, sourceName

    Set file = fso.OpenTextFile(filePath, 1)
    headers = ParseCSVLine(file.ReadLine)
    fout.WriteLine FormatLine(headers)

    If isFile1 Then
        joinCol = CInt(configDict("File1JoinColumn"))
        sourceName = configDict("File1Keyword")
    Else
        joinCol = CInt(configDict("File2JoinColumn"))
        sourceName = configDict("File2Keyword")
    End If

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
