' This VBScript performs a FULL OUTER JOIN on two CSV files — "Loan and Address FM" and "Loan Next Due Date"
' It selects the most recently published file for each type, combines them row-by-row based on the first column (Account/Member Number).

' Date 05/13/2025     Anthonym@hillary.com 


Option Explicit

Dim fso, folder, file
Dim folderPath, file1Path, file2Path

Sub startDoc
    Set fso = CreateObject("Scripting.FileSystemObject")
    folderPath = "S:\Tmp"
    Set folder = fso.GetFolder(folderPath)

    file1Path = ""
    file2Path = ""

    ' 🔍 Step 1: Locate latest files regardless of date
    Dim latest1Date, latest2Date
    latest1Date = #1/1/1900#
    latest2Date = #1/1/1900#

    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "csv" Then
            If InStr(file.Name, "Loan and Address FM") > 0 Then
                If file.DateLastModified > latest1Date Then
                    latest1Date = file.DateLastModified
                    file1Path = file.Path
                End If
            ElseIf InStr(file.Name, "Loan Next Due Date") > 0 Then
                If file.DateLastModified > latest2Date Then
                    latest2Date = file.DateLastModified
                    file2Path = file.Path
                End If
            End If
        End If
    Next

    ' 🔁 Step 2: Process if both files exist
    If file1Path <> "" And file2Path <> "" Then
        Call FullOuterJoinCSV(file1Path, file2Path, folderPath & "\Combined_Latest.txt")
    Else
        If file1Path = "" Then Fout.WriteLine "Missing: Loan and Address FM"
        If file2Path = "" Then Fout.WriteLine "Missing: Loan Next Due Date"
        Exit Sub
    End If

End Sub



' 🧠 FULL OUTER JOIN PROCEDURE with column alignment
Sub FullOuterJoinCSV(file1Path, file2Path, outputPath)
    Dim file1, file2, headers1, headers2
    Dim dict1, dict2, line, fields, key
    Dim processedKeys, k, emptyRow1, emptyRow2, matchCount, unmatchedCount

    Set dict1 = CreateObject("Scripting.Dictionary")
    Set dict2 = CreateObject("Scripting.Dictionary")
    Set processedKeys = CreateObject("Scripting.Dictionary")

    Set file1 = fso.OpenTextFile(file1Path, 1)
	line = file1.ReadLine
	line = Replace(line, Chr(34), "")
    headers1 = Split(line, ",")

    Do Until file1.AtEndOfStream
        line = file1.ReadLine
        If Trim(line) <> "" Then
            line = Replace(line, Chr(34), "")
            fields = Split(line, ",")
			
            key = LCase(Replace(Trim(fields(0)), " ", ""))
            If Not dict1.Exists(key) Then
                dict1.Add key, fields
            End If
        End If
    Loop
    file1.Close

    Set file2 = fso.OpenTextFile(file2Path, 1)
    line = file2.ReadLine
	line = Replace(line, Chr(34), "")
    headers2 = Split(line, ",")
	
    Do Until file2.AtEndOfStream
        line = file2.ReadLine
		
        If Trim(line) <> "" Then
			line = Replace(line, Chr(34), "")
			fields = Split(line, ",")

			' 🔧 Trim each field in the array
			Dim i
			For i = 0 To UBound(fields)
				fields(i) = Trim(fields(i))
			Next

			key = LCase(Replace(fields(0), " ", ""))
			If Not dict2.Exists(key) Then
				dict2.Add key, fields
			End If
		End If
    Loop
    file2.Close


    fout.WriteLine FormatLine(headers1) & FormatLine(headers2)

    matchCount = 0
    unmatchedCount = 0

    For Each k In dict1.Keys
        If dict2.Exists(k) Then
            fout.WriteLine FormatLine(dict1(k)) & FormatLine(dict2(k))
            matchCount = matchCount + 1
        Else
            fout.WriteLine FormatLine(dict1(k)) & Space(Len(FormatLine(headers2)))
        End If
        processedKeys.Add k, True
    Next

    For Each k In dict2.Keys
        If Not processedKeys.Exists(k) Then
            fout.WriteLine Space(Len(FormatLine(headers1))) & FormatLine(dict2(k))
            unmatchedCount = unmatchedCount + 1
        End If
    Next

    fout.Close


End Sub

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