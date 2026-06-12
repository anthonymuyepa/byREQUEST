Dim fCount: fCount = 1
Dim combinedHeader: combinedHeader = ""
Dim paramsDict              ' Dictionary to store parsed parameter settings
Dim fileDicts(4)            ' Array to hold dictionaries for File1 to File4

Sub startDoc()

    ' Step 1: Load parameters from external configuration file
    Set paramsDict = ReadParametersFile("S:\CarolinaTrust\byREQUEST\Scripts\UniWyo_params.txt") 
	
    Dim folder, path, i, filekey, joinCols, keyword
    folder = paramsDict("FolderPath") ' Base folder where files are located

    ' Step 2: Attempt to load up to 4 files (there could be less) based on keywords and join columns
    For i = 1 To 4
        filekey = "File" & i & "Keyword"

        If paramsDict.Exists(filekey) Then
            keyword = Trim(paramsDict(filekey))

            If keyword <> "" Then
                path = FindFileMatchingKeyword(folder, keyword)

                If Trim(path) <> "" Then
                    If paramsDict.Exists("File" & i & "JoinColumns") Then
                        joinCols = paramsDict("File" & i & "JoinColumns")
						
                        ' Attempt to load the file into a dictionary
                        On Error Resume Next
                        Set fileDicts(i) = LoadFileToDict(path, joinCols)
                        If Err.Number <> 0 Then
                            MsgBox "Error loading file: " & path & vbCrLf & "Error: " & Err.Description
                            Err.Clear
                        End If
                        On Error GoTo 0
                    End If
                End If
            End If
        End If
    Next

    ' Step 3: Print all loaded dictionary contents
    PrintAllDictionaries fileDicts

End Sub


Function ReadParametersFile(filePath)
    Dim dict, fso, file, line, parts

    Set dict = CreateObject("Scripting.Dictionary")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(filePath) Then
        MsgBox "Parameters file not found: " & filePath
        Set ReadParametersFile = dict
        Exit Function
    End If

    Set file = fso.OpenTextFile(filePath, 1)
    Do Until file.AtEndOfStream
        line = Trim(file.ReadLine)
        If line <> "" And InStr(line, "=") > 0 Then
            parts = Split(line, "=")
            dict(Trim(parts(0))) = Trim(parts(1))
        End If
    Loop
    file.Close

    Set ReadParametersFile = dict
End Function


Function FindFileMatchingKeyword(folderPath, keyword)
    Dim fso, folder, file, match
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    match = "" ' default to empty in case nothing found

    For Each file In folder.Files
        If InStr(1, file.Name, keyword, 1) > 0 Then ' case-insensitive
            match = folderPath & "\" & file.Name
            Exit For
        End If
    Next

    If match = "" Then
        fout.WriteLine "No file found in folder '" & folderPath & "' matching keyword: " & keyword
    End If

    FindFileMatchingKeyword = match
End Function



Function LoadFileToDict(filePath, joinCols)
    Dim fso, file, line, header, row, key, i, dict
    Dim existing, headerCols, prefixedCols(), temp, joinColArray

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set dict = CreateObject("Scripting.Dictionary")

    If Not fso.FileExists(filePath) Then
        MsgBox "File not found: " & filePath
        Exit Function
    End If

    Set file = fso.OpenTextFile(filePath, 1)

    ' Read and process header
    If Not file.AtEndOfStream Then header = file.ReadLine
    headerCols = ParseCSVLine(header)

    ReDim prefixedCols(UBound(headerCols))
    For i = 0 To UBound(headerCols)
        prefixedCols(i) = "F" & fCount & "_" & Trim(headerCols(i))
    Next

    If combinedHeader <> "" Then
        combinedHeader = combinedHeader & "," & Join(prefixedCols, ",")
    Else
        combinedHeader = Join(prefixedCols, ",")
    End If
    fCount = fCount + 1

    ' Normalize joinCols to array
    If InStr(joinCols, ",") > 0 Then
        joinColArray = Split(joinCols, ",")
    Else
        ReDim temp(0)
        temp(0) = joinCols
        joinColArray = temp
    End If

    ' Read and process each line
    Do Until file.AtEndOfStream
        line = Trim(file.ReadLine)
        If Len(line) > 0 Then
            row = ParseCSVLine(line)

            ' Build join key
            key = ""
            For i = 0 To UBound(joinColArray)
                If CInt(joinColArray(i)) <= UBound(row) Then
                    key = key & Trim(row(CInt(joinColArray(i)))) & "|"
                Else
                    Exit Function
                End If
            Next
            If Len(key) > 0 Then key = Left(key, Len(key) - 1)

            ' Store row in dictionary
            If Not dict.Exists(key) Then
                dict.Add key, Array(row)
            Else
                existing = dict(key)
                If IsArray(existing) Then
                    dict(key) = AppendArray(existing, Array(row))
                Else
                    Exit Function
                End If
            End If
        End If
    Loop

    file.Close
    Set LoadFileToDict = dict
End Function


Function ParseCSVLine(line)
    Dim result(), field, inQuotes, c, i
    ReDim result(0)
    field = ""
    inQuotes = False

    For i = 1 To Len(line)
        c = Mid(line, i, 1)

        If c = """" Then
            If inQuotes And Mid(line, i + 1, 1) = """" Then
                field = field & """"
                i = i + 1
            Else
                inQuotes = Not inQuotes
            End If
        ElseIf c = "," And Not inQuotes Then
            result(UBound(result)) = Trim(field)
            ReDim Preserve result(UBound(result) + 1)
            field = ""
        Else
            field = field & c
        End If
    Next

    result(UBound(result)) = Trim(field)
    ParseCSVLine = result
End Function





Function AppendArray(arr1, arr2)
    Dim result(), i, totalLen

    If IsEmpty(arr1) Then
        ' If arr1 is empty, just return arr2
        ReDim result(UBound(arr2))
        For i = 0 To UBound(arr2)
            result(i) = arr2(i)
        Next
    Else
        ' Combine both arrays
        totalLen = UBound(arr1) + 1 + UBound(arr2) + 1
        ReDim result(totalLen - 1)

        For i = 0 To UBound(arr1)
            result(i) = arr1(i)
        Next

        For i = 0 To UBound(arr2)
            result(UBound(arr1) + 1 + i) = arr2(i)
        Next
    End If

    AppendArray = result
End Function





Function JoinTwoFileDicts(dictA, dictB)
    Dim joined, allKeys, fkey, i, j, rowsA, rowsB, rowA, rowB, resultRow
    Set joined = CreateObject("Scripting.Dictionary")
    Set allKeys = CreateObject("Scripting.Dictionary")

    ' Gather all unique keys
    For Each fkey In dictA.Keys
        allKeys(fkey) = 1
    Next
    For Each fkey In dictB.Keys
        allKeys(fkey) = 1
    Next

    ' Join rows for each key
    For Each fkey In allKeys.Keys
        If dictA.Exists(fkey) Then
            rowsA = dictA(fkey)
        Else
            rowsA = Array(Array("")) ' 1 row with 1 blank column
        End If

        If dictB.Exists(fkey) Then
            rowsB = dictB(fkey)
        Else
            rowsB = Array(Array(""))
        End If

        ' Cartesian join
        For i = 0 To UBound(rowsA)
            For j = 0 To UBound(rowsB)
                rowA = rowsA(i)
                rowB = rowsB(j)
				
                resultRow = JoinArrays(rowA, rowB)
                If Not joined.Exists(fkey) Then
                    joined.Add fkey, Array(resultRow)
                Else
                    joined(fkey) = AppendArray(joined(fkey), Array(resultRow))
                End If
            Next
        Next
    Next

    ' ' DEBUG
    ' For Each fkey In joined.Keys
        ' MsgBox "JOINED KEY: " & fkey & vbCrLf & "Rows: " & UBound(joined(fkey)) + 1
    ' Next

    Set JoinTwoFileDicts = joined
End Function



Function JoinArrays(arr1, arr2)
    Dim i, total, result
    total = -1
    If IsArray(arr1) Then total = total + UBound(arr1) + 1
    If IsArray(arr2) Then total = total + UBound(arr2) + 1
    ReDim result(total)
    For i = 0 To UBound(arr1)
        result(i) = arr1(i)
    Next
    For i = 0 To UBound(arr2)
        result(UBound(arr1) + 1 + i) = arr2(i)
    Next
    JoinArrays = result
End Function


' Computes maximum column widths from combinedData

Function ComputeMaxWidths(dict)
    Dim rowKey, rows, rawRow, row, i, j
    Dim maxWidths()
    Dim initialized
    initialized = False

    For Each rowKey In dict.Keys
        rows = dict(rowKey)
        For i = 0 To UBound(rows)
            rawRow = rows(i)

            ' FLATTEN if needed (if row is an array of arrays)
            If IsArray(rawRow(0)) Then
                row = FlattenRow(rawRow)
            Else
                row = rawRow
            End If

            ' DEBUG: Show row content
            'MsgBox "Row Key: " & rowKey & vbCrLf & "Row(" & i & "): " & Join(row, ",")

            If Not initialized Then
                ReDim maxWidths(UBound(row))
                For j = 0 To UBound(maxWidths)
                    maxWidths(j) = 0
                Next
                initialized = True
            End If

            ' Measure width of each column
            For j = 0 To UBound(row)
                Dim cellVal, cellLen
                If IsNull(row(j)) Then
                    cellVal = ""
                Else
                    cellVal = Trim(CStr(row(j)))
                End If

                cellLen = Len(cellVal)
                If cellLen > maxWidths(j) Then
                    maxWidths(j) = cellLen
                End If
            Next
        Next
    Next

    ' Apply padding and minimum width
    For j = 0 To UBound(maxWidths)
        maxWidths(j) = maxWidths(j) + PADDING
        If maxWidths(j) < MIN_WIDTH Then
            maxWidths(j) = MIN_WIDTH
        End If
    Next

    ' DEBUG: Final column widths
    'MsgBox "Final column widths (padded): " & Join(maxWidths, ", ")

    ComputeMaxWidths = maxWidths
End Function

' Flattens an array of arrays into a single array
Function FlattenRow(nested)
    Dim flat(), i, j, k
    k = 0

    For i = 0 To UBound(nested)
        k = k + UBound(nested(i)) + 1
    Next

    ReDim flat(k - 1)
    k = 0

    For i = 0 To UBound(nested)
        For j = 0 To UBound(nested(i))
            flat(k) = nested(i)(j)
            k = k + 1
        Next
    Next

    FlattenRow = flat
End Function



Function FormatLine(fields, colWidths)
    Dim result, i
    result = ""
    For i = 0 To UBound(fields)
        result = result & PadRight(Trim(fields(i)), colWidths(i) + PADDING)
    Next
    FormatLine = result
End Function


Function PadRight(value, width)
    PadRight = Left(value & Space(width), width)
End Function


Sub PrintAllDictionaries(fileDicts)
    Dim i, dict, key, row, j

    fout.WriteLine "========== DICTIONARY CONTENTS =========="

    For i = 1 To 4
        If Not IsEmpty(fileDicts(i)) Then
            Set dict = fileDicts(i)
            fout.WriteLine "----- File " & i & " -----"
            For Each key In dict.Keys
                fout.WriteLine "Key: " & key
                If IsArray(dict(key)) Then
                    For j = 0 To UBound(dict(key))
                        row = dict(key)(j)
                        If IsArray(row) Then
                            fout.WriteLine "  Row " & j & ": " & Join(row, ", ")
                        Else
                            fout.WriteLine "  Row " & j & ": " & row
                        End If
                    Next
                Else
                    fout.WriteLine "  Value: " & dict(key)
                End If
            Next
        Else
            fout.WriteLine "File " & i & " not loaded or empty."
        End If
    Next

    fout.WriteLine "========================================="
End Sub
