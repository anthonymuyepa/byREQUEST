' Rewritten VBScript to perform a  full outer join across 4 files, aligned, with robust join column handling and chaining


Option Explicit

' Global variables
Dim paramsDict              ' Dictionary to store parsed parameter settings
Dim fout                    ' Output file object, provided globally by ByRequest
Dim fileDicts(4)            ' Array to hold dictionaries for File1 to File4
Dim combinedData            ' Dictionary to hold the final joined dataset
Dim combinedHeader 

combinedHeader =""
Const PADDING = 3            ' Global setting: extra padding for visual spacing
Const MIN_WIDTH = 3

Sub startDoc()

    ' Step 1: Load parameters from external configuration file
    Set paramsDict = ReadParametersFile("S:\CarolinaTrust\byREQUEST\Scripts\UniWyo_params.txt")

    Dim folder, path, i, filekey, joinCols, keyword
    folder = paramsDict("FolderPath") ' Base folder where files are located

    ' Step 2: Attempt to load up to 4 files based on keywords and join columns
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
                    Else
                        'MsgBox "JoinColumns missing for File" & i & " — skipping"
                    End If
                Else
                    'MsgBox "No file found for keyword: " & keyword
                End If
            Else
               ' MsgBox "Skipping File" & i & ": Empty keyword"
            End If
        End If
    Next

    ' Step 3: Initialize combinedData with File1's dictionary
    If Not fileDicts(1) Is Nothing Then
        Set combinedData = fileDicts(1)
    Else
        'MsgBox "File1 dictionary is missing — cannot proceed with join."
        Exit Sub
    End If

    ' Step 4: Join remaining files (File2 to File4) with combinedData
    For i = 2 To 4
        If paramsDict.Exists("File" & i & "Keyword") Then
            If Not fileDicts(i) Is Nothing Then
                Set combinedData = JoinTwoFileDicts(combinedData, fileDicts(i))
                'MsgBox "Dictionary File" & i & " finished"
            Else
                'MsgBox "File" & i & "Dict is Nothing"
            End If
        Else
            'MsgBox "File" & i & " not defined in parameters"
        End If
    Next

' You can override the default file combination sequence and number of files joined
' by commenting/uncommenting the lines below. The combination order matters and affects output alignment.

' Set combinedData = fileDicts(1)
' Set combinedData = JoinTwoFileDicts(combinedData, fileDicts(2))
' Set combinedData = JoinTwoFileDicts(combinedData, fileDicts(3))




    ' Step 5: Ensure we have data to write
    If combinedData.Count = 0 Then
        'MsgBox "No rows to output — combinedData is empty"
        Exit Sub
    End If

    ' Step 6: Write the final joined data to the output file
    Dim colWidths, keyOut, rowArr, j
    

    ' Compute column widths for formatting output
    colWidths = ComputeMaxWidths(combinedData)
	If combinedHeader <> "" Then
		Dim columns
		columns = Split(combinedHeader, ",")

		Dim trimmedColumns()
		ReDim trimmedColumns(UBound(columns))
	
		For j = 0 To UBound(columns)
			trimmedColumns(j) = Trim(columns(j))
		Next

		fout.WriteLine FormatLine(trimmedColumns, colWidths)
	End If

    ' Write each row of the joined data to the output file
    For Each keyOut In combinedData.Keys
        rowArr = combinedData(keyOut)
        For j = 0 To UBound(rowArr)
            fout.WriteLine FormatLine(rowArr(j), colWidths)
        Next
    Next

End Sub

Function ReadParametersFile(paramPath)
    Dim fso, file, line, parts, paramsDict, key, val
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.OpenTextFile(paramPath, 1)
    Set paramsDict = CreateObject("Scripting.Dictionary")

    Do Until file.AtEndOfStream
        line = Trim(file.ReadLine)

        If Len(line) > 0 And InStr(line, "=") > 0 Then
            parts = Split(line, "=")
            key = Trim(parts(0))

            If UBound(parts) >= 1 Then
                val = Trim(parts(1))

                ' Only add if value is not empty
                If val <> "" Then
                    If InStr(key, "JoinColumns") > 0 Then
                        paramsDict.Add key, Split(val, ",")
                    Else
                        paramsDict.Add key, val
                    End If
                Else
                    
                End If
            Else
                
            End If
        End If
    Loop

    file.Close

    ' ' DEBUG: Show all parameters
    ' Dim k, v, i
    ' For Each k In paramsDict.Keys
        ' v = paramsDict(k)
        ' If IsArray(v) Then
            ' MsgBox k & " = " & Join(v, ",")
        ' Else
            ' MsgBox k & " = " & v
        ' End If
    ' Next

    Set ReadParametersFile = paramsDict
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
    Dim existing

    'MsgBox "LoadFileToDict 1"
    Set fso = CreateObject("Scripting.FileSystemObject")
    'MsgBox "LoadFileToDict: " & filePath

    If Not fso.FileExists(filePath) Then
        MsgBox "File not found: " & filePath
        Exit Function
    End If

    Set file = fso.OpenTextFile(filePath, 1)
    Set dict = CreateObject("Scripting.Dictionary")

    ' Read header and ignore
    If Not file.AtEndOfStream Then header = file.ReadLine

	If combinedHeader <> "" Then
		combinedHeader = combinedHeader & ", " & header
	Else
		combinedHeader = header
	End If
	
	
    Do Until file.AtEndOfStream
        line = Trim(file.ReadLine)
        If Len(line) > 0 Then
            row = Split(line, ",")

            ' Build join key
            key = ""
            For i = 0 To UBound(joinCols)
                If CInt(joinCols(i)) <= UBound(row) Then
                    key = key & Trim(row(CInt(joinCols(i)))) & "|"
                Else
                    'MsgBox "Join column index out of bounds in file: " & filePath
                    Exit Function
                End If
            Next
            If Len(key) > 0 Then key = Left(key, Len(key) - 1)

            ' Store row in dictionary (handle duplicates)
            If Not dict.Exists(key) Then
                dict.Add key, Array(row)
            Else
                existing = dict(key)
                If IsArray(existing) Then
                    dict(key) = AppendArray(existing, Array(row))
                Else
                    ' MsgBox "Type mismatch at key: " & key & vbCrLf & "Expected array, got: " & TypeName(existing)
                    Exit Function
                End If
            End If
        End If
    Loop

    file.Close

    ' ' DEBUG
    ' For Each key In dict.Keys
        ' MsgBox "Key: " & key & vbCrLf & "RowCount: " & UBound(dict(key)) + 1
    ' Next

    Set LoadFileToDict = dict
End Function





Function AppendArray(arr, newItem)
    Dim i, newArr

    If IsEmpty(arr) Then
        ReDim newArr(0)
        newArr(0) = newItem(0)
    Else
        ReDim newArr(UBound(arr) + 1)
        For i = 0 To UBound(arr)
            newArr(i) = arr(i)
        Next
        newArr(UBound(arr) + 1) = newItem(0)
    End If

    AppendArray = newArr
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
