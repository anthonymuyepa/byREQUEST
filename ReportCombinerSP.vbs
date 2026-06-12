' ==========================================================================================
'
' Script: Multi-File Full Outer Join Processor
' Description:
' This VBScript reads up to 4 CSV files and performs a full outer join using a common
' join key (defined per file via parameter configuration). It aligns and formats the output
' into a tabular report with properly spaced headers.
'
' Author:         Anthonym@hillary.com
' Company:        Hillary Software
' Date:           June 7, 2025
'
' Input Expectations:
'   • All input files are standard comma-separated CSV files with headers in the first row.
'   • Join columns are defined per file using 0-based indices in the parameters file.
'
' Parameters File Format (example):
'   FolderPath=S:\Reports
'   File1Keyword=Transactions
'   File1JoinColumns=0
'   File2Keyword=Owners
'   File2JoinColumns=0
'   File3Keyword=Statuses
'   File3JoinColumns=0
'
' Join Type:
'   • Full outer join across files, preserving all rows
'   • Supports Cartesian product of matches where keys have multiple rows
'
' Output:
'   • Aligned output with dynamically generated headers
'   • Supports 3 output modes: single-row per key, flattened with duplicates, full Cartesian
'=====================================================================================

Dim fCount: fCount = 1
Dim combinedHeader: combinedHeader = ""
Dim paramsDict              ' Dictionary to store parsed parameter settings
Dim fileDicts(4)            ' Array to hold dictionaries for File1 to File4
Const PADDING = 2           ' Global setting: extra padding for visual spacing
Const MIN_WIDTH = 3

Sub startDoc()
    ' Step 1: Load parameters from external configuration file
    Set paramsDict = ReadParametersFile("S:\CarolinaTrust\byREQUEST\Scripts\Uniwyo_params.txt") 

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
                            fout.Writeline "Error loading file: " & path & vbCrLf & "Error: " & Err.Description
                            Err.Clear
                        End If
                        On Error GoTo 0
                    End If
                End If
            End If
        End If
    Next

    ' Optional Debug: Print all loaded dictionary contents
    ' PrintAllDictionaries fileDicts

    ' Step 3: Merge dictionaries across all loaded files
    Dim mergedDict
    Set mergedDict = MergeDictionaries(fileDicts)
	'PrintMergedDictionary (mergedDict)
    ' ' Step 4: Flatten and write final output based on selected mode
    ' If flattenMode = "CARTESIAN" Then
        ' FlattenMergedDictCartesian mergedDict
    ' ElseIf flattenMode = "PRESERVEDUPLICATES" Then
        ' FlattenMergedDictPreserveDuplicates mergedDict
    ' Else
         'FlattenMergedDict mergedDict ' Default: only first row per file
		 FlattenMergedDict mergedDict

    ' End If
End Sub



Sub PrintMergedDictionary(mergedDict)
    Dim key, rows, i, j, row, joinedRow

    For Each key In mergedDict.Keys
        fout.WriteLine "=== Key: " & key & " ==="
        rows = mergedDict(key)
        joinedRow = Array() ' initialize as empty array

        For i = 0 To 3
            'MsgBox "rows(" & i & ") is type: " & TypeName(rows(i))
            If IsArray(rows(i)) Then
                For j = 0 To UBound(rows(i))
                    row = rows(i)(j)
                    fout.WriteLine "  File" & (i + 1) & " Row " & (j + 1) & ": " & Join(row, " | ")
                    joinedRow = AppendArray(joinedRow, row)
                Next
            Else
                fout.WriteLine "  File" & (i + 1) & ": No data"
            End If
        Next

        ' Print the fully joined row
        If IsArray(joinedRow) Then
            fout.WriteLine "  Joined Row: " & Join(joinedRow, " | ")
        Else
            fout.WriteLine "  Joined Row: (empty)"
        End If

        fout.WriteLine ""
    Next
End Sub

Sub FlattenMergedDict(mergedDict)
    Dim key, i, rowByFile(4), row, flatRow, fileIndex, fileRows, j
    Dim flatFields, rawx
    Dim allFlatRows(), rowCount
    rowCount = 0
    ReDim allFlatRows(10000) ' Expand later if needed

    ' First row is the header
    flatFields = Split(combinedHeader, ",")
    allFlatRows(rowCount) = flatFields
    rowCount = rowCount + 1

    ' Loop through each key in the merged dictionary
    For Each key In mergedDict.Keys
        Dim rows, joinedRow
        rows = mergedDict(key)
        joinedRow = ""

        For i = 0 To 3
            If IsArray(rows(i)) Then
                ' Use first row from the source
                row = rows(i)(0)
                joinedRow = joinedRow & Join(row, ",") & ","
            Else
                ' Pad with the right number of commas
                joinedRow = joinedRow & String(CountHeaderColumns(i + 1), ",")
            End If
        Next

        ' Remove trailing comma
        If Right(joinedRow, 1) = "," Then
            joinedRow = Left(joinedRow, Len(joinedRow) - 1)
        End If

        flatFields = Split(joinedRow, ",")
        If rowCount > UBound(allFlatRows) Then ReDim Preserve allFlatRows(rowCount + 1000)
        allFlatRows(rowCount) = flatFields
        rowCount = rowCount + 1
    Next

    ' Trim final array
    ReDim Preserve allFlatRows(rowCount - 1)

    ' Align and print
    Dim colWidths
    colWidths = ComputeMaxColumnWidths(allFlatRows)

    For i = 0 To UBound(allFlatRows)
        fout.WriteLine FormatLine(allFlatRows(i), colWidths)
    Next
End Sub


Function AppendArraysSafe(arr1, arr2)
    Dim newArr(), i, n1, n2
    If Not IsArray(arr1) Then
        ReDim arr1(-1)
    End If
    If Not IsArray(arr2) Then
        ReDim arr2(-1)
    End If

    n1 = UBound(arr1)
    n2 = UBound(arr2)
    ReDim newArr(n1 + n2 + 1)

    For i = 0 To n1
        newArr(i) = arr1(i)
    Next
    For i = 0 To n2
        newArr(n1 + 1 + i) = arr2(i)
    Next

    AppendArraysSafe = newArr
End Function

Function AppendEmptyCols(arr, count)
    Dim newArr(), i, oldLen
    If Not IsArray(arr) Then
        ReDim arr(-1)
    End If

    oldLen = UBound(arr)
    ReDim newArr(oldLen + count + 1)

    For i = 0 To oldLen
        newArr(i) = arr(i)
    Next
    For i = 0 To count - 1
        newArr(oldLen + 1 + i) = ""
    Next

    AppendEmptyCols = newArr
End Function














Function CountHeaderColumns(fileIndex)
    Dim parts, i, count, prefix
    prefix = "F" & fileIndex & "_"
    parts = Split(combinedHeader, ",")
    count = 0
    For i = 0 To UBound(parts)
        If Left(parts(i), Len(prefix)) = prefix Then
            count = count + 1
        End If
    Next
    CountHeaderColumns = count
End Function













Function MergeDictionaries(fileDicts)
    Dim mergedDict, i, key, rowSet, row, entry, j
    Set mergedDict = CreateObject("Scripting.Dictionary")

    For i = 1 To 4
        If Not IsEmpty(fileDicts(i)) Then
            For Each key In fileDicts(i).Keys
                rowSet = fileDicts(i)(key)

                ' Initialize entry if new key
                If Not mergedDict.Exists(key) Then
                    entry = Array(Empty, Empty, Empty, Empty) ' index 0 = File1, 1 = File2, ...
                    mergedDict.Add key, entry
                End If

                entry = mergedDict(key)

                ' Append rows from this file to their corresponding slot
                If IsArray(rowSet) Then
                    If IsEmpty(entry(i - 1)) Then
                        entry(i - 1) = rowSet
                    Else
                        entry(i - 1) = AppendArray(entry(i - 1), rowSet)
                    End If
                End If

                mergedDict(key) = entry
            Next
        End If
    Next

    Set MergeDictionaries = mergedDict
End Function



Function ReadParametersFile(filePath)
    Dim dict, fso, file, line, parts

    Set dict = CreateObject("Scripting.Dictionary")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(filePath) Then
        fout.Writeline "Parameters file not found: " & filePath
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
        fout.Writeline "File not found: " & filePath
        Exit Function
    End If

    Set file = fso.OpenTextFile(filePath, 1)

    ' Read and process header
    If Not file.AtEndOfStream Then header = file.ReadLine
    headerCols = ParseCSVLine(header)

    ReDim prefixedCols(UBound(headerCols))
    For i = 0 To UBound(headerCols)
		prefixedCols(i) = EscapeCSVField("F" & fCount & "_" & Trim(headerCols(i)))
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

Function EscapeCSVField(val)
    If InStr(val, ",") > 0 Or InStr(val, """") > 0 Or InStr(val, vbCrLf) > 0 Then
        val = Replace(val, """", """""")
        EscapeCSVField = """" & val & """"
    Else
        EscapeCSVField = val
    End If
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
    totalLen = -1

    If IsArray(arr1) Then totalLen = totalLen + UBound(arr1) + 1
    If IsArray(arr2) Then totalLen = totalLen + UBound(arr2) + 1

    If totalLen < 0 Then
        ReDim result(-1)
    Else
        ReDim result(totalLen)
        Dim pos : pos = 0

        If IsArray(arr1) Then
            For i = 0 To UBound(arr1)
                result(pos) = arr1(i)
                pos = pos + 1
            Next
        End If

        If IsArray(arr2) Then
            For i = 0 To UBound(arr2)
                result(pos) = arr2(i)
                pos = pos + 1
            Next
        End If
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
        ' fout.Writeline "JOINED KEY: " & fkey & vbCrLf & "Rows: " & UBound(joined(fkey)) + 1
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

Function ComputeMaxColumnWidths(allRows)
    Dim colCount, colWidths(), i, j
    colCount = UBound(allRows(0))
    ReDim colWidths(colCount)

    For j = 0 To colCount
        colWidths(j) = Len(allRows(0)(j))
    Next

    For i = 1 To UBound(allRows)
        For j = 0 To colCount
            If Len(allRows(i)(j)) > colWidths(j) Then
                colWidths(j) = Len(allRows(i)(j))
            End If
        Next
    Next

    ComputeMaxColumnWidths = colWidths
End Function




Function ComputeColumnWidths(header)
    Dim parts, colWidths(), i
    parts = Split(header, ",")
    ReDim colWidths(UBound(parts))
    For i = 0 To UBound(parts)
        colWidths(i) = Len(Trim(parts(i)))
    Next
    ComputeColumnWidths = colWidths
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
    Dim result, i, safeText, width
    result = ""
    For i = 0 To UBound(fields)
        ' Handle nulls
        If IsNull(fields(i)) Then
            safeText = "Malawi"
        Else
            safeText = Trim(fields(i))
        End If

        ' Default width if missing
        If i > UBound(colWidths) Then
            width = Len(safeText)
        Else
            width = colWidths(i)
        End If

        result = result & PadRight(safeText, width + 2)
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


Sub FlattenMergedDictCartesian(mergedDict)
    Dim key, i, j, rowByFile(4), row, combo, allCombos
    Dim fields(), colOffset, k, colWidths
    Dim allFlatRows(), rowCount
    Dim headerFields

    rowCount = 0
    ReDim allFlatRows(10000)

    ' Add the header first
    headerFields = Split(combinedHeader, ",")
    allFlatRows(rowCount) = headerFields
    rowCount = rowCount + 1

    ' Loop each merged key
    For Each key In mergedDict.Keys
        ' Load row arrays for each file
        For i = 1 To 4
            If IsArray(mergedDict(key)(i - 1)) Then
                rowByFile(i) = mergedDict(key)(i - 1)
            Else
                ReDim row(CountHeaderColumns(i) - 1)
                For j = 0 To UBound(row)
                    row(j) = ""
                Next
			Dim tempRowSet
			ReDim tempRowSet(0)
			tempRowSet(0) = row
			rowByFile(i) = tempRowSet
            End If
        Next

        ' Get all combinations
        allCombos = CartesianProduct4(rowByFile(1), rowByFile(2), rowByFile(3), rowByFile(4))

        ' Flatten each combo into a row of strings
        For Each combo In allCombos
            ReDim fields(UBound(headerFields))
            colOffset = 0

            For i = 0 To 3
                row = combo(i)
                For k = 0 To UBound(row)
                    fields(colOffset + k) = row(k)
                Next
                colOffset = colOffset + CountHeaderColumns(i + 1)
            Next

            If rowCount > UBound(allFlatRows) Then ReDim Preserve allFlatRows(rowCount + 1000)
            allFlatRows(rowCount) = fields
            rowCount = rowCount + 1
        Next
    Next

    ' Trim the collected rows
    ReDim Preserve allFlatRows(rowCount - 1)

    ' Compute column widths
    colWidths = ComputeMaxColumnWidths(allFlatRows)

    ' Output aligned rows
    For i = 0 To rowCount - 1
        fout.WriteLine FormatLine(allFlatRows(i), colWidths)
    Next
End Sub



Function CartesianProduct4(rowByFile)
    Dim arr1, arr2, arr3, arr4
    Dim i, j, k, l
    Dim result(), count, combo(3)

    ' If any array is missing, replace with one empty array of N empty strings
    If Not IsArray(rowByFile(1)) Or UBound(rowByFile(1)) < 0 Then arr1 = Array(Array("", "", "")) Else arr1 = rowByFile(1)
    If Not IsArray(rowByFile(2)) Or UBound(rowByFile(2)) < 0 Then arr2 = Array(Array("", "", "")) Else arr2 = rowByFile(2)
    If Not IsArray(rowByFile(3)) Or UBound(rowByFile(3)) < 0 Then arr3 = Array(Array("", "", "")) Else arr3 = rowByFile(3)
    If Not IsArray(rowByFile(4)) Or UBound(rowByFile(4)) < 0 Then arr4 = Array(Array("", "", "")) Else arr4 = rowByFile(4)

    count = 0
    ReDim result(10000)

    For i = 0 To UBound(arr1)
        For j = 0 To UBound(arr2)
            For k = 0 To UBound(arr3)
                For l = 0 To UBound(arr4)
                    combo(0) = arr1(i)
                    combo(1) = arr2(j)
                    combo(2) = arr3(k)
                    combo(3) = arr4(l)
                    If count > UBound(result) Then ReDim Preserve result(count + 1000)
                    result(count) = combo
                    count = count + 1
                Next
            Next
        Next
    Next

    ReDim Preserve result(count - 1)
    CartesianProduct = result
End Function



Sub FlattenMergedDictPreserveDuplicates(mergedDict)
    Dim key, fileIndex, fileRows, rowByFile(4), i, j, flatRow
    Dim allFlatRows(), rowCount, combo
    rowCount = 0

    ' Step 1: collect all flattened rows (preserving all rows from all files)
    For Each key In mergedDict.Keys
        ' Initialize rowByFile array
        For i = 1 To 4
            rowByFile(i) = Array()
        Next

        ' Load the rows per file
        For fileIndex = 1 To 4
            If mergedDict(key)(fileIndex - 1) <> Empty Then
                fileRows = mergedDict(key)(fileIndex - 1)
                If IsArray(fileRows) Then
                    rowByFile(fileIndex) = fileRows
                End If
            End If
        Next

        ' Cartesian product across file rows
        Dim combos
        combos = CartesianProduct4(rowByFile(1), rowByFile(2), rowByFile(3), rowByFile(4))


        ' Store all combinations
        For Each combo In combos
            ReDim Preserve allFlatRows(rowCount)
            allFlatRows(rowCount) = FlattenCombo(combo)
            rowCount = rowCount + 1
        Next
    Next

    ' Step 2: Compute column widths
    Dim colWidths
    colWidths = ComputeMaxColumnWidths(allFlatRows)

    ' Step 3: Print header
    fout.WriteLine FormatLine(Split(combinedHeader, ","), colWidths)

    ' Step 4: Print all flattened lines
    For i = 0 To UBound(allFlatRows)
        fout.WriteLine FormatLine(allFlatRows(i), colWidths)
    Next
End Sub





Function FlattenCombo(combo)
    Dim flatRow(), totalCols, i, j, row
    totalCols = 0
    For i = 1 To 4
        totalCols = totalCols + CountHeaderColumns(i)
    Next
    ReDim flatRow(totalCols - 1)

    Dim colStart
    colStart = 0
    For i = 1 To 4
        If IsArray(combo(i - 1)) Then
            row = combo(i - 1)
            For j = 0 To CountHeaderColumns(i) - 1
                If j <= UBound(row) Then
                    flatRow(colStart + j) = row(j)
                Else
                    flatRow(colStart + j) = ""
                End If
            Next
        Else
            For j = 0 To CountHeaderColumns(i) - 1
                flatRow(colStart + j) = ""
            Next
        End If
        colStart = colStart + CountHeaderColumns(i)
    Next

    FlattenCombo = flatRow
End Function

