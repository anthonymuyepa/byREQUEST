Option Explicit

Dim gFSO
Dim gShell
Dim gCfg

Dim gIniPath
Dim gSourceRoot
Dim gDestination
Dim gPatterns
Dim gDaysToGoBack

Dim gCopiedCount
Dim gSkippedCount
Dim gMatchedCount
Dim gFolderCount
Dim gErrorCount

Sub StartDoc()

    Set gFSO = CreateObject("Scripting.FileSystemObject")
    Set gShell = CreateObject("WScript.Shell")

    gIniPath = "S:\ConnexCU\Daily_Copy_DNA_Files.ini"

    gCopiedCount = 0
    gSkippedCount = 0
    gMatchedCount = 0
    gFolderCount = 0
    gErrorCount = 0

    fout.WriteLine "DNA FILE COPY PROCESS"
    fout.WriteLine String(110, "-")

    Set gCfg = LoadIni(gIniPath)

    gSourceRoot = Trim(GetCfg(gCfg, "Paths", "SourceRoot", ""))
    gDestination = Trim(GetCfg(gCfg, "Paths", "Destination", ""))
    gPatterns = Trim(GetCfg(gCfg, "Files", "Patterns", ""))
    gDaysToGoBack = Trim(GetCfg(gCfg, "Options", "DaysToGoBack", "0"))

    If gSourceRoot = "" Then
        fout.WriteLine "ERROR: Missing [Paths] SourceRoot in INI."
        Exit Sub
    End If

    If gDestination = "" Then
        fout.WriteLine "ERROR: Missing [Paths] Destination in INI."
        Exit Sub
    End If

    If gPatterns = "" Then
        fout.WriteLine "ERROR: Missing [Files] Patterns in INI."
        Exit Sub
    End If

    fout.WriteLine "INI File              : " & gIniPath
    fout.WriteLine "Source Root           : " & gSourceRoot
    fout.WriteLine "Destination           : " & gDestination
    fout.WriteLine "Patterns              : " & gPatterns
    fout.WriteLine "DaysToGoBack : " & gDaysToGoBack
    fout.WriteLine ""

    If Not gFSO.FileExists(gIniPath) Then
        fout.WriteLine "ERROR: INI file not found."
        Exit Sub
    End If

    If Not gFSO.FolderExists(gSourceRoot) Then
        fout.WriteLine "ERROR: Source root folder not found: " & gSourceRoot
        Exit Sub
    End If

    If Not gFSO.FolderExists(gDestination) Then
        On Error Resume Next
        gFSO.CreateFolder gDestination
        If Err.Number <> 0 Then
            fout.WriteLine "ERROR: Could not create destination folder: " & gDestination
            fout.WriteLine "DETAIL: " & Err.Description
            Err.Clear
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0
        fout.WriteLine "Created destination folder: " & gDestination
        fout.WriteLine ""
    End If

    fout.WriteLine "File Status"
    fout.WriteLine String(110, "-")

    If CStr(gDaysToGoBack) = "1" Then
        ProcessAllDateFolders gSourceRoot, gDestination, SplitPatterns(gPatterns)
    Else
        ProcessLatestDateFolder gSourceRoot, gDestination, SplitPatterns(gPatterns)
    End If

End Sub

Sub ProcessLine()
End Sub

Sub CloseDoc()
    fout.WriteLine ""
    fout.WriteLine String(110, "-")
    fout.WriteLine "Folders Processed : " & gFolderCount
    fout.WriteLine "Matched Files     : " & gMatchedCount
    fout.WriteLine "Copied Files      : " & gCopiedCount
    fout.WriteLine "Skipped Files     : " & gSkippedCount
    fout.WriteLine "Errors            : " & gErrorCount
    fout.WriteLine "Completed."
End Sub

Sub ProcessLatestDateFolder(sourceRoot, destPath, filePatterns)
    Dim latestFolder

    latestFolder = GetNewestDateFolder(sourceRoot)

    If latestFolder = "" Then
        fout.WriteLine "ERROR: No valid yyyymmdd folders found under: " & sourceRoot
        Exit Sub
    End If

    gFolderCount = gFolderCount + 1
    fout.WriteLine "Processing latest folder: " & latestFolder
    CopyMatchingFiles latestFolder, destPath, filePatterns
End Sub

Sub ProcessAllDateFolders(sourceRoot, destPath, filePatterns)
    Dim folders, i

    folders = GetSortedDateFolders(sourceRoot)

    If IsEmpty(folders) Then
        fout.WriteLine "ERROR: No valid yyyymmdd folders found under: " & sourceRoot
        Exit Sub
    End If

    For i = LBound(folders) To UBound(folders)
        gFolderCount = gFolderCount + 1
        fout.WriteLine "Processing folder: " & folders(i)
        CopyMatchingFiles folders(i), destPath, filePatterns
    Next
End Sub

Function GetNewestDateFolder(parentPath)
    Dim parentFolder, subFolder
    Dim bestName, folderName

    bestName = ""
    Set parentFolder = gFSO.GetFolder(parentPath)

    For Each subFolder In parentFolder.SubFolders
        folderName = Trim(subFolder.Name)
        If IsYYYYMMDD(folderName) Then
            If bestName = "" Then
                bestName = folderName
            ElseIf folderName > bestName Then
                bestName = folderName
            End If
        End If
    Next

    If bestName = "" Then
        GetNewestDateFolder = ""
    Else
        GetNewestDateFolder = gFSO.BuildPath(parentPath, bestName)
    End If
End Function

Function GetSortedDateFolders(parentPath)
    Dim parentFolder, subFolder
    Dim temp(), count, i, j, hold
    Dim folderName

    count = 0
    Set parentFolder = gFSO.GetFolder(parentPath)

    For Each subFolder In parentFolder.SubFolders
        folderName = Trim(subFolder.Name)
        If IsYYYYMMDD(folderName) Then
            ReDim Preserve temp(count)
            temp(count) = folderName
            count = count + 1
        End If
    Next

    If count = 0 Then Exit Function

    For i = 0 To UBound(temp) - 1
        For j = i + 1 To UBound(temp)
            If temp(i) > temp(j) Then
                hold = temp(i)
                temp(i) = temp(j)
                temp(j) = hold
            End If
        Next
    Next

    For i = 0 To UBound(temp)
        temp(i) = gFSO.BuildPath(parentPath, temp(i))
    Next

    GetSortedDateFolders = temp
End Function

Function IsYYYYMMDD(folderName)
    Dim yyyy, mm, dd, dt

    IsYYYYMMDD = False

    If Len(folderName) <> 8 Then Exit Function
    If Not IsNumeric(folderName) Then Exit Function

    yyyy = CInt(Left(folderName, 4))
    mm = CInt(Mid(folderName, 5, 2))
    dd = CInt(Right(folderName, 2))

    If yyyy < 2000 Or yyyy > 2100 Then Exit Function
    If mm < 1 Or mm > 12 Then Exit Function
    If dd < 1 Or dd > 31 Then Exit Function

    On Error Resume Next
    dt = DateSerial(yyyy, mm, dd)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    If Year(dt) = yyyy And Month(dt) = mm And Day(dt) = dd Then
        IsYYYYMMDD = True
    End If
End Function

Sub CopyMatchingFiles(sourcePath, destPath, filePatterns)
    Dim sourceFolder, file
    Dim matched, i, destFilePath, patternText
    Dim localMatchCount

    localMatchCount = 0

    If Not gFSO.FolderExists(sourcePath) Then
        fout.WriteLine "ERROR: Source folder does not exist: " & sourcePath
        gErrorCount = gErrorCount + 1
        Exit Sub
    End If

    If Not gFSO.FolderExists(destPath) Then
        gFSO.CreateFolder destPath
    End If

    Set sourceFolder = gFSO.GetFolder(sourcePath)

    fout.WriteLine "Scanning folder: " & sourcePath

    For Each file In sourceFolder.Files
        matched = False
        fout.WriteLine "FOUND   : " & file.Name

        For i = LBound(filePatterns) To UBound(filePatterns)
            patternText = Trim(filePatterns(i))
            patternText = Replace(patternText, "*", "")
            patternText = Replace(patternText, """", "")

            If patternText <> "" Then
                If InStr(1, UCase(file.Name), UCase(patternText), vbTextCompare) > 0 Then
                    matched = True
                    Exit For
                End If
            End If
        Next

        If matched Then
            localMatchCount = localMatchCount + 1
            gMatchedCount = gMatchedCount + 1
            destFilePath = gFSO.BuildPath(destPath, file.Name)

            If gFSO.FileExists(destFilePath) Then
                gSkippedCount = gSkippedCount + 1
                fout.WriteLine "SKIPPED : " & file.Name
            Else
                On Error Resume Next
                gFSO.CopyFile file.Path, destFilePath, False
                If Err.Number <> 0 Then
                    gErrorCount = gErrorCount + 1
                    fout.WriteLine "ERROR   : Copy failed"
                    fout.WriteLine "          From: " & file.Path
                    fout.WriteLine "          To  : " & destFilePath
                    fout.WriteLine "          Msg : " & Err.Description
                    Err.Clear
                Else
                    gCopiedCount = gCopiedCount + 1
                    fout.WriteLine "COPIED  : " & file.Name
                End If
                On Error GoTo 0
            End If
        End If
    Next

    If localMatchCount = 0 Then
        fout.WriteLine "NO FILES MATCHED IN: " & sourcePath
    End If
End Sub

Function SplitPatterns(s)
    Dim cleaned, parts, i
    cleaned = Replace(s, vbCr, "")
    cleaned = Replace(cleaned, vbLf, "")
    parts = Split(cleaned, ",")
    For i = LBound(parts) To UBound(parts)
        parts(i) = Trim(parts(i))
    Next
    SplitPatterns = parts
End Function

Function LoadIni(filePath)
    Dim dict, ts, line, section, key, value, p, fullKey

    Set dict = CreateObject("Scripting.Dictionary")
    section = ""

    If Not gFSO.FileExists(filePath) Then
        fout.WriteLine "ERROR: INI file not found: " & filePath
        Set LoadIni = dict
        Exit Function
    End If

    Set ts = gFSO.OpenTextFile(filePath, 1, False)

    Do Until ts.AtEndOfStream
        line = Trim(ts.ReadLine)
        If line <> "" Then
            If Left(line, 1) <> ";" And Left(line, 1) <> "#" Then
                If Left(line, 1) = "[" And Right(line, 1) = "]" Then
                    section = Trim(Mid(line, 2, Len(line) - 2))
                Else
                    p = InStr(line, "=")
                    If p > 0 Then
                        key = Trim(Left(line, p - 1))
                        value = Trim(Mid(line, p + 1))
                        fullKey = UCase(section & "." & key)
                        dict(fullKey) = value
                    End If
                End If
            End If
        End If
    Loop

    ts.Close
    Set LoadIni = dict
End Function

Function GetCfg(dict, section, key, defaultValue)
    Dim fullKey
    fullKey = UCase(section & "." & key)
    If dict.Exists(fullKey) Then
        GetCfg = dict(fullKey)
    Else
        GetCfg = defaultValue
    End If
End Function