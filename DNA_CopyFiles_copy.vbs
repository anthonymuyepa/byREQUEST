Option Explicit

Dim gFSO
Dim gShell
Dim gCfg

Dim gIniPath
Dim gSourceRoot
Dim gDestination
Dim gPatterns

Dim gCopiedCount
Dim gSkippedCount
Dim gMatchedCount
Dim gLatestFolder

Sub StartDoc()

    Set gFSO = CreateObject("Scripting.FileSystemObject")
    Set gShell = CreateObject("WScript.Shell")

    gIniPath = "C:\Scripts\Daily_Copy_DNA_Files.ini"

    gCopiedCount = 0
    gSkippedCount = 0
    gMatchedCount = 0
    gLatestFolder = ""

    fout.WriteLine "DNA FILE COPY PROCESS"
    fout.WriteLine String(90, "-")

    Set gCfg = LoadIni(gIniPath)

    gSourceRoot = Trim(GetCfg(gCfg, "Paths", "SourceRoot", ""))
    gDestination = Trim(GetCfg(gCfg, "Paths", "Destination", ""))
    gPatterns = Trim(GetCfg(gCfg, "Files", "Patterns", ""))

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

    fout.WriteLine "INI File      : " & gIniPath
    fout.WriteLine "Source Root   : " & gSourceRoot
    fout.WriteLine "Destination   : " & gDestination
    fout.WriteLine "Patterns      : " & gPatterns
    fout.WriteLine ""

    If Not gFSO.FolderExists(gSourceRoot) Then
        fout.WriteLine "ERROR: Source root folder not found: " & gSourceRoot
        Exit Sub
    End If

    If Not gFSO.FolderExists(gDestination) Then
        gFSO.CreateFolder gDestination
        fout.WriteLine "Created destination folder: " & gDestination
    End If

    gLatestFolder = GetNewestDateFolder(gSourceRoot)

    If gLatestFolder = "" Then
        fout.WriteLine "ERROR: No valid yyyymmdd folders found under: " & gSourceRoot
        Exit Sub
    End If

    fout.WriteLine "Latest Folder : " & gLatestFolder
    fout.WriteLine ""
    fout.WriteLine "File Status"
    fout.WriteLine String(90, "-")

    CopyMatchingFiles gLatestFolder, gDestination, SplitPatterns(gPatterns)

End Sub

Sub ProcessLine()
    ' No per-line processing needed for this job.
End Sub

Sub CloseDoc()

    fout.WriteLine ""
    fout.WriteLine String(90, "-")
    fout.WriteLine "Matched Files : " & gMatchedCount
    fout.WriteLine "Copied Files  : " & gCopiedCount
    fout.WriteLine "Skipped Files : " & gSkippedCount
    fout.WriteLine "Completed."

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

Function IsYYYYMMDD(folderName)
    Dim yyyy, mm, dd, dt

    IsYYYYMMDD = False

    If Len(folderName) <> 8 Then Exit Function
    If Not IsNumeric(folderName) Then Exit Function

    yyyy = CInt(Left(folderName, 4))
    mm   = CInt(Mid(folderName, 5, 2))
    dd   = CInt(Right(folderName, 2))

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
    Dim matched, i, destFilePath

    Set sourceFolder = gFSO.GetFolder(sourcePath)

    For Each file In sourceFolder.Files
        matched = False

        For i = LBound(filePatterns) To UBound(filePatterns)
            If Trim(filePatterns(i)) <> "" Then
                If UCase(file.Name) Like UCase(Trim(filePatterns(i))) Then
                    matched = True
                    Exit For
                End If
            End If
        Next

        If matched Then
            gMatchedCount = gMatchedCount + 1
            destFilePath = gFSO.BuildPath(destPath, file.Name)

            If gFSO.FileExists(destFilePath) Then
                gSkippedCount = gSkippedCount + 1
                fout.WriteLine "SKIPPED : " & file.Name
            Else
                gFSO.CopyFile file.Path, destFilePath, False
                gCopiedCount = gCopiedCount + 1
                fout.WriteLine "COPIED  : " & file.Name
            End If
        End If
    Next
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