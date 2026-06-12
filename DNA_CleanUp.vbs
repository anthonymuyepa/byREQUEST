Option Explicit

Dim gFSO
Dim gCfg

Dim gIniPath
Dim gDestination
Dim gPatterns
Dim gRetentionDays

Dim gDeletedCount
Dim gSkippedCount
Dim gMatchedCount
Dim gErrorCount

Sub StartDoc()

    Set gFSO = CreateObject("Scripting.FileSystemObject")

    gIniPath = "S:\ConnexCU\Daily_Copy_DNA_Files.ini"

    gDeletedCount = 0
    gSkippedCount = 0
    gMatchedCount = 0
    gErrorCount = 0

    fout.WriteLine "DNA FILE CLEANUP PROCESS"

    Set gCfg = LoadIni(gIniPath)

    Dim sDestination, sPatterns, sRetentionDays
    Dim arrPatterns

    sDestination   = CStr(GetCfg(gCfg, "Paths", "Destination", ""))
    sPatterns      = CStr(GetCfg(gCfg, "Files", "Patterns", ""))
    sRetentionDays = CStr(GetCfg(gCfg, "Options", "RetentionDays", "3"))

    fout.WriteLine "DEBUG Destination raw  : [" & sDestination & "]"
    fout.WriteLine "DEBUG Patterns raw     : [" & sPatterns & "]"
    fout.WriteLine "DEBUG Retention raw    : [" & sRetentionDays & "]"
 
    gDestination = Trim(sDestination)
    gPatterns    = Trim(sPatterns)
 
	gRetentionDays = CInt(0 & GetCfg(gCfg, "Options", "RetentionDays", "3"))
	If gRetentionDays <= 0 Then gRetentionDays = 3
	
	
    If gDestination = "" Then
        fout.WriteLine "ERROR: Missing [Paths] Destination in INI."
        Exit Sub
    End If
 
    If gPatterns = "" Then
        fout.WriteLine "ERROR: Missing [Files] Patterns in INI."
        Exit Sub
    End If
 
    If Not gFSO.FileExists(gIniPath) Then
        fout.WriteLine "ERROR: INI file not found: " & gIniPath
        Exit Sub
    End If

    If Not gFSO.FolderExists(gDestination) Then
        fout.WriteLine "ERROR: Destination folder not found: " & gDestination
        Exit Sub
    End If

    arrPatterns = SplitPatterns(gPatterns)

    fout.WriteLine ""
    fout.WriteLine String(110, "-")
    fout.WriteLine "INI File      : " & gIniPath
    fout.WriteLine "Destination   : " & gDestination
    fout.WriteLine "Patterns      : " & gPatterns
    fout.WriteLine "RetentionDays : " & gRetentionDays
    fout.WriteLine "Pattern count : " & (UBound(arrPatterns) - LBound(arrPatterns) + 1)
    fout.WriteLine ""
    fout.WriteLine "File Status"
    fout.WriteLine String(110, "-")

    On Error Resume Next
    CleanupOldFiles gDestination, arrPatterns, gRetentionDays
    If Err.Number <> 0 Then
        fout.WriteLine "ERROR: CleanupOldFiles failed"
        fout.WriteLine "ERROR NUMBER: " & Err.Number
        fout.WriteLine "ERROR DESC  : " & Err.Description
        gErrorCount = gErrorCount + 1
        Err.Clear
    End If
    On Error GoTo 0

    fout.WriteLine "DEBUG: Returned from CleanupOldFiles"

End Sub

Sub ProcessLine()
    ' No line-by-line processing needed
End Sub

Sub CloseDoc()

    fout.WriteLine ""
    fout.WriteLine String(110, "-")
    fout.WriteLine "Matched Files : " & gMatchedCount
    fout.WriteLine "Deleted Files : " & gDeletedCount
    fout.WriteLine "Skipped Files : " & gSkippedCount
    fout.WriteLine "Errors        : " & gErrorCount
    fout.WriteLine "Completed."

End Sub



Sub CleanupOldFiles(folderPath, filePatterns, retentionDays)
    Dim folderObj, file
    Dim i, matched, patternText, ageDays

    fout.WriteLine "DEBUG: Entered CleanupOldFiles"
    fout.WriteLine "DEBUG: folderPath = [" & folderPath & "]"
    fout.WriteLine "DEBUG: retentionDays = [" & retentionDays & "]"

    If Not gFSO.FolderExists(folderPath) Then
        fout.WriteLine "ERROR: Folder does not exist: " & folderPath
        gErrorCount = gErrorCount + 1
        Exit Sub
    End If

    Set folderObj = gFSO.GetFolder(folderPath)

    fout.WriteLine "DEBUG: Folder opened successfully"

    For Each file In folderObj.Files
        matched = False
        fout.WriteLine "FOUND   : " & file.Path

        For i = LBound(filePatterns) To UBound(filePatterns)
            patternText = Trim(CStr(filePatterns(i)))

            If patternText <> "" Then
                fout.WriteLine "CHECK   : [" & file.Name & "] AGAINST [" & patternText & "]"
                If InStr(1, UCase(file.Name), UCase(patternText), vbTextCompare) > 0 Then
                    matched = True
                    Exit For
                End If
            End If
        Next

        If matched Then
            gMatchedCount = gMatchedCount + 1
            ageDays = DateDiff("d", file.DateLastModified, Now)

            If ageDays >= retentionDays Then
                On Error Resume Next
                file.Delete True
                If Err.Number <> 0 Then
                    gErrorCount = gErrorCount + 1
                    fout.WriteLine "ERROR   : Could not delete " & file.Path
                    fout.WriteLine "          AgeDays  : " & ageDays
                    fout.WriteLine "          Modified : " & file.DateLastModified
                    fout.WriteLine "          Msg      : " & Err.Description
                    Err.Clear
                Else
                    gDeletedCount = gDeletedCount + 1
                    fout.WriteLine "DELETED : " & file.Path & _
                                   " | AgeDays=" & ageDays & _
                                   " | Modified=" & file.DateLastModified
                End If
                On Error GoTo 0
            Else
                gSkippedCount = gSkippedCount + 1
                fout.WriteLine "KEPT    : " & file.Path & _
                               " | AgeDays=" & ageDays & _
                               " | Modified=" & file.DateLastModified
            End If
        Else
            fout.WriteLine "NO MATCH: " & file.Path
        End If

        fout.WriteLine ""
    Next

    fout.WriteLine "DEBUG: Finished scanning folder"
End Sub
 

Function SplitPatterns(s)
    Dim cleaned, parts, i

    cleaned = Replace(s, vbCr, "")
    cleaned = Replace(cleaned, vbLf, "")
    cleaned = Replace(cleaned, """", "")
    cleaned = Replace(cleaned, "*", "")
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

    GetCfg = defaultValue
    fullKey = UCase(section & "." & key)

    If IsObject(dict) Then
        If dict.Exists(fullKey) Then
            If IsNull(dict(fullKey)) Then
                GetCfg = defaultValue
            Else
                GetCfg = CStr(dict(fullKey))
            End If
        End If
    End If
End Function