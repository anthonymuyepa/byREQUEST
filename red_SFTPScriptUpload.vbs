'===============================================================
' byREQUEST SFTP Upload One File At A Time
'
' Purpose:
'   - byREQUEST calls this script once per spool file
'   - Reads current file using Spoolfile.Name
'   - Finds matching file directly in C:\SafireFileDrop\
'   - Uploads only that one file to SFTP directory
'   - Optionally moves uploaded file to C:\SafireFileDrop\Archive
'
' Notes:
'   - No folder upload
'   - No recursive upload
'   - No status log file
'   - Messages are written using fout.WriteLine
'===============================================================

Option Explicit

'===============================================================
' CONFIGURATION
'===============================================================

Dim WINSCP_PATH
Dim SAVED_SITE_NAME

Dim BASE_FOLDER
Dim ARCHIVE_FOLDER
Dim REMOTE_UPLOAD_FOLDER

Dim MOVE_TO_ARCHIVE_AFTER_UPLOAD

WINSCP_PATH = "C:\Program Files (x86)\WinSCP\WinSCP.com"
SAVED_SITE_NAME = "MySavedSFTPSite"

BASE_FOLDER = "C:\SafireFileDrop"
ARCHIVE_FOLDER = BASE_FOLDER & "\Archive"

REMOTE_UPLOAD_FOLDER = "/"

' True  = move uploaded file to Archive after successful upload
' False = leave uploaded file in C:\SafireFileDrop
MOVE_TO_ARCHIVE_AFTER_UPLOAD = True


'===============================================================
' GLOBAL VARIABLES
'===============================================================

Dim fso
Dim shell

Dim gScriptFile
Dim gExitCode
Dim gHasError

Dim gSpoolFileName
Dim gMatchedLocalFile

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

gHasError = False
gExitCode = 0
gSpoolFileName = ""
gMatchedLocalFile = ""


'===============================================================
' byREQUEST ENTRY POINTS
'===============================================================

Sub StartDoc()

    On Error Resume Next

    WriteMsg "===================================================="
    WriteMsg "SFTP upload started: " & Now

    EnsureFolder BASE_FOLDER

    If MOVE_TO_ARCHIVE_AFTER_UPLOAD = True Then
        EnsureFolder ARCHIVE_FOLDER
    End If

    ' Temporary WinSCP script goes to Windows TEMP folder.
    gScriptFile = shell.ExpandEnvironmentStrings("%TEMP%") & "\SFTP_Upload_WinSCP_Script_" & GetDateTimeStamp() & ".txt"

    ' byREQUEST current spool file name
    gSpoolFileName = Spoolfile.Name

    If Err.Number <> 0 Then
        WriteMsg "ERROR: Unable to read Spoolfile.Name - " & Err.Description
        gHasError = True
        Err.Clear
        Exit Sub
    End If

    gSpoolFileName = fso.GetFileName(gSpoolFileName)

    WriteMsg "byREQUEST Spoolfile.Name: " & gSpoolFileName
    WriteMsg "Base folder: " & BASE_FOLDER
    WriteMsg "Remote upload folder: " & REMOTE_UPLOAD_FOLDER
    WriteMsg "Move to archive after upload: " & CStr(MOVE_TO_ARCHIVE_AFTER_UPLOAD)

    If Not fso.FileExists(WINSCP_PATH) Then
        WriteMsg "ERROR: WinSCP.com was not found at: " & WINSCP_PATH
        gHasError = True
        Exit Sub
    End If

    gMatchedLocalFile = FindMatchingLocalFile(BASE_FOLDER, gSpoolFileName)

    If Len(gMatchedLocalFile) = 0 Then
        WriteMsg "ERROR: No matching file found in base folder for spool file: " & gSpoolFileName
        gHasError = True
        Exit Sub
    End If

    WriteMsg "Matched local file: " & gMatchedLocalFile

    CreateWinSCPScript

    If gHasError = True Then
        WriteMsg "ERROR: WinSCP script could not be created."
        Exit Sub
    End If

    RunWinSCP

    If gExitCode = 0 Then
        WriteMsg "WinSCP upload completed successfully."

        If MOVE_TO_ARCHIVE_AFTER_UPLOAD = True Then
            MoveUploadedFileToArchive gMatchedLocalFile
        Else
            WriteMsg "Archive disabled. Uploaded file left in base folder."
        End If
    Else
        WriteMsg "ERROR: WinSCP upload failed. Exit code: " & gExitCode
        gHasError = True
    End If

End Sub


Sub ProcessLine(ByVal line)

    ' Upload-only job does not process input lines.

End Sub


Sub CloseDoc()

    On Error Resume Next

    If Len(gScriptFile) > 0 Then
        If fso.FileExists(gScriptFile) Then
            fso.DeleteFile gScriptFile, True
        End If
    End If

    If gHasError = True Then
        WriteMsg "SFTP upload completed with errors: " & Now
    Else
        WriteMsg "SFTP upload completed successfully: " & Now
    End If

    WriteMsg "===================================================="

End Sub


'===============================================================
' FIND MATCHING LOCAL FILE
'===============================================================

Function FindMatchingLocalFile(ByVal folderPath, ByVal spoolName)

    On Error Resume Next

    Dim folder
    Dim file
    Dim spoolBase
    Dim fileBase

    Dim exactMatch
    Dim baseMatch
    Dim startsWithMatch

    FindMatchingLocalFile = ""

    exactMatch = ""
    baseMatch = ""
    startsWithMatch = ""

    If Not fso.FolderExists(folderPath) Then
        WriteMsg "ERROR: Base folder does not exist: " & folderPath
        Exit Function
    End If

    spoolName = Trim(fso.GetFileName(spoolName))
    spoolBase = GetBaseNameNoExt(spoolName)

    Set folder = fso.GetFolder(folderPath)

    WriteMsg "Searching base folder only. Subfolders ignored."
    WriteMsg "Spool name: " & spoolName
    WriteMsg "Spool base name: " & spoolBase

    For Each file In folder.Files

        fileBase = GetBaseNameNoExt(file.Name)

        WriteMsg "Checking candidate file: " & file.Name

        ' 1. Exact full filename match
        If LCase(file.Name) = LCase(spoolName) Then
            exactMatch = file.Path
            Exit For
        End If

        ' 2. Base filename match without extension
        If Len(baseMatch) = 0 Then
            If LCase(fileBase) = LCase(spoolBase) Then
                baseMatch = file.Path
            End If
        End If

        ' 3. Starts-with match
        If Len(startsWithMatch) = 0 Then
            If LCase(fileBase) Like LCase(spoolBase) & "*" Then
                startsWithMatch = file.Path
            End If
        End If

    Next

    If Len(exactMatch) > 0 Then
        FindMatchingLocalFile = exactMatch
        WriteMsg "Match type: exact full filename"
    ElseIf Len(baseMatch) > 0 Then
        FindMatchingLocalFile = baseMatch
        WriteMsg "Match type: base filename without extension"
    ElseIf Len(startsWithMatch) > 0 Then
        FindMatchingLocalFile = startsWithMatch
        WriteMsg "Match type: starts-with base filename"
    Else
        WriteMsg "No matching local file found."
    End If

End Function


Function GetBaseNameNoExt(ByVal fileName)

    On Error Resume Next

    Dim cleanName
    Dim dotPos

    cleanName = Trim(fso.GetFileName(fileName))
    dotPos = InStrRev(cleanName, ".")

    If dotPos > 1 Then
        GetBaseNameNoExt = Left(cleanName, dotPos - 1)
    Else
        GetBaseNameNoExt = cleanName
    End If

End Function


'===============================================================
' CREATE WINSCP SCRIPT
'===============================================================

Sub CreateWinSCPScript()

    On Error Resume Next

    Dim ts
    Dim localFileName

    localFileName = fso.GetFileName(gMatchedLocalFile)

    Set ts = fso.CreateTextFile(gScriptFile, True)

    If Err.Number <> 0 Then
        WriteMsg "ERROR creating WinSCP script file: " & Err.Description
        gHasError = True
        Err.Clear
        Exit Sub
    End If

    ts.WriteLine "option batch abort"
    ts.WriteLine "option confirm off"
    ts.WriteLine "open " & Chr(34) & SAVED_SITE_NAME & Chr(34)
    ts.WriteLine ""

    ts.WriteLine "# Upload one matched file only"
    ts.WriteLine "cd " & Chr(34) & REMOTE_UPLOAD_FOLDER & Chr(34)
    ts.WriteLine "lcd " & Chr(34) & BASE_FOLDER & Chr(34)

    ' Upload only the matched file.
    ' No wildcard. No directory upload. No recursion.
    ts.WriteLine "put " & Chr(34) & localFileName & Chr(34)

    ts.WriteLine ""
    ts.WriteLine "exit"

    ts.Close

    If Err.Number <> 0 Then
        WriteMsg "ERROR writing WinSCP script: " & Err.Description
        gHasError = True
        Err.Clear
    Else
        WriteMsg "Created temporary WinSCP script."
        WriteMsg "WinSCP will upload only: " & localFileName
    End If

End Sub


'===============================================================
' RUN WINSCP
'===============================================================

Sub RunWinSCP()

    On Error Resume Next

    Dim cmd

    cmd = Chr(34) & WINSCP_PATH & Chr(34) & _
          " /script=" & Chr(34) & gScriptFile & Chr(34)

    WriteMsg "Running WinSCP upload command."

    gExitCode = shell.Run(cmd, 0, True)

    If Err.Number <> 0 Then
        WriteMsg "ERROR running WinSCP: " & Err.Description
        gHasError = True
        gExitCode = 999
        Err.Clear
    End If

End Sub


'===============================================================
' OPTIONAL ARCHIVE
'===============================================================

Sub MoveUploadedFileToArchive(ByVal filePath)

    On Error Resume Next

    Dim fileName
    Dim destFile

    If Not fso.FileExists(filePath) Then
        WriteMsg "WARNING: Uploaded file no longer exists. Cannot archive: " & filePath
        Exit Sub
    End If

    EnsureFolder ARCHIVE_FOLDER

    fileName = fso.GetFileName(filePath)
    destFile = ARCHIVE_FOLDER & "\" & GetDateTimeStamp() & "_" & fileName

    fso.MoveFile filePath, destFile

    If Err.Number <> 0 Then
        WriteMsg "WARNING: Could not move uploaded file to archive: " & Err.Description
        Err.Clear
    Else
        WriteMsg "Moved uploaded file to archive: " & destFile
    End If

End Sub


'===============================================================
' HELPER FUNCTIONS
'===============================================================

Sub EnsureFolder(ByVal folderPath)

    On Error Resume Next

    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If

    If Err.Number <> 0 Then
        WriteMsg "ERROR creating folder: " & folderPath & " - " & Err.Description
        gHasError = True
        Err.Clear
    End If

End Sub


Sub WriteMsg(ByVal msg)

    On Error Resume Next

    fout.WriteLine msg

    If Err.Number <> 0 Then
        Err.Clear
    End If

End Sub


Function GetDateStamp()

    GetDateStamp = Year(Date) & _
                   Right("0" & Month(Date), 2) & _
                   Right("0" & Day(Date), 2)

End Function


Function GetDateTimeStamp()

    GetDateTimeStamp = Year(Now) & _
                       Right("0" & Month(Now), 2) & _
                       Right("0" & Day(Now), 2) & "_" & _
                       Right("0" & Hour(Now), 2) & _
                       Right("0" & Minute(Now), 2) & _
                       Right("0" & Second(Now), 2)

End Function