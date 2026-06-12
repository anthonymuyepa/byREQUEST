'===============================================================
' byREQUEST SFTP Download Script
'
' Purpose:
'   - Connect to saved WinSCP site: FileServeOptis
'   - Download all files from remote download folder
'   - Download files only, no subfolders
'   - Save files to local folder:
'       C:\SafireFileDrop\FiservFileIntake
'   - Write messages to byREQUEST using fout.WriteLine
'
' Notes:
'   - This script is NOT per file.
'   - It connects once and downloads all available files.
'   - It does not upload.
'===============================================================

Option Explicit

'===============================================================
' CONFIGURATION
'===============================================================

Dim WINSCP_PATH
Dim SAVED_SITE_NAME

Dim REMOTE_DOWNLOAD_FOLDER
Dim LOCAL_DOWNLOAD_FOLDER

Dim DELETE_REMOTE_AFTER_DOWNLOAD

WINSCP_PATH = "C:\Program Files (x86)\WinSCP\WinSCP.com"

' Saved WinSCP site name
SAVED_SITE_NAME = "FileServeOptis"

' Remote download folder
' Change this if the remote source folder is not root.
REMOTE_DOWNLOAD_FOLDER = "/"

' Local destination folder
LOCAL_DOWNLOAD_FOLDER = "C:\SafireFileDrop\FiservFileIntake"

' True  = delete files from remote server after successful download
' False = leave files on remote server
DELETE_REMOTE_AFTER_DOWNLOAD = False


'===============================================================
' GLOBAL VARIABLES
'===============================================================

Dim fso
Dim shell

Dim gScriptFile
Dim gExitCode
Dim gHasError

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

gHasError = False
gExitCode = 0


'===============================================================
' byREQUEST ENTRY POINTS
'===============================================================

Sub StartDoc()

    On Error Resume Next

    WriteMsg "===================================================="
    WriteMsg "SFTP download started: " & Now
    WriteMsg "Saved site: " & SAVED_SITE_NAME
    WriteMsg "Remote download folder: " & REMOTE_DOWNLOAD_FOLDER
    WriteMsg "Local download folder: " & LOCAL_DOWNLOAD_FOLDER
    WriteMsg "Delete remote files after download: " & CStr(DELETE_REMOTE_AFTER_DOWNLOAD)

    EnsureFolder LOCAL_DOWNLOAD_FOLDER

    If Not fso.FileExists(WINSCP_PATH) Then
        WriteMsg "ERROR: WinSCP.com was not found at: " & WINSCP_PATH
        gHasError = True
        Exit Sub
    End If

    ' Temporary WinSCP script goes to Windows TEMP folder.
    gScriptFile = shell.ExpandEnvironmentStrings("%TEMP%") & "\SFTP_Download_WinSCP_Script_" & GetDateTimeStamp() & ".txt"

    CreateWinSCPScript

    If gHasError = True Then
        WriteMsg "ERROR: WinSCP download script could not be created."
        Exit Sub
    End If

    RunWinSCP

    If gExitCode = 0 Then
        WriteMsg "WinSCP download completed successfully."
    Else
        WriteMsg "ERROR: WinSCP download failed. Exit code: " & gExitCode
        gHasError = True
    End If

End Sub


Sub ProcessLine(ByVal line)

    ' Download job does not process input report lines.

End Sub


Sub CloseDoc()

    On Error Resume Next

    If Len(gScriptFile) > 0 Then
        If fso.FileExists(gScriptFile) Then
            fso.DeleteFile gScriptFile, True
        End If
    End If

    If gHasError = True Then
        WriteMsg "SFTP download completed with errors: " & Now
    Else
        WriteMsg "SFTP download completed successfully: " & Now
    End If

    WriteMsg "===================================================="

End Sub


'===============================================================
' CREATE WINSCP SCRIPT
'===============================================================

Sub CreateWinSCPScript()

    On Error Resume Next

    Dim ts

    Set ts = fso.CreateTextFile(gScriptFile, True)

    If Err.Number <> 0 Then
        WriteMsg "ERROR creating WinSCP download script file: " & Err.Description
        gHasError = True
        Err.Clear
        Exit Sub
    End If

    ts.WriteLine "option batch abort"
    ts.WriteLine "option confirm off"
    ts.WriteLine "open " & Chr(34) & SAVED_SITE_NAME & Chr(34)
    ts.WriteLine ""

    ts.WriteLine "# Download files only. Do not download directories."
    ts.WriteLine "cd " & Chr(34) & REMOTE_DOWNLOAD_FOLDER & Chr(34)
    ts.WriteLine "lcd " & Chr(34) & LOCAL_DOWNLOAD_FOLDER & Chr(34)

    If DELETE_REMOTE_AFTER_DOWNLOAD = True Then
        ts.WriteLine "get -delete -filemask=" & Chr(34) & "|*/" & Chr(34) & " *"
    Else
        ts.WriteLine "get -filemask=" & Chr(34) & "|*/" & Chr(34) & " *"
    End If

    ts.WriteLine ""
    ts.WriteLine "exit"

    ts.Close

    If Err.Number <> 0 Then
        WriteMsg "ERROR writing WinSCP download script: " & Err.Description
        gHasError = True
        Err.Clear
    Else
        WriteMsg "Created temporary WinSCP download script."
        WriteMsg "Download command will get files only, no subfolders."
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

    WriteMsg "Running WinSCP download command."

    gExitCode = shell.Run(cmd, 0, True)

    If Err.Number <> 0 Then
        WriteMsg "ERROR running WinSCP: " & Err.Description
        gHasError = True
        gExitCode = 999
        Err.Clear
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


Function GetDateTimeStamp()

    GetDateTimeStamp = Year(Now) & _
                       Right("0" & Month(Now), 2) & _
                       Right("0" & Day(Now), 2) & "_" & _
                       Right("0" & Hour(Now), 2) & _
                       Right("0" & Minute(Now), 2) & _
                       Right("0" & Second(Now), 2)

End Function