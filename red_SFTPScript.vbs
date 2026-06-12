'===============================================================
' byREQUEST SFTP Automation Script
' Purpose:
'   - Connect to SFTP server using WinSCP
'   - Download files from remote root /
'   - Upload files to remote /uploads
'   - Archive uploaded files after successful upload
'   - Log all activity
'
' Schedule:
'   Configure the byREQUEST spooler to run every 3 hours.
'
' Requirements:
'   - WinSCP installed on the byREQUEST server
'   - SFTP credentials or saved WinSCP site
'   - No shell access required on the SFTP server
'
'===============================================================

Option Explicit

'===============================================================
' CONFIGURATION
'===============================================================

Dim WINSCP_PATH
Dim USE_SAVED_SITE
Dim SAVED_SITE_NAME

Dim SFTP_HOST
Dim SFTP_USERNAME
Dim SFTP_PASSWORD
Dim SFTP_HOSTKEY

Dim REMOTE_DOWNLOAD_FOLDER
Dim REMOTE_UPLOAD_FOLDER

Dim BASE_FOLDER
Dim LOCAL_DOWNLOAD_FOLDER
Dim LOCAL_UPLOAD_FOLDER
Dim LOCAL_ARCHIVE_FOLDER
Dim LOG_FOLDER

Dim DELETE_REMOTE_AFTER_DOWNLOAD
Dim ARCHIVE_LOCAL_UPLOADS_AFTER_SUCCESS

WINSCP_PATH = "C:\Program Files (x86)\WinSCP\WinSCP.com"

' Option 1: Use saved WinSCP site
' Recommended for production because password is not stored in this script.
USE_SAVED_SITE = False
SAVED_SITE_NAME = "MySavedSFTPSite"

' Option 2: Use direct SFTP connection
SFTP_HOST = "test.rebex.net"
SFTP_USERNAME = "demo"
SFTP_PASSWORD = "password"
SFTP_HOSTKEY = "*"


SFTP_HOSTKEY = "ssh-rsa 2048 REPLACE_WITH_REAL_HOSTKEY"

REMOTE_DOWNLOAD_FOLDER = "/"
REMOTE_UPLOAD_FOLDER = "/uploads"

BASE_FOLDER = "C:\SFTPJob"
LOCAL_DOWNLOAD_FOLDER = BASE_FOLDER & "\Download"
LOCAL_UPLOAD_FOLDER = BASE_FOLDER & "\Upload"
LOCAL_ARCHIVE_FOLDER = BASE_FOLDER & "\Archive"
LOG_FOLDER = BASE_FOLDER & "\Logs"

' Set to True only if files should be removed from the SFTP server after download.
DELETE_REMOTE_AFTER_DOWNLOAD = False

' Prevents same local upload files from being uploaded repeatedly every 3 hours.
ARCHIVE_LOCAL_UPLOADS_AFTER_SUCCESS = True


'===============================================================
' GLOBAL OBJECTS
'===============================================================

Dim fso
Dim shell
Dim gScriptFile
Dim gWinSCPLogFile
Dim gStatusLogFile
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

    EnsureFolder BASE_FOLDER
    EnsureFolder LOCAL_DOWNLOAD_FOLDER
    EnsureFolder LOCAL_UPLOAD_FOLDER
    EnsureFolder LOCAL_ARCHIVE_FOLDER
    EnsureFolder LOG_FOLDER

    gStatusLogFile = LOG_FOLDER & "\SFTP_Status_" & GetDateStamp() & ".log"
    gWinSCPLogFile = LOG_FOLDER & "\WinSCP_" & GetDateTimeStamp() & ".log"
    gScriptFile = BASE_FOLDER & "\SFTP_WinSCP_Script_" & GetDateTimeStamp() & ".txt"

    WriteLog "===================================================="
    WriteLog "SFTP automation started: " & Now
    WriteLog "Remote download folder: " & REMOTE_DOWNLOAD_FOLDER
    WriteLog "Remote upload folder: " & REMOTE_UPLOAD_FOLDER
    WriteLog "Local download folder: " & LOCAL_DOWNLOAD_FOLDER
    WriteLog "Local upload folder: " & LOCAL_UPLOAD_FOLDER

    If Not fso.FileExists(WINSCP_PATH) Then
        WriteLog "ERROR: WinSCP.com was not found at: " & WINSCP_PATH
        gHasError = True
        Exit Sub
    End If

    CreateWinSCPScript

    If gHasError = True Then
        WriteLog "ERROR: WinSCP script could not be created."
        Exit Sub
    End If

    RunWinSCP

    If gExitCode = 0 Then
        WriteLog "WinSCP completed successfully."

        If ARCHIVE_LOCAL_UPLOADS_AFTER_SUCCESS = True Then
            ArchiveUploadedFiles
        End If
    Else
        WriteLog "ERROR: WinSCP failed. Exit code: " & gExitCode
        WriteLog "Check WinSCP log: " & gWinSCPLogFile
        gHasError = True
    End If

End Sub


Sub ProcessLine(ByVal line)

    ' This SFTP job does not need to process report lines.
    ' byREQUEST may still call ProcessLine depending on the spooler type.
    ' Leave this intentionally blank.

End Sub


Sub CloseDoc()

    On Error Resume Next

    If fso.FileExists(gScriptFile) Then
        fso.DeleteFile gScriptFile, True
    End If

    If gHasError = True Then
        WriteLog "SFTP automation completed with errors: " & Now
    Else
        WriteLog "SFTP automation completed successfully: " & Now
    End If

    WriteLog "===================================================="

End Sub


'===============================================================
' CREATE WINSCP SCRIPT
'===============================================================

Sub CreateWinSCPScript()

    On Error Resume Next

    Dim ts
    Set ts = fso.CreateTextFile(gScriptFile, True)

    If Err.Number <> 0 Then
        WriteLog "ERROR creating WinSCP script file: " & Err.Description
        gHasError = True
        Err.Clear
        Exit Sub
    End If

    ts.WriteLine "option batch abort"
    ts.WriteLine "option confirm off"

    If USE_SAVED_SITE = True Then
        ts.WriteLine "open " & Chr(34) & SAVED_SITE_NAME & Chr(34)
    Else
        ts.WriteLine "open sftp://" & EncodeUrlPart(SFTP_USERNAME) & ":" & EncodeUrlPart(SFTP_PASSWORD) & "@" & SFTP_HOST & "/ -hostkey=" & Chr(34) & SFTP_HOSTKEY & Chr(34)
    End If

    ts.WriteLine ""

    ' Download from remote root /
    ts.WriteLine "# Download files from remote root"
    ts.WriteLine "cd " & Chr(34) & REMOTE_DOWNLOAD_FOLDER & Chr(34)
    ts.WriteLine "lcd " & Chr(34) & LOCAL_DOWNLOAD_FOLDER & Chr(34)

    If DELETE_REMOTE_AFTER_DOWNLOAD = True Then
        ts.WriteLine "get -delete *.*"
    Else
        ts.WriteLine "get *.*"
    End If

    ts.WriteLine ""

    ' Upload to remote /uploads
    ts.WriteLine "# Upload files to remote uploads folder"
    ts.WriteLine "cd " & Chr(34) & REMOTE_UPLOAD_FOLDER & Chr(34)
    ts.WriteLine "lcd " & Chr(34) & LOCAL_UPLOAD_FOLDER & Chr(34)
    ts.WriteLine "put *.*"

    ts.WriteLine ""
    ts.WriteLine "exit"

    ts.Close

    If Err.Number <> 0 Then
        WriteLog "ERROR writing WinSCP script: " & Err.Description
        gHasError = True
        Err.Clear
    Else
        WriteLog "Created WinSCP script: " & gScriptFile
    End If

End Sub


'===============================================================
' RUN WINSCP
'===============================================================

Sub RunWinSCP()

    On Error Resume Next

    Dim cmd

    cmd = Chr(34) & WINSCP_PATH & Chr(34) & _
          " /script=" & Chr(34) & gScriptFile & Chr(34) & _
          " /log=" & Chr(34) & gWinSCPLogFile & Chr(34)

    WriteLog "Running WinSCP command."
    WriteLog cmd

    gExitCode = shell.Run(cmd, 0, True)

    If Err.Number <> 0 Then
        WriteLog "ERROR running WinSCP: " & Err.Description
        gHasError = True
        gExitCode = 999
        Err.Clear
    End If

End Sub


'===============================================================
' ARCHIVE UPLOADED FILES
'===============================================================

Sub ArchiveUploadedFiles()

    On Error Resume Next

    Dim folder
    Dim file
    Dim destFile

    If Not fso.FolderExists(LOCAL_UPLOAD_FOLDER) Then
        WriteLog "Upload folder does not exist. Nothing to archive."
        Exit Sub
    End If

    Set folder = fso.GetFolder(LOCAL_UPLOAD_FOLDER)

    If folder.Files.Count = 0 Then
        WriteLog "No uploaded files found to archive."
        Exit Sub
    End If

    WriteLog "Archiving uploaded files."

    For Each file In folder.Files

        destFile = LOCAL_ARCHIVE_FOLDER & "\" & GetDateTimeStamp() & "_" & file.Name

        Err.Clear
        fso.MoveFile file.Path, destFile

        If Err.Number <> 0 Then
            WriteLog "WARNING: Could not archive file: " & file.Path & " - " & Err.Description
            Err.Clear
        Else
            WriteLog "Archived uploaded file: " & file.Name
        End If

    Next

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
        WriteLog "ERROR creating folder: " & folderPath & " - " & Err.Description
        gHasError = True
        Err.Clear
    End If

End Sub


Sub WriteLog(ByVal msg)

    On Error Resume Next

    Dim ts

    If Len(gStatusLogFile) = 0 Then
        gStatusLogFile = "C:\SFTPJob\SFTP_Status_" & GetDateStamp() & ".log"
    End If

    Set ts = fso.OpenTextFile(gStatusLogFile, 8, True)
    ts.WriteLine FormatDateTime(Now, 2) & " " & FormatDateTime(Now, 3) & " - " & msg
    ts.Close

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


Function EncodeUrlPart(ByVal value)

    ' Minimal URL encoding for username/password.
    ' Handles common special characters used in passwords.
    Dim result

    result = value
    result = Replace(result, "%", "%25")
    result = Replace(result, " ", "%20")
    result = Replace(result, "@", "%40")
    result = Replace(result, ":", "%3A")
    result = Replace(result, "/", "%2F")
    result = Replace(result, "\", "%5C")
    result = Replace(result, "#", "%23")
    result = Replace(result, "&", "%26")
    result = Replace(result, "?", "%3F")
    result = Replace(result, "+", "%2B")

    EncodeUrlPart = result

End Function