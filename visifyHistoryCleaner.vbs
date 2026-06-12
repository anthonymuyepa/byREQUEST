' ============================================================================
' ProcessLine snippet
' ----------------------------------------------------------------------------
' Assumes the host script provides:
'   srec  - current line read from input
'   fout  - open output TextStream (created with CreateTextFile)
'
' Call LoadExcludeList() ONCE before the main read loop.
' Call ProcessLine() for each input line.
' ============================================================================

Const EXCLUDE_FILE   = "C:\Data\ExcludeList.txt"
Const CASE_SENSITIVE = False   ' True = case-sensitive substring match

Dim gExcludeDict      ' module-level dictionary of exclusion tokens
Dim gExcludeKeys      ' cached array of keys for fast iteration

' ----------------------------------------------------------------------------
' LoadExcludeList - call once at startup
' ----------------------------------------------------------------------------
Sub startDoc
    Dim fso, ts, token

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set gExcludeDict = CreateObject("Scripting.Dictionary")

    If CASE_SENSITIVE Then
        gExcludeDict.CompareMode = 0   ' BinaryCompare
    Else
        gExcludeDict.CompareMode = 1   ' TextCompare
    End If

    If Not fso.FileExists(EXCLUDE_FILE) Then
        WScript.Echo "ERROR: Exclude list not found: " & EXCLUDE_FILE
        WScript.Quit 1
    End If

    Set ts = fso.OpenTextFile(EXCLUDE_FILE, 1, False)
    Do While Not ts.AtEndOfStream
        token = Trim(ts.ReadLine)
        If Len(token) > 0 Then
            If Not gExcludeDict.Exists(token) Then gExcludeDict.Add token, True
        End If
    Loop
    ts.Close

    gExcludeKeys = gExcludeDict.Keys
End Sub

' ----------------------------------------------------------------------------
' ProcessLine - writes srec to fout unless it contains an excluded token
' ----------------------------------------------------------------------------
Sub ProcessLine()
    Dim hay, needle, i

    If CASE_SENSITIVE Then
        hay = srec
    Else
        hay = UCase(srec)
    End If

    For i = 0 To UBound(gExcludeKeys)
        If CASE_SENSITIVE Then
            needle = gExcludeKeys(i)
        Else
            needle = UCase(gExcludeKeys(i))
        End If
        If InStr(1, hay, needle, 0) > 0 Then Exit Sub   ' skip line
    Next

    fout.WriteLine srec
End Sub