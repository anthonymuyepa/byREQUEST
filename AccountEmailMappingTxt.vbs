Option Explicit

' =========================
' CONFIG
' =========================
Const EMAIL_LIST_PATH = "R:\byREQUEST\Templates\acct_email_list.txt"  ' CSV-style TXT
Const DEFAULT_EMAIL = "dfdfd@dfd.net"
 
 
Sub StartDoc()
    Dim acct, emailAddr, filename
	filename = Spoolfile.Name
	
    acct = Replace(LCase(filename), ".pdf", "")
	 
    emailAddr = FindEmailFromTxt(EMAIL_LIST_PATH, acct)
 MsgBox "Got emailAddr" & emailAddr
    Spoolfile.User1 = emailAddr
	
End Sub

' =========================
' FIND EMAIL DIRECTLY IN TXT
' =========================


Function FindEmailFromTxt(path, acct)
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then
        FindEmailFromTxt = DEFAULT_EMAIL
        Exit Function
    End If

    Dim ts: Set ts = fso.OpenTextFile(path, 1, False) ' ForReading
    Dim line, row, acctInFile, emails, isFirst, i
    isFirst = True

    Do Until ts.AtEndOfStream
        line = Trim(ts.ReadLine)
        If line <> "" Then
            row = Split(line, ",")                ' CSV only
            For i = 0 To UBound(row)
                row(i) = StripQuotes(Trim(CStr(row(i))))
            Next

            ' header check (first non-empty line)
            If isFirst Then
                isFirst = False
                If InStr(LCase(row(0)), "acct") > 0 Or InStr(LCase(row(0)), "account") > 0 Then
                    ' header -> skip this line
                Else
                    ' treat as data
                    acctInFile = CStr(row(0))
                    If acctInFile = acct Then
                        emails = BuildEmailList(row)
                        ts.Close
                        FindEmailFromTxt = IIf(emails <> "", emails, DEFAULT_EMAIL)
                        Exit Function
                    End If
                End If
            Else
                acctInFile = CStr(row(0))
                If acctInFile = acct Then
                    emails = BuildEmailList(row)
                    ts.Close
                    FindEmailFromTxt = IIf(emails <> "", emails, DEFAULT_EMAIL)
                    Exit Function
                End If
            End If
        End If
    Loop

    ts.Close
    FindEmailFromTxt = DEFAULT_EMAIL
End Function

' Join columns 2..N with ';' (also normalizes any ';' lists inside a single cell)
Function BuildEmailList(row)
    Dim emails, i, cell
    emails = ""
    For i = 1 To UBound(row)
        cell = Trim(row(i))
        If cell <> "" Then
            ' Normalize spaces, handle semicolon-separated lists
            cell = Replace(cell, " ", "")
            Dim parts, j
            parts = Split(cell, ";")
            For j = 0 To UBound(parts)
                If IsValidEmail(parts(j)) Then
                    If emails <> "" Then emails = emails & ";"
                    emails = emails & parts(j)
                End If
            Next
        End If
    Next
    BuildEmailList = emails
End Function



Function GetField(arr, idx)
    If IsArray(arr) And idx >= 0 And idx <= UBound(arr) Then
        GetField = arr(idx)
    Else
        GetField = ""
    End If
End Function

Function StripQuotes(s)
    If Len(s) >= 2 Then
        If (Left(s,1) = """" And Right(s,1) = """") _
        Or (Left(s,1) = "'" And Right(s,1) = "'") Then
            StripQuotes = Mid(s, 2, Len(s)-2)
            Exit Function
        End If
    End If
    StripQuotes = s
End Function


Function IsValidEmail(addr)
    Dim atPos, dotPos
    addr = Trim(LCase(addr))
    If addr = "" Then
        IsValidEmail = False
        Exit Function
    End If
    atPos = InStr(addr, "@")
    dotPos = InStrRev(addr, ".")
    ' Must contain "@" and a "." after "@", and no spaces
    If atPos > 1 And dotPos > atPos + 1 And InStr(addr, " ") = 0 Then
        IsValidEmail = True
    Else
        IsValidEmail = False
    End If
End Function

