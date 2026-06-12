Option Explicit

' =========================
' CONFIG
' =========================
Const EMAIL_LIST_PATH = "R:\byREQUEST\Templates\acct_email_list.txt"  ' CSV-style TXT
Const DEFAULT_EMAIL = "anthonym@hillary.com"
 
 
Sub StartDoc()
    Dim acct, emailAddr, filename
	filename = Spoolfile.Name
	
    acct = Replace(LCase(filename), ".pdf", "")
	 
    emailAddr = FindEmailFromTxt(EMAIL_LIST_PATH, acct)
 MsgBox "Got emailAddr: " & emailAddr

Dim emails, i

' Assume emailAddr is already defined, e.g.:
' emailAddr = "user1@example.com;user2@example.com;user3@example.com"

emails = Split(emailAddr, ";")

For i = 0 To UBound(emails)
    If i >= 9 Then Exit For ' Stop after 9
    Select Case i
        Case 0: Spoolfile.User1 = Trim(emails(i))
        Case 1: Spoolfile.User2 = Trim(emails(i))
        Case 2: Spoolfile.User3 = Trim(emails(i))
        Case 3: Spoolfile.User4 = Trim(emails(i))
        Case 4: Spoolfile.User5 = Trim(emails(i))
        Case 5: Spoolfile.User6 = Trim(emails(i))
        Case 6: Spoolfile.User7 = Trim(emails(i))
        Case 7: Spoolfile.User8 = Trim(emails(i))
        Case 8: Spoolfile.User9 = Trim(emails(i))
    End Select
Next	
	
End Sub

' =========================
' FIND EMAIL DIRECTLY IN TXT
' =========================


Function FindEmailFromTxt(path, acct)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(path) Then
        FindEmailFromTxt = DEFAULT_EMAIL
        Exit Function
    End If

    Dim ts
    Set ts = fso.OpenTextFile(path, 1, False) ' ForReading

    Dim line, row, acctInFile, emails, isHeader
    isHeader = True  ' assume first line is header

    Do Until ts.AtEndOfStream
        line = Trim(ts.ReadLine)
        If line <> "" Then
            row = Split(line, ",")
            Dim i
            For i = 0 To UBound(row)
                row(i) = StripQuotes(Trim(CStr(row(i))))
            Next

            ' Skip header row
            If isHeader And InStr(LCase(row(0)), "acct") > 0 Then
                isHeader = False
            Else
                isHeader = False
                acctInFile = CStr(row(0))

                If acctInFile = acct Then
                    ' ---- NEW LOGIC: combine all e-mail fields ----
                    emails = ""
                    For i = 1 To UBound(row)
                        If Trim(row(i)) <> "" Then
                            If emails <> "" Then emails = emails & ";"
                            emails = emails & Trim(row(i))
                        End If
                    Next

                    ts.Close
                    If emails <> "" Then
                        FindEmailFromTxt = emails
                    Else
                        FindEmailFromTxt = DEFAULT_EMAIL
                    End If
                    Exit Function
                End If
            End If
        End If
    Loop

    ts.Close
    FindEmailFromTxt = DEFAULT_EMAIL
End Function


Function ChooseOnePerRow(primary, secondary)
    primary = Trim(primary)
    secondary = Trim(secondary)
    If primary <> "" Then
        ChooseOnePerRow = NormalizeEmailList(primary)
    Else
        ChooseOnePerRow = secondary
    End If
End Function

Function NormalizeEmailList(list)
    Dim noSpaces
	noSpaces = Replace(list, " ", "")
    Dim parts, i, out
    parts = Split(noSpaces, ";")
    For i = 0 To UBound(parts)
        If Trim(parts(i)) <> "" Then
            If out <> "" Then out = out & ";"
            out = out & parts(i)
        End If
    Next
    NormalizeEmailList = out
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
