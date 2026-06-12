Option Explicit

Dim gHeaderPart
Dim gReportDate, gHeaderParsed
Dim gSeqNumber

Sub StartDoc()

    Dim i, line

    fout.WriteLine "Period     Account No Name                       Type     Description                     LastPaid   Due Date   Days DQ         Balance   Code"

    ' ---------------------------------------------------
    ' Read the first few lines only to locate the header
    ' ---------------------------------------------------
    For i = 1 To 2
        line = GetLine

        If Not gHeaderParsed Then
            If InStr(1, line, "Delinquent Consumer", vbTextCompare) > 0 And _
               InStr(1, line, "Seq", vbTextCompare) > 0 Then

                gHeaderPart = line

                ' Extract report period date from:
                ' Credit Union Consumer Secured 02/28/26 at 03/01/26 01:39 Seq 266006 Page
                gReportDate = GetEffectiveDate(line)
 
                gHeaderParsed = True
                Exit For
            End If
        End If
    Next

End Sub

Sub ProcessLine()

    If Trim(srec) = "" Then Exit Sub
    If gReportDate = "" Then Exit Sub

    ' Only process lines that appear to start with an account number
    If IsNumeric(Left(Trim(srec), 1)) Then

        ' Keep only expected detail-line lengths
        If Len(srec) > 70  Then


            fout.WriteLine _
                PadRight(gReportDate, 10) & " " & _
                srec

        End If
    End If

End Sub

Function GetEffectiveDate(ByVal sLine)
    Dim re, m

    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "(\d{2}/\d{2}/\d{2})\s+at\s+\d{2}/\d{2}/\d{2}"
    re.IgnoreCase = True
    re.Global = False

    If re.Test(sLine) Then
        Set m = re.Execute(sLine)(0)
        GetEffectiveDate = m.SubMatches(0)
    Else
        GetEffectiveDate = ""
    End If
End Function


Function PadRight(ByVal sText, ByVal n)
    Dim s
    s = CStr(sText)
    If Len(s) >= n Then
        PadRight = Left(s, n)
    Else
        PadRight = s & String(n - Len(s), " ")
    End If
End Function