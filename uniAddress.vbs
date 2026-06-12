Option Explicit

' --- Settings ---
Const PRINT_HEADER_EACH_PAGE = False   ' True = print "Account#" header every time it appears (per page)
Const MIN_BODY_LINE_LEN       = 4      ' ignore super-short lines (<4 chars)

' --- State ---
Dim headerPrinted, reachedEnd

Sub startDoc
    headerPrinted = False
    reachedEnd    = False
End Sub

Sub ProcessLine
    If reachedEnd Then Exit Sub

    Dim line: line = srec

    ' Strip form-feed and trailing spaces; keep leading spaces for alignment
    line = Replace(line, Chr(12), "")
    line = RTrim(line)

    If Len(line) = 0 Then Exit Sub  ' skip blanks

    ' Detect the header line (allow leading spaces)
    If StartsWith(LTrim(line), "Account#", vbTextCompare) Then
        If (Not headerPrinted) Or PRINT_HEADER_EACH_PAGE Then
            fout.WriteLine line
            headerPrinted = True
        End If
        Exit Sub
    End If

    ' If we haven't seen the header yet, ignore everything until we do
    If Not headerPrinted Then Exit Sub

    ' Skip decorative / repeated page header lines
    If IsSkippableHeaderLine(line) Then Exit Sub

    ' Write body lines that are non-trivial
    If Len(line) >= MIN_BODY_LINE_LEN Then
        fout.WriteLine line
    End If
End Sub

' -------- Helpers --------

Function StartsWith(text, prefix, cmp)
    StartsWith = (Left(text, Len(prefix)) = prefix)
    If cmp = vbTextCompare Then StartsWith = (LCase(Left(text, Len(prefix))) = LCase(prefix))
End Function

Function IsSkippableHeaderLine(line)
    ' Skip: dashed rules, org banner, column heading
    If InStr(1, line, "---------------", vbBinaryCompare) > 0 Then IsSkippableHeaderLine = True: Exit Function
    If InStr(1, line, "UniWyo Credit", vbTextCompare)    > 0 Then IsSkippableHeaderLine = True: Exit Function
    If InStr(1, line, "Account Name", vbTextCompare)      > 0 Then IsSkippableHeaderLine = True: Exit Function
    IsSkippableHeaderLine = False
End Function

