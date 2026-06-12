Option Explicit

'==========================================================
' byREQUEST VBScript - Flatten 3-srec records by CONCAT
'
' Adds 3 important fields from report header:
'   ReportName  (left 4 chars of ERTR header line)
'   Post_Date   (mid(srec,10,8) from ERTR header line)
'   ERTR_File   (Spoolfile.Name)
'
' Record pattern:
'   L1: first two chars are digits (TC)
'   L2: starts with "RTN" after leading spaces
'   L3: starts with "COMPANY NAME" after leading spaces
'
' Action for each record:
'   Pad L1/L2/L3 to 140, concat, remove headings,
'   then PREFIX Post_Date and APPEND ReportName + ERTR_File
'==========================================================

Dim PAD_LEN
Dim gStage, gL1, gL2

Dim gReportName, gPostDate, gERTRFile
Dim gGotHeaderInfo

Sub StartDoc()
    gStage = 0
    PAD_LEN = 140

    gReportName = ""
    gPostDate   = ""
    gERTRFile   = ""
    gGotHeaderInfo = False

	fout.WriteLine "Post_Date  TC  RDFI R&T#  DFI ACCOUNT        RTN AMOUNT INDIVIDUAL ID   INDIVIDUAL NAME         DD A  TRACE NUMBER    RRC ORC DRC  SRC BH#               RTN TRACE#       RTN RDFI# DEATHDATE INFO                                                      COMPANY NAME     COMPANY ID ERTR_File"
    On Error Resume Next
    gERTRFile = Spoolfile.Name
    On Error GoTo 0
End Sub

Sub ProcessLine()

    '---- capture ERTR header once (first occurrence) ----
    If Not gGotHeaderInfo Then
        If UCase(Left(srec, 4)) = "ERTR" Then
            gReportName = Left(srec, 4)
            ' ERTR     02-02-26 08:42:08 ...
			Dim rawDate, p, mm, dd, yy, dt

			rawDate = Trim(Mid(srec, 10, 8))   ' "02-02-26"

			p = Split(rawDate, "-")
			If UBound(p) = 2 Then
				mm = CLng(p(0))
				dd = CLng(p(1))
				yy = CLng(p(2))

				If yy < 100 Then yy = 2000 + yy   ' adjust 2-digit year

				dt = DateSerial(yy, mm, dd)

				gPostDate = Right("0" & Month(dt), 2) & "/" & _
							Right("0" & Day(dt), 2) & "/" & _
							Year(dt)
			Else
				gPostDate = rawDate   ' fallback if unexpected format
			End If

			gGotHeaderInfo = True

        End If
    End If

    Select Case gStage
        Case 0
            If IsRecordStart(srec) Then
                gL1 = srec
                gStage = 1
            End If

        Case 1
            If IsLine2(srec) Then
                gL2 = srec
                gStage = 2
            ElseIf IsRecordStart(srec) Then
                gL1 = srec
                gStage = 1
            Else
                gStage = 0
            End If

        Case 2
            If IsLine3(srec) Then
                Dim outLine
                outLine = PadRight(gL1, PAD_LEN) & _
                          PadRight(gL2, PAD_LEN) & _
                          srec

                outLine = CleanHeaders(outLine)

                ' Prefix date, append report + file
                ' (spaces between tokens so downstream split is easy)
                outLine = gPostDate & " " & outLine & " " & gERTRFile

                PrintLine outLine
            End If
            gStage = 0
    End Select

End Sub

Sub CloseDoc()
End Sub

'-----------------------------
' Helpers
'-----------------------------
Function IsRecordStart(ByVal s)
    If Len(s) < 2 Then
        IsRecordStart = False
        Exit Function
    End If
    IsRecordStart = (IsDigit(Mid(s, 1, 1)) And IsDigit(Mid(s, 2, 1)))
End Function

Function IsLine2(ByVal s)
    IsLine2 = (UCase(Left(LTrim(s), 3)) = "RTN")
End Function

Function IsLine3(ByVal s)
    IsLine3 = (UCase(Left(LTrim(s), 12)) = "COMPANY NAME")
End Function

Function IsDigit(ByVal ch)
    IsDigit = (Len(ch) = 1 And Asc(ch) >= 48 And Asc(ch) <= 57)
End Function

Function PadRight(ByVal s, ByVal n)
    s = Nz(s)
    If Len(s) >= n Then
        PadRight = Left(s, n)
    Else
        PadRight = s & String(n - Len(s), " ")
    End If
End Function

Function CleanHeaders(ByVal s)
    ' Line 2 headers
    s = Replace(s, "RTN TRACE# = ", "")
    s = Replace(s, "RTN RDFI# = ", "")
    s = Replace(s, "DEATH DATE = ", "")
    s = Replace(s, "INFO = ", "")

    ' Line 3 headers
    s = Replace(s, "COMPANY NAME = ", "")
    s = Replace(s, "COMPANY ID = ", "")

    CleanHeaders = s
End Function

Function Nz(ByVal v)
    If IsNull(v) Or IsEmpty(v) Then
        Nz = ""
    Else
        Nz = CStr(v)
    End If
End Function

Sub PrintLine(ByVal s)
    fout.WriteLine s
End Sub
