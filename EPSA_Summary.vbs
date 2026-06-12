'============================================================
' byREQUEST Transform Script (Buffered Excel Write)
' Fixed-width line (srec) -> Excel (.xlsx)
'
' - Buffers rows and writes in batches
' - Lazy-init Excel in ProcessLine
' - Saves in CloseDoc
' - Output filename with timestamp + random suffix
' - NO HEADER ROW added (data starts at row 1)
' - Positions array = start of NEXT column (inclusive)
' - Forces ALL cells to TEXT format – no auto-number/date/scientific conversion
'============================================================
Option Explicit

'==========================================================
' EASY TO CHANGE:
' Output folder where the XLSX will be written
'==========================================================
Dim OUTPUT_DIR 

OUTPUT_DIR = GetOutputDir("S:\Finance_Department\Reports\IT_Reports\EPSA\") 
 

 
Dim gFso
Dim gPositions, gStarts, gLengths, gColCount
Dim gXlApp, gWb, gWs
Dim gOutFile
Dim gExcelReady
Dim gRowExcel
Dim gBufSize, gBufCount
Dim gBuf

Sub StartDoc()
    On Error Resume Next

    Set gFso = CreateObject("Scripting.FileSystemObject")

    '-----------------------------
    ' Column boundaries (fixed positions)
    '-----------------------------
    gPositions = Array(11,23,35,51,76,92,100,108,116,128,140,156,168,183,208,218,243,247,267,297,305,315,330,344,356,368,372,376,386,400,414,449)
    gColCount = UBound(gPositions) + 1

    ReDim gStarts(gColCount - 1)
    ReDim gLengths(gColCount - 1)

    gStarts(0) = 1

    Dim i
    For i = 1 To gColCount - 1
        gStarts(i) = gPositions(i - 1)
    Next

    For i = 0 To gColCount - 2
        gLengths(i) = gStarts(i + 1) - gStarts(i)
    Next

    gLengths(gColCount - 1) = gPositions(UBound(gPositions)) - gStarts(gColCount - 1)

    '-----------------------------
    ' Build output file name from Spoolfile.Name
    '-----------------------------
    Dim folderPath, baseName, stamp, rndNum

    folderPath = OUTPUT_DIR
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    If Not gFso.FolderExists(folderPath) Then
        'Hard fail if folder is wrong (better than silently going nowhere)
        WScript.Quit 1
    End If

    baseName = SpoolFile.name
	gOutFile = GetUniqueOutputPath(folderPath, baseName)

    '-----------------------------
    ' Buffer init
    '-----------------------------
    gBufSize = 2000
    gBufCount = 0
    ReDim gBuf(gBufSize - 1, gColCount - 1)

    gExcelReady = False
    gRowExcel = 1

    On Error GoTo 0
End Sub



Sub ProcessLine()
    On Error Resume Next
    
    If gColCount = 0 Then StartDoc
    If Len(Trim(srec)) = 0 Then Exit Sub
    
    EnsureExcelReady
    If gWs Is Nothing Then Exit Sub
    
    Dim col, seg
    For col = 0 To gColCount - 1
        seg = ""
        If Len(srec) >= gStarts(col) Then
            seg = Mid(srec, gStarts(col), gLengths(col))
        End If
        gBuf(gBufCount, col) = Trim(seg)
    Next
    
    gBufCount = gBufCount + 1
    If gBufCount >= gBufSize Then
        FlushBuffer
    End If
    
    On Error GoTo 0
End Sub

Sub EnsureExcelReady()
    On Error Resume Next
    
    If gExcelReady = True Then Exit Sub
    
    Set gXlApp = CreateObject("Excel.Application")
    If Err.Number <> 0 Or gXlApp Is Nothing Then
        Err.Clear
        Exit Sub
    End If
    
    gXlApp.Visible       = False
    gXlApp.ScreenUpdating = False
    gXlApp.EnableEvents   = False
    gXlApp.DisplayAlerts  = False
    gXlApp.Calculation    = -4135   ' xlCalculationManual
    
    Set gWb = gXlApp.Workbooks.Add
    Set gWs = gWb.Sheets(1)
    gWs.Name = "Data"
    
    ' Force entire sheet to TEXT format before any data is written
    gWs.Cells.NumberFormat = "@"
    
    gExcelReady = True
    
    On Error GoTo 0
End Sub

Sub FlushBuffer()
    On Error Resume Next
    
    If gBufCount <= 0 Then Exit Sub
    If gWs Is Nothing Then Exit Sub
    
    gWs.Cells(gRowExcel, 1).Resize(gBufCount, gColCount).Value = gBuf
    
    gRowExcel = gRowExcel + gBufCount
    gBufCount = 0
    
    On Error GoTo 0
End Sub


'Excel constants (late-binding safe)
Const xlUp = -4162
Const xlDelimited = 1
Const xlGeneralFormat = 1


Sub FormatColumnMNumeric(ByVal wb)
    On Error Resume Next

    Dim ws, lastRow, rng, data, i, s

    If wb Is Nothing Then Exit Sub
    Set ws = wb.Worksheets(1)
    If ws Is Nothing Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, "M").End(xlUp).Row
    If lastRow < 2 Then Exit Sub   'row 1 header

    Set rng = ws.Range("M2:M" & lastRow)

    'Read column into array (fast)
    data = rng.Value

    For i = 1 To UBound(data, 1)
        If Not IsEmpty(data(i, 1)) And Not IsNull(data(i, 1)) Then

            'Convert everything through a cleaned string path
            s = CStr(data(i, 1))

            'Clean hidden / common junk
            s = Replace(s, Chr(160), "")   'NBSP
            s = Replace(s, vbTab, "")
            s = Replace(s, " ", "")
            s = Replace(s, ",", "")        'remove thousand separators
            s = Replace(s, "$", "")
            s = Replace(s, "'", "")

            'Handle (123.45) negative format
            If Left(s, 1) = "(" And Right(s, 1) = ")" Then
                s = "-" & Mid(s, 2, Len(s) - 2)
            End If

            If Len(s) > 0 And IsNumeric(s) Then
                data(i, 1) = CDbl(s)       'force real numeric
            End If
        End If
    Next

    'Write back and format
    rng.Value = data
    rng.NumberFormat = "#,##0.00"

    On Error GoTo 0
End Sub



Sub CloseDoc()
    On Error Resume Next

    '-----------------------------------
    ' Basic guards
    '-----------------------------------
    If gWs Is Nothing Then Exit Sub
    If gWb Is Nothing Then Exit Sub
    If gXlApp Is Nothing Then Exit Sub

    '-----------------------------------
    ' Flush to Excel
    '-----------------------------------
    FlushBuffer
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If

    '-----------------------------------
    ' Format as text
    '-----------------------------------
    gWs.UsedRange.NumberFormat = "@"
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If

    '-----------------------------------
    ' Save output workbook
    '-----------------------------------
    gWb.SaveAs gOutFile, 51   ' 51 = xlsx
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If

    '-----------------------------------
    ' Macro Section
    '-----------------------------------
    Dim macroWorkbookPath
    Dim wbMacro
    Dim macroCall

    macroWorkbookPath = "S:\CarolinaTrust\byREQUEST\byREQUEST.xls"

    If Not gFso.FileExists(macroWorkbookPath) Then Exit Sub

    Set wbMacro = gXlApp.Workbooks.Open(macroWorkbookPath)
    If Err.Number <> 0 Or wbMacro Is Nothing Then
        Err.Clear
        Exit Sub
    End If

    '-----------------------------------
    ' FORCE: Output workbook + FIRST sheet active
    '-----------------------------------

	gWb.Activate
	gWb.Worksheets(1).Activate
	gWb.Worksheets(1).Range("A1").Select

	'NEW: Column M to numeric BEFORE pivot macro
	FormatColumnMNumeric gWb

	'Run macro
	macroCall = "'" & wbMacro.Name & "'!Module27.EPSA_ACH_Combined"
'	gXlApp.Run macroCall

    If Err.Number <> 0 Then
        Err.Clear
    End If

    '-----------------------------------
    ' Save AFTER macro (important)
    '-----------------------------------
    gWb.Save

    '-----------------------------------
    ' Cleanup
    '-----------------------------------
    wbMacro.Close False
    gWb.Close False
    gXlApp.Quit

    Set wbMacro = Nothing
    Set gWs = Nothing
    Set gWb = Nothing
    Set gXlApp = Nothing
    Set gFso = Nothing

    On Error GoTo 0
End Sub


Function GetUniqueOutputPath(ByVal folderPath, ByVal baseName)
    On Error Resume Next
    
    Dim candidate, highestN, file, nm, n, posTilde, posDot
    
    highestN = 0
    
    'Check plain name first
    candidate = folderPath & baseName & ".xlsx"
    
    If Not gFso.FileExists(candidate) Then
        GetUniqueOutputPath = candidate
        Exit Function
    End If
    
    'Scan folder for ~N versions
    Dim f
    For Each f In gFso.GetFolder(folderPath).Files
        
        nm = f.Name
        
        'Must start with baseName and end with .xlsx
        If LCase(Left(nm, Len(baseName))) = LCase(baseName) Then
            
            posTilde = InStrRev(nm, "~")
            posDot = InStrRev(nm, ".")
            
            If posTilde > 0 And posDot > posTilde Then
                n = Mid(nm, posTilde + 1, posDot - posTilde - 1)
                
                If IsNumeric(n) Then
                    If CLng(n) > highestN Then
                        highestN = CLng(n)
                    End If
                End If
            End If
            
        End If
    Next
    
    'Next number
    highestN = highestN + 1
    
    GetUniqueOutputPath = folderPath & baseName & "~" & highestN & ".xlsx"
    
    On Error GoTo 0
End Function



Function GetOutputDir(ByVal baseDir)
    Dim fso, yr, mon, p1, p2

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Right(baseDir, 1) <> "\" Then baseDir = baseDir & "\"

    yr = CStr(Year(Date))
    mon = Right("0" & Month(Date), 2)

    p1 = baseDir & yr
    p2 = p1 & "\" & mon

    If Not fso.FolderExists(p1) Then fso.CreateFolder p1
    If Not fso.FolderExists(p2) Then fso.CreateFolder p2

    GetOutputDir = p2 & "\"

    Set fso = Nothing
End Function 