Option Explicit

Dim ExcelApp, wb, ws
Dim basePath, location, locations
Dim reportYear, currentMonth, i

Sub StartDoc()
    reportYear = Year(Now)
    currentMonth = Month(Now)
    basePath = "R:\" ' Root folder for reports

    locations = Array( _
        "Carthage", "Daingerfield", "Diana", "Gilmer", "Hallsville", "Henderson", "Hughes Springs", _
        "Jefferson", "Kilgore", "LeTourneau", "Longview", "Longview Loop", "Lufkin", "Marshall", "Marshall North", _
        "Mt Pleasant North", "Nacogdoches", "Spring Hill", "Tyler", "Wildwood" _
    )

    Set ExcelApp = CreateObject("Excel.Application")
    ExcelApp.Visible = False
    ExcelApp.DisplayAlerts = False
	



    For i = 0 To UBound(locations)
        location = locations(i)
        Dim filePath, wbName, mnthName
        filePath = basePath & reportYear & " FSR Production Report- " & location & ".xlsx"
		
		
        mnthName = MonthName(currentMonth)
		'mnthName ="March"
		
        ' Open or reuse workbook
	
		
		Dim fso
		Set fso = CreateObject("Scripting.FileSystemObject")
        Set wb = IsWorkbookOpen(ExcelApp, filePath)
		
		If wb Is Nothing Then
			If fso.FileExists(filePath) Then
				On Error Resume Next
				Set wb = ExcelApp.Workbooks.Open(filePath)
				If Err.Number <> 0 Then
					fout.Writeline "Error opening workbook: " & Err.Description
					Err.Clear
					Set wb = Nothing
				End If
				On Error GoTo 0
			Else
				fout.Writeline "File not found: " & filePath
			End If
		End If

		
        If Not wb Is Nothing Then
			fout.writeline  "Using Master File: "  & filePath
			

            For Each ws In wb.Sheets
                ' Parse sheet name to get empName and optionally empID
                Dim sheetName, empName, empID, sepIndex
                sheetName = ws.Name
                sepIndex = InStr(sheetName, " - ")
                If sepIndex > 0 Then
                    empName = Trim(Left(sheetName, sepIndex - 1))
                    empID = Trim(Mid(sheetName, sepIndex + 3))
                Else
                    empName = sheetName
                    empID = ""
                End If
				mnthName = "March"
				'MsgBox currentMonth & " " &  mnthName & " " &  empName & " " &  empID
				
				
                ' Open the source file: Total Accounts Assisted by User Report - <location>.xlsx
                Dim sourcePath, sourceFile, sourceWb, sourceWs
                sourcePath = "R:\byrequest\reports\Lobby Tracking\" & mnthName & "\"
                sourceFile = "Total Accounts Assisted by User Report - " & location & "Loop.xlsx"
				'MsgBox "Using Source File: "  & sourceFile
                On Error Resume Next
                Set sourceWb = ExcelApp.Workbooks.Open(sourcePath & sourceFile)
                On Error GoTo 0
				'MsgBox "Source File Opened: " & sourcePath & sourceFile
				
                If Not sourceWb Is Nothing Then
                    Set sourceWs = sourceWb.Sheets(1)

                    Dim rowIdx, lastRow, sourceID, sourceName, assisted, provided, productsSold
                    lastRow = sourceWs.Cells(sourceWs.Rows.Count, 2).End(-4162).Row ' xlUp
		
                    For rowIdx = 10 To lastRow
                        sourceID = Trim(sourceWs.Cells(rowIdx, 2).Value)
                        sourceName = Trim(sourceWs.Cells(rowIdx, 3).Value)
						
						MsgBox sourceID & " " & sourceName

                        If (empID <> "" And empID = sourceID) Or (LCase(sourceName) = LCase(empName)) Then
                            If IsNumeric(sourceWs.Cells(rowIdx, 39).Value) And IsNumeric(sourceWs.Cells(rowIdx, 40).Value) Then
                                assisted = CLng(sourceWs.Cells(rowIdx, 39).Value)
                                provided = CLng(sourceWs.Cells(rowIdx, 40).Value)
                                productsSold = provided - assisted
                                ws.Cells(currentMonth + 1, 2).Value = productsSold
                            End If
                            Exit For
                        End If
                    Next

                    sourceWb.Close False
                    Set sourceWb = Nothing
                End If
            Next

            wb.Close True
            Set wb = Nothing
        End If
    Next

    ExcelApp.Quit
    Set ExcelApp = Nothing
End Sub

Function IsWorkbookOpen(xlApp, filePath)
    Dim wbName, i
    wbName = Mid(filePath, InStrRev(filePath, "\") + 1)
    For i = 1 To xlApp.Workbooks.Count
        If LCase(xlApp.Workbooks(i).Name) = LCase(wbName) Then
            Set IsWorkbookOpen = xlApp.Workbooks(i)
            Exit Function
        End If
    Next
    Set IsWorkbookOpen = Nothing
End Function
