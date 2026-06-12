Sub EPSA_ACH_Combined()
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo CleanUp

    Dim wsRaw As Worksheet, wsData As Worksheet
    Dim lastCol As Long, col As Long
    Dim headerValue As String, cleanHeader As String

    ' Use the current sheet (where you pasted the ACH report) as raw data
    Set wsRaw = ActiveSheet
    Set wsData = wsRaw
    
    ' Move ERTR section to its own sheet BEFORE cleaning headers / pivots
    Create_ERTRSummary wsData


    ' -------------------------------------------------
    ' 1. Clean & standardize headers for ACH_EPSA table format
    ' -------------------------------------------------
    lastCol = wsData.Cells(1, Columns.Count).End(xlToLeft).column
    
    ' First, let's see what headers we actually have
    Debug.Print "=== ORIGINAL HEADERS ==="
    For col = 1 To lastCol
        Debug.Print "Col " & col & ": '" & wsData.Cells(1, col).Value & "'"
    Next col
    
    ' Clean headers - remove invalid characters for pivot tables
    For col = 1 To lastCol
        headerValue = CStr(wsData.Cells(1, col).Value)
        
        ' Clean the header - remove invalid characters
        cleanHeader = CleanPivotHeader(headerValue)
        
        ' Apply standardized naming
        cleanHeader = StandardizeHeaderName(cleanHeader)
        
        ' Ensure it's not empty and not too long
        If Len(cleanHeader) = 0 Then
            cleanHeader = "Column_" & col
        ElseIf Len(cleanHeader) > 255 Then
            cleanHeader = Left(cleanHeader, 255)
        End If
        
        ' Set the cleaned header
        wsData.Cells(1, col).Value = cleanHeader
        
        ' Debug: Show transformation
        Debug.Print "Col " & col & ": '" & headerValue & "' -> '" & cleanHeader & "'"
    Next col

    ' Debug: Show final headers
    Debug.Print "=== FINAL HEADERS ==="
    For col = 1 To lastCol
        Debug.Print "Col " & col & ": '" & wsData.Cells(1, col).Value & "'"
    Next col
    
    ' Check for blank cells in data range
    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).row
    If lastRow < 2 Then
        MsgBox "No data found! Please ensure you have pasted the ACH report data.", vbExclamation
        GoTo CleanUp
    End If
    
    ' Check for any blank headers
    For col = 1 To lastCol
        If Len(Trim(wsData.Cells(1, col).Value)) = 0 Then
            wsData.Cells(1, col).Value = "Column_" & col
        End If
    Next col

    ' -------------------------------------------------
    ' 2. Create pivot tables using database column names
    ' -------------------------------------------------
    CreatePivot_OverviewByCompany wsData
    CreatePivot_DebitsVsCredits wsData
    'CreatePivot_Top10MembersByAmount wsData
    CreatePivot_BatchSummary wsData
    CreatePivot_TransactionsByDate wsData
    'CreatePivot_TopIndividuals wsData
    CreatePivot_StatusSummary wsData
    CreatePivot_TranTypeSummary wsData
    'MoveERTRReportLeft 6

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    If Err.Number <> 0 Then
        MsgBox "Error: " & Err.Description & vbCrLf & "Error Number: " & Err.Number & vbCrLf & _
               "Check Immediate Window (Ctrl+G) for header debug info.", vbCritical
        Debug.Print "ERROR: " & Err.Description & " (Error #" & Err.Number & ")"
    Else
        ' Activate the first pivot tab
        On Error Resume Next
        ThisWorkbook.Worksheets("Overview by Company").Activate
        On Error GoTo 0
        'MsgBox "ACH Dashboard created successfully! 8 pivot tables created on separate tabs.", vbInformation
    End If
End Sub

' ================================================================
' Clean header for PivotTable compatibility
' ================================================================
Function CleanPivotHeader(headerText As String) As String
    Dim result As String
    result = Trim(headerText)
    
    ' Remove single quotes
    result = Replace(result, "'", "")
    
    ' Remove double quotes
    result = Replace(result, """", "")
    
    ' Remove special characters that cause issues
    result = Replace(result, "!", "")
    result = Replace(result, "@", "")
    result = Replace(result, "#", "")
    result = Replace(result, "$", "")
    result = Replace(result, "%", "")
    result = Replace(result, "^", "")
    result = Replace(result, "&", "")
    result = Replace(result, "*", "")
    result = Replace(result, "(", "")
    result = Replace(result, ")", "")
    result = Replace(result, "-", "")
    result = Replace(result, "=", "")
    result = Replace(result, "+", "")
    result = Replace(result, "[", "")
    result = Replace(result, "]", "")
    result = Replace(result, "{", "")
    result = Replace(result, "}", "")
    result = Replace(result, "\", "")
    result = Replace(result, "|", "")
    result = Replace(result, ";", "")
    result = Replace(result, ":", "")
    result = Replace(result, "'", "")
    result = Replace(result, """", "")
    result = Replace(result, "<", "")
    result = Replace(result, ">", "")
    result = Replace(result, "?", "")
    result = Replace(result, "/", "")
    result = Replace(result, "~", "")
    result = Replace(result, "`", "")
    
    ' Replace multiple spaces with single space
    While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Wend
    
    ' Replace spaces with underscores for Excel compatibility
    result = Replace(result, " ", "_")
    
    ' Remove leading/trailing underscores
    result = Trim(result)
    If Left(result, 1) = "_" Then result = Mid(result, 2)
    If Right(result, 1) = "_" Then result = Left(result, Len(result) - 1)
    
    ' Ensure first character is a letter
    If Len(result) > 0 Then
        If Not (Asc(UCase(Left(result, 1))) >= 65 And Asc(UCase(Left(result, 1))) <= 90) Then
            result = "Col_" & result
        End If
    End If
    
    CleanPivotHeader = result
End Function

' ================================================================
' Standardize header names to known database column names
' ================================================================
Function StandardizeHeaderName(headerText As String) As String
    Dim result As String
    result = UCase(Trim(headerText))
    
    ' Map common variations to standard names
    Select Case result
        ' Core fields
        Case "REPORT", "RPT"
            StandardizeHeaderName = "Report"
        Case "POST_DATE", "POSTDATE", "DATE", "TRANS_DATE"
            StandardizeHeaderName = "Post_Date"
        Case "POST_TIME", "POSTTIME", "TIME"
            StandardizeHeaderName = "Post_Time"
        Case "SETTLEMENT_DATE", "SETTLEMENTDATE", "SETTL_DATE"
            StandardizeHeaderName = "Settlement_Date"
        Case "COMPANY_NAME", "COMPANYNAME", "CO_NAME", "SOURCE_NAME"
            StandardizeHeaderName = "Company_Name"
        Case "ACH_BATCH_ID", "ACHBATCHID", "BATCH_ID", "BATCH"
            StandardizeHeaderName = "ACH_Batch_ID"
        Case "PCTL", "PCTL_"
            StandardizeHeaderName = "PCTL"
        Case "LCTL", "LCTL_"
            StandardizeHeaderName = "LCTL"
        Case "STATUS", "ST"
            StandardizeHeaderName = "Status"
        Case "TRAN_TYPE", "TRANTYPE", "TYPE_"
            StandardizeHeaderName = "Tran_Type"
        Case "TRAN_CODE", "TRANCODE", "TC", "TRANS_CODE"
            StandardizeHeaderName = "Tran_Code"
        Case "EFFECTIVE_DATE", "EFFECTIVEDATE", "EFF_DATE"
            StandardizeHeaderName = "Effective_Date"
        Case "AMOUNT", "AMT"
            StandardizeHeaderName = "Amount"
        Case "MEMBER_NBR", "MEMBERNBR", "MEMBER", "MEMBER_", "MEM_NBR"
            StandardizeHeaderName = "Member_Nbr"
        Case "ACCT", "ACCOUNT", "ACCOUNT_"
            StandardizeHeaderName = "Acct"
        Case "MBR_NAME", "MBRNAME", "NAME", "MEMBER_NAME"
            StandardizeHeaderName = "Mbr_Name"
        Case "FEE", "FEE_AMT"
            StandardizeHeaderName = "Fee"
        Case "ACH_IND_NAME", "ACHINDNAME", "INDIVIDUAL_NAME"
            StandardizeHeaderName = "ACH_Ind_Name"
        Case "DD"
            StandardizeHeaderName = "DD"
        Case "TRACE_NBR", "TRACENBR", "TRACE", "TRACE_"
            StandardizeHeaderName = "Trace_Nbr"
        Case "DESCRIPTION", "DESC"
            StandardizeHeaderName = "Description"
        Case "TYPE" ' Separate from Tran_Type
            StandardizeHeaderName = "Type"
        Case "POST_SEQ", "POSTSEQ"
            StandardizeHeaderName = "Post_Seq"
        Case "FILE_POST_SEQ", "FILEPOSTSEQ"
            StandardizeHeaderName = "File_Post_Seq"
        Case "DATECOL", "DATE_COL"
            StandardizeHeaderName = "DateCol"
            
        ' Additional fields
        Case "SOURCE_NAME"
            StandardizeHeaderName = "Source_Name"
        Case "BATCH"
            StandardizeHeaderName = "Batch"
        Case "CHECK_NUMBER", "CHECKNUMBER", "CHECK"
            StandardizeHeaderName = "Check_Number"
        Case "SDC_NUMBER", "SDCNUMBER", "SDC"
            StandardizeHeaderName = "SDC_Number"
        Case "S_L", "S/L"
            StandardizeHeaderName = "S_L"
        Case "RR"
            StandardizeHeaderName = "RR"
        Case "RDFI_ID", "RDFIID", "RDFI"
            StandardizeHeaderName = "RDFI_ID"
        Case "FILE_ACC_NUMBER", "FILEACCNUMBER", "ACC_NUMBER"
            StandardizeHeaderName = "File_Acc_Number"
        Case "INDIVIDUAL_ID", "INDIVIDUALID", "INDIV_ID"
            StandardizeHeaderName = "Individual_ID"
        Case "MESSAGE", "MSG"
            StandardizeHeaderName = "Message"
        Case "SOURCE_MESSAGE", "SOURCEMESSAGE"
            StandardizeHeaderName = "Source_Message"
        Case "HASHKEY", "HASH_KEY"
            StandardizeHeaderName = "HashKey"
        Case Else
            ' Return cleaned text as-is
            StandardizeHeaderName = headerText
    End Select
End Function

' ================================================================
' Reusable pivot creator with error handling
' ================================================================
Function CreateACHPivotOnNewTab(wsData As Worksheet, tabName As String, pivotName As String) As PivotTable
    Dim wsPivot As Worksheet
    Dim pc As PivotCache, pt As PivotTable, rng As Range
    Dim lr As Long, lc As Long
    
    On Error GoTo ErrorHandler
    
    ' Delete existing sheet if it exists
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(tabName).Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler

    ' Create new worksheet
    Set wsPivot = Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsPivot.Name = tabName

    ' Define data range - ensure we have headers
    lr = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).row
    lc = wsData.Cells(1, Columns.Count).End(xlToLeft).column
    
    ' Verify we have data
    If lr < 2 Then
        Err.Raise 1001, , "No data found for pivot table"
    End If
    
    If lc < 1 Then
        Err.Raise 1002, , "No columns found for pivot table"
    End If
    
    Set rng = wsData.Range(wsData.Cells(1, 1), wsData.Cells(lr, lc))

    ' Create pivot cache and table
    Set pc = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rng)
    Set pt = wsPivot.PivotTables.Add(PivotCache:=pc, TableDestination:=wsPivot.Range("A3"), TableName:=pivotName)
    
    Set CreateACHPivotOnNewTab = pt
    Exit Function
    
ErrorHandler:
    Debug.Print "Error creating pivot '" & pivotName & "': " & Err.Description
    Set CreateACHPivotOnNewTab = Nothing
End Function

' ================================================================
' Safe pivot field adder with error handling
' ================================================================
Sub SafeAddPivotField(pt As PivotTable, fieldName As String, orientation As XlPivotFieldOrientation)
    On Error Resume Next
    Dim pf As PivotField
    Set pf = pt.PivotFields(fieldName)
    If Err.Number = 0 Then
        pf.orientation = orientation
        pf.Position = 1
    Else
        Debug.Print "Field not found for pivot: " & fieldName
    End If
    Err.Clear
End Sub

' ================================================================
' Safe data field adder with error handling
' ================================================================
Sub SafeAddDataField(pt As PivotTable, fieldName As String, caption As String, functionType As XlConsolidationFunction)
    On Error Resume Next
    Dim pf As PivotField
    Set pf = pt.PivotFields(fieldName)
    If Err.Number = 0 Then
        With pt
            .AddDataField pf, caption, functionType
            .DataFields(.DataFields.Count).NumberFormat = "$#,##0.00"
        End With
    Else
        Debug.Print "Data field not found: " & fieldName
    End If
    Err.Clear
End Sub

' ================================================================
' 1. Overview by Company_Name
' ================================================================
Private Sub CreatePivot_OverviewByCompany(wsData As Worksheet)
    Dim pt As PivotTable
    Set pt = CreateACHPivotOnNewTab(wsData, "Overview by Company", "PivotByCompany")
    
    If pt Is Nothing Then Exit Sub
    
    With pt
        SafeAddPivotField pt, "Company_Name", xlRowField
        SafeAddPivotField pt, "Tran_Code", xlColumnField
        SafeAddDataField pt, "Amount", "Total Amount", xlSum
        
        ' Try to add count as well
        On Error Resume Next
        .AddDataField .PivotFields("Amount"), "Count", xlCount
        .DataFields(2).NumberFormat = "#,##0"
        Err.Clear
        
        .RowGrand = True
        .ColumnGrand = True
    End With
    
    ' Format the sheet
    With pt.Parent
        .Range("A1").Value = "ACH Overview by Company Name"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Columns.AutoFit
    End With
End Sub

' ================================================================
' 2. Debits vs Credits Summary by Tran_Code
' ================================================================
Private Sub CreatePivot_DebitsVsCredits(wsData As Worksheet)
    Dim pt As PivotTable
    Set pt = CreateACHPivotOnNewTab(wsData, "Debits vs Credits", "DebitsCredits")
    
    If pt Is Nothing Then Exit Sub
    
    With pt
        SafeAddPivotField pt, "Tran_Code", xlRowField
        SafeAddPivotField pt, "Tran_Type", xlRowField
        SafeAddDataField pt, "Amount", "Total Amount", xlSum
        
        ' Try to add count
        On Error Resume Next
        .AddDataField .PivotFields("Amount"), "Count", xlCount
        .DataFields(2).NumberFormat = "#,##0"
        Err.Clear
        
        .RowGrand = True
    End With
    
    ' Format the sheet
    With pt.Parent
        .Range("A1").Value = "ACH Debits vs Credits Summary"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Columns.AutoFit
    End With
End Sub

' ================================================================
' 3. Top Members by Dollar Amount (simplified)
' ================================================================
Private Sub CreatePivot_Top10MembersByAmount(wsData As Worksheet)
    Dim pt As PivotTable
    Set pt = CreateACHPivotOnNewTab(wsData, "Top Members", "TopMembers")
    
    If pt Is Nothing Then Exit Sub
    
    With pt
        ' Add Member_Nbr as row field
        On Error Resume Next
        .PivotFields("Member_Nbr").orientation = xlRowField
        
        ' Add Tran_Code as column field
        .PivotFields("Tran_Code").orientation = xlColumnField
        Err.Clear
        
        ' Add data fields
        .AddDataField .PivotFields("Amount"), "Total Amount", xlSum
        .DataFields(1).NumberFormat = "$#,##0.00"
        
        ' Add Count of Tran_Code

        
        .RowGrand = True
        .ColumnGrand = True
    End With
    
    ' Format the sheet
    With pt.Parent
        .Range("A1").Value = "Top Members by Total Amount"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Columns.AutoFit
    End With
End Sub

' ================================================================
' 4. Batch Summary (simplified)
' ================================================================
Private Sub CreatePivot_BatchSummary(wsData As Worksheet)
    Dim pt As PivotTable
    Set pt = CreateACHPivotOnNewTab(wsData, "Batch Summary", "BatchSummary")
    
    If pt Is Nothing Then Exit Sub
    
    With pt
        ' Try ACH_Batch_ID first, then Batch
        On Error Resume Next
        .PivotFields("ACH_Batch_ID").orientation = xlRowField
                ' Add Tran_Code as column field
        .PivotFields("Tran_Code").orientation = xlColumnField

        .AddDataField .PivotFields("Tran_Code"), "Count", xlCount
     
        If Err.Number <> 0 Then
            Err.Clear
            .PivotFields("Batch").orientation = xlRowField
        End If
        On Error GoTo 0
        
        SafeAddDataField pt, "Amount", "Total Amount", xlSum
        
        .RowGrand = True
    End With
    
    ' Format the sheet
    With pt.Parent
        .Range("A1").Value = "ACH Batch Summary"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Columns.AutoFit
    End With
End Sub

' ================================================================
' 5. Transactions by Date (simplified)
' ================================================================
Private Sub CreatePivot_TransactionsByDate(wsData As Worksheet)
    Dim pt As PivotTable
    Set pt = CreateACHPivotOnNewTab(wsData, "By Date", "ByDate")
    
    If pt Is Nothing Then Exit Sub
    
    With pt
        ' Try Post_Date first, then DateCol
        On Error Resume Next
        .PivotFields("Post_Date").orientation = xlRowField
                ' Add Tran_Code as column field
        .PivotFields("Tran_Code").orientation = xlColumnField

        .AddDataField .PivotFields("Tran_Code"), "Count", xlCount
        
        
        
        If Err.Number <> 0 Then
            Err.Clear
            .PivotFields("DateCol").orientation = xlRowField
        End If
        On Error GoTo 0
        
        SafeAddDataField pt, "Amount", "Total Amount", xlSum
        
        .RowGrand = True
    End With
    
    ' Format the sheet
    With pt.Parent
        .Range("A1").Value = "ACH Transactions by Date"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Columns.AutoFit
    End With
End Sub

' ================================================================
' 6. Create simplified pivot tables based on available fields
' ================================================================
Private Sub CreatePivot_TopIndividuals(wsData As Worksheet)
    Dim pt As PivotTable
    Set pt = CreateACHPivotOnNewTab(wsData, "Individuals", "Individuals")
    
    If pt Is Nothing Then Exit Sub
    
    With pt
        ' Try ACH_Ind_Name
        On Error Resume Next
        .PivotFields("ACH_Ind_Name").orientation = xlRowField
                ' Add Tran_Code as column field
        .PivotFields("Tran_Code").orientation = xlColumnField

        .AddDataField .PivotFields("Tran_Code"), "Count", xlCount
        .DataFields(2).NumberFormat = "#,##0"
        
        
        If Err.Number <> 0 Then
            Err.Clear
            ' Try Individual_Name or other variations
            .PivotFields("Individual_Name").orientation = xlRowField
        End If
        On Error GoTo 0
        
        SafeAddDataField pt, "Amount", "Total Amount", xlSum
        
        .RowGrand = True
    End With
    
    ' Format the sheet
    With pt.Parent
        .Range("A1").Value = "ACH Transactions by Individual"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Columns.AutoFit
    End With
End Sub

' ================================================================
' 7. Status Summary (simplified)
' ================================================================
Private Sub CreatePivot_StatusSummary(wsData As Worksheet)
    Dim pt As PivotTable
    Set pt = CreateACHPivotOnNewTab(wsData, "Status", "Status")
    
    If pt Is Nothing Then Exit Sub
    
    With pt
        ' Rows
        SafeAddPivotField pt, "Status", xlRowField
        SafeAddPivotField pt, "Tran_Code", xlRowField
        SafeAddPivotField pt, "Type", xlRowField
        
        ' Columns
        SafeAddPivotField pt, "Tran_Type", xlColumnField
        
        ' Values
        SafeAddDataField pt, "Amount", "Total Amount", xlSum
        .AddDataField .PivotFields("Tran_Code"), "Count", xlCount
        
        .RowGrand = True
    End With
    
    ' Format the sheet
    With pt.Parent
        .Range("A1").Value = "Transaction Status Summary"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Columns.AutoFit
    End With
End Sub
' ================================================================
' 8. Transaction Type Summary (enhanced with field checking)
' ================================================================
Private Sub CreatePivot_TranTypeSummary(wsData As Worksheet)
    Dim pt As PivotTable
    Set pt = CreateACHPivotOnNewTab(wsData, "Tran Type Summary", "TranTypeSummary")
    
    If pt Is Nothing Then Exit Sub
    
    ' Check which fields are available
    Dim hasTranType As Boolean, hasFilePostSeq As Boolean
    Dim hasACHBatchID As Boolean, hasType As Boolean, hasAmount As Boolean
    
    On Error Resume Next
    hasTranType = (pt.PivotFields("Tran_Type").Name = "Tran_Type")
    hasFilePostSeq = (pt.PivotFields("File_Post_Seq").Name = "File_Post_Seq")
    hasACHBatchID = (pt.PivotFields("ACH_Batch_ID").Name = "ACH_Batch_ID")
    hasType = (pt.PivotFields("Type").Name = "Type")
    hasAmount = (pt.PivotFields("Amount").Name = "Amount")
    On Error GoTo 0
    
    With pt
        ' Add available fields
        If hasTranType Then
            .PivotFields("Tran_Type").orientation = xlRowField
            .PivotFields("Tran_Type").Position = 1
        End If
        
        If hasFilePostSeq Then
            .PivotFields("File_Post_Seq").orientation = xlRowField
            .PivotFields("File_Post_Seq").Position = 2
        End If
        
        If hasACHBatchID Then
            .PivotFields("ACH_Batch_ID").orientation = xlRowField
            .PivotFields("ACH_Batch_ID").Position = 3
        End If
        
        If hasType Then
            .PivotFields("Type").orientation = xlColumnField
        End If
        
        ' Add data fields
        If hasAmount Then
            ' Sum of Amount
            .AddDataField .PivotFields("Amount"), "Total Amount", xlSum
            .DataFields(1).NumberFormat = "$#,##0.00"
            
            ' Count of Transactions (using Amount field)
            .AddDataField .PivotFields("Amount"), "Count of Transactions", xlCount
            .DataFields(2).NumberFormat = "#,##0"
        End If
        
        ' Try to add Count of ACH_Batch_ID (if field exists)
        If hasACHBatchID Then
            On Error Resume Next
            .AddDataField .PivotFields("ACH_Batch_ID"), "Count of Batches", xlCount
            If Err.Number = 0 Then
                .DataFields(.DataFields.Count).NumberFormat = "#,##0"
            End If
            Err.Clear
        End If
        
        .RowGrand = True
        .ColumnGrand = True
        
        ' Format layout
        .RowAxisLayout xlTabularRow
        .RepeatAllLabels xlRepeatLabels
    End With
    
    ' Format the sheet
    With pt.Parent
        .Range("A1").Value = "Transaction Type Summary"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        
        ' Add field availability info
        Dim infoText As String
        infoText = "Fields: "
        If hasTranType Then infoText = infoText & "Tran_Type ? "
        If hasFilePostSeq Then infoText = infoText & "File_Post_Seq ? "
        If hasACHBatchID Then infoText = infoText & "ACH_Batch_ID ? "
        If hasType Then infoText = infoText & "Type ? "
        If hasAmount Then infoText = infoText & "Amount ? "
        
        .Range("A2").Value = infoText
        .Range("A2").Font.Size = 9
        .Range("A2").Font.Italic = True
        
        .Columns.AutoFit
    End With
End Sub

Private Sub Create_ERTRSummary(wsData As Worksheet)

    Dim wsPivot As Worksheet
    Dim wsERTR As Worksheet
    Dim ertrHeaderRow As Long
    Dim lastRow As Long
    Dim rngCut As Range

    '------------------------------------------------------------
    ' 1) Create new worksheet (your preferred method)
    '------------------------------------------------------------
    Set wsPivot = Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsPivot.Name = "ERTR Report"

    '------------------------------------------------------------
    ' 2) Find ERTR start row in the CURRENT combined sheet
    '    - look for "Report" in column A after row 2
    '------------------------------------------------------------
    ertrHeaderRow = 0
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).row

    Dim r As Long
    For r = 3 To lastRow
        If InStr(1, CStr(wsData.Cells(r, 1).Value), "Report", vbTextCompare) > 0 Then
            ertrHeaderRow = r
            Exit For
        End If
    Next r

    If ertrHeaderRow = 0 Then
    
        Application.DisplayAlerts = False
        wsPivot.Range("A1").Value = "REPORT    ACH RECONCILIATION SUMMARY"
        wsPivot.Range("A1:C1").Merge
        Application.DisplayAlerts = True
        Exit Sub
    End If

    '------------------------------------------------------------
    ' 3) Cut ERTR block A:AG from header row to last row
    '    and paste into ERTR Report sheet at A1
    '------------------------------------------------------------
    Set rngCut = wsData.Range(wsData.Cells(ertrHeaderRow, 1), wsData.Cells(lastRow, 33)) ' A=1, AG=33
    rngCut.Cut Destination:=wsPivot.Range("A1")

    '------------------------------------------------------------
    ' 4) Delete placeholder columns on ERTR sheet
    '    Placeholders (from your ERTR row output):
    '    P,Q,R,U,V,W,X,AA,AB,AC,AD,AF
    '------------------------------------------------------------
    Application.DisplayAlerts = False
' Delete columns from highest to lowest index

    wsPivot.Columns(30).Delete ' AD
    wsPivot.Columns(27).Delete ' AA
    wsPivot.Columns(26).Delete ' Z
    wsPivot.Columns(25).Delete ' Y
    wsPivot.Columns(24).Delete ' X
    wsPivot.Columns(23).Delete ' W
    wsPivot.Columns(22).Delete ' V
    wsPivot.Columns(20).Delete ' T
    wsPivot.Columns(18).Delete ' R
    wsPivot.Columns(17).Delete ' Q
    wsPivot.Columns(16).Delete ' P
    wsPivot.Columns(14).Delete ' N
    wsPivot.Columns(12).Delete ' L
    wsPivot.Columns(8).Delete ' H
    wsPivot.Columns(7).Delete  ' G
    wsPivot.Range("A1").Value = "REPORT    ACH RECONCILIATION SUMMARY"
    Application.DisplayAlerts = True


    ' Auto-fit all columns in ERTR Report
    wsPivot.Columns.AutoFit
    
    
    
    'CREATE ERTR PIVOT
    
    Dim ertrRaw As Worksheet, ertrData As Worksheet
    Set ertrRaw = ActiveSheet
    Set ertrData = ertrRaw

    CreatePivot_ERTRStatusSummary ertrData

End Sub



Private Function FindERTRHeaderRow(ws As Worksheet) As Long
    Dim r As Long
    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    For r = 1 To lastRow
        If UCase(Trim(CStr(ws.Cells(r, 1).Value))) = "ERTR" _
           And UCase(Trim(CStr(ws.Cells(r, 2).Value))) = "POST_DATE" Then
            FindERTRHeaderRow = r
            Exit Function
        End If
    Next r

    FindERTRHeaderRow = 0
End Function


Private Function BuildERTRDataSheet(wsData As Worksheet) As Worksheet

    Dim wsERTR As Worksheet
    Dim ertrHeaderRow As Long
    Dim lastRow As Long
    Dim rngCopy As Range

    ertrHeaderRow = FindERTRHeaderRow(wsData)
    If ertrHeaderRow = 0 Then Exit Function

    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).row

    Set wsERTR = Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsERTR.Name = "ERTR Returns"

    ' Copy the whole ERTR block starting from the real header row
    Set rngCopy = wsData.Range(wsData.Cells(ertrHeaderRow, 1), wsData.Cells(lastRow, 33)) ' A:AG
    rngCopy.Copy Destination:=wsERTR.Range("A1")

    Application.DisplayAlerts = False

    wsERTR.Columns(1).Delete

    Application.DisplayAlerts = True

    ' Convert EPSA_Amount to numeric if needed
    Call NormalizeERTRAmount(wsERTR)

    wsERTR.Columns.AutoFit

    Set BuildERTRDataSheet = wsERTR
End Function

Private Sub NormalizeERTRAmount(ws As Worksheet)

    Dim amtCol As Long
    Dim lastRow As Long
    Dim r As Long
    Dim v As String

    amtCol = HeaderColumn(ws, "EPSA_Amount")
    If amtCol = 0 Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, amtCol).End(xlUp).row

    For r = 2 To lastRow
        v = Trim(CStr(ws.Cells(r, amtCol).Value))
        If v <> "" Then
            v = Replace(v, ",", "")
            If IsNumeric(v) Then
                ws.Cells(r, amtCol).Value = CDbl(v)
            End If
        End If
    Next r

    ws.Columns(amtCol).NumberFormat = "#,##0.00"
End Sub

Private Function HeaderColumn(ws As Worksheet, headerName As String) As Long
    Dim lastCol As Long
    Dim c As Long

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).column

    For c = 1 To lastCol
        If UCase(Trim(CStr(ws.Cells(1, c).Value))) = UCase(headerName) Then
            HeaderColumn = c
            Exit Function
        End If
    Next c

    HeaderColumn = 0
End Function

Private Sub CreatePivot_ERTRStatusSummary(wsData As Worksheet)

    Dim wsERTR As Worksheet
    Dim pt As PivotTable

    Set wsERTR = BuildERTRDataSheet(wsData)
    If wsERTR Is Nothing Then Exit Sub

    Set pt = CreateACHPivotOnNewTab(wsERTR, "ERTR Status", "ERTR_Status")
    If pt Is Nothing Then Exit Sub

    With pt
        ' Rows
        SafeAddPivotField pt, "Status", xlRowField
        SafeAddPivotField pt, "Tran_Code", xlRowField
        SafeAddPivotField pt, "Type", xlRowField

        ' Columns
        SafeAddPivotField pt, "Tran_Type", xlColumnField

        ' Values
        SafeAddDataField pt, "EPSA_Amount", "Total Amount", xlSum
        .AddDataField .PivotFields("Tran_Code"), "Count", xlCount

        .RowGrand = True
    End With

    With pt.Parent
        .Range("A1").Value = "ERTR Transaction Status Summary"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Columns.AutoFit
    End With
End Sub
