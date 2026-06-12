'===========================================================
' DVC101A -> FILE 2 BUILDER (Flatten tiers + Enrich from Excel)
'
' Lookup: "XType_Ttls.xlsx"
'   Columns: Savings Type | Savings Class | Description | NCUA CODE | Count | Ttl_Balance
'
' Output format (File 2):
'   Run_Date, Share_Type, Description, Class/Code, Savings Class, NCUA CODE,
'   Min_to_Earn, Var_Rate_Mthd, Balance_To, Rate/Differential, First_Offered
'
' Notes:
' - Uses your working widths (DO NOT CHANGE)
' - Flattens continuation tier rows by inheriting last product header context
' - Skips repeating headers/noise lines
'===========================================================

Option Explicit

Dim reNum, Run_Date
Dim widths, i

' Your working widths (for File 3-style lines)
widths = Array(4, 7, 5, 6, 6, 4, 4, 13, 6, 11, 10, 7, 8, 11, 5, 5, 5, 5, 10)

' Lookup dictionary: key = Savings Type (string)
' value = Array(SavingsClass, Description, NCUACode)
Dim dictShareTypes

' --- Options ---
Dim OPT_FlagMissingLookup   ' If True: write placeholder text when lookup missing
Dim OPT_SkipIfNoLookup      ' If True: skip output row when lookup missing

OPT_FlagMissingLookup = True
OPT_SkipIfNoLookup    = False


' Current (last-seen) header context for carrying down to tier rows
Dim curTYPE, curCLASS, curCODE, curPLAN, curPOST_PER, curERN_PER, curDIV_CDE
Dim curMIN_TO_EARN, curVAR_RTHD, curROLLBACK, curGRACE_DEP, curGRACE_WITH, curRENEW, curDATE_OFFERED
Dim tierIndex

Public Sub StartDoc

    ' Load lookup first
    Set dictShareTypes = CreateObject("Scripting.Dictionary")
    LoadShareTypesLookup "S:\SCCCU\DVC101A\SXType_Ttls.xlsx", dictShareTypes

    ' Output header (File 2)
    fout.WriteLine "Run_Date       Share_Type     Description                   Class/Code     Savings Class                           NCUA CODE      Min_to_Earn    Var_Rate_Mthd  Balance_To     Rate/Differential   First_Offered"

    ' Read Run_Date from header (first few lines)
    Dim SrecLocal, p
    For i = 1 To 6
        SrecLocal = GetLine
        p = InStr(1, SrecLocal, "DVC101A", vbTextCompare)
        If p = 1 Then
            Run_Date = Trim(Mid(SrecLocal, p + 7, 9))
            If Len(Run_Date) = 6 Then
                Run_Date = Left(Run_Date, 2) & "-" & Mid(Run_Date, 3, 2) & "-" & Right(Run_Date, 2)
            End If
            Exit For
        End If
    Next

    Set reNum = CreateObject("VBScript.RegExp")
    reNum.Pattern = "^\d+$"

    ' Clear context
    curTYPE = "" : tierIndex = 0
End Sub

Sub ProcessLine

    '-----------------------------
    ' Skip noise / repeating headers
    '-----------------------------
    If Len(Trim(srec)) = 0 Then Exit Sub

    ' Form feed / pagination artifacts
    If InStr(1, srec, Chr(12), vbBinaryCompare) > 0 Then Exit Sub

    ' Known headers
    If InStr(1, srec, "DVC101A", vbTextCompare) > 0 Then Exit Sub
    If InStr(1, srec, "SANTA", vbTextCompare) > 0 Then Exit Sub
    If InStr(1, srec, "MIN BAL TO", vbTextCompare) > 0 Then Exit Sub
    If InStr(1, srec, "DIVIDEND", vbTextCompare) > 0 Then Exit Sub
    If InStr(1, srec, "TYPE", vbTextCompare) > 0 Then Exit Sub
    If InStr(1, srec, "TOTL RCD", vbTextCompare) > 0 Then Exit Sub
    If InStr(1, srec, "END-OF-REPORT", vbTextCompare) > 0 Then Exit Sub

    ' Separator lines
    If InStr(srec, "----") > 0 Then Exit Sub

    '-----------------------------
    ' Parse fixed-width fields
    '-----------------------------
    Dim vals
    vals = SplitFields(srec)

    ' Map fields by position (File 3 flattened schema)
    Dim vTYPE, vCLASS, vCODE, vPLAN, vPOST_PER, vERN_PER, vDIV_CDE
    Dim vMIN_TO_EARN, vVAR_RTHD, vBAL_LIMIT, vRATE_DIFF, vINDEX, vRED_DIV_RATE, vREDBAL_LIMIT
    Dim vROLLBACK, vGRACE_DEP, vGRACE_WITH, vRENEW, vDATE_OFFERED

    vTYPE         = vals(0)
    vCLASS        = vals(1)
    vCODE         = vals(2)
    vPLAN         = vals(3)
    vPOST_PER     = vals(4)
    vERN_PER      = vals(5)
    vDIV_CDE      = vals(6)
    vMIN_TO_EARN  = vals(7)
    vVAR_RTHD     = vals(8)
    vBAL_LIMIT    = vals(9)
    vRATE_DIFF    = vals(10)
    vINDEX        = vals(11)
    vRED_DIV_RATE = vals(12)
    vREDBAL_LIMIT = vals(13)
    vROLLBACK     = vals(14)
    vGRACE_DEP    = vals(15)
    vGRACE_WITH   = vals(16)
    vRENEW        = vals(17)
    vDATE_OFFERED = vals(18)

    ' Decide if this is a Product row or a continuation tier row
    Dim isProductRow, isTier
    isProductRow = (Len(vTYPE) > 0 And reNum.Test(vTYPE))
    isTier = (Not isProductRow) And (Len(vBAL_LIMIT) > 0 Or Len(vRATE_DIFF) > 0 Or Len(vRED_DIV_RATE) > 0 Or Len(vREDBAL_LIMIT) > 0)

    ' 1) Product row: store context + output as tier 1
    If isProductRow Then

        curTYPE         = vTYPE
        curCLASS        = vCLASS
        curCODE         = vCODE
        curPLAN         = vPLAN
        curPOST_PER     = vPOST_PER
        curERN_PER      = vERN_PER
        curDIV_CDE      = vDIV_CDE
        curMIN_TO_EARN  = vMIN_TO_EARN
        curVAR_RTHD     = vVAR_RTHD
        curROLLBACK     = vROLLBACK
        curGRACE_DEP    = vGRACE_DEP
        curGRACE_WITH   = vGRACE_WITH
        curRENEW        = vRENEW
        curDATE_OFFERED = vDATE_OFFERED

        tierIndex = 0
        tierIndex = tierIndex + 1
        If Len(vINDEX) = 0 Then vINDEX = CStr(tierIndex)

        ' Emit File 2 row (enriched)
        WriteFile2Row curTYPE, curCLASS, curCODE, curMIN_TO_EARN, curVAR_RTHD, vBAL_LIMIT, vRATE_DIFF, curDATE_OFFERED
        Exit Sub
    End If

    ' 2) Tier row: inherit context and output as next tier
    If isTier Then
        If Len(curTYPE) = 0 Then Exit Sub ' orphan tier

        tierIndex = tierIndex + 1
        If Len(vINDEX) = 0 Then vINDEX = CStr(tierIndex)

        WriteFile2Row curTYPE, curCLASS, curCODE, curMIN_TO_EARN, curVAR_RTHD, vBAL_LIMIT, vRATE_DIFF, curDATE_OFFERED
    End If

End Sub
 
'===========================================================
' Output: File 2 Row (15-char aligned columns)
'===========================================================
Sub WriteFile2Row(shareType, cls, code, minToEarn, varMthd, balTo, rateDiff, firstOffered)

    Dim desc, savClass, ncua, classCode
    Dim arr

	desc = "" : savClass = "" : ncua = ""

	Dim key, hasLookup
	key = CStr(shareType)
	hasLookup = dictShareTypes.Exists(key)

	If hasLookup Then
		arr = dictShareTypes(key)
		savClass = CStr(arr(0))
		desc     = CStr(arr(1))
		ncua     = CStr(arr(2))
	Else
		' Option 1: Skip rows that have no lookup match
		If OPT_SkipIfNoLookup Then Exit Sub

		' Option 2: Flag missing lookup values in the output
		If OPT_FlagMissingLookup Then
			desc = "**MISSING LOOKUP**"
			savClass = ""
			ncua = ""
		End If
	End If


    classCode = cls & "/" & code

    fout.WriteLine _
        PadSpaces(Run_Date, 15) & _
        PadSpaces(shareType, 15) & _
        PadSpaces(desc, 30) & _
        PadSpaces(classCode, 15) & _
        PadSpaces(savClass, 40) & _
        PadSpaces(ncua, 15) & _
        PadSpaces(minToEarn, 15) & _
        PadSpaces(varMthd, 15) & _
        PadSpaces(balTo, 15) & _
        PadSpaces(rateDiff, 20) & _
        PadSpaces(firstOffered, 15)

End Sub

'===========================================================
' Fixed-width slicer (uses widths)
'===========================================================
Function SplitFields(line)
    Dim arr(), j, p, w, s, totalLen
    ReDim arr(UBound(widths))

    totalLen = 0
    For j = 0 To UBound(widths)
        totalLen = totalLen + widths(j)
    Next

    s = line
    If Len(s) < totalLen Then s = s & Space(totalLen - Len(s))

    p = 1
    For j = 0 To UBound(widths)
        w = widths(j)
        arr(j) = Trim(Mid(s, p, w))
        p = p + w
    Next

    SplitFields = arr
End Function

Function PadSpaces(val, sp)
    PadSpaces = Left(CStr(val) & Space(sp), sp)
End Function

'===========================================================
' Excel Lookup Loader: C:\SCCCU\ShareTypes.xlsx
' Builds: dict(key= Savings Type) = Array(SavingsClass, Description, NCUACode)
'===========================================================
Sub LoadShareTypesLookup(xlsxPath, dict)

    On Error Resume Next

    Dim xl, wb, ws
    Dim r, lastRow
    Dim cType, cClass, cDesc, cNcua
    Dim vType, vClass, vDesc, vNcua

    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False

    Set wb = xl.Workbooks.Open(xlsxPath, False, True) ' ReadOnly=True
    Set ws = wb.Worksheets(1)

    ' Find columns by header row (row 1)
    cType  = FindHeaderCol(ws, "Savings Type")
    cClass = FindHeaderCol(ws, "Savings Class")
    cDesc  = FindHeaderCol(ws, "Description")
    cNcua  = FindHeaderCol(ws, "NCUA CODE")

    ' If any header not found, exit cleanly
    If cType = 0 Or cClass = 0 Or cDesc = 0 Or cNcua = 0 Then
        wb.Close False
        xl.Quit
        Set ws = Nothing : Set wb = Nothing : Set xl = Nothing
        Exit Sub
    End If

    ' Determine last row based on Savings Type column
    lastRow = ws.Cells(ws.Rows.Count, cType).End(-4162).Row ' -4162 = xlUp

    For r = 2 To lastRow
        vType  = Trim(CStr(ws.Cells(r, cType).Value))
        If Len(vType) = 0 Then
            ' allow blanks inside sheet? skip
        Else
            vClass = Trim(CStr(ws.Cells(r, cClass).Value))
            vDesc  = Trim(CStr(ws.Cells(r, cDesc).Value))
            vNcua  = Trim(CStr(ws.Cells(r, cNcua).Value))

            ' Store/overwrite by Savings Type
            dict(CStr(vType)) = Array(vClass, vDesc, vNcua)
        End If
    Next

    wb.Close False
    xl.Quit

    Set ws = Nothing
    Set wb = Nothing
    Set xl = Nothing

    On Error GoTo 0
End Sub

Function FindHeaderCol(ws, headerText)
    Dim c, lastCol, v
    FindHeaderCol = 0

    lastCol = ws.Cells(1, ws.Columns.Count).End(-4159).Column ' -4159 = xlToLeft
    For c = 1 To lastCol
        v = Trim(CStr(ws.Cells(1, c).Value))
        If StrComp(v, headerText, vbTextCompare) = 0 Then
            FindHeaderCol = c
            Exit Function
        End If
    Next
End Function
