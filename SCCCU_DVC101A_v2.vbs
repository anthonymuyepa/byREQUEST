'==========================
' DVC101A File 3 FLATTENER
' - Drop-in replacement
' - Uses your working widths
' - Flattens continuation (tier) lines by inheriting the last header row
' - Outputs 15-char aligned columns (including Run_Date)
'==========================

Option Explicit

Dim reNum, Run_Date
Dim widths, i


widths = Array(4, 7, 5, 6, 6, 4, 4, 13, 6, 11, 10, 7, 8, 11, 5, 5, 5, 5, 10)

' Current (last-seen) header context for carrying down to tier rows
Dim curTYPE, curCLASS, curCODE, curPLAN, curPOST_PER, curERN_PER, curDIV_CDE
Dim curMIN_TO_EARN, curVAR_RTHD, curROLLBACK, curGRACE_DEP, curGRACE_WITH, curRENEW, curDATE_OFFERED
Dim tierIndex

Public Sub StartDoc

    fout.WriteLine "Run_Date       TYPE           CLASS          CODE           PLAN           POST_PER       ERN_PER     DIV_CDE           MIN_TO_EARN    VAR_RTHD      BAL_LIMIT       RATE/DIFF       INDEX         RED_DIV_RATE   REDBAL_LIMIT   ROLLBACK       GRACE_DEP      GRACE_WITH     RENEW          DATE_OFFERED"

    ' Read Run_Date from header (first few lines)
    Dim SrecLocal, p
    For i = 1 To 4
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
    If Len(Trim(srec)) = 0 Then Exit Sub                    ' blank line

    If InStr(1, srec, Chr(12), vbBinaryCompare) > 0 Then Exit Sub  ' form feed (FF)

    If InStr(1, srec, "DVC101A", vbTextCompare) > 0 Then Exit Sub
 
    If InStr(1, srec, "SANTA", vbTextCompare) > 0 Then Exit Sub
    If InStr(srec, "MIN BAL TO") > 0 Then Exit Sub
	If InStr(srec, "DIVIDEND") > 0 Then 

	 Exit Sub
	End if
	If InStr(srec, "----") > 0 Then 
	 Exit Sub
	End if

    If InStr(srec, "POST ERN DIV") > 0 Then Exit Sub
    If InStr(1, srec, "TYPE", vbTextCompare) > 0 Then Exit Sub
    If InStr(1, srec, "NUMBER", vbBinaryCompare) > 0 Then Exit Sub
    If InStr(1, srec, "TOTL RCD", vbTextCompare) > 0 Then Exit Sub	
    If InStr(1, srec, "END-OF-REPORT", vbTextCompare) > 0 Then Exit Sub	
	
	
	
	
    Dim vals, pos, f
    vals = SplitFields(srec)

    ' Map fields by position (matches your header order)
    Dim vTYPE, vCLASS, vCODE, vPLAN, vPOST_PER, vERN_PER, vDIV_CDE
    Dim vMIN_TO_EARN, vVAR_RTHD, vBAL_LIMIT, vRATE_DIFF, vINDEX, vRED_DIV_RATE, vREDBAL_LIMIT
    Dim vROLLBACK, vGRACE_DEP, vGRACE_WITH, vRENEW, vDATE_OFFERED

    vTYPE        = vals(0)
    vCLASS       = vals(1)
    vCODE        = vals(2)
    vPLAN        = vals(3)
    vPOST_PER    = vals(4)
    vERN_PER     = vals(5)
    vDIV_CDE     = vals(6)
    vMIN_TO_EARN = vals(7)
    vVAR_RTHD    = vals(8)
    vBAL_LIMIT   = vals(9)
    vRATE_DIFF   = vals(10)
    vINDEX       = vals(11)
    vRED_DIV_RATE= vals(12)
    vREDBAL_LIMIT= vals(13)
    vROLLBACK    = vals(14)
    vGRACE_DEP   = vals(15)
    vGRACE_WITH  = vals(16)
    vRENEW       = vals(17)
    vDATE_OFFERED= vals(18)

	
    ' Decide if this is a Product row or a continuation tier row
    Dim isProductRow, isTier
    isProductRow = (Len(vTYPE) > 0 And reNum.Test(vTYPE))
    isTier   = (Not isProductRow) And (Len(vBAL_LIMIT) > 0 Or Len(vRATE_DIFF) > 0 Or Len(vRED_DIV_RATE) > 0 Or Len(vREDBAL_LIMIT) > 0)

    ' 1) Header row: store context + output as tier 1 (or given index)
    If isProductRow Then

        curTYPE        = vTYPE
        curCLASS       = vCLASS
        curCODE        = vCODE
        curPLAN        = vPLAN
        curPOST_PER    = vPOST_PER
        curERN_PER     = vERN_PER
        curDIV_CDE     = vDIV_CDE
        curMIN_TO_EARN = vMIN_TO_EARN
        curVAR_RTHD    = vVAR_RTHD
        curROLLBACK    = vROLLBACK
        curGRACE_DEP   = vGRACE_DEP
        curGRACE_WITH  = vGRACE_WITH
        curRENEW       = vRENEW
        curDATE_OFFERED= vDATE_OFFERED

        tierIndex = 0
        tierIndex = tierIndex + 1
        If Len(vINDEX) = 0 Then vINDEX = CStr(tierIndex)

        WriteFlatRow curTYPE, curCLASS, curCODE, curPLAN, curPOST_PER, curERN_PER, curDIV_CDE, _
                     curMIN_TO_EARN, curVAR_RTHD, vBAL_LIMIT, vRATE_DIFF, vINDEX, vRED_DIV_RATE, vREDBAL_LIMIT, _
                     curROLLBACK, curGRACE_DEP, curGRACE_WITH, curRENEW, curDATE_OFFERED
        Exit Sub
    End If

    ' 2) Continuation/tier row: inherit context and output as next tier
    If isTier Then
        If Len(curTYPE) = 0 Then Exit Sub ' orphan tier, ignore

        tierIndex = tierIndex + 1
        If Len(vINDEX) = 0 Then vINDEX = CStr(tierIndex)

        ' If reduced fields are blank on tier lines, keep them blank (or carry down if you prefer)
        WriteFlatRow curTYPE, curCLASS, curCODE, curPLAN, curPOST_PER, curERN_PER, curDIV_CDE, _
                     curMIN_TO_EARN, curVAR_RTHD, vBAL_LIMIT, vRATE_DIFF, vINDEX, vRED_DIV_RATE, vREDBAL_LIMIT, _
                     curROLLBACK, curGRACE_DEP, curGRACE_WITH, curRENEW, curDATE_OFFERED
    End If

End Sub

'----------------------------
' Fixed-width slicer (uses widths)
'----------------------------
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

'----------------------------
' Writer: 15-char aligned output
'----------------------------
Sub WriteFlatRow(t, c1, c2, plan, postp, ernp, divc, mte, vr, bal, rate, idx, rrate, rbal, rb, gd, gw, rn, dt)

    Dim outLine
    outLine = _
        PadSpaces(Run_Date, 15) & _
        PadSpaces(t, 15) & _
        PadSpaces(c1, 15) & _
        PadSpaces(c2, 15) & _
        PadSpaces(plan, 15) & _
        PadSpaces(postp, 15) & _
        PadSpaces(ernp, 15) & _
        PadSpaces(divc, 15) & _
        PadSpaces(mte, 15) & _
        PadSpaces(vr, 15) & _
        PadSpaces(bal, 15) & _
        PadSpaces(rate, 15) & _
        PadSpaces(idx, 15) & _
        PadSpaces(rrate, 15) & _
        PadSpaces(rbal, 15) & _
        PadSpaces(rb, 15) & _
        PadSpaces(gd, 15) & _
        PadSpaces(gw, 15) & _
        PadSpaces(rn, 15) & _
        PadSpaces(dt, 15)

    fout.WriteLine outLine
End Sub

Function PadSpaces(val, sp)
    PadSpaces = Left(CStr(val) & Space(sp), sp)
End Function
