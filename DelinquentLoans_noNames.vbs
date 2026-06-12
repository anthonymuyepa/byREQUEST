'************************************************************
' Author: ACM, HillarySoftware
' Date: 04 Feb 2025
' Summary: Processes loan data, filters out "Method B DQ Calculation" loans,
'          categorizes them by delinquency period, and outputs totals.
'************************************************************

Option Explicit

' Variable Declarations
Dim StartPick, OriginDate, LastPayDate, DelqDays, LBalance, LType, LName
Dim LoanTypeCounts, LoanTypeBalances, LoanTypeNames, DelinquencyPeriods
Dim LoanType, LoanTypeName, Period, LineCount, CurrentPeriod, sRec
Dim bufferLine

' Initialize dictionaries
Set LoanTypeCounts = CreateObject("Scripting.Dictionary")
Set LoanTypeBalances = CreateObject("Scripting.Dictionary")
Set LoanTypeNames = CreateObject("Scripting.Dictionary")
Set DelinquencyPeriods = CreateObject("Scripting.Dictionary")

LineCount = 0
CurrentPeriod = ""
bufferLine = ""

' Header for report
Dim headerLine
headerLine = "Delinquency Period    Loan Type    Loan Type Name                          Count      Balance"

' === MAIN LINE PROCESSING ===
Private Sub ProcessLine
    LineCount = LineCount + 1

    ' Detect delinquency period
    If InStr(sRec, "Days DQ") > 0 Then
        CurrentPeriod = Trim(sRec)
        If Not DelinquencyPeriods.Exists(CurrentPeriod) Then
            DelinquencyPeriods.Add CurrentPeriod, True
        End If
    End If

    ' If the last buffered line was a loan, check if this is the Method B line
    If bufferLine <> "" Then
        If InStr(sRec, "Subject to Method B DQ Calculation") > 0 Then
            bufferLine = ""  ' It's Method B, skip it
            Exit Sub
        Else
            ' Not Method B, process the loan

            Call ProcessLoanLine(bufferLine)
            bufferLine = ""
        End If
    End If

    ' Buffer this line if it's a loan line
    If InStr(Left(sRec, 4), "L") > 0 Then
        bufferLine = sRec

    End If
End Sub

' === PROCESS A LOAN LINE ===
Private Sub ProcessLoanLine(dRec)
    Dim compositeKey
    LType = Mid(dRec, 48, 4)

	
    LName = Trim(Mid(dRec, 12, 30))
    LBalance = CDbl(Trim(Mid(dRec, 58, 15)))
    compositeKey = CurrentPeriod & LType

    If Not LoanTypeNames.Exists(LType) Then
        LoanTypeNames.Add LType, LName
    End If

    If Not LoanTypeCounts.Exists(compositeKey) Then
        LoanTypeCounts.Add compositeKey, 0
        LoanTypeBalances.Add compositeKey, 0
    End If

    LoanTypeCounts(compositeKey) = LoanTypeCounts(compositeKey) + 1
    LoanTypeBalances(compositeKey) = LoanTypeBalances(compositeKey) + LBalance
End Sub

' === SORTING FUNCTION ===
Private Function GetSortedKeys(dict)
    Dim keysArray, i, j, temp
    keysArray = dict.Keys
    For i = UBound(keysArray) To 1 Step -1
        For j = 0 To i - 1
            If keysArray(j) > keysArray(j + 1) Then
                temp = keysArray(j)
                keysArray(j) = keysArray(j + 1)
                keysArray(j + 1) = temp
            End If
        Next
    Next
    GetSortedKeys = keysArray
End Function

' === OUTPUT THE REPORT ===
Private Sub CloseDoc()
    Dim PeriodKey, LoanKey, LoanTypeName
    Dim OutputLine
    Dim FirstRow
    Dim PeriodTotalCount
    Dim PeriodTotalBalance
    Dim LastProcessedPeriod
    Dim LastPeriodTotalCount
    Dim LastPeriodTotalBalance

    fout.WriteLine headerLine

    ' Loop through delinquency periods
    For Each PeriodKey In DelinquencyPeriods.Keys
        FirstRow = True
        PeriodTotalCount = 0
        PeriodTotalBalance = 0

        ' Get matching loan keys
        Dim sortedLoanKeys
        Set sortedLoanKeys = CreateObject("Scripting.Dictionary")

        For Each LoanKey In LoanTypeCounts.Keys
            If Left(LoanKey, Len(PeriodKey)) = PeriodKey Then
                sortedLoanKeys.Add LoanKey, LoanTypeCounts(LoanKey)
            End If
        Next

        Dim sortedKeys
        sortedKeys = GetSortedKeys(sortedLoanKeys)

        For Each LoanKey In sortedKeys
            LoanType = Mid(LoanKey, Len(PeriodKey) + 1)
            If LoanTypeNames.Exists(LoanType) Then
                LoanTypeName = LoanTypeNames(LoanType)
            Else
                LoanTypeName = "Unknown Loan Type"
            End If

            If FirstRow Then
                OutputLine = PeriodKey & Space(25 - Len(PeriodKey)) & LoanType & Space(10 - Len(LoanType)) & _
                             Space(40) & LoanTypeCounts(LoanKey) & Space(9) & _
                             FormatCurrency(LoanTypeBalances(LoanKey))
                FirstRow = False
            Else
                OutputLine = Space(25) & LoanType & Space(10 - Len(LoanType))& _
                             Space(40) & LoanTypeCounts(LoanKey) & Space(9) & _
                             FormatCurrency(LoanTypeBalances(LoanKey))
            End If

            fout.WriteLine OutputLine
            PeriodTotalCount = PeriodTotalCount + LoanTypeCounts(LoanKey)
            PeriodTotalBalance = PeriodTotalBalance + LoanTypeBalances(LoanKey)
        Next

        ' Period subtotal
        If PeriodTotalCount > 0 Or PeriodTotalBalance > 0 Then
            fout.WriteLine Space(25) & "SUBTOTAL" & Space(45) & PeriodTotalCount & Space(9) & FormatCurrency(PeriodTotalBalance)
            fout.WriteLine ""
        End If

        LastProcessedPeriod = PeriodKey
        LastPeriodTotalCount = PeriodTotalCount
        LastPeriodTotalBalance = PeriodTotalBalance
    Next

    ' Final total
    If LastProcessedPeriod <> "" And (LastPeriodTotalCount > 0 Or LastPeriodTotalBalance > 0) Then
        fout.WriteLine Space(25) & "SUBTOTAL" & Space(45) & LastPeriodTotalCount & Space(9) & FormatCurrency(LastPeriodTotalBalance)
    End If
End Sub
