'************************************************************
' Author: ACM, HillarySoftware (Updated by AM)
' Date: 04 Feb 2025
' Summary: Processes loan data, filters out "Method B DQ Calculation" loans,
'          categorizes them by delinquency period, and outputs totals.
'************************************************************

Option Explicit

' Variable Declarations
Dim StartPick, OriginDate, LastPayDate, DelqDays, LBalance, LType, LName
Dim LoanTypeCounts, LoanTypeBalances, LoanTypeNames, DelinquencyPeriods
Dim LoanType, LoanTypeName, Period, LineCount, CurrentPeriod, sRec
Dim bufferLine, isBuffering, methodBFound

' Initialize dictionaries
Set LoanTypeCounts = CreateObject("Scripting.Dictionary")
Set LoanTypeBalances = CreateObject("Scripting.Dictionary")
Set LoanTypeNames = CreateObject("Scripting.Dictionary")
Set DelinquencyPeriods = CreateObject("Scripting.Dictionary")

LineCount = 0
CurrentPeriod = ""
bufferLine = ""
isBuffering = False
methodBFound = False

' Header for report
Dim headerLine
headerLine = "Delinquency Period    Loan Type    Loan Type Name                          Count      Balance"

' === MAIN LINE PROCESSING ===
Private Sub ProcessLine
    LineCount = LineCount + 1

    ' Detect delinquency period
    If InStr(sRec, "Days DQ") > 0 Then
        Call FlushBufferIfNeeded()  ' Flush any leftover loan before changing section
        CurrentPeriod = Trim(sRec)
        If Not DelinquencyPeriods.Exists(CurrentPeriod) Then
            DelinquencyPeriods.Add CurrentPeriod, True
        End If
        Exit Sub
    End If

    ' Check for Method B tag while buffering
    If isBuffering And InStr(sRec, "Subject to Method B DQ Calculation") > 0 Then
        methodBFound = True
        Exit Sub
    End If

    ' Check if a new loan has started
    If InStr(Left(sRec, 4), "L") > 0 Then
        Call FlushBufferIfNeeded() ' Finalize previous loan before starting new one
        bufferLine = sRec
        isBuffering = True
        methodBFound = False
        Exit Sub
    End If

    ' All other lines are ignored unless they trigger one of the above
End Sub

' === FLUSH BUFFER ===
Private Sub FlushBufferIfNeeded()
    If isBuffering Then
        If Not methodBFound Then
            Call ProcessLoanLine(bufferLine)
        End If
        bufferLine = ""
        isBuffering = False
        methodBFound = False
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
    ' Final safety flush in case last loan was never ended
    Call FlushBufferIfNeeded()

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
                             LoanTypeName & Space(40 - Len(LoanTypeName)) & LoanTypeCounts(LoanKey) & Space(9) & _
                             FormatCurrency(LoanTypeBalances(LoanKey))
                FirstRow = False
            Else
                OutputLine = Space(25) & LoanType & Space(10 - Len(LoanType)) & LoanTypeName & _
                             Space(40 - Len(LoanTypeName)) & LoanTypeCounts(LoanKey) & Space(9) & _
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
