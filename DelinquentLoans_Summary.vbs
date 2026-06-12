Option Explicit

' Variables Declaration
Dim StartPick, OriginDate, LastPayDate, DelqDays, LBalance, LType, LName
Dim LoanTypeCounts, LoanTypeBalances, LoanTypeNames, DelinquencyPeriods
Dim LoanType, LoanTypeName, Period, LineCount, CurrentPeriod

' Dictionaries for tracking data
Set LoanTypeCounts = CreateObject("Scripting.Dictionary")
Set LoanTypeBalances = CreateObject("Scripting.Dictionary")
Set LoanTypeNames = CreateObject("Scripting.Dictionary")
Set DelinquencyPeriods = CreateObject("Scripting.Dictionary")

LineCount = 0
CurrentPeriod = ""

' Header for the output
Dim headerLine
headerLine = "Delinquency Period    Loan Type    Loan Type Name                          Count      Balance"

Private Sub ProcessLine
    LineCount = LineCount + 1

    ' Detect and handle delinquency period header
    If InStr(sRec, "Days DQ") > 0 Then
        CurrentPeriod = Trim(sRec)
        If Not DelinquencyPeriods.Exists(CurrentPeriod) Then
            DelinquencyPeriods.Add CurrentPeriod, True
        End If
    End If

    ' Process loan type data
    If IsNumeric(Left(sRec, 10)) Then
        ' Skip Account Number and Account Name lines
        Exit Sub
    ElseIf Trim(Left(sRec, 4)) = "L" Then
        LType = Mid(sRec, 48, 4)  ' Extract Loan Type
        LName = Trim(Mid(sRec, 12, 30)) ' Extract Loan Type Name
        LBalance = CDbl(Trim(Mid(sRec, 58, 15))) ' Extract Balance

        ' Dynamically store the loan type name
        If Not LoanTypeNames.Exists(LType) Then
            LoanTypeNames.Add LType, LName
        End If

        ' Initialize loan type counters if not already set
        If Not LoanTypeCounts.Exists(CurrentPeriod & LType) Then
            LoanTypeCounts.Add CurrentPeriod & LType, 0
            LoanTypeBalances.Add CurrentPeriod & LType, 0
        End If

        ' Increment counts and balances
        LoanTypeCounts(CurrentPeriod & LType) = LoanTypeCounts(CurrentPeriod & LType) + 1
        LoanTypeBalances(CurrentPeriod & LType) = LoanTypeBalances(CurrentPeriod & LType) + LBalance
    End If
End Sub

Private Sub CloseDoc()
    Dim PeriodKey, LoanKey, LoanTypeName
    Dim OutputLine  
    Dim FirstRow 
    Dim PeriodTotalCount  
    Dim PeriodTotalBalance  
    Dim ProcessedPeriods

    fout.WriteLine headerLine

    ' Track processed periods to prevent duplicate subtotals
    Set ProcessedPeriods = CreateObject("Scripting.Dictionary")

    ' Loop through delinquency periods and loan types
    For Each PeriodKey In DelinquencyPeriods.Keys
        FirstRow = True ' Flag to track the first row of the period
        PeriodTotalCount = 0 ' Reset subtotal count for the period
        PeriodTotalBalance = 0 ' Reset subtotal balance for the period

        For Each LoanKey In LoanTypeCounts.Keys
            If Left(LoanKey, Len(PeriodKey)) = PeriodKey Then
                LoanType = Mid(LoanKey, Len(PeriodKey) + 1)
                If LoanTypeNames.Exists(LoanType) Then
                    LoanTypeName = LoanTypeNames(LoanType)
                Else
                    LoanTypeName = "Unknown Loan Type"
                End If

                ' Only include the delinquency period for the first row
                If FirstRow Then
                    OutputLine = PeriodKey & Space(25 - Len(PeriodKey)) & LoanType & Space(10 - Len(LoanType)) & _
                                 LoanTypeName & Space(40 - Len(LoanTypeName)) & LoanTypeCounts(LoanKey) & Space(9) & _
                                 FormatCurrency(LoanTypeBalances(LoanKey))
                    FirstRow = False ' Set flag to false after the first row
                Else
                    OutputLine = Space(25) & LoanType & Space(10 - Len(LoanType)) & LoanTypeName & _
                                 Space(40 - Len(LoanTypeName)) & LoanTypeCounts(LoanKey) & Space(9) & _
                                 FormatCurrency(LoanTypeBalances(LoanKey))
                End If

                fout.WriteLine OutputLine

                ' Accumulate totals for the period
                PeriodTotalCount = PeriodTotalCount + LoanTypeCounts(LoanKey)
                PeriodTotalBalance = PeriodTotalBalance + LoanTypeBalances(LoanKey)
            End If
        Next

        ' Write the subtotal line for the period if not zero
        If PeriodTotalCount > 0 Or PeriodTotalBalance > 0 Then
            fout.WriteLine Space(25) & "SUBTOTAL" & Space(45) & PeriodTotalCount & Space(9) & FormatCurrency(PeriodTotalBalance)
            fout.WriteLine "" ' Add blank line after each period
        End If

        ' Mark this period as processed
        ProcessedPeriods.Add PeriodKey, True
    Next

    ' Ensure the last subtotal is printed even if not followed by another period
    If Not ProcessedPeriods.Exists(CurrentPeriod) Then
        fout.WriteLine Space(25) & "SUBTOTAL" & Space(45) & PeriodTotalCount & Space(9) & FormatCurrency(PeriodTotalBalance)
    End If
End Sub