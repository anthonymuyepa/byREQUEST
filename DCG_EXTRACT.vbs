'===============================================
' Anthonym@hillary
' Last Edit: 26 May 2025

' Script: Format CSV-Like Output for Alignment


' Description:
' This script 
' - Splits each line into columns
' - Calculates the maximum width of each column
' - Pads each field so that columns align nicely

'===============================================


Option Explicit
Const DELIM = ","
Const PAD = 1 
Dim allLines(), lineParts, maxWidths()
Dim i, j, line, rowCount, colCount, header
header = "Account No,Class Code,Type Code,Current Bal,Part Bal Owned,Part Bal Sold,Part % Owned,Part % Sold,Interest Rate,Reprice Index,Maturity Date,P&I Payment,P&I Part Pay,Pay Frequency,Pay Type Code,Next Pay Date,Interest Marg,Rate Chg Freq,Next Reprice,Rate Ceiling,Rate Floor,Cap per Period,Floor per Per,Accrual Code,Line of Cr ID,PPP Identif,Forecl Status,Original Term,Orig Date,Fix/Adj Code,Non-Accrual,Customer ID,Business Line,Household ID,Bank Code,Cashout Ref %,Collateral Cd,FICO Score,Full Doc %,Int Credit Rt,Delinq Status,Current LTV,Original LTV,Original Bal,Orig Branch,Orig Officer,Primary Res %,Property Type,State Code,CallRpt Code,M&A ID,NAICS Code,Teaser Mos,Teaser Left,Conv End Date,Convertible,Conv Start Dt,Country Code,County Code,Index Lookbk,Late Chgs YTD,Orig Int Rate,Orig Reprice,SingleFam %,GL Code,Prin Pay %,P+I Pay %,Escrow Bal,Escrow Pay,Forecl State,First Pay Dt,LL Reserve Cd,Plan Code,Purpose Code,Orig Prepay,Curr Prepay $,Curr Prepay %,Orig Part Bal,Gross Int Rt,Book Value,Par Value,Market Value,Cap Freq"
header ="Account Number,Class Code,Type Code,Current Account Balance,Participation Balance Owned *,Participation Balance Sold *,Participation Percentage Owned *,Participation Percentage Sold *,Interest Rate,Repricing Index Code,Maturity Date,P&I Payment,P&I Payment Participations,Payment Frequency,Payment Type - P&IIOPO Code,Next Payment Date,Interest Margin,Interest Rate Change Frequency,Next Repricing Date,Lifetime Rate Ceiling,Lifetime Rate Floor,Periodic Change Cap,Periodic Change Floor,Accrual Method Code, Line of Credit Identifier, PPP Identifier,Foreclosure Status Code,Original Term,Origination Date,Fixed or Adjustable Code,Non-Accrual Status,Customer ID,Business Line, Household ID,Bank Code,Cash Out Refinance Ratio,Collateral Code,FICO Score,Full Documented Ratio,Internal Credit Rating,Loan Delinquent,Current Loan to Value (LTV),Original LTV,Original Balance,Originating Branch,Originating Loan Officer,Primary Residential Ratio,Property Type,State Code,Call Report Code,M&A Identifier,NAICS Code,Teaser Months,Teaser Months Remaining,End of Conversion,Convertible Flag,Start of Conversion,Country Code,County Code,Index Lookback Days,Late Charges Paid YTD,Original Interest Rate,Original Repricing Date,Single Family Residential Ratio,General Ledger Code,Principal Payment Percent,Principal+Interest Payment %,Escrow Balance,Escrow Payment,Foreclosure State Code,First Payment Date,Loan Loss Reserve Code,Plan Code,Purpose Code,Original Prepayment Penalty,Current Prepayment Penalty Amount,Current Prepayment Penalty %,Original Participation Balance,Gross Interest Rate,Book Value,Par Value,Market Value,PERIODIC CAP FREQUENCY"


Sub StartDoc()
	'Skip the header, from removing percentage symbols 
	srec = getline
	ReDim Preserve allLines(rowCount)
    allLines(rowCount) = srec
    rowCount = rowCount + 1

End Sub


Sub processLine
	srec = replace(srec,"%","")
	ReDim Preserve allLines(rowCount)
    allLines(rowCount) = srec
    rowCount = rowCount + 1

End Sub


Sub closeDoc()
 
' Determine max column widths
	ReDim maxWidths(0)
		dim allLn
		allLn = UBound(allLines)

		For i = 0 To allLn
			lineParts = SplitCsvLine(allLines(i))
		
			If UBound(lineParts) > UBound(maxWidths) Then
				ReDim Preserve maxWidths(UBound(lineParts))
			End If
			For j = 0 To UBound(lineParts)
				If Len(lineParts(j)) > maxWidths(j) Then
					maxWidths(j) = Len(lineParts(j))
				End If
			Next
		Next
		

	For i = 0 To allLn
		lineParts = SplitCsvLine(allLines(i))
		line = ""
		dim linesall
		linesall = UBound(lineParts)
		
		
		For j = 0 To linesall
			line = line & PadRight(lineParts(j), maxWidths(j) + PAD)
		Next
		fout.WriteLine line
	Next
End Sub

' Helper: pad string to right
Function PadRight(s, width)
    PadRight = s & Space(width - Len(s))
End Function


Function SplitCsvLine(str)
    Dim i, c, part, inQuotes, out(), index
    index = 0
    part = ""
    inQuotes = False
    ReDim out(0)

    For i = 1 To Len(str)
		
        c = Mid(str, i, 1)
        If c = """" Then
            inQuotes = Not inQuotes
        ElseIf c = DELIM And Not inQuotes Then
			
            out(index) = part
            index = index + 1
            ReDim Preserve out(index)
            part = ""
        Else
            part = part & c
        End If
    Next
	
    out(index) = part  ' capture the final piece
    SplitCsvLine = out
End Function

