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
Const PAD = 5 
Dim allLines(), lineParts, maxWidths()
Dim i, j, line, rowCount, colCount


Sub StartDoc()

End Sub


Sub processLine

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

