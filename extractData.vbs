

Option Explicit

Dim objShell, fso, objFolder, objFile, objRegExp, objMatches, objMatch, strText
Dim pdf2textExe, inputPdf, outputTxt, folderPath
'Dim fout     ' fout is globally defined. 


Sub startDoc()
	' Configure paths
	pdf2textExe = "S:\CarolinaTrust\byREQUEST\Pdf2Text\pdftotext.exe"
	folderPath = "S:\Tmp\" ' Folder with all PDF invoices

	' Create File System Object
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set objShell = CreateObject("WScript.Shell")


	If Not fso.FolderExists(folderPath) Then
	Fout.WriteLine "Folder does not exist: " & folderPath
	Exit sub
	End If

	Set objFolder = fso.GetFolder(folderPath)

	For Each objFile In objFolder.Files
		If LCase(fso.GetExtensionName(objFile.Name)) = "pdf" Then
			inputPdf = objFile.Path


			
			outputTxt = folderPath & "" & fso.GetBaseName(objFile.Name) & ".txt"


			' Convert PDF to text
			If fso.FileExists(pdf2textExe) Then
				objShell.Run Chr(34) & pdf2textExe & Chr(34) & " -table " & Chr(34) & inputPdf & Chr(34) & " " & Chr(34) & outputTxt & Chr(34), 0, True
			End If
			MsgBox "Extract Complete: view file in " & outputTxt

			' Read text output
			If fso.FileExists(outputTxt) Then
				Dim textStream
				Set textStream = fso.OpenTextFile(outputTxt, 1)
				strText = textStream.ReadAll
				
				textStream.Close

				' Extract fields
				Dim invoiceDate, invoiceNumber, paymentTerms, dueDate, orderDate, poNumber, customerNumber, amountDue
				invoiceDate = GetMatch(strText, "INVOICE DATE\s*[:\-]?\s*(\d{2}/\d{2}/\d{2})")
				invoiceNumber = GetMatch(strText, "INVOICE NUMBER\s*[:\-]?\s*(\S+)")
				paymentTerms = GetMatch(strText, "PAYMENT TERMS\s*[:\-]?\s*([\w\s]+)")
				dueDate = GetMatch(strText, "DUE DATE\s*[:\-]?\s*(\d{2}/\d{2}/\d{2})")
				orderDate = GetMatch(strText, "ORDER DATE\s*[:\-]?\s*(\d{2}/\d{2}/\d{2})")
				poNumber = GetMatch(strText, "PURCHASE ORDER NUMBER\s*[:\-]?\s*(\S+)")
				customerNumber = GetMatch(strText, "CUSTOMER NUMBER\s*[:\-]?\s*(\S+)")
				amountDue = GetMatch(strText, "AMOUNT DUE\s*[:\-]?\s*\$?([\d,]+\.\d{2})")
				
				MsgBox invoiceDate & "   " & invoiceNumber & "   " & paymentTerms 
				 
			End If
		End If

	Next

	Fout.Writeline "Batch processing completed."
End sub


Function GetMatch(strTxt, pattern)
    Dim re, matches
    
    Set re = New RegExp
    re.Pattern = pattern
    re.IgnoreCase = True
    re.Global = False
    
    Set matches = re.Execute(strTxt)
    
    If matches.Count > 0 Then
        GetMatch = Trim(matches(0).SubMatches(0))
    Else
        GetMatch = ""
    End If
End Function


Sub WriteOutput(invoiceDate,invoiceNumber, paymentTerms)
	fout.WriteLine invoiceDate & "   " & invoiceNumber & "   " & paymentTerms 
	MsgBox invoiceDate & "   " & invoiceNumber & "   " & paymentTerms 
				'& "   " & dueDate & "   " & orderDate & "   " & poNumber & "   " & customerNumber & "   " & amountDue 
End sub