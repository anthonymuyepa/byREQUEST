

Option Explicit

 

Sub StartDoc() 

Dim fso, folder, file, xlApp, wb, ws, fout
Dim r, valA, valC, filePath

' Initialize objects
Set fso = CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder("S:\JHU\ACHreportCombine")
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
xlApp.DisplayAlerts = False

fout.WriteLine "🚀 Scanning first 500 rows of Sheet1"

For Each file In folder.Files
    If LCase(fso.GetExtensionName(file.Name)) = "xlsx" Then
        If InStr(file.Name, "04-") > 0 Then
            filePath = file.Path
            fout.WriteLine "📄 File: " & file.Name

            Set wb = xlApp.Workbooks.Open(filePath, False, True) ' Open read-only
            Set ws = wb.Sheets(1)

            For r = 1 To 500
                valA = Trim(CStr(ws.Cells(r, 1).Text))
                valC = Trim(CStr(ws.Cells(r, 3).Text))

                If valA = "" And valC = "" Then
                    ' Optionally break early if we hit many empties in a row
                End If

                fout.WriteLine "🔎 Row " & r & " A=[" & valA & "] C=[" & valC & "]"

                If InStr(UCase(valC), "TOTAL") > 0 Then
                    fout.WriteLine "✅ Found TOTAL row at " & r
                End If
            Next

            wb.Close False
            Set ws = Nothing
            Set wb = Nothing
        End If
    End If
Next

' Cleanup
xlApp.Quit
Set xlApp = Nothing
Set fso = Nothing
Set folder = Nothing

fout.WriteLine "✅ Manual row scan complete."




End Sub
