Option Explicit
'Attribute VB_Name = "Module1"

' IF THIS SCRIPT IS MODIFIED YOU MUST UPDATE HOW THE VERSION IS USED IN ADVANCED INSTALLER
' HOPEFULLY IF SOME HAS MODIFIED THIS SCRIPT ON SOME OTHER MACHINE OTHER THAN brqdev.hillary.local
' THEY HAVE UPDATED THIS FILE on brqdev.hillary.local in \data\byRequest\v5\staging\scripts\control.vbs  (or sent the modification to support@hillary.com) so that when
' we do an update we don't lose the modifications
' Last modification date
'   July 20,2024  Ashcraft  Fix bug in PeekAhead.  Forced getline to return 1 space when a line of 0 length was read from brqTemp1.data
'                           This prevents reading past a FF when doing a peekahead and find an input record of length 0
'   Nov. 2, 2022  Ashcraft   Removed fin.close and fout.close in BinaryMainLoop
'   March 3, 2022 Ashcraft  Merged several copies of control.vbs to try to get the "right"
'                                       version in the installer file
'   5-2-2020   Ashcraft   removed the close of fin and fout at the end of BinaryMainLoop  Probably not a problem bit it was wrong since BinaryMainLoop does not use fin and fout

'   8-23-2019  Ashcraft  Add function JoinWithNLinesWithVariableLength
'   4-26-2017  Ashcraft
'   Added BinaryMainLoop so that binary xfers can have a preprocessing script that doesn't corrupt the original binary file
'   1-22-2016  Ashcraft
'   Added StartDoc() and CloseDoc() back into MainLoop
'   6-26-2014  Ashcraft
'     Added local variables s1-s9 into SetUserVariablesBasedOnFilename()
'   6-23-2014  Ashcraft/Sellitto
'   6-24-2014
'      Fixed bug in StrRepl -- Zane
'   6-23-2014 Added 'by Ref' to SplitString subroutine,
'             Dim'd inName in SetUserVariablesBasedOnFilename
'   6-6-2014 removed 'msgbox from getline()
'   Added "Page " and "PAGE " to DeletePage function
' last modification date  5-28-2014  Ashcraft
' merged and cleaned up which subroutines and functions are latest 5-28-2014  Ashcraft
' control.vbs now contains the following subroutines and functions
'     CombineGroupOfLines(nextRec,maxLength,startExpression,endExpression)
'     DeletePage ()
'     GetLine()
'     Initialize
'     JoinWithNextLine
'     JoinWithNLines (n)
'     PeekAhead
'     Rearrange(oldstring,numpieces,startpositions,lengthofpositions)
'     ReturnNthLine(n)
'     SetUserVariablesBasedOnFilename (firstRec)
'     SkipNLines(n)
'     SplitString(StringToSplit,SplitDelimiter,parts1,parts2,parts3,parts4,parts5,parts6,parts7,parts8,parts9)
'     StretchLine
'     StrRepl(sRec, sWith, iWhere)
' 11-20-2019 Sellitto Added function "lookupToExcel" to lookup Excel values, given filename, column to return,
'     value to match, column in which the matching value is found


Const ForReading = 1
Dim fIn           ' brqTemp1, for reading
           
Dim leftWidth
Dim addWidth
Dim fOut          ' brqTemp2, for writing
Dim sRec          ' Input buffer from brqTemp1
Dim iPg           ' Current page number
Dim iLn           ' Line number of page
Dim iRn           ' Record number in brqTemp1
Dim sTempPath     ' Path of Temp folder
Dim sAppPath      ' Application Path
Dim sScriptsPath   ' Scripts path
Dim foundStartOfGroup ' Have we found the start of a group
Dim groupOfLines  ' group of lines being built
Dim xxxRec     ' I would rather have had a Static variable inside GetLine

Dim PeekAheadRecord ' record saved when PeekAhead is called and then returned
                    ' when the next getLine() is called


Private Sub StartDoc()
' dummy that can be overwritten in script file
                                                                         
End Sub
Private Sub CloseDoc()
  'msgbox " dummy that can be overwritten in script file"
End Sub
Function GetProductVersion(myFile)
        ' Based on code by Maputi on StackOverflow.com:
        ' http://stackoverflow.com/a/2990698
        Dim arrTranslations
        Dim i
        Dim objFolder, objFolderItem, objShell
        Dim strFileName, strPropertyName, strParentFolder, strVersion
        ' Note that property names are language dependent, so you may have to add the lower case property name for your own language
        Set arrTranslations = CreateObject("System.Collections.ArrayList")
        arrTranslations.Add "product version" ' English
        arrTranslations.Add "productversie"   ' Dutch
        strVersion = ""
        strFileName = objFSO.GetFileName(myFile)
        strParentFolder = objFSO.GetParentFolderName(myFile)
        Set objShell = CreateObject("Shell.Application")
        Set objFolder = objShell.Namespace(strParentFolder)
        Set objFolderItem = objFolder.ParseName(strFileName)
        For i = 0 To 300
                strPropertyName = objFolder.GetDetailsOf(objFolder.Items, i)
                If arrTranslations.Contains(LCase(strPropertyName)) Then
                        ' Product Version
                        strVersion = objFolder.GetDetailsOf(objFolderItem, i)
                        Exit For
                End If
        Next
        Set objFolderItem = Nothing
        Set objFolder = Nothing
        Set objShell = Nothing
        Set arrTranslations = Nothing
        ' Replace commas by dots
        strVersion = Replace(strVersion, ",", ".")
        ' Remove spaces
        strVersion = Replace(strVersion, " ", "")
        GetProductVersion = strVersion
End Function
Private Function JoinWithNLines(n)
Dim NewLine, posct, nextRec
  
  NewLine = sRec
  
  Dim l
  posct = 133 - Len(sRec)

  For l = 1 To (n - 1)
      nextRec = GetLine
      NewLine = NewLine & Space(posct) & nextRec
      posct = 133 - Len(nextRec)

  Next
  
  JoinWithNLines = NewLine

End Function

Private Function JoinWithNLinesWithVariableLength(n, MaxLineLen)
Dim NewLine, posct, nextRec
  
  NewLine = sRec
  
  Dim l
  posct = MaxLineLen - Len(sRec)
  
  For l = 1 To (n - 1)
      nextRec = GetLine
      If posct < 1 Then
         posct = 1
      End If
      NewLine = NewLine & Space(posct) & nextRec
      posct = MaxLineLen - Len(nextRec)
  Next
  
  JoinWithNLinesWithVariableLength = NewLine

End Function


'
' rearrange a string
'    newstring = rearrange(oldstring, numpieces,startpositions,lengthofpieces)
'    oldstring is the original string
'    numpieces is the number of pieces that are to be extracted and recombined to create newstring
'    startpositions() is an array containing the beginning position of the strings to be used
'    lengthofpositions() is an array that corresponds to the length of the position specified by startpositions()

Private Function rearrange(oldstring, numpieces, startpositions, lengthofpositions)
    Dim i
    Dim nextPiece
    rearrange = ""
    For i = 0 To numpieces - 1
        nextPiece = Mid(oldstring, startpositions(i), lengthofpositions(i))
        rearrange = rearrange + nextPiece
    Next
End Function


' SKip N lines
Private Sub SkipNLines(n)
    Dim tmp      ' dummy variable to hold last line being skipped
    tmp = ReturnNthLine(n)
End Sub

' read the next n lines and return the last line read
Private Function ReturnNthLine(n)
    Dim i        ' loop counter
    For i = 1 To n
        ReturnNthLine = GetLine
    Next
End Function


    
'combine lines containing the start expression until find a new line
'with  the finish  expression
Private Sub CombineGroupOfLines(nextRec, maxLength, startExpression, endExpression)
    'nextRec is srec
    'maxLength is used to pad lines so columns line up
    'startExpression is executed to determine if we are at the beginning
    '      of a next section
    '
    ' blank lines are skipped (?)
    ' heading lines are output so that byRequest can clean them up later
    Dim spacesForAlignment     ' computed number of spaces to make each line
                               '  same length
    Dim currentRec               ' contains the record to be used in the
                                 '   evaluation of start and stop expressions
    Dim Group                    ' contains the group of lines being combined
    Dim peek                     ' contains the peekahead record
    Dim tmp
    Dim startVal
    Dim endVal
    Dim val
    currentRec = nextRec
    tmp = ""
    Group = ""
    groupOfLines = ""
    execute startExpression
    execute endExpression

    If startVal = True Or endVal = True Then
        groupOfLines = groupOfLines & currentRec & Space(maxLength - Len(currentRec))
        peek = PeekAhead
        currentRec = peek
        If startVal = True Or endVal = True Then
            If startVal = True And Len(groupOfLines) > 0 Then
               ' found the start of a new group
               fOut.WriteLine groupOfLines
               Exit Sub
            ElseIf endVal = True Then
               tmp1 = peek     ' the line we just peeked at
               groupOfLines = groupOfLines & tmp1
               fOut.WriteLine groupOfLines
               Exit Sub
            Else
               groupOfLines = groupOfLines & Space(maxLength - Len(currentRec))
            End If
        Else
          fOut.WriteLine groupOfLines
          tmp = PeekAhead
          currentRec = GetLine
          execute startExpression
          execute endExpression
        End If
    Else
       fOut.WriteLine currentRec
    End If
End Sub

Private Sub JoinWithNextLine()
   fOut.WriteLine Left(sRec, leftWidth) & Mid(GetLine, leftWidth - addWidth + 1)
End Sub

' ---------------------------------------------

Private Sub StretchLine()
   fOut.WriteLine Left(sRec, leftWidth) & Space(addWidth) & Mid(sRec, leftWidth + 1)
End Sub

' ---------------------------------------------

Private Sub StrRepl(sRec, sWith, iWhere)
   Dim NewRec
   
   If Len(sRec) < iWhere - 1 Then sRec = sRec & Space(iWhere - Len(sRec) - 1)
   
   NewRec = Left(sRec, iWhere - 1) & sWith
   
   If Len(sRec) > iWhere + Len(sWith) Then NewRec = NewRec & Mid(sRec, iWhere + Len(sWith))
   
   sRec = NewRec
End Sub


' ---------------------------------------------

Private Function Initialize()
   On Error Resume Next
   Initialize = 1
   PeekAheadRecord = ""
   xxxRec = ""
   sAppPath = Options.Folder(eApplication)
   sTempPath = Options.Folder(eTempFiles)
   sScriptsPath = Options.Folder(eScripts)
   
   Err.Clear
   
   Err.Clear
   Dim oFSO: Set oFSO = CreateObject("Scripting.FileSystemObject")
   
   Err.Clear
   Set fIn = oFSO.OpenTextFile(sTempPath & "brqTemp1.dat", 1)
   If Err.Number <> 0 Then
      MsgBox "Initialize: Unable to open " & sTempPath & "brqTemp1.dat - " & Err.Description & " err=" & Err.Number
      Initialize = 0
   Else
                Err.Clear
                oFSO.deletefile (sTempPath & "brqTemp2.dat")
                Err.Clear
                Set fOut = oFSO.createTextFile(sTempPath & "brqTemp2.dat", True)
                Dim myErr
                Dim myErrDescription
                myErrDescription = Err.Description
                myErr = Err.Number
                'msggox "err after create brqtemp2 " & myerr & " " & myErrDescription
                If myErr <> 0 Then
                        MsgBox "Initialize:Unable to create file " & sTempPath & "brqTemp2.dat" & " - " & myErrDescription
                        fIn.Close
                        Initialize = 0
                End If
        End If
        If Initialize = 1 Then

                iPg = 1
                iLn = 0
                iRn = 0
    End If
End Function

' ---------------------------------------------



Private Function GetLine()
   Dim p
   Dim Rec
   
   'If Len(xxxRec) = 0 Then
      If fIn.AtEndOfStream Then
         If Len(PeekAheadRecord) > 0 Then
            xxxRec = PeekAheadRecord
            PeekAheadRecord = ""
         Else
            GetLine = ""
            Exit Function
         End If
      Else
         If Len(PeekAheadRecord) > 0 Then
            xxxRec = PeekAheadRecord
            PeekAheadRecord = ""
         Else
            xxxRec = fIn.ReadLine
            If Len(xxxRec) = 0 Then
                xxxRec = " "    'Using len xxxRec to indicate no value available for peekahead so force xxxRec to have a len when first read
            End If
            If InStr(xxxRec, "11514") > 0 Then
                xxxRec = xxxRec
            End If
         End If
      End If
   'End If
         
   p = InStr(2, xxxRec, Chr(12))
   
   If p > 0 Then 'There's a page break part way through this record
      Rec = Left(xxxRec, p - 1)
      xxxRec = Mid(xxxRec, p)
   Else
      Rec = xxxRec
      xxxRec = ""
   End If
   
   iRn = iRn + 1
   iLn = iLn + 1

   If Len(Rec) > 0 Then
      If Left(Rec, 1) = Chr(12) Then
         iPg = iPg + 1
         iLn = 1
      End If
   End If
   If Len(Rec) < leftWidth Then Rec = Rec & Space(leftWidth - Len(Rec))
   GetLine = Rec
End Function

Private Function PeekAhead()
   If fIn.AtEndOfStream Then
      PeekAhead = ""
      PeekAheadRecord = ""
      Exit Function
   End If
   PeekAheadRecord = GetLine
   PeekAhead = PeekAheadRecord
End Function

Function readBinary(StrPath)

    Dim oFSO: Set oFSO = CreateObject("Scripting.FileSystemObject")
    Dim oFile: Set oFile = oFSO.GetFile(StrPath)

    If IsNull(oFile) Then
                MsgBox ("File not found: " & StrPath)
                Exit Function
        End If

    With oFile.OpenAsTextStream()
        readBinary = .Read(oFile.Size)
        .Close
    End With

End Function

Function writeBinary(strBinary, StrPath)

    Dim oFSO: Set oFSO = CreateObject("Scripting.FileSystemObject")

    ' below lines pupose: checks that write access is possible!
    Dim oTxtStream: Set oTxtStream = oFSO.createTextFile(StrPath)
    On Error Resume Next
    If Err.Number <> 0 Then MsgBox (Err.message): Exit Function
    On Error GoTo 0
    Set oTxtStream = Nothing
    ' end check of write access

    With oFSO.createTextFile(StrPath)
        .Write (strBinary)
        .Close
    End With

End Function
' ---------------------------------------------
' BinaryMainLoop   used to preprocess binary files so we can use StartDoc and CloseDoc without corrupting file
Sub BinaryMainLoop()
   sTempPath = Options.Folder(eTempFiles)
   StartDoc
   Err.Clear
   writeBinary readBinary(sTempPath & "brqTemp1.dat"), sTempPath & "brqTemp2.dat"
   If Err Then
     MsgBox "Error in BinaryMainLoop " & Err.Description
      Exit Sub
   End If
   CloseDoc
End Sub

Sub MainLoop()
        On Error Resume Next
        Dim ret
        ret = Initialize
        'msgbox ("control ret=" & ret)
   If ret = 1 Then
      StartDoc
          'msgbox ("control after StartDoc")
      Do While (Not fIn.AtEndOfStream) Or (Len(PeekAheadRecord) > 0)
         sRec = GetLine
         Call ProcessLine
      Loop
          'msgbox ("control before closedoc")
      CloseDoc
          'msgbox ("control after closedoc")

      fIn.Close
      fOut.Close
   End If
   Set fIn = Nothing
   Set fOut = Nothing
End Sub


'    File User variables script written by Hillary Software, Inc. 2014
'    Call from vbs script, pass in Spoolfile name containing double underscores as delimiter
'    and up to 9 User variable fields.
'    Returns up to 9 pieces of the file name to be used as Spoolfile.User1 - 9
'    Calling program must include all parameters in the 'call' statement, but Spoolfile.Name
'    You need NOT have all 9 pieces. Must explicitly define 1 parameter in calling program for
'        the value to be passed back to cause the code to be executed only once.
' Sample code for calling program:
'     Call SetUserVariablesBasedOnFilename (firstRec)
' ---------------------------------------------

Private Sub SplitString(StringToSplit, SplitDelimiter, ByRef parts1, ByRef parts2, ByRef parts3, ByRef parts4, ByRef parts5, ByRef parts6, ByRef parts7, ByRef parts8, ByRef parts9)
Dim Parts
 
 Parts = Split(StringToSplit, SplitDelimiter, -1)   ' Delimiter is Double Underscore
 On Error Resume Next
   parts1 = Parts(0)
   parts2 = Parts(1)
   parts3 = Parts(2)
   parts4 = Parts(3)
   parts5 = Parts(4)
   parts6 = Parts(5)
   parts7 = Parts(6)
   parts8 = Parts(7)
   parts9 = Parts(8)

End Sub

' --------------------------------------------
Private Sub SetUserVariablesBasedOnFilename(firstRec)
Dim s1, s2, s3, s4, s5, s6, s7, s8, s9
Dim inName
inName = Spoolfile.Name
 If firstRec = 0 Then
    Call SplitString(inName, "__", s1, s2, s3, s4, s5, s6, s7, s8, s9) ' Split spoolfile name with double-underscore delimiter
    Spoolfile.User1 = s1                                                ' Set User variables for use in naming files.
    Spoolfile.User2 = s2
    Spoolfile.User3 = s3
    Spoolfile.User4 = s4
    Spoolfile.User5 = s5
    Spoolfile.User6 = s6
    Spoolfile.User7 = s7
    Spoolfile.User8 = s8
    Spoolfile.User9 = s9
   
  End If

End Sub

Private Sub DeletePage()
Dim p1, p2, p3, p4
  ' Strip out page numbers by looking for key words
  ' "PAGE:"  "Page:" 'PAGE ' or 'Page '
  p1 = InStr(sRec, "PAGE:")
  p2 = InStr(sRec, "Page:")
  p3 = InStr(sRec, "PAGE ")
  p4 = InStr(sRec, "Page ")
  If p1 > 0 Then
     sRec = Left(sRec, p1 - 1)
  ElseIf p2 > 0 Then
     sRec = Left(sRec, p2 - 1)
  ElseIf p3 > 0 Then
     sRec = Left(sRec, p3 - 1)
  ElseIf p4 > 0 Then
     sRec = Left(sRec, p4 - 1)
  End If
End Sub

Private Function ExcelLookup(StrPath, colNum, findCol, matchValue)
 ' find matchval in colNum and then return corresponding value in findCol
        ExcelLookup = ""
    Dim intRowCount, intColCount
        Dim r, xlApp, xlWb, sht1
        Set xlApp = CreateObject("Excel.Application")
        xlApp.DisplayAlerts = False
        Set xlWb = xlApp.WorkBooks.Open(StrPath)
        Set sht1 = xlWb.sheets(1)
        intRowCount = sht1.UsedRange.Rows.Count
        r = 0
        Do Until r = intRowCount
                r = r + 1
                If StrComp(Trim(sht1.cells(r, colNum).Text), Trim(matchValue)) = 0 Then  ' Column colNum contains the value to be matched
           ExcelLookup = Trim(sht1.cells(r, findCol))        ' Column findCol contains value to be returned
                   Exit Do
                End If
        Loop
        Set sht1 = Nothing
        xlWb.Close
        Set xlWb = Nothing
        xlApp.Quit
        Set xlApp = Nothing
End Function
Private Function lookupToExcel(StrPath, colNum, findCol, matchValue)
 ' nice function but looks like someone modified it so I can't use it since I don't know what it is being used for NOW.  Copied to create new function that doesn
 ' what this one did originally.  Why would someone write a nice function and then have someone butcher it to do something else    ???   Steve
    Dim returnValue
    returnValue = "0.00"                    ' If this is zero, value was not found in Previous Month Excel file
    'Dim intRowCount, intColCount

        'Dim r, xlApp, xlWb, sht1
        'Set xlApp=CreateObject("Excel.Application")
        'xlApp.DisplayAlerts=False
        'set xlWb=xlApp.WorkBooks.Open(StrPath)
        'set sht1=xlWb.sheets(1)
        'intRowCount=sht1.UsedRange.Rows.Count
        'intColCount=sht1.UsedRange.Columns.Count
        r = 0
        Do Until r = intRowCount
                r = r + 1
                If Trim(sht1.cells(r, colNum)) = Trim(matchValue) Then    ' Column colNum contains the value to be matched
                        returnValue = Trim(sht1.cells(r, findCol))        ' Column findCol contains value to be returned
                        Exit Do
                End If
        Loop
        'Set sht1=nothing
        'xlWb.Close
        'set xlWb=nothing
        'xlApp.Quit
        'set xlApp=nothing
  
     lookupToExcel = returnValue
End Function
