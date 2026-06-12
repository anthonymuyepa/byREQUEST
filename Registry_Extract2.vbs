Option Explicit

Dim SpName,Filter,SplTemp,FileName,PubAddr,PubFormat,PubTrans 
Dim Lnct, HKey, InDent,LocalValue, Location, LocalWhat, LocalType
Dim outRec

OutRec=Space(260)
Const OverwriteTrue = True
Const ReplaceChr = " "

Lnct=0
InDent = 16



Private Sub ProcessLine()

   If InStr(sRec,"\byREQUEST\Spoolers\") > 0 Then
      writeRec
   End If
   
   If InStr(sRec,"FullReportName") >  0 Then
      InDent= 2
   ElseIf InStr(sRec,"Name") > 0 Then
   
      SpName=Mid(sRec,9,80)
      SpName=Replace(SpName,"""","")  
   End If
 
   If InStr(sRec,"Filter") > 0 Then
      Filter=Mid(sRec,11,120)
      Filter=Replace(Filter,"""","")
   End If
 
   If InStr(sRec,"Host") > 0 Then
      HKey=Mid(sRec,9,3)
      HKey=Replace(HKey,"""","")
   End If

   If InStr(sRec,"Location") > 0 Then
      Location = Mid(sRec,19,80)
      Location=Replace(Location,"""","")
   End If

   If InStr(sRec,"Spooler Template") > 0 Then
      SplTemp=Mid(sRec,21,80)
      SplTemp=Replace(SplTemp,"""","")
   End If

   If InStr(sRec,"Publishing Transport") > 0 Then
      PubTrans=Mid(sRec,25,40)
      PubTrans=Replace(PubTrans,"""","")
   End If
   
   If InStr(sRec,"Publishing Address") > 0 Then
      PubAddr=Mid(sRec,23,120)
      PubAddr=Replace(PubAddr,"""","")
   End If   
      
   If InStr(sRec,"Publishing Format") > 0 Then
     PubFormat=Mid(sRec,21,40)
     PubFormat=Replace(PubFormat,"""","")
   End If

   If InStr(sRec,"LocalWhat") > 0 Then
     LocalWhat=Mid(sRec,19,20)
     LocalWhat=Replace(LocalWhat,"""","")
   End If

   If InStr(sRec,"LocalType") > 0 Then
     LocalType=Mid(sRec,19,20)
     LocalType=Replace(LocalType,"""","")
   End If

   If InStr(sRec,"LocalValue") > 0 Then
     LocalValue=Mid(sRec,14,80)
     LocalValue=Replace(LocalValue,"""","")
   End If

End Sub

Sub writeRec()
    

    outRec=SpName & Space(42 - Len(SpName)) & HKey & " " & Location & Space(42 - Len(Location)) & Filter & Space(32 - Len(Filter)) & LocalValue & Space(26 - Len(LocalValue)) & " [" 
    outRec=outRec & SplTemp & "]" & Space(32 - Len(SplTemp)) & PubAddr & Space(60 - Len(PubAddr)) & PubFormat  
    outRec=outRec & Space(14 - Len(PubFormat)) & PubTrans & Space(28 - Len(PubTrans)) & LocalWhat & " " & LocalType 

    fout.Writeline outRec

 

    outRec=Space(260)

End Sub