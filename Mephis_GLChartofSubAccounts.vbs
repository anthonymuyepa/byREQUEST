Rem  +=======================================================+
Rem  |  Memphis City - GL Chart of SubAccounts
Rem  |  Hillary Software  - March 2019
Rem  |  Edited 11/12/2020 ZTS
Rem  |  Changed to create function that returns ThreeDigit Letter date named ThreeDigit
Rem  |  Made sub StoreDates() to be run on first line of file to store day from file inner data
Rem  |    u7 is Day in ##
Rem  |    u8 is Month in ##
Rem  |   u9 is Month is @@@ where @ is month character three letters
Rem  |
Rem  |  Rev 2-8-2021 for just 2 or 3 lines per entry / cds
Rem  |  Rev 2-23-2021 to bring down correct Acct Num / cds
Rem  +=======================================================+

Option Explicit

dim newRec, tmpStr,endSw,firstSw,comm1,comm2,firstHd,detHd1,detHd2,hdrAcctNum,hdrAcctName,sLen,totSw 
dim Hdr(12)
hdrAcctNum=Space(16) 
hdrAcctName=Space(42)

endSw=0
firstSw=0 
firstHd=0
detHd1="GL Account number GL Account Name                                       SubAccount   Type     Name                            Offset GL        Acquisition    Maturity   Start       Effective  Post Through "
detHd2="  Original value      Limit value    Std Accrual     Current value  Accrued amount  Comments          "
 
Private Sub Processline


 If InStr(sRec,"GL Chart of SubAccounts") > 0  Then    ' Top of page Headings  5 lines
   If firstHd < 1 Then
    StoreDates()
    StoreHeader
    Else  bypassHeader    
   End If
    firstHd=1
  End If

  If InStr(sRec,"Offset GL ") > 0 Then
    If firstHd < 1 Then
     fout.Writeline Space(2)
     WriteGLHeader
     endSw = 2
    End If
  End If
 
  If InStr(sRec,"-----------") > 0 Then
     sRec=Getline
  End If

  If InStr(sRec,"Sub Account Total") > 0 Then
     totSw=1
     writeDetail
     sRec=Getline
  End If

  If InStr(sRec,"following GL SubAccounts") > 0 Then          ' Store the account number and name, output as first 2 columns of Detail
    ' dim sLen
     sLen=Len(sRec) - 67
     sRec=Replace(sRec,"<"," ")
     sRec=Replace(sRec,">"," ")
     hdrAcctNum=Mid(sRec,51,16) 
     hdrAcctName=Mid(sRec,68,sLen) & Space(52 - sLen)
     totSw=0
'  msgbox sRec
  End If
 
  If IsNumeric(Mid(sRec,1,4)) Then    
       writeDetail
       tmpStr=sRec & "  "
  End If

  If InStr(sRec,"Original:") > 0 Then      ' append to first line
     If totSw < 1 Then                     ' Use this to avoid adding the Totals part
        tmpStr = tmpStr & Mid(sRec,12,16) & Mid(sRec,38,16) & Mid(sRec,64,16) & Mid(sRec,88,18) & Mid(sRec,117,15)
     End If
  End If

  If InStr(sRec,"Comments:") > 0 then      
    tmpstr=tmpStr & Space(2) & Mid(sRec,11,42)
  End If


End Sub

Sub StoreHeader
    dim h
    fout.Writeline sRec
    'Spoolfile.User1="Fixed Assets"
    Do While h < 4    
       sRec=Getline
      h=h + 1     
    Loop

    fout.Writeline detHd1 & detHd2        ' Column headings
    fout.Writeline sRec & sRec & "-----------------------------------"     ' Should be a line of dashes
 
End Sub

Sub writeDetail
'  Get rid of inline headings
 
   If firstSw > 0 Then
      If Len(Trim(tmpStr)) > 30 Then
         fout.Writeline hdrAcctNum & Space(2) & hdrAcctName & Space(2) & tmpStr
         tmpStr=Space(360)
      End If
   End If

   firstSw = 1
  
End Sub

Sub WriteTotalHeader
 '   firstHd=1
    dim h
    h=0
    Hdr(0) = Chr(12) & hdr(0)   ' add form feed
    Hdr(6) = "Tab: Totals"
    Do While h < 8                     
       fout.Writeline Hdr(h)
       h = h + 1
    Loop

End Sub
       
Sub WriteGLHeader
    dim i
    Hdr(6) = "Tab: GL Account"
    Do While i < 8
       fout.Writeline Hdr(i)
       i = i + 1
    Loop

End Sub

Sub formatGL

    If InStr(sRec,"Exception ") > 0 Then                             ' To align the totals
       sRec=Replace(sRec,":"," ")
    End If

    sRec=Space(390) & Mid(sRec,1,17) & Space(2) & Mid(sRec,18,52) & Mid(sRec,108,25)    ' Gets rid of extra spaces btw amount and date
   
    fout.Writeline sRec

End Sub

Sub bypassHeader        ' Do not include headings on subsequent pages
    sRec=Getline
    sRec=Getline
    sRec=Getline
    sRec=Getline
 
'  msgbox "Srec is " & sRec

End Sub

Private Sub StoreDates()
'  msgbox("in StoreDates")  & Mid(sRec,83,2) & "++"
 
  dim m, y, d, tmp
  m = Mid(sRec,83,2)    ' The position within sRec may vary from one file to another - this is the internal Run or As-of Date
  d = Mid(sRec,86,2)
  y = Mid(sRec,89,2)
  tmp = DateAdd("d",-1,d & "-" & ThreeDigit(m) & "-" & y)
  spoolfile.User2 = year(tmp)
  spoolfile.User3 = day(tmp)
  spoolfile.User4 = month(tmp)
  spoolfile.User5 = ThreeDigit(month(tmp))
  spoolfile.User6 = y
  spoolfile.User7 = d
  spoolfile.User8 = m
  spoolfile.User9 = ThreeDigit(m)
  'msgbox(spoolfile.User7)
  'msgbox(spoolfile.User8)
  'msgbox(spoolfile.User9)

End Sub

Private Function ThreeDigit(MonthDigit)
  'msgbox("In ThreeDigit " & MonthDigit)
  Dim MonthArray
    Dim Variable4

    MonthArray = Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
    Variable4 = Cint(MonthDigit)
    'msgbox(Variable4)
  ThreeDigit=MonthArray(Variable4-1)

End Function
