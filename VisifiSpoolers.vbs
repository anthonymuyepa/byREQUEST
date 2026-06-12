Option Explicit
Dim filename, dstgroup,  SpoolerNo
 
Sub startDoc
	fout.WriteLine "Windows Registry Editor Version 5.00"
	fout.WriteLine "" 

	SpoolerNo = 99
End Sub

Sub ProcessLine()
       If srec <> "" Then
	    SpoolerNo = SpoolerNo + 1
		fout.WriteLine "" 
        filename = ""
        dstgroup = ""
        filename = Split(srec, "|")(0)
        dstgroup = Split(srec, "|")(1)
		
	
		fout.WriteLine "[HKEY_CURRENT_USER\Software\Hillary Software\byREQUEST\Spoolers\" & SpoolerNo & "]"
        fout.WriteLine """Name""=""" & filename & """"
        fout.WriteLine """Enabled""=dword:ffffffff"
        fout.WriteLine """Host""=""H4"""
        fout.WriteLine """Filter""=""" & filename & "*"""
        fout.WriteLine """Domain""=dword:00000001"
        fout.WriteLine """SymToPublish""=dword:00000000"
        fout.WriteLine """SymOverRidePathString""="""""
        fout.WriteLine """SymOptOverRidePathString""=dword:00000000"
        fout.WriteLine """SymSearchString""="""""
        fout.WriteLine """SymIsADataFile""=dword:00000000"
        fout.WriteLine """SymFullReportName""=dword:00000000"
        fout.WriteLine """SymDaysBackToSearch""=dword:00000000"
        fout.WriteLine """SymDefaultDaysBackToSearch""=dword:00000004"
        fout.WriteLine """SymStartLookingOnLine""=dword:00000000"
        fout.WriteLine """SymSearchStartsAtPosition""=dword:00000000"
        fout.WriteLine """SymUseDaysBackToSearch""=dword:00000000"
        fout.WriteLine """SymUseStartLookingOnLine""=dword:00000000"
        fout.WriteLine """SymUseSearchStartsAtPosition""=dword:00000000"
        fout.WriteLine """Content""=dword:00000001"
        fout.WriteLine """Strip Vax Extensions""=dword:00000000"
        fout.WriteLine """Location""=""C:\\Data\\SFTP\\home\\Visifi\\"""
        fout.WriteLine """ChkReady""=dword:ffffffff"
        fout.WriteLine """ChkDefer""=dword:ffffffff"
        fout.WriteLine """ChkSpSave""=dword:ffffffff"
        fout.WriteLine """ChkProblm""=dword:00000000"
        fout.WriteLine """ChkStdList""=dword:00000000"
        fout.WriteLine """ChkStalled""=dword:ffffffff"
        fout.WriteLine """ChkActive""=dword:00000000"
        fout.WriteLine """ChkCreate""=dword:00000000"
        fout.WriteLine """ChkUnindexed""=dword:ffffffff"
        fout.WriteLine """Suppress Blank Lines""=dword:00000000"
        fout.WriteLine """Suppress Blank Pages""=dword:00000001"
        fout.WriteLine """Skip Initial Lines""=dword:00000000"
        fout.WriteLine """LocalYes""=dword:00000000"
        fout.WriteLine """LocalWhat""=dword:ffffffff"
        fout.WriteLine """LocalType""=dword:ffffffff"
        fout.WriteLine """LocalValue""="""""
        fout.WriteLine """Templates Path""="""""
		
		If InStr(filename,"RGLACR10") > 0 Then
			fout.WriteLine """Spooler Template""=""EOM Reports_LandScape"""
			fout.WriteLine """Pre-Processing""=""Script Echo"""
		ElseIf InStr(filename,"RGL07011MP") > 0 _
			Or InStr(filename,"RGL07011QP") > 0 _
			Or InStr(filename,"RGL07011YP") > 0 _
			Or InStr(filename,"RGL07012MP") > 0 _
			Or InStr(filename,"RGL07012QP") > 0 _
			Or InStr(filename,"RGLFTI10RF") > 0 _
			Or InStr(filename,"RGLFTI20DF") > 0 _
			Or InStr(filename,"RGLFTI20RF") > 0 _
			Or InStr(filename,"RGLFTI30RF") > 0 _
			Or InStr(filename,"RGLTBL10F") > 0 _
			Or InStr(filename,"RMGGLS10") > 0 _
			Or InStr(filename,"RMG25020") > 0 _
			Or InStr(filename,"RMGGLS20") > 0 Then

			fout.WriteLine """Spooler Template""=""Accounting_LandScape"""
			fout.WriteLine """Pre-Processing""=""Script Echo"""
			
		ElseIf InStr(filename,"RMG58020") > 0 Then

			fout.WriteLine """Spooler Template""=""Management_LandScape"""
			fout.WriteLine """Pre-Processing""=""Script Echo"""
		
		Else
			fout.WriteLine """Spooler Template""="""""
			fout.WriteLine """Pre-Processing""="""""
		End If


        fout.WriteLine """Publishing""=dword:00000001"
		
		If InStr(filename,"RBX00")>0 then 
		    fout.WriteLine """Publishing Transport""=""Discard"""
			fout.WriteLine """Publishing Address""=""(none)""" 
			fout.WriteLine """Publishing Format""=""Text"""
			fout.WriteLine """Publishing Delete""=dword:00000001"
			fout.WriteLine """Publishing Command""="""""
		Else
			fout.WriteLine """Publishing Transport""=""PC Network"""
			fout.WriteLine """Publishing Address""=""\\\\server1\\shared files\\" & dstgroup & "\\!0-!2-!3\\!n""" 
			fout.WriteLine """Publishing Format""=""Direct PDF"""
			fout.WriteLine """Publishing Delete""=dword:00000000"
			fout.WriteLine """Publishing Command""="""""
		End if
		
        fout.WriteLine """Schedule Key""=""Schedule 2"""
        fout.WriteLine """Chain To""="""""
		
       


    End If
	fout.WriteLine "" 
	
End Sub

Sub CloseDoc

	    SpoolerNo = SpoolerNo + 1
		fout.WriteLine "" 

		fout.WriteLine "[HKEY_CURRENT_USER\Software\Hillary Software\byREQUEST\Spoolers\" & SpoolerNo & "]"
        fout.WriteLine """Name""=""Z_UnPublished"""
        fout.WriteLine """Enabled""=dword:ffffffff"
        fout.WriteLine """Host""=""H4"""
        fout.WriteLine """Filter""=""*"""
        fout.WriteLine """Domain""=dword:00000001"
        fout.WriteLine """SymToPublish""=dword:00000000"
        fout.WriteLine """SymOverRidePathString""="""""
        fout.WriteLine """SymOptOverRidePathString""=dword:00000000"
        fout.WriteLine """SymSearchString""="""""
        fout.WriteLine """SymIsADataFile""=dword:00000000"
        fout.WriteLine """SymFullReportName""=dword:00000000"
        fout.WriteLine """SymDaysBackToSearch""=dword:00000000"
        fout.WriteLine """SymDefaultDaysBackToSearch""=dword:00000004"
        fout.WriteLine """SymStartLookingOnLine""=dword:00000000"
        fout.WriteLine """SymSearchStartsAtPosition""=dword:00000000"
        fout.WriteLine """SymUseDaysBackToSearch""=dword:00000000"
        fout.WriteLine """SymUseStartLookingOnLine""=dword:00000000"
        fout.WriteLine """SymUseSearchStartsAtPosition""=dword:00000000"
        fout.WriteLine """Content""=dword:00000001"
        fout.WriteLine """Strip Vax Extensions""=dword:00000000"
        fout.WriteLine """Location""=""C:\\Data\\SFTP\\home\\Visifi\\"""
        fout.WriteLine """ChkReady""=dword:ffffffff"
        fout.WriteLine """ChkDefer""=dword:ffffffff"
        fout.WriteLine """ChkSpSave""=dword:ffffffff"
        fout.WriteLine """ChkProblm""=dword:00000000"
        fout.WriteLine """ChkStdList""=dword:00000000"
        fout.WriteLine """ChkStalled""=dword:ffffffff"
        fout.WriteLine """ChkActive""=dword:00000000"
        fout.WriteLine """ChkCreate""=dword:00000000"
        fout.WriteLine """ChkUnindexed""=dword:ffffffff"
        fout.WriteLine """Suppress Blank Lines""=dword:00000000"
        fout.WriteLine """Suppress Blank Pages""=dword:00000001"
        fout.WriteLine """Skip Initial Lines""=dword:00000000"
        fout.WriteLine """LocalYes""=dword:00000000"
        fout.WriteLine """LocalWhat""=dword:ffffffff"
        fout.WriteLine """LocalType""=dword:ffffffff"
        fout.WriteLine """LocalValue""="""""
        fout.WriteLine """Templates Path""="""""
        fout.WriteLine """Spooler Template""="""""
        fout.WriteLine """Pre-Processing""="""""
        fout.WriteLine """Publishing""=dword:00000001"
        fout.WriteLine """Publishing Transport""=""PC Network"""
        fout.WriteLine """Publishing Address""=""\\\\server1\\shared files\\UnPublished\\!0-!2-!3\\!n""" 
        fout.WriteLine """Publishing Format""=""Direct PDF"""
        fout.WriteLine """Publishing Delete""=dword:00000000"
        fout.WriteLine """Publishing Command""="""""
        fout.WriteLine """Schedule Key""=""Schedule 4"""
        fout.WriteLine """Chain To""="""""
 

End Sub
