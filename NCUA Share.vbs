
dim hdr 
hdr="S        ACCOUNT #/ID                NAME                                             ADDRESS                               CITY                   ST        ZIPCODE          EXTRA ADDRESS                         BALANCE                SH CD        SSN/TIN #           DT. GRNTD         MAT. DATE          DIVRATE       ACT. DT.       AMT. FRZN                 ACT. CODE           INT ACCR.       FM DATE           FM USER"


Sub StartDoc()
	fout.writeline hdr
End Sub







Private Sub ProcessLine()
    srec = Replace(srec, vbTab, "        ")
    fout.WriteLine srec
End Sub