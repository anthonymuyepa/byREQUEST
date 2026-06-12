' Date Created: 2025-09-18
' Client Name : CarolinaTrust
' Description : 
'   This script processes NCUA loan data file records, outputs them with
'   a standardized header row, and replaces tabs with spaces
'   for alignment in the final text output.
' ============================================================

' -----------------------------
' Define the header row with column names
' -----------------------------
dim hdr 
hdr="L        ACCOUNT#-ID                 NAME                                             ADDRESS                               CITY                   ST        ZIP CODE         EXTRA ADDRESS                         LOAN TYPE    PAYMENT $              PURP CODE    PMNT COUNT PAY FREQ  LOAN DATE         ORIGINAL AMT           INT RATE    INT CODE      CURRENT BAL            LAST ACTIVITY     LASTACTCD    DUE DATE          INT DUE           CREDIT LIMIT          SOC SEC #           DAYS DQ     DQ31-60    DQ61-90    DQ91-120   DQ121+ INSIDERCODES  OFFICER/CC  CREDITSCORE CHG OFF AMT          RISK GRADE        PMTS LEFT   COLL CODE    LAST FM DATE  LAST FM USER   BRANCH"

Sub StartDoc()
	fout.writeline hdr
End Sub


Private Sub ProcessLine()
    srec = Replace(srec, vbTab, "        ")
    fout.WriteLine srec
End Sub