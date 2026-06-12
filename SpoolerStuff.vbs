
Dim gHeaderPart
Dim gReportDate, gSeqNumber, gHeaderParsed, gTargetMetric

Sub StartDoc()
 

    Dim Sep
    Sep = String(60, "-")

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.ID: " & Spoolfile.ID

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.Name: " & Spoolfile.Name

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.Account: " & Spoolfile.Account

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.Owner: " & Spoolfile.Owner

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.Job: " & Spoolfile.Job

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.JSNum: " & Spoolfile.JSNum

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.Device: " & Spoolfile.Device

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.Form: " & Spoolfile.Form

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.Group: " & Spoolfile.Group

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.System: " & Spoolfile.System

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.State: " & Spoolfile.State

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.Priority: " & CStr(Spoolfile.Priority)

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.NumPages: " & CStr(Spoolfile.NumPages)

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.KB: " & CStr(Spoolfile.KB)

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.Stamp: " & CStr(Spoolfile.Stamp)

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.DaysOld: " & CStr(Spoolfile.DaysOld)

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.MinsOld: " & CStr(Spoolfile.MinsOld)

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.Content: " & CStr(Spoolfile.Content)

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.ContentDesc: " & Spoolfile.ContentDesc

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.Spacing: " & CStr(Spoolfile.Spacing)

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.SpacingDesc: " & Spoolfile.SpacingDesc

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.Layout: " & CStr(Spoolfile.Layout)

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.LayoutDesc: " & Spoolfile.LayoutDesc

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.LPP: " & CStr(Spoolfile.LPP)

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.HostPath: " & Spoolfile.HostPath

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.Location: " & Spoolfile.Location

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.Key: " & Spoolfile.Key

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.HostKey: " & Spoolfile.HostKey

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.Indexed: " & CStr(Spoolfile.Indexed)

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.ProcStatus: " & CStr(Spoolfile.ProcStatus)

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.HeaderLength: " & CStr(Spoolfile.HeaderLength)

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.Folder: " & Spoolfile.Folder

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.IconName: " & Spoolfile.IconName

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.User0: " & Spoolfile.User0

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.User1: " & Spoolfile.User1

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.User2: " & Spoolfile.User2

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.User3: " & Spoolfile.User3

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.User4: " & Spoolfile.User4

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.User5: " & Spoolfile.User5

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.User6: " & Spoolfile.User6

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.User7: " & Spoolfile.User7

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.User8: " & Spoolfile.User8

    Fout.WriteLine Sep
    Fout.WriteLine "Spoolfile.User9: " & Spoolfile.User9

    Fout.WriteLine Sep

End Sub
