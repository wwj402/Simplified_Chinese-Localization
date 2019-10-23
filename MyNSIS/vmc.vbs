Function OnFirstSandboxOwner
    msgbox "The sandbox owner is: " + GetCurrentProcessName
    Origin = GetEnvironmentVariable("TS_ORIGIN")
    LastSlash = InStrRev(Origin, "\")
    SourcePath = Left(Origin, LastSlash)
    BinPath = ExpandPath("%SystemSystem%\drivers\vstor2-mntapi10-shared.sys")
    Set WshShell = CreateObject("WScript.Shell")
    WshShell.Run ("sc create vstor2-mntapi10-shared binpath= """ + BinPath + """ type= service start= demand displayname= ""Vstor2 MntApi 1.0 Driver (shared)""")
    BinPath = ExpandPath("%SystemSystem%\drivers\bmdrvr.sys")
    WshShell.Run ("sc create bmdrvr binpath= """ + BinPath + """ type= service start= demand displayname= ""Modified Clusters Tracking Driver""")

End Function

Function OnFirstParentExit
    msgbox "Quiting application: " + GetCurrentProcessName
End Function