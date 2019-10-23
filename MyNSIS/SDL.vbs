' OnFirstSandboxOwner
' *Called only when an application first locks the sandbox. This callback is not called 
' if a second copy of the same application uses the same sandbox while the first copy runs. 
' If the first application spawns a subprocess and quits, the second subprocess locks the 
' sandbox and prevents this callback from running until all subprocesses quit and the 
' application runs again.
' OnFirstParentStart
' *Called before running a ThinApp executable file regardless of 
' whether the sandbox is simultaneously owned by another captured executable file.
' OnFirstParentExit
' *Called when the first parent process exits. If a parent process runs a child process and 
' quits, this callback is called even if the child process continues to run.
' OnLastProcessExit
' *Called when the last process owning the sandbox exits. If a parent process runs a child
' process and quits, this callback is called when the last child process exits.


Function OnFirstParentStart
    AppPath = ExpandPath("%ProgramFilesDir%\SDL Passolo 2018")
    AppExe = "PassoloLauncher18.0.97.0.exe"
    ' msgbox AppPath
    SetEnvironmentVariable "PSLHOME", AppPath
    ' HandleId = ExecuteVirtualProcess(AppPath & "\" & AppExe)
    ' WaitForProcess(HandleId, 0)
End Function