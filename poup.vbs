Function CheckAndHandleProcess(procName)
    Dim wmi, procs, p, choice
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set procs = wmi.ExecQuery("Select * from Win32_Process Where Name='" & procName & "'")

    If procs.Count > 0 Then
        choice = MsgBox(procName & " is running." & vbCrLf & _
                        "Click YES to close it or NO to cancel installation.", _
                        vbYesNo + vbExclamation, "Installation Warning")

        If choice = vbYes Then
            For Each p In procs
                p.Terminate
            Next
            MsgBox procName & " has been closed. Proceed with installation.", vbInformation
        Else
            MsgBox "Installation cancelled.", vbCritical
            WScript.Quit
        End If
    End If
End Function

'--- Call the function ---
Call CheckAndHandleProcess("excel.exe")