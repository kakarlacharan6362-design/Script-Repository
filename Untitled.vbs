'Defining Variables
Set oShell = CreateObject("WScript.Shell")
set oEnv = oShell.Environment("PROCESS")
Set oFSO = CreateObject("Scripting.FileSystemObject")
strSysDir = oShell.ExpandEnvironmentStrings("%SystemDrive%")
strPgmDir = oShell.ExpandEnvironmentStrings("%programdata%")
strProgDir = oShell.ExpandEnvironmentStrings("%ProgramFiles(x86)%")
strtemp = oShell.ExpandEnvironmentStrings("%temp%")
ScriptFullName = wscript.scriptFullname
ScriptFolderName = left(scriptFullName, InStrRev(ScriptFullName, "\"))
oEnv("SEE_MASK_NOZONECHECKS") = 1
'##################################################################################################################################

strcmd1 =  "Dism.exe /Online /Enable-Feature /FeatureName:MSMQ-Server /All /NoRestart"
rtnVal = oShell.Run(strcmd1,0, True)


'##################################################################################################################################

Sub subCreateFolderPath(strPath)

	Dim oFSO
	Set oFSO = CreateObject("Scripting.FileSystemObject")

	Dim aFolders
	aFolders = Split(strPath,"\")

	Dim strFolder
	strFolder =aFolders(0)

	Dim i
	For i=1 To UBound(aFolders)-1
		strFolder = strFolder & "\" & aFolders(i) 
		If Not oFSO.FolderExists(strFolder) = True Then
			oFSO.CreateFolder(strFolder)
		End If

	Next

	Set oFSO = Nothing
End Sub

'=========================================================================================='==========================================================================================
Public Function KillProcess(ProcessName)
	Dim objWMIService, colProcessList, objProcess
	Const strComputer = "."
	Set objWMIService = GetObject("winmgmts:" _
	& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colProcessList = objWMIService.ExecQuery ("SELECT * FROM Win32_Process")
	For Each objProcess In colProcessList
		If LCase(objProcess.Name) =  LCase(ProcessName) Then
			objProcess.Terminate()
			Exit For
		End If
	Next
	WScript.Sleep 2000
	If ProcessActive(ProcessName) Then KillProcess(ProcessName)
	Set objWMIService = Nothing
	Set colProcessList = Nothing
End Function 

'=========================================================================================='==========================================================================================
'=========================================================================================='==========================================================================================
Public Function ProcessActive(Proc_Name)
    Dim Process_set : Set Process_set = GetObject("winmgmts:").ExecQuery("select * from Win32_Process")
    Dim Process
    ProcessActive = False
    For Each Process In Process_set ' Enumerate all processes
        If LCase(Proc_Name) = LCase(Process.Name) Then
            ProcessActive = True
            Exit Function
        End If
    Next
End Function

'=========================================================================================='==========================================================================================



Public Function regExists(regKey)
set oShell = CreateObject("wscript.shell")
	On Error Resume Next
	regExists= oShell.RegRead (regKey)
	If not isEmpty(regExists) then	
	      regExists=True 
	Else
	      regExists=False
	End If
End Function