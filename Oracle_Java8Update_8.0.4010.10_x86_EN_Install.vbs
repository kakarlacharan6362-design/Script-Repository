' This is the standard install script for the AVIS engagement.
' Simply indicate the Package Name, Vendor, and Version below.
' Package Name:   Oracle_Java8Update_8.0.4010.10_x86_EN
' Package Vendor: Oracle 
' Package Version:8.0.4010.10
' Generic Prerequisites
'################################################################################################################################################
Set objShell = CreateObject("WScript.Shell")
set oEnv = objShell.Environment("PROCESS")
Set oFS = CreateObject("Scripting.FileSystemObject")
Set objRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
Dim strProcess,strProcess2, strProcess3, strProcess4

strProgDir = objShell.ExpandEnvironmentStrings("%ProgramFiles(x86)%")
strProgData = objShell.ExpandEnvironmentStrings("%Programdata%")
strDate = Now()
struser = objShell.ExpandEnvironmentStrings("%username%")
ScriptFullName = wscript.scriptFullname
ScriptFolderName = left(scriptFullName, InStrRev(ScriptFullName, "\"))
oEnv("SEE_MASK_NOZONECHECKS") = 1

LogFolder = strsysdir & "\DXC\"
LogFolders = strsysdir & "\DXC\SCCM\"

If Not(oFS.FolderExists(LogFolder)) Then
    oFS.CreateFolder(LogFolder)
End If

If Not(oFS.FolderExists(LogFolders)) Then
    oFS.CreateFolder(LogFolders)
End If
'###########################################################################################################################

' Uninstalling previous version

'##############################################################################################################################
Dim appName, appVersion, appVendor, logName, logFolder, rtnVal
Dim objShell, objWMIService, strIdentifyingNumber, colSoftware, objSoftware, strQuerySoftToRemove, strCmdLine, strComputer, WinDir, InstallationLog, InstallationLogDebug  
Dim strWMIQuery1, strProcess1, strSysDir
Dim strFromPathName, strProgData, ScriptFolderName
Dim oFS, FLD

' Set the path to the folder where the current script is located
ScriptFolderName = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))

'==========================================================================================
'Start of Script (MAIN)
'=========================================================================================
strQuerySoftToRemove = "SELECT * FROM Win32_Product WHERE Name like 'Java % Update%'"

'==========================================================================================
WinDir = objShell.ExpandEnvironmentStrings("%WINDIR%")
strSysDir = objShell.ExpandEnvironmentStrings("%SystemDrive%")
strProgData = objShell.ExpandEnvironmentStrings("%ProgramData%")
strComputer = "."

Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colSoftware = objWMIService.ExecQuery(strQuerySoftToRemove)

For Each objSoftware in colSoftware
    appVendor = objSoftware.Vendor
    appName = objSoftware.Name
    appVersion = objSoftware.version
    logName = appVendor & "_" & appName & "_" & appVersion & "_"
    logFolder = strSysDir & "\DXC\SCCM\"
    
    subCreateFolderPath logFolder

    InstallationLog = logFolder & logName & "PackageUninstaller.log"
    InstallationLogDebug = logFolder & logName & "PackageUninstaller_SCRIPT.log"
    InstallationLog = Replace(InstallationLog, " ", "_")
    InstallationLogDebug = Replace(InstallationLogDebug, " ", "_")
    DebugLog 0, "***************************************************************************************"
    DebugLog 0, "***************************************************************************************"
    DebugLog 0, "Running script " & Wscript.ScriptFullName

    strIdentifyingNumber = objSoftware.IdentifyingNumber
    strCmdLine = "msiexec.exe /X " & strIdentifyingNumber & " /l*v " & Chr(34) & InstallationLog & Chr(34) & " /qn"

    DebugLog 0, "Checking if javaw is running in the machine or not..."
    If ProcessActive("javaw.exe") Then
        DebugLog 0, "javaw.exe is running in the machine. Killing the process..."
        DebugLog 0, " "	
        KillProcess "javaw.exe"
    End If

    rtnVal = ""
    DebugLog 0, "Executing the command: '" & strCmdLine & "'"
    rtnVal = objShell.Run(strCmdLine, 0, True)
    
    ' Capture return code
    If rtnVal = 0 Or rtnVal = 3010 Then
        DebugLog 0, "The application " & appName & " " & appVersion & " was uninstalled successfully with return " & rtnVal

        ' Additional commands for registry modification
        If regExists("HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\JavaSoft\Java Update\Policy\EnableJavaUpdate") Then
            objShell.RegDelete "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\JavaSoft\Java Update\Policy\EnableJavaUpdate"
        End If
    Else
        ' If one component or app fails, exit the script
        DebugLog 2, "The application " & appName & " " & appVersion & " uninstallation failed with return " & rtnVal	
        DebugLog 0, "***************************************************************************************"
        DebugLog 0, "***************************************************************************************"
        'End
    End If
Next

' If previous version is not present, install the current version

'###########################################################################################################################

' Installing current version

'##############################################################################################################################
strFromPathName = ScriptFolderName
Set oFS = CreateObject("Scripting.FileSystemObject")
Set FLD = oFS.GetFolder(strFromPathName)


For Each aItem In FLD.Files 
   If LCase(Right(Cstr(aItem.Name), 3)) = "exe" Then
   
strcmd = chr(34) & aItem & Chr(34) & " /s INSTALL_SILENT=1 /L "  & chr(34) & strSysDir & "\DXC\SCCM\Oracle_Java8Update_401_8.0.4010.10_x86_EN_Install.log" & chr(34)
rtnVal = objShell.Run(strcmd,0, True)
  
   End If
Next

If rtnVal = 0 Or rtnVal = 3010 Then

objShell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\JavaSoft\Java Update\Policy\EnableJavaUpdate", "0", "REG_DWORD"

If oFS.FileExists(strProgData &"\Microsoft\Windows\Start Menu\Programs\Java\Check For Updates.lnk") Then
oFS.Deletefile(strProgData &"\Microsoft\Windows\Start Menu\Programs\Java\Check For Updates.lnk") 
End If

If oFS.FileExists(strProgData &"\Microsoft\Windows\Start Menu\Programs\Java\Get Help.url") Then
oFS.Deletefile(strProgData &"\Microsoft\Windows\Start Menu\Programs\Java\Get Help.url") 
End If

If oFS.FileExists(strProgData &"\Microsoft\Windows\Start Menu\Programs\Java\Visit Java.com.url") Then
oFS.Deletefile(strProgData &"\Microsoft\Windows\Start Menu\Programs\Java\Visit Java.com.url") 
End If

End If
'################################################################################################################################################
Set objShell = Nothing

'################################################################################################################################################
'End of Script (MAIN)
'=========================================================================================='==========================================================================================
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
'Function for Check & write The Registry in machine
'=========================================================================================='==========================================================================================
Public Function Writereg(regkey,regvalue,regtype)
	On Error Resume Next
	Dim regexists, objFS, objShell
	Set objFS = CreateObject("Scripting.FileSystemObject")
	Set objShell = CreateObject( "WScript.Shell" )
	DebugLog 0,"Writing the  '" & regkey & "' in the machine"
	regwrite = objShell.RegWrite(regkey,regvalue , regtype) 
End Function 
'=========================================================================================='==========================================================================================
'=========================================================================================='==========================================================================================

'Create a log file parm type of notification & msg
Sub DebugLog(iType,smsg)
	'iType:'0 = Info'1 = Warning'2 = Error
	on error resume next
	dim oFSO,objShell,oFile
	set oFSO = CreateObject("Scripting.FileSystemObject")
	set objShell = CreateObject("Wscript.shell")
	set oFile = oFSO.openTextFile(InstallationLogDebug,8,TRUE)
	if ucase(smsg) = "BLANK" then 
		oFile.WriteLine 
	else
		oFile.WriteLine now() & vbTAB & iType & vbTAB & smsg
	End If
	oFile.Close
	set oFile = Nothing
	set oFSO = Nothing
	set objShell = Nothing
	err.Clear 
	On Error Goto 0
End Sub
'=========================================================================================='==========================================================================================
'=========================================================================================='==========================================================================================
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
'===========================================================================================================================================================

Function regExists(regKey)
	On Error Resume Next
	regExists = objShell.RegRead(regKey)
	If not isEmpty(regExists) then	
	      regExists=True 
	Else
	      regExists=False
	End If
End Function