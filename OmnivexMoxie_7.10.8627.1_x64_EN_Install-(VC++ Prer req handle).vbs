' This is the standard install script for the / engagement.
' Simply indicate the Package Name, Vendor, and Version below.
' Package Name: OmnivexMoxie_7.10.8627.1_x64_EN
' Package Vendor: Omnivex
' Package Version: 7.10.8627.1
' Generic Prerequisites : Microsoft Visual C++ 2013 Redistributable x64, Microsoft Visual C++ 2015-2019 Redistributable x64 
'##################################################################################################################################
'Defining Variables
On Error Resume Next
Set oShell = CreateObject("WScript.Shell")
set oEnv = oShell.Environment("PROCESS")
Set oFSO = CreateObject("Scripting.FileSystemObject")
strWindir = oShell.ExpandEnvironmentStrings("%Windir%")
strSysDir = oShell.ExpandEnvironmentStrings("%SystemDrive%")
strProgDir = oShell.ExpandEnvironmentStrings("%ProgramFiles(x86)%")
strDir = oShell.ExpandEnvironmentStrings("%programdata%")
strtemp = oShell.ExpandEnvironmentStrings("%temp%")
strPublic = oShell.expandenvironmentstrings("%Public%")
ScriptFullName = wscript.scriptFullname
ScriptFolderName = left(scriptFullName, InStrRev(ScriptFullName, "\"))
oEnv("SEE_MASK_NOZONECHECKS") = 1
'##################################################################################################################################

strLogFolder = strSysDir & "\\"
strLogFolder1 = strSysDir & "\\SCCM\"
If Not(oFSO.FolderExists(strLogFolder)) Then 
oFSO.CreateFolder(strLogFolder)
End If

If Not(oFSO.FolderExists(strLogFolder1)) Then 
oFSO.CreateFolder(strLogFolder1)
End If

'#############Uninstalling Omnivex_Moxie_7.2.6368.1_EN##################

If IsInstalled("{E5840E1F-2F18-4443-BED6-7BEE6ED14EB6}") Then
	
	KillProcess("Omnivex Moxie Player.exe")
	KillProcess("Omnivex Data Server Connection.exe")
	KillProcess("Omnivex Moxie Client Logging Agent.exe")
	
	Ret =  StatusService ("Omnivex Moxie Agent")
	If Ret = "Running" Then  
		subStopService ("Omnivex Moxie Agent")		
	End If
	
	Ret =  StatusService ("Omnivex Moxie Client Logging Agent")
	If Ret = "Running" Then  
		subStopService ("Omnivex Moxie Client Logging Agent")		
	End If
		
	strcmd1 =  "msiexec.exe /x" & "{E5840E1F-2F18-4443-BED6-7BEE6ED14EB6}" & " /qn /l*v " & chr(34) & "%SystemDrive%\\SCCM\Omnivex_Moxie_7.2.6368.1_EN_Uninstall.log" & chr(34)
	rtnVal = oShell.Run(strcmd1,0, True)
	WScript.sleep(5000)
	
	If rtnVal = 0 Or rtnVal = 3010 Then
		If IsInstalled("{D4AD39AD-091E-4D33-BB2B-59F6FCB8ADC3}") then
			strcmd2 =  "msiexec.exe /x" & "{D4AD39AD-091E-4D33-BB2B-59F6FCB8ADC3}" & " /qn /l*v " & chr(34) & "%SystemDrive%\\SCCM\MSFTSQLServerCompact3.5SP2_x64_EN_Uninstall.log" & chr(34)
			rtnVal1 = oShell.Run(strcmd2,0, True)
			WScript.sleep(5000)
		End If
		
		If IsInstalled("{3A9FC03D-C685-4831-94CF-4EDFD3749497}") then
			strcmd3 =  "msiexec.exe /x" & "{3A9FC03D-C685-4831-94CF-4EDFD3749497}" & " /qn /l*v " & chr(34) & "%SystemDrive%\\SCCM\MSFTSQLServerCompact3.5SP2_EN_Uninstall.log" & chr(34)
			rtnVal2 = oShell.Run(strcmd3,0, True)
			WScript.sleep(5000)
		End If
		
		If (oFSO.FileExists(strProgDir & "\Omnivex\Moxie\Agent\Preload\AliasConnection.xml")) Then 
			oFSO.DeleteFile(strProgDir & "\Omnivex\Moxie\Agent\Preload\AliasConnection.xml")
		End if
				
	 End If	 
End If

'##################################################### Prerequisite ###############################################################
'Installing Microsoft Visual C++ 2013 Redistributable 64-bit 12.0.21005.01

If regExists("HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\VisualStudio\12.0\VC\Runtimes\x64\Version") Then

	strKey1 = oShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\VisualStudio\12.0\VC\Runtimes\x64\Version")
	
	if (strKey1 < "v12.0.21005.01") Then
		On Error Resume Next
		strcmd1 = chr(34) & ScriptFolderName & "vcredist2013_x64.exe" & Chr(34) & " /install /quiet /norestart /log " & chr(34) & "%SystemDrive%\\SCCM\Microsoft_VisualC++2013_12.0.21005.01_x64_EN_Install.log" & chr(34)
		'MsgBox (strcmd1)
		rtnVal1 = oShell.Run(strcmd1,0, True)
		WScript.sleep(5000)
	End If

Else
	On Error Resume Next
	strcmd1 = chr(34) & ScriptFolderName & "vcredist2013_x64.exe" & Chr(34) & " /install /quiet /norestart /log " & chr(34) & "%SystemDrive%\\SCCM\Microsoft_VisualC++2013_12.0.21005.01_x64_EN_Install.log" & chr(34)
	rtnVal1 = oShell.Run(strcmd1,0, True)
	WScript.sleep(5000)
End If

'Installing Microsoft Visual C++ 2015-2019 Redistributable 64-bit 14.28.29325.02
If regExists("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\X64\Version") Then

	strKey2 = oShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\X64\Version")
	
	if (strKey2 < "v14.28.29325.02") Then
		On Error Resume Next
		strcmd2 = chr(34) & ScriptFolderName & "vcredist2019_x64.exe" & Chr(34) & " /install /quiet /norestart /log " & chr(34) & "%SystemDrive%\\SCCM\Microsoft_VisualC++2015-2019_14.28.29325.02_x64_EN_Install.log" & chr(34)
		rtnVal2 = oShell.Run(strcmd2,0, True)
		WScript.sleep(5000)
	End If

Else
	On Error Resume Next
	strcmd2 = chr(34) & ScriptFolderName & "vcredist2019_x64.exe" & Chr(34) & " /install /quiet /norestart /log " & chr(34) & "%SystemDrive%\\SCCM\Microsoft_VisualC++2015-2019_14.28.29325.02_x64_EN_Install.log" & chr(34)
	rtnVal2 = oShell.Run(strcmd2,0, True)
	WScript.sleep(5000)
End If

'############################################################ INSTALLATION #################################################################################

strcmd1  = "msiexec.exe /i " & chr(34) & ScriptFolderName & "Omnivex Moxie Setup.msi" & chr(34) & " TRANSFORMS="& chr(34) & ScriptFolderName & "OmnivexMoxie_7.10.8627.1_x64_EN.Mst" & chr(34)  &" /qn /l*v " & chr(34) & "%SystemDrive%\\SCCM\OmnivexMoxie_7.10.8627.1_x64_EN_Install.log" & chr(34)
rtnVal = oShell.Run(strcmd1,0, True)

WScript.sleep(10000)
If rtnVal = 0 Or rtnVal = 3010 Then
		
		Ret =  StatusService ("Omnivex Moxie Agent")
		If Ret = "Running" Then  
			subStopService ("Omnivex Moxie Agent")		
		End If 
	
		If (oFSO.FileExists(strProgDir & "\Omnivex\Moxie\Agent\Preload\AliasConnection.xml")) Then 
			oFSO.DeleteFile(strProgDir & "\Omnivex\Moxie\Agent\Preload\AliasConnection.xml")
		End if
		oFSO.CopyFile ScriptFolderName & "AliasConnection.xml" , strProgDir & "\Omnivex\Moxie\Agent\Preload\AliasConnection.xml", True
		
		wscript.sleep 10000
		Ret2 =  StatusService ("Omnivex Moxie Agent")
		If Ret2 = "Stopped" Then 
	    	subStartService ("Omnivex Moxie Agent")
		End If 
		
	wscript.quit rtnVal
Else
	wscript.quit rtnVal
End If 


'################################################################################################################################## 
Public Function IsInstalled(strGUID)
	On Error Resume Next
	Dim objInstaller
	Dim strProduct, strInstalledProducts
	
	Set objInstaller = WScript.CreateObject("WindowsInstaller.Installer")
	Set strInstalledProducts = objInstaller.Products
	 
IsInstalled = False
	For Each strProduct In strInstalledProducts
		If UCase(strProduct) = UCase(strGUID) Then
			'        If objInstaller.ProductInfo(strProduct, "ProductName") = strGUID Then
			'        If objInstaller.ProductInfo(strProduct, "PackageCode") = strGUID Then
			IsInstalled = True
		End If
	Next
End Function 

Set oShell = Nothing


'****************************************************************************************************************************************************************

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

Public Function regExists(regKey)
	On Error Resume Next
	regExists = oShell.RegRead(regKey)
	If not isEmpty(regExists) then	
	      regExists=True 
	Else
	      regExists=False
	End If
End Function

'=========================================================================================='==========================================================================================
'----------------------------------------------------------------------
'Name	: subStopService
'Purpose: To stop  a service
'Input	: strService: Name of service
'Return	: NA
'----------------------------------------------------------------------
Function subStopService(strService)


	'On Error Resume Next

	Dim strComputer,errReturn
	strComputer = "."

	Dim objWMI
	Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\"& strComputer &"\root\cimv2")

	Dim objCols
	Set objCols = objWMI.ExecQuery("Select * from Win32_Service Where Name = '"& strService & "'")

	Dim objCol

	For Each objCol In objCols
		errReturn = objCol.StopService()
	Next

	subStopService = errReturn
	
	Set objCols = Nothing
	Set objWMI = Nothing

End Function

'----------------------------------------------------------------------
'Name	: subStartService
'Purpose: To start  a service
'Input	: strService: Name of service
'Return	: NA
'----------------------------------------------------------------------
Function subStartService(strService)


	On Error Resume Next

	Dim strComputer,errReturn
	strComputer = "."

	Dim objWMI
	Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\"& strComputer &"\root\cimv2")

	Dim objCols
	Set objCols = objWMI.ExecQuery("Select * from Win32_Service Where Name = '"& strService & "'")

	Dim objCol

	For Each objCol In objCols
		errReturn =	objCol.StartService()
	Next
	
	subStartService = ErrReturn

	Set objCols = Nothing
	Set objWMI = Nothing

End Function  

Function subCopyFile(strsrc, strdes)
	
	On Error Resume Next
	

	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	
	If objFSO.FileExists(strsrc) Then		
 	
 		subCopyFile = objFSO.CopyFile(strsrc, strdes, 0)
 	End If
 	
 	If objFSO.FileExists(strdes) Then		
 	
 		subCopyFile = "0"
 	Else
 		subCopyFile = "1"
 	End If

	Set objFSO = Nothing

End Function

Function StatusService(strService)

	On Error Resume Next


	Dim strComputer,ret
	strComputer = "."

	Dim objWMI
	Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\"& strComputer &"\root\cimv2")

	Dim objCols
	Set objCols = objWMI.ExecQuery("Select * from Win32_Service Where Name = '"& strService & "'")

	Dim objCol

	For Each objCol In objCols

	Ret = objCol.State
		
	Next
	
	If ret = "" Then		
		StatusService = "Service Not Found"
	Else
	  	StatusService = ret
	End If   	
	
	Set objCols = Nothing
	Set objWMI = Nothing

End Function

