'*********************************************************************************************************************************
' Package Name: DellSecureWorks_RedCloak_Uninstallation_2.8_x64_EN_R01
' Package Vendor: DellSecureWorks
' Package Version: 2.8
'*********************************************************************************************************************************
'=========================================================================================
'Global Script Constants
'=========================================================================================
Const EVENT_SUCCESS = 0
Const EVENT_ERROR = 1
Const EVENT_WARNING = 2
Const EVENT_INFORMATION = 4


'==========================================================================================
'Start of Script (MAIN)
'==========================================================================================
Const HKEY_CURRENT_USER   = &H80000001
Const HKEY_LOCAL_MACHINE  = &H80000002
Const HKU_USERS = &H80000003

On Error Resume Next
strComputer = "."
Dim objWMIService, colSoftware, objSoftware, strComputer, strProductCode
Dim oFileSys
Set oFileSys = CreateObject("Scripting.FileSystemObject")
Set oReg = GetObject("winmgmts:!root/default:StdRegProv")
Set objShell = CreateObject("WScript.Shell")
set oEnv = objShell.Environment("PROCESS")
set objFSO = CreateObject("Scripting.FileSystemObject")
set ObjNetwork = CreateObject("WScript.Network")
Set oShell = CreateObject("WScript.Shell")
set oEnv = objShell.Environment("PROCESS")
Set oFS = CreateObject("Scripting.FileSystemObject")
ScriptFullName = wscript.scriptFullname
ScriptFolderName = left(scriptFullName, InStrRev(ScriptFullName, "\"))
strProgramFiles = objShell.ExpandEnvironmentStrings("%ProgramFiles%")
strProgramFiles86 = objShell.ExpandEnvironmentStrings("%ProgramFiles(x86)%")
strSysDir = objShell.ExpandEnvironmentStrings("%SystemDrive%")
strWindir = objShell.ExpandEnvironmentStrings("%Windir%")
oEnv("SEE_MASK_NOZONECHECKS") = 1

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colSoftware = objWMIService.ExecQuery("SELECT * FROM Win32_Product")

'******************************************************************************
'Log naming Constants
'******************************************************************************
strAppVendor = "DellSecureWorks"
strAppName = "RedCloak_Uninstallation"
strAppVersion = "2.8"
strAppLang = "EN"
strAppArch = "x64"
'******************************************************************************
strLogName = "DellSecureWorks_RedCloak_Uninstallation_2.8_x64_EN_R01"
strLogFolder = strWindir & "\Logs\"
If Not(objFSO.FolderExists(strLogFolder)) Then CreateFolder(strLogFolder)
strInstallationLog = strLogFolder & strLogName & "_Install.log"

'***************************************************************************
'Log Header Information STARTS
'***************************************************************************
DebugLog 0,"***************************************************************************************"
DebugLog 0,"This is a Customised Log for Installation of " & strAppName  & " " & strAppVersion
DebugLog 0,"***************************************************************************************"
DebugLog 0, " "
DebugLog 0, strAppVendor & " Install Started - "& Now
DebugLog 0,"Installer = " & strScriptPath & "Install.vbs"
DebugLog 0,"Script Version = 1.0"
DebugLog 0,"User = " & ObjNetwork.UserName
DebugLog 0,"Computer Name = " & ObjNetwork.ComputerName
DebugLog 0,"Log File = " & strInstallationLog
DebugLog 0,"***************************************************************************************"
'***************************************************************************
'Log Header Information ENDS
'***************************************************************************
'Uninstalling DellSecureWorks Ignition  2.8.1.0 EN
If IsInstalled("{4795500B-CE63-49A9-BF42-5343A6E607AF}") Then
	DebugLog 0,"{4795500B-CE63-49A9-BF42-5343A6E607AF} found On the machine."
	DebugLog 0, "Executing the uninstallation of DellSecureWorks Ignition 2.8.1.0 EN"
	strCmdLine = "msiexec.exe /x " & "{4795500B-CE63-49A9-BF42-5343A6E607AF}" & " /qn /l*v " & chr(34) & strLogFolder & "DellSecureWorks_Ignition_2.8.1.0_EN_R01_MSI_Uninstall.log" & Chr(34)
	DebugLog 0, "Executing '" & strCmdLine & "'"

	intRetNo = objShell.Run (strCmdLine, 0, True)
	DebugLog 0, "Execution of  '" & strCmdLine & "' return with '"  & intRetNo &"'"	
	If (intRetNo = 0) Or (intRetNo = 3010) Then
	DebugLog 0,"DellSecureWorks Ignition 2.8.1.0 EN uninstalled successfully. Return Code : " & intRetNo

	    'Uninstalling DellSecureWorks RedCloak  2.8.1.0 EN
		If IsInstalled("{81CAFEDD-311C-43D2-BB55-BE2746D2CD99}") Then
			DebugLog 0,"{81CAFEDD-311C-43D2-BB55-BE2746D2CD99} found On the machine."
			DebugLog 0, "Executing the uninstallation of DellSecureWorks RedCloak  2.8.1.0 EN"
			strCmdLine = chr(34) & strWindir & "\System32\msiexec.exe" & chr(34) & " /x " & chr(34) & "{81CAFEDD-311C-43D2-BB55-BE2746D2CD99}" & Chr(34) & " /qn /l*v " & Chr(34) & strWindir & "\logs\DellSecureWorks_RedCloak_2.8.1.0_EN_R01_Uninstall.log" & Chr(34)
			DebugLog 0,strCmdLine
		
			iReturn= objShell.Run(strCmdLine, 0, true)		
			If (iReturn = 0) Or (iReturn = 3010) Then
			    objShell.RegDelete "HKEY_LOCAL_MACHINE\SYSTEM\\Packages\RedCloak 2.8.1.0\"
				objShell.RegDelete "HKEY_LOCAL_MACHINE\SOFTWARE\RedCloak\Hostel\"
				DebugLog 0,"DellSecureWorks RedCloak  2.8.1.0 EN uninstalled successfully. Return Code : " & iReturn
			Else
				DebugLog 2,"Error occurred while uninstalling DellSecureWorks RedCloak  2.8.1.0 EN. Return Code : " & iReturn
				WScript.quit iReturn
		    End If			
		Else 
			DebugLog 0,"DellSecureWorks RedCloak  2.8.1.0 EN with code {81CAFEDD-311C-43D2-BB55-BE2746D2CD99} Is Not found On the machine."
		End If

   Else
	 DebugLog 2,"Error occurred while uninstalling DellSecureWorks Ignition 2.8.1.0 EN. Return Code : " & intRetNo
 	 WScript.quit intRetNo
    End If			
Else 
	DebugLog 0,"DellSecureWorks Ignition 2.8.1.0 EN with code {4795500B-CE63-49A9-BF42-5343A6E607AF} Is Not found On the machine."
End If 	

'DellSecureWorks RedCloak  2.8.5.0 EN
If IsInstalled("{990A15D3-201D-4B3B-981A-1EA2D0ED3721}") Then
	DebugLog 0,"{990A15D3-201D-4B3B-981A-1EA2D0ED3721} found On the machine."
	DebugLog 0, "Executing the uninstallation of DellSecureWorks RedCloak  2.8.5.0 EN"
	strCmdLine = chr(34) & strWindir & "\System32\msiexec.exe" & chr(34) & " /x " & chr(34) & "{990A15D3-201D-4B3B-981A-1EA2D0ED3721}" & Chr(34) & " /qn /l*v " & Chr(34) & strWindir & "\logs\DellSecureWorks_RedCloak_2.8.5.0_EN_R01_Uninstall.log" & Chr(34)
	DebugLog 0,strCmdLine

	iReturn= objShell.Run(strCmdLine, 0, true)		
	If (iReturn = 0) Or (iReturn = 3010) Then
		DebugLog 0,"DellSecureWorks RedCloak  2.8.5.0 EN uninstalled successfully. Return Code : " & iReturn
	Else
		DebugLog 2,"Error occurred while uninstalling DellSecureWorks RedCloak  2.8.5.0 EN. Return Code : " & iReturn
		WScript.quit iReturn
    End If			
Else 
	DebugLog 0,"DellSecureWorks RedCloak  2.8.5.0 EN with code {990A15D3-201D-4B3B-981A-1EA2D0ED3721} Is Not found On the machine."
End If

objShell.RegWrite "HKEY_LOCAL_MACHINE\SYSTEM\\Packages\DellSecureWorks RedCloak Uninstallation\Uninstall", "01", "REG_SZ"

DebugLog 0, "Writing Detection key"

set objShell = Nothing

WScript.Quit(0)	


'********************************************************************************************************************************
'                                        Below are the functions do not change them
'*********************************************************************************************************************************
								
'------------------------------------------------------------------------
' Name		: IsProcessActive
' Input		: IsProcessActive (Process Name)
' Return	: Boolean value
' Comments	: None
'------------------------------------------------------------------------
Function IsProcessActive(strProName)	
		Dim objProcess
		Set objProcess = GetObject("winmgmts:").execquery("select * from Win32_Process where Name='" & strProName & "' and KernelModeTime>1")
		IsProcessActive = False

		If objProcess.count > 0 Then
			IsProcessActive = True
			exit function
		End If
End Function


'***************************************************************************************	
	'	Function to install or uninstall using command line and silent parameters.
	'	Usage : Inside main use Call Install ("command.exe", "Parameters")
'***************************************************************************************
Function UnInstall(strCmd, strParam)
	'On Error Resume Next
        Dim oShell, ret
	Err.Clear
	DebugLog 0,"Execution started"
	
	Set oShell = createObject("wScript.Shell")
	Err.Clear
	
	'***************************************************************************************	
	'	Remove this comment for below MSGBOX if you want to check what is the command line with arguments
		'msgbox strCmd & strParam
	'***************************************************************************************
	If strParam <> "" Then
		strCmd = strCmd & strParam
	Else
		DebugLog 0,"No parametres were passed for the Exe."
	End If 
	
	DebugLog 0,"Running Command: "& strCmd
	ret=oShell.Run(strCmd,0,True)
	If ret <> "0" AND ret <> "3010" Then
		DebugLog 2,"Problem To Run UnInstallation. ReturnCode = "&ret & ". ErrorCode = "&Err.Number
		DebugLog 2,"More info is available here: "&InstallationLog
		ExitCode ReturnCode
	End If
	Set oShell = Nothing
	Err.Clear
	On Error Goto 0
End Function

'***************************************************************************************	
	'	Sub for Exit Code which will log the Exit code number
'***************************************************************************************

Sub ExitCode (number)
	DebugLog 0,"##Exit! ExitCode= "&number
	Wscript.Quit number
End Sub

'***************************************************************************************	
	'	Sub used to Write values to Log
	'	Usage : Any where in the script part you can call it as below
	'			DebugLog 0,"Simple Log message"
	'			DebugLog 1,"Warning message"
	'			DebugLog 2,"Error message"
'***************************************************************************************
Sub DebugLog(iType,smsg)
	'iType:'0 = Info'1 = Warning'2 = Error
	'on error resume next
	dim oFSO,oShell,oFile
	set oFSO = CreateObject("Scripting.FileSystemObject")
	set oShell = CreateObject("Wscript.shell")
	set oFile = oFSO.openTextFile(strInstallationLog,8,TRUE)
	if ucase(smsg) = "BLANK" then 
		oFile.WriteLine 
	else
		oFile.WriteLine now() & vbTAB & itype & vbTAB & smsg
	End If
	oFile.Close
	set oFile = Nothing
	set oFSO = Nothing
	set oShell = Nothing
	err.Clear 
	On Error Goto 0
End Sub

'***************************************************************************************	
	'	Class for using properties which will set basic Environment variables required for Vista OS
	'   obj.UserName		Will give the current logged in Username
	'   obj.TempDir			Will give the path for User Temp 
	'   obj.UserDomain		Will give the Domain of current logged in User
	'   obj.ComputerName	Will give the Computer name where this script will be executed 
	'   obj.ProgramFiles	Will give the path of Program Files folder (c:\Program Files\)
	'   obj.WinDir			Will give the path of Windows folder (c:\Windows\)
	'   obj.Root			Will give the Root Dir (c:\) in most of the systems
	'	obj.PublicU			Will give the Public Folder (C:\Users\Public)
	' 	USE APPROPRIATE VALUES BY CREATING OBJECT FOR CLASS (ex : 	Set CC = new SysProps 
	'									Rootdir = CC.Root )
'***************************************************************************************

Class SysProps
	' Gets the username for the environment string %username%
	public property get UserName
		dim oShell
		set oShell = CreateObject("Wscript.Shell")
		UserName = oshell.ExpandEnvironmentStrings("%username%")
		set oShell = Nothing
	end property
	' Gets the Temp Directory for the environment string %Temp%
	public property get TempDir
		dim oShell
		set oShell = CreateObject("Wscript.Shell")
		TempDir = oshell.ExpandEnvironmentStrings("%Temp%")
		set oShell = Nothing
	end property
	' Gets the domainnname for the environment string %domainname%
	public Property get UserDomain
		Dim oShell
		Set oShell = createObject("wScript.Shell")
		UserDomain = oshell.ExpandEnvironmentStrings("%userdomain%")
		set oShell = Nothing
	end property
	' Gets the computername for the environment string %computername%
	Public Property get ComputerName
		dim oShell
		set oShell = CreateObject("wscript.shell")
		ComputerName = oshell.ExpandEnvironmentStrings("%computername%")
		set oShell = Nothing
	end Property
	' Gets the Program Files Folder for the environment string %ProgramFiles%

	Public Property get ProgramFiles
		dim oShell, OSType
		set oShell = CreateObject("wscript.shell")
		OSType = ChkOS()
		If (OSType = "32-bit") then
		ProgramFiles = oshell.ExpandEnvironmentStrings("%ProgramFiles%")
		Else
		ProgramFiles = oshell.ExpandEnvironmentStrings("%ProgramFiles(x86)%")
		End If
		ProgramFiles = ProgramFiles & "\"
		set oShell = Nothing
	end Property



	' Gets the Windows Folder for the environment string %windir%
	Public Property get WinDir
		dim oShell
		set oShell = CreateObject("wscript.shell")
		WinDir = oshell.ExpandEnvironmentStrings("%windir%")
		WinDir = WinDir & "\"
		set oShell = Nothing
	end Property
	' Gets the SYSTEM ROOT(c:) drive for the environment string %SystemRoot%
	Public Property get Root
		dim oShell
		set oShell = CreateObject("wscript.shell")
		Root = oshell.ExpandEnvironmentStrings("%SystemDrive%")
		Root = Root & "\"
		set oShell = Nothing
	end Property
	Public Property get PublicU
		dim oShell
		set oShell = CreateObject("wscript.shell")
		Root = oshell.ExpandEnvironmentStrings("%Public%")
		Root = PublicU & "\"
		set oShell = Nothing
	end Property
End Class

Public Function IsInstalled(strGUID)
	
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

Public Function regExists(regKey)

	On Error Resume Next
	regExists = objShell.RegRead(regKey)
	If not isEmpty(regExists) then	
	      regExists=True 
	Else
	      regExists=False
	End If
End Function

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
'*****************************************************************************************

Function CheckArchitecture 
	On Error Resume Next
	
	Dim s
	Dim oShell
	set oShell = CreateObject("Wscript.Shell")
	Err.Clear

	s = oShell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")

	'WriteLog LOG_INFO, "PROCESSOR_ARCHITECTURE = " & s
	If Instr(s,"64") > 0 Then
		CheckArchitecture = 64

	Else
		'In case 64 bit machine - need to check PROCESSOR_ARCHITEW6432 as well since PROCESSOR_ARCHITECTURE will return 32 bit.
		s = ""
		s = oShell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITEW6432%")
		If Instr(s,"%") > 0 Then
			CheckArchitecture = 32	'Default to 32 bit OS architecture

		Else
		'	WriteLog LOG_INFO, "PROCESSOR_ARCHITEW6432 = " & s
			If Instr(s,"64") > 0 Then
				CheckArchitecture = 64

			Else
				CheckArchitecture = 32

			End If
		End If
	End If
	'WriteLog LOG_INFO, "Architecture = " & CheckArchitecture & "-bit"
	
	'WriteLog LOG_INFO, "(CheckArchitecture) Err = " & Err.Number & " : " & Err.Description
End Function
'*****************************************************************************************
Function RemoveStringsInFile(strPath, arrFiles())
Dim objShell, objFS, strSysDir, strFile, i
Set objShell = CreateObject("Wscript.shell")
Set objFS = CreateObject("Scripting.FileSystemObject")
strSysDir = objShell.ExpandEnvironmentStrings("%SystemDrive%")
If (Right(strSysDir, 1) <> "\") Then
strSysDir = strSysDir & "\"
End If
On Error Resume Next
For i = 0 To UBound(arrFiles)
strFile = strSysDir & strPath & arrFiles(i)
objFS.DeleteFile strFile, True
Next
Set objFS = Nothing
Set objShell = Nothing 
End Function
'*****************************************************************************************
Public Function regExists(regKey)
On Error Resume Next
    regExists = objShell.RegRead(regKey)
If not isEmpty(regExists) then
    regExists=True
Else
    regExists=False
End If
End Function
