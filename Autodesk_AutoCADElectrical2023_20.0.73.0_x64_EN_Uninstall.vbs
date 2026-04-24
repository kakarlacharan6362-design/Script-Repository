' This is the standard Uninstall script for the DXC/Ahold engagement.
' Simply indicate the Package Name, Vendor, and Version below.
' Package Name: Autodesk_AutoCADElectrical2023_20.0.73.0_x64_EN
' Package Vendor: Autodesk
' Package Version: 20.0.73.0
'##################################################################################################################################
'Defining Variables
Dim oFileSys
Set oFileSys = CreateObject("Scripting.FileSystemObject")
set ObjNetwork = CreateObject("WScript.Network")
Set objShell = CreateObject("WScript.Shell")
set oEnv = objShell.Environment("PROCESS")
Set oFS = CreateObject("Scripting.FileSystemObject")
set objFSO = CreateObject("Scripting.FileSystemObject")
strSysDir = objShell.ExpandEnvironmentStrings("%SystemDrive%")
strAppDir = objShell.ExpandEnvironmentStrings("%appdata%")
strPgmDir = objShell.ExpandEnvironmentStrings("%programdata%")
strProgramFiles = objShell.ExpandEnvironmentStrings("%ProgramFiles%")
strPgmDirx86 = objShell.ExpandEnvironmentStrings("%programfiles(x86)%")
ScriptFullName = wscript.scriptFullname
ScriptFolderName = left(scriptFullName, InStrRev(ScriptFullName, "\"))
oEnv("SEE_MASK_NOZONECHECKS") = 1

strLogName = "Autodesk_AutoCADElectrical2023_20.0.73.0_x64_EN"
strLogFolder = strSysDir & "\Windows\Logs\"
If Not(objFSO.FolderExists(strLogFolder)) Then CreateFolder(strLogFolder)
strInstallationLog = strLogFolder & strLogName & "_UnInstallDebug.log"

'***************************************************************************
'Log Header Information STARTS
'***************************************************************************
DebugLog 0,"***************************************************************************************"
DebugLog 0,"This is a Customised Log for UnInstallation of " & strAppName  & " " & strAppVersion
DebugLog 0,"***************************************************************************************"
DebugLog 0, " "
DebugLog 0, strAppVendor & "UnInstall Started - "& Now
DebugLog 0,"UnInstaller = " & strScriptPath & "UnInstallDebug.vbs"
DebugLog 0,"Script Version = 1.0"
DebugLog 0,"User = " & ObjNetwork.UserName
DebugLog 0,"Computer Name = " & ObjNetwork.ComputerName
DebugLog 0,"Log File = " & strInstallationLog
DebugLog 0,"***************************************************************************************"
'***************************************************************************
'Log Header Information ENDS
'***************************************************************************

'##################################################################################################################################
objShell.run "taskkill.exe /T /F /IM AdskIdentityManager.exe" , 0 , True
objShell.run "taskkill.exe /T /F /IM AutodeskDesktopApp.exe" , 0 , True
objShell.run "taskkill.exe /T /F /IM AdskAccessServiceHost.exe" , 0 , True
objShell.run "taskkill.exe /T /F /IM AdSSO.exe" , 0 , True
objShell.run "taskkill.exe /T /F /IM AdskLicensingService.exe" , 0 , True
objShell.run "taskkill.exe /T /F /IM GenuineService.exe" , 0 , True
objShell.run "taskkill.exe /T /F /IM acad.exe" , 0 , True
objShell.run "taskkill.exe /T /F /IM install_manager.exe" , 0 , True
objShell.run "taskkill.exe /T /F /IM AdRefMan.exe" , 0 , True
objShell.run "taskkill.exe /T /F /IM LogAnalyzer.exe" , 0 , True
objShell.run "taskkill.exe /T /F /IM AcSignApply.exe" , 0 , True
objShell.run "taskkill.exe /T /F /IM AcStdBatch.dll" , 0 , True
objShell.run "taskkill.exe /T /F /IM AdskAccessCore.exe" , 0 , True
objShell.run "taskkill.exe /T /F /IM FNPLicensingService64.exe" , 0 , True
objShell.run "taskkill.exe /T /F /IM AceMigrationAssistant.exe" , 0 , True




WScript.sleep(2000)


on Error Resume Next
'******************************************Autodesk Material Library 2023 ***********************************************

strcmd10 = "msiexec.exe /x" & "{8E133591-B0FD-4DB0-B60E-FB593CAF72B0}" & " /qn /l*v " & chr(34) & strSysDir & "\Windows\Logs\Autodesk_MaterialLibrary2023_21.0.1.1_x86_EN_Uninstall.log" & chr(34)

DebugLog 0, "Executing '" & strcmd10 & "'"
rtnVal10 = objShell.Run(strcmd10,0, True)
If rtnVal10 = 0 Or rtnVal10 = 3010 Then	
		
		DebugLog 0, "Uninstalled  the 'Autodesk_MaterialLibrary2023_21.0.1.1_x86_EN' successfully and returned: " & rtnVal10
		DebugLog 0, "Completed Uninstallation Without Errors"

Else
	   DebugLog 0, "Error Uninstalling the Autodesk_MaterialLibrary2023_21.0.1.1_x86_EN and return with: " & rtnVal10
	   DebugLog 0, "Completed Uninstallation With Errors: " & rtnVal10
	   
end If

'******************************************Autodesk_AppManager_3.3.o_x64_EN ***********************************************


strcmd9 = "msiexec.exe /x" & "{4EF1F1D4-E74F-45A8-AF89-95907847D484} " & " /qn /l*v " & chr(34) & strSysDir & "\Windows\Logs\Autodesk_AppManager_3.3.o_x86_EN_Uninstall.log" & chr(34)

DebugLog 0, "Executing '" & strcmd9 & "'"
rtnVal9 = objShell.Run(strcmd9,0, True)
If rtnVal9 = 0 Or rtnVal9 = 3010 Then	
		
		DebugLog 0, "Uninstalled  the 'Autodesk_AppManager_3.3.o_x86_EN_Uninstall' successfully and returned: " & rtnVal9
		DebugLog 0, "Completed Uninstallation Without Errors"

Else
	   DebugLog 0, "Error Uninstalling the Autodesk_AppManager_3.3.o_x86_EN_Uninstall and return with: " & rtnVal9
	   DebugLog 0, "Completed Uninstallation With Errors: " & rtnVal10
	   
end If

'******************************************Autodesk Material Library Base Resolution Image Library 2023 ***********************************************

strcmd8= "msiexec.exe /x" & "{3B564A94-BA47-4E42-ACD6-B5C35291210B}" & " /qn /l*v " & chr(34) & strSysDir & "\Windows\Logs\Autodesk_MaterialLibraryBaseResolutionImageLibrary2023_21.0.1.1_x86_EN_Uninstall.log" & chr(34)

DebugLog 0, "Executing '" & strcmd8 & "'"
rtnVal8 = objShell.Run(strcmd8,0, True)	
If rtnVal8 = 0 Or rtnVal8 = 3010 Then	
		
		DebugLog 0, "Uninstalled  the 'Autodesk_MaterialLibraryBaseResolutionImageLibrary2023_21.0.1.1_x86_EN' successfully and returned: " & rtnVal8
		DebugLog 0, "Completed Uninstallation Without Errors"

Else
	DebugLog 0, "Error Uninstalling the Autodesk_MaterialLibraryBaseResolutionImageLibrary2023_21.0.1.1_x86_EN and return with: " & rtnVal8
	DebugLog 0, "Completed Uninstallation With Errors: " & rtnVal8
	
end If

'******************************************Autodesk  Featured Apps ***********************************************

strcmd7= "msiexec.exe /x" & "{DE8DA5A8-C311-4F2B-B1C3-27A8BC154154}" & " /qn /l*v " & chr(34) & strSysDir & "\Windows\Logs\Autodesk_FeaturedApps_3.3.o _x86_EN_Uninstall.log" & chr(34)

DebugLog 0, "Executing '" & strcmd7 & "'"
rtnVal7 = objShell.Run(strcmd7,0, True)	
If rtnVal7 = 0 Or rtnVal7 = 3010 Then	
		
		DebugLog 0, "Uninstalled  the 'Autodesk_FeaturedApps_3.3.o _x86_EN' successfully and returned: " & rtnVal7
		DebugLog 0, "Completed Uninstallation Without Errors"

Else
	DebugLog 0, "Error Autodesk_FeaturedApps_3.3.o _x86_EN and return with: " & rtnVal7
	DebugLog 0, "Completed Uninstallation With Errors: " & rtnVal7
	
end If

'******************************************Autodesk_SingleSignOnComponent_13.7.7.1807_x64_EN ***********************************************

strcmd6 = "msiexec.exe /x" & "{88003D19-C1C3-402D-A162-42D9B924266C}" & " /qn /l*v " & chr(34) & strSysDir & "\Windows\Logs\Autodesk_SingleSignOnComponent_13.3.3.1803_x64_EN_Uninstall.log" & chr(34)

DebugLog 0, "Executing '" & strcmd6 & "'"
rtnVal6 = objShell.Run(strcmd6,0, True)
If rtnVal6 = 0 Or rtnVal6 = 3010 Then	
		
		DebugLog 0, "Uninstalled  the 'Autodesk_SingleSignOnComponent_13.3.3.1803_x64_EN' successfully and returned: " & rtnVal6
		DebugLog 0, "Completed Uninstallation Without Errors"
		
Else
	   DebugLog 0, "Error Uninstalling the Autodesk_SingleSignOnComponent_13.3.3.1803_x64_EN and return with: " & rtnVal6
	   DebugLog 0, "Completed Uninstallation With Errors: " & rtnVal6
	
end If


'******************************************Autodesk_AutoCADPerformanceFeedbackTool_1.3.12_x86_EN***********************************************

strcmd3 = "msiexec.exe /x" & "{293C8AB2-59FA-4C6E-A707-EE7457D8F567}" & " /qn /l*v " & chr(34) & strSysDir & "\Windows\Logs\Autodesk_AutoCADPerformanceFeedbackTool_1.3.12_x86_EN_Uninstall.log" & chr(34)
DebugLog 0, "Executing '" & strcmd3 & "'"
rtnVal3 = objShell.Run(strcmd3,0, True)

If rtnVal3 = 0 Or rtnVal3 = 3010 Then	
		
			DebugLog 0, "Uninstalled  the 'Autodesk_AutoCADPerformanceFeedbackTool_1.3.12_x86_EN' successfully and returned: " & rtnVal3
			DebugLog 0, "Completed Uninstallation Without Errors"
			
			
			
		Else
			DebugLog 0, "Error Uninstalling the Autodesk_AutoCADPerformanceFeedbackTool_1.3.12_x86_EN and return with: " & rtnVal3
			DebugLog 0, "Completed Uninstallation With Errors: " & rtnVal3

End If

'******************************************AutoCAD Open in Desktop ***********************************************

strcmd5 = "msiexec.exe /x" & "{2B8E195A-0082-4B8F-9284-0FCCB6017C23}" & " /qn /l*v " & chr(34) & strSysDir & "\Windows\Logs\Autodesk_OpeninDesktop_1.o.26.o_x64_EN_Uninstall.log" & chr(34)

DebugLog 0, "Executing '" & strcmd5 & "'"
rtnVal5 = objShell.Run(strcmd5,0, True)
If rtnVal5 = 0 Or rtnVal5 = 3010 Then	
		
		DebugLog 0, "Uninstalled  the 'Autodesk_OpeninDesktop_1.o.26.o_x64_EN' successfully and returned: " & rtnVal5
		DebugLog 0, "Completed Uninstallation Without Errors"
		
Else
	   DebugLog 0, "Error Uninstalling the Autodesk_OpeninDesktop_1.o.26.o_x64_EN and return with: " & rtnVal5
	   DebugLog 0, "Completed Uninstallation With Errors: " & rtnVal5
	
end If

'******************************************Autodesk Save to Web and Mobile ***********************************************

strcmd4 = "msiexec.exe /x" & "{5AB49421-ADA1-4512-9E47-0AE9906F6A28}" & " /qn /l*v " & chr(34) & strSysDir & "\Windows\Logs\Autodesk_SavetoWebandMobile_3.0.30_x64_EN_Uninstall.log" & chr(34)

DebugLog 0, "Executing '" & strcmd4 & "'"
rtnVal4 = objShell.Run(strcmd4,0, True)
If rtnVal4 = 0 Or rtnVal4 = 3010 Then	
		
		DebugLog 0, "Uninstalled  the 'Autodesk_OpeninDesktop_1.o.26.o_x64_EN' successfully and returned: " & rtnVal4
		DebugLog 0, "Completed Uninstallation Without Errors"
		
Else
	   DebugLog 0, "Error Uninstalling the Autodesk_OpeninDesktop_1.o.26.o_x64_EN and return with: " & rtnVal4
	   DebugLog 0, "Completed Uninstallation With Errors: " & rtnVal4
	
end If

If regExists("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{DE663064-1A7C-3129-832C-6D36950C3559}\") Then	
strcmd =chr(34)& strProgramFiles & "\Autodesk\AdODIS\V1\Installer.exe" & chr(34) & " -q -i uninstall --trigger_point system -m " &chr(34)& strPgmDir & "\Autodesk\ODIS\metadata\{DE663064-1A7C-3129-832C-6D36950C3559}\bundleManifest.xml" & chr(34) & " -x " & chr(34) & strPgmDir & "\Autodesk\ODIS\metadata\{DE663064-1A7C-3129-832C-6D36950C3559}\SetupRes\manifest.xsd" & chr(34) & " --extension_manifest " & chr(34)& strPgmDir & "\Autodesk\ODIS\metadata\{DE663064-1A7C-3129-832C-6D36950C3559}\setup_ext.xml" & chr(34) & " --extension_manifest_xsd " & chr(34) & strPgmDir & "\Autodesk\ODIS\metadata\{DE663064-1A7C-3129-832C-6D36950C3559}\SetupRes\manifest_ext.xsd" & chr(34)


DebugLog 0, "Executing '" & strcmd & "'"
rtnVal = objShell.Run(strcmd,0, True)
DebugLog 0, "Execution of  '" & strcmd & "' return with '"  & rtnVal &"'"	
If rtnVal = 0 Or rtnVal = 3010 or rtnVal =1604 Then	
		
		DebugLog 0, "Uninstalled  the 'Autodesk_AutoCADElectrical2023_20.0.73.0_x64_EN' successfully and returned: " & rtnVal
		DebugLog 0, "Completed Uninstallation Without Errors"


	
Else
	    DebugLog 0, "Error Uninstalling the Autodesk_AutoCADElectrical2023_20.0.73.0_x64_EN and return with: " & rtnVal
	    DebugLog 0, "Completed Uninstallation With Errors: " & rtnVal
	
end If

End If

'******************************************Autodesk Genuine Service***********************************************

strcmd1 = "msiexec.exe /x" & "{D207E870-6397-417E-B7DD-720BFBE589A3}" & " REBOOT=ReallySuppress /qn /l*v " & chr(34) & strSysDir & "\Windows\Logs\Autodesk_GenuineServices_7.5.0.226_x86_EN_Uninstall.log" & chr(34)
DebugLog 0, "Executing '" & strcmd1 & "'"
rtnVal1 = objShell.Run(strcmd1,0, True)

If rtnVal1 = 0 Or rtnVal1 = 3010 or rtnVal1 =1604 Then	
		
			DebugLog 0, "Uninstalled  the 'Autodesk_GenuineServices_7.5.0.226_x86_EN' successfully and returned: " & rtnVal1
			DebugLog 0, "Completed Uninstallation Without Errors"
			
			
			
		Else
			DebugLog 0, "Error Uninstalling the Autodesk_GenuineServices and return with: " & rtnVal1
			DebugLog 0, "Completed Uninstallation With Errors: " & rtnVal1

End If

'******************************************Autodesk Identity Manager ***********************************************

strcmd7= chr(34) & strProgramFiles & "\Autodesk\AdskIdentityManager\uninstall.exe" & chr(34) & " --mode unattended"

DebugLog 0, "Executing '" & strcmd7 & "'"
rtnVal7 = objShell.Run(strcmd7,0, True)	
If rtnVal7 = 0 Or rtnVal7 = 3010 Then	
		
		DebugLog 0, "Uninstalled  the 'Autodesk_IdentityManager_1.14.0.3_x64_EN' successfully and returned: " & rtnVal7
		DebugLog 0, "Completed Uninstallation Without Errors"

Else
	DebugLog 0, "Error Uninstalling the Autodesk_IdentityManager_1.14.0.3_x64_EN and return with: " & rtnVal7
	DebugLog 0, "Completed Uninstallation With Errors: " & rtnVal7
	
end If

'******************************************Autodesk Access***********************************************

If regExists("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{A3158B3E-5F28-358A-BF1A-9532D8EBC811}\") Then	
strcmd2 =chr(34)& strProgramFiles & "\Autodesk\AdODIS\V1\Installer.exe" & chr(34) & " -q -i uninstall --trigger_point system -m " &chr(34)& strPgmDir & "\Autodesk\ODIS\metadata\{A3158B3E-5F28-358A-BF1A-9532D8EBC811}\pkg.access.xml" & chr(34) & " -x " & chr(34)& strProgramFiles & "\Autodesk\AdODIS\V1\SetupRes\manifest.xsd" & chr(34) & " --manifest_type package " & chr(34)


DebugLog 0, "Executing '" & strcmd2 & "'"
rtnVal2 = objShell.Run(strcmd2,0, True)
DebugLog 0, "Execution of  '" & strcmd2 & "' return with '"  & rtnVal2 &"'"	
If rtnVal2 = 0 Or rtnVal2 = 3010 Then	
		
		DebugLog 0, "Uninstalled  the 'Autodesk_Access_2.13.2.57x64_EN' successfully and returned: " & rtnVal2
		DebugLog 0, "Completed Uninstallation Without Errors"


	
Else
	    DebugLog 0, "Error Uninstalling the Autodesk_Access_2.13.2.57x64_EN and return with: " & rtnVal2
	    DebugLog 0, "Completed Uninstallation With Errors: " & rtnVal2
	
end If

End If


If oFS.FolderExists(strSysDir & "\Autodesk") Then 
oFS.DeleteFolder strSysDir & "\Autodesk" , True
End If

If oFS.FolderExists(strProgramFiles86 & "\Autodesk") Then 
oFS.DeleteFolder strProgramFiles86 &"\Autodesk" , True
End If

If oFS.FolderExists(strProgramFiles & "\Autodesk") Then 
oFS.DeleteFolder strProgramFiles & "\Autodesk" , True
End If

If oFS.FolderExists(strPgmDir &"\Autodesk") Then 
oFS.DeleteFolder strPgmDir &"\Autodesk", True
End If

If regExists("HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Autodesk\") then
objShell.Regdelete "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Autodesk\", True
End If

If regExists("HKEY_LOCAL_MACHINE\SOFTWARE\Autodesk\") then
objShell.Regdelete "HKEY_LOCAL_MACHINE\SOFTWARE\Autodesk\", True
End If

If Not regExists("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{DE663064-1A7C-3129-832C-6D36950C3559}\") Then	
		
		If regExists("HKEY_LOCAL_MACHINE\SOFTWARE\Ahold\AutoCADElectrical2023\") then
		objShell.Regdelete "HKEY_LOCAL_MACHINE\SOFTWARE\Ahold\AutoCADElectrical2023\"
		End If

End IF


		
'##################################################################################################################################
Set oFS = Nothing
Set objShell = Nothing
'##################################################################################################################################

 Public Function regExists(regKey)
	On Error Resume Next
	regExists = objShell.RegRead(regKey)
	If not isEmpty(regExists) then	
	      regExists=True 
	Else
	      regExists=False
	End If
End Function


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
'##################################################################################################################