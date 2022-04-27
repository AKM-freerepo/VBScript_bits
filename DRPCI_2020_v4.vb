'scrubbed to remove proprietary info with *** instead of normal commands
' PCI Re-mediation Install updated for 2020
' Last Modified on 18Feb20
' Added  r******d bios support and i3 r******d
' Added CA*Dopts shuffle
' Modes = RNA, LOCKBIOS, UNLOCKBIOS, S*oSETUP, FIXCA*D, ""
On Error Resume Next
Rebtard = WScript.Arguments.Item(0) 'Should be blank or Mode
Const sTest = 0 '1 for test, 0 for deploy
Const ForReading = 1, ForWriting = 2, ForAppending = 8, HideWindow = 0, ShowWindow = 1
Const ScriptWait = True, ScriptProceed = False, CreateYes = True, CreateNo = False
Const InDir = "c:\s***\install\" 'or installation drop location
Const IHLog = "c:\s***\install\installhistory.log"  'make caps log
Const CIMV2 = "winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2"
Dim oFS, oShell
Dim sOKFile, sPackageName, sPackageLog, S*ompName, sHWType, sStoreNumber, iLaneNumber, S*ountryCode, sLock, sDW, sStatus
Dim i**itFailure, iVerifyCounter, sWMLog1, sWMLog2, sWMLog3, sS**iptN,  sSWLevel, sPCINADM, sHWSpec, sPlatform, sBoardType
i**itFailure = 0
iVerifyCounter = 1
Set oFS = **eateObject("Scripting.FileSystemObject")
Set oShell = **eateObject("WScript.Shell")
S*ompName = oShell.RegRead("HKLM\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName\ComputerName")
sDevice = oShell.RegRead("HKLM\SOFTWARE\N**\S***\CurrentVersion\Load Controller\R**T******l")  'type of device if not by pc name
sPackageName = "DRPCI"
sPackageLog = "c:\s***\install\" & sPackageName & ".log"
sOKFile = "c:\s***\S***Update" & sPackageName & ".OK"
sHWType = UCase(oShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\N**\S***-*******m\ObservedOptions\HWType"))
sHWSpec = UCase(oShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\N**\S***-*******m\ObservedOptions\HWSpec"))
sPlatform = UCase(oShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\N**\Image_Info\Platform"))
sStoreNumber = Mid(S*ompName, 8, 4) 'provided store number makes sense
S*ountryCode = Right(S*ompName, 2)
sS***Release = oShell.RegRead("HKLM\Software\N**\S***\Des**iption\Model")
iLaneNumber = Right(oShell.RegRead("HKLM\Software\N**\S***\CurrentVersion\S***TB\TerminalNumber"), 2)
If iLaneNumber = "" Then
	iLaneNumber = Mid(S*ompName, 4, 2)
End If
sS**iptN = sPackageName & ".vbs"
sSWLevel = RegRead("HKLM\Software\N**\S***\Installation\PatchLevel")
sBoardType = ""
sLiteType1 = ""
sLiteType2 = ""
Call BoardCheck
If IsNull(Rebtard) Then
	Rebtard = ""
End If
If Rebtard = "" Then
	Rebtard = ""
Else	
	Rebtard = UCase(Rebtard)
End If	
'Begin
CopyCat
Select Case Rebtard
	Case "LOCKBIOS"
		Call Bios
		Log "BIOS should now be locked. End.", sPackageLog
		Wscript.Quit 1
	Case "UNLOCKBIOS"
		Call Bios
		Log "BIOS should now be unlocked. End.", sPackageLog
		Wscript.Quit 1
	Case "FIXCA*D"
		Call FixCA*DOpt
		Log "CA*DOpts files should match hardware.  End", sPackageLog
		Wscript.Quit 1
	Case "RNA"
		Call RNA
		Log "User RunNonAdmin should be completed. End.", sPackageLog
		Wscript.Quit 1
	Case "S*oSETUP"
		Call Checks
	Case ""
		Call Checks
	Case Else
		Wscript.Quit 6
End Select		
Sub Checks
	If oFS.FolderExists("c:\s***\install\backup\DRPCI\") = False Then
		Run "cmd /c mkdir c:\s***\install\backup\DRPCI\", 0, True, 1
		Run "cmd /c mkdir c:\s***\install\backup\DRPCI\s***\", 0, True, 1
		Run "cmd /c mkdir c:\s***\install\backup\DRPCI\s***\config\", 0, True, 1
		Run "cmd /c mkdir c:\s***\install\backup\DRPCI\s***\bin\", 0, True, 1
		Run "cmd /c mkdir c:\s***\install\backup\DRPCI\s***\install\", 0, True, 1
	End If	
	sPCINADM = RegRead("HKEY_LOCAL_MACHINE\Software\N**\S***\Installation\PCINonAdmin")
	sPCIPASS = RegRead("HKEY_LOCAL_MACHINE\Software\N**\S***\Installation\PCIPassword")
	If sPCINADM = "" Then
		sPCINADM = "No"
	End If
	If sPCIPASS = "" Then
		sPCIPASS = "No"
	End If	
	If sPCINADM = "Yes" Then
		Log "Noting that Disaster Recovery PCI ran on PCI system.", sPackageLog
		Ws**ipt.Quit 4
	End If
	Select Case sDevice 'should **eate sDevice to:  R*P or S*O
		Case "Yes"'R*P
			sDevice = "R*P"
		Case "No"
			sDevice = "S*O"
	End Select
	DNSVerify 'Now skips lab site domains
	If Right(sStoreNumber, 2) = "99" Then
		If Left(sStoreNumber, 1) = "0" Then
			Log "Leaving testing user in place for lab testing.", sPackageLog
		End If
	Else	
		Run "cmd /c Net user /del testing", 0, True, 2
		Log "Removed testing account information", sPackageLog
	End If	
	Log "Attempting to make changes for Time Zone.", sPackageLog
	If oFS.FileExists("c:\s***\bin\N**TZUtil.exe") = True Then
		Run "cmd /c c:\s***\bin\N**TZUtil.exe", 0, True, 1
		Sleep 5
		Log "Manage Time Zone action attempted.", sPackageLog		
	ElseIf oFS.FileExists("c:\s***\install\N**TZUtil.exe") = True Then
		Run "cmd /c copy c:\s***\install\N**TZUtil.exe c:\s***\bin\N**TZUtil.exe", 0, True, 1
		Run "cmd /c c:\s***\bin\N**TZUtil.exe", 0, True, 1 
		Sleep 5
		Log "Manage Time Zone action attempted.", sPackageLog		
	End If		
	Log "  Starting Installation s**ipt for " & sPackageName, sPackageLog
	If sDevice = "R*P" Then
		Call PCIStart
	ElseIf sDevice = "S*O" Then
		Call PCIStart
	Else
		Log "S**ipt unable to determine device type, exiting s**ipt.", sPackageLog
		HandleOKFile "Error"  
		' DelStart
		WScript.Quit 4
	End If
End Sub
Sub PCIStart()
	Log "All files appear to be present for install to proceed...", sPackageLog
	Call InstallStart 'nonunique
	Call PCIInstall
	If Rebtard = "" Then
		Call RNA
	End If	
	Call BIOSInstall
	Call InstallCleanUp
	Call InstallFinish
	Wscript.Quit 1
End Sub
'=====================================================
'	Install sub routines
'=====================================================
Sub PCIInstall()
	Log "This stage of the update will attempt to change the s*** password and turn on restricted shell.", sPackageLog
	oShell.RegWrite "HKEY_LOCAL_MACHINE\Software\N**\S***\Installation\PCIPassword", "No", "REG_SZ"
	Log "Added registry key for installation.", sPackageLog	
	Call DRWTSN
	Call FixCA*DOpt
	FCopy "c:\s***\install\S*o\s***\bin\SendS***.exe", "c:\s***\bin\SendS***.exe", 1
	FCopy "c:\s***\install\S*o\s***\bin\SendS***U.exe", "c:\s***\bin\SendS***U.exe", 1
	FCopy "c:\s***\install\S*o\s***\bin\GetR*pNetLanes.sh", "c:\s***\bin\GetR*pNetLanes.sh", 1
	FCopy "c:\s***\install\S*o\s***\bin\rr_tx_query.sh", "c:\s***\bin\rr_tx_query.sh", 1
	FCopy "c:\s***\install\S*o\s***\bin\CompareR*pLanes.sh", "c:\s***\bin\CompareR*pLanes.sh", 1
	FCopy "c:\s***\install\S*o\s***\bin\Autologon.exe", "c:\s***\bin\Autologon.exe", 1
	FCopy "c:\s***\install\S*o\s***\bin\ChangePwd.exe", "c:\s***\bin\ChangePwd.exe", 1
	FCopy "c:\s***\install\S*o\s***\bin\checkS***LogonEvents.sh", "c:\s***\bin\checkS***LogonEvents.sh", 1
	FCopy "c:\s***\install\S*o\s***\bin\cygrunsrv_Rlp.sh", "c:\s***\bin\cygrunsrv_Rlp.sh", 1
	FCopy "c:\s***\install\S*o\s***\bin\ExplorerNo-LaunchpadYes-WMT.bat", "c:\s***\bin\ExplorerNo-LaunchpadYes-WMT.bat", 1
	FCopy "c:\s***\install\S*o\s***\bin\LogonS**ipt.sh", "c:\s***\bin\LogonS**ipt.sh", 1
	FCopy "c:\s***\install\S*o\s***\bin\NetUserGetInfo.exe", "c:\s***\bin\NetUserGetInfo.exe", 1
	FCopy "c:\s***\install\S*o\s***\bin\RlpNetUse.sh", "c:\s***\bin\RlpNetUse.sh", 1
	FCopy "c:\s***\install\S*o\s***\bin\RlpToggleDebug.sh", "c:\s***\bin\RlpToggleDebug.sh", 1
	FCopy "c:\s***\install\S*o\s***\bin\WMLogonJP.exe", "c:\s***\bin\WMLogonJP.exe", 1
	FCopy "c:\s***\install\S*o\s***\bin\WMLogon.exe", "c:\s***\bin\WMLogon.exe", 1
	FCopy "c:\s***\install\S*o\s***\bin\WMLogon_change_S*O.bat", "c:\s***\bin\WMLogon_change_S*O.bat	", 1
	FCopy "c:\s***\install\S*o\s***\bin\WMLogon_change_S*O.EXE", "c:\s***\bin\WMLogon_change_S*O.EXE", 1
	FCopy "c:\s***\install\S*o\s***\bin\WMLogon_deploy_S*O.bat", "c:\s***\bin\WMLogon_deploy_S*O.bat", 1
	FCopy "c:\s***\install\S*o\s***\bin\WMLogon_deploy_S*O.EXE", "c:\s***\bin\WMLogon_deploy_S*O.EXE", 1
	Log "Copied S*o bin files.", sPackageLog
	FCopy "c:\s***\install\S*o\s***\config\rlpConfig.rc", "c:\s***\config\rlpConfig.rc", 0
	Run "cmd /c copy c:\s***\install\S*o\s***\config\*.enc c:\s***\config\", 0, True, 1
	Log "Copied S*o config files.", sPackageLog
	FCopy "c:\s***\install\S*o\s***\cygwin\bin\ukill.exe", "c:\s***\cygwin\bin\", 0
	Log "Copied S*o cygwin files.", sPackageLog
	If S*ountryCode = "JP" Then
		FCopy "c:\s***\install\S*o\s***\config\C*DDOpts.JP", "c:\s***\config\C*DDOpts.000", 0	
		FCopy "c:\s***\install\S*o\s***\bin\WMLogonJP.exe", "c:\s***\bin\WMLogon.exe", 1
		FCopy "c:\s***\install\S*o\s***\config\JPpwMcxAdmin.enc", "c:\s***\config\pwM*xAdmin.enc", 0
		FCopy "c:\s***\install\S*o\s***\config\JPpwMcxS***.enc", "c:\s***\config\pwM*xS***.enc", 0
		FCopy "c:\s***\install\S*o\s***\config\JPpwMcxS***_Alt.enc", "c:\s***\config\pwM*xS***_Alt.enc", 0
		FCopy "c:\s***\install\S*o\s***\config\JPpwMcxS***_Old.enc", "c:\s***\config\pwM*xS***_Old.enc", 0
		FCopy "c:\s***\install\S*o\s***\config\JPpwS*oAdmin.enc", "c:\s***\config\pwS*oAdmin.enc", 0
		FCopy "c:\s***\install\S*o\s***\config\JPpwS*oS***.enc", "c:\s***\config\pwS*oS***.enc", 0
		FCopy "c:\s***\install\S*o\s***\config\JPpwS*oS***_Alt.enc", "c:\s***\config\pwS*oS***_Alt.enc", 0
		FCopy "c:\s***\install\S*o\s***\config\JPpwS*oS***_Old.enc", "c:\s***\config\pwS*oS***_Old.enc", 0
		Log "Finished emplacing Japan files.", sPackageLog	
	End If		
	Log "All files should now be copied for installation to proceed", sPackageLog	
	'Add in Group Policy changes
	Run "cmd /c Net Accounts /MINPWLENGTH:0", 0, True, 1
	Run "cmd /c NET ACCOUNTS /MINPWAGE:0", 0, True, 1
	Run "cmd /c NET ACCOUNTS /MAXPWAGE:UNLIMITED", 0, True, 1
	Run "cmd /c c:\s***\bin\WMLogon_change_S*o.bat /s", 0, True, 1
	Sleep 5
	Log "Attempted password change to a new complex version.  Warning Password no longer known.", sPackageLog
	Sleep 5
	Set oFile = oFS.OpenTextFile("c:\s***\logs\WMLogon.log", ForReading, False)
	sWMLog1 = oFile.ReadAll
	oFile.Close
	Run "cmd /c copy c:\s***\install\S*o\s***\config\*.enc c:\s***\config\", 0, True, 1
	If S*ountryCode = "JP" Then
		FCopy "c:\s***\install\S*o\s***\config\CA*DOpts.JP", "c:\s***\config\C*DDOpts.000", 0	
		FCopy "c:\s***\install\S*o\s***\bin\WMLogonJP.exe", "c:\s***\bin\WMLogon.exe", 0
		FCopy "c:\s***\install\S*o\s***\config\JPpwMcxAdmin.enc", "c:\s***\config\pwM*xAdmin.enc", 0
		FCopy "c:\s***\install\S*o\s***\config\JPpwMcxS***.enc", "c:\s***\config\pwM*xS***.enc", 0
		FCopy "c:\s***\install\S*o\s***\config\JPpwMcxS***_Alt.enc", "c:\s***\config\pwM*xS***_Alt.enc", 0
		FCopy "c:\s***\install\S*o\s***\config\JPpwMcxS***_Old.enc", "c:\s***\config\pwM*xS***_Old.enc", 0
		FCopy "c:\s***\install\S*o\s***\config\JPpwS*oAdmin.enc", "c:\s***\config\pwS*oAdmin.enc", 0
		FCopy "c:\s***\install\S*o\s***\config\JPpwS*oS***.enc", "c:\s***\config\pwS*oS***.enc", 0
		FCopy "c:\s***\install\S*o\s***\config\JPpwS*oS***_Alt.enc", "c:\s***\config\pwS*oS***_Alt.enc", 0
		FCopy "c:\s***\install\S*o\s***\config\JPpwS*oS***_Old.enc", "c:\s***\config\pwS*oS***_Old.enc", 0
	End If		
	If Instr(sWMLog1, "CHANGEPASSWORD: EXIT WITH STATUS CODE: 0") > 0 Then
		oShell.RegWrite "HKEY_LOCAL_MACHINE\Software\N**\S***\Installation\PCIPassword", "Yes", "REG_SZ"
		oShell.RegWrite "HKEY_LOCAL_MACHINE\Software\N**\S***\Installation\PCIRestricted", "No", "REG_SZ"
		Run "cmd /c c:\s***\bin\WMLogon_deploy_S*o.bat /s", 0, True, 1
		Log "Ran Deployment s**ipt after 0 code.", sPackageLog
	ElseIf Instr(sWMLog1, "CHANGEPASSWORD: EXIT WITH STATUS CODE: -1") > 0 Then		
		oShell.RegWrite "HKEY_LOCAL_MACHINE\Software\N**\S***\Installation\PCIPassword", "Yes", "REG_SZ"
		oShell.RegWrite "HKEY_LOCAL_MACHINE\Software\N**\S***\Installation\PCIRestricted", "No", "REG_SZ"
		Run "cmd /c c:\s***\bin\WMLogon_deploy_S*o.bat /s", 0, True, 1
		Log "Ran Deployment script after -1 code.", sPackageLog	
	Else
		Log "Error during install, will require intervention.", sPackageLog
		HandleOkFile "Error"
		Log "S**ipt unable to enable Restricted Shell.  Exiting s**ipt.", sPackageLog
		Ws**ipt.Quit 2 
	End If	
	Sleep 5	
	FCopy "c:\s***\install\backup\DRPCI\s***\bin\SendS***.exe", "c:\s***\bin\SendS***.exe", 0
	FCopy "c:\s***\install\backup\DRPCI\s***\bin\SendS***U.exe", "c:\s***\bin\SendS***U.exe", 0
	Log "Restricted Shell now activated.", sPackageLog
	'allowance for new logons**ipt
	FMove "c:\s***\bin\LogonS**ipt.sh", "c:\s***\bin\LogonS**ipt.bash", 0
	FCopy "c:\s***\install\S*o\s***\bin\LogonS**ipt.exe", "c:\s***\bin\LogonS**ipt.exe", 0
	'FCopy "c:\s***\install\S*o\s***\bin\S***WinShell-WMT-PowerShell.reg", "c:\s***\bin\S***WinShell-WMT-PowerShell.reg", 0
	FCopy "c:\s***\install\S*o\s***\config\LogonS**ipt.ini", "c:\s***\config\LogonS**ipt.ini", 0
	FCopy "c:\s***\install\S*o\s***\bin\ShellReadyEvent.exe", "c:\s***\bin\ShellReadyEvent.exe", 0
	Run "cmd /c REGEDIT /S c:\s***\bin\S***WinShell-WMT-PowerShell.reg", 0, True, 1
	Run "cmd /c PowerShell -Command " & Chr(34) & "& {New-EventLog -Source S***LogonWmt -LogName Application}" & Chr(34), 0, True, 1
	Del "c:\users\s***\AppData\Roaming\Mi**osoft\Windows\" & Chr(34) & "Start Menu" & Chr(34) & "\Programs\Startup\S*oSetup.lnk"
	FCopy "c:\s***\bin\S***WinShell-WMT-PowerShell.reg", "c:\s***\bin\S***WinShell-WMT-Wait.reg", 0
	Log "LogonS**ipt actions attempted", sPackageLog
	FCopy "c:\s***\install\S*o\n**ra.exe", "c:\s***\bin\n**ra.exe", 1
	Run "cmd /c ""C:\WINDOWS\Mi**osoft.NET\Framework\v2.0.50727\InstallUtil c:\s***\bin\n**ra.exe""", 0, True, 1
	Run "cmd /c sc start n**ra_service", 0, True, 1
	Log "N**RA installed on this device", sPackageLog	
	Run "cmd /c REGEDIT /S c:\s***\install\S*o\StickyKey.reg", 0, True, 1
	If S*ountryCode = "JP" Then
		FCopy "c:\s***\config\CA*DOpts.JP", "c:\s***\config\CA*DOpts.000", 0	
	End If
	TestLog "AllLaneS*ommon test"
	If sDevice = "RAP" Then
		Set oFile = oFS.OpenTextFile("c:\install\S*oSetup.log", ForReading, True)
		sSSL = oFile.ReadAll
		oFile.Close
		If Instr(sSSL, "End Enable Bag Scale on RAP procedure") > 0 Then 
			Select Case S*ountryCode
				Case "UK"
					FCopy "c:\s***\install\S*o\s***\config\UK_CEALC_3.xml", "c:\s***\config\ConfigEntity-AllLaneS*ommon.xml", 0
					Log "Over-wrote UK ConfigEntity-AllLaneS*ommon.xml after S*oSetup.", sPackageLog
				Case "JP"
					FCopy "c:\s***\install\S*o\s***\config\JP_CEALC_3.xml", "c:\s***\config\ConfigEntity-AllLaneS*ommon.xml", 0
					Log "Over-wrote JP ConfigEntity-AllLaneS*ommon.xml after S*oSetup.", sPackageLog
			End Select
		End If
	End If		
	TestLog "Passed the point of AllLaneS*ommon"		
	Set oFile = oFS.OpenTextFile("c:\s***\logs\WMLogon.log", ForReading, False)
	sWMLog2 = oFile.ReadAll
	oFile.Close
	If Instr(sWMLog2, "DEPLOYMENT: EXIT WITH STATUS CODE: 0") > 0 Then 
		Log "Preparing to update registry after 0 code.", sPackageLog
		oShell.RegWrite "HKEY_LOCAL_MACHINE\Software\N**\S***\Installation\PCIRestricted", "Yes", "REG_SZ"
		Log "Registry updated to include PCI key", sPackageLog	
	ElseIf Instr(sWMLog2, "DEPLOYMENT: EXIT WITH STATUS CODE: -1") > 0 Then
		Log "Preparing to update registry after -1 code.", sPackageLog
		oShell.RegWrite "HKEY_LOCAL_MACHINE\Software\N**\S***\Installation\PCIRestricted", "Yes", "REG_SZ"
		Log "Registry updated to include PCI key", sPackageLog	
	Else
		Log "Error during install, will require intervention.", sPackageLog
		HandleOkFile "Error"
		Log "Error during install, will require intervention.", sPackageLog
		Ws**ipt.Quit 2 
	End If
	If sDW = 1 Then
		FMove "c:\s***\config\DRExclude.drpci.bak", "c:\s***\config\DRExclude.ini", 0
		FMove "c:\s***\config\DRNoBoot.drpci.bak", "c:\s***\config\DRNoBoot.ini", 0
	End If
End Sub
Sub RNA() 
	Call DRWTSN
	Log "This stage of the update will attempt to demote s*** from Admin group.", sPackageLog
	oShell.RegWrite "HKEY_LOCAL_MACHINE\Software\N**\S***\Installation\PCINonAdmin", "No", "REG_SZ"
	TestLog "PCI Non-admin showing as: " & sPCINADM
	Run "cmd /c ntrights -u s*** +r SeNetworkLogonRight", 0, True, 1 
	FCopy "c:\s***\install\S*o\RunNonAdmin.bat", "c:\s***\bin\RunNonAdmin.bat", 1
	Sleep 3
	sPCINADM = RegRead("HKEY_LOCAL_MACHINE\Software\N**\S***\Installation\PCINonAdmin")
	Run "cmd /c c:\s***\bin\RunNonAdmin.bat s*** s*** /NOLOGIN /S", 0, True, 1
	iCounter = 0
	Do 
		Sleep 300
		sPCINADM = RegRead("HKEY_LOCAL_MACHINE\Software\N**\S***\Installation\PCINonAdmin")
		iCounter = iCounter + 1		
	Loop Until sPCINADM = "Yes" Or iCounter = 3
	TestLog "iCounter = " & iCounter
	TestLog "PCI Non-admin now showing as: " & sPCINADM 
	Log " Sleeping for a bit after running Non-Admin...", sPackageLog
	Sleep 20
	Run "cmd /c ntrights -u s*** +r SeNetworkLogonRight", 0, True, 1
	If S*ountryCode = "UK" Then
		Run "cmd /c sends*** -FilePermissions s*** " & Chr(34) & "c:\Proxy" & Chr(34) & " grant all", 0, True, 1
	End If
	If sDW = 1 Then
		FMove "c:\s***\config\DRExclude.drpci.bak", "c:\s***\config\DRExclude.ini", 0
		FMove "c:\s***\config\DRNoBoot.drpci.bak", "c:\s***\config\DRNoBoot.ini", 0
	End If 
End Sub	
Sub InstallCleanUp()
	' DelStart
	Log "Cleaning up temporary files from installation.", sPackageLog
	Run "cmd /c del c:\s***\config\*.enc /s", 0, True, 1
	Run "cmd /c del c:\s***\config\*.unenc /s", 0, True, 1
	Log "Deleted en**ypted files.", sPackageLog
	Log "End of DRPCI run.", sPackageLog
End Sub
Sub BIOSInstall()
	If oFS.FileExists("c:\s***\install\POSTDRPCI.vbs") = True Then
		Run "cmd /c c:\s***\install\POSTDRPCI.vbs", 0, True, 1 'added in post pci fix for datapump/S*ogoal
		Sleep 60  'end post pci	
	End If	
	Do While oFS.FileExists("c:\s***\firmware\terminal\UpdateBIOS.exe") = False
		Run "cmd /c ""cd c:\s***\install & msiexec /package UpdateBIOSsetup.msi /quiet""", 0, True, 1
		Sleep 120
		Log "Attempted to unpack UpdateBIOS", sPackageLog
	Loop
	If S*ountryCode = "JP" And sDevice = "S*o" Then
		FCopy "c:\s***\config\CA*DOpts.JP", "c:\s***\config\CA*DOpts.000", 0
	End If		
	Run "cmd /c net stop tsf", 0, True, 1
	Run "cmd /c sc config tsf start= auto", 0, True, 1
	Log "Set up files for BIOS Update", sPackageLog
	Log "**eating Bios Update registry entry.", sPackageLog
	oShell.RegWrite "HKEY_LOCAL_MACHINE\Software\N**\S***\Installation\PCIBiosUpdate", "No", "REG_SZ"
	Log "Preparing to run BIOS Update.", sPackageLog
	Call Bios
	Sleep 10
End Sub
Sub FixCA*DOpt
	Select Case S*ountryCode
		Case "JP"
			If oFS.FileExists("c:\s***\config\CA*DOpts.JP") = False Then
				FCopy "c:\s***\config\CA*DOpts.000", "c:\s***\config\CA*DOpts.JP", 0	
			End If	
		Case "UK"
			If sPlatform = "R6C" Then
				FCopy "c:\s***\config\CA*Dopts.0R6", "c:\s***\config\CA*DOpts.000", 0
			End If
			If sPlatform = "R6N" Then
				FCopy "c:\s***\config\CA*Dopts.0RN", "c:\s***\config\CA*DOpts.000", 0
			End If	
			If sPlatform = "SS90" Then
				FCopy "c:\s***\config\CA*Dopts.ss90", "c:\s***\config\CA*DOpts.000", 0
			End If
		Case "US"
			If sPlatform = "SS90" Then
				FCopy "c:\s***\config\CA*Dopts.00.ss90", "c:\s***\config\CA*Dopts.000", 0
			End If	
			If InStr(sPlatform, "PLUS") > 1 Then
				If InStr(sPlatform, "NARROW") > 0 Then
					FCopy "c:\s***\config\CA*Dopts.00.narrow", "c:\s***\config\CA*Dopts.000", 0
				Else		
					Call CheckLite
					If sLiteType1 = "BNR" Then
						If sLiteType2 = "s***5" Then
							FCopy "c:\s***\config\CA*Dopts.00.r6lite", "c:\s***\config\CA*Dopts.000", 0
						ElseIf sLiteType2 = "s***6" Then
							FCopy "c:\s***\config\CA*Dopts.00.r6lmp", "c:\s***\config\CA*Dopts.000", 0	
						End If
					ElseIf sLiteType1 = "GSR" Then
						If sLiteType2 = "s***5" Then
							FCopy "c:\s***\config\CA*Dopts.00.r6lmm", "c:\s***\config\CA*Dopts.000", 0
						ElseIf sLiteType2 = "s***6" Then
							FCopy "c:\s***\config\CA*Dopts.00.r6pplus", "c:\s***\config\CA*Dopts.000", 0 
						End If
					End If
				End If
			End If		
			If InStr(sPlatform, "NARROW") > 0 Then
				FCopy "c:\s***\config\CA*Dopts.00.narrow", "c:\s***\config\CA*Dopts.000", 0
			End If		
	End Select
End Sub
Sub CheckLite
	FileName = "c:\s***\install\litecheck.txt"
	Run "cmd /c c:\s***\bin\EnumUsb.exe /c c:\s***\config\EnumUsb.ini /o " & FileName, 0, True, 5
	Set mefile = oFS.OpenTextFile(FileName, ForReading, False)
	sLiteType = mefile.ReadAll
	mefile.Close
	If InStr(sLiteType, "BNR") > 0 Then
		sLiteType1 = "BNR"
	ElseIf InStr(sLiteType, "ATM") > 0 Then
		sLiteType1 = "GSR"
	Else
		Log "Failure determining Dispenser type", sPackageLog
	End If
	TestLog "Showing dispenser type as:  " & sLiteType1 & "!"
	If InStr(sLiteType, "K590") > 0 Then
		sLiteType2 = "s***5"
	ElseIf InStr(sLiteType, "NHPI") > 0 Then
		sLiteType2 = "s***6"
	Else
		Log "Failure determining Printer type", sPackageLog
	End If	
	TestLog "Showing printer type as:  " & sLiteType2 & "!"
End Sub
'######################################################################################################################
'	                       Everything below this point must remain package non-specific
'######################################################################################################################
Sub DRWTSN
	sDW = 0
	If oFS.FileExists("c:\s***\config\DRExclude.ini") = True Then
		Set oFile = oFS.OpenTextFile("c:\s***\config\DRExclude.ini", ForReading, True)
		sIni = oFile.ReadAll
		oFile.Close
		If Instr(sIni, "SendS***.exe") > 0 Then 
			If Instr(sIni, "SSPSWOSUtil.exe") > 0 Then
				sDW = 0
			Else
				sDW = 1
			End If
		Else 
			sDW = 1
		End If
	Else
		sDW = 1
	End If
	If sDW = 1 Then
		FCopy "c:\s***\config\DRExclude.ini", "c:\s***\config\DRExclude.drpci.bak", 0
		FCopy "c:\s***\config\DRNoBoot.ini", "c:\s***\config\DRNoBoot.drpci.bak", 0
		Set eFile = oFS.OpenTextFile("c:\s***\config\DRExclude.ini", ForAppending, True)
		eFile.WriteLine "SS*oUI.exe"
		eFile.WriteLine "SendS***.exe"
		eFile.WriteLine "SendS***U.exe"
		eFile.WriteLine "SS**WOSUtil.exe"
		eFile.Close
		Set fFile = oFS.OpenTextFile("c:\s***\config\DRNoBoot.ini", ForAppending, True)
		fFile.WriteLine "SS*oUI.exe"
		fFile.WriteLine "SendS***.exe"
		fFile.WriteLine "SendS***U.exe"
		fFile.WriteLine "SS**WOSUtil.exe"
		fFile.Close	
	End If	
End Sub
Sub Bios
	If Rebtard = "UNLOCKBIOS" Then
		sLock = "u"
	Else
		sLock = "l"
	End If
	Select Case sBoardType
		Case "P*C*N*"
			FCopy "c:\s***\install\S*o\Bios\p****o_" & sLock & ".rom", "c:\s***\bin\UpdateBIOS\afuwin32\po**no_" & sLock & ".rom", 0
			Run "cmd /c " & Chr(34) & "cd c:\s***\bin\UpdateBIOS\afuwin32\ & afuwin.exe p**ono_" & sLock & ".rom /N /R" & Chr(34), 0, True, 1
			Sleep 30	
		Case "T*L*A*E*A"
			FCopy "c:\s***\install\S*o\Bios\ta*****ega_" & sLock & ".map", "c:\s***\bin\UpdateBIOS\t*****ga_" & sLock & ".map", 0
			Run "cmd /c " & Chr(34) & " cd c:\s***\bin\UpdateBIOS\ & WinSetCMOS_SA.exe L:tall**ega_" & sLock & ".map /S" & Chr(34), 0, True, 1
			Sleep 30			
		Case "B******L"
			FCopy "c:\s***\install\S*o\Bios\br***ol_" & sLock & ".map", "c:\s***\bin\UpdateBIOS\br***ol_" & sLock & ".map", 0
			Run "cmd /c " & Chr(34) & "cd c:\s***\bin\UpdateBIOS\ & WinSetCMOS_SA.exe L:br***ol_" & sLock & ".map /S" & Chr(34), 0, True, 1
			Sleep 30			
		Case "BR****L_II"
			b2up = RegRead("HKEY_LOCAL_MACHINE\Software\N**\Image_Info\BIOS")
			If b2up = "" Then
				FCopy "c:\s***\install\S*o\Bios\bios_q67.rom", "c:\s***\bin\UpdateBIOS\afuwin32\bios_q67.rom", 0	
				Run "cmd /c " & Chr(34) & "cd c:\s***\bin\UpdateBIOS\afuwin32\ & afuwin.exe bios_q67.rom /P /B /N /R /X /Q" & Chr(34), 0, True, 1
				Sleep 60
				oShell.RegWrite "HKEY_LOCAL_MACHINE\Software\N**\Image_Info\BIOS", "8.6.6.0", "REG_SZ"
			End If
			FCopy "c:\s***\install\S*o\Bios\br****l2_" & sLock & ".rom", "c:\s***\bin\UpdateBIOS\afuwin32\br****l2_" & sLock & ".rom", 0
			Run "cmd /c " & Chr(34) & "cd c:\s***\bin\UpdateBIOS\afuwin32\ & afuwin.exe br****l2_" & sLock & ".rom /N /R" & Chr(34), 0, True, 1
			Sleep 30	
		Case "M*NACO"
			b2up = RegRead("HKEY_LOCAL_MACHINE\Software\N**\Image_Info\BIOS")
			If b2up = "" Then
				b3up = RegRead("HKEY_LOCAL_MACHINE\HARDWARE\DES**IPTION\System\BIOS\BIOSReleaseDate")
				If InStr(b3up, "03/02/2018") < 1 Then
					FCopy "c:\s***\install\S*o\Bios\bios_q87.rom", "c:\s***\bin\UpdateBIOS\afuwin32\bios_q87.rom", 0	
					Run "cmd /c " & Chr(34) & "cd c:\s***\bin\UpdateBIOS\afuwin32\ & afuwin.exe bios_q87.rom /P /B /N /R /X /Q" & Chr(34), 0, True, 1
					Sleep 60
				End If	
				oShell.RegWrite "HKEY_LOCAL_MACHINE\Software\N**\Image_Info\BIOS", "9.1.3.1", "REG_SZ"
			End If		
			If sPlatform = "R6L" Then
				FCopy "c:\s***\install\S*o\Bios\m*naco_r6lite_" & sLock & ".rom", "c:\s***\bin\UpdateBIOS\afuwin32\m*naco_r6lite_" & sLock & ".rom", 0
				Run "cmd /c " & Chr(34) & "cd c:\s***\bin\UpdateBIOS\afuwin32\ & afuwin.exe m*naco_r6lite_" & sLock & ".rom /N /R" & Chr(34), 0, True, 1
				Sleep 30
			End If	
			If sPlatform = "R6LP" Then
				FCopy "c:\s***\install\S*o\Bios\m*naco_r6plus_" & sLock & ".rom", "c:\s***\bin\UpdateBIOS\afuwin32\m*naco_r6plus_" & sLock & ".rom", 0
				Run "cmd /c " & Chr(34) & "cd c:\s***\bin\UpdateBIOS\afuwin32\ & afuwin.exe m*naco_r6plus_" & sLock & ".rom /N /R" & Chr(34), 0, True, 1
				Sleep 30
			End If			
			If sPlatform = "SS90" Then
				FCopy "c:\s***\install\S*o\Bios\m*naco_ss90_" & sLock & ".rom", "c:\s***\bin\UpdateBIOS\afuwin32\m*naco_ss90_" & sLock & ".rom", 0
				Run "cmd /c " & Chr(34) & "cd c:\s***\bin\UpdateBIOS\afuwin32\ & afuwin.exe m*naco_ss90_" & sLock & ".rom /N /R" & Chr(34), 0, True, 1
				Sleep 30
			End If	
			If sHWSpec = "7702" Then
				FCopy "c:\s***\install\S*o\Bios\m*naco_r6rap_" & sLock & ".rom", "c:\s***\bin\UpdateBIOS\afuwin32\m*naco_r6rap_" & sLock & ".rom", 0
				Run "cmd /c " & Chr(34) & "cd c:\s***\bin\UpdateBIOS\afuwin32\ & afuwin.exe m*naco_r6rap_" & sLock & ".rom /N /R" & Chr(34), 0, True, 1
				Sleep 30	
			End If
			If sPlatform = "R6" Then
				FCopy "c:\s***\install\S*o\Bios\m*naco_r6_" & sLock & ".rom", "c:\s***\bin\UpdateBIOS\afuwin32\m*naco_r6_" & sLock & ".rom", 0
				Run "cmd /c " & Chr(34) & "cd c:\s***\bin\UpdateBIOS\afuwin32\ & afuwin.exe m*naco_r6_" & sLock & ".rom /N /R" & Chr(34), 0, True, 1
				Sleep 30	
			End If	
		Case "R*CHMOND"
			If S*ountryCode = "US" Then
				If sHWSpec = "NARROW" Then
					FCopy "c:\s***\install\S*o\Bios\r*chmond_r6lnplus_" & sLock & ".rom", "c:\s***\bin\UpdateBIOS\afuwin32\r*chmond_r6lnplus_" & sLock & ".rom", 0
					Run "cmd /c " & Chr(34) & "cd c:\s***\bin\UpdateBIOS\afuwin32\ & afuwin.exe r*chmond_r6lnplus_" & sLock & ".rom /N /R" & Chr(34), 0, True, 1
					Sleep 30
				End If				
				If sHWSpec = "LITE_PLUS" Then
					FCopy "c:\s***\install\S*o\Bios\r*chmond_r6pplus_" & sLock & ".rom", "c:\s***\bin\UpdateBIOS\afuwin32\r*chmond_r6pplus_" & sLock & ".rom", 0
					Run "cmd /c " & Chr(34) & "cd c:\s***\bin\UpdateBIOS\afuwin32\ & afuwin.exe r*chmond_r6pplus_" & sLock & ".rom /N /R" & Chr(34), 0, True, 1
					Sleep 30
				End If			
				If sHWSpec = "7703" Then
					FCopy "c:\s***\install\S*o\Bios\r*chmond_r6rap_" & sLock & ".rom", "c:\s***\bin\UpdateBIOS\afuwin32\r*chmond_r6rap_" & sLock & ".rom", 0
					Run "cmd /c " & Chr(34) & "cd c:\s***\bin\UpdateBIOS\afuwin32\ & afuwin.exe r*chmond_r6rap_" & sLock & ".rom /N /R" & Chr(34), 0, True, 1
					Sleep 30
				End If	
			End If
			If S*ountryCode = "UK" Then
				FCopy "c:\s***\install\S*o\Bios\r*chmond_i3_" & sLock & ".rom", "c:\s***\bin\UpdateBIOS\afuwin32\r*chmond_i3_" & sLock & ".rom", 0
				Run "cmd /c " & Chr(34) & "cd c:\s***\bin\UpdateBIOS\afuwin32\ & afuwin.exe r*chmond_i3_" & sLock & ".rom /N /R" & Chr(34), 0, True, 1
				Sleep 30
			End If	
		Case Else
			Abort "Hardware"
	End Select	
	If sLock = "u" Then
		oShell.RegWrite "HKEY_LOCAL_MACHINE\Software\N**\S***\Installation\PCIBiosUpdate", "No", "REG_SZ"
	Else
		oShell.RegWrite "HKEY_LOCAL_MACHINE\Software\N**\S***\Installation\PCIBiosUpdate", "Yes", "REG_SZ"
	End If	
	Log "Finished running BIOS command for " & sBoardType , sPackageLog
End Sub
Sub BoardCheck
	FileName = "c:\s***\install\boardtype.txt"
	Run "cmd /c wmic baseboard get Product /value |find " & Chr(34) & "=" & Chr(34) & "> " & FileName, 0, True, 5
	Set mefile = oFS.OpenTextFile(FileName, ForReading, False)
	sBoardType = mefile.ReadLine
	mefile.Close
	TestLog "sBoardType showing as:  " & sBoardType
	sBoardType = Trim(sBoardType)
	Do While(ASC(Right(sBoardType,1))<33 Or ASC(Right(sBoardType,1))>127)
		sBoardType = Left(sBoardType, Len(sBoardType)-1)
	Loop
	Do While(ASC(Left(sBoardType,1))<33 Or ASC(Left(sBoardType,1))>127)
		sBoardType = Right(sBoardType, Len(sBoardType)-1)
		sBoardType = Trim(sBoardType)
	Loop 'should output "Product=
	Select Case sBoardType
		Case "Product=m*naco"
			sBoardType = "m*naco"
		Case "Product=T*****dega"
			sBoardType = "T*****DEGA"
		Case "Product=Br***ol"
			sBoardType = "BR***OL"
		Case "Product=Br****l_II"
			sBoardType = "BR****L_II"
		Case "Product=P*cono"
			sBoardType = "P*CONO"	
		Case "Product=r*chmond"		
			sBoardType = "r*chmond"
		Case Else
			sBoardType = "Unknown Motherboard"
	End Select	
	TestLog "motherboard showing as:  " & sBoardType & " ."
End Sub
Sub FBackup(bfile)
	bfile2 = Replace(bfile,"c:\", "c:\s***\install\backup\DRPCI\")
	If oFS.FileExists(bfile2) = False Then
		Run "cmd /c copy " & bfile & " " & bfile2, 0, True, 1
	End If
End Sub
Sub FCopy(ofile, nfile, backup)
	If oFS.FileExists(ofile) = True Then
		If backup = 1 Then
			If oFS.FileExists(nfile) = True Then
				FBackup nfile
			End IF	
		End If	
		Run "cmd /c copy  " & ofile & " " & nfile, 0, True, 1
	Else
		Log "Unable to copy " & ofile & ".", sPackageLog
	End If	
End Sub
Sub FMove(ofile, nfile, backup)
	If oFS.FileExists(ofile) = True Then
		If backup = 1 Then
			If oFS.FileExists(nfile) = True Then
				FBackup nfile
			End IF	
		End If		
		Run "cmd /c move " & ofile & " " & nfile, 0, True, 1
	Else
		Log "Unable to move " & ofile & ".", sPackageLog
	End If	
End Sub
Sub Abort(sScope)'	Abort s**ipt w/logging
	Select Case sScope
		Case "Hardware"
			Log "HARDWARE CHECK: device is missing **itical hardware components or is incorrect type of device for package; aborting installation", sPackageLog
		Case "Software"
			Log "SOFTWARE CHECK: package not intended for current software level; aborting installation", sPackageLog
		Case "Missing"
			Log "FAILURE: " & icritFailure & " Critical file(s) are missing; aborting installation", sPackageLog
	End Select
	Log "=====================================================" & vbNewLine, sPackageLog
	WS**ipt.Quit
End Sub
Sub InstallStart()'	Common start-of-install tasks
	Log "=====================================================", sPackageLog
	Log "BEGIN: " & sPackageName & " installation", sPackageLog
	Log "DEVICE: " & sDevice & " [" & iLaneNumber & "] at " & sStoreNumber & "." & S*ountryCode, sPackageLog
	Log "CURRENT SOFTWARE LEVEL: " & sSWLevel, sPackageLog
	OKFile "Create"
	DNSVerify
	TestLog "Completed all InstallStart tasks."
End Sub
Sub InstallFinish()'	Common end-of-install tasks
	OKFile "Final"
	' DelStart
	Log "COMPLETE: " & sPackageName & " installation", sPackageLog
	Log "=====================================================" & vbNewLine, sPackageLog
	Log "COMPLETE: " & sPackageName & " installation", IHLog
	Log "=====================================================", IHLog
	Log "<--- " & sPackageName & " now installed on this device --->" & vbNewLine, IHLog
	' Sleep (iLaneNumber Mod 2) * 90
	' Reboot
End Sub
Sub Reboot()'	Reboot unit
	If sTest = 0 Then
		Run "C:\WINDOWS\system32\shutdown.exe -f -r -t 30", HideWindow, S**iptProceed, 120
		Log "Executing Initsys.exe to reboot system, shutdown.exe must have failed.", sPackageLog
		Run "C:\S***\BIN\initsys.exe", HideWindow, S**iptProceed, 0
	ElseIf sTest = 1 Then
		TestLog "Reboot sub called.  Ending script."
		Run "cmd /c c:\windows\system32\cmd.exe", 1, True, 50
		Ws**ipt.Quit 66
	End If	
End Sub
Sub OKFile(sMode)'	".OK" file handling
	Dim sText
	Select Case sMode
		Case "Delete"
			Log "ACTION: ensure that " & sOKFile & " is removed", sPackageLog
			Del sOKFile
			Exit Sub
		Case "Final"
			sText = ": complete"
		Case "Error"
			sText = ": possible installation error"
			Set oFile = oFS.OpenTextFile("c:\s***\Error" & sPackageName, ForWriting, True)
			oFile.WriteLine sOKFile & sText
			oFile.Close
		Case "Create"
			sText = ": installation in progress"
	End Select
	If oFS.FileExists(sOKFile) = False Then
		Set oFile = oFS.OpenTextFile(sOKFile, ForWriting, True)
		oFile.WriteLine sOKFile & sText
		oFile.Close
		Log "created " & sOKFile, sPackageLog
	End If
End Sub
Sub VerifyFileExists(sArgument)'	Verify that **itical files exist
	If oFS.FileExists(sArgument) = False Then
		Log "   **ITICAL FAILURE: Required file " & sArgument & " does not exist", sPackageLog
		Inc i**itFailure, 1
	End If
End Sub
Sub DNSVerify()'  DNS Suffix Search List check/fix
	Dim iCounter, sValue, sExisting, sData
	iCounter = 0
	sValue = "HKLM\System\CurrentControlSet\Services\TCPIP\Parameters\SearchList"
	sData = "secure.wa**mart.com,s0" & sStoreNumber & "." & LCase(S*ountryCode) & ".wa**mart.com,wa**mart.com"
	Do While iCounter <= 3
		sExisting = oShell.RegRead(sValue)
		If InStr(sExisting, sData) Then
			Log "DNS: suffix search list OK", sPackageLog
			Exit Do
		ElseIf InStr(sExisting, "lab.net") Then
			Log "DNS: Lab site; not altering DNSSSL", sPackageLog
			Exit Do
		Else
			Log "DNS: suffix search list is '" & sExisting & "'", sPackageLog
			Log "DNS: suffix search list incorrect; correcting", sPackageLog
			oShell.RegWrite sValue, sData, "REG_SZ"
			Sleep 30
			Inc iCounter, 1
		End If
	Loop
End Sub
Sub Popup(sText)'	Display full-s**een popup
	If sTest = 0 Then
		Run "cmd /c " & InDir & "cover1.exe " & sPackageName & " " & sText, HideWindow, S**iptProceed, 0
	ElseIf sTest = 1 Then
		TestLog "Popup: " & sText
	End If	
End Sub
Sub Log(sEntry, sFile)'	Log events to file
	Dim oFile
	Set oFile = oFS.OpenTextFile(sFile, ForAppending, **eateYes)
	oFile.WriteLine Replace(Now, "/", "-") & " " & sEntry
	oFile.Close
	Sleep 1
End Sub
Sub TestLog(sEntry)
	If sTest = 1 Then
		Dim oFile
		Set oFile = oFS.OpenTextFile(sPackageLog, ForAppending, **eateYes)
		oFile.WriteLine Replace(Now, "/", "-") & " " & "Test Log: " & sEntry
		oFile.Close
		Sleep 1
	End If	
End Sub
Sub Inc(iIn, iMod)'	Increment/decrement an integer value
	iIn = iIn + iMod
End Sub
Sub Sleep(iSleep)'	Script sleep
	WS**ipt.Sleep (iSleep * 1000)
End Sub
Sub Run(sCmd, iStyle, bWait, iSleep)'	Run a process
	If sTest = 1 Then
		TestLog "Run : " & sCmd
	End If	
	oShell.Run sCmd, iStyle, bWait
	Sleep iSleep
End Sub
Sub Del(sSpec)'	Delete a file/folder
	If oFS.FileExists(sSpec) = True Then
		oFS.DeleteFile(sSpec)
		Sleep 15
	ElseIf oFS.FolderExists(sSpec) = True Then
		oFS.DeleteFolder(sSpec)
		Sleep 15
	End If
End Sub
Sub LogErr
	If Err.Number <> 0 Then
		Log "LInt[" & Err.Number & "] / Hex[" & Hex(Err.Number) & "]   Source[" & Err.Source & "]   Error[" & vb**LF & Err.Des**iption & "]", sPackageLog
		Err.Clear
	End If
End Sub
Function RegRead(sPath)
	On Error Resume Next
	RegRead = Null : RegRead = oShell.RegRead(sPath) : LogErr
	If IsNull(RegRead) Then RegRead = ""
End Function
Sub CopyCat
	Dim WMI : Set WMI = GetObject(CIMV2)
	Dim ProcList : Set ProcList = WMI.ExecQuery("SELECT * FROM Win32_Process WHERE CommandLine LIKE '%" & WS**ipt.S**iptName & "%'") 
	If ProcList.Count > 1 Then
		Log "**NOTE**: Aborting concurrent run; installation proceeds", sPackageLog
		WS**ipt.Quit
	End If
	Set ProcList = Nothing : Set WMI = Nothing
End Sub
'======================================================
' End of script
'======================================================