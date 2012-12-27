Option Name: ConfigLogPath
Option Value: \\fileserver\logs$
Option Description: Folder that the final log will be copied to. It will be named: scriptname + computername + timestamp

Option Name: LegacyOSCertificateHotfixFolder
Option Value: \\fileserver\publicSoftware\hotfix  
Option Description: Folder where the legacy certificate hotfixes are stored.

Option Name: LegacyOSHotfix_XP_x32
Option Value: WindowsXP-KB968730-x86-ENU.exe
Option Description: Filename (without path) of the downloaded and extracted hotfix.

Option Name: LegacyOSHotfix_XP2003_x64
Option Value: WindowsServer2003.WindowsXP-KB968730-x64-ENU.exe
Option Description: Filename (without path) of the downloaded and extracted hotfix.

Option Name: CertHotFixID
Option Value: KB968730
Option Description: KB number of the XP\2003 hotfix to install. This must be the same string as "hotfixID" listed the Powershell command (gwmi win32_QuickFixEngineering).
		
Option Name: ClientLocation
Option Value: \\fileserver\publicSoftware\sccmClient
Option Description: Path to the client installer and supporting files (ccmsetup.exe).		

Option Name: AutoHotFix
Option Value: \\fileserver\publicSoftware\sccmClient
Option Description: path to the client hotfix. We assume that you put the hotfixes in <autoHotFixPath>\x64, and the <autoHotFixPath>\i386 (respectively).
		
Option Name: ForceReinstall
Option Value: True
Option Description: Forces reinstall of the client, even if the client appears to be healthy.

Option Name: ForceUninstall
Option Value: True
Option Description: Forces uninstall of the client before any other operations. Useful for client corruption issues where an uninstall is needed before a reinstall will work.
		
Option Name: OtherInstallationProperties
Option Value: /mp:https://yourmp.domain.com /usepkicert  SMSMP=https://yourmp.domain.com RESETKEYINFORMATION=TRUE DNSSUFFIX=asc.ohio-state.edu CCMCERTSEL="SubjectAttr:OU = your department"
Option Description: Any additional ccmsetup.exe parameters that you'd like to add.
