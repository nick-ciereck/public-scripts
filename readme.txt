Readme for ConfigMgrStartup.xml

INSTALLATION:

This script will install the 2012 SCCM Client. 

1. Make a folder to hold your SCCM Client setup files. The path should look something like this \\fileserver\publicSoftware\sccmClient.

2. Place the "ConfigMgrStartup.vbs" and "ConfigMgrStartup.xml" files in the \sccmClient directory.

3. To implement hotfixes, In the \sccmClient directory make three directories called "I386", "X64" and "LegacyOSHotfix". Place the SCCM client hotfix(es) in the proper 32 or 64 bit directories.
Place the legacy hotfixes in the "LegacyOSHotfix" directory. Ensure the options in the xml file have been updated.

4. To implement error logging, make a "logs" folder on a server share. Note: Ensure that for each client computer, the "SYSTEM" account has write permissions to the logs folder. 

Note: The xml options CAN have trailing slashes. 

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
Option Value: /mp:https://yourmp.domain.com /usepkicert  SMSMP=https://yourmp.domain.com RESETKEYINFORMATION=TRUE DNSSUFFIX=Your.DNS.com CCMCERTSEL="SubjectAttr:OU = your department"
Option Description: Any additional ccmsetup.exe parameters that you'd like to add.
