' ConfigMgr Startup Script
' Version 3.09
' Initial writer: Jason Sandys; http://blog.configmgrftw.com
' Extended significantly by: Nick Ciereck; http://win1337ist.wordpress.com
' Refactored by: John Puskar; http://windowsmasher.wordpress.com
'

'To Do
'  * install client loop which waits for ccmsetup.exe should log once per minute
'  * clientConfig should not honor ForceReinstall or ForceUninstall if the option exists and does not equal true.
'  * forceUninstall should only uninstall right before the client is installed (because of a client health failure, or forceReinstall).

'known issues: (f)InWinPE will return true if the system is -not- in WinPE but an X: exists.


Option Explicit

Dim g_fso, g_WshShell, g_logPathandName, g_startTime
Set g_fso = CreateObject ("Scripting.FileSystemObject")
Set g_WshShell = WScript.CreateObject("WScript.Shell")

' Custom Variables added by Nick Ciereck
Dim g_finalLog 
Dim strComputerName 
Dim strLogName
Dim hotFixPathAndName
Dim objWMIService
Dim colObjProc
Dim ObjProc
Dim hotFixPathAndFile 
Dim g_sRepairsRun

Const CONFIGMGRLOGPATH = 				 "ConfigLogPath"

'Cert HotFix Path
Const optLegacyCertHotFixpath =  "LegacyOSCertificateHotfixFolder"
Const xp_X86_CertHotFix = "WindowsXP-KB968730-x86-ENU.exe"
Const Srvr_2003_X86_CertHotFix = "WindowsServer2003-KB968730-x86-ENU.exe"
Const xp_2003_X64_CertHotFix = "WindowsServer2003.WindowsXP-KB968730-x64-ENU.exe"

'Const hotFixPath = 					"AutoHotfix"
Const hotFix64Name = "configmgr2012ac-rtm-kb2717295-x64.msp"
Const hotFix32Name = "configmgr2012ac-rtm-kb2717295-i386.msp"	

'Cert hotfix ID
Const CERTHOTFIXID = "CertHotFixID"
Const PROCESSOR_ARCHITECTURE_X86  = 0
Const PROCESSOR_ARCHITECTURE_IA64 = 6
Const PROCESSOR_ARCHITECTURE_X64  = 9					

'Default Options

Const OPTION_LOCALADMIN =										"LocalAdmin"
Const OPTION_LOCALADMIN_GROUP =							"LocalAdminGroup"
Const OPTION_AGENTVERSION =									"AgentVersion"
Const OPTION_DEFAULT_RUNINTERVAL =					"MinimumInterval"
Const OPTION_CACHESIZE =										"CacheSize"
Const OPTION_INSTALLPATH =									"ClientLocation"
Const OPTION_SITECODE =											"SiteCode"
Const OPTION_OTHERINSTALLPROPS = 						"OtherInstallationProperties"
Const OPTION_MAXLOGFILE_SIZE =							"MaxLogFile"
Const OPTION_ERROR_LOCATION =								"ErrorLocation"
Const OPTION_AUTOHOTFIX =										"AutoHotfix"
Const OPTION_STARTUPDELAY =									"Delay"
Const OPTION_WMISCRIPT =										"WMIScript"
Const OPTION_WMISCRIPT_ASYNCH =							"WMIScriptAsynch"
Const OPTION_WMISCRIPTOPTIONS =							"WMIScriptOptions"
Const OPTION_FORCE_REINSTALL = 							"ForceReinstall"
Const OPTION_FORCE_UNINSTALL = 							"ForceUninstall"
Const OPTION_XP_2003_X64_CERTHOTFIX = 			"LegacyOSHotfix_XP2003_x64"
Const OPTION_XP_X32_CERTHOTFIX = 		      	"LegacyOSHotfix_XP_x32"
Const OPTION_2003_X32_CERTHOTFIX = 		      "LegacyOSHotfix_2003_x32"
Const DEFAULT_REGISTRY_LOCATION =						"HKLM\Software\ConfigMgrStartup"
Const DEFAULT_LOCALADMIN_GROUP =						"Administrators"
Const DEFAULT_AGENTVERSION =								"4.00.6487.2000"
Const DEFAULT_RUN_INTERVAL =								12
Const DEFAULT_CACHESIZE =										"5120"
Const DEFAULT_REGISTRY_LASTRUN_VALUE =			"Last Run"
Const DEFAULT_REGISTRY_LOGLOCATION_VALUE =	"Log Location"
Const DEFAULT_REGISTRY_LASTRESULT_VALUE =		"Last Execution Result"
Const DEFAULT_CONFIGFILE_PARAMETER =				"config"
Const DEFAULT_MAXLOGFILE_SIZE =							"6144"
Const DEFAULT_EVENTLOG_PREFIX =							"ConfigMgr StartUp Script -- "
Const DEFAULT_WMISCRIPT_ASYNCH =						"1"
Const DEFAULT_XP_2003_X64_CERTHOTFIX =			"WindowsServer2003.WindowsXP-KB968730-x64-ENU.exe"

'Messages and Outputs
Const MSG_EndOfScript =										"END_OF_SCRIPT"
Const MSG_BlankSpace = 										" "
Const MSG_OK =														"...OK"
Const MSG_NOTOK =													"...FAILED"
Const MSG_FOUND = 												"...found"
Const MSG_NOTFOUND =											"...not found"
Const MSG_MAIN_BEGIN =										"Beginning Execution at "
Const MSG_MAIN_FINISH =										"Finished Execution at "
Const MSG_ELAPSED_TIME = 									"Total script execution time is "
Const MSG_DIVIDER =												"----------------------------------------"
Const MSG_LOGMSG_CLIENTSTATUS =        		"Client Check "
Const MSG_LOGMSG_FILEERROR =        			"Unable to create or update error file at "
Const MSG_LOGMSG_FILEOK =        					"Successfully created or updated error file at "
Const MSG_LOGMSG_FILEDELETEERROR = 				"Unable to remove error file at "
Const MSG_LOGMSG_FILEDELETEOK = 					"Successfully removed error file at "
Const MSG_OPENCONFIG_NOT_SPECIFIED =			"Configuration file not specified on command-line with config switch"
Const MSG_OPENCONFIG_DOESNOTEXIST =				"The configuration file does not exist: "
Const MSG_OPENCONFIG_PARSEERROR	=					"The specified configuration file contains a parsing error: "
Const MSG_OPENCONFIG_OPENED	=							"Opened configuration file: "
Const MSG_LOADOPTIONS_STARTED = 					"Loading Options and Parameters from configuration file"
Const MSG_LOADOPTIONS_OPTIONLOADED =			"Option loaded: "
Const MSG_LOADOPTIONS_PARAMLOADED =				"Parameter loaded: "
Const MSG_LASTRUN_VERIFYING =							"Verifying Last Run time from Registry: "
Const MSG_LASTRUN_NOLASTRUN =							"No last run time recorded in registry"
Const MSG_LASTRUN_TIME =									"Last run time: "
Const MSG_LASTRUN_TIMENOTOK =							"Existing because last run time was less than expected number of hours ago: "
Const MSG_LASTRESULT_VERIFYING =					"Verifying Last Result from Registry: "
Const MSG_LASTRESULT_NOLASTRESULT =				"No last result recorded in registry"
Const MSG_LASTRESULT_RESULT =							"Last execution result: "
Const MSG_LASTRESULT_FAIL =								"Failed"
Const MSG_LASTRESULT_SUCCEED = 						"Succeeded"
Const MSG_CHECKWMI_ERROR = 								"Error Connecting to WMI: "
Const MSG_CHECKWMI_SUCCESS = 							"Successfully Connected to WMI"
Const MSG_WMISCRIPT_NOTFOUND =						"Could not find WMI Script: "
Const MSG_WMISCRIPT_EXECUTING =						"Executing WMI Script: "
Const MSG_WMISCRIPT_EXECUTINGASYNCH =			"Asynchonously executing WMI Script: "
Const MSG_WMISCRIPT_OPTIONS =							"WMI Script Options: "
Const MSG_WMISCRIPT_ERROR =								"Failed to successfully run the WMI Script: "
Const MSG_WMISCRIPT_SUCCESS =							"Successfully ran the WMI Script."
Const MSG_WMISCRIPT_SUCCESSASYNCH =				"Successfully started the WMI Script, not waiting for results."
Const MSG_CHECKSERVICE_START = 						"START: Service Check..."
Const MSG_CHECKSERVICE_STARTMODE =				"...expected StartMode of "
Const MSG_CHECKSERVICE_STARTMODEOK =			"...set start mode to "
Const MSG_CHECKSERVICE_STARTMODEFAIL =		"...failed to set start mode with error: "
Const MSG_CHECKSERVICE_STATE =						"...expected State of "
Const MSG_CHECKSERVICE_STARTEDOK =				"...started service"
Const MSG_CHECKSERVICE_STARTEDFAIL =			"...could not start service, error: "
Const MSG_CHECKSERVICE_STOPPEDOK =				"...stopped service"
Const MSG_CHECKSERVICE_STOPPEDFAIL =			"...could not stop service, error: "
Const MSG_CHECKADMINSHARE_START = 				"START: Admin Share Check..."
Const MSG_CHECKADMINSHARE_SETSUCCESS =		"...set AutoShareWks registry value, a reboot is required to create the Admin$ share."
Const MSG_CHECKADMINSHARE_SETFAIL =				"...unable to set AutoShareWks registry value: "
Const MSG_CHECKREGISTRY_START =						"START: Registry Check..."
Const MSG_CHECKREGISTRY_EXPECTED =				"...expected value of "
Const MSG_CHECKREGISTRY_ENFORCEOK =				"...successfully updated value"
Const MSG_CHECKREGISTRY_ENFORCEFAIL =			"...failed to update value"
Const MSG_CHECKLOCALADMIN_START = 				"START: Local Admin Check..."
Const MSG_CHECKLOCALADMIN_ALREADYMEMBER =	"...user already member"
Const MSG_CHECKLOCALADMIN_ADDMEMBEROK =		"...user add successful"
Const MSG_CHECKLOCALADMIN_ADDMEMBERFAIL =	"...user add failed with error: "
Const MSG_CHECKCLIENT_START =							"START: Checking Client Status..."
Const MSG_CHECKCLIENT_WMINOTFOUND =				" *Cannot connect to ConfigMgr WMI Namespace: "
Const MSG_CHECKCLIENT_VERSION =						"START: Getting ConfigMgr agent version..."
Const MSG_CHECKCLIENT_VERSIONNOTFOUND =		" *Cannot determine ConfigMgr agent version"
Const MSG_CHECKCLIENT_OLDVERSION =				" *Old version of agent found: "
Const MSG_CHECKCLIENT_CCMEXEC =						"START: Checking SMS Agent Host Status..."
Const MSG_CHECKCLIENT_VERSIONFOUND =			" *SMS Agent Host version: "
Const MSG_CHECKCLIENT_MOVEDLOG =					"Moved log file to "
Const MSG_CHECKCLIENT_MOVELOGFAIL =				"Unable to move log file, error: "
Const MSG_CHECKCLIENT_GETLOGDIRECTORY =		"Unable to get agent log directory with error: "
Const MSG_CHECKCLIENT_VERSIONEXPECTED =		" *Expected client version: "
Const MSG_CHECKCACHE_START =							"START: Check agent cache..."
Const MSG_CHECKCACHE_CREATEFAIL =					" *Could not create UIResourceManager with error: "
Const MSG_CHECKCACHE_WMIFAIL =						" *Could not retrieve the cache object from WMI with error: "
Const MSG_CHECKCACHE_WMIWRITEFAIL =				" *Could not set the cache size in WMI with error: "
Const MSG_CHECKCACHE_CACHEFAIL =					" *Could not retrieve agent cache size with error: "
Const MSG_CHECKCACHE_SETSIZE =						" *Set cache size to "
Const MSG_CHECKCACHE_SIZEOK =	  					" *Current cache size is "
Const MSG_INSTALLCLIENT_START =						"START: Client Install..."
Const MSG_INSTALLCLIENT_PATHCHECK =				" *Checking for ccmsetup in "
Const MSG_INSTALLCLIENT_COMMANDLINE =			" *Initiating client install with command-line: "
Const MSG_INSTALLCLIENT_SUCCESS =					" *Successfully initiated CCMSetup"
Const MSG_INSTALLCLIENT_FAILED =					" *Failed to initiate CCMSetup with error: "
Const MSG_AUTOPATCH_COMMANDLINE =					" *Discovering client hotfixes from: "
Const MSG_AUTOPATCH_DIRERROR =						" *Unable to open the hotfix folder: "
Const MSG_AUTOPATCH_FOUNDHOTFIX =					"  ...Found hotfix: "
Const MSG_HOTFIX_FILEVERIFY =							" *Verifying hotfix accessibility: "
Const MSG_HOTFIX_DUPLICATE =							"  ...Hotfix already added: "
Const MSG_HOTFIX_MULTIPLE =								" *Multiple hotfixes specified, cannot verify accessibility "
Const MSG_CHECKASSIGNMENT_START =					"START: Checking client assignment..."
Const MSG_CHECKASSIGNMENT_OK =						" *Client assigned to site " 
Const MSG_CHECKASSIGNMENT_NOTOK =					" *Client not assigned to site, initiaing (re-)install"
Const MSG_VERIFYINPUT_OK =								"The input variables from the xml have been verified as correct"
Const MSG_VERIFYINPUT_FAILED =						"FAILED Verification of xml variables"

'= Functions =

Sub Main
		Dim argsNamed
		Dim WshShell
		Dim xmlConfig
		Dim configOptions
		Dim parameters
		Dim bClientOK
		Dim lastResult
		Dim clientVerCheck
		Dim bContinue
		Dim bInstallClient
		Dim iClientStatus
		Dim i
		Dim msg
		Dim bAction
		Dim OSVer, OSArch
		Dim bLegacyHotfixNeeded
		Dim clientInstalled
		Dim bSkipAction
		Dim bResult
		
		Const CLIENT_INSTALLED = 1
		Const CLIENT_CLIENT_UPDATE_REQUIRED = 2
		Const CLIENT_CACHE_FAILURE = 3
		Const CLIENT_ASSIGNMENT_FAIlURE = 4
		Const CLIENT_CHECK_FAILED = 6
		Const CLIENT_UNKNOWN_ERROR = 8
		
		bInstallClient = False
		bContinue = True
		bClientOK = False
	
		Set argsNamed = WScript.Arguments.Named
		Set configOptions = WScript.CreateObject("Scripting.Dictionary")
		Set parameters = WScript.CreateObject("Scripting.Dictionary")
		
		'Load Config
		If OpenConfig (argsNamed, xmlConfig) Then
			LoadOptions xmlConfig, configOptions, parameters
		Else
			bContinue = False 
		End If
		
		'Delay the script if the option says to do so.
		If configOptions.Exists(OPTION_STARTUPDELAY) Then
			Delay CInt(configOptions.Item(OPTION_STARTUPDELAY))
		End If
	
		'Check the script and system environments to see if the script can run and complete.
		lastResult = GetLastResult(configOptions)
		If Check_OSWMI(configOptions) <> True Then
			bContinue = False
		ElseIf CheckServices(xmlConfig) <> True Then
			bContinue = False
		ElseIf Conform_AdminShareOptions <> True Then
			bContinue = False
		ElseIf Confirm_RegistryOptions(xmlConfig) <> True Then
			bContinue = False
		ElseIf Conform_LocalAdminOptions(configOptions) <> True Then
			bContinue = False
		ElseIf LastRunOK(configOptions) <> True Then
			bContinue = False
		End If
		
		'Check the current OS for legacy hotfixes
		If(bContinue) Then
			msg = "Checking if the client needs a legacy OS hotfix."
	    WriteLogMsg msg, 1, 1, 0
			bLegacyHotfixNeeded = False
			bLegacyHotfixNeeded = CheckOSNeedsCertHotfix(configOptions)
			If bLegacyHotfixNeeded = True Then
	    	msg = "Installing the legacy OS certificate hotfix."
	    	WriteLogMsg msg, 1, 1, 0
	    	bResult = False
	    	GetOSVersionAndArch OSVer, OSArch
	    	msg = "Read OS Version as """ & OSVer & """ and OS Architecture as """ & OSArch & """."
	    	WriteLogMsg msg, 1, 1, 0
	    	bResult = InstallLegacyOSHotfix(configOptions, OSVer, OSArch)
		 	  msg = "Re-checking legacy OS hotfix installation status after the install."
	    	WriteLogMsg msg, 1, 1, 0
		 	  bResult = CheckOSNeedsCertHotfix(configOptions)
				If bResult = False Then
					WriteLogMsg "OS Hotfix Installed", 1, 1, 0
					bContinue = True
				Else
					writeLogMsg "OS Hotfix Install failed.", 1, 1, 0
					bContinue = False
				End If
			End If
		End If
		
		'obey forceUninstall
		If configOptions.Exists(OPTION_FORCE_UNINSTALL) Then
			If configOptions.Item(OPTION_FORCE_UNINSTALL) = True Then
				bAction = UninstallClient(configOptions, parameters)
			End If
		End If
		
		'obey forceReinstall
		'If bContinue = True Then
			If configOptions.Exists(OPTION_FORCE_REINSTALL) Then
				If configOptions.Item(OPTION_FORCE_REINSTALL) = "True" Then
					writeLogMsg "Reinstalling client immediately per config options.", 1, 1, 0
					bAction = InstallClient(configOptions, parameters)
				End If
			End If
		'End If
		
		'Test client health and repair if necessary
		Dim arrActions(20)
		Dim action
		Dim bResults
		arrActions(0) = "clientInstalled"
		arrActions(1) = "clientWmiConnectivity"
		arrActions(2) = "clientServicesRunning"
		arrActions(3) = "clientVersionLevel"
		arrActions(4) = "clientCacheSizeOK"
		arrActions(5) = "clientSiteAssignment"
		For Each action in arrActions
			If action = "" OR action = null Then
				Exit For
			End If
			
			'Skip site assignment if we just installed
			bSkipAction = False
			If InStr(g_sRepairsRun, "clientInstalled") _
				OR InStr(g_sRepairsRun, "clientWMIConnectivity") _
				OR InStr(g_sRepairsRun, "clientServicesRunning") _
				OR InStr(g_sRepairsRun, "clientVersionLevel") Then
					If action = "clientSiteAssignment" Then
						bSkipAction = True
					End If
			End If
			
			If bSkipAction = False Then
				msg = "Running the following client check: """ +  action + """."
				WriteLogMsg msg, 1, 1, 0
				bResults = false
				bResults = Run_ClientCheck (action,configOptions)
				If bResults = True Then
					msg = " *The following action passed: """ + action + """."
					WriteLogMsg msg, 1, 1, 0
				Else
					msg = " *The following action failed """ + action + """."
					bResults = False
					WriteLogMsg msg, 1, 1, 0
					msg = "Attempting to repair..."
					WriteLogMsg msg, 1, 1, 0
					bResults = Repair_ClientCheck (action, configOptions, parameters)
					bResults = Run_ClientCheck (action,configOptions)
					If bResults = True Then
						g_sRepairsRun = g_sRepairsRun + "," + action
					End If
				End If
			End If
		Next
		
'		If configOptions.Exists(OPTION_ERROR_LOCATION) Then
'			WriteErrorFile clientError, configOptions.Item(OPTION_ERROR_LOCATION), lastResult												
'		End If
	WriteLogMsg "Script execution complete.", 1, 1, 0
	WriteFinalLogMsg configOptions
End Sub

Function Repair_ClientCheck(sCheckName, configOptions, parameters)
	Dim bResults
	bResults = False
	'''''''''WORK NEEDED
	'''case clientVersion could maybe just upgrade instead of reinstall
	Select Case sCheckName
		Case "clientInstalled"
			bResults = InstallClient(configOptions, parameters)
		Case "clientWmiConnectivity"
			bResults = InstallClient(configOptions, parameters)
		Case "clientServicesRunning"
			bResults = InstallClient(configOptions, parameters)
		Case "clientVersionLevel"
			bResults = InstallClient(configOptions, parameters)
		Case "clientCacheSizeOK"
			bResults = Fix_ClientCacheSize(configOptions)
		Case "clientSiteAssignment"
			bResults = InstallClient(configOptions, parameters)
	End Select
	Repair_ClientCheck = bResults
End Function

Function Run_ClientCheck(sCheckName, configOptions)
	Dim bResults
	bResults = False
	Select Case sCheckName
		Case "clientInstalled"
			bResults = Check_IsClientInstalled(configOptions)
		Case "clientWmiConnectivity"
			bResults = Check_ClientWmiConnectivity(configOptions)
		Case "clientServicesRunning"
			bResults = Check_ClientServicesRunning(configOptions)
		Case "clientVersionLevel"
			bResults = Check_ClientVersionLevel(configOptions)
		Case "clientCacheSizeOK"
			bResults = Check_ClientCacheSizeOK(configOptions)
		Case "clientSiteAssignment"
			bResults = Check_ClientSiteAssignment(configOptions)
	End Select
	Run_ClientCheck = bResults
End Function

Function Check_ClientSiteAssignment(configOptions)
	Dim smsClient, siteCode
	Dim errorCode
	Dim bResults
	bResults = False
	
	WriteLogMsg MSG_CHECKASSIGNMENT_START, 1, 1, 1
	On Error Resume Next
	Err.Clear
	Set smsClient = CreateObject ("Microsoft.SMS.Client")
	errorCode = Err.Number
	siteCode = smsClient.GetAssignedSite
	On Error GoTo 0
	
	If Len(siteCode) = 0 Or errorCode <> 0 Then
		WriteLogMsg MSG_CHECKASSIGNMENT_NOTOK & siteCode, 1, 1, 0
		bResults = False
	Else
		WriteLogMsg MSG_CHECKASSIGNMENT_OK & siteCode, 1, 1, 0
		bResults = True
	End If
	Set smsClient = Nothing
	
	If bResults = True Then
  	Check_ClientSiteAssignment = True
  Else
  	Check_ClientSiteAssignment = False
  End If
End Function

Function Check_ClientCacheSizeOK(configOptions)
	Dim uiResManager, cache
	Dim errorCode, errorMsg
	Dim desiredCacheSize
	Dim msg, bResults, bContinue
	
	bResults = False
	bContinue = True
	WriteLogMsg MSG_CHECKCACHE_START, 1, 1, 1
	
	'Bind to CCM Resource Manager
	On Error Resume Next
	Err.Clear
	Set uiResManager = CreateObject("UIResource.UIResourceMgr")
	errorCode = Err.Number
	errorMsg = Err.Description & " (" & Err.Number & ")"
	On Error GoTo 0
	If errorCode <> 0 Then
		WriteLogMsg MSG_CHECKCACHE_CREATEFAIL & errorMsg, 2, 1, 0
		bResults = False
		bContinue = False
	End If
	
	'Retrieve data from CCM Resource Manager
	If bContinue = True Then
		On Error Resume Next
		Err.Clear
		Set cache = uiResManager.GetCacheInfo
		errorCode = Err.Number	
		errorMsg = Err.Description & " (" & Err.Number & ")"
		On Error GoTo 0
		If errorCode <> 0 Then
		   Set uiResManager = Nothing
		   WriteLogMsg MSG_CHECKCACHE_CACHEFAIL & errorMsg, 2, 1, 0
		   bResults = False
		   bContinue = False
		End If
	End If
	
	'Compare expected vs read
	If bContinue = True Then
		desiredCacheSize = CInt(GetOptionValue(OPTION_CACHESIZE, DEFAULT_CACHESIZE, configOptions))
		If cache.TotalSize < desiredCacheSize Then
			msg = " *Error: client cache size is currently """ & cache.TotalSize & """."
			WriteLogMsg msg, 1, 1, 0
			msg = " *Error: This is less than the desired minimum cache size of """ & desiredCacheSize & """ listed in the XML options."
			WriteLogMsg msg, 1, 1, 0
			bResults = False
		Else
			WriteLogMsg MSG_CHECKCACHE_SIZEOK & cache.TotalSize, 1, 1, 0
			bResults = True
		End If
		Set uiResManager = Nothing
  End If
  
  If bResults = True Then
  	Check_ClientCacheSizeOK = True
  Else
  	Check_ClientCacheSizeOK = False
  End If
End Function

Function Fix_ClientCacheSize(configOptions)
	Dim uiResManager, cache
	Dim errorCode, errorMsg
	Dim desiredCacheSize
	Dim msg, bResults, bContinue
	
	bResults = False
	bContinue = True
	WriteLogMsg MSG_CHECKCACHE_START, 1, 1, 1
	
	'Bind to CCM Resource Manager
	On Error Resume Next
	Err.Clear
	Set uiResManager = CreateObject("UIResource.UIResourceMgr")
	errorCode = Err.Number
	errorMsg = Err.Description & " (" & Err.Number & ")"
	On Error GoTo 0
	If errorCode <> 0 Then
		WriteLogMsg MSG_CHECKCACHE_CREATEFAIL & errorMsg, 2, 1, 0
		bResults = False
		bContinue = False
	End If
	
	'Retrieve data from CCM Resource Manager
	If bContinue = True Then
		On Error Resume Next
		Err.Clear
		Set cache = uiResManager.GetCacheInfo
		errorCode = Err.Number	
		errorMsg = Err.Description & " (" & Err.Number & ")"
		On Error GoTo 0
		If errorCode <> 0 Then
		   Set uiResManager = Nothing
		   WriteLogMsg MSG_CHECKCACHE_CACHEFAIL & errorMsg, 2, 1, 0
		   bResults = False
		   bContinue = False
		End If
	End If
	
	'Compare expected vs read
	If bContinue = True Then
		desiredCacheSize = CInt(GetOptionValue(OPTION_CACHESIZE, DEFAULT_CACHESIZE, configOptions))
		If cache.TotalSize < desiredCacheSize Then
'			msg = " *Error: client cache size is currently """ & cache.TotalSize & """."
'			WriteLogMsg msg, 1, 1, 0
'			msg = " *Error: This is less than the desired minimum cache size of """ & desiredCacheSize & """ listed in the XML options."
'			WriteLogMsg msg, 1, 1, 0
			bResults = False
		Else
			WriteLogMsg MSG_CHECKCACHE_SIZEOK & cache.TotalSize, 1, 1, 0
			bResults = True
		End If
  End If
  
  'Change the cache
  If bContinue = True Then
		If cache.TotalSize <> desiredCacheSize Then
			cache.TotalSize = desiredCacheSize
			WriteLogMsg MSG_CHECKCACHE_SETSIZE & desiredCacheSize, 1, 1, 0
		Else
			WriteLogMsg MSG_CHECKCACHE_SIZEOK & cache.TotalSize, 1, 1, 0
		End If
	  Set uiResManager = Nothing
  End If
  
  bResults = Check_ClientCacheSizeOK(configOptions)
  Fix_ClientCacheSize = bResults
End Function

Function Check_ClientVersionLevel(configOptions)
	Dim wmi, ccmWMI, errorCode, errorMsg
	Dim clientProperties, clientProp
	Dim sActualClientVersion, expectedVersion
	Dim configMgrLogPath
	Dim clientCheck
	Dim bResults
	Dim bContinue
	Dim msg
	bContinue = True
	bResults = False
	
	'Bind to CCM WMI
	On Error Resume Next
	Err.Clear
	Set ccmWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\ccm")
  errorCode = Err.Number
	errorMsg = Err.Description & " (" & Err.Number & ")"
	On Error GoTo 0
	If errorCode <> 0 Then
		WriteLogMsg MSG_CHECKCLIENT_WMINOTFOUND & errorMsg, 2, 1, 0
		bResults = False
		bContinue = False
	End If
	
	'Get the greatest CCM component version
	If bContinue = True Then
		sActualClientVersion = Find_CurrentClientVersion(ccmwmi)
		If sActualClientVersion = null OR sActualClientVersion = "" Then
			bContinue = False
			msg = " *ERROR: Could not determine client version."
			WriteLogMsg msg, 1, 1, 0
		End If
	End If
	
	'Compare real and expected versions
	If bContinue = True Then
		expectedVersion = GetOptionValue(OPTION_AGENTVERSION, DEFAULT_AGENTVERSION, configOptions)
		WriteLogMsg MSG_CHECKCLIENT_VERSIONEXPECTED & expectedVersion, 1, 1, 0
		If Compare_ClientVersionStrings(expectedVersion, sActualClientVersion) = False Then
			'Wscript.echo "The current client version is old, ConfigMgr will install and hotfix will be applied"
			WriteLogMsg MSG_CHECKCLIENT_OLDVERSION & sActualClientVersion, 2, 1, 0
			bResults = False
		Else
			'Wscript.echo "The current client is up to date, hotfix and client will not be installed"
			WriteLogMsg MSG_CHECKCLIENT_VERSIONFOUND & sActualClientVersion, 1, 1, 0
			bResults = True
		End If
	End If
	'return
	If bResults = True Then
		Check_ClientVersionLevel = True
	Else
		Check_ClientVersionLevel = False
	End If
End Function

Function Find_CurrentClientVersion(ccmWmi)
		Dim clientProperties
		Dim clientProp
		Dim sActualClientVersion
		Dim msg
		Dim retval
		
		WriteLogMsg MSG_CHECKCLIENT_VERSION, 1, 1, 0
		Set clientProperties = ccmWmi.ExecQuery("Select * from SMS_Client")
		On Error Resume Next
		For Each clientProp In clientProperties
			'msg = "Iterating current WMI property version number: " + clientProp.ClientVersion
			'WriteLogMsg msg, 1, 1, 0
			
			'set initial if needed
			If sActualClientVersion = null OR sActualClientVersion = "" Then
				sActualClientVersion = clientProp.ClientVersion
			End If
			
			'Check current iteration
			If clientProp.ClientVersion = Null or clientProp.ClientVersion = "" Then
				Exit For
			End If
			
			'check current versoin against next iteration
			If Compare_ClientVersionStrings (sActualClientVersion, clientProp.ClientVersion) = False Then
				sActualClientVersion = clientProp.ClientVersion
			End If
		Next
		If sActualClientVersion = null OR sActualClientVersion = "" OR sActualClientVersion = "0" Then
			WriteLogMsg MSG_CHECKCLIENT_VERSIONNOTFOUND & errorMsg, 2, 1, 0
		End If
	
	If sActualClientVersion = null Then
		retval = False
	Else
		retval = sActualClientVersion
	End If
	Find_CurrentClientVersion = Retval
End Function

Function Check_ClientServicesRunning(configOptions)
	Dim wmi, errorCode, errorMsg
	Dim bResults
	
	WriteLogMsg MSG_CHECKCLIENT_CCMEXEC, 1, 1, 0
	Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	If CheckService(wmi, "CCMExec", "Running", "Auto", True) = False Then
	  bResults = False
	Else
		bResults = True
	End If
	Set wmi = Nothing
	
	If bResults = True Then
		Check_ClientServicesRunning = True
	Else
		Check_ClientServicesRunning = False
	End If
End Function

Function Check_ClientWmiConnectivity(configOptions)
	Dim wmi, ccmWMI, errorCode, errorMsg
	Dim bResults

	On Error Resume Next
	Err.Clear
	Set ccmWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\ccm")
  errorCode = Err.Number
	errorMsg = Err.Description & " (" & Err.Number & ")"
	'On Error GoTo 0
	Set ccmWMI = Nothing
	
	If errorCode <> 0 Then
		WriteLogMsg MSG_CHECKCLIENT_WMINOTFOUND & errorMsg, 2, 1, 0
		bResults = False
		Exit Function
	Else
		bResults = True
	End If
	
	If bResults = True Then
		Check_ClientWmiConnectivity = True
	Else
		Check_ClientWmiConnectivity = False
	End If
End Function

Function Check_IsClientInstalled(configOptions)
	Dim wmi
	Dim bResults
	
	WriteLogMsg "Checking whether client is installed", 1, 1, 0
	
	'Look for ccmexec service
	WriteLogMsg MSG_CHECKCLIENT_CCMEXEC, 1, 1, 0
	Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	If CheckService(wmi, "CCMExec", "Running", "Auto", True) = False Then
		bResults = False
	Else
		bResults = True
	End If
	
	'We 'outta check for the msi installer property instead...
	'''''WORK NEEDED
	Set wmi = Nothing
	
	If bResults = False Then
		Check_IsClientInstalled = False
	Else
		Check_IsClientInstalled = True
	End If
End Function

Function Compare_ClientVersionStrings(ByVal expectedVersion, ByVal currentVersion)
	'Returns true if the 'current version' argument is greater than the 'expected version' argument.
	Dim currentVersionArray, expectedVersionArray
	Dim versionPartCount, bResults
	currentVersionArray = Split(currentVersion, ".", -1, 1)
	expectedVersionArray = Split(expectedVersion, ".", -1, 1)
	bResults = True
	For versionPartCount = 0 To 3
		If currentVersionArray(versionPartCount) > expectedVersionArray(versionPartCount) Then
			Exit For
		ElseIf currentVersionArray(versionPartCount) < expectedVersionArray(versionPartCount) Then
			bResults = False
			Exit For
		End If
	Next
	
	If bResults = False Then
		Compare_ClientVersionStrings = False
	Else
		Compare_ClientVersionStrings = True
	End If
End Function

Sub WriteErrorFile (ByVal clientErr, ByVal errorLocation, ByVal lastExecutionResult)
	Dim badLog, badLogFileName
	Dim network
	Dim errorCode, errorMsg
	Dim registryLocation
	
	Set network = WScript.CreateObject("WScript.Network")
	badLogFileName = errorLocation & "\" & network.ComputerName & ".log"

	registryLocation = DEFAULT_REGISTRY_LOCATION & "\" & DEFAULT_REGISTRY_LASTRESULT_VALUE

	If clientErr = True Then

		WriteLogMsg MSG_LOGMSG_CLIENTSTATUS & MSG_LASTRESULT_FAIL, 1, 1, 0
		g_WshShell.LogEvent 1, MSG_LOGMSG_CLIENTSTATUS & MSG_LASTRESULT_FAIL
		g_WshShell.RegWrite registryLocation, MSG_LASTRESULT_FAIL, "REG_SZ"

		On Error Resume Next
		Err.Clear

		Set badLog = g_fso.OpenTextFile(badLogFileName, 8, True)
	    errorCode = Err.Number
	    errorMsg = Err.Description & " (" & Err.Number & ")"
    
   		On Error GoTo 0
    
	    If errorCode <> 0 Then
	    	WriteLogMsg MSG_LOGMSG_FILEERROR & badLogFileName & ": " & errorMsg, 3, 1, 0
	    	Exit Sub
	    End If

		badLog.WriteLine Date & " " & Time
		badLog.Close
		
    	WriteLogMsg MSG_LOGMSG_FILEOK & badLogFileName, 1, 1, 0
		
	Else 
		WriteLogMsg MSG_LOGMSG_CLIENTSTATUS & MSG_LASTRESULT_SUCCEED, 1, 1, 0
		g_WshShell.LogEvent 0, MSG_LOGMSG_CLIENTSTATUS & MSG_LASTRESULT_SUCCEED
		g_WshShell.RegWrite registryLocation, MSG_LASTRESULT_SUCCEED, "REG_SZ"

		If lastExecutionResult = False Then
			On Error Resume Next
			Err.Clear
	
			g_fso.DeleteFile badLogFileName, True
		    errorCode = Err.Number
		    errorMsg = Err.Description & " (" & Err.Number & ")"
	    
	   		On Error GoTo 0
	    
		    If errorCode <> 0 Then
		    	WriteLogMsg MSG_LOGMSG_FILEDELETEERROR & badLogFileName & ": " & errorMsg, 2, 1, 0
		    Else
		    	WriteLogMsg MSG_LOGMSG_FILEDELETEOK & badLogFileName, 1, 1, 0
			End If
		End If

	End If				

End Sub

Sub Delay(ByVal delayTime)

	Dim countdown
	
	For countdown = delayTime To 1 Step -1
		WriteLogMsg countdown & "...", 1, 1, 0
		WScript.Sleep(1000)
	Next
	
End Sub

Function GetLastResult(ByRef options)
	Dim registryLocation, lastResult
	Dim errorCode
	
	GetLastResult = True
	registryLocation = DEFAULT_REGISTRY_LOCATION & "\" & DEFAULT_REGISTRY_LASTRESULT_VALUE
	WriteLogMsg MSG_LASTRESULT_VERIFYING & registryLocation, 1, 1, 0
	
	On Error Resume Next
	Err.Clear
	lastResult = g_WshShell.RegRead(registryLocation)
	errorCode = Err.Number
	On Error GoTo 0
	
	If errorCode <> 0 Then
		'Wscript.echo "The errorcode is 0"
		WriteLogMsg MSG_LASTRESULT_NOLASTRESULT, 1, 1, 0
		lastResult = MSG_LASTRESULT_SUCCEED
	Else	
		WriteLogMsg MSG_LASTRESULT_RESULT & lastResult, 1, 1, 0
	End If
	
	If lastResult = MSG_LASTRESULT_FAIL Then
		GetLastResult = False
	End If
End Function

Function OpenConfig(ByRef args, ByRef config)
	Dim configFilename
	OpenConfig = False
	
	'Check for the proper command line arguments
	If Not ( args.Exists ( DEFAULT_CONFIGFILE_PARAMETER ) ) Then
	 	' Print the proper usage and return
	 	WriteLogMsg MSG_OPENCONFIG_NOT_SPECIFIED, 3, 1, 1
		Exit Function
	End If
	
	configFilename = args.Item ( DEFAULT_CONFIGFILE_PARAMETER )
	'Check to make sure the specified config file exists
	If Not g_fso.FileExists ( configFilename ) Then
		WriteLogMsg MSG_OPENCONFIG_DOESNOTEXIST & configFilename, 3, 1, 1
		Exit Function
	End If

	Set config  = CreateObject ( "Msxml2.DOMDocument" )
	' Load the whole XML config file at once
	config.async = False
	config.load ( configFilename )
	
	' Check the file to make sure it is valid XML
	If config.parseError.errorCode <> 0 Then
		WriteLogMsg MSG_OPENCONFIG_PARSEERROR & config.parseError.reason, 3, 1, 0
		Exit Function
	Else
		' Set our XML query language to XPath
		config.setProperty "SelectionLanguage", "XPath"
		WriteLogMsg MSG_OPENCONFIG_OPENED & configFilename, 1, 1, 1
	End If
	OpenConfig = True
End Function

Sub LoadOptions(ByRef config, ByRef options, ByRef parameters)
	Dim optionsNodes, optionNode
	Dim optionName, optionValue
	Dim paramNodes, paramNode
	Dim paramName, paramValue 
	Dim Verify 'As Boolean 
	
	Verify = true
	Set optionsNodes = config.documentElement.selectNodes ( "/Startup/Option" )
	WriteLogMsg MSG_LOADOPTIONS_STARTED, 1, 1, 0

	For Each optionNode In optionsNodes
		optionName = optionNode.getAttribute("Name")
		optionValue = optionNode.text

		If Len(optionValue) > 0 Then
			
			SanitizeOptionInput optionName, optionValue
			
			Verify = VerifyInput(optionName, optionValue)
			If Verify = false then
				WriteLogMsg MSG_VERIFYINPUT_FAILED, 1, 1, 0
			'else
				'WriteLogMsg MSG_VERIFYINPUT_OK, 1, 1, 0
			End if
			'end if 
			'WriteLogMsg MSG_VERIFYINPUT_OK, 1, 1, 0
			
			options.Add optionName, optionValue
			WriteLogMsg MSG_LOADOPTIONS_OPTIONLOADED & optionName & ": '" & optionValue & "'", 1, 1, 0
			
		End If
		
	Next
	
	Set paramNodes = config.documentElement.selectNodes ( "/Startup/InstallProperty" )
	For Each paramNode In paramNodes
		paramName = paramNode.getAttribute("Name")
		paramValue = paramNode.text
		If Len(paramValue) > 0 Then
			parameters.Add paramName, paramValue
			WriteLogMsg MSG_LOADOPTIONS_PARAMLOADED & paramName & ": '" & paramValue & "'", 1, 1, 0
		End If
	Next
End Sub

Sub WriteLogMsg(msg, msgtype, echomsg, eventlog) 
	Dim outmsg, theTime, logfile
	theTime = Time
  
	outmsg = "<![LOG[" & msg & "]LOG]!><time="
	outmsg = outmsg & """" & DatePart("h", theTime) & ":" & DatePart("n", theTime) & ":" & DatePart("s", theTime) & ".000+0"""
	outmsg = outmsg & " date=""" & Replace(Date, "/", "-") & """"
	outmsg = outmsg & " component=""" & WScript.ScriptName & """ context="""" type=""" & msgtype & """ thread="""" file=""" & WScript.ScriptName & """>"
	
	If msgtype = 6 Then
		outmsg =  msg 
	End If
	
	On Error Resume Next
	Set logfile = g_fso.OpenTextFile(g_logPathandName, 8, True)
	logfile.WriteLine outmsg
	logfile.Close
	On Error GoTo 0

	'If echomsg = 1 Then
		'WScript.Echo msg
	'End If
	
	If eventlog = 1 Then
	 g_WshShell.LogEvent 0, DEFAULT_EVENTLOG_PREFIX & msg
	End If
End Sub

Sub WriteFinalLogMsg(ByRef options)
	Dim registryLocation
	Dim logfile
	Dim logFileSize, maxLogFileSize
	Dim finishTime
	
	registryLocation = DEFAULT_REGISTRY_LOCATION & "\" & DEFAULT_REGISTRY_LOGLOCATION_VALUE
	g_WshShell.RegWrite registryLocation, g_logPathandName, "REG_SZ"
	finishTime = Now
	
	WriteLogMsg MSG_MAIN_FINISH & finishTime, 1, 1, 0
	WriteLogMsg MSG_ELAPSED_TIME & DateDiff ("h", g_startTime, finishTime) & ":" & DateDiff ("n", g_startTime, finishTime) & ":" & DateDiff ("s", g_startTime, finishTime), 1, 1, 0
	WriteLogMsg MSG_DIVIDER, 1, 0, 0
	
	Set logfile = g_fso.GetFile(g_logPathandName)
	
	'Write Blank lines at end of script
	WriteLogMsg MSG_EndOfScript, 6, 0, 0
	WriteLogMsg MSG_BlankSpace, 6, 0, 0
	WriteLogMsg MSG_BlankSpace, 6, 0, 0
	WriteLogMsg MSG_BlankSpace, 6, 0, 0
	WriteLogMsg MSG_BlankSpace, 6, 0, 0
	
	logFileSize = logfile.Size / 1024
	maxLogFileSize = GetOptionValue(OPTION_MAXLOGFILE_SIZE, DEFAULT_MAXLOGFILE_SIZE, options)
	
	If logFileSize > maxLogFileSize Then
		g_fso.CopyFile g_logPathandName, g_logPathandName & ".old", True
		g_fso.DeleteFile g_logPathandName, True
	End If
	
	g_WshShell.LogEvent 0, MSG_MAIN_FINISH & g_logPathandName
	MoveFinalLogToLogShare g_logPathandName, options
	g_fso.DeleteFile g_logPathandName, True
End Sub

Sub WriteOpeningBlock
	Dim registryLocation
	Dim logfile, logpath
	Dim userEnv
	Dim errorCode
	Dim logFileSize, maxLogFileSize

	Set userEnv = g_WshShell.Environment("Process")
	logpath = userEnv("TEMP")
	registryLocation = DEFAULT_REGISTRY_LOCATION & "\" & DEFAULT_REGISTRY_LOGLOCATION_VALUE
	
	On Error Resume Next
	Err.Clear
	g_logPathandName = g_WshShell.RegRead (registryLocation)
	errorCode = Err.Number                               
	On Error GoTo 0
	
	'Open logfile
	If errorCode = 0 Then
		On Error Resume Next
		Set logfile = g_fso.OpenTextFile(g_logPathandName, 8, True)
		errorCode = Err.Number
		On Error GoTo 0
		
		If errorCode <> 0 Then
			g_logPathandName = logpath & "\" & WScript.ScriptName & ".log"
		Else
			logfile.Close
		End If
	Else
		g_logPathandName = logpath & "\" & WScript.ScriptName & ".log"
	End If
	
	'Begin the log
	WriteLogMsg MSG_DIVIDER, 1, 1, 0
	WriteLogMsg MSG_MAIN_BEGIN & g_startTime, 1, 1, 1			
End Sub

Function LastRunOK(ByRef options)
	Dim registryLocation, lastRunTime, lastRunInterval, minimumRunInterval
	Dim errorCode
	
	registryLocation = DEFAULT_REGISTRY_LOCATION & "\" & DEFAULT_REGISTRY_LASTRUN_VALUE
	WriteLogMsg MSG_LASTRUN_VERIFYING & registryLocation, 1, 1, 0
	
	On Error Resume Next
	Err.Clear
	lastRunTime = g_WshShell.RegRead(registryLocation)
	errorCode = Err.Number
	On Error GoTo 0
	
	minimumRunInterval = CInt(GetOptionValue(OPTION_DEFAULT_RUNINTERVAL, DEFAULT_RUN_INTERVAL, options))
	If errorCode <> 0 Then
		WriteLogMsg MSG_LASTRUN_NOLASTRUN, 1, 1, 0
		lastRunInterval = minimumRunInterval + 1
	Else	
		lastRunInterval = DateDiff("h", lastRunTime, Now) 
		WriteLogMsg MSG_LASTRUN_TIME & lastRunTime, 1, 1, 0
	End If
	
	If lastRunInterval < minimumRunInterval Then
		LastRunOK = False
		'Wscript.echo "Last run not ok"
		WriteLogMsg MSG_LASTRUN_TIMENOTOK & minimumRunInterval, 1, 1, 0
	Else
		g_WshShell.RegWrite registryLocation, Now, "REG_SZ"
		'Wscript.echo "Last run ok"
		LastRunOK = True
		'LastRunOK = False
	End If
End Function

Function Check_OSWMI (ByRef options)
	Dim errorCode, errorMsg
	Dim fixScript, fixScriptOptions, fixScriptPath
	Dim fixScriptAsynch
	Dim wmiOK
	
	'Wscript.echo "Check_OSWMI starting now."
	wmiOK = CheckWMIConnectivity(options)
	'Wscript.echo "The value of wmiOk is " & wmiOK
	Check_OSWMI = wmiOK
	If wmiOK = True Then
		'Wscript.echo "WMI_OK, Check WMI is now running"
		fixScript = GetOptionValue(OPTION_WMISCRIPT, "0", options) 
		fixScriptAsynch = GetOptionValue(OPTION_WMISCRIPT_ASYNCH, DEFAULT_WMISCRIPT_ASYNCH, options)
		fixScriptOptions = GetOptionValue(OPTION_WMISCRIPTOPTIONS, "", options)
		
		If fixScript <> "0" Then
			'fixScriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "") & fixScript
			fixScriptPath = fixScript
			'Wscript.echo "The value of fixScript path is " & fixScriptPath
			If Not g_fso.FileExists(fixScriptPath) Then
		  		WriteLogMsg MSG_WMISCRIPT_NOTFOUND & fixScriptPath, 3, 1, 0
			ElseIf fixScriptAsynch = "0" Then
		  	WriteLogMsg MSG_WMISCRIPT_EXECUTING & fixScriptPath, 1, 1, 0
		  Else
		  	WriteLogMsg MSG_WMISCRIPT_EXECUTINGASYNCH & fixScriptPath, 1, 1, 0
		  End If
		  
		  If fixScriptOptions <> "" Then
		  	fixScriptOptions = Replace(fixScriptOptions, "%logpath%", options.Item(OPTION_ERROR_LOCATION))
		  	WriteLogMsg MSG_WMISCRIPT_OPTIONS & fixScriptOptions, 1, 1, 0
			End If
		  
			On Error Resume Next
			Err.Clear
			'Wscript.echo "The WMIDiag script is being run now" 
			'Wscript.echo "The value of fixScript optins is " & fixScriptOptions
			g_WshShell.Run "cscript.exe " & Chr(34) & fixScriptPath & Chr(34) & " " & fixScriptOptions, 0, Eval(fixScriptAsynch = "0")
			On Error GoTo 0
			
		 	If errorCode <> 0 Then
		  	WriteLogMsg MSG_WMISCRIPT_ERROR & errorCode, 3, 1, 0
			ElseIf fixScriptAsynch <> "0" Then
		  	WriteLogMsg MSG_WMISCRIPT_SUCCESSASYNCH, 1, 1, 0
			Else
				WriteLogMsg MSG_WMISCRIPT_SUCCESS, 1, 1, 0
				If fixScriptAsynch = "0" Then
				Check_OSWMI = CheckWMIConnectivity(options)
				End If
			End If
		End If
	End If
End Function

Function CheckWMIConnectivity (ByRef options)
	Dim wmi
	Dim errorCode, errorMsg

	On Error Resume Next
	Err.Clear
	Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	errorCode = Err.Number
	errorMsg = Err.Description & " (" & Err.Number & ")"
	
	If errorCode <> 0 Then
		CheckWMIConnectivity = False
		WriteLogMsg MSG_CHECKWMI_ERROR & errorMsg, 3, 1, 0
	Else
		CheckWMIConnectivity = True
		WriteLogMsg MSG_CHECKWMI_SUCCESS, 1, 1, 0
	End If
	
	Set wmi = Nothing
End Function

Function CheckServices(ByRef config)
	Dim wmi, errorCode, errorMsg
	Dim serviceCheckNodes, serviceToCheck, serviceName, expectedServiceState, expectedServiceStartMode, enforce
	Dim returnCode
	
	CheckServices = True
	Set serviceCheckNodes = config.documentElement.selectNodes ( "/Startup/ServiceCheck" )
	
	On Error Resume Next
	Err.Clear
	Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	errorCode = Err.Number
	errorMsg = Err.Description & " (" & Err.Number & ")"
	On Error GoTo 0
	
	If errorCode <> 0 Then
		WriteLogMsg MSG_CHECKWMI_ERROR & errorMsg, 3, 1, 0
		CheckServices = False
	Else
		WriteLogMsg MSG_CHECKSERVICE_START, 1, 1, 1
		For Each serviceToCheck In serviceCheckNodes
			serviceName = serviceToCheck.getAttribute("Name")
			expectedServiceState = serviceToCheck.getAttribute("State")
			expectedServiceStartMode = serviceToCheck.getAttribute("StartMode")
			enforce = serviceToCheck.getAttribute("Enforce")
			If Not CheckService(wmi, serviceName, expectedServiceState, expectedServiceStartMode, enforce) Then
				CheckServices = False
			End If
		Next
	End If
	Set wmi = Nothing
End Function

Function CheckService(ByRef wmi, serviceName, expectedServiceState, expectedServiceStartMode, enforce)
	Dim service
	Dim msg
	Dim serviceStatus
	Dim returnCode, errorCode, errorMsg
	Dim newStartMode
	
	serviceStatus = 1
	CheckService = True
	
	On Error Resume Next
	Err.Clear
	Set service = wmi.Get("Win32_Service.Name='" & serviceName & "'")
	errorCode = Err.Number
	errorMsg = Err.Description & " (" & Err.Number & ")"
	On Error GoTo 0
	
	If errorCode <> 0 Then
		msg = " *" & serviceName & MSG_NOTFOUND & ": " & errorMsg
		WriteLogMsg msg & MSG_NOTFOUND, 2, 1, 0
		CheckService = False
		Exit Function
	End If
	
	msg = " *" & service.Name
	If IsObject(service) Then
		msg = msg & MSG_FOUND & " (" & service.State & "," & service.StartMode & ")"
		If service.StartMode <> expectedServiceStartMode Then
			msg = msg & MSG_CHECKSERVICE_STARTMODE & expectedServiceStartMode			
			If enforce = "True" Then
				If expectedServiceStartMode = "Auto" Then
				 newStartMode = "Automatic"
				Else
					newStartMode = expectedServiceStartMode
				End If
				
				returnCode = service.ChangeStartMode(newStartMode)
				If returnCode = 0 Then
					msg = msg & MSG_CHECKSERVICE_STARTMODEOK & newStartMode
				Else
					msg = msg & MSG_CHECKSERVICE_STARTMODEFAIL & returnCode
					serviceStatus = 3
					CheckService = False
				End If
			End If
		End If
		
		If service.State <> expectedServiceState Then
			msg = msg & MSG_CHECKSERVICE_STATE & expectedServiceState
			If enforce = "True" Then
				If expectedServiceState = "Running" Then
					returnCode = service.StartService()
					WScript.Sleep(15000)
					
					If returnCode = 0 Then
						msg = msg &  MSG_CHECKSERVICE_STARTEDOK
					Else
						msg = msg &  MSG_CHECKSERVICE_STARTEDFAIL & returnCode
						serviceStatus = 3
						CheckService = False
					End If

				ElseIf expectedServiceState = "Stopped" Then
					returnCode = service.StopService()
					If returnCode = 0 Then
						msg = msg &  MSG_CHECKSERVICE_STOPPEDOK
					Else
						msg = msg &  MSG_CHECKSERVICE_STOPPEDFAIL & returnCode
						serviceStatus = 3
						CheckService = False
					End If
				End If						
			End If	
		End If	
	Else
		msg = msg & MSG_NOTFOUND
		serviceStatus = 2
	End If
	
	If serviceStatus = 1 Then
		msg = msg & MSG_OK
	Else
		msg = msg & MSG_NOTOK
	End If
	WriteLogMsg msg, serviceStatus, 1, 0
End Function

Function Conform_AdminShareOptions()
	Dim wmi, adminShare, adminShareRegValue, errorMsg, errorCode
	Dim msg, status
	
	adminShareRegValue = 1
	status = 1
	WriteLogMsg MSG_CHECKADMINSHARE_START, 1, 1, 1

	Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set adminShare = wmi.Get("Win32_Share.Name='ADMIN$'")
	msg = " *Admin$"
	
	If IsObject(adminShare) Then
		msg = msg & MSG_FOUND
	Else
		msg = msg &  MSG_NOTFOUND
		status = 3
		
		On Error Resume Next
		Err.Clear
		adminShareRegValue = g_WshShell.RegRead("HKLM\SYSTEM\CurrentControlSet\Services\LanManServer\Parameters\AutoShareWks")
		errorCode = Err.Number
		errorMsg = Err.Description & " (" & Err.Number & ")"
		On Error GoTo 0

		If errorCode = 0 And adminShareRegValue = 0 Then
			On Error Resume Next
			Err.Clear
			g_WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\LanManServer\Parameters\AutoShareWks", 1, "REG_DWORD"
			errorCode = Err.Number
			errorMsg = Err.Description & " (" & Err.Number & ")"
			On Error GoTo 0

			If errorCode = 0 Then
				msg = msg & MSG_CHECKADMINSHARE_SETSUCCESS
			Else
				msg = msg & MSG_CHECKADMINSHARE_SETFAIL & errorMsg
				status = 3
			End If
		End If
	End If
	
	If status = 1 Then
		msg = msg & MSG_OK
		Conform_AdminShareOptions = True
	Else
		msg = msg & MSG_NOTOK
		Conform_AdminShareOptions = False
	End If
	
	WriteLogMsg msg, status, 1, 0
	Set wmi = Nothing
End Function

Function Confirm_RegistryOptions(ByRef config)
	Dim registryCheckNodes, registryValueToCheck
	Dim regKey, regValue, expectedValue, valueType, enforce
	Dim errorCode
	Dim actualValue
	Dim msg
	Dim regStatus
	
	Confirm_RegistryOptions = True
	Set registryCheckNodes = config.documentElement.selectNodes ( "/Startup/RegistryValueCheck" )
	WriteLogMsg MSG_CHECKREGISTRY_START, 1, 1, 1
	
	For Each registryValueToCheck In registryCheckNodes
		regKey = registryValueToCheck.getAttribute("Key")
		regValue = registryValueToCheck.getAttribute("Value")
		expectedValue = registryValueToCheck.getAttribute("Expected")
		enforce = registryValueToCheck.getAttribute("Enforce")
		valueType = registryValueToCheck.getAttribute("Type")
		
		regStatus = 1
		If valueType = "REG_DWORD" Then
			expectedValue = CInt(expectedValue)
		End If
		
		On Error Resume Next
		Err.Clear
		actualValue = g_WshShell.RegRead(regKey & "\" & regValue)
		errorCode = Err.Number
		On Error GoTo 0
		msg = " *" & regKey & "\" & regValue
		
		If errorCode <> 0 Then
			msg = msg & MSG_NOTFOUND
			regStatus = 2
		Else
			msg = msg & MSG_FOUND & " (" & actualValue & ")"
			If actualValue <> expectedValue Then
				msg = msg & MSG_CHECKREGISTRY_EXPECTED & expectedValue
				regStatus = 2
				Confirm_RegistryOptions = False
			Else
				enforce = False
			End If
		End If
		
		If enforce = "True" Then
			On Error Resume Next
			Err.Clear
			g_WshShell.RegWrite regKey & "\" & regValue, expectedValue, valueType
			errorCode = Err.Number
			On Error GoTo 0
	
			If errorCode = 0 Then
				msg = msg & MSG_CHECKREGISTRY_ENFORCEOK
				Confirm_RegistryOptions = True
			Else
				msg = msg &  MSG_CHECKREGISTRY_ENFORCEFAIL
				regStatus = 3
				Confirm_RegistryOptions = False
			End If
		End If
		
		If Confirm_RegistryOptions = True Then
			msg = msg & MSG_OK
		Else
			msg = msg & MSG_NOTOK
		End If
		WriteLogMsg msg, regStatus, 1, 0
	Next
End Function

Function Conform_LocalAdminOptions(ByRef options)
	Dim adminAccountName
	Dim localAdminGroupName, localAdminGroup
	Dim errorCode, errorMsg, msg, status
	Dim network
	
	status = 1
	Conform_LocalAdminOptions = True
	If options.Exists(OPTION_LOCALADMIN) Then
		WriteLogMsg MSG_CHECKLOCALADMIN_START, 1, 1, 1
		
		adminAccountName = options.Item(OPTION_LOCALADMIN)
		localAdminGroupName = GetOptionValue(OPTION_LOCALADMIN_GROUP,DEFAULT_LOCALADMIN_GROUP, options)
		msg = " *" & localAdminGroupName
		Set network = WScript.CreateObject("WScript.Network")
		
		On Error Resume Next
		Err.Clear
		Set localAdminGroup = GetObject("WinNT://" & network.ComputerName & "/" & localAdminGroupName & ",group")
		errorCode = Err.Number
		errorMsg = Err.Description & " (" & Err.Number & ")"
		On Error GoTo 0
		
		If errorCode = 0 Then
			msg = msg & MSG_FOUND & "...(" & adminAccountName & ")"
		    If localAdminGroup.IsMember("WinNT://" & adminAccountName) Then
		    	msg = msg & MSG_CHECKLOCALADMIN_ALREADYMEMBER
		    Else
					On Error Resume Next
					Err.Clear
	        localAdminGroup.Add ("WinNT://"& adminAccountName)
	        errorCode = Err.Number
	        errorMsg = Err.Description & " (" & Err.Number & ")"
	        On Error GoTo 0
	        	
	        If errorCode = 0 Then
	        	msg = msg & MSG_CHECKLOCALADMIN_ADDMEMBEROK
	        Else
	        	msg = msg & MSG_CHECKLOCALADMIN_ADDMEMBERFAIL & errorMsg
	        	status = 3
	        End If
				End If
		Else
			msg = msg & MSG_NOTFOUND
			status = 2
		End If
		
		If status = 1 Then
			msg = msg & MSG_OK
	 		Conform_LocalAdminOptions = True
		Else
			msg = msg & MSG_NOTOK
	 		Conform_LocalAdminOptions = False
		End If
		WriteLogMsg msg, status, 1, 0
	End If
End Function

Sub CheckCache(ByRef options)
	Dim uiResManager, cache
	Dim errorCode, errorMsg
	Dim desiredCacheSize
	
	WriteLogMsg MSG_CHECKCACHE_START, 1, 1, 1

	On Error Resume Next
	Err.Clear
	Set uiResManager = CreateObject("UIResource.UIResourceMgr")
	errorCode = Err.Number
	errorMsg = Err.Description & " (" & Err.Number & ")"
	On Error GoTo 0
	
	If errorCode <> 0 Then
	   	WriteLogMsg MSG_CHECKCACHE_CREATEFAIL & errorMsg, 2, 1, 0
	    Exit Sub
	End If
	
	On Error Resume Next
	Err.Clear
	Set cache = uiResManager.GetCacheInfo
	errorCode = Err.Number	
	errorMsg = Err.Description & " (" & Err.Number & ")"
	On Error GoTo 0

	If errorCode <> 0 Then
	   Set uiResManager = Nothing
	   WriteLogMsg MSG_CHECKCACHE_CACHEFAIL & errorMsg, 2, 1, 0
	   Exit sub
	End If
	
	desiredCacheSize = CInt(GetOptionValue(OPTION_CACHESIZE, DEFAULT_CACHESIZE, options))
	If cache.TotalSize <> desiredCacheSize Then
		cache.TotalSize = desiredCacheSize
		WriteLogMsg MSG_CHECKCACHE_SETSIZE & desiredCacheSize, 1, 1, 0
	Else
		WriteLogMsg MSG_CHECKCACHE_SIZEOK & cache.TotalSize, 1, 1, 0
	End If
  Set uiResManager = Nothing
End Sub

Sub CheckCacheDuringStartup(ByRef options)
	Dim cache
	Dim errorCode, errorMsg
	Dim cacheSize, desiredCacheSize
	
	WriteLogMsg MSG_CHECKCACHE_START, 1, 1, 1

	On Error Resume Next
	Err.Clear
	Set cache = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\ccm\SoftMgmtAgent:CacheConfig='Cache'")
	errorCode = Err.Number
	errorMsg = Err.Description & " (" & Err.Number & ")"
	On Error GoTo 0

	If errorCode <> 0 Then
	   	WriteLogMsg MSG_CHECKCACHE_WMIFAIL & errorCode, 2, 1, 0
	    Exit Sub
	End If
	
	cacheSize = cache.Size
	WriteLogMsg " *Current cache size is: " & cacheSize, 1, 1, 1
	desiredCacheSize = CInt(GetOptionValue(OPTION_CACHESIZE, DEFAULT_CACHESIZE, options))
	If cacheSize <> desiredCacheSize Then
		cache.Size = desiredCacheSize
		On Error Resume Next
		Err.Clear
		cache.Put_
		errorCode = Err.Number
		errorMsg = Err.Description & " (" & Err.Number & ")"
		On Error GoTo 0
	
		If errorCode <> 0 Then
		   	WriteLogMsg MSG_CHECKCACHE_WMIWRITEFAIL & errorMsg, 2, 1, 0
		    Exit Sub
		End If
		WriteLogMsg MSG_CHECKCACHE_SETSIZE & desiredCacheSize, 1, 1, 0
	Else
		WriteLogMsg MSG_CHECKCACHE_SIZEOK & cacheSize, 1, 1, 0
	End If
	Set cache = Nothing
End Sub

Function CheckAssignment
	Dim smsClient, siteCode
	Dim errorCode
	
	WriteLogMsg MSG_CHECKASSIGNMENT_START, 1, 1, 1
	On Error Resume Next
	Err.Clear
	Set smsClient = CreateObject ("Microsoft.SMS.Client")
	errorCode = Err.Number
	siteCode = smsClient.GetAssignedSite
	On Error GoTo 0
	
	If Len(siteCode) = 0 Or errorCode <> 0 Then
		WriteLogMsg MSG_CHECKASSIGNMENT_NOTOK & siteCode, 1, 1, 0
		CheckAssignment = False
	Else
		WriteLogMsg MSG_CHECKASSIGNMENT_OK & siteCode, 1, 1, 0
		CheckAssignment = True
	End If

	Set smsClient = Nothing
End Function

Function UninstallClient(ByRef options, ByRef parameters)
	Dim fsp, mp, slp, cacheSize, siteCode
	Dim installPath
	Dim commandLine
	Dim returnCode
	Dim param, paramValue
	Dim installPatchProperty
	Dim errorCode, errorMsg, msg
		
	installPatchProperty = ""
	WriteLogMsg "Uninstalling Client", 1, 1, 0
	
	If options.Exists(OPTION_INSTALLPATH) Then
		cacheSize = GetOptionValue(OPTION_CACHESIZE, DEFAULT_CACHESIZE, options)
		TrimTrailingSlash commandLine
		commandLine = options.Item(OPTION_INSTALLPATH) & "\ccmsetup.exe"
		msg = MSG_INSTALLCLIENT_PATHCHECK & commandLine
		
		If Not g_fso.FileExists(commandLine) Then
			WriteLogMsg msg & MSG_NOTFOUND & " " & commandLine, 3, 1, 0
			UninstallClient = False
			Exit Function		
		End If

		WriteLogMsg msg & MSG_FOUND, 1, 1, 0
		commandLine = commandline & " /uninstall"
		
		WriteLogMsg MSG_INSTALLCLIENT_COMMANDLINE & commandLine, 1, 1, 0
		'returnCode = g_WshShell.Run(commandLine, 0, true)
		returnCode = RunClientInstall(commandLine)
		If returnCode = true Then
		  WriteLogMsg "Install failed: " & returnCode, 3, 1, 0
		  unInstallClient = False
		Else
		  WriteLogMsg "Install succeeded.", 1, 1, 0
		  unInstallClient = True
    End If
	End If
End Function

Function InstallClient(ByRef options, ByRef parameters)
	Dim fsp, mp, slp, cacheSize, siteCode
	Dim installPath
	Dim commandLine
	Dim returnCode
	Dim param, paramValue
	Dim installPatchProperty
	Dim errorCode, errorMsg, msg
		
	installPatchProperty = ""
	WriteLogMsg MSG_INSTALLCLIENT_START, 1, 1, 0
	
	If options.Exists(OPTION_INSTALLPATH) Then
		cacheSize = GetOptionValue(OPTION_CACHESIZE, DEFAULT_CACHESIZE, options)
		TrimTrailingSlash commandLine
		commandLine = options.Item(OPTION_INSTALLPATH) & "\ccmsetup.exe"
		msg = MSG_INSTALLCLIENT_PATHCHECK & commandLine
		
		If Not g_fso.FileExists(commandLine) Then
			WriteLogMsg msg & MSG_NOTFOUND & " " & commandLine, 3, 1, 0
			InstallClient = False
			Exit Function		
		End If

		WriteLogMsg msg & MSG_FOUND, 1, 1, 0
		commandLine = commandLine & " SMSCACHESIZE=" & cacheSize
		
		'get site code
		If options.Exists(OPTION_SITECODE) Then
			commandLine = commandLine & " SMSSITECODE=" & options.Item(OPTION_SITECODE)
		End If
		
		'other props
		If options.Exists(OPTION_OTHERINSTALLPROPS) Then
			commandLine = commandLine & " " & options.Item(OPTION_OTHERINSTALLPROPS)
		End If
		
		'Section responsible for the hotfix
		For Each param In parameters.Keys
			'check if param equals patch
			If Instr(1, param, "PATCH", 1) Then
				If InStr(1, parameters.Item(param), ";", 1) Then
					WriteLogMsg MSG_HOTFIX_MULTIPLE, 1, 1, 0
					installPatchProperty = installPatchProperty & parameters.Item(param)
				Else
					msg = MSG_HOTFIX_FILEVERIFY & parameters.Item(param)
					If g_fso.FileExists(parameters.Item(param)) Then				
						If Len(installPatchProperty) > 0 Then
							installPatchProperty = installPatchProperty & ";"
						End If
						installPatchProperty = installPatchProperty & parameters.Item(param)
						WriteLogMsg msg & MSG_FOUND, 1, 1, 0				
					Else
						WriteLogMsg msg & MSG_NOTFOUND & ": " & errorMsg, 2, 1, 0				
					End If
				End If
			Else
				commandLine = commandLine & " " & param & "=" & parameters.Item(param)
			End If
		Next
		
		'Wscript.echo "HOTFIX is being put into the ccmsetup.exe command line"
		If options.Exists(OPTION_AUTOHOTFIX) Then
			If Len(installPatchProperty) > 0 Then
				installPatchProperty = installPatchProperty & ";"
			End If
			AutoHotfix options.Item(OPTION_AUTOHOTFIX), installPatchProperty
		End If
		If Len(installPatchProperty) > 0 Then
			commandLine = commandLine & " PATCH=""" & installPatchProperty & """"
		End If
		
		WriteLogMsg MSG_INSTALLCLIENT_COMMANDLINE & commandLine, 1, 1, 0
		'returnCode = g_WshShell.Run(commandLine, 0, true)
		returnCode = RunClientInstall(commandLine)
		If returnCode = true Then
		  WriteLogMsg MSG_INSTALLCLIENT_FAILED & returnCode, 3, 1, 0
		  InstallClient = False
		Else
		  WriteLogMsg MSG_INSTALLCLIENT_SUCCESS, 1, 1, 0
		  InstallClient = True
    End If
	End If
End Function

Function GetOptionValue(ByVal optionName, ByVal defaultValue, ByRef options)
	If options.Exists(optionName) Then
		GetOptionValue = options.Item(optionName)
	Else
		GetOptionValue = defaultValue
	End If
End Function

Sub AutoHotfix(ByVal hotfixDirectory, ByRef patchProperty)
	Dim errorCode, errorMsg
	Dim hfDir
	Dim finalDirLocation
	Dim OSVer

	WriteLogMsg MSG_AUTOPATCH_COMMANDLINE & hotfixDirectory, 1, 1, 0
  
  On Error Resume Next
	Err.Clear
	hfDir = g_fso.GetFolder(hotfixDirectory)
	errorCode = Err.Number	
	errorMsg = Err.Description & " (" & Err.Number & ")"
	On Error GoTo 0

	If errorCode <> 0 Then
	   WriteLogMsg MSG_AUTOPATCH_DIRERROR & errorMsg, 2, 1, 0
	   Exit Sub
	End If
	
	'Trim Trailing slashes before checking the architecture
	TrimTrailingSlash hfDir
	
	' need to modify directory based on architecture
	'Wscript.echo "Starting check os now"
	CheckOSType OSVer
	'Wscript.echo "The OSVer being returned is " & OSVer
	Select Case OSVer
		'If 32 bit look in .\i386
		Case 32
			hfDir = hfDir & "\i386\"
		'If 64 bit look in .\x64
		Case 64
			hfDir = hfDir & "\x64\"
	End Select
	'Wscript.echo "The Directory that will be searched in is " & hfDir
	Set finalDirLocation = g_fso.GetFolder(hfdir)
	
	FindHotfixes finalDirLocation, patchProperty
End Sub

Sub FindHotfixes(ByVal directory, ByRef patchProperty)
	Dim file, subfolder
	For Each file In directory.Files
    	If LCase(g_fso.GetExtensionName(file)) = "msp" Then
    		If InStr(1, patchProperty, file.Name, 1) Then
    			WriteLogMsg MSG_HOTFIX_DUPLICATE & file.Name, 2, 1, 0
				Else    		
		  		If Len(patchProperty) > 0 Then
						patchProperty = patchProperty & ";"
					End If
				
					WriteLogMsg MSG_AUTOPATCH_FOUNDHOTFIX & file.Path, 1, 1, 0
					'Wscript.echo "The file path is " & file.path
					patchProperty = patchProperty & file.Path
					'Wscript.echo "The patchProperty is " & patchProperty
				End If
    	End If
	Next
  
  'For Each subfolder in directory.subfolders
	'	FindHotfixes subfolder, patchProperty
	'Next
	'FindHotFixTest options, patchProperty
End Sub

Function InWinPE
	Dim sysEnv, systemDrive
	Set sysEnv = g_WshShell.Environment("PROCESS")
	systemDrive = sysEnv("SYSTEMDRIVE")
	
	InWinPE = Eval(SystemDrive = "X:")
End Function

Function IsAdmin
	On Error Resume Next
  CreateObject("WScript.Shell").RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")
  
  If err.number = 0 Then 
  	IsAdmin = True
  Else
  	IsAdmin = False
  	'WScript.Echo "User is not a local admin."
  End If
  On Error GoTo 0
End Function

Sub MoveFinalLogToLogShare(ByVal g_logPathandName, ByRef options)
	'Variables
	Dim strDate
	Dim paddedDate
	Dim strPadDate
	Dim remoteLog
	Dim localLog
	Dim ReadAllTextFile
	
	'Get Computer name 
	strComputerName = g_WshShell.ExpandEnvironmentStrings("%COMPUTERNAME%")

	'Make the file name of the new log file to be placed on \\winfs\logs
	strLogName = strComputerName & "_" & "ConfigMgrStartupLog.txt"
	
	'Attach the Winfs\logs UNC path to the new log file name.
	g_finalLog  = options.Item(CONFIGMGRLOGPATH) & strLogName
	
	'Open existing log file and prep it for appending
	set remoteLog = g_fso.OpenTextFile(g_finalLog, 8, true, true )
		
	'Open log file on local machine and read the contents to a variable
	set localLog = g_fso.OpenTextFile(g_logPathandName, 1, True)
	ReadAllTextFile = localLog.ReadAll
		
	'write the contents of the client log file into the final log file share 
	remoteLog.Write(ReadAllTextFile)
		
	'Close the winfs\logs log file and C:\windows\temp log file
	remoteLog.close
	localLog.close
End Sub

Sub GetOSVersionAndArch(ByRef OSVer, ByRef OSArch)
	'Wscript.echo "The CheckOS sub is now running"
	Dim oArch
	Dim sVer
	Dim colObjOS
	Dim iWidth
	Dim oOS
	Dim OSName
	
	strComputerName = g_WshShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
	'Get Operating system info
	'Connect the wmiservices to target computer 
	Set objWMIService = GetObject("winmgmts:\\" & strComputerName & "\root\cimv2")
	Set colObjOS = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
	
	For Each oOS in colObjOS
		sVer = oOS.version
		OSName = oOS.Caption
		If(Left(sVer,3)) >= 6.0 Then 'If Win7
			oArch = oOS.OSArchitecture
			'Wscript.Echo "The Architecture of the proc is sArch" & 
			Select Case oArch
				Case "32-bit"
					OSArch = 32
				Case "64-bit"
					OSArch = 64
			End Select
			OSVer = "win7"
		End If
		If(Left(sVer,3)) = 5.2 Then '2003 or xp(x64)
			oArch = oOS.OSArchitecture
			'Wscript.Echo "The Architecture of the proc is sArch" & 
			Select Case oArch
				Case "32-bit"
					OSArch = 32
					OSVer = "2003"
				Case "64-bit"
					OSArch = 64
					if Left(OSNAME, 10) = "Windows XP" Then
						OSVer = "xp"
					else
						OSVer = "2003"
					End If
			End Select
		Else 'XP'
			Set colObjProc = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
			For Each objProc in colObjProc
				iWidth = objProc.AddressWidth				
				If iWidth = 32 or iWidth = 64 Then
					OSArch = iWidth
					OSVer = "xp"
				End If
			Next
			OSVer = "xp"
		End If	
	Next

	'Wscript.Echo "The OS Architecture is " & OSVer & "-bit."
End Sub 

Sub CheckOSType(ByRef OSArch)
	'Wscript.echo "The CheckOS sub is now running"
	Dim oArch
	Dim sVer
	Dim colObjOS
	Dim iWidth
	Dim oOS
	
	strComputerName = g_WshShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
	'Get Operating system info
	'Connect the wmiservices to target computer 
	Set objWMIService = GetObject("winmgmts:\\" & strComputerName & "\root\cimv2")
	Set colObjOS = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
	
	For Each oOS in colObjOS
		sVer = oOS.version
		If(Left(sVer,3)) >= 6.0 Then 'If Win7
			oArch = oOS.OSArchitecture
			Select Case oArch
				Case "32-bit"
					OSArch = 32
				Case "64-bit"
					OSArch = 64
			End Select
		Else 'XP
			Set colObjProc = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
			For Each objProc in colObjProc
				iWidth = objProc.AddressWidth				
				If iWidth = 32 or iWidth = 64 Then
					OSArch = iWidth
				End If
			Next
		End If
	Next
End Sub 

Function IsLegacyOS()
	
	Dim oArch
	Dim sVer
	Dim colObjOS
	Dim iWidth
	Dim oOS
	Dim bLegacyOS
	
	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	Set colObjOS = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
	
	For Each oOS in colObjOS
		sVer = oOS.version
		If(Left(sVer,3)) >= 6.0 Then 'If Win7 then false
			bLegacyOS = false
		Else 'XP or 2003
			bLegacyOS = true
		End If
	Next
	
	IsLegacyOS = bLegacyOS
End Function

Sub TrimTrailingSlash(ByRef path)
	'Wscript.echo "The TrimTrailingSlash sub is now running"
    If Right(path,1)="\" Then
      path = Left(path,Len(path)-1)
      'Wscript.echo "The value of path is now " & path
    End If
End Sub 

Sub SanitizeOptionInput (ByRef optionName, ByRef optionValue)
	If optionName = "ClientLocation" Then
		TrimTrailingSlash optionValue
		'Wscript.echo "Sanitized input of ClientLocation and it looks like " & optionValue
	ElseIf optionName = "AutoHotFix" Then
		TrimTrailingSlash optionValue
		'Wscript.echo "Sanitized input of AutoHotFix and it looks like " & optionValue
	end if
End Sub

Function VerifyInput (ByVal optionName, ByVal optionValue)
	VerifyInput = true
	
	If optionName = "ClientLocation" Then
		If Not g_fso.FolderExists(optionValue) Then
			VerifyInput = false
		End If
	ElseIf optionName = "AutoHotFix" Then
		If Not g_fso.FolderExists(optionValue) Then
			VerifyInput = false
		End If
	End If
End Function


Function CheckClientNeedsHotfix(ByVal ConfigOptions)
	Dim clientVersion, expectedVersion
	Dim clientProperties, clientProp
	Dim  ccmWMI
	
	'set wmiObj
	Set ccmWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\ccm")
	
	'Get clientProperties from wmi
	Set clientProperties = ccmWMI.ExecQuery("Select * from SMS_Client")
	
	'get expected Client version from xml
	expectedVersion = GetOptionValue(OPTION_AGENTVERSION, DEFAULT_AGENTVERSION, ConfigOptions)
	clientVersion = "0"
	For each clientProp in clientProperties
		if expectedVersion > clientProp.ClientVersion then
			CheckClientNeedsHotfix = true
		else
			CheckClientNeedsHotfix = false
		End if
	Next
	
	'Check to see if client is up to date
	'If CheckClientVersion(expectedVersion, clientVersion) = False Then
	'	Wscript.echo "The Client version is old and Needs to be updated"
	'	CheckOSNeedsHotfix = true
		'write to log file
	'Else
	'	Wscript.echo "The Client Version is up to date"
	'	CheckOSNeedsHotfix = false
		'write to log file
	'End If
End Function

Function CheckOSNeedsCertHotfix(ByVal ConfigOptions)
	Dim expectedVersion
	Dim colEngObj, engObj
	Dim WMIService
	Dim strComputerName
	Dim HotFixID
	Dim bHotfixNeeded
	Dim bLegacyOS
	Dim msg
	
	Set WMIService = GetObject("winmgmts:\\.\root\cimv2")
	Set colEngObj = WMIService.ExecQuery("SELECT * FROM Win32_QuickFixEngineering")
	
	bHotfixNeeded = true
	
	'check for legacy os
	bLegacyOS = IsLegacyOS
	If bLegacyOS = false Then
		msg = "Client OS is newer than XP\2003 and does not need a certificate hotfix."
		writelogmsg msg, 1, 1, 0
		bHotfixNeeded = false
	End If

	'check if hotfix is installed if legacy OS
	If bLegacyOS = True Then
		bHotfixNeeded = true
		For each engObj in colEngObj'
			If engObj.HotFixID = ConfigOptions.Item(CERTHOTFIXID) Then
				msg = "The certificate hotfix is installed on the client."
				writelogmsg msg, 1, 1, 0
				bHotfixNeeded = false
				Exit For
			End If
		Next
		If bHotfixNeeded = True Then
			msg = "The client OS is considered  ""legacy"", but does not have the following required certificate hotfix installed: " & ConfigOptions.Item(CERTHOTFIXID)
			writelogmsg msg, 1, 1, 0
		End If
	End If
	
	CheckOSNeedsCertHotfix = bHotfixNeeded
End Function


Function InstallLegacyOSHotfix(ByVal ConfigOptions, ByVal OSVer, ByVal OSArch)
	Dim HotFixPath
	Dim commandLine
	Dim returnCode
	Dim msg
	Dim bContinue	
	Dim bFail
	Dim sHFXFilename
	Dim sSilentHFXArgs
	Dim Results
	'Wscript.echo "The OSVer is " & OSVer
	'Wscript.echo "The OSArch is " & OSArch
	bContinue = true
	
	Select Case OSVer
		Case "2003"
			Select Case OSArch
				Case 32
					sHFXFilename = ConfigOptions.Item(OPTION_2003_X32_CERTHOTFIX)
				Case 64
					sHFXFilename = ConfigOptions.Item(OPTION_XP_2003_X64_CERTHOTFIX)
				Case Else
					sHFXFilename = False 
			End Select
		Case "xp"
			Select Case OSArch
				Case 32
					sHFXFilename = ConfigOptions.Item(OPTION_XP_X32_CERTHOTFIX)
				Case 64
					sHFXFilename = ConfigOptions.Item(OPTION_XP_2003_X64_CERTHOTFIX)
					
				Case Else
					sHFXFilename = False
			End Select
		Case Else
			bFail = True
	End Select
	
	
	If sHFXFilename = False Then
		msg = "ERROR: Failed to determine hotfix filename. Check that the hotfix filename options are specified in your XML config file."
		WriteLogMsg msg, 1, 1, 0
		bContinue = False
	End If
	
	'build the install command & run installer
	If bContinue = True Then
		'build string
		sSilentHFXArgs = "/quiet" 
		HotFixPath = ConfigOptions.Item(optLegacyCertHotFixpath)
		TrimTrailingSlash(HotFixPath)
		commandLine = commandLine & HotFixPath & "\" & sHFXFilename & " " & sSilentHFXArgs
		'run command
		WriteLogMsg MSG_INSTALLCLIENT_COMMANDLINE & commandLine, 1, 1, 0
		returnCode = g_WshShell.Run(commandLine, 0, true)
		'Log output
		If returnCode <> 0 Then
		  Msg = "ERROR: Hotfix install returned error code: """ & returnCode & """."
		  WriteLogMsg msg, 3, 1, 0
		Else
		  Msg = "Hotfix install returned error code 0 (success)."
		  WriteLogMsg msg, 1, 1, 0
	  End If
	End If
	
	'Check if Cert hot fix needs installed again
	If CheckOSNeedsCertHotfix(ConfigOptions) = False Then
		Results = True
	Else
		Results = False
	End If
	InstallLegacyOSHotfix = Results
End Function 

Function RunClientInstall(ByVal commandLine)
 	Dim oWMI
 	Dim ccmProcess
 	Dim i
	Dim timeoutSeconds
	Dim msg
	Dim bResults
	Dim ccmPID
	Dim bSetupRunning
	Dim colProcesses
	Dim objWMIService
	Dim objProcess
	
	bResults = False
 	Set oWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
  
  'Run the command line
 	Set ccmProcess = g_WshShell.Exec(commandLine)
 	ccmPID = ccmProcess.ProcessID
 	
  'loop to make sure ccmsetup.exe process ends before continuing on
 	timeoutSeconds = 10 * 60 '10mins * 60 sec\min	
	bSetupRunning = True
 	i = 0
 	Do while i < timeoutSeconds
 		'wscript.echo "loop #" & i
 		Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
		Set colProcesses = objWMIService.ExecQuery ("Select * from Win32_Process")
		bSetupRunning = False
		For each objProcess in colProcesses 
			If objProcess.Name = "ccmsetup.exe" Then
				bSetupRunning = True
 			End If
 		Next
 		If bSetupRunning = False Then
 			Exit Do
 		End If
 		wscript.sleep 1000
 		i = i + 1
 	Loop
 	
 	If (i > timeoutSeconds) then
  	msg = " *ERROR: The install client function timed out."
  	bResults = False
  End If
  RunClientInstall = True
End Function


''' Begin

WriteOpeningBlock
Main
