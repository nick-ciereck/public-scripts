'This script is from http://www.microsoft.com/technet/scriptcenter/resources/qanda/jul07/hey0727.mspx
'Adapted by John Puskar on 03.26.08
'Option Explicit
On Error Resume Next

Dim objSysInfo, objNetwork, strUserPath, objUser
Dim strGroup, strGroupPath, objGroup, strGroupName
Dim arrComputerName, arrOU, strName, strComputerOU
Dim strCurrentUser, strCurrentPrinter
Dim arrOldPrinters(200) 

'Grab all currently mapped printers
Set objSysInfo = CreateObject("ADSystemInfo")
Set objNetwork = CreateObject("Wscript.Network")
Set oPrinters = objNetwork.EnumPrinterConnections

'Lab Printers
arrOldPrinters(0) = "\\Print\NW2105 - Xerox Phaser 8550"
arrOldPrinters(1) = "\\Print\MP3033 - HP LaserJet 4000"
arrOldPrinters(2) = "\\Print\NW2105 - HP Color LaserJet CP3505DT"
arrOldPrinters(3) = "\\Print\NW2105 - HP Color LaserJet CP3505DP"
arrOldPrinters(4) = "\\Print\NW1118 - Xerox Phaser 8550DT"
arrOldPrinters(5) = "\\Print-server\HP4100PCL"
arrOldPrinters(6) = "\\Print-server\HP4100PS"
arrOldPrinters(7) = "\\Print-server\2105nw-phaser"
arrOldPrinters(8) = "\\Print-server\2105 NW PCL"
arrOldPrinters(9) = "\\Print-server\2105 NW PS"
arrOldPrinters(10) ="\\Print-server\2105NW"
arrOldPrinters(11) ="\\Print-server\2105NW PCL"
arrOldPrinters(12) ="\\Print-server\2105NWP"
arrOldPrinters(13) ="\\Print-server\old.donotuse.2105nw-phaser"
arrOldPrinters(14) ="\\Print-server\old.donotuse.2105NWP"
arrOldPrinters(15) ="\\Print-server\old.donotuse.2105NW PCL"
arrOldPrinters(16) ="\\Print-server\old.donotuse.2105NW"
arrOldPrinters(17) ="\\Print-server\old.donotuse.2105 NW PS"
arrOldPrinters(18) ="\\Print-server\old.donotuse.2105 NW PCL"
arrOldPrinters(19) ="\\Print-Server\2047MP"
arrOldPrinters(20) ="\\Print-server\0008mp-colorPS"
arrOldPrinters(21) ="\\Print-server\0008MP-hp4000"
arrOldPrinters(22) ="\\Print-Server\undergrad-organic"
arrOldPrinters(23) ="\\Print-Server\MassPS"
arrOldPrinters(24) ="\\Print-Server\MassPCL"
arrOldPrinters(25) ="\\Print-Server\HP4100PS"
'Faculty Printers
arrOldPrinters(26) ="\\Print-server\EL4011 - HP LaserHet 1320"
arrOldPrinters(27) ="\\Print-server\EL4011 - HP LaserJet 1320"
arrOldPrinters(28) ="\\Print\EL4087 - HP LaserJet 1320"
arrOldPrinters(29) ="\\print\MP3051 - HP LaserJet 2200"
arrOldPrinters(30) ="\\Print-server\biochem"
arrOldPrinters(31) ="\\Print-server\HPLJ3550Color"
arrOldPrinters(32) ="\\Print-server\400JL-HP2015 PCL"
arrOldPrinters(33) ="\\Print-server\400JL-HP2015 PS"
arrOldPrinters(34) ="\\Print-server\old.400JL-HP2015 PS"
arrOldPrinters(35) ="\\Print-server\old.400JL-HP2015 PS"
arrOldPrinters(36) ="\\Print\EL3040 - Lexmark C530"
arrOldPrinters(37) ="\\Print-Server\305JL"
arrOldPrinters(38) ="\\Print-Server\badjic2300"
arrOldPrinters(39) ="\\Print-Server\3073EL"
arrOldPrinters(40) ="\\Print-server\3051MPPCL"
arrOldPrinters(41) ="\\Print-server\3051MPPS"
arrOldPrinters(42) ="\\Print-server\3051mp-hp3800-pcl"
arrOldPrinters(43) ="\\Print-server\3051mp-hp3800-ps"
arrOldPrinters(44) ="\\Print-server\wu4250PS"
arrOldPrinters(45) ="\\Print-server\wu4250PCL"
arrOldPrinters(46) ="\\Print\EL1033 - HP LaserJet 4050 Series"
arrOldPrinters(47) ="\\Print-Server\0106NWPCL"
arrOldPrinters(48) ="\\Print-server\0106NWPS"
arrOldPrinters(49) ="\\Print-server\ugak"
arrOldPrinters(50) ="\\Print-server\3040EL-Laser PS"
arrOldPrinters(51) ="\\Print-server\3040EL-Laser PCL"
arrOldPrinters(52) ="\\Print-server\3040EL-HP2430PS"
arrOldPrinters(53) ="\\Print-server\3040EL-HP2430PCL"
arrOldPrinters(54) ="\\Print-server\3040EL-LexmarkPS"
arrOldPrinters(55) ="\\Print-server\3040EL-LexmarkPCL"
arrOldPrinters(56) ="\\Print-Server\36MP"
arrOldPrinters(57) ="\\Print-Server\24CE"
arrOldPrinters(58) ="\\Print\MP3035 - HP LaserJet 2015"
arrOldPrinters(59) ="\\Print-server\3035MP-HP2015dnPCL"
arrOldPrinters(60) ="\\Print-server\3035MP-HP2600n"
arrOldPrinters(61) ="\\Print-server\1102nw-hp2015dnPS"
arrOldPrinters(62) ="\\Print-server\1118NW-Color PCL"
arrOldPrinters(63) ="\\Print-server\4112NW"
arrOldPrinters(64) ="\\Print-server\4112NWP"
arrOldPrinters(65) ="\\Print-server\PlatzLJ5"
arrOldPrinters(66) ="\\Print-server\4011el-1320n PCL"
arrOldPrinters(67) ="\\Print-server\4011el-1320n PS"
arrOldPrinters(68) ="\\Print-Server\tsai4200"
arrOldPrinters(69) ="\\Print-Server\315JLColor"
arrOldPrinters(70) ="\\Print-Server\3035MP-HP2600n"
arrOldPrinters(71) ="\\Print-Server\2025MP-HP2105dnPS"
arrOldPrinters(72) ="\\Print-Server\2025MP-HP2105dnPCL"
'Personnel Office Printers
arrOldPrinters(73) ="\\Print-Server\nw1104-lj3055 pcl"
arrOldPrinters(74) ="\\Print-Server\nw1104-lj3055 ps"
arrOldPrinters(75) ="\\Print-Server\1104 NW PCL"
arrOldPrinters(76) ="\\Print-Server\1104 NW PS"
arrOldPrinters(77) ="\\Print-Server\1104"
arrOldPrinters(78) ="\\Print-Server\1104_brother"
arrOldPrinters(79) ="\\Print-Server\personnel-brother"
'Front Office Printers - NW1118
arrOldPrinters(80) = "\\Print\NW1118 - HP Color LaserJet 3505"
arrOldPrinters(81) = "\\Print\NW1118 - Xerox Phaser 8550"
arrOldPrinters(82) = "\\Print-Server\1118NW-Color PCL"
arrOldPrinters(83) = "\\Print-Server\1118NW-Color PS"
arrOldPrinters(84) = "\\Print-Server\1118 NW PCL"
arrOldPrinters(85) = "\\Print-Server\1118 NW PS"
'Front Office Printers - CE100
arrOldPrinters(86) = "\\Print-Server\fermium"
arrOldPrinters(87) = "\\Print-Server\frontoffice"
arrOldPrinters(88) = "\\Print-server\100CE-HP2300N"
'Graduate Studies Office Printers
arrOldPrinters(89) = "\\Print-server\Triton"
arrOldPrinters(90) = "\\Print-Server\Venus"
arrOldPrinters(91) = "\\Print-Server\Venus2"
arrOldPrinters(92) = "\\Print-Server\Mars"
'Other old printers
arrOldPrinters(93) = "\\Print2\NW2105 - HP Color LaserJet CP3505"
arrOldPrinters(94) = "\\Print2\NW2105 - HP LaserJet 4100"
arrOldPrinters(95) = "\\Print2\NW2105 - HP LaserJet P4015dn"
arrOldPrinters(96) = "\\Print\NW1118 - HP Color LaserJet CP3505"
arrOldPrinters(97) = "\\Print\NW2105 - HP LaserJet 4100"
'Old Uniprint Printers
arrOldPrinters(98) = "\\uniprint\chemspool1"
arrOldPrinters(99) = "\\uniprint\mp2047-hp4240spool"
arrOldPrinters(100) = "\\uniprint\mp0008-hp4000spool"

'Remove known old Printers
'only use odd numbers in oPrinters. The collection returned looks like ([even]"printer port",[odd]"printer name") etc...
For i = 1 to oPrinters.Count Step 2
	strCurrentPrinter = oPrinters.Item(i)
	For Each oldPrinter In arrOldPrinters
		If LCase(strCurrentPrinter) = LCase(oldPrinter) Then  'LCase fixes issues with "\\print.." vs. "\\Print.."
			objNetwork.RemovePrinterConnection oldPrinter
		End If
	Next
	
	'
	'REPLACE DIRECTLY UPGRADED PRINTERS
	'
	
	
	'02/14/11
	If LCase(strCurrentPrinter) = LCase("\\print\CE018 - HP Color LaserJet 4700") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\CE018 - HP Color LaserJet 4700"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\CE018 - HP LaserJet 4000") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\CE018 - HP LaserJet 4000"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\CE018 - HP LaserJet 2055dn") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\CE018 - HP LaserJet P2055dn"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\CE100A - HP LaserJet 2055dn") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\CE100A - HP LaserJet 2055dn"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\CE100C - HP LaserJet CP3525dn") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\CE100C - HP LaserJet CP3525dn"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\NW1136 - HP LaserJet 4000") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\NW1136 - HP LaserJet 4000"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\NW2140 - Lotus") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\NW2140 - Lotus"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\NW2146 - HP Color LaserJet 4650") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\NW2146 - HP Color LaserJet 4650"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\JL400 - HP Color LaserJet 3550") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\JL400 - HP Color LaserJet 3550"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\EL3067 - HP LaserJet 5") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\EL3067 - HP LaserJet 5"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\EL3067 - HP LaserJet 5") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\EL3067 - HP LaserJet 5"
	End If
	
	'02/15/11
	If LCase(strCurrentPrinter) = LCase("\\print\CE100G - HP LaserJet 2430") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\CE100G - HP LaserJet 2430"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\CE140 - HP LaserJet 4240") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\CE140 - HP LaserJet 4240"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\CE140 - HP LaserJet P2055dn") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\CE140 - HP LaserJet P2055dn"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\CE144 - HP LaserJet 4000") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\CE144 - HP LaserJet 4000"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\CE145 - HP LaserJet 1200") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\CE145 - HP LaserJet 1200"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\CE160A - Brother HL-5250DN") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\CE160A - Brother HL-5250DN"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\CE280 - HP LaserJet 4 Plus") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\CE280 - HP LaserJet 4 Plus"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\CE380B - HP LaserJet 4 Plus") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\CE380B - HP LaserJet 4 Plus"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\CE431 - HP LaserJet 2430") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\CE431 - HP LaserJet 2430"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\SL3132 - HP Color LaserJet 2600n") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\SL3132 - HP Color LaserJet 2600n"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\SL3132 - HP OfficeJetPro 8500") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\SL3132 - HP OfficeJetPro 8500"
	End If
	
	'02/16/11
	If LCase(strCurrentPrinter) = LCase("\\print\EL0034 - HP LaserJet P2035n") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\EL0034 - HP LaserJet P2035n"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\EL0052 - HP LaserJet 2420") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\EL0052 - HP LaserJet 2420"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\EL0056 - HP LaserJet CP1518ni") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\EL0056 - HP LaserJet CP1518ni"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\EL0056 - HP LaserJet P3005") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\EL0056 - HP LaserJet P3005"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\EL0070 - HP Color LaserJet 2600") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\EL0070 - HP Color LaserJet 2600"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\EL0073 - HP Color LaserJet 2605") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\EL0073 - HP Color LaserJet 2605dn"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\EL1025 - HP LaserJet P2055dn") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\EL1025 - HP LaserJet P2055dn"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\EL1031 - HP LaserJet 4 Plus") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\EL1031 - HP LaserJet 4 Plus"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\EL1033 - HP LaserJet 4050") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\EL1033 - HP LaserJet 4050"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\EL1043 - HP LaserJet 2420") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\EL1043 - HP LaserJet 2420"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\EL1043 - HP LaserJet 4240") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\EL1043 - HP LaserJet 4240"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\EL1045 - HP LaserJet 4250") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\EL1045 - HP LaserJet 4250"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\EL1048 - HP LaserJet CP2025dn") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\EL1048 - HP Color LaserJet CP2025dn"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\EL1058 - HP LaserJet 2300") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\EL1058 - HP LaserJet 2300"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\EL1069 - XRay Laser") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\EL1069 - XRay Laser"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\EL1086 - HP LaserJet P3005") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\EL1086 - HP LaserJet P3005"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\EL2048 - HP LaserJet CP3525") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\EL2048 - HP LaserJet CP3525"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\EL3040 - HP LaserJet 2430") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\EL3040 - HP LaserJet 2430"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\EL3040 - Lexmark C532") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\EL3040 - Lexmark C532"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\EL3040 - Lexmark E360dn") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\EL3040 - Lexmark E360dn"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\EL3040 - Lexmark T640") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\EL3040 - Lexmark T640"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\EL3042 - Lexmark E360dn") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\EL3042 - Lexmark E360dn"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\EL3048 - HP LaserJet p2035n") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\EL3048 - HP LaserJet p2035n"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\EL3069 - HP LaserJet P2035n") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\EL3069 - HP LaserJet P2035n"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\EL4021 - Brother HL2170w") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\EL4021 - Brother HL-2170W"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\EL4031 - HP LaserJet 2300") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\EL4031 - HP LaserJet 2300"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\EL4047 - HP LaserJet 4000") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\EL4047 - HP LaserJet 4000"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\EL4083 - HP LaserJet 1320") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\EL4083 - HP LaserJet 1320"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\EL4086 - HP LaserJet P2035n") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\EL4086 - HP LaserJet P2035n"
	End If
	
	'02/17/11
	If LCase(strCurrentPrinter) = LCase("\\print\JL416 - Lexmark T650") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\JL416 - Lexmark T650"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\JL416 - HP LaserJet P2015") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\JL416 - HP LaserJet P2015"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\JL400 - HP LaserJet P2015") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\JL400 - HP LaserJet P2015"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\JL315 - HP LaserJet 4000") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\JL315 - HP LaserJet 4000"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\JL305 - HP LaserJet 1320") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\JL305 - HP LaserJet 1320"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\JL100 - HP LaserJet P2055dn") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\JL100 - HP LaserJet P2055dn"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\JL100 - HP LaserJet 4000") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\JL100 - HP LaserJet 4000"
	End If
	
	'2
	If LCase(strCurrentPrinter) = LCase("\\print\MP0023 - HP LaserJet P3005") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\MP0023 - HP LaserJet P3005"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\MP0036 - HP LaserJet 3600") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\MP0036 - HP LaserJet 3600"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\MP2029 - HP LaserJet P2015") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\MP2029 - HP LaserJet P2015"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\MP2045 - HP LaserJet P3005n") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\MP2045 - HP LaserJet P3005n"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\MP3021 - Xerox Phaser 6280dn") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\MP3021 - Xerox Phaser 6280dn"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\MP3033 - HP LaserJet P2055dn") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\MP3033 - HP LaserJet P2055dn"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\MP3033 - Xerox Phaser 8550") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\MP3033 - Xerox Phaser 8550"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\MP3035 - HP LaserJet P2105") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\MP3035 - HP LaserJet P2105"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\MP3035 - Xerox Phaser 6280dn") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\MP3035 - Xerox Phaser 6280dn"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\MP3046 - HP LaserJet 4000") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\MP3046 - HP LaserJet 4000"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\MP3051 - HP Color LaserJet 3800") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\MP3051 - HP Color LaserJet 3800"
	End If
	
	If LCase(strCurrentPrinter) = LCase("\\print\NW0105 - HP LaserJet P2035n") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\NW0105 - HP LaserJet P2035n"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\NW0106 - HP Color LaserJet 3700") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\NW0106 - HP Color LaserJet 3700"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\NW0106 - HP LaserJet 5M") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\NW0106 - HP LaserJet 5M"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\NW1102 - HP LaserJet P2015") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\NW1102 - HP LaserJet P2015"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\NW1102 - HP OfficeJet 7680 All-In-One") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\NW1102 - HP OfficeJet 7600 All-In-One"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\NW1104 - HP LaserJet P2015") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\NW1104 - HP LaserJet P2015"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\NW1110 - HP LaserJet CP3525x") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\NW1110 - HP LaserJet CP3525x"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\NW1118 - HP LaserJet P2015") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\NW1118 - HP LaserJet P2015"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\NW1118 - Xerox Phaser 8860") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\NW1118 - Xerox Phaser 8860"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\NW1130 -  HP LaserJet P2035n") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\NW1130 - HP LaserJet P2035n"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\NW2102 - HP LaserJet 1300") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\NW2102 - HP LaserJet 1300"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\NW2106 - HP Color LaserJet 2605") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\NW2106 - HP Color LaserJet 2605"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\NW2128 - HP LaserJet 4 Plus") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\NW2128 - HP LaserJet 4 Plus"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\NW2133 - HP LaserJet P2055dn") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\NW2133 - HP LaserJet P2055dn"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\NW3106 - HP LaserJet P2055dn") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\NW3106 - HP LaserJet P2055dn"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\NW3144 - HP LaserJet 2035n") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\NW3144 - HP LaserJet 2035n"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\NW4104 - HP LaserJet 1100") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\NW4104 - HP LaserJet 1100"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\NW4105 - Hp LaserJet P2035n") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\NW4105 - HP LaserJet P2035n"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\NW4107 - HP LaserJet P2035n") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\NW4107 - HP LaserJet P2035n"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\NW4112 - HP LaserJet P2055dn") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\NW4112 - HP LaserJet P2055dn"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\NW4120 - HP LaserJet 4000") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\printers\NW4120 - HP LaserJet 4000"
	End If
	
	
'	If LCase(strCurrentPrinter) = LCase("\\print\") Then
'		objNetwork.RemovePrinterConnection strCurrentPrinter
'		objNetwork.AddWindowsPrinterConnection "\\printers\"
'	End If
	
	'Install new copiers
	If LCase(strCurrentPrinter) = LCase("\\copiers\BI576 - Canon iR7200 Copier") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\copiers\BI576 - Canon iR4570 Copier"
	End If
	If LCase(strCurrentPrinter) = LCase("\\copiers\CE100 - Canon iRC6800 Copier") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\copiers\CE100 - Canon iR8095 Copier"
	End If
	If LCase(strCurrentPrinter) = LCase("\\copiers\NW1118 - Canon iR7095 Copier") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\copiers\NW1118 - Canon iR8095 Copier"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\CE100 - Canon iR C6800") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\copiers\CE100 - Canon iRC6800 Copier"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\ce100 - canon ir6800c copier") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\copiers\CE100 - Canon iRC6800 Copier"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\CE380 - Canon iR7095 Copier") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\copiers\CE380 - Canon iR7095 Copier"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\EL1058 - Canon iR7200 Copier") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\copiers\EL1058 - Canon iR7200 Copier"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\JL100 - Canon iR5570 Copier") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\copiers\JL100 - Canon iR5570 Copier"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\MP2060B - Canon iR8500 Copier") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\copiers\MP2060B - Canon iR8500 Copier"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\MP3033 - Canon iR7200 Copier") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\copiers\MP3033 - Canon iR7200 Copier"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\NW1118 - Canon iR7095 Copier") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\copiers\NW1118 - Canon iR7095 Copier"
	End If
	'NW2105 Lab Printers
	'msgbox strcurrentprinter
	If LCase(strCurrentPrinter) = LCase("\\print\NW2105 - HP Color LaserJet CP3505") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\metered-printers\NW2105 - HP Color LaserJet CP3505"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\NW2105 - HP LaserJet 4100") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\metered-printers\NW2105 - HP LaserJet 4100"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\NW2105 - HP LaserJet P4015dn") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\metered-printers\NW2105 - HP LaserJet P4015dn"
	End If
	'MP0008 Lab Printers
	If LCase(strCurrentPrinter) = LCase("\\print\MP0008 - HP Color LaserJet 3800") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\metered-printers\MP0008 - HP Color LaserJet 3800"
	End If
	If LCase(strCurrentPrinter) = LCase("\\print\MP0008 - HP LaserJet 4000") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\metered-printers\MP0008 - HP LaserJet 4000"
	End If
	'MP2047 Lab Printer
	If LCase(strCurrentPrinter) = LCase("\\print\MP2047 - HP LaserJet 4240") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\metered-printers\MP2047 - HP LaserJet 4240"
	End If
	'MP2060 Lab Printer
	If LCase(strCurrentPrinter) = LCase("\\print\MP2060 - HP LaserJet 4050N") Then
		objNetwork.RemovePrinterConnection strCurrentPrinter
		objNetwork.AddWindowsPrinterConnection "\\metered-printers\MP2060 - HP LaserJet 4240"
	End If
Next


'Map by Group
strUserPath = "LDAP://" & objSysInfo.UserName
Set objUser = GetObject(strUserPath)

For Each strGroup in objUser.MemberOf
	strGroupPath = "LDAP://" & strGroup
	Set objGroup = GetObject(strGroupPath)
	strGroupName = objGroup.CN

	Select Case strGroupName
		Case "Accounting"
			objNetwork.AddWindowsPrinterConnection "\\printers\NW1104 - HP LaserJet 3055"
			objNetwork.AddWindowsPrinterConnection "\\printers\NW1118 - Xerox Phaser 8860"
		Case "Allen Group"
			objNetwork.AddWindowsPrinterConnection "\\allen06\allenPrinter"
			objNetwork.AddWindowsPrinterConnection "\\allen02\0130hp2100"
			objNetwork.AddWindowsPrinterConnection "\\yaglaser\yaghp2100"
		Case "Babu Group"
			'objNetwork.AddWindowsPrinterConnection "\\printers\EL3073 - HP LaserJet 4000"
			objNetwork.AddWindowsPrinterConnection "\\metered-printers\NW2105 - HP Color LaserJet CP3505"
			objNetwork.AddWindowsPrinterConnection "\\metered-printers\NW2105 - HP LaserJet P4015dn"
		Case "Biochemistry Users"
			objNetwork.AddWindowsPrinterConnection "\\printers\BI776 - Xerox Workcenter 5655"
		Case "Computer Support"
			'objNetwork.AddWindowsPrinterConnection "\\Print\NW2105 - HP Color LaserJet CP3505"
			'objNetwork.AddWindowsPrinterConnection "\\Print\NW2105 - HP LaserJet 4100"
			'objNetwork.AddWindowsPrinterConnection "\\Print\NW2105 - HP LaserJet P4015dn"
			'objNetwork.AddWindowsPrinterConnection "\\Print\NW1118 - Xerox Phaser 8860"
		Case "Dalbey Group"
			objNetwork.AddWindowsPrinterConnection "\\printers\JL305 - HP LaserJet 1320"
			objNetwork.AddWindowsPrinterConnection "\\printers\JL100 - HP LaserJet 4000"
		Case "Dutta Group"
			objNetwork.AddWindowsPrinterConnection "\\printers\MP3051 - HP Color LaserJet 3800"
			objNetwork.AddWindowsPrinterConnection "\\legacyprint\MP3051 - HP LaserJet 2200"
		Case "Grad Studies"
			objNetwork.AddWindowsPrinterConnection "\\printers\EL1043 - HP LaserJet 4240"
			objNetwork.AddWindowsPrinterConnection "\\printers\EL1031 - HP LaserJet 4 Plus"
			objNetwork.AddWindowsPrinterConnection "\\printers\EL1033 - HP LaserJet 4050"
		Case "Harris Group"
			objNetwork.AddWindowsPrinterConnection "\\printers\MP2045 - HP LaserJet P3005n"
		Case "Jackman Group"
			objNetwork.AddWindowsPrinterConnection "\\printers\BI742 - Dell 1720dn"
		Case "Jaroniec Group"
			objNetwork.AddWindowsPrinterConnection "\\printers\EL1086 - HP LaserJet P3005"
		Case "Kohler Group"
			objNetwork.AddWindowsPrinterConnection "\\printers\NW0106 - HP Color LaserJet 3700"
			objNetwork.AddWindowsPrinterConnection "\\printers\NW0106 - HPLaserJet5M"
		Case "Mattson Group"
			objNetwork.AddWindowsPrinterConnection "\\printers\EL4086 - HP LaserJet P2035n"
		Case "Magliery Group"
			objNetwork.AddWindowsPrinterConnection "\\printers\EL3040 - HP LaserJet 2430"
			objNetwork.AddWindowsPrinterConnection "\\printers\EL3040 - Lexmark C532"
			objNetwork.AddWindowsPrinterConnection "\\printers\EL3040 - Lexmark E360dn"
			objNetwork.AddWindowsPrinterConnection "\\printers\EL3042 - Lexmark E360dn"
		Case "McCoy Group"
			objNetwork.AddWindowsPrinterConnection "\\metered-printers\NW2105 - HP Color LaserJet CP3505"
			'objNetwork.AddWindowsPrinterConnection "\\Print\NW2116 - HP LaserJet 2420"
		Case "Miller Group"
			objNetwork.AddWindowsPrinterConnection "\\printers\MP0036 - HP LaserJet 3600"
			objNetwork.AddWindowsPrinterConnection "\\printers\CE018 - HP Color LaserJet 4700"
			objNetwork.AddWindowsPrinterConnection "\\printers\CE018 - HP LaserJet P2055dn"
			'objNetwork.AddWindowsPrinterConnection "\\Print\Miller - Loaner Printer"
		Case "Musier Group"
			objNetwork.AddWindowsPrinterConnection "\\printers\MP3035 - HP Color LaserJet CP2600n"
		Case "Administrative Support"
			objNetwork.AddWindowsPrinterConnection "\\printers\NW1118 - Xerox Phaser 8860"
		Case "Pei Group"
			objNetwork.AddWindowsPrinterConnection "\\peiuv-vis\pei5550"
			objNetwork.AddWindowsPrinterConnection "\\printers\JL400 - HP Color LaserJet 3550"
			objNetwork.AddWindowsPrinterConnection "\\printers\JL400 - HP LaserJet P2015"
		Case "Personnel"
			objNetwork.AddWindowsPrinterConnection "\\printers\NW1102 - HP LaserJet P2015"
			objNetwork.AddWindowsPrinterConnection "\\printers\NW1102 - HP OfficeJet 7600 All-In-One"
			objNetwork.AddWindowsPrinterConnection "\\printers\NW1118 - HP LaserJet P2015"
		Case "Physical Chemistry Student Lecture Series"
			objNetwork.AddWindowsPrinterConnection "\\metered-printers\NW2105 - HP Color LaserJet CP3505"
			objNetwork.AddWindowsPrinterConnection "\\metered-printers\NW2105 - HP LaserJet P4015dn"
		Case "Paquette Group"
			objNetwork.AddWindowsPrinterConnection "\\paquette2\PaquetteHP"
		Case "Platz Group"
			objNetwork.AddWindowsPrinterConnection "\\print\EL1075 - Brother HL2070n"
		Case "Parquette Group"
			objNetwork.AddWindowsPrinterConnection "\\printers\NW4112 - HP LaserJet P2055dn"
			objNetwork.AddWindowsPrinterConnection "\\metered-printers\NW2105 - HP Color LaserJet CP3505"
			objNetwork.AddWindowsPrinterConnection "\\metered-printers\NW2105 - HP LaserJet P4015dn"
			'objNetwork.AddWindowsPrinterConnection "\\Print\NW4112 - HP LaserJet P2055dn"
		Case "Shore Group"
			objNetwork.AddWindowsPrinterConnection "\\shorepc2\HPLaser"
		Case "Stambuli Group"
			objNetwork.AddWindowsPrinterConnection "\\printers\EL4083 - HP LaserJet 1320"
		Case "Taylor Group"
			objNetwork.AddWindowsPrinterConnection "\\Taylorlab\TaylorPrinter"
		Case "Woodward Group"
			objNetwork.AddWindowsPrinterConnection "\\woodward14\HP LaserJet 2100"
			objNetwork.AddWindowsPrinterConnection "\\metered-printers\NW2105 - HP Color LaserJet CP3505"
			objNetwork.AddWindowsPrinterConnection "\\metered-printers\NW2105 - HP LaserJet P4015dn"
		Case "WOW Group"
			objNetwork.AddWindowsPrinterConnection "\\legacyprint\SL3132 - HP Color LaserJet 2600n"
		Case "Wu Group"
			objNetwork.AddWindowsPrinterConnection "\\Print\EL1045 - HP LaserJet 4250"
	End Select
Next


'Map by OU
strName = objSysInfo.ComputerName
arrComputerName = Split(strName, ",")
arrOU = Split(arrComputerName(1), "=")
strComputerOU = arrOU(1)

Select Case strComputerOU
	Case "Accounting"
		objNetwork.AddWindowsPrinterConnection "\\Print\NW1104 - HP LaserJet 3055"
		objNetwork.SetDefaultPrinter "\\Print\NW1104 - HP LaserJet 3055"
'	Case "Admin Associates"
'		objNetwork.AddWindowsPrinterConnection "\\Print\NW1118 - HP LaserJet 2300"
'		objNetwork.AddWindowsPrinterConnection "\\Print\NW1118 - HP Color LaserJet CP3505"
'		objNetwork.SetDefaultPrinter "\\Print\NW1118 - HP LaserJet 2300"
	Case "CE0140"
		objNetwork.AddWindowsPrinterConnection "\\Print\CE140 - HP LaserJet P2055dn"
		objNetwork.SetDefaultPrinter "\\Print\CE140 - HP LaserJet P2055dn"
	Case "CE0160"
		objNetwork.AddWindowsPrinterConnection "\\uniprint1\chemspool1"
		objNetwork.SetDefaultPrinter "\\uniprint1\chemspool1"
	Case "CE0170"
		objNetwork.AddWindowsPrinterConnection "\\uniprint1\chemspool1"
		objNetwork.SetDefaultPrinter "\\uniprint1\chemspool1"
	Case "CE0400"
		objNetwork.AddWindowsPrinterConnection "\\printers\CE431 - HP LaserJet 2430"
		objNetwork.AddWindowsPrinterConnection "\\copiers\CE380 - Canon iR7095 Copier"
		objNetwork.SetDefaultPrinter "\\printers\CE431 - HP LaserJet 2430"
	Case "MP2045"
		objNetwork.AddWindowsPrinterConnection "\\printers\MP2045 - HP LaserJet P3005n"
		objNetwork.SetDefaultPrinter "\\printers\MP2045 - HP LaserJet P3005n"
	Case "MP2047"
		'objNetwork.AddWindowsPrinterConnection "\\metered-printers\MP2047 - HP LaserJet 4240"
		objNetwork.AddWindowsPrinterConnection "\\uniprint1\mp2047-hp4240spool"
		objNetwork.SetDefaultPrinter "\\uniprint1\mp2047-hp4240spool"
	Case "MP2060"
		objNetwork.AddWindowsPrinterConnection "\\printers\MP2060 - HP LaserJet 4240"
		objNetwork.SetDefaultPrinter "\\printers\MP2060 - HP LaserJet 4240"
	Case "MP0008"
		'objNetwork.AddWindowsPrinterConnection "\\metered-printers\MP0008 - HP Color LaserJet 3800"
		'objNetwork.AddWindowsPrinterConnection "\\metered-printers\MP0008 - HP LaserJet 4000"
		objNetwork.AddWindowsPrinterConnection "\\uniprint1\mp0008-hp3800"
		objNetwork.AddWindowsPrinterConnection "\\uniprint1\mp0008-hp4000spool"
		objNetwork.SetDefaultPrinter "\\uniprint1\mp0008-hp4000spool"
	Case "NW2105"
		objNetwork.AddWindowsPrinterConnection "\\metered-printers\NW2105 - HP Color LaserJet CP3505"
		objNetwork.AddWindowsPrinterConnection "\\metered-printers\NW2105 - HP LaserJet 4100"
		objNetwork.AddWindowsPrinterConnection "\\metered-printers\NW2105 - HP LaserJet P4015dn"
		objNetwork.SetDefaultPrinter "\\metered-printers\NW2105 - HP LaserJet P4015dn"
	Case "Surface Analysis"
		objNetwork.AddWindowsPrinterConnection "\\printers\EL0056 - HP LaserJet CP1518ni"
		objNetwork.AddWindowsPrinterConnection "\\printers\EL0056 - HP LaserJet P3005"
		objNetwork.AddWindowsPrinterConnection "\\metered-printers\NW2105 - HP Color LaserJet CP3505"
		objNetwork.SetDefaultPrinter "\\printers\EL0056 - HP LaserJet P3005"
	'Case "Undergraduate Office"
	'	'objNetwork.AddWindowsPrinterConnection "\\printers\CE100 - HP LaserJet 2300"
	Case "MP2045"
		objNetwork.AddWindowsPrinterConnection "\\printers\MP2045 - HP LaserJet P3005n"
	Case "Kiosks"
		objNetwork.AddWindowsPrinterConnection " \\uniprint1\chemspool1"
End Select

'Map by User
'Disabled - The group mapping part of this script no longer sets default printer, so this should be taken care of.
'strCurrentUser = objNetwork.UserName
'Select Case strCurrentUser
	'Case "hadad"
	'	objNetwork.SetDefaultPrinter "\\Print\EL1033 - HP LaserJet 4050"
	'Case "bforan"
	'	objNetwork.AddWindowsPrinterConnection "\\Print-Server\nw2146-hplj4650PS"
	'	objNetwork.SetDefaultPrinter "\\Print-Server\nw2146-hplj4650PS"
	'Case "pei"
	'	objNetwork.AddWindowsPrinterConnection "\\Print\JL100 - HP LaserJet 4000"
	'	objNetwork.SetDefaultPrinter "\\Print\JL100 - HP LaserJet 4000"
	'Case "mbrett"
	'	objNetwork.SetDefaultPrinter "HP LaserJet 1020"
'End Select

'Map by computer name
strDN = objSysInfo.ComputerName
arrDN = Split(strDN, ",")
arrComputername = Split(arrDN(0), "=")
strComputername = arrComputername(1)
'msgbox strComputerName
If InStr(strComputerName,"-IR100") Then
	objNetwork.AddWindowsPrinterConnection "\\printers\CE431 - Brother MFC-8510DN"
	objNetwork.SetDefaultPrinter "\\printers\CE431 - Brother MFC-8510DN"
End If