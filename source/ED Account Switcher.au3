#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=arrows.ico
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Fileversion=1.0.2
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <array.au3>
#include <file.au3>
#include <GUIConstantsEx.au3>
#include <ColorConstants.au3>
#include <GuiComboBox.au3>

Global $title = "E:D Account Switcher 1.0.2"
Global $basename=stringleft(@ScriptName , StringInStr(@ScriptName , "." , 0, -1)-1)
Global $sIniname=@scriptdir&'\'&$basename&".ini"

Global $launcherDIR = "startup"

#cs
if not FileExists('EDlaunch.exe') Then
	MsgBox(0,"Error","Put the "&@ScriptName&" in the same directory as the ED Launcher !",0,0)
	Exit
endif
#ce

; It's possible to minimize user envolvement even further by grabbing the path from the currently running launcher before closing it. This is done in the last function, down at the bottom.


$Form1 = GUICreate($title, 370, 180)

$Group1 = GUICtrlCreateGroup("Switch Account", 5, 5, 360, 50)
Global $Combo1 = GUICtrlCreateCombo("", 15, 25, 200, 20, BitOR($GUI_SS_DEFAULT_COMBO, $CBS_DROPDOWNLIST))
_GUICtrlComboBox_SetCueBanner($Combo1, "Choose Account")

$apply = GUICtrlCreateButton("Launch / Apply", 255, 24, 90, 23)
GUICtrlCreateGroup("", -99, -99, 1, 1)

$Group2 = GUICtrlCreateGroup("Add Currently Active Account to DB", 5, 70, 360, 45)
$makeNewAcc = GUICtrlCreateButton("Add New Account", 120, 85, 130, 23)
GUICtrlCreateGroup("", -99, -99, 1, 1)

$Group3 = GUICtrlCreateGroup("EDLaunch Status", 5, 130, 360, 45)
global $infotxt = GUICtrlCreateLabel("", 15, 149, 340, 16)
GUICtrlCreateGroup("", -99, -99, 1, 1)

GUISetState(@SW_SHOW)

Global $aAccountsDataDB[1][8]
_CurrentToDB(1)

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $apply
			$combo = GUICtrlRead($Combo1)
			_SwitchToAccount($combo)
		Case $makeNewAcc
			_CurrentToDB()
	EndSwitch

	_checkEDLaunch()
WEnd

Func _SwitchToAccount($account)
	local $dif = 0, $return
	If $account = "" Then Return -1

    If $launcherDIR <> "" Then ; Check if the EDlaunch process is running.
		GUICtrlSetColor($infotxt, $COLOR_WHITE)
		GUICtrlSetBkColor($infotxt, $COLOR_RED)
		GUICtrlSetData($infotxt, " Stopping current EDlaunch...")
		Sleep(1000)

		; keep trying to close the launcher for 10 seconds in case it is frozen
		local $begin = timerinit()
		Do
			ProcessClose("edlaunch.exe")
			sleep(500)
			$dif = timerdiff($begin)
			if $dif >= 10000 Then
				$return = MsgBox(49, $title, "Couldn't close the launcher, please do so manually.")
				if $return = 2 Then
					MsgBox(16, $title, "Cannot continue. Exiting.")
					Exit
				EndIf
			EndIf
		until not processexists("edlaunch.exe")

		Sleep(1000)
    EndIf

	Local $hFile = FileOpen($aAccountsDataDB[0][0], 0)
	Local $content = FileRead($hFile)
	FileClose($hFile)

	For $i = 1 To UBound($aAccountsDataDB) - 1
		If $aAccountsDataDB[$i][0] == $account Then ExitLoop
	Next

	; the values for these settings should always be present in the launcher's config file
	$content = StringRegExpReplace($content, '(?Ui)(.*"userSettings".*PublicKeyToken=)(.*)(".*)', "${1}" & $aAccountsDataDB[$i][1] & "${3}")
	$content = StringRegExpReplace($content, '(?Ui)(.*"FORCServerSupport.Properties.Settings".*PublicKeyToken=)(.*)(".*)', "${1}" & $aAccountsDataDB[$i][2] & "${3}")
	$content = StringRegExpReplace($content, '(?Ui)(.*"ClientSupport.Properties.Settings".*PublicKeyToken=)(.*)(".*)', "${1}" & $aAccountsDataDB[$i][3] & "${3}")
	$content = StringRegExpReplace($content, '(?Ui)(.*"CBViewModel.Properties.Settings".*PublicKeyToken=)(.*)(".*)', "${1}" & $aAccountsDataDB[$i][4] & "${3}")
	$content = StringRegExpReplace($content, '(?Ui)(.*"MachineToken".*\s*?<value>)(.*)(</value>)', "${1}" & $aAccountsDataDB[$i][5] & "${3}")


	; however users may or may not have chosen to save these values:

	; username
	$valuestyleUN = StringRegExp($content, '(?Uism)<setting name="UserName".*>\s.*(?:<value>(.*)</value>|<value />)', 1)
	If $aAccountsDataDB[$i][6] <> "none" Then
		If Not IsArray($valuestyleUN) Then
			; malformed launcher config
		EndIf
		If Not StringInStr($valuestyleUN[0], "<value />") Then
			$content = StringRegExpReplace($content, '(?Ui)(.*"UserName".*\s*?<value>)(.*)(</value>)', "${1}" & $aAccountsDataDB[$i][6] & "${3}")
		Else
			$content = StringRegExpReplace($content, '(?Ui)(.*"UserName".*\s*?)(<value />)', "${1}" & '<value>' & $aAccountsDataDB[$i][6] & '</value>')
		EndIf
	Else
		If Not StringInStr($valuestyleUN[0], "<value />") Then
			$content = StringRegExpReplace($content, '(?Ui)(.*"UserName".*\s*?)(<value>.*</value>)', "${1}" & '<value />')
		EndIf
	EndIf

	; password
	$valuestylePW = StringRegExp($content, '(?Uism)<setting name="Password".*>\s.*(?:<value>(.*)</value>|<value />)', 1)
	If $aAccountsDataDB[$i][7] <> "none" Then
		If Not IsArray($valuestylePW) Then
			; malformed launcher config
		EndIf
		If Not StringInStr($valuestylePW[0], "<value />") Then
			$content = StringRegExpReplace($content, '(?Ui)(.*"Password".*\s*?<value>)(.*)(</value>)', "${1}" & $aAccountsDataDB[$i][7] & "${3}")
		Else
			$content = StringRegExpReplace($content, '(?Ui)(.*"Password".*\s*?)(<value />)', "${1}" & '<value>' & $aAccountsDataDB[$i][7] & '</value>')
		EndIf
	Else
		If Not StringInStr($valuestylePW[0], "<value />") Then
			$content = StringRegExpReplace($content, '(?Ui)(.*"Password".*\s*?)(<value>.*</value>)', "${1}" & '<value />')
		EndIf
	EndIf

	Local $hFile = FileOpen($aAccountsDataDB[0][0], 2)
	FileWrite($hFile, $content)
	FileClose($hFile)

    ; Start the ED launcher program.		If it was found to be running previously.
	if fileexists($launcherDIR) Then
		GUICtrlSetState($apply, $GUI_DISABLE)
		GUICtrlSetColor($infotxt, $COLOR_WHITE)
		GUICtrlSetBkColor($infotxt, $COLOR_BLUE)
		GUICtrlSetData ($infotxt, " Starting EDlauncher... takes a moment... ")
		Run($launcherDIR)

		; WinWaitActive() alone won't work on some systems as the launcher window might pop up inactive behind the E:D Account Switcher window
		; "Dangerous" might eventually change to "Horizons" or something else entirely, so I think it's better to only search the title for "Elite" and "Launcher" but make sure the window class also matches.

		$dif = 0 ; was declared local above, just need to zero it just to make sure
		$return = "" ; same here
		$begin = timerinit(); also was declared above and only needs to be reinitialized
		Do
			sleep(100)
			$dif = timerdiff($begin)
			if $dif >= 10000 Then
				$return = MsgBox(49, $title, "EDLauncher didn't start, please start it manually.")
				if $return = 2 Then
					MsgBox(16, $title, "Cannot continue. Exiting.")
					Exit
				EndIf
			EndIf
		until winexists("[REGEXPTITLE:(?si)Elite(.*?)Launcher; REGEXPCLASS:(?si)HwndWrapper(.*?)]", "") ; wait for the window to come up but don't care about whether or not it's active

		$hWnd = wingethandle("[REGEXPTITLE:(?si)Elite(.*?)Launcher; REGEXPCLASS:(?si)HwndWrapper(.*?)]") ; grab the handle of the window...

		$dif = 0
		$begin = timerinit()
		While 1 ; ... and wait for it to become responsive
			$aRet = DLLCall('user32.dll', "bool", "IsHungAppWindow", "hwnd", $hWnd)
			If @error Or $aRet[0] = 0 Then ExitLoop
			sleep(100)
			$dif = timerdiff($begin)
			if $dif >= 10000 Then
				MsgBox(16, $title, "EDLauncher started but didn't become responsive. Cannot continue. Exiting.")
				Exit
			EndIf
		WEnd

		GUICtrlSetBkColor($infotxt, $GUI_BKCOLOR_TRANSPARENT)
		GUICtrlSetColor($infotxt, $COLOR_BLACK)
		GUICtrlSetData ($infotxt, "")
		GUICtrlSetState($apply, $GUI_ENABLE)
	EndIf

	$launcherDIR = "startup" ; the lazy way out...
	_checkEDLaunch()
EndFunc   ;==>_SwitchToAccount

Func _SetCombo()
	;_arraydisplay($aAccountsDataDB, "_setCombo() start")
	If UBound($aAccountsDataDB) >= 2 Then
		For $i = 1 To UBound($aAccountsDataDB) - 1
			If $aAccountsDataDB[0][1] == $aAccountsDataDB[$i][1] And _
					$aAccountsDataDB[0][2] == $aAccountsDataDB[$i][2] And _
					$aAccountsDataDB[0][3] == $aAccountsDataDB[$i][3] And _
					$aAccountsDataDB[0][4] == $aAccountsDataDB[$i][4] And _
					$aAccountsDataDB[0][5] == $aAccountsDataDB[$i][5] And _
					$aAccountsDataDB[0][6] == $aAccountsDataDB[$i][6] And _
					$aAccountsDataDB[0][7] == $aAccountsDataDB[$i][7] Then
				_GUICtrlComboBox_SetCurSel($Combo1, $i - 1)
				Return 1
			EndIf
		Next
	EndIf
EndFunc   ;==>_SetCombo

Func _AccountsDataFromDB()
	; assume new account DB if there are no saved accounts or EDswitch.ini was cleaned
	If Not FileExists($sIniname) Then Return 0
	Local $aSectionnames = IniReadSectionNames($sIniname)
	If Not IsArray($aSectionnames) Then Return 0

	_ArraySort($aSectionnames, 0, 1)

	For $i = 1 To $aSectionnames[0]

		$aSection = IniReadSection($sIniname, $aSectionnames[$i])
		If Not IsArray($aSection) Then ContinueLoop

		; make sure DB entry keys exist before adding them to the array
		$idx1 = _ArraySearch($aSection, "userSettings", 1)
		$idx2 = _ArraySearch($aSection, "FORCServerSupport.Properties.Settings", 1)
		$idx3 = _ArraySearch($aSection, "ClientSupport.Properties.Settings", 1)
		$idx4 = _ArraySearch($aSection, "CBViewModel.Properties.Settings", 1)
		$idx5 = _ArraySearch($aSection, "MachineToken", 1)
		$idx6 = _ArraySearch($aSection, "UserName", 1)
		$idx7 = _ArraySearch($aSection, "Password", 1)
		If $idx1 = -1 Or $idx2 = -1 Or $idx3 = -1 Or $idx4 = -1 Or $idx5 = -1 Or $idx6 = -1 Or $idx7 = -1 Then ContinueLoop

		; we also need to make sure the value for every key at least contains something
		If StringStripWS($aSection[$idx1][1], 8) = "" _
				Or StringStripWS($aSection[$idx2][1], 8) = "" _
				Or StringStripWS($aSection[$idx3][1], 8) = "" _
				Or StringStripWS($aSection[$idx4][1], 8) = "" _
				Or StringStripWS($aSection[$idx5][1], 8) = "" _
				Or StringStripWS($aSection[$idx6][1], 8) = "" _
				Or StringStripWS($aSection[$idx7][1], 8) = "" Then ContinueLoop

		;anything that's left at this point is for the launcher to check for validity. It's good enough for us though, so we add it to our DB array
		ReDim $aAccountsDataDB[UBound($aAccountsDataDB) + 1][8]
		$aAccountsDataDB[UBound($aAccountsDataDB) - 1][0] = $aSectionnames[$i]
		$aAccountsDataDB[UBound($aAccountsDataDB) - 1][1] = $aSection[$idx1][1]
		$aAccountsDataDB[UBound($aAccountsDataDB) - 1][2] = $aSection[$idx2][1]
		$aAccountsDataDB[UBound($aAccountsDataDB) - 1][3] = $aSection[$idx3][1]
		$aAccountsDataDB[UBound($aAccountsDataDB) - 1][4] = $aSection[$idx4][1]
		$aAccountsDataDB[UBound($aAccountsDataDB) - 1][5] = $aSection[$idx5][1]
		$aAccountsDataDB[UBound($aAccountsDataDB) - 1][6] = $aSection[$idx6][1]
		$aAccountsDataDB[UBound($aAccountsDataDB) - 1][7] = $aSection[$idx7][1]
	Next
	;_arraydisplay($aAccountsDataDB)
EndFunc   ;==>_AccountsDataFromDB

Func _AccountDataFromLauncher()
	Local $sFile = _getMostCurrentUserConfigPath()
	If $sFile = -1 Then
		MsgBox(16, $title, "Launcher settings not found. Make sure E:D is installed and you have logged in at least once.")
		Exit
	EndIf

	Local $hFile = FileOpen($sFile, 0)
	Local $content = FileRead($hFile)
	FileClose($hFile)

	$_1 = _getPublicKeyToken($content, "userSettings")
	$_2 = _getPublicKeyToken($content, "FORCServerSupport.Properties.Settings")
	$_3 = _getPublicKeyToken($content, "ClientSupport.Properties.Settings")
	$_4 = _getPublicKeyToken($content, "CBViewModel.Properties.Settings")
	$_5 = _getSetting($content, "MachineToken")
	$_6 = _getSetting($content, "UserName")
	If $_6 = -1 Then $_6 = "none"
	$_7 = _getSetting($content, "Password")
	If $_7 = -1 Then $_7 = "none"
	;msgbox(0, "test", $_6 & @crlf & $_7)

	If $_1 = -1 Or $_2 = -1 Or $_3 = -1 Or $_4 = -1 Or $_5 = -1 Then
		MsgBox(16, $title, "Unexpected launcher settings. Either the configuration file has been updated to a new layout or it was corrupted." & @CRLF & @CRLF & "Cannot continue." & @CRLF & @CRLF & "Exiting.")
		Exit
	EndIf

	$aAccountsDataDB[UBound($aAccountsDataDB) - 1][0] = $sFile
	$aAccountsDataDB[UBound($aAccountsDataDB) - 1][1] = $_1
	$aAccountsDataDB[UBound($aAccountsDataDB) - 1][2] = $_2
	$aAccountsDataDB[UBound($aAccountsDataDB) - 1][3] = $_3
	$aAccountsDataDB[UBound($aAccountsDataDB) - 1][4] = $_4
	$aAccountsDataDB[UBound($aAccountsDataDB) - 1][5] = $_5
	$aAccountsDataDB[UBound($aAccountsDataDB) - 1][6] = $_6
	$aAccountsDataDB[UBound($aAccountsDataDB) - 1][7] = $_7
	;_arraydisplay($aAccountsDataDB)
EndFunc   ;==>_AccountDataFromLauncher

Func _CurrentToDB($startup = 0)
	; make sure we got up to date data from all data sources
	Dim $aAccountsDataDB[1][8]
	_AccountDataFromLauncher()
	_AccountsDataFromDB()
	_DBtoCombo()
	_SetCombo()

	If UBound($aAccountsDataDB) >= 2 Then
		For $i = 1 To UBound($aAccountsDataDB) - 1
			If $aAccountsDataDB[0][1] = $aAccountsDataDB[$i][1] And _
					$aAccountsDataDB[0][2] = $aAccountsDataDB[$i][2] And _
					$aAccountsDataDB[0][3] = $aAccountsDataDB[$i][3] And _
					$aAccountsDataDB[0][4] = $aAccountsDataDB[$i][4] And _
					$aAccountsDataDB[0][5] = $aAccountsDataDB[$i][5] And _
					$aAccountsDataDB[0][6] = $aAccountsDataDB[$i][6] And _
					$aAccountsDataDB[0][7] = $aAccountsDataDB[$i][7] Then
				If $startup = 0 Then MsgBox(64, $title, "The launcher's current account settings are already saved in the accounts DB as:" & @CRLF & @CRLF & $aAccountsDataDB[$i][0])
				Return 0
			EndIf
		Next
	EndIf

	If $startup = 1 And MsgBox(36, $title, "Current E:D account not found in account switcher DB. Shall I add it?") = 7 Then Return 0

	$sSectionName = InputBox($title, "Enter a name for your new account", "Account " & UBound($aAccountsDataDB))
	If $sSectionName = "" Then
		MsgBox(48, $title, "Account name must not be empty.")
		Return 0
	EndIf
	;_arraydisplay($aAccountsDataDB)
	IniWrite($sIniname, $sSectionName, "userSettings", $aAccountsDataDB[0][1])
	IniWrite($sIniname, $sSectionName, "FORCServerSupport.Properties.Settings", $aAccountsDataDB[0][2])
	IniWrite($sIniname, $sSectionName, "ClientSupport.Properties.Settings", $aAccountsDataDB[0][3])
	IniWrite($sIniname, $sSectionName, "CBViewModel.Properties.Settings", $aAccountsDataDB[0][4])
	IniWrite($sIniname, $sSectionName, "MachineToken", $aAccountsDataDB[0][5])
	IniWrite($sIniname, $sSectionName, "UserName", $aAccountsDataDB[0][6])
	IniWrite($sIniname, $sSectionName, "Password", $aAccountsDataDB[0][7])

	; rebuild the DB array with all data available and make the appropriate changes to the GUI
	Dim $aAccountsDataDB[1][8]
	_AccountDataFromLauncher()
	_AccountsDataFromDB()
	_DBtoCombo()
	_SetCombo()
	;_arraydisplay($aAccountsDataDB)
EndFunc   ;==>_CurrentToDB

Func _DBtoCombo()
	Local $sComboContent = ""
	For $i = 1 To UBound($aAccountsDataDB) - 1
		$sComboContent &= $aAccountsDataDB[$i][0] & "|"
	Next
	_GUICtrlComboBox_ResetContent($Combo1)
	GUICtrlSetData($Combo1, $sComboContent)
EndFunc   ;==>_DBtoCombo

Func _getPublicKeyToken($content, $section)
	Local $regex = StringRegExp($content, '(?Uis:' & $section & '.*publickeytoken=(.*)")', 1)
	If IsArray($regex) Then
		Return $regex[0]
	EndIf
	Return -1
EndFunc   ;==>_getPublicKeyToken

Func _getSetting($content, $setting)
	Local $regex = StringRegExp($content, '(?i)<setting name="' & $setting & '".*\s{1,}?.*<value>(.*)</value>', 1)
	If IsArray($regex) Then
		Return $regex[0]
	EndIf
	Return -1
EndFunc   ;==>_getSetting

Func _getMostCurrentUserConfigPath()
	Local $aFiles = _FileListToArrayRec("C:\Users\" & @UserName & "\AppData\Local\Frontier_Developments\", "user.config", 1, 1, 0, 2)
	If $aFiles = "" Then Return -1
	Local $aTest[0], $fullpath, $versionpath
	For $i = 1 To UBound($aFiles) - 1
		$fullpath = $aFiles[$i]
		$versionpath = StringRegExp($fullpath, "(?i:.*\\(.*\\user.config))", 1)
		If IsArray($versionpath) Then
			ReDim $aTest[UBound($aTest) + 1]
			$aTest[UBound($aTest) - 1] = $versionpath[0]
		EndIf
	Next
	_ArraySort($aTest, 1, 0)
	Local $idx = _ArraySearch($aFiles, $aTest[0], 0, 0, 1, 1)
	;_arraydisplay($aFiles, $idx)
	Return $aFiles[$idx]
EndFunc   ;==>_getMostCurrentUserConfigPath

Func _ProcessGetLocation($iPID)
	Local $aProc = DllCall('kernel32.dll', 'hwnd', 'OpenProcess', 'int', BitOR(0x0400, 0x0010), 'int', 0, 'int', $iPID)
	If $aProc[0] = 0 Then Return SetError(1, 0, '')
	Local $vStruct = DllStructCreate('int[1024]')
	DllCall('psapi.dll', 'int', 'EnumProcessModules', 'hwnd', $aProc[0], 'ptr', DllStructGetPtr($vStruct), 'int', DllStructGetSize($vStruct), 'int_ptr', 0)
	Local $aReturn = DllCall('psapi.dll', 'int', 'GetModuleFileNameEx', 'hwnd', $aProc[0], 'int', DllStructGetData($vStruct, 1), 'str', '', 'int', 2048)
	If StringLen($aReturn[3]) = 0 Then Return SetError(2, 0, '')
	Return $aReturn[3]
EndFunc   ;==>_ProcessGetLocation

Func _checkEDLaunch()
	If ProcessExists("edlaunch.exe") and (($launcherDIR = "") or ($launcherDIR = "startup")) Then
		$list = ProcessList("edlaunch.exe")
		$launcherDIR = _ProcessGetLocation($list[1][1]) ; $launcherDIR then contains the full path to the launcher, including "edlaunch.exe"
		GUICtrlSetColor($infotxt, $COLOR_WHITE)
		GUICtrlSetBkColor($infotxt, $COLOR_BLUE)
		GUICtrlSetData($infotxt, " EDlaunch found, running in launch mode")
	EndIf

	If not ProcessExists("edlaunch.exe") and ($launcherDIR <> "") Then
		$launcherDIR = ""
		GUICtrlSetColor($infotxt, $COLOR_WHITE)
		GUICtrlSetBkColor($infotxt, $COLOR_BLUE)
		GUICtrlSetData($infotxt, " EDlaunch NOT found, running in apply mode")
	EndIf
EndFunc
