#include-once

; UDF Version 0.0.0.2 - Use at your own risk, subject to script breaking changes


Global $bDebugging = False

Global Const $__aKPErrorDesc = [ _
	"Success", _
	"Run error, check paths", _
	"Invalid password combination, verify unlock key", _
	"Unknown error" _
]

Global Enum $__KPError_Success = 0, _
			$__KPError_BadCommand, _
			$__KPError_CompositeKey, _
			$__KPError_Unknown

Global Enum $__KPUnlock_Password, _
			$__KPUnlock_PasswordEnc, _
			$__KPUnlock_KeyFile, _
			$__KPUnlock_UserAccount

#include <AutoItConstants.au3> ; STDOUT_CHILD STDIN_CHILD
#include <StringConstants.au3> ; StripLeading + StripTrailing

#Region ### Notes ###
#CS
This UDF allows you to edit a KeePass database.
N.B.: If the database is open and unsaved in KeePass when KeePass attempts to save, it will ask the user how to resolve: Overwrite, reload, or cancel.

Derived directly from the documentation here: https://keepass.info/help/v2_dev/scr_sc_index.html

{PASSWORD_ENC}
	I don't understand what it is used for or how it is used.
	If you need this implemented, please let me know and possibly help me figure out what it is used for and how.

REQUIRED: Download KPScript and extract to KeePass' installation directory. Download link --> https://keepass.info/plugins.html#kpscript
"The KPScript.exe file needs to be copied into the directory where KeePass is installed (where the KeePass.exe file is)."
My (full) copy installed itself to "C:\Program Files (x86)\KeePass Password Safe 2\"

List of Commands provided by KPScript
	(N.B. return values aren't parsed well, that's step 2. First, get it working, then make it pretty)

Created
	ListGroups
	ListEntries
	GetEntryString
	GenPw
	EstimateQuality
	DeleteEntry
	AddEntry
	EditEntry

Unimplemented
	MoveEntry
	Import
	Export
	Sync
	ChangeMasterKey
	DetachBins

#CE
#EndRegion ### Notes ###

; #FUNCTION# ====================================================================================================================
; Name ..........: _KPScript_EntryAdd
; Description ...: Create a new entry in a KeePass database
; Syntax ........: _KPScript_EntryAdd($sKPScript, $sDatabaseFile, $aUnlockKey, $sEntryProperties)
; Parameters ....: $sKPScript           - full path to KPScript.exe.
;                  $sDatabaseFile       - full path to the kdbx file.
;                  $aUnlockKey          - an Unlock array from _KPScript_UnlockKey.
;                  $aEntryProperties    - an array returned from _KPScript_EntryProperties.
; Return values .: Success - True
;                  Failure - False and sets @error to 1
; Author ........: Seadoggie01
; Modified ......:
; Remarks .......: This method WILL create a group if you pass a new group name or path. If no path is passed, it will be
;                  created in the base directory.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _KPScript_EntryAdd($sKPScript, $sDatabaseFile, $aUnlockKey, $aEntryProperties)

	; Because some things don't make sense, this command uses a different syntax. Ugh!
	; Replace all -set- in the command with -
	Local $sEntryProperties = ""
	For $i=0 To UBound($aEntryProperties) - 1
		$sEntryProperties &= StringReplace($aEntryProperties[$i][0], "-set-", "-") & $aEntryProperties[$i][1]
	Next

	__KPScript_Run($sKPScript, "AddEntry", __KPScript_StringWrap($sDatabaseFile) & " " & $sEntryProperties, $aUnlockKey)
	If @error Then Return SetError(1, @error, False)

	Return True

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _KPScript_EntryProperties
; Description ...: Builds a list of properties to be set when adding/editing an entry
; Syntax ........: _KPScript_EntryProperties([$vTitle = Default[, $sUserName = Default[, $sPassword = Default[, $sURL = Default[,
;                  $sNotes = Default[, $sGroupName = Default[, $sGroupPath = Default[, $iIcon = Default[, $iCustIcon = Default[,
;                  $bExpires = Default[, $sExpiryTime = Default]]]]]]]]]]])
; Parameters ....: $vTitle              - [optional] a variant value. Default is Default.
;                  $sUserName           - [optional] a string value. Default is Default.
;                  $sPassword           - [optional] a string value. Default is Default.
;                  $sURL                - [optional] a string value. Default is Default.
;                  $sNotes              - [optional] a string value. Default is Default.
;                  $sGroupName          - [optional] a string value. Default is Default.
;                  $sGroupPath          - [optional] a string value. Default is Default.
;                  $iIcon               - [optional] an integer value. Default is Default.
;                  $iCustIcon           - [optional] an integer value. Default is Default.
;                  $bExpires            - [optional] a boolean value. Default is Default.
;                  $sExpiryTime         - [optional] a string value. Default is Default.
; Return values .: An 0-based, 2D array of properties to be set
; Author ........: Seadoggie01
; Modified ......: August 18, 2021
; Remarks .......: Empty strings will be used to set a value to empty. Use Default to ignore a column
; Related .......: _EntryAdd, _EntryEdit
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _KPScript_EntryProperties($vTitle = Default, $sUserName = Default, $sPassword = Default, $sURL = Default, $sNotes = Default, $sGroupName = Default, $sGroupPath = Default, $iIcon = Default, $iCustIcon = Default, $bExpires = Default, $sExpiryTime = Default)

	Local $aProperties = [["Title",Default],["UserName",Default],["Password",Default],["URL",Default],["Notes",Default],["GroupName",Default],["GroupPath",Default],["Icon",Default],["CustomIcon",Default],["Expires",Default],["ExpiryTime",Default]]
	If IsArray($vTitle) Then
		$aProperties = $vTitle
	Else
		If Not IsKeyword($vTitle) Then $aProperties[0][1] = $vTitle
		If Not IsKeyword($sUserName) Then $aProperties[1][1] = $sUserName
		If Not IsKeyword($sPassword) Then $aProperties[2][1] = $sPassword
		If Not IsKeyword($sURL) Then $aProperties[3][1] = $sURL
		If Not IsKeyword($sNotes) Then $aProperties[4][1] = $sNotes
		If Not IsKeyword($sGroupName) Then $aProperties[5][1] = $sGroupName
		If Not IsKeyword($sGroupPath) Then $aProperties[6][1] = $sGroupPath
		If Not IsKeyword($iIcon) Then $aProperties[7][1] = $iIcon
		If Not IsKeyword($iCustIcon) Then $aProperties[8][1] = $iCustIcon
		If Not IsKeyword($bExpires) Then $aProperties[9][1] = $bExpires
		If Not IsKeyword($sExpiryTime) Then $aProperties[10][1] = $sExpiryTime
	EndIf

	Local $aRetProperties[UBound($aProperties)][2], $iIndex = 0

	; For each property
	For $i=0 To UBound($aProperties) - 1
		; If the value isn't a keyword
		If Not IsKeyword($aProperties[$i][1]) Then
			; If it's a special column
			If StringInStr("Icon|CustomIcon|Expires|ExpiryTime|", $aProperties[$i][0] & "|") Then
				$aRetProperties[$iIndex][0] = " -setx-" & $aProperties[$i][0]
			Else
				$aRetProperties[$iIndex][0] = " -set-" & $aProperties[$i][0]
			EndIf
			$aRetProperties[$iIndex][1] = __KPScript_StringWrap($aProperties[$i][1])
			$iIndex += 1
		EndIf
	Next
	ReDim $aRetProperties[$iIndex][2]
	Return $aRetProperties

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _KPScript_EntryDelete
; Description ...: Deletes one or more existing entries from a database.
; Syntax ........: _KPScript_EntryDelete($sKPScript, $sDatabaseFile, $aUnlockKey, $sIdentification)
; Parameters ....: $sKPScript           - full path to KPScript.exe.
;                  $sDatabaseFile       - full path to the kdbx file.
;                  $aUnlockKey          - an Unlock array from _KPScript_UnlockKey.
;                  $sIdentification     - a string value from _KPScript_EntryIdentify.
; Return values .: Success - True
;                  Failure - False and sets @error to 1
; Author ........: Seadoggie01
; Modified ......: July 7, 2020
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _KPScript_EntryDelete($sKPScript, $sDatabaseFile, $aUnlockKey, $sIdentification)

	Local $vRet = __KPScript_Run($sKPScript, "DeleteEntry", __KPScript_StringWrap($sDatabaseFile) & " " & $sIdentification, $aUnlockKey)
	If @error Then Return SetError(1, @error, False)

	#ToDo: Check if this returns true, otherwise discard and return true
	Return $vRet ; True?

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _KPScript_EntryDeleteAll
; Description ...: Deletes all entries (in all subgroups) from a database.
; Syntax ........: _KPScript_EntryDeleteAll($sKPScript, $sDatabaseFile, $aUnlockKey)
; Parameters ....: $sKPScript           - full path to KPScript.exe.
;                  $sDatabaseFile       - full path to the kdbx file.
;                  $aUnlockKey          - an Unlock array from _KPScript_UnlockKey.
; Return values .: None
; Author ........: Seadoggie01
; Modified ......: July 7, 2020
; Remarks .......: What can I say except, rm -rf?
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _KPScript_EntryDeleteAll($sKPScript, $sDatabaseFile, $aUnlockKey)

	Local $vRet = __KPScript_Run($sKPScript, "DeleteAllEntries", __KPScript_StringWrap($sDatabaseFile), $aUnlockKey)
	If @error Then Return SetError(1, @error, False)

	Return $vRet

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _KPScript_EntryEdit
; Description ...: Edit an existing entry in a Database
; Syntax ........: _KPScript_EntryEdit($sKPScript, $sDatabaseFile, $aUnlockKey, $sIdentification, $aEntryProperties)
; Parameters ....: $sKPScript           - full path to KPScript.exe.
;                  $sDatabaseFile       - full path to the kdbx file.
;                  $aUnlockKey          - an Unlock array from _KPScript_UnlockKey.
;                  $sIdentification     - a string value from _KPScript_EntryIdentify.
;                  $aEntryProperties    - an array returned from _KPScript_EntryProperties.
; Return values .: Success - True
;                  Failure - False and sets @error:
;                  |1 - __KPScript_Run and sets @extended to @error value
; Author ........: Seadoggie
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _KPScript_EntryEdit($sKPScript, $sDatabaseFile, $aUnlockKey, $sIdentification, $aEntryProperties)

	Local $sEntryProperties = ""
	For $i=0 To UBound($aEntryProperties) - 1
		$sEntryProperties &= $aEntryProperties[$i][0] & $aEntryProperties[$i][1]
	Next

	Local $vRet = __KPScript_Run($sKPScript, "EditEntry", __KPScript_StringWrap($sDatabaseFile) & " " & $sIdentification & " " & $sEntryProperties, $aUnlockKey)
	If @error Then Return SetError(1, @error, False)

	Return $vRet ; True?

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _KPScript_EntryIdentify
; Description ...: Creates a string used in identifying entries
; Syntax ........: _KPScript_EntryIdentify($vIdentify[, $sValue = Default])
; Parameters ....: $vIdentify           - a string identifying the column or a 0-based 2D Array of [['column', 'value'], ...].
;                  $sValue              - [optional] a string value. Default is Default (only use if $vIdentify is an array).
; Return values .: Success: the string formatted to indentify entries
;                  Failure: False and sets @error = 1 - an array with incorrect dimensions was passed as $vIdentify
; Author ........: Seadoggie01
; Modified ......: August 18, 2021
; Remarks .......: Pass "all" as $vIdentify to match all entries
; Related .......: _KPScript_EntryList, _KPScript_EntryStringGet, _KPScript_EntryEdit, _KPScript_DeleteEntry, _KPScript_MoveEntry
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _KPScript_EntryIdentify($vIdentify, $sValue = Default)

	Local $aIdentify[1][2]
	If IsArray($vIdentify) Then
		If UBound($vIdentify, 0) <> 2 Then Return SetError(1, 0, False)
		If UBound($vIdentify, 2) <> 2 Then Return SetError(1, 0, False)
		$aIdentify = $vIdentify
	Else
		$aIdentify[0][0] = $vIdentify
		$aIdentify[0][1] = $sValue
	EndIf

	Local $sIdentify = ""

	For $i=0 To UBound($aIdentify) - 1
		Switch StringLower($aIdentify[$i][0])
			Case "uuid", "tags", "expires", "group", "grouppath"
				$aIdentify[$i][0] = "-refx-" & $aIdentify[$i][0]
			Case "expired"
				$aIdentify[$i][0] = "-refx-" & $aIdentify[$i][0]
				; must be bool, convert it
				$aIdentify[$i][1] = $aIdentify[$i][1] ? True : False
			Case "all"
				; Ignore all other input, this matches everything
				Return "-refx-All"
			Case Else
				; Assume the user has a custom column
				$aIdentify[$i][0] = "-ref-" & StringReplace($aIdentify[$i][0], " ", "")
		EndSwitch

		$sIdentify &= $aIdentify[$i][0] & ":"
		$sIdentify &= __KPScript_StringWrap($aIdentify[$i][1]) & " "
	Next

	$sIdentify = StringTrimRight($sIdentify, 1)

	Return $sIdentify

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _KPScript_EntryList
; Description ...: Lists all entries in a format that easily machine-readable. The output is not intended to be printed/used directly.
; Syntax ........: _KPScript_EntryList($sKPScript, $sDatabaseFile, $aUnlockKey, $sIdentification)
; Parameters ....: $sKPScript           - full path to KPScript.exe.
;                  $sDatabaseFile       - full path to the kdbx file.
;                  $aUnlockKey          - an Unlock array from _KPScript_UnlockKey.
;                  $sIdentification     - a string value from _KPScript_EntryIdentify.
; Return values .: a string containing all matching entries, properties are seperated by newlines, entries by 2 newlines
; Author ........: Seadoggie01
; Modified ......: July 7, 2020
; Remarks .......: This output is not parsed and I don't plan to parse it. Please feel free to contribute if you've correctly parsed it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _KPScript_EntryList($sKPScript, $sDatabaseFile, $aUnlockKey, $sIdentification)

	Local $vRet = __KPScript_Run($sKPScript, "ListEntries", __KPScript_StringWrap($sDatabaseFile) & " " & $sIdentification, $aUnlockKey)
	If @error Then Return SetError(1, @error, False)

	Return $vRet

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _KPScript_EntryStringGet
; Description ...: Retrieves a column's value of a record from the database file
; Syntax ........: _KPScript_EntryStringGet($sKPScript, $sDatabaseFile, $aUnlockKey, $sField, $sIdentification[, $bFail = False])
; Parameters ....: $sKPScript           - full path to KPScript.exe.
;                  $sDatabaseFile       - full path to the kdbx file.
;                  $aUnlockKey          - an Unlock array from _KPScript_UnlockKey.
;                  $sField              - name of the field to return.
;                  $sIdentification     - a string value from _KPScript_EntryIdentify.
;                  $bFail               - [optional] a boolean value. Default is False.
; Return values .: Success - the column value
;                  Failure - False and sets @error:
;                  |1 - __KPScript_Run and sets @extended to @error value
; Author ........: Seadoggie01
; Modified ......: July 7, 2020
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _KPScript_EntryStringGet($sKPScript, $sDatabaseFile, $aUnlockKey, $sField, $sIdentification, $bFail = False)

	If $bFail Then ConsoleWrite("#ToDo: Fail" & @CRLF)
	$sField = "-Field:" & __KPScript_StringWrap($sField)
	Local $sParam = $sField & " " & $sIdentification

	Local $vRet = __KPScript_Run($sKPScript, "GetEntryString", __KPScript_StringWrap($sDatabaseFile) & " " & $sParam & " -Spr", $aUnlockKey)
	If @error Then Return SetError(1, @error, False)

	Return $vRet

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _KPScript_EstimateQuality
; Description ...: Estimates the quality (in bits) of the password specified
; Syntax ........: _KPScript_EstimateQuality($sKPScript, $sPassword)
; Parameters ....: $sKPScript           - full path to KPScript.exe.
;                  $sPassword           - a string value.
; Return values .: Success - quality in bits
;                  Failure - False and sets @error:
;                  |1 - __KPScript_Run and sets @extended to @error value
; Author ........: Seadoggie01
; Modified ......: July 7, 2020
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _KPScript_EstimateQuality($sKPScript, $sPassword)

	Local $sOutput = __KPScript_Run($sKPScript, "EstimateQuality", "-text:" & __KPScript_StringWrap($sPassword))
	If @error Then Return SetError(1, @error, False)

	Return $sOutput

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _KPScript_GenPassword
; Description ...: Generates passwords using a profile.
; Syntax ........: _KPScript_GenPassword($sKPScript[, $iCount = 1[, $sProfile = Default]])
; Parameters ....: $sKPScript           - full path to KPScript.exe.
;                  $iCount              - [optional] passwords to generate. Default is 1.
;                  $sProfile            - [optional] name of the profile to use. Default is the default generator.
; Return values .: Success - A list (string) of passwords seperated by newlines
;                  Failure - False and sets @error to 1
; Author ........: Seadoggie01
; Modified ......: July 7, 2020
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _KPScript_GenPassword($sKPScript, $iCount = 1, $sProfile = Default)

	Local $sParams = ""
	If $iCount <> 1 Then $sParams = "-count:" & $iCount & " "
	If $sProfile <> Default Then $sParams &= "-profile:" & __KPScript_StringWrap($sProfile)

	Local $sPassword = __KPScript_Run($sKPScript, "GenPw", StringStripWS($sParams, $STR_STRIPTRAILING))
	If @error Then Return SetError(1, @error, False)

	Return $sPassword

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _KPScript_GroupsList
; Description ...: Lists the groups in a KeePass database file
; Syntax ........: _KPScript_GroupsList($sKPScript, $sDatabaseFile, $aUnlockKey)
; Parameters ....: $sKPScript           - full path to KPScript.exe.
;                  $sDatabaseFile       - full path to the kdbx file.
;                  $aUnlockKey          - an Unlock array from _KPScript_UnlockKey.
; Return values .: Not tested?!
; Author ........: Seadoggie01
; Modified ......: July 7, 2020
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _KPScript_GroupsList($sKPScript, $sDatabaseFile, $aUnlockKey)

	Local $vRet = __KPScript_Run($sKPScript, "ListGroups", __KPScript_StringWrap($sDatabaseFile), $aUnlockKey)
	If @error Then Return SetError(1, @error, False)

	Return $vRet

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _KPScript_UnlockKey
; Description ...: Creates an array of options to use with database related functions.
; Syntax ........: _KPScript_UnlockKey([$sPassword = Default[, $sPasswordEnc = Default[, $sKeyFile = Default[, $bUserAccount = False]]]])
; Parameters ....: $sPassword           - [optional] a string value. Default is Default.
;                  $sPasswordEnc        - [optional] a string value. Default is Default.
;                  $sKeyFile            - [optional] a string value. Default is Default.
;                  $bUserAccount        - [optional] a boolean value. Default is False.
; Return values .: a 0-based 1D array of options
; Author ........: Seadoggie01
; Modified ......: July 7, 2020
; Remarks .......: Use at least 1 method for unlocking the database
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _KPScript_UnlockKey($sPassword = Default, $sPasswordEnc = Default, $sKeyFile = Default, $bUserAccount = False)

	Local $aUnlockKey[4]
	If Not IsKeyword($sPassword) Then $aUnlockKey[$__KPUnlock_Password] = __KPScript_StringWrap($sPassword)
	If Not IsKeyword($sPasswordEnc) Then $aUnlockKey[$__KPUnlock_PasswordEnc] = __KPScript_StringWrap($sPasswordEnc)
	If Not IsKeyword($sKeyFile) Then $aUnlockKey[$__KPUnlock_KeyFile] = __KPScript_StringWrap($sKeyFile)
	$aUnlockKey[$__KPUnlock_UserAccount] = $bUserAccount

	Return $aUnlockKey

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _KPScript_ErrorDesc
; Description ...: Gets the description for a $__KPError_* value
; Syntax ........: _KPScript_ErrorDesc($__KPError)
; Parameters ....: $__KPError           - the error.
; Return values .: Success - the description
;                  Failure - False and sets @error to 1: invalid error value
; Author ........: Seadoggie
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _KPScript_ErrorDesc($__KPError)

	If $__KPError > $__KPError_Unknown Or $__KPError < $__KPError_Success Then $__KPError = $__KPError_Unknown
	Return $__aKPErrorDesc[$__KPError]

EndFunc

#Region ### Internal Functions ###

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __KPScript_Run
; Description ...: Runs KPScript.exe with the provided options
; Syntax ........: __KPScript_Run($sKPScript[, $sCmd = ""[, $sParam = ""[, $aUnlockKey = Default]]])
; Parameters ....: $sKPScript           - a string value.
;                  $sCmd                - [optional] a string value. Default is "".
;                  $sParam              - [optional] a string value. Default is "".
;                  $aUnlockKey          - [optional] an array of unknowns. Default is Default.
; Return values .: Success - the parsed data from StdOut
;                  Failure - False and sets @error:
;                  |1 - Error executing command, @extended is set to error from Run
;                  |2 - Error reading StdOut data
;                  |2 - Error parsing data, @extended is set to $__KPError_*
; Author ........: Seadoggie01
; Modified ......: July 7, 2020
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __KPScript_Run($sKPScript, $sCmd = "", $sParam = "", $aUnlockKey = Default)

	; Wrap with quotes as needed
	If StringInStr($sKPScript, " ") Then $sKPScript = '"' & $sKPScript & '"'

	If $sCmd <> "" Then $sCmd = " -c:" & $sCmd
	If $sParam <> "" Then $sParam = " " & $sParam
	If Not IsKeyword($aUnlockKey) Then
		If $aUnlockKey[$__KPUnlock_UserAccount] Then $sParam &= " -useraccount"
		If $aUnlockKey[$__KPUnlock_KeyFile] Then $sParam &= " -keyprompt"
	EndIf
	Local $sFullCommand = $sKPScript & $sCmd & $sParam
	If $bDebugging Then ConsoleWrite($sFullCommand & @CRLF)

	_UDF_Debug($sFullCommand, __KPScript_Run)

	Local $iPID = Run($sFullCommand, "", @SW_HIDE, $STDOUT_CHILD + $STDIN_CHILD)
	If @error Then Return SetError($__KPError_BadCommand, @error, False)

	If Not IsKeyword($aUnlockKey) Then

		Local $sOutput, $sLastLine, $iPos
		While True
			$sOutput &= StdoutRead($iPID)
			$iPos = StringInStr($sOutput, @CRLF, 0, -1)
			$sLastLine = $iPos ? StringRight($sOutput, StringLen($sOutput) - $iPos - 1) : $sOutput
			Switch $sLastLine
				Case "Password: "
					StdinWrite($iPID, $aUnlockKey[$__KPUnlock_Password] & @CRLF)
				Case "Key File: "
					StdinWrite($iPID, $aUnlockKey[$__KPUnlock_KeyFile] & @CRLF)
				Case "User Account (Y/N): "
					StdinWrite($iPID, $aUnlockKey[$__KPUnlock_UserAccount] ? "y" : "n" & @CRLF)
					ExitLoop
				Case ""
					; Nothing
			EndSwitch
			Sleep(100)
			If Not ProcessExists($iPID) Then ExitLoop
		WEnd
	EndIf

	ProcessWaitClose($iPID)

	$sOutput &= StdoutRead($iPID)
	If @error Then Return SetError(2, 0, False)

	If $bDebugging Then ConsoleWrite("Output: " & $sOutput)

	$sOutput = StringStripWS($sOutput, $STR_STRIPLEADING + $STR_STRIPTRAILING)
	$iPos = StringInStr($sOutput, @CRLF, 0, -1)
	$sLastLine = $iPos ? StringRight($sOutput, StringLen($sOutput) - $iPos - 1) : $sOutput

	If StringLeft($sLastLine, 4) = "OK: " Then
		Return StringStripWS(StringReplace($sOutput, "OK: Operation completed successfully.", ""), $STR_STRIPLEADING + $STR_STRIPTRAILING)
	ElseIf StringLeft($sOutput, 3) = "E: " Then
		Return SetError(3, __KPScript_KnownErrors($sOutput), $sOutput)
	Else
		ConsoleWrite("Found: " & StringLeft($sLastLine, 2) & @CRLF)
		If $bDebugging Then ConsoleWrite("Other last line: " & $sLastLine & @CRLF)
	EndIf

	Return $sOutput

EndFunc

Func __KPScript_KnownErrors($sOutput)

	If StringInStr($sOutput, "E: The composite key is invalid!") Then Return $__KPError_CompositeKey
	Return $__KPError_Unknown

EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __KPScript_StringWrap
; Description ...: Wrapes text in quotes if it contains a space
; Syntax ........: __KPScript_StringWrap($sText)
; Parameters ....: $sText               - a string value.
; Return values .: $sText, possibly wrapped in quotes
; Author ........: Seadoggie01
; Modified ......: July 7, 2020
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __KPScript_StringWrap($sText)

	If StringInStr($sText, " ") Then $sText = '"' & $sText & '"'
	Return $sText

EndFunc

#EndRegion ### Internal Functions ###
