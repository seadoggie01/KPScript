#include-once

; UDF Version 0.0.0.1

#include <AutoItConstants.au3> ; STDOUT_CHILD STDIN_CHILD
#include <StringConstants.au3> ; StripLeading + StripTrailing

#Region ### Notes ###
#CS
This UDF allows you to edit a KeePass database.
N.B.: If the database is open and unsaved in KeePass, when KeePass attempts to save, it will ask the user how to resolve: Overwrite, reload, or cancel.

Derived directly from the documentation here: https://keepass.info/help/v2_dev/scr_sc_index.html

{PASSWORD_ENC}
	I don't understand what it is used for or how it is used.
	If you need this implemented, please let me know and possibly help me figure out what it is used for and how.

REQUIRED: Download KPScript and extract to KeePass' installation directory. Download link --> https://keepass.info/plugins.html#kpscript
"The KPScript.exe file needs to be copied into the directory where KeePass is installed (where the KeePass.exe file is)."
My (full) copy installed itself to "C:\Program Files (x86)\KeePass Password Safe 2\"

Testing was completed using the FULL version of KeePass (Installed via installer), however, it "should" be compatible with Zip version

List of Commands provided by KPScript
	(N.B. return values aren't parsed well, that's step 2. First, get it working, then make it pretty)

Working
	ListGroups: Done
	ListEntries: Done
	GetEntryString: Done
	GenPw: Done
	EstimateQuality: Done
	DeleteEntry: Done

Untested but implemented
	DeleteAllEntries: Done?

Unimplemented
	AddEntry
	EditEntry
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
; Description ...:
; Syntax ........: _KPScript_EntryAdd($sKPScript, $sDatabaseFile, $aUnlockKey, $sEntryProperties)
; Parameters ....: $sKPScript           - full path to KPScript.exe.
;                  $sDatabaseFile       - full path to the kdbx file.
;                  $aUnlockKey          - an Unlock array from _KPScript_UnlockKey.
;                  $sEntryProperties    - a string returned from _KPScript_EntryProperties.
; Return values .: None
; Author ........: Seadoggie01
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _KPScript_EntryAdd($sKPScript, $sDatabaseFile, $aUnlockKey, $sEntryProperties)

	Local $vRet = __KPScript_Run($sKPScript, "AddEntry", __KPScript_StringWrap($sDatabaseFile) & " " & $sEntryProperties, $aUnlockKey)
	If @error Then Return SetError(1, @error, False)

	Return $vRet

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _KPScript_EntryDelete
; Description ...: Deletes one or more existing entries from a database.
; Syntax ........: _KPScript_EntryDelete($sKPScript, $sDatabaseFile, $aUnlockKey, $sIdentification)
; Parameters ....: $sKPScript           - full path to KPScript.exe.
;                  $sDatabaseFile       - full path to the kdbx file.
;                  $aUnlockKey          - an Unlock array from _KPScript_UnlockKey.
;                  $sIdentification     - a string value.
; Return values .: Success - True
;                  Failure - False and sets @error to 1
; Author ........: Seadoggie01
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _KPScript_EntryDelete($sKPScript, $sDatabaseFile, $aUnlockKey, $sIdentification)

	Local $vRet = __KPScript_Run($sKPScript, "DeleteEntry", __KPScript_StringWrap($sDatabaseFile) & " " & $sIdentification, $aUnlockKey)
	If @error Then Return SetError(1, @error, False)

	Return True

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
; Modified ......:
; Remarks .......:
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
; Name ..........: _KPScript_EntryIdentify
; Description ...:
; Syntax ........: _KPScript_EntryIdentify($vIdentify[, $sValue = Default])
; Parameters ....: $vIdentify           - a variant value.
;                  $sValue              - [optional] a string value. Default is Default.
; Return values .: None
; Author ........: Seadoggie01
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _KPScript_EntryIdentify($vIdentify, $sValue = Default)

	Local $aIdentify[0][0]
	If IsArray($vIdentify) Then
		$aIdentify = $vIdentify
	Else
		ReDim $aIdentify[1][2]
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
				; must be bool
				$aIdentify[$i][1] = $aIdentify[$i][1] ? True : False
			Case "all"
				; Ignore all other input, this matches everything
				Return "-refx-All"
			Case Else
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
;                  $sIdentification     - a string value.
; Return values .: None
; Author ........: Seadoggie01
; Modified ......:
; Remarks .......:
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
; Name ..........: _KPScript_EntryProperties
; Description ...:
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
; Return values .: None
; Author ........: Seadoggie01
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _KPScript_EntryProperties($vTitle = Default, $sUserName = Default, $sPassword = Default, $sURL = Default, $sNotes = Default, $sGroupName = Default, $sGroupPath = Default, $iIcon = Default, $iCustIcon = Default, $bExpires = Default, $sExpiryTime = Default)

	Local $aProperties = [["Title",""],["UserName",""],["Password",""],["URL",""],["Notes",""],["GroupName",""],["GroupPath",""],["Icon",""],["CustomIcon",""],["Expires",""],["ExpiryTime",""]]
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

	Local $sProperties = ""
	For $i=0 To UBound($aProperties) - 1
		If $aProperties[$i][1] <> "" Then
			If StringInStr("Icon|CustomIcon|Expires|ExpiryTime|", $aProperties[$i][0] & "|") Then
				$sProperties &= " -setx" & $aProperties[$i][0] & ":" & __KPScript_StringWrap($aProperties[$i][1])
			Else
				$sProperties &= " -" & $aProperties[$i][0] & ":" & __KPScript_StringWrap($aProperties[$i][1])
			EndIf
		EndIf
	Next

	Return $sProperties

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _KPScript_EntryStringGet
; Description ...:
; Syntax ........: _KPScript_EntryStringGet($sKPScript, $sDatabaseFile, $aUnlockKey, $sField, $sIdentification[, $bFail = False])
; Parameters ....: $sKPScript           - full path to KPScript.exe.
;                  $sDatabaseFile       - full path to the kdbx file.
;                  $aUnlockKey          - an Unlock array from _KPScript_UnlockKey.
;                  $sField              - a string value.
;                  $sIdentification     - a string value.
;                  $bFail               - [optional] a boolean value. Default is False.
; Return values .: None
; Author ........: Seadoggie01
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _KPScript_EntryStringGet($sKPScript, $sDatabaseFile, $aUnlockKey, $sField, $sIdentification, $bFail = False)

	If $bFail Then Exit ConsoleWrite("#ToDo: Fail" & @CRLF)
	$sField = "-Field:" & __KPScript_StringWrap($sField)
	Local $sParam = $sField & " " & $sIdentification

	Local $vRet = __KPScript_Run($sKPScript, "GetEntryString", __KPScript_StringWrap($sDatabaseFile) & " " & $sParam, $aUnlockKey)
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
;                  Failure - False and sets @error to 1
; Author ........: Seadoggie01
; Modified ......:
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
; Return values .: Success - A list of password (string)
;                  Failure - False and sets @error to 1
; Author ........: Seadoggie01
; Modified ......:
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
; Description ...:
; Syntax ........: _KPScript_GroupsList($sKPScript, $sDatabaseFile, $aUnlockKey)
; Parameters ....: $sKPScript           - full path to KPScript.exe.
;                  $sDatabaseFile       - full path to the kdbx file.
;                  $aUnlockKey          - an Unlock array from _KPScript_UnlockKey.
; Return values .: None
; Author ........: Seadoggie01
; Modified ......:
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
; Syntax ........: _KPScript_UnlockKey([$sPassword = Default[, $sPasswordEnc = Default[, $sKeyFile = Default[,
;                  $bUserAccount = False]]]])
; Parameters ....: $sPassword           - [optional] a string value. Default is Default.
;                  $sPasswordEnc        - [optional] a string value. Default is Default.
;                  $sKeyFile            - [optional] a string value. Default is Default.
;                  $bUserAccount        - [optional] a boolean value. Default is False.
; Return values .: a 0-based 1D array of options
; Author ........: Seadoggie01
; Modified ......:
; Remarks .......: Use at least 1 method for unlocking the database
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _KPScript_UnlockKey($sPassword = Default, $sPasswordEnc = Default, $sKeyFile = Default, $bUserAccount = False)

	Local $aUnlockKey[4]
	If Not IsKeyword($sPassword) Then $aUnlockKey[0] = __KPScript_StringWrap($sPassword)
	If Not IsKeyword($sPasswordEnc) Then $aUnlockKey[1] = __KPScript_StringWrap($sPasswordEnc)
	If Not IsKeyword($sKeyFile) Then $aUnlockKey[2] = __KPScript_StringWrap($sKeyFile)
	$aUnlockKey[3] = $bUserAccount

	Return $aUnlockKey

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
;                  |2 - Error parsing data, @extended is set to error from __KPScript_ParseData
; Author ........: Seadoggie01
; Modified ......:
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
		If $aUnlockKey[3] Then $sParam &= " -useraccount"
		$sParam &= " -keyprompt"
	EndIf
	Local $sFullCommand = $sKPScript & $sCmd & $sParam

	Local $iPID = Run($sFullCommand, "", @SW_HIDE, $STDOUT_CHILD + $STDIN_CHILD)
	If @error Then Return SetError(1, @error, False)

	If Not IsKeyword($aUnlockKey) Then

		Local $sOutput, $sLastLine, $iPos
		While True
			$sOutput = StdoutRead($iPID)
			$iPos = StringInStr($sOutput, @CRLF, 0, -1)
			$sLastLine = $iPos ? StringRight($sOutput, StringLen($sOutput) - $iPos - 1) : $sOutput
			Switch $sLastLine
				Case "Password: "
					StdinWrite($iPID, $aUnlockKey[0] & @CRLF)
				Case "Key File: "
					StdinWrite($iPID, $aUnlockKey[1] & @CRLF)
				Case "User Account (Y/N): "
					StdinWrite($iPID, $aUnlockKey[3] ? "y" : "n" & @CRLF)
					ExitLoop
				Case ""
					; Nothing
			EndSwitch
			Sleep(100)
			If Not ProcessExists($iPID) Then ExitLoop
		WEnd
	EndIf

	ProcessWaitClose($iPID)

	Local $sParsed = __KPScript_ParseOutput($iPID)
	If @error Then Return SetError(2, @error, False)

	Return $sParsed

EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __KPScript_ParseOutput
; Description ...: Checks the last line of output from a command for an error and strips it.
; Syntax ........: __KPScript_ParseOutput($iPID)
; Parameters ....: $iPID                - an integer value.
; Return values .: Success - the parsed output
;                  Failure - Sets @error:
;                  |1 - Unable to read from Stdout, Returns False
;                  |2 - Execution returned an error, Returns output
; Author ........: Seadoggie01
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __KPScript_ParseOutput($iPID)

	Local $sOutput = StdoutRead($iPID)
	If @error Then Return SetError(1, 0, False)

	$sOutput = StringStripWS($sOutput, $STR_STRIPLEADING + $STR_STRIPTRAILING)
	Local $sLastLine, $iPos = StringInStr($sOutput, @CRLF, 0, -1)
	If $iPos Then
		$sLastLine = StringRight($sOutput, StringLen($sOutput) - $iPos - 1)
	Else
		$sLastLine = $sOutput
	EndIf
	If StringLeft($sLastLine, 3) = "OK:" Then
		Return StringStripWS(StringReplace($sOutput, "OK: Operation completed successfully.", ""), $STR_STRIPLEADING + $STR_STRIPTRAILING)
	ElseIf StringLeft($sLastLine, 2) = "E:" Then
		Return SetError(2, 0, $sOutput)
	EndIf

	Return $sOutput

EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __KPScript_StringWrap
; Description ...: Wrapes text in quotes if it contains a space
; Syntax ........: __KPScript_StringWrap($sText)
; Parameters ....: $sText               - a string value.
; Return values .: $sText, possibly wrapped in quotes
; Author ........: Seadoggie01
; Modified ......:
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
