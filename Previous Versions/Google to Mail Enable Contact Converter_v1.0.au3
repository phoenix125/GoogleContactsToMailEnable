#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Resources\phoenix.ico
#AutoIt3Wrapper_Outfile=Builds\GoogleContactsToMailEnable_v1.0.exe
#AutoIt3Wrapper_Outfile_x64=Builds\GoogleContactsToMailEnable_64-bit(x64)_v1.0.exe
#AutoIt3Wrapper_Res_Comment=By Phoenix125.com
#AutoIt3Wrapper_Res_Description=Google Contacts vCard to Mail Enable Importer
#AutoIt3Wrapper_Res_Fileversion=1.0
#AutoIt3Wrapper_Res_ProductName=Google Contacts To Mail Enable Importer
#AutoIt3Wrapper_Res_ProductVersion=1.0
#AutoIt3Wrapper_Res_CompanyName=http://www.Phoenix125.com
#AutoIt3Wrapper_Res_LegalCopyright=http://www.Phoenix125.com
#AutoIt3Wrapper_Res_SaveSource=y
#AutoIt3Wrapper_Run_AU3Check=n
#AutoIt3Wrapper_AU3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_Tidy_Stop_OnError=n
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/mo
#AutoIt3Wrapper_Res_Icon_Add=Resources\discord.ico
#AutoIt3Wrapper_Res_Icon_Add=Resources\info.ico
#AutoIt3Wrapper_Res_Icon_Add=Resources\forum.ico
#AutoIt3Wrapper_Res_Icon_Add=Resources\manual.ico
#AutoIt3Wrapper_Res_Icon_Add=Resources\refresh.ico
#AutoIt3Wrapper_Res_Icon_Add=Resources\refreshnotice.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
Global $aUtilName = "Google Contacts to Mail Enable"
Global $aUtilFileName = StringStripWS($aUtilName, 8)
Global $aFolderTemp = @ScriptDir & "\" & $aUtilFileName & "UtilFiles\"
If @Compiled = 0 Then
	Global $aIconFile = @ScriptDir & "\Google Contacts to Mail Enable Icons.exe"
Else
	Global $aIconFile = @ScriptFullPath
EndIf
If Not FileExists($aFolderTemp) Then
	Do
		DirCreate($aFolderTemp)
	Until FileExists($aFolderTemp)
EndIf
FileInstall("K:\AutoIT\_MyProgs\GoogleToMailEnable\Resources\phoenixlogo.jpg", $aFolderTemp, 0)
FileInstall("K:\AutoIT\_MyProgs\GoogleToMailEnable\Resources\MailEnableLogo1.jpg", $aFolderTemp, 0)
FileInstall("K:\AutoIT\_MyProgs\GoogleToMailEnable\Resources\GoogleLogo.jpg", $aFolderTemp, 0)
FileInstall("K:\AutoIT\_MyProgs\GoogleToMailEnable\Resources\GoogleExport1.jpg", $aFolderTemp, 0)
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <File.au3>
#include <GUIConstantsEx.au3>
#include <Inet.au3>
#include <StaticConstants.au3>
#include <String.au3>
#include <Timers.au3>
#include <WindowsConstants.au3>
Local $aUtilVer = "v1.0" ; (2021-02-20)
Global $aUtilVerNumber = 0
Opt("GUIOnEventMode", 1)
Global $aNumberOfContacts = 0
Global $aIniFile = @ScriptDir & "\" & $aUtilFileName & ".ini"
Global $aINIHeader = " --------------- " & StringUpper($aUtilName) & " --------------- "
Global $cSWButtonStartServer = "0xA5B89B" ; Faded Green
Global $cSWButtonStopServer = "0xB89B9B" ; Faded Red
Global $cSWButtonRestartUtil = "0xD3D17F" ; Faded Yellow
Global $cButtonFadedBlue = "0xCBDDFF" ; Light Blue
Global $cSWTitle = "0xDC143C" ; Darker Red
Global $aServerUpdateLinkVer = "http://www.phoenix125.com/share/mailenable/latestver.txt"
Global $aServerUpdateLinkDL = "http://www.phoenix125.com/share/mailenable/GoogleContactsToMailEnable.zip"
Global $aUpdateAvailableTF = False

_IniRead()
_IniWrite()
_ShowGUI()
While True ;**** Loop Until Closed ****
	Sleep(100)
WEnd
Func _SetGlobalVariablesGoogle($tMax)
	Global $gFN[$tMax]
	Global $gNameFirst[$tMax]
	Global $gNameMiddle[$tMax]
	Global $gNameLast[$tMax]
	Global $gNamePre[$tMax]
	Global $gNameSuffix[$tMax]
	Global $gNickName[$tMax]
	Global $gEmailMain[$tMax]
	Global $gPhoneWork[$tMax]
	Global $gPhoneHome[$tMax]
	Global $gPhoneOther[$tMax]
	Global $gPhoneMobile[$tMax]
	Global $gPhoneWorkFax[$tMax]
	Global $gAddressHomePOBox[$tMax] ;
	Global $gAddressHomeStreet1[$tMax]
	Global $gAddressHomeStreet2[$tMax]
	Global $gAddressHomeCity[$tMax]
	Global $gAddressHomeState[$tMax]
	Global $gAddressHomeZIP[$tMax]
	Global $gAddressHomeCountry[$tMax]
	Global $gAddressWorkPOBox[$tMax] ;
	Global $gAddressWorkStreet1[$tMax]
	Global $gAddressWorkStreet2[$tMax]
	Global $gAddressWorkCity[$tMax]
	Global $gAddressWorkState[$tMax]
	Global $gAddressWorkZIP[$tMax]
	Global $gAddressWorkCountry[$tMax]
	Global $gAddressOtherPOBox[$tMax] ;
	Global $gAddressOtherStreet1[$tMax]
	Global $gAddressOtherStreet2[$tMax]
	Global $gAddressOtherCity[$tMax]
	Global $gAddressOtherState[$tMax]
	Global $gAddressOtherZIP[$tMax]
	Global $gAddressOtherCountry[$tMax]
	Global $gBirthday[$tMax]
	Global $gCompany[$tMax]
	Global $gDepartment[$tMax]
	Global $gURLWork[$tMax]
	Global $gJobTitle[$tMax]
	Global $gHomePage[$tMax]
	Global $gRelatedLabel[$tMax]
	Global $gRelatedName[$tMax]
	Global $gNote[$tMax]
	Global $gCustom[$tMax]
	Global $gURLBlog[$tMax]
	Global $gURLProfile[$tMax]
	Global $gCategories[$tMax]
	Global $gPhotoURL[$tMax]
	Global $gEmailWork[$tMax]
	Global $gPhonePager[$tMax]
	Global $gEmailOther[$tMax]
	Global $gChatYAHOO[$tMax]
	Global $gChatAIM[$tMax]
	Global $gChatGTALK[$tMax]
	Global $gChatICQ[$tMax]
	Global $gChatQQ[$tMax]
	Global $gChatSkype[$tMax]
	Global $gChatMSN[$tMax]
	Global $gChatJabber[$tMax]
	Global $gCategNames[0]
	Global $gCategContacts[0]
EndFunc   ;==>_SetGlobalVariablesGoogle
Func _SetGlobalVariablesMailEnable($tMax)
	Global $mFN[$tMax]
	Global $mN[$tMax]
	Global $mNICKNAME[$tMax]
	Global $mEMAIL_PREF_INTERNET[$tMax]
	Global $mTEL_HOME_VOICE[$tMax]
	Global $mTEL_CELL_VOICE[$tMax]
	Global $mTEL_PAGER_VOICE[$tMax]
	Global $mNOTE[$tMax]
	Global $mTEL_WORK_VOICE[$tMax]
	Global $mTEL_WORK_FAX[$tMax]
	Global $mADR_HOME[$tMax]
	Global $mADR_WORK[$tMax]
	Global $mADR_OTHER[$tMax]
	Global $mBDAY[$tMax]
	Global $mORG[$tMax]
	Global $mURL_WORK[$tMax]
	Global $mTITLE[$tMax]
	Global $mFileName[$tMax]
	Global $mIndex[$tMax]
	Global $mUID[$tMax]

	Global $mEMAIL_WORK_INTERNET[$tMax]
	Global $mEMAIL_OTHER_INTERNET[$tMax]
	Global $mTYPE_MSMESSENGER[$tMax]
EndFunc   ;==>_SetGlobalVariablesMailEnable
Func _IniRead()
	Global $aExportFolder = _RemoveTrailingSlash(IniRead($aIniFile, $aINIHeader, "ExportFolder", "C:\Program Files (x86)\Mail Enable\Postoffices"))
	Global $gGoogleSourceFile = IniRead($aIniFile, $aINIHeader, "GoogleSourceFile", @ScriptDir & "\contacts.vcf")
	Global $aAddGroupsToNotesYN = IniRead($aIniFile, $aINIHeader, "AddGroupsToNotesYN", "no")
	Global $aDownloadPhotosYN = IniRead($aIniFile, $aINIHeader, "DownloadPhotosYN", "yes")
	Global $aCopyPhonesNosToMobileYN = IniRead($aIniFile, $aINIHeader, "CopyPhonesNosToMobileYN", "no")
	Global $aCopyWorkorOtherEmailsToPrefYN = IniRead($aIniFile, $aINIHeader, "CopyWorkorOtherEmailsToPrefYN", "yes")
	Global $aReformatPhoneNumbersYN = IniRead($aIniFile, $aINIHeader, "ReformatPhoneNumbersYN", "yes")
	Global $aWorkAliases = IniRead($aIniFile, $aINIHeader, "WorkAliases", "Work,Office,Corporate")
	Global $aMobileAliases = IniRead($aIniFile, $aINIHeader, "MobileAliases", "Mobile,Cell")
	Global $aHomeAliases = IniRead($aIniFile, $aINIHeader, "HomeAliases", "Home")
	Global $aCreateGroupsYN = IniRead($aIniFile, $aINIHeader, "CreateGroupsYN", "yes")
	Global $aOpenFolderWhenDoneYN = IniRead($aIniFile, $aINIHeader, "OpenFolderWhenDoneYN", "yes")
	Global $aCheckForUpdatesYN = IniRead($aIniFile, $aINIHeader, "CheckForUpdatesYN", "yes")
	Global $aIndex_XML = $aExportFolder & "\_index.xml"
	Global $aGroups_XML = $aExportFolder & "\_GROUPS.XML"
	Global $aUtilLastVerNumber = IniRead($aIniFile, $aINIHeader, "UtilLastVerNumber", $aUtilVerNumber)
EndFunc   ;==>_IniRead
Func _IniWrite()
	IniWrite($aIniFile, $aINIHeader, "ExportFolder", $aExportFolder)
	IniWrite($aIniFile, $aINIHeader, "GoogleSourceFile", $gGoogleSourceFile)
	IniWrite($aIniFile, $aINIHeader, "AddGroupsToNotesYN", $aAddGroupsToNotesYN)
	IniWrite($aIniFile, $aINIHeader, "DownloadPhotosYN", $aDownloadPhotosYN)
	IniWrite($aIniFile, $aINIHeader, "CopyPhonesNosToMobileYN", $aCopyPhonesNosToMobileYN)
	IniWrite($aIniFile, $aINIHeader, "CopyWorkorOtherEmailsToPrefYN", $aCopyWorkorOtherEmailsToPrefYN)
	IniWrite($aIniFile, $aINIHeader, "ReformatPhoneNumbersYN", $aReformatPhoneNumbersYN)
	IniWrite($aIniFile, $aINIHeader, "WorkAliases", $aWorkAliases)
	IniWrite($aIniFile, $aINIHeader, "MobileAliases", $aMobileAliases)
	IniWrite($aIniFile, $aINIHeader, "HomeAliases", $aHomeAliases)
	IniWrite($aIniFile, $aINIHeader, "CreateGroupsYN", $aCreateGroupsYN)
	IniWrite($aIniFile, $aINIHeader, "OpenFolderWhenDoneYN", $aOpenFolderWhenDoneYN)
	IniWrite($aIniFile, $aINIHeader, "CheckForUpdatesYN", $aCheckForUpdatesYN)
	IniWrite($aIniFile, $aINIHeader, "UtilLastVerNumber", $aUtilVerNumber)
EndFunc   ;==>_IniWrite
Func _ReadGoogleVCAR()
	Global $gFileReadPre
	Global $gFileRead[0]
	ControlSetText($aSplash, "", "Static1", "Reading Google VCF File." & @CRLF & @CRLF & "Timer:" & Round(_Timer_Diff($aTimerStart) / 1000) & "s")
	_FileReadToArray($gGoogleSourceFile, $gFileReadPre)
	$gUbound = UBound($gFileReadPre) - 1
	$tCnt = 0
	For $t0 = 0 To $gUbound
		$tCnt += 1
		If $tCnt = 100 Then
			ControlSetText($aSplash, "", "Static1", "Combining split fields." & @CRLF & "Processing line " & $t0 & " of " & $gUbound & _
					@CRLF & @CRLF & "Timer:" & Round(_Timer_Diff($aTimerStart) / 1000) & "s")
			$tCnt = 0
		EndIf
		Local $tTxt = $gFileReadPre[$t0]
		If $t0 <= $gUbound - 1 Then
			If StringLeft($gFileReadPre[$t0 + 1], 1) = " " Then
				Do
					If $t0 <= $gUbound + 1 Then
						$t0 += 1
						$tTxt &= StringTrimLeft($gFileReadPre[$t0], 1)
					EndIf
				Until StringLeft($gFileReadPre[$t0 + 1], 1) <> " "
			EndIf
		EndIf
		_ArrayAdd($gFileRead, $tTxt)
	Next
	$gUbound = UBound($gFileRead) - 1
	For $t1 = 0 To ($gUbound)
		If $gFileRead[$t1] = "BEGIN:VCARD" Then
			$aNumberOfContacts += 1
			ControlSetText($aSplash, "", "Static1", "Counting number of contacts." & @CRLF & "Found " & $aNumberOfContacts & " contacts" & _
					@CRLF & @CRLF & "Timer:" & Round(_Timer_Diff($aTimerStart) / 1000) & "s")
		EndIf
	Next
	_SetGlobalVariablesGoogle($aNumberOfContacts)
	Local $tNo = -1
	Local $tCnt = 0
	For $t1 = 0 To $gUbound
		If $gFileRead[$t1] = "BEGIN:VCARD" Then
			$tCnt += 1
			If $tCnt = 10 Then
				ControlSetText($aSplash, "", "Static1", "Reading Gmail VCF file." & @CRLF & "Processing contact " & $t1 & " of " & $aNumberOfContacts & _
						@CRLF & @CRLF & "Timer:" & Round(_Timer_Diff($aTimerStart) / 1000) & "s")
				$tCnt = 0
			EndIf
			Local $tNextTemp = ""
			Local $tRelatedTemp = ""
			Local $tExtendedPre = ""
			Local $tExtendedPost = ""
			$tNo += 1
			If $tNo >= $aNumberOfContacts Then
				MsgBox(0, "Kim", "Error. More More contacts than calculated total." & @CRLF & "Count:" & $tNo & " | Total:" & $aNumberOfContacts, 30)
				ExitLoop
			EndIf
			For $t2 = $t1 To $gUbound
				If $gFileRead[$t2] = "END:VCARD" Then ExitLoop
				Local $tExtendedPre = $gFileRead[$t2]
				Local $tPre = "!NoColon!"
				Local $tPost = ""
				For $t3 = 1 To StringLen($gFileRead[$t2])
					If StringMid($gFileRead[$t2], $t3, 1) = ":" Then
						$tPre = StringLeft($gFileRead[$t2], ($t3 - 1))
						$tPost = StringRight($gFileRead[$t2], (StringLen($gFileRead[$t2]) - $t3))
						ExitLoop
					EndIf
				Next
				Local $tNoNumber = StringRegExpReplace($tPre, "\d", "")
				If StringLeft($tNoNumber, 5) = "item." Then $tNoNumber = StringTrimLeft($tNoNumber, 5)
				Local $tNumber = StringRegExpReplace($tPre, "[A-Za-z.,=;-]", "")
				If $tPre = "FN" Then
					$gFN[$tNo] = $tPost
				ElseIf $tPre = "N" Then
					Local $tSplit = StringSplit($tPost, ";")
					If UBound($tSplit) > 0 Then $gNameLast[$tNo] = $tSplit[1]
					If UBound($tSplit) > 1 Then $gNameFirst[$tNo] = $tSplit[2]
					If UBound($tSplit) > 2 Then $gNameMiddle[$tNo] = $tSplit[3]
					If UBound($tSplit) > 3 Then $gNamePre[$tNo] = $tSplit[4]
					If UBound($tSplit) > 4 Then $gNameSuffix[$tNo] = $tSplit[5]
				ElseIf $tPre = "NICKNAME" Then
					$gNickName[$tNo] = $tPost
				ElseIf $tPre = "TEL;TYPE=WORK" Then
					If $gPhoneWork[$tNo] = "" Then
						$gPhoneWork[$tNo] = $tPost
					Else
						$gCustom[$tNo] &= "Work:" & $tNextTemp & @CRLF
					EndIf
				ElseIf $tPre = "TEL;TYPE=HOME" Then
					If $gPhoneHome[$tNo] = "" Then
						$gPhoneHome[$tNo] = $tPost
					Else
						$gCustom[$tNo] &= "Home:" & $tNextTemp & @CRLF
					EndIf
				ElseIf $tPre = "TEL;TYPE=PAGER" Then
					If $gPhonePager[$tNo] = "" Then
						$gPhonePager[$tNo] = $tPost
					Else
						$gCustom[$tNo] &= "Pager:" & $tNextTemp & @CRLF
					EndIf
				ElseIf $tPre = "TEL" Then
					$gPhoneOther[$tNo] &= "Phone Other:" & $tPost & @CRLF
				ElseIf $tPre = "TEL;TYPE=CELL" Then
					If $gPhoneMobile[$tNo] = "" Then
						$gPhoneMobile[$tNo] = $tPost
					Else
						$gCustom[$tNo] &= "Cell:" & $tNextTemp & @CRLF
					EndIf
				EndIf
				If $tRelatedTemp <> "" Then
					$gRelatedName[$tNo] = $tRelatedTemp
					Local $tTxt = StringReplace($tPost, "_$!<", "")
					$tTxt = StringReplace($tTxt, ">!$_", "")
					If $tNoNumber = "X-ABLabel" Then $gRelatedLabel[$tNo] = $tTxt
					$tRelatedTemp = ""
					$tNoNumber = ""
				EndIf
				If $tNoNumber = "EMAIL;TYPE=INTERNET;TYPE=HOME" Then
					If $gEmailMain[$tNo] = "" Then
						$gEmailMain[$tNo] = $tPost
					Else
						$gCustom[$tNo] &= "Home Email:" & $tPost & @CRLF
					EndIf
				ElseIf $tNoNumber = "EMAIL;TYPE=INTERNET;TYPE=WORK" Then
					If $gEmailWork[$tNo] = "" Then
						$gEmailWork[$tNo] = $tPost
					Else
						$gCustom[$tNo] &= "Work Email:" & $tPost & @CRLF
					EndIf
				ElseIf $tNoNumber = "EMAIL;TYPE=INTERNET" Then
					If $gEmailMain[$tNo] = "" Then
						$gEmailMain[$tNo] = $tPost
					Else
						If $t1 = $gUbound Then
							$gCustom[$tNo] &= "Email:" & $tPost & @CRLF
						Else
							Local $tPre1 = ""
							Local $tPost1 = ""
							For $t3 = 1 To StringLen($gFileRead[$t2 + 1])
								If StringMid($gFileRead[$t2 + 1], $t3, 1) = ":" Then
									$tPre1 = StringLeft($gFileRead[$t2 + 1], ($t3 - 1))
									$tPost1 = StringRight($gFileRead[$t2 + 1], (StringLen($gFileRead[$t2 + 1]) - $t3))
									ExitLoop
								EndIf
							Next
							Local $tNoNumber1 = StringRegExpReplace($tPre1, "\d", "")
							If StringLeft($tNoNumber1, 5) = "item." Then $tNoNumber1 = StringTrimLeft($tNoNumber1, 5)
							Local $tLabel1 = _ConvertCustomLabel($tPost1)
							If $tNoNumber1 = "X-ABLabel" Then
								If $gEmailOther[$tNo] = "" Then
									$gEmailOther[$tNo] = $tPost
								Else
									$gCustom[$tNo] &= $tLabel1 & ":" & $tPost & @CRLF
								EndIf
							Else
								If $gEmailOther[$tNo] = "" Then
									$gEmailOther[$tNo] = $tPost
								Else
									$gCustom[$tNo] &= "Email:" & $tPost & @CRLF
								EndIf
							EndIf
						EndIf
						$tLabel = ""
					EndIf
				ElseIf $tNoNumber = "TITLE" Then
					$gJobTitle[$tNo] = $tPost
				ElseIf $tNoNumber = "TEL" Then
					$tNextTemp = $tPost
				ElseIf $tNoNumber = "URL" Then
					$tNextTemp = $tPost
				ElseIf $tNoNumber = "X-ABRELATEDNAMES" Then
					$tRelatedTemp = $tPost
				ElseIf $tNoNumber = "X-ABDATE" Then
					$tNextTemp = $tPost
				ElseIf $tNoNumber = "PHOTO" Then
					$gPhotoURL[$tNo] = StringReplace($tPost, "\:", ":")
				ElseIf $tNoNumber = "ORG" Then
					Local $tSplit = StringSplit($tPost, ";")
					If UBound($tSplit) > 1 Then $gCompany[$tNo] = $tSplit[1]
					If UBound($tSplit) > 2 Then $gDepartment[$tNo] = $tSplit[2]
				ElseIf $tNoNumber = "X-ABLabel" Then
					If StringLeft($tPost, 4) = "_$!<" Then
						Local $tString = StringTrimLeft($tPost, 4)
						$tPost = StringTrimRight($tString, 4)
					EndIf
					Local $tLabel = _ConvertCustomLabel($tPost)
					If $tLabel = "Home Fax" Then
						If $gPhoneOther[$tNo] = "" Then
							$gPhoneOther[$tNo] &= "Home Fax:" & $tNextTemp & @CRLF
						Else
							$gCustom[$tNo] &= "Home Fax:" & $tNextTemp & @CRLF
						EndIf
						$tLabel = ""
					ElseIf $tLabel = "Work Fax" Then
						If $gPhoneWorkFax[$tNo] = "" Then
							$gPhoneWorkFax[$tNo] = $tNextTemp
						Else
							$gCustom[$tNo] &= "Work Fax:" & $tNextTemp & @CRLF
						EndIf
						$tLabel = ""
					ElseIf $tLabel = "Home Page" Then
						If $gHomePage[$tNo] = "" Then
							$gHomePage[$tNo] = $tNextTemp
						Else
							$gCustom[$tNo] &= "Home Page:" & $tNextTemp & @CRLF
						EndIf
						$tLabel = ""
					ElseIf $tLabel = "Blog" Then
						If $gURLBlog[$tNo] = "" Then
							$gURLBlog[$tNo] = $tNextTemp
						Else
							$gCustom[$tNo] &= "Blog:" & $tNextTemp & @CRLF
						EndIf
						$tLabel = ""
					ElseIf $tLabel = "Profile" Then
						$gURLProfile[$tNo] = $tNextTemp
						$tLabel = ""
					Else
						If $tLabel <> "" And $tNextTemp <> "" Then
							Local $tFound = False
							Local $tHome = StringSplit($aHomeAliases, ",")
							If UBound($tHome) > 1 Then
								For $t = 1 To (UBound($tHome) - 1)
									If $tHome[$t] = $tLabel Then
										If $gPhoneHome[$tNo] = "" Then
											$gPhoneHome[$tNo] = $tNextTemp
											$tFound = True
											ExitLoop
										EndIf
									EndIf
								Next
							EndIf
							Local $tWork = StringSplit($aWorkAliases, ",")
							If UBound($tWork) > 1 And $tFound = False Then
								For $t = 1 To (UBound($tWork) - 1)
									If $tWork[$t] = $tLabel Then
										If $gPhoneWork[$tNo] = "" Then
											$gPhoneWork[$tNo] = $tNextTemp
											$tFound = True
											ExitLoop
										EndIf
									EndIf
								Next
							EndIf
							Local $tMobile = StringSplit($aMobileAliases, ",")
							If UBound($tMobile) > 1 And $tFound = False Then
								For $t = 1 To (UBound($tMobile) - 1)
									If $tMobile[$t] = $tLabel Then
										If $gPhoneMobile[$tNo] = "" Then
											$gPhoneMobile[$tNo] = $tNextTemp
											$tFound = True
											ExitLoop
										EndIf
									EndIf
								Next
							EndIf
							If $tFound = False Then $gCustom[$tNo] &= $tLabel & ":" & $tNextTemp & @CRLF
						EndIf
					EndIf
					$tNextTemp = ""
				EndIf
				If $tPre = "ADR;TYPE=HOME" Then
					Local $tSplit = StringSplit($tPost, ";")
					If UBound($tSplit) > 1 Then $gAddressHomePOBox[$tNo] = $tSplit[1]
					If UBound($tSplit) > 2 Then $gAddressHomeStreet2[$tNo] = $tSplit[2]
					If UBound($tSplit) > 3 Then $gAddressHomeStreet1[$tNo] = $tSplit[3]
					If UBound($tSplit) > 4 Then $gAddressHomeCity[$tNo] = $tSplit[4]
					If UBound($tSplit) > 5 Then $gAddressHomeState[$tNo] = $tSplit[5]
					If UBound($tSplit) > 6 Then $gAddressHomeZIP[$tNo] = $tSplit[6]
					If UBound($tSplit) > 7 Then $gAddressHomeCountry[$tNo] = $tSplit[7]
				ElseIf $tPre = "ADR;TYPE=WORK" Then
					Local $tSplit = StringSplit($tPost, ";")
					If UBound($tSplit) > 1 Then $gAddressWorkPOBox[$tNo] = $tSplit[1]
					If UBound($tSplit) > 2 Then $gAddressWorkStreet2[$tNo] = $tSplit[2]
					If UBound($tSplit) > 3 Then $gAddressWorkStreet1[$tNo] = $tSplit[3]
					If UBound($tSplit) > 4 Then $gAddressWorkCity[$tNo] = $tSplit[4]
					If UBound($tSplit) > 5 Then $gAddressWorkState[$tNo] = $tSplit[5]
					If UBound($tSplit) > 6 Then $gAddressWorkZIP[$tNo] = $tSplit[6]
					If UBound($tSplit) > 7 Then $gAddressWorkCountry[$tNo] = $tSplit[7]
				ElseIf $tPre = "ADR" Then
					Local $tSplit = StringSplit($tPost, ";")
					If UBound($tSplit) > 1 Then $gAddressOtherPOBox[$tNo] = $tSplit[1]
					If UBound($tSplit) > 2 Then $gAddressOtherStreet2[$tNo] = $tSplit[2]
					If UBound($tSplit) > 3 Then $gAddressOtherStreet1[$tNo] = $tSplit[3]
					If UBound($tSplit) > 4 Then $gAddressOtherCity[$tNo] = $tSplit[4]
					If UBound($tSplit) > 5 Then $gAddressOtherState[$tNo] = $tSplit[5]
					If UBound($tSplit) > 6 Then $gAddressOtherZIP[$tNo] = $tSplit[6]
					If UBound($tSplit) > 7 Then $gAddressOtherCountry[$tNo] = $tSplit[7]
				ElseIf $tPre = "URL;TYPE=WORK" Then
					$gURLWork[$tNo] = $tPost
				ElseIf $tPre = "BDAY" Then
					$gBirthday[$tNo] = $tPost
				ElseIf $tPre = "NOTE" Then
					$gNote[$tNo] = $tPost
				ElseIf $tPre = "YAHOO" Then
					$gChatYAHOO[$tNo] = $tPost
				ElseIf $tPre = "AIM" Then
					$gChatAIM[$tNo] = $tPost
				ElseIf $tPre = "GTALK" Then
					$gChatGTALK[$tNo] = $tPost
				ElseIf $tPre = "ICQ" Then
					$gChatICQ[$tNo] = $tPost
				ElseIf $tPre = "JABBER" Then
					$gChatJabber[$tNo] = $tPost
				ElseIf $tPre = "SKYPE" Then
					$gChatSkype[$tNo] = $tPost
				ElseIf $tPre = "QQ" Then
					$gChatQQ[$tNo] = $tPost
				ElseIf $tPre = "MSN" Then
					$gChatMSN[$tNo] = $tPost
				ElseIf $tPre = "CATEGORIES" Then
					$gCategories[$tNo] = $tPost
				EndIf
			Next
			$gPhoneOther[$tNo] = StringTrimRight($gPhoneOther[$tNo], 2)
		EndIf
	Next
EndFunc   ;==>_ReadGoogleVCAR
Func _ConvertGoogleToMailEnable()
	Local $tCnt = 0
	For $tNo = 0 To ($aNumberOfContacts - 1)
		$tCnt += 1
		If $tCnt = 10 Then
			ControlSetText($aSplash, "", "Static1", "Converting to Mail Enable files." & @CRLF & "Processing contact " & $tNo & " of " & $aNumberOfContacts & _
					@CRLF & @CRLF & "Timer:" & Round(_Timer_Diff($aTimerStart) / 1000) & "s")
			$tCnt = 0
		EndIf
		Local $tAdd2Home = ""
		Local $tAddPOHome = ""
		Local $tAdd2Other = ""
		Local $tAddPOOther = ""
		Local $tAdd2Work = ""
		Local $tAddPOWork = ""
		$mFN[$tNo] = _ReplaceChar($gFN[$tNo])
		$mN[$tNo] = _ReplaceChar($gNameLast[$tNo] & ";" & $gNameFirst[$tNo] & ";" & $gNameMiddle[$tNo] & ";" & $gNameSuffix[$tNo] & ";" & $gNamePre[$tNo])
		$mNICKNAME[$tNo] = _ReplaceChar($gNickName[$tNo])
		$mEMAIL_PREF_INTERNET[$tNo] = $gEmailMain[$tNo]
		$mEMAIL_WORK_INTERNET[$tNo] = $gEmailWork[$tNo]
		$mEMAIL_OTHER_INTERNET[$tNo] = $gEmailOther[$tNo]
		$mTEL_WORK_VOICE[$tNo] = _ReplaceChar($gPhoneWork[$tNo])
		$mTEL_HOME_VOICE[$tNo] = _ReplaceChar($gPhoneHome[$tNo])
		$mTEL_CELL_VOICE[$tNo] = _ReplaceChar($gPhoneMobile[$tNo])
		$mTEL_PAGER_VOICE[$tNo] = _ReplaceChar($gPhonePager[$tNo])
		$mNOTE[$tNo] = StringReplace($gNote[$tNo], "\n", @CRLF) & @CRLF & @CRLF
		If $gChatGTALK[$tNo] <> "" Then
			If $mTYPE_MSMESSENGER[$tNo] = "" Then
				$mTYPE_MSMESSENGER[$tNo] = _ReplaceChar($gChatGTALK[$tNo])
			Else
				$mNOTE[$tNo] &= "Chat Google Talk:" & $gChatGTALK[$tNo] & @CRLF
			EndIf
		EndIf
		If $gChatSkype[$tNo] <> "" Then
			If $mTYPE_MSMESSENGER[$tNo] = "" Then
				$mTYPE_MSMESSENGER[$tNo] = _ReplaceChar($gChatSkype[$tNo])
			Else
				$mNOTE[$tNo] &= "Chat Skype:" & $gChatSkype[$tNo] & @CRLF
			EndIf
		EndIf
		If $gChatYAHOO[$tNo] <> "" Then
			If $mTYPE_MSMESSENGER[$tNo] = "" Then
				$mTYPE_MSMESSENGER[$tNo] = _ReplaceChar($gChatYAHOO[$tNo])
			Else
				$mNOTE[$tNo] &= "Chat Yahoo:" & $gChatYAHOO[$tNo] & @CRLF
			EndIf
		EndIf
		If $gChatMSN[$tNo] <> "" Then
			If $mTYPE_MSMESSENGER[$tNo] = "" Then
				$mTYPE_MSMESSENGER[$tNo] = _ReplaceChar($gChatMSN[$tNo])
			Else
				$mNOTE[$tNo] &= "Chat MSN:" & $gChatMSN[$tNo] & @CRLF
			EndIf
		EndIf
		If $gChatICQ[$tNo] <> "" Then
			If $mTYPE_MSMESSENGER[$tNo] = "" Then
				$mTYPE_MSMESSENGER[$tNo] = _ReplaceChar($gChatICQ[$tNo])
			Else
				$mNOTE[$tNo] &= "Chat ICQ:" & $gChatICQ[$tNo] & @CRLF
			EndIf
		EndIf
		If $gChatAIM[$tNo] <> "" Then
			If $mTYPE_MSMESSENGER[$tNo] = "" Then
				$mTYPE_MSMESSENGER[$tNo] = _ReplaceChar($gChatAIM[$tNo])
			Else
				$mNOTE[$tNo] &= "Chat AIM:" & $gChatAIM[$tNo] & @CRLF
			EndIf
		EndIf
		If $gChatQQ[$tNo] <> "" Then
			If $mTYPE_MSMESSENGER[$tNo] = "" Then
				$mTYPE_MSMESSENGER[$tNo] = _ReplaceChar($gChatQQ[$tNo])
			Else
				$mNOTE[$tNo] &= "Chat QQ:" & $gChatQQ[$tNo] & @CRLF
			EndIf
		EndIf
		If $gChatJabber[$tNo] <> "" Then
			If $mTYPE_MSMESSENGER[$tNo] = "" Then
				$mTYPE_MSMESSENGER[$tNo] = _ReplaceChar($gChatJabber[$tNo])
			Else
				$mNOTE[$tNo] &= "Chat Jabber:" & $gChatJabber[$tNo] & @CRLF
			EndIf
		EndIf
		If $gHomePage[$tNo] <> "" Then $mNOTE[$tNo] &= "Home Page:" & $gHomePage[$tNo] & @CRLF
		If $gPhoneOther[$tNo] <> "" Then $mNOTE[$tNo] &= $gPhoneOther[$tNo] & @CRLF
		If $gURLBlog[$tNo] <> "" Then $mNOTE[$tNo] &= "Blog:" & $gURLBlog[$tNo] & @CRLF
		If $gURLProfile[$tNo] <> "" Then $mNOTE[$tNo] &= "Profile:" & $gURLProfile[$tNo] & @CRLF
		If $gRelatedLabel[$tNo] <> "" Then $mNOTE[$tNo] &= $gRelatedLabel[$tNo] & ": " & $gRelatedName[$tNo] & @CRLF
		If $gCategories[$tNo] <> "" Then
			Local $tSplit = StringSplit($gCategories[$tNo], ",", 2)
			For $t = 0 To (UBound($tSplit) - 1)
				$tSplit[$t] = _ConvertCustomLabel($tSplit[$t])
				$tSplit[$t] = _RemoveSpecialChars($tSplit[$t])
			Next
			$gCategories[$tNo] = _ArrayToString($tSplit, ",")
			If $aAddGroupsToNotesYN = "yes" Then $mNOTE[$tNo] &= "Categories:" & $gCategories[$tNo] & @CRLF
		EndIf
		$mTEL_WORK_FAX[$tNo] = $gPhoneWorkFax[$tNo]
		If $gAddressHomeStreet2[$tNo] <> "" Then _ReplaceChar($tAdd2Home = ", " & $gAddressHomeStreet2[$tNo])
		If $gAddressHomePOBox[$tNo] <> "" Then _ReplaceChar($tAddPOHome = $gAddressHomePOBox[$tNo] & ", ")
		$mADR_HOME[$tNo] = _ReplaceChar(";;" & $tAddPOHome & $gAddressHomeStreet1[$tNo] & $tAdd2Home & ";" & $gAddressHomeCity[$tNo] & ";" & _
				$gAddressHomeState[$tNo] & ";" & $gAddressHomeZIP[$tNo] & ";" & $gAddressHomeCountry[$tNo])
		If $gAddressWorkStreet2[$tNo] <> "" Then _ReplaceChar($tAdd2Work = ", " & $gAddressWorkStreet2[$tNo])
		If $gAddressWorkPOBox[$tNo] <> "" Then $tAddPOWork = $gAddressWorkPOBox[$tNo] & ", "
		$mADR_WORK[$tNo] = _ReplaceChar(";;" & $tAddPOWork & $gAddressWorkStreet1[$tNo] & $tAdd2Work & ";" & $gAddressWorkCity[$tNo] & ";" & _
				$gAddressWorkState[$tNo] & ";" & $gAddressWorkZIP[$tNo] & ";" & $gAddressWorkCountry[$tNo])
		If $gAddressOtherStreet2[$tNo] <> "" Then $tAdd2Other = ", " & $gAddressOtherStreet2[$tNo]
		If $gAddressOtherPOBox[$tNo] <> "" Then $tAddPOOther = $gAddressOtherPOBox[$tNo] & ", "
		$mADR_OTHER[$tNo] = _ReplaceChar(";;" & $tAddPOOther & $gAddressOtherStreet1[$tNo] & $tAdd2Other & ";" & $gAddressOtherCity[$tNo] & ";" & _
				$gAddressOtherState[$tNo] & ";" & $gAddressOtherZIP[$tNo] & ";" & $gAddressOtherCountry[$tNo])
		If $gBirthday[$tNo] <> "" Then $mBDAY[$tNo] = StringLeft($gBirthday[$tNo], 4) & "-" & StringMid($gBirthday[$tNo], 5, 2) & "-" & StringMid($gBirthday[$tNo], 7, 2) & "T00:00:00"
		$mORG[$tNo] = _ReplaceChar($gCompany[$tNo] & ";" & $gDepartment[$tNo])
		$mURL_WORK[$tNo] = StringReplace($gURLWork[$tNo], "\:", ":")
		$mTITLE[$tNo] = _ReplaceChar($gJobTitle[$tNo])
		$mNOTE[$tNo] &= $gCustom[$tNo]
		$mNOTE[$tNo] = StringReplace($mNOTE[$tNo], "http\:", "http:")
		$mNOTE[$tNo] = _ReplaceChar($mNOTE[$tNo])
		$mFileName[$tNo] = _GenerateFileName() & ".VCF"
		If StringLen(_RemoveSpecialChars($mADR_HOME[$tNo])) < 5 Then $mADR_HOME[$tNo] = ""
		If StringLen(_RemoveSpecialChars($mADR_OTHER[$tNo])) < 5 Then $mADR_OTHER[$tNo] = ""
		If StringLen(_RemoveSpecialChars($mADR_WORK[$tNo])) < 5 Then $mADR_WORK[$tNo] = ""
	Next
EndFunc   ;==>_ConvertGoogleToMailEnable
Func _GenerateFileName()
	Local $tFileName = ""
	For $t = 1 To 32
		Local $tRandom = Random(0, 16)
		If $tRandom < 10 Then
			$tFileName &= Chr(48 + $tRandom)
		Else
			$tFileName &= Chr(55 + $tRandom)
		EndIf
	Next
	Return $tFileName
EndFunc   ;==>_GenerateFileName
Func _ReplaceChar($tTxt)
	Local $tReturn = StringReplace($tTxt, "&", "&amp;")
	Return $tReturn
EndFunc   ;==>_ReplaceChar
Func _ConvertPhoneNumbers($tNum)
	If $mTEL_CELL_VOICE[$tNum] <> "" Then $mTEL_CELL_VOICE[$tNum] = _FormatPhoneNumber($mTEL_CELL_VOICE[$tNum])
	If $mTEL_HOME_VOICE[$tNum] <> "" Then $mTEL_HOME_VOICE[$tNum] = _FormatPhoneNumber($mTEL_HOME_VOICE[$tNum])
	If $mTEL_PAGER_VOICE[$tNum] <> "" Then $mTEL_PAGER_VOICE[$tNum] = _FormatPhoneNumber($mTEL_PAGER_VOICE[$tNum])
	If $mTEL_WORK_FAX[$tNum] <> "" Then $mTEL_WORK_FAX[$tNum] = _FormatPhoneNumber($mTEL_WORK_FAX[$tNum])
	If $mTEL_WORK_VOICE[$tNum] <> "" Then $mTEL_WORK_VOICE[$tNum] = _FormatPhoneNumber($mTEL_WORK_VOICE[$tNum])
EndFunc   ;==>_ConvertPhoneNumbers
Func _WriteMailEnable()
	For $tNo = 0 To ($aNumberOfContacts - 1)
		Local $tArray[0]
		ControlSetText($aSplash, "", "Static1", "Writing Mail Enable files." & @CRLF & "Processing contact " & $tNo & " of " & $aNumberOfContacts & _
				@CRLF & @CRLF & "Timer:" & Round(_Timer_Diff($aTimerStart) / 1000) & "s")
		_ArrayAdd($tArray, "BEGIN:VCARD")
		_ArrayAdd($tArray, "VERSION:2.1")
		If StringReplace($mADR_HOME[$tNo], ";", "") <> "" Then
			_ArrayAdd($tArray, "ADR;HOME:" & $mADR_HOME[$tNo])
		Else
			_ArrayAdd($tArray, "ADR;HOME:;;;;;;")
		EndIf
		If StringReplace($mADR_OTHER[$tNo], ";", "") <> "" Then
			_ArrayAdd($tArray, "ADR;OTHER:" & $mADR_OTHER[$tNo])
		Else
			_ArrayAdd($tArray, "ADR;OTHER:;;;;;;")
		EndIf
		If StringReplace($mADR_WORK[$tNo], ";", "") <> "" Then
			_ArrayAdd($tArray, "ADR;WORK:" & $mADR_WORK[$tNo])
		Else
			_ArrayAdd($tArray, "ADR;WORK:;;;;;;")
		EndIf
;~ 		If  <> "" Then _ArrayAdd($tArray, "AGENT:" & )
		If $mBDAY[$tNo] <> "" Then _ArrayAdd($tArray, "BDAY:" & $mBDAY[$tNo])
		If $mEMAIL_OTHER_INTERNET[$tNo] <> "" Then _ArrayAdd($tArray, "EMAIL;OTHER;INTERNET:" & $mEMAIL_OTHER_INTERNET[$tNo])
		If $mEMAIL_PREF_INTERNET[$tNo] <> "" Then _ArrayAdd($tArray, "EMAIL;PREF;INTERNET:" & $mEMAIL_PREF_INTERNET[$tNo])
		If $mEMAIL_WORK_INTERNET[$tNo] <> "" Then _ArrayAdd($tArray, "EMAIL;WORK;INTERNET:" & $mEMAIL_WORK_INTERNET[$tNo])
		If $mTYPE_MSMESSENGER[$tNo] <> "" Then _ArrayAdd($tArray, "IM;TYPE-MSMESSENGER:" & $mTYPE_MSMESSENGER[$tNo])
		If $mFN[$tNo] <> "" Then _ArrayAdd($tArray, "FN:" & $mFN[$tNo])
		If $mN[$tNo] <> "" Then
			_ArrayAdd($tArray, "N:" & $mN[$tNo])
		Else
			_ArrayAdd($tArray, "N:;;;;")
		EndIf
		If $mNICKNAME[$tNo] <> "" Then _ArrayAdd($tArray, "NICKNAME:" & $mNICKNAME[$tNo])
		If $mORG[$tNo] <> "" Then
			_ArrayAdd($tArray, "ORG:" & $mORG[$tNo])
		Else
			_ArrayAdd($tArray, "ORG:;")
		EndIf
		If $aReformatPhoneNumbersYN = "yes" Then _ConvertPhoneNumbers($tNo)
		If $mTEL_CELL_VOICE[$tNo] <> "" Then _ArrayAdd($tArray, "TEL;CELL;VOICE:" & $mTEL_CELL_VOICE[$tNo])
		If $mTEL_HOME_VOICE[$tNo] <> "" Then _ArrayAdd($tArray, "TEL;HOME;VOICE:" & $mTEL_HOME_VOICE[$tNo])
		If $mTEL_PAGER_VOICE[$tNo] <> "" Then _ArrayAdd($tArray, "TEL;PAGER;VOICE:" & $mTEL_PAGER_VOICE[$tNo])
		If $mTEL_WORK_FAX[$tNo] <> "" Then _ArrayAdd($tArray, "TEL;WORK;FAX:" & $mTEL_WORK_FAX[$tNo])
		If $mTEL_WORK_VOICE[$tNo] <> "" Then _ArrayAdd($tArray, "TEL;WORK;VOICE:" & $mTEL_WORK_VOICE[$tNo])
		If $mTITLE[$tNo] <> "" Then _ArrayAdd($tArray, "TITLE:" & $mTITLE[$tNo])
		If $mURL_WORK[$tNo] <> "" Then _ArrayAdd($tArray, "URL;WORK:" & $mURL_WORK[$tNo])
		Local $tNote = StringReplace($mNOTE[$tNo], @CRLF, "")
		$tNote = _RemoveExtraSpaces($tNote)
		If $tNote <> "" Then
			$mNOTE[$tNo] = StringReplace($mNOTE[$tNo], @CRLF, "=0A")
			$mNOTE[$tNo] = StringReplace($mNOTE[$tNo], "\,", ",")
			$mNOTE[$tNo] = StringReplace($mNOTE[$tNo], "\:", ":")
			_ArrayAdd($tArray, "NOTE;ENCODING=QUOTED-PRINTABLE;CHARSET=UTF-8:" & $mNOTE[$tNo])
		EndIf
		If $aDownloadPhotosYN = "yes" Then
			If $gPhotoURL[$tNo] <> "" Then
				Local $tPhoto = StringRegExpReplace(_ConvertToBase64(_INetGetSource($gPhotoURL[$tNo])), @LF, "")
				If StringLen($tPhoto) > 40 Then _ArrayAdd($tArray, "PHOTO;TYPE=JPEG;ENCODING=BASE64:")
				Local $tStop = False
				Local $tTxt = " "
				Do
					$tTxt = StringLeft($tPhoto, 75)
					_ArrayAdd($tArray, " " & $tTxt)
					If StringLen($tTxt) < 75 Then
						$tStop = True
					Else
						$tPhoto = StringTrimLeft($tPhoto, 75)
					EndIf
				Until $tStop
				_ArrayAdd($tArray, "")
			EndIf
		EndIf
		_ArrayAdd($tArray, "REV:20210201T024607")
		_ArrayAdd($tArray, "END:VCARD")
		_FileWriteFromArray($aExportFolder & "\" & $mFileName[$tNo], $tArray)
	Next
EndFunc   ;==>_WriteMailEnable
Func _UpdateIndex()
	_GetLastUID()
	For $tNo = 0 To ($aNumberOfContacts - 1)
		Local $tELEMNT_ID = ""
		Local $tADR_HOME = ""
		Local $tADR_OTHER = ""
		Local $tADR_WORK = ""
		Local $tAGENT = ""
		Local $tBDAY = ""
		Local $tATTACHMENTS = ""
		Local $tEMAIL_OTHER_INTERNET = ""
		Local $tEMAIL_PREF_INTERNET = ""
		Local $tEMAIL_WORK_INTERNET = ""
		Local $tFN = ""
		Local $tIMPORTANCE = ""
		Local $tIM_TYPE_MSMESSENGER = ""
		Local $tN = ""
		Local $tNICKNAME = ""
		Local $tORG = ""
		Local $tREAD = ""
		Local $tSIZE = ""
		Local $tSTATE = ""
		Local $tTEL_CELL_VOICE = ""
		Local $tTEL_HOME_VOICE = ""
		Local $tTEL_PAGER_VOICE = ""
		Local $tTEL_WORK_VOICE = ""
		Local $tTITLE = ""
		Local $tELEMENT_END = ""
		$aLastUIDDec += 1
		$mUID[$tNo] = $aLastUIDDec
		$tELEMNT_ID = '<ELEMENT ID="' & $mFileName[$tNo] & '" UID="' & $mUID[$tNo] & '">'
		If $mADR_HOME[$tNo] <> "" Then
			$tADR_HOME = '<ADR_HOME>' & $mADR_HOME[$tNo] & '</ADR_HOME>'
		Else
			$tADR_HOME = '<ADR_HOME>;;;;;;</ADR_HOME>'
		EndIf
		If $mADR_OTHER[$tNo] <> "" Then
			$tADR_OTHER = '<ADR_OTHER>' & $mADR_OTHER[$tNo] & '</ADR_OTHER>'
		Else
			$tADR_OTHER = '<ADR_OTHER>;;;;;;</ADR_OTHER>'
		EndIf
		If $mADR_WORK[$tNo] <> "" Then
			$tADR_WORK = '<ADR_WORK>' & $mADR_WORK[$tNo] & '</ADR_WORK>'
		Else
			$tADR_WORK = '<ADR_WORK>;;;;;;</ADR_WORK>'
		EndIf
;~ 		If   <> "" Then $tAGENT = '<AGENT>' &  & '</AGENT>'
		$tATTACHMENTS = '<ATTACHMENTS>0</ATTACHMENTS>'
		If $aCopyWorkorOtherEmailsToPrefYN = "yes" Then
			If $mEMAIL_PREF_INTERNET[$tNo] = "" Then
				If $mEMAIL_WORK_INTERNET[$tNo] <> "" Then
					$mEMAIL_PREF_INTERNET[$tNo] = $mEMAIL_WORK_INTERNET[$tNo]
				ElseIf $mEMAIL_OTHER_INTERNET[$tNo] <> "" Then
					$mEMAIL_PREF_INTERNET[$tNo] = $mEMAIL_OTHER_INTERNET[$tNo]
				EndIf
			EndIf
		EndIf
		If $aCopyPhonesNosToMobileYN = "yes" Then
			If $mTEL_CELL_VOICE[$tNo] = "" Then
				If $mTEL_HOME_VOICE[$tNo] <> "" Then
					$mTEL_CELL_VOICE[$tNo] = $mTEL_HOME_VOICE[$tNo]
				ElseIf $mTEL_WORK_VOICE[$tNo] <> "" Then
					$mTEL_CELL_VOICE[$tNo] = $mTEL_WORK_VOICE[$tNo]
				ElseIf $mTEL_PAGER_VOICE[$tNo] <> "" Then
					$mTEL_CELL_VOICE[$tNo] = $mTEL_PAGER_VOICE[$tNo]
				EndIf
			EndIf
		EndIf
		If $mBDAY[$tNo] <> "" Then $tBDAY = '<BDAY>' & $mBDAY[$tNo] & '</BDAY>'
		If $mEMAIL_OTHER_INTERNET[$tNo] <> "" Then $tEMAIL_OTHER_INTERNET = '<EMAIL_OTHER_INTERNET>' & $mEMAIL_OTHER_INTERNET[$tNo] & '</EMAIL_OTHER_INTERNET>'
		If $mEMAIL_PREF_INTERNET[$tNo] <> "" Then $tEMAIL_PREF_INTERNET = '<EMAIL_PREF_INTERNET>' & $mEMAIL_PREF_INTERNET[$tNo] & '</EMAIL_PREF_INTERNET>'
		If $mEMAIL_WORK_INTERNET[$tNo] <> "" Then $tEMAIL_WORK_INTERNET = '<EMAIL_WORK_INTERNET>' & $mEMAIL_WORK_INTERNET[$tNo] & '</EMAIL_WORK_INTERNET>'
		If $mFN[$tNo] <> "" Then $tFN = '<FN>' & $mFN[$tNo] & '</FN>'
		$tIMPORTANCE = '<IMPORTANCE>0</IMPORTANCE>'
		If $mTYPE_MSMESSENGER[$tNo] <> "" Then $tIM_TYPE_MSMESSENGER = '<IM_TYPE-MSMESSENGER>' & $mTYPE_MSMESSENGER[$tNo] & '</IM_TYPE-MSMESSENGER>'
		Local $tDate = _DateDiff('s', "1977/05/14 00:00:00", _NowCalc())
		$tMODIFIEDDATE = '<MODIFIEDDATE>' & $tDate & '</MODIFIEDDATE>'
		If $mN[$tNo] <> "" Then
			$tN = '<N>' & $mN[$tNo] & '</N>'
		Else
			$tN = '<N>;;;;</N>'
		EndIf

		If $mNICKNAME[$tNo] <> "" Then $tNICKNAME = '<NICKNAME>' & $mNICKNAME[$tNo] & '</NICKNAME>'
		If $mORG[$tNo] <> "" Then
			$tORG = '<ORG>' & $mORG[$tNo] & '</ORG>'
		Else
			$tORG = '<ORG>;</ORG>'
		EndIf
		$tREAD = '<READ>0</READ>'
		$tSIZE = '<SIZE>0</SIZE>'
		$tSTATE = '<STATE>0</STATE>'
		If $mTEL_CELL_VOICE[$tNo] <> "" Then $tTEL_CELL_VOICE = '<TEL_CELL_VOICE>' & $mTEL_CELL_VOICE[$tNo] & '</TEL_CELL_VOICE>'
		If $mTEL_HOME_VOICE[$tNo] <> "" Then $tTEL_HOME_VOICE = '<TEL_HOME_VOICE>' & $mTEL_HOME_VOICE[$tNo] & '</TEL_HOME_VOICE>'
		If $mTEL_PAGER_VOICE[$tNo] <> "" Then $tTEL_PAGER_VOICE = '<TEL_PAGER_VOICE>' & $mTEL_PAGER_VOICE[$tNo] & '</TEL_PAGER_VOICE>'
		If $mTEL_WORK_VOICE[$tNo] <> "" Then $tTEL_WORK_VOICE = '<TEL_WORK_VOICE>' & $mTEL_WORK_VOICE[$tNo] & '</TEL_WORK_VOICE>'
		If $mTITLE[$tNo] <> "" Then $tTITLE = '<TITLE>' & $mTITLE[$tNo] & '</TITLE>'
		$tELEMENT_END = '</ELEMENT>'
		$mIndex[$tNo] = $tELEMNT_ID & $tADR_HOME & $tADR_OTHER & $tADR_WORK & $tAGENT & $tATTACHMENTS & $tBDAY & $tEMAIL_OTHER_INTERNET & $tEMAIL_PREF_INTERNET & _
				$tEMAIL_WORK_INTERNET & $tFN & $tIMPORTANCE & $tMODIFIEDDATE & $tIM_TYPE_MSMESSENGER & $tN & $tNICKNAME & $tORG & $tREAD & $tSIZE & $tSTATE & _
				$tTEL_CELL_VOICE & $tTEL_HOME_VOICE & $tTEL_PAGER_VOICE & $tTEL_WORK_VOICE & $tTITLE & $tELEMENT_END
	Next
	Local $tFileRead = FileRead($aIndex_XML, 100000000)
	If StringInStr($tFileRead, "<BASEELEMENT LASTUID=") Then
	Else
		$tFileRead = '<BASEELEMENT LASTUID="x00000000"></BASEELEMENT>'
	EndIf
	Local $tAdd = ""
	For $tNo = 0 To ($aNumberOfContacts - 1)
		$tAdd &= $mIndex[$tNo]
	Next
	$tFileRead = StringReplace($tFileRead, '</BASEELEMENT>', $tAdd & '</BASEELEMENT>')
	Local $tUpdatedLastUID = StringLower(Hex($aLastUIDDec, 8))
	$tFileRead = StringRegExpReplace($tFileRead, '<BASEELEMENT LASTUID="x........">', '<BASEELEMENT LASTUID="x' & $tUpdatedLastUID & '">')
	If FileExists($aIndex_XML) Then FileMove($aIndex_XML, $aExportFolder & "\" & _FileNameRemoveExt($aIndex_XML) & "_Backup_" & @YEAR & "-" & @MON & "-" & @MDAY & ".xml", 1)
	FileWrite($aIndex_XML, $tFileRead)
EndFunc   ;==>_UpdateIndex
Func _WriteCategories()
	For $tNo = 0 To ($aNumberOfContacts - 1)
		ControlSetText($aSplash, "", "Static1", "Prosessing categories." & @CRLF & "Processing contact " & $tNo & " of " & $aNumberOfContacts & _
				@CRLF & @CRLF & "Timer:" & Round(_Timer_Diff($aTimerStart) / 1000) & "s")
		If $gCategories[$tNo] <> "" Then
			Local $tSplit = StringSplit($gCategories[$tNo], ",", 2)
			For $t = 0 To (UBound($tSplit) - 1)
				$tSplit[$t] = _ConvertCustomLabel($tSplit[$t])
				$tSplit[$t] = _RemoveSpecialChars($tSplit[$t])
				Local $tFound = False
				For $i = 0 To (UBound($gCategNames) - 1)
					If $tSplit[$t] = $gCategNames[$i] Then
						$gCategContacts[$i] &= $mFileName[$tNo] & @CRLF
						$tFound = True
						ExitLoop
					EndIf
				Next
				If $tFound = False Then
					_ArrayAdd($gCategNames, $tSplit[$t])
					_ArrayAdd($gCategContacts, $mFileName[$tNo] & @CRLF)
				EndIf
			Next
		EndIf
	Next
	Global $mGroupsFileName[0]
	Local $tAdd = ""
	Local $tFileRead = FileRead($aGroups_XML, 100000000)
	If StringInStr($tFileRead, "<BASEELEMENT>") Then
	Else
		Local $tFileRead = '<BASEELEMENT>' & @CRLF & '</BASEELEMENT>'
		FileWrite($aGroups_XML, $tFileRead)
	EndIf
	For $tNo = 0 To (UBound($gCategNames) - 1)
		_ArrayAdd($mGroupsFileName, _GenerateFileName() & ".GRP")
		FileWrite($aExportFolder & "\" & $mGroupsFileName[$tNo], $gCategContacts[$tNo])
		$tAdd &= '  <ELEMENT ID="' & $mGroupsFileName[$tNo] & '">' & @CRLF & '    <NAME>' & $gCategNames[$tNo] & '</NAME>' & @CRLF & '  </ELEMENT>' & @CRLF
	Next
	$tFileRead = StringReplace($tFileRead, '</BASEELEMENT>', $tAdd & '</BASEELEMENT>')
	If FileExists($aGroups_XML) Then FileMove($aGroups_XML, $aExportFolder & "\" & _FileNameRemoveExt($aGroups_XML) & "_Backup_" & @YEAR & "-" & @MON & "-" & @MDAY & ".XML")
	FileWrite($aGroups_XML, $tFileRead)
EndFunc   ;==>_WriteCategories
Func _FileNameRemoveExt($tPath)
	Local $xFileName = _PathSplit($tPath, "", "", "", "")
	If UBound($xFileName) > 4 Then
		Return $xFileName[3]
	Else
		Return "[ERROR]"
	EndIf
EndFunc   ;==>_FileNameRemoveExt
Func _RemoveExtraSpaces($tText)  ; Replace more than 1 spaces with only 1
	Return StringRegExpReplace($tText, "(?m)\h{2,}", " ")
EndFunc   ;==>_RemoveExtraSpaces
Func _GetLastUID()
	Local $tFileRead = FileRead($aIndex_XML, 50)
	Global $aLastUIDHex = _ArrayToString(_StringBetween($tFileRead, '<BASEELEMENT LASTUID="x', '">'))
	Global $aLastUIDDec = Dec($aLastUIDHex)
EndFunc   ;==>_GetLastUID
Func _RemoveComma($tTxt)
	Return StringReplace($tTxt, ",", " ")
EndFunc   ;==>_RemoveComma
Func _FormatPhoneNumber($tNumber)
	If StringLen(StringRegExpReplace($tNumber, "[^0-9]", "")) > 7 Then
		Local $tReturn = ""
		For $tX1 = 1 To StringLen($tNumber)
			Local $tDig1 = StringMid($tNumber, $tX1, 1)
			Local $tDig2 = StringRegExpReplace($tDig1, "[^0-9]", "")
			If $tDig2 <> "" Then
				$tReturn &= $tDig2
				If StringLeft($tReturn, 1) = "1" Then $tReturn = StringTrimLeft($tReturn, 1)
				If StringLen($tReturn) >= 10 Then
					$tReturn &= StringRight($tNumber, (StringLen($tNumber) - $tX1))
					ExitLoop
				EndIf
			EndIf
		Next
	Else
		Return $tNumber
	EndIf
	If StringLen($tReturn) >= 10 Then
		$tReturn = "(" & StringMid($tReturn, 1, 3) & ") " & StringMid($tReturn, 4, 3) & "-" & StringMid($tReturn, 7, 4) & " " & StringRight($tReturn, (StringLen($tNumber) - 15))
		$tReturn = StringReplace($tReturn, "  ", " ")
	ElseIf StringLen($tReturn) = 7 Then
		$tReturn = StringMid($tReturn, 1, 3) & "-" & StringMid($tReturn, 4, 4)
	Else
	EndIf
	Return $tReturn
EndFunc   ;==>_FormatPhoneNumber
Func _Splash($tTxt, $tTime = 0, $tWidth = 400, $tHeight = 125)
	Local $tPID = SplashTextOn($aUtilName, $tTxt, $tWidth, $tHeight, -1, -1, $DLG_MOVEABLE, "")
	If $tTime > 0 Then
		Sleep($tTime)
		SplashOff()
	EndIf
	Return $tPID
EndFunc   ;==>_Splash
Func _RemoveTrailingSlash($aString)
	Local $bString = StringRight($aString, 1)
	If $bString = "\" Then $aString = StringTrimRight($aString, 1)
	Return $aString
EndFunc   ;==>_RemoveTrailingSlash
Func _ConvertCustomLabel($tTxt)
	If StringIsUpper(StringLeft($tTxt, 1)) Then
		Return $tTxt
	Else
		Local $tLen = StringLen($tTxt)
		Local $tChar = StringMid($tTxt, 1, 1)
		Local $tNew = StringUpper($tChar)
		For $t = 2 To $tLen
			Local $tChar = StringMid($tTxt, $t, 1)
			If StringIsUpper($tChar) Then
				$tNew &= " " & StringUpper($tChar)
			Else
				$tNew &= $tChar
			EndIf
		Next
		Return StringReplace($tNew, "  ", " ")
	EndIf
EndFunc   ;==>_ConvertCustomLabel
Func _RemoveSpecialChars($aString)
	Return StringRegExpReplace($aString, "(?i)([^a-z0-9-_ .])", "")
EndFunc   ;==>_RemoveSpecialChars
Func _ConvertToBase64($tDat1)
	$tDat = StringToBinary($tDat1)
	$objXML = ObjCreate("MSXML2.DOMDocument")
	$objNode = $objXML.createElement("b64")
	$objNode.dataType = "bin.base64"
	$objNode.nodeTypedValue = $tDat
	Return $objNode.Text
EndFunc   ;==>_ConvertToBase64
Func UtilUpdate($tLink, $tDL, $tUtil, $tUtilName, $tSplash = 0, $tUpdate = "show")
	Local $tUtilUpdateAvailableTF = False
	$aUpdateAutoUtil = False
	If $tUpdate = "show" Then _Splash("Checking for update", 0, 350, 75)
	Local $tVer[2]
	Local $sFilePath = $aFolderTemp & $aUtilName & "_latest_ver.tmp"
	$iGet = _InetGetMulti(20, $sFilePath, $tLink)
	If $iGet = "Error" Then
		If $tUpdate = "show" Then _Splash("Update check failed." & @CRLF & "Please try again later.", 2000, 350, 75)
	Else
		$tVer = StringSplit($iGet, "^", 2)
		If UBound($tVer) < 2 Then Return False
		Local $tTxt1 = ReplaceCRLF(ReplaceCRwithCRLF($tVer[1]))
		If $tVer[0] = $tUtil Then
			If $tUpdate = "show" Then _Splash($aUtilName & " up to date.", 2000, 350, 75)
		Else
			$tUtilUpdateAvailableTF = True
			If ($tUpdate = "show") Or ($tUpdate = "auto") Or ($tUpdate = "UpdateOnly") Then
				SplashOff()
				If ($tUpdate = "Auto") And ($aUpdateAutoUtil = "yes") Then
					Local $tMB = 6
				Else
					Local $tMB = MsgBox($MB_YESNOCANCEL, $aUtilVer, "New " & $aUtilName & " update available. " & @CRLF & "Installed version: " & $tUtil & @CRLF & _
							"Latest version: " & $tVer[0] & @CRLF & @CRLF & _
							"Notes: " & @CRLF & $tVer[1] & @CRLF & @CRLF & _
							"Click (YES) to download update to " & @CRLF & @ScriptDir & @CRLF & _
							"Click (NO) to stop checking for updates." & @CRLF & _
							"Click (CANCEL) to skip this update check.", 15)
				EndIf
				If $tMB = 6 Then
					_Splash(" Downloading latest version of " & @CRLF & $tUtilName)
					Local $tZIP = @ScriptDir & "\" & $tUtilName & "_" & $tVer[0] & ".zip"
					If FileExists($tZIP) Then FileDelete($tZIP)
					If FileExists($tUtilName & "_" & $tVer[0] & ".exe") Then FileDelete($tUtilName & "_" & $tVer[0] & ".exe")
					If FileExists(@ScriptDir & "\readme.txt") Then FileDelete(@ScriptDir & "\readme.txt")
					InetGet($tDL, $tZIP, 1)
					_ExtractZipAll($tZIP, @ScriptDir)
					If Not FileExists(@ScriptDir & "\" & $tUtilName & "_" & $tVer[0] & ".exe") Then
						SplashOff()
						$tMB = MsgBox($MB_OKCANCEL, $aUtilVer, "Utility update download failed . . . " & @CRLF & _
								"Go to """ & $tLink & """ to download latest version." & @CRLF & @CRLF & _
								"Click (OK), (CANCEL), or wait 60 seconds, to resume current version.", 60)
					Else
						SplashOff()
						If ($tUpdate = "Auto") And ($aUpdateAutoUtil = "yes") Then
							$tMB = MsgBox($MB_OKCANCEL, $aUtilVer, "Auto utility update download complete. . . " & @CRLF & @CRLF & _
									"Click (OK) to run new version or wait 60 seconds (servers will remain running) OR" & @CRLF & _
									"Click (CANCEL) to resume current version.", 60)
							If $tMB = 1 Then ; OK
							ElseIf $tMB = -1 Then
								$tMB = 1 ; OK
							ElseIf $tMB = 2 Then ; CANCEL
							EndIf
						Else
							$tMB = MsgBox($MB_OKCANCEL, $aUtilVer, "Utility update download complete. . . " & @CRLF & @CRLF & _
									"Click (OK) to run new version (servers will remain running) OR" & @CRLF & _
									"Click (CANCEL), or wait 15 seconds, to resume current version.", 15)
						EndIf
						If $tMB = 1 Then
							Local $xArray[13]
							$xArray[0] = '@echo off'
							$xArray[1] = 'echo --------------------------------------------'
							$xArray[2] = 'echo  Waiting 5 seconds for shutdown to complete'
							$xArray[3] = 'echo --------------------------------------------'
							$xArray[4] = 'timeout 5'
							$xArray[5] = 'start "Starting ' & $aUtilName & '" "' & @ScriptDir & "\" & $tUtilName & "_" & $tVer[0] & ".exe" & '"'
							$xArray[6] = 'echo --------------------------------------------'
							$xArray[7] = 'echo  ' & $aUtilName & ' started . . .'
							$xArray[8] = 'echo --------------------------------------------'
							$xArray[9] = 'timeout 3'
							$xArray[10] = 'exit'
							Local $tBatFile = $aFolderTemp & $aUtilName & "_Delay_Restart.bat"
							FileDelete($tBatFile)
							_FileWriteFromArray($tBatFile, $xArray)
							If FileExists($tBatFile) Then
								Run($tBatFile)
							Else
								Run(@ScriptDir & "\" & $tUtilName & "_" & $tVer[0] & ".exe")
							EndIf
							Exit
						Else
							_Splash("Utility update check canceled by user." & @CRLF & "Resuming utility . . .", 2000)
						EndIf
					EndIf
				ElseIf $tMB = 7 Then
					$aCheckForUpdatesYN = "no"
					IniWrite($aIniFile, $aINIHeader, "CheckForUpdatesYN", $aCheckForUpdatesYN)
					_Splash("Utility update check disabled." & @CRLF & "To enable update check, change [CheckForUpdatesYN=yes] in the .ini.", 5000, 500)
				ElseIf $tMB = 2 Then
;~ 					_Splash("Utility update check canceled by user." & @CRLF & "Resuming utility . . .", 2000)
				EndIf
			EndIf
		EndIf
	EndIf
	Return $tUtilUpdateAvailableTF
EndFunc   ;==>UtilUpdate
Func _InetGetMulti($tCnt, $tFile, $tLink1, $tLink2 = "0")
	FileDelete($tFile)
	Local $i = 0
	Local $tTmp1 = InetGet($tLink1, $tFile, $INET_FORCERELOAD, $INET_DOWNLOADBACKGROUND)
	Do
		Sleep(100)
		$i += 1
	Until InetGetInfo($tTmp1, $INET_DOWNLOADCOMPLETE) Or $i = $tCnt
	InetClose($tTmp1)
	If $i = $tCnt And $tLink2 <> "0" Then
		$tTmp2 = InetGet($tLink2, $tFile, $INET_FORCERELOAD, $INET_DOWNLOADBACKGROUND)
		Do
			Sleep(100)
			$i += 1
		Until InetGetInfo($tTmp2, $INET_DOWNLOADCOMPLETE) Or $i = $tCnt
		InetClose($tTmp2)
	EndIf
	Local $hFileOpen = FileOpen($tFile, 0)
	Local $hFileRead = FileRead($hFileOpen, 100000000)
	If $hFileOpen = -1 Then
		InetClose($tTmp1)
		Sleep(200)
		FileClose($hFileOpen)
		Local $hFileRead = _INetGetSource($tLink1)
		If @error Then
			If $tLink2 <> "0" Then
				$hFileRead = _INetGetSource($tLink2)
				If @error Then
					Return "Error" ; Error
				Else
					FileClose($hFileOpen)
					FileDelete($tFile)
					FileWrite($tFile, $hFileRead)
				EndIf
			Else
				Return True ; Error
			EndIf
		Else
			FileClose($hFileOpen)
			FileDelete($tFile)
			FileWrite($tFile, $hFileRead)
		EndIf
	Else
		FileClose($hFileOpen)
	EndIf
	Return $hFileRead ; No error
EndFunc   ;==>_InetGetMulti
Func ReplaceCRwithCRLF($sString) ; Initial Regular expression by Melba23 with a new suggestion by Ascend4nt and modified By guinness.
	Return StringRegExpReplace($sString, '(*BSR_ANYCRLF)\R', @CRLF) ; Idea by Ascend4nt
EndFunc   ;==>ReplaceCRwithCRLF
Func ReplaceCRLF($tMsg0)
	Return StringReplace($tMsg0, @CRLF, "|")
EndFunc   ;==>ReplaceCRLF
; #FUNCTION# ;===============================================================================
; https://www.autoitscript.com/forum/topic/101529-sunzippings-zipping/?tab=comments#comment-721988
; Name...........: _ExtractZip
; Description ...: Extracts file/folder from ZIP compressed file
; Syntax.........: _ExtractZip($sZipFile, $sDestinationFolder)
; Parameters ....: $sZipFile - full path to the ZIP file to process
;                  $sDestinationFolder - folder to extract to. Will be created if it does not exsist exist.
; Return values .: Success - Returns 1
;                          - Sets @error to 0
;                  Failure - Returns 0 sets @error:
;                  |1 - Shell Object creation failure
;                  |2 - Destination folder is unavailable
;                  |3 - Structure within ZIP file is wrong
;                  |4 - Specified file/folder to extract not existing
; Author ........: trancexx, modifyed by corgano
;
;==========================================================================================
Func _ExtractZipAll($sZipFile, $sDestinationFolder, $sFolderStructure = "")
	Local $i
	Do
		$i += 1
		$sTempZipFolder = @TempDir & "\Temporary Directory " & $i & " for " & StringRegExpReplace($sZipFile, ".*\\", "")
	Until Not FileExists($sTempZipFolder) ; this folder will be created during extraction

	Local $oShell = ObjCreate("Shell.Application")

	If Not IsObj($oShell) Then
		Return SetError(1, 0, 0) ; highly unlikely but could happen
	EndIf

	Local $oDestinationFolder = $oShell.NameSpace($sDestinationFolder)
	If Not IsObj($oDestinationFolder) Then
		DirCreate($sDestinationFolder)
;~         Return SetError(2, 0, 0) ; unavailable destionation location
	EndIf

	Local $oOriginFolder = $oShell.NameSpace($sZipFile & "\" & $sFolderStructure) ; FolderStructure is overstatement because of the available depth
	If Not IsObj($oOriginFolder) Then
		Return SetError(3, 0, 0) ; unavailable location
	EndIf

	Local $oOriginFile = $oOriginFolder.Items() ;get all items
	If Not IsObj($oOriginFile) Then
		Return SetError(4, 0, 0) ; no such file in ZIP file
	EndIf

	; copy content of origin to destination
	$oDestinationFolder.CopyHere($oOriginFile, 20) ; 20 means 4 and 16, replaces files if asked

	DirRemove($sTempZipFolder, 1) ; clean temp dir

	Return 1 ; All OK!

EndFunc   ;==>_ExtractZipAll
Func _ShowGUI()
	#Region ### START Koda GUI section ### Form=K:\AutoIT\_MyProgs\GoogleToMailEnable\Koda\googletome(b7).kxf
	Global $W_MainWindow = GUICreate("Google Contacts to Mail Enable", 902, 578, -1, -1, BitOR($GUI_SS_DEFAULT_GUI, $WS_SIZEBOX, $WS_THICKFRAME))
	GUISetIcon($aIconFile, 99)
	GUISetBkColor(0x808080)
	GUISetOnEvent($GUI_EVENT_CLOSE, "W_MainWindowClose")
	GUISetOnEvent($GUI_EVENT_MINIMIZE, "W_MainWindowMinimize")
	GUISetOnEvent($GUI_EVENT_MAXIMIZE, "W_MainWindowMaximize")
	GUISetOnEvent($GUI_EVENT_RESTORE, "W_MainWindowRestore")
	Global $G_MainWindow = GUICtrlCreateGroup("", 4, -3, 891, 571)
	GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
	Global $G_SourceGroup = GUICtrlCreateGroup("Source: Google vCard", 23, 85, 855, 267)
	GUICtrlSetFont(-1, 9, 400, 0, "MS Sans Serif")
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	Global $I_SourceFile = GUICtrlCreateInput("contacts.vcf", 31, 320, 471, 21)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetOnEvent(-1, "I_SourceFileChange")
	GUICtrlSetTip(-1, "Click to select the source .vcf file")
	Global $L_SourceInstructions = GUICtrlCreateLabel("Export your contacts from Google as vCard (for IOS Contacts)", 34, 104, 443, 23)
	GUICtrlSetFont(-1, 12, 400, 2, "Tahoma")
	GUICtrlSetColor(-1, 0xFFFF00)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetOnEvent(-1, "L_SourceInstructionsClick")
	Global $P_GoogleContactsExportImage = GUICtrlCreatePic($aFolderTemp & "GoogleExport1.jpg", 505, 96, 368, 218)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetTip(-1, "Click to enlarge image")
	GUICtrlSetOnEvent(-1, "P_GoogleContactsExportImageClick")
	Global $L_SelectSource = GUICtrlCreateLabel("Select Source VCF File:", 31, 294, 190, 23)
	GUICtrlSetFont(-1, 12, 800, 2, "Tahoma")
	GUICtrlSetColor(-1, 0xFFFF00)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetOnEvent(-1, "L_SelectSourceClick")
	Global $L_SourceDefault = GUICtrlCreateLabel("(Default: contacts.vcf)", 590, 322, 130, 20)
	GUICtrlSetFont(-1, 10, 400, 0, "Tahoma")
	GUICtrlSetColor(-1, 0x00FF00)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetOnEvent(-1, "L_SourceDefaultClick")
	Global $B_SelectSourceFile = GUICtrlCreateButton("Select File", 503, 318, 83, 25)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetBkColor(-1, $cButtonFadedBlue)
	GUICtrlSetTip(-1, "Click to select the source .vcf file")
	GUICtrlSetOnEvent(-1, "B_SelectSourceFileClick")
	Global $C_ImportPhotos = GUICtrlCreateCheckbox("Import photos", 36, 136, 117, 17)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetTip(-1, "Import Google Contact photos")
	GUICtrlSetOnEvent(-1, "C_ImportPhotosClick")
	Global $C_ReformatPhoneNos = GUICtrlCreateCheckbox("Reformat phone numbers", 36, 159, 145, 17)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetTip(-1, "Convert phone numbers to (111) 222-3333 format. ie. +1 231 453.6754  --> (231) 453-6754")
	GUICtrlSetOnEvent(-1, "C_ReformatPhoneNosClick")
	Global $C_ImportGroups = GUICtrlCreateCheckbox("Import Groups", 36, 182, 117, 17)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetTip(-1, "Import and create Groups")
	GUICtrlSetOnEvent(-1, "C_ImportGroupsClick")
	Global $C_CopyPhone = GUICtrlCreateCheckbox("Copy Phone #s to Preferred (M)", 219, 134, 261, 17)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetTip(-1, "If contact does not have a mobile number, copy Home/Work/Pager to Mobile so that it shows on main screen")
	GUICtrlSetOnEvent(-1, "C_CopyPhoneClick")
	Global $C_CopyEmail = GUICtrlCreateCheckbox("Copy Emails to Preferred", 219, 158, 261, 17)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetTip(-1, "If contact does not have a preferred (main) email, copy Work/Other email to preferred so that it shows on contact page.")
	GUICtrlSetOnEvent(-1, "C_CopyEmailClick")
	Global $L_WorkPhoneAliases = GUICtrlCreateLabel("Work Phone Aliases", 36, 238, 100, 17)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetTip(-1, "(Comma separated) If no work phone, import these custom labels as work phone. ")
	GUICtrlSetOnEvent(-1, "L_WorkPhoneAliasesClick")
	Global $L_HomePhoneAliases = GUICtrlCreateLabel("Home Phone Aliases", 36, 262, 102, 17)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetTip(-1, "(Comma separated) If no home phone, import these custom labels as home phone. ie. If you used custom label, Personal, then any email labelled as Personal will import as Home")
	GUICtrlSetOnEvent(-1, "L_HomePhoneAliasesClick")
	Global $I_WorkAliases = GUICtrlCreateInput("I_WorkAliases", 144, 235, 359, 21)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetOnEvent(-1, "I_WorkAliasesChange")
	GUICtrlSetTip(-1, "(Comma separated) If no work phone, import these custom labels as work phone. ")
	Global $I_HomeAliases = GUICtrlCreateInput("I_HomeAliases", 144, 259, 359, 21)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetOnEvent(-1, "I_HomeAliasesChange")
	GUICtrlSetTip(-1, "(Comma separated) If no home phone, import these custom labels as home phone. ie. If you used custom label, Personal, then any email labelled as Personal will import as Home")
	Global $L_MobilePhoneAlias = GUICtrlCreateLabel("Mobile Phone Aliases", 36, 213, 105, 17)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetTip(-1, "(Comma separated) If no mobile phone, import these custom labels as mobile phone. ie. If you used custom label, Cell, then any phone labelled as Cell will import as Mobile.")
	GUICtrlSetOnEvent(-1, "L_MobilePhoneAliasClick")
	Global $I_MobileAliases = GUICtrlCreateInput("I_MobileAliases", 144, 211, 359, 21)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetOnEvent(-1, "I_MobileAliasesChange")
	GUICtrlSetTip(-1, "(Comma separated) If no mobile phone, import these custom labels as mobile phone. ie. If you used custom label, Cell, then any phone labelled as Cell will import as Mobile.")
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	Global $P_MailEnableLogo = GUICtrlCreatePic($aFolderTemp & "MailEnableLogo1.jpg", 23, 21, 173, 58, BitOR($GUI_SS_DEFAULT_PIC, $WS_BORDER))
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetTip(-1, "Click to visit https://www.mailenable.com")
	GUICtrlSetOnEvent(-1, "P_MailEnableLogoClick")
	Global $L_Title = GUICtrlCreateLabel("Google Contacts to Mail Enable", 228, 31, 484, 41)
	GUICtrlSetFont(-1, 24, 800, 0, "MS Sans Serif")
	GUICtrlSetColor(-1, $cSWTitle)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetTip(-1, "Click to open the GitHub source website.")
	GUICtrlSetOnEvent(-1, "L_TitleClick")
	Global $G_ImportGroup = GUICtrlCreateGroup("Import", 22, 491, 253, 61)
	GUICtrlSetFont(-1, 9, 400, 0, "MS Sans Serif")
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	Global $B_Exit = GUICtrlCreateButton("Exit", 203, 512, 58, 29)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetTip(-1, "Click to exit")
	GUICtrlSetBkColor(-1, $cSWButtonStopServer)
	GUICtrlSetOnEvent(-1, "B_ExitClick")
	Global $B_ImportContacts = GUICtrlCreateButton("START (Import Contacts)", 37, 512, 158, 29)
	GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetBkColor(-1, $cSWButtonStartServer)
	GUICtrlSetTip(-1, "Click to start the importation process.")
	GUICtrlSetOnEvent(-1, "B_ImportContactsClick")
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	Global $P_GoogleContactsLogo = GUICtrlCreatePic($aFolderTemp & "GoogleLogo.jpg", 819, 21, 58, 58)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetTip(-1, "Click to open https://contacts.google.com")
	GUICtrlSetOnEvent(-1, "P_GoogleContactsLogoClick")
	Global $G_OutputGroup = GUICtrlCreateGroup("Output: Mail Enable Mailbox Contact Folder", 22, 363, 857, 115)
	GUICtrlSetFont(-1, 9, 400, 0, "MS Sans Serif")
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	Global $I_OutputContactsFolder = GUICtrlCreateInput("Contacts", 31, 417, 475, 21)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetOnEvent(-1, "I_OutputContactsFolderChange")
	GUICtrlSetTip(-1, "Click to select the output folder: typically the {MailBox}/Contacts folder")
	Global $L_SelectContactsFolder = GUICtrlCreateLabel("Select Mail Enable Mailbox Contact Folder:", 33, 392, 352, 23)
	GUICtrlSetFont(-1, 12, 800, 2, "Tahoma")
	GUICtrlSetColor(-1, 0xFFFF00)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetOnEvent(-1, "L_SelectContactsFolderClick")
	Global $L_SelectContactsDefault = GUICtrlCreateLabel("(Default: C:\Program Files (x86)\Mail Enable\Postoffices\example.com\MAILROOT\joe\Contacts)", 31, 442, 555, 20)
	GUICtrlSetFont(-1, 10, 400, 0, "Tahoma")
	GUICtrlSetColor(-1, 0x00FF00)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetOnEvent(-1, "L_SelectContactsDefaultClick")
	Global $B_OutputContactsSelectFolder = GUICtrlCreateButton("Select Folder", 507, 415, 83, 25)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetBkColor(-1, $cButtonFadedBlue)
	GUICtrlSetTip(-1, "Click to select the output folder: typically the {MailBox}/Contacts folder")
	GUICtrlSetOnEvent(-1, "B_OutputContactsSelectFolderClick")
	Global $C_OpenContactsFolder = GUICtrlCreateCheckbox("Open Contacts Folder When Done", 664, 448, 201, 21)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetTip(-1, "Open the Contacts folder in File Explorer when complete.")
	GUICtrlSetOnEvent(-1, "C_OpenContactsFolderClick")
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	Global $L_FooterPhoenixURL = GUICtrlCreateLabel("http://www.Phoenix125.com", 644, 535, 168, 20)
	GUICtrlSetFont(-1, 10, 400, 0, "Tahoma")
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetOnEvent(-1, "L_FooterPhoenixURLClick")
	Global $L_FooterProgName = GUICtrlCreateLabel($aUtilName & " " & $aUtilVer, 581, 507, 231, 20, $SS_RIGHT)
	GUICtrlSetFont(-1, 10, 400, 0, "Tahoma")
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetOnEvent(-1, "L_FooterProgNameClick")
	Global $P_PhoenixLogo = GUICtrlCreatePic($aFolderTemp & "phoenixlogo.jpg", 818, 495, 60, 60)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetTip(-1, "Click to visit http://www.phoenix125.com")
	GUICtrlSetOnEvent(-1, "P_PhoenixLogoClick")
	Global $P_About = GUICtrlCreateIcon($aIconFile, 202, 534, 511, 24, 24)
;~ 	Global $P_About = GUICtrlCreatePic($aFolderTemp & "phoenixlogo.jpg", 534, 511, 24, 24)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetTip(-1, "About")
	GUICtrlSetOnEvent(-1, "P_AboutClick")
	Global $P_Discord = GUICtrlCreateIcon($aIconFile, 201, 499, 511, 24, 24)
;~ 	Global $P_Discord = GUICtrlCreatePic($aFolderTemp & "phoenixlogo.jpg", 499, 511, 24, 24)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetTip(-1, "Discord Page")
	GUICtrlSetOnEvent(-1, "P_DiscordClick")
	Global $P_Discussion = GUICtrlCreateIcon($aIconFile, 203, 463, 511, 24, 24)
;~ 	Global $P_Discussion = GUICtrlCreatePic($aFolderTemp & "phoenixlogo.jpg", 463, 511, 24, 24)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetTip(-1, "Discussion Forum")
	GUICtrlSetOnEvent(-1, "P_DiscussionClick")
	Global $P_Help = GUICtrlCreateIcon($aIconFile, 204, 425, 511, 24, 24)
;~ 	Global $P_Help = GUICtrlCreatePic($aFolderTemp & "phoenixlogo.jpg", 425, 511, 24, 24)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetTip(-1, "Online Help")
	GUICtrlSetOnEvent(-1, "P_HelpClick")
	Global $P_Update = GUICtrlCreateIcon($aIconFile, 205, 359, 511, 24, 24)
;~ 	Global $P_Update = GUICtrlCreatePic($aFolderTemp & "phoenixlogo.jpg", 359, 511, 24, 24)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetTip(-1, "Check for updates")
	GUICtrlSetOnEvent(-1, "P_UpdateClick")
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###
	_UpdateGUI()
	If $aCheckForUpdatesYN = "yes" Then
		$aUpdateAvailableTF = UtilUpdate($aServerUpdateLinkVer, $aServerUpdateLinkDL, $aUtilVer, "GoogleContactsToMailEnable", 0, "UpdateOnly")
		If $aUpdateAvailableTF Then
			GUICtrlSetImage($P_Update, $aIconFile, 206)
		Else
			GUICtrlSetImage($P_Update, $aIconFile, 205)
		EndIf
	EndIf
EndFunc   ;==>_ShowGUI
Func _UpdateGUI()
	GUICtrlSetData($I_OutputContactsFolder, $aExportFolder)
	GUICtrlSetData($I_SourceFile, $gGoogleSourceFile)
	GUICtrlSetData($I_MobileAliases, $aMobileAliases)
	GUICtrlSetData($I_WorkAliases, $aWorkAliases)
	GUICtrlSetData($I_HomeAliases, $aHomeAliases)
	If $aDownloadPhotosYN = "yes" Then
		GUICtrlSetState($C_ImportPhotos, $GUI_CHECKED)
	Else
		GUICtrlSetState($C_ImportPhotos, $GUI_UNCHECKED)
	EndIf
	If $aCopyPhonesNosToMobileYN = "yes" Then
		GUICtrlSetState($C_CopyPhone, $GUI_CHECKED)
	Else
		GUICtrlSetState($C_CopyPhone, $GUI_UNCHECKED)
	EndIf
	If $aCopyWorkorOtherEmailsToPrefYN = "yes" Then
		GUICtrlSetState($C_CopyEmail, $GUI_CHECKED)
	Else
		GUICtrlSetState($C_CopyEmail, $GUI_UNCHECKED)
	EndIf
	If $aReformatPhoneNumbersYN = "yes" Then
		GUICtrlSetState($C_ReformatPhoneNos, $GUI_CHECKED)
	Else
		GUICtrlSetState($C_ReformatPhoneNos, $GUI_UNCHECKED)
	EndIf
	If $aCreateGroupsYN = "yes" Then
		GUICtrlSetState($C_ImportGroups, $GUI_CHECKED)
	Else
		GUICtrlSetState($C_ImportGroups, $GUI_UNCHECKED)
	EndIf
	If $aOpenFolderWhenDoneYN = "yes" Then
		GUICtrlSetState($C_OpenContactsFolder, $GUI_CHECKED)
	Else
		GUICtrlSetState($C_OpenContactsFolder, $GUI_UNCHECKED)
	EndIf
	If $aUpdateAvailableTF Then
		GUICtrlSetImage($P_Update, $aIconFile, 206)
	Else
		GUICtrlSetImage($P_Update, $aIconFile, 205)
	EndIf
EndFunc   ;==>_UpdateGUI
Func _ReadGUI()
	$aExportFolder = _RemoveTrailingSlash(GUICtrlRead($I_OutputContactsFolder))
	$gGoogleSourceFile = GUICtrlRead($I_SourceFile)
EndFunc   ;==>_ReadGUI
Func B_ExitClick()
	_Exit()
EndFunc   ;==>B_ExitClick
Func B_ImportContactsClick()
	Global $aTimerStart = TimerInit()
	_ReadGUI()
	_IniWrite()
	Global $aStartText = $aUtilName & @CRLF & @CRLF
	Global $aSplash = _Splash($aStartText)
	_ReadGoogleVCAR()
	_SetGlobalVariablesMailEnable($aNumberOfContacts)
	_ConvertGoogleToMailEnable()
	_WriteMailEnable()
	_UpdateIndex()
	If $aCreateGroupsYN = "yes" Then _WriteCategories()
	If $aOpenFolderWhenDoneYN = "yes" Then ShellExecute($aExportFolder)
	SplashOff()
	MsgBox(0, $aUtilName, $aStartText & "Complete. " & $aNumberOfContacts & " contacts imported." & @CRLF & @CRLF & "Timer:" & Round(_Timer_Diff($aTimerStart) / 1000) & "s")
EndFunc   ;==>B_ImportContactsClick
Func B_OutputContactsSelectFolderClick()
	Local $tExportFolder = FileSelectFolder("Please select Mail Enable Contact (Output) Folder.", $aExportFolder)
	If @error Then
		GUICtrlSetData($I_OutputContactsFolder, $aExportFolder)
	Else
		$aExportFolder = _RemoveTrailingSlash($tExportFolder)
	EndIf
	GUICtrlSetData($I_OutputContactsFolder, $aExportFolder)
	IniWrite($aIniFile, $aINIHeader, "ExportFolder", $aExportFolder)
	$aIndex_XML = $aExportFolder & "\_index.xml"
	$aGroups_XML = $aExportFolder & "\_GROUPS.XML"
EndFunc   ;==>B_OutputContactsSelectFolderClick
Func B_SelectSourceFileClick()
	Local $tGoogleSourceFile = FileOpenDialog("Please select source file", $gGoogleSourceFile, "VCF File (*.vcf)", 3, "contacts.vcf")
	If @error Then
		GUICtrlSetData($I_SourceFile, $gGoogleSourceFile)
	Else
		$gGoogleSourceFile = $tGoogleSourceFile
	EndIf
	GUICtrlSetData($I_SourceFile, $gGoogleSourceFile)
	IniWrite($aIniFile, $aINIHeader, "GoogleSourceFile", $gGoogleSourceFile)
EndFunc   ;==>B_SelectSourceFileClick
Func C_CopyEmailClick()
	If GUICtrlRead($C_CopyEmail) = $GUI_CHECKED Then
		$aCopyWorkorOtherEmailsToPrefYN = "yes"
	Else
		$aCopyWorkorOtherEmailsToPrefYN = "no"
	EndIf
	IniWrite($aIniFile, $aINIHeader, "CopyWorkorOtherEmailsToPrefYN", $aCopyWorkorOtherEmailsToPrefYN)
	_UpdateGUI()
EndFunc   ;==>C_CopyEmailClick
Func C_CopyPhoneClick()
	If GUICtrlRead($C_CopyPhone) = $GUI_CHECKED Then
		$aCopyPhonesNosToMobileYN = "yes"
	Else
		$aCopyPhonesNosToMobileYN = "no"
	EndIf
	IniWrite($aIniFile, $aINIHeader, "CopyPhonesNosToMobileYN", $aCopyPhonesNosToMobileYN)
	_UpdateGUI()
EndFunc   ;==>C_CopyPhoneClick
Func C_ImportGroupsClick()
	If GUICtrlRead($C_ImportGroups) = $GUI_CHECKED Then
		$aCreateGroupsYN = "yes"
	Else
		$aCreateGroupsYN = "no"
	EndIf
	IniWrite($aIniFile, $aINIHeader, "CreateGroupsYN", $aCreateGroupsYN)
	_UpdateGUI()
EndFunc   ;==>C_ImportGroupsClick
Func C_ImportPhotosClick()
	If GUICtrlRead($C_ImportPhotos) = $GUI_CHECKED Then
		$aDownloadPhotosYN = "yes"
	Else
		$aDownloadPhotosYN = "no"
	EndIf
	IniWrite($aIniFile, $aINIHeader, "DownloadPhotosYN", $aDownloadPhotosYN)
	_UpdateGUI()
EndFunc   ;==>C_ImportPhotosClick
Func C_OpenContactsFolderClick()
	If GUICtrlRead($C_OpenContactsFolder) = $GUI_CHECKED Then
		$aOpenFolderWhenDoneYN = "yes"
	Else
		$aOpenFolderWhenDoneYN = "no"
	EndIf
	IniWrite($aIniFile, $aINIHeader, "OpenFolderWhenDoneYN", $aOpenFolderWhenDoneYN)
	_UpdateGUI()
EndFunc   ;==>C_OpenContactsFolderClick
Func C_ReformatPhoneNosClick()
	If GUICtrlRead($C_ReformatPhoneNos) = $GUI_CHECKED Then
		$aReformatPhoneNumbersYN = "yes"
	Else
		$aReformatPhoneNumbersYN = "no"
	EndIf
	IniWrite($aIniFile, $aINIHeader, "ReformatPhoneNumbersYN", $aReformatPhoneNumbersYN)
	_UpdateGUI()
EndFunc   ;==>C_ReformatPhoneNosClick
Func I_HomeAliasesChange()
EndFunc   ;==>I_HomeAliasesChange
Func I_MobileAliasesChange()
EndFunc   ;==>I_MobileAliasesChange
Func P_GoogleContactsExportImageClick()
	ShellExecute($aFolderTemp & "GoogleExport1.jpg")
EndFunc   ;==>P_GoogleContactsExportImageClick
Func W_MainWindowClose()
	_Exit()
EndFunc   ;==>W_MainWindowClose
Func W_MainWindowMaximize()
EndFunc   ;==>W_MainWindowMaximize
Func W_MainWindowMinimize()
EndFunc   ;==>W_MainWindowMinimize
Func W_MainWindowRestore()
EndFunc   ;==>W_MainWindowRestore
Func I_OutputContactsFolderChange()
	$aExportFolder = _RemoveTrailingSlash(GUICtrlRead($I_OutputContactsFolder))
	IniWrite($aIniFile, $aINIHeader, "ExportFolder", $aExportFolder)
	$aIndex_XML = $aExportFolder & "\_index.xml"
	$aGroups_XML = $aExportFolder & "\_GROUPS.XML"
EndFunc   ;==>I_OutputContactsFolderChange
Func I_SourceFileChange()
	$gGoogleSourceFile = GUICtrlRead($I_SourceFile)
	IniWrite($aIniFile, $aINIHeader, "GoogleSourceFile", $gGoogleSourceFile)
EndFunc   ;==>I_SourceFileChange
Func I_WorkAliasesChange()
EndFunc   ;==>I_WorkAliasesChange
Func L_FooterPhoenixURLClick()
	ShellExecute("http://www.phoenix125.com")
EndFunc   ;==>L_FooterPhoenixURLClick
Func L_FooterProgNameClick()
	ShellExecute("https://github.com/phoenix125/GoogleContactsToMailEnable")
EndFunc   ;==>L_FooterProgNameClick
Func L_HomePhoneAliasesClick()
EndFunc   ;==>L_HomePhoneAliasesClick
Func L_MobilePhoneAliasClick()
EndFunc   ;==>L_MobilePhoneAliasClick
Func L_SelectContactsDefaultClick()
EndFunc   ;==>L_SelectContactsDefaultClick
Func L_SelectContactsFolderClick()
EndFunc   ;==>L_SelectContactsFolderClick
Func L_SelectSourceClick()
EndFunc   ;==>L_SelectSourceClick
Func L_SourceDefaultClick()
EndFunc   ;==>L_SourceDefaultClick
Func L_SourceInstructionsClick()
EndFunc   ;==>L_SourceInstructionsClick
Func L_TitleClick()
	ShellExecute("https://github.com/phoenix125/GoogleContactsToMailEnable")
EndFunc   ;==>L_TitleClick
Func L_WorkPhoneAliasesClick()
EndFunc   ;==>L_WorkPhoneAliasesClick
Func P_AboutClick()
	MsgBox($MB_SYSTEMMODAL, $aUtilName, $aUtilName & @CRLF & "Version: " & $aUtilVer & @CRLF & @CRLF & "Install Path: " & @ScriptDir & _
			@CRLF & @CRLF & "Discord: http://discord.gg/EU7pzPs" & @CRLF & "Website: http://www.phoenix125.com", 15)
EndFunc   ;==>P_AboutClick
Func P_DiscordClick()
	ShellExecute("http://discord.gg/EU7pzPs")
EndFunc   ;==>P_DiscordClick
Func P_DiscussionClick()
	ShellExecute("https://phoenix125.createaforum.com/google-contacts-to-mail-enable-discussion/")
EndFunc   ;==>P_DiscussionClick
Func P_GoogleContactsLogoClick()
	ShellExecute("https://contacts.google.com")
EndFunc   ;==>P_GoogleContactsLogoClick
Func P_HelpClick()
	ShellExecute("http://www.phoenix125.com/share/mailenable/ReadMe.pdf")
EndFunc   ;==>P_HelpClick
Func P_MailEnableLogoClick()
	ShellExecute("https://www.mailenable.com")
EndFunc   ;==>P_MailEnableLogoClick
Func P_PhoenixLogoClick()
	ShellExecute("http://www.phoenix125.com")
EndFunc   ;==>P_PhoenixLogoClick
Func P_UpdateClick()
	$aUpdateAvailableTF = UtilUpdate($aServerUpdateLinkVer, $aServerUpdateLinkDL, $aUtilVer, "GoogleContactsToMailEnable", 0, "Show")
	If $aUpdateAvailableTF Then
		GUICtrlSetImage($P_Update, $aIconFile, 206)
	Else
		GUICtrlSetImage($P_Update, $aIconFile, 205)
	EndIf
EndFunc   ;==>P_UpdateClick
Func _Exit()
	GUIDelete($W_MainWindow)
	MsgBox(0, $aUtilName, "Thank you for using " & $aUtilName & "." & @CRLF & @CRLF & "Please report any problems or comments to: " & @CRLF & _
			"Discord: http://discord.gg/EU7pzPs or " & @CRLF & "email: kim@phoenix125.com" & @CRLF & @CRLF & "Visit http://www.Phoenix125.com", 20)
	Exit
EndFunc   ;==>_Exit
