; #####################
; # MAIN CONTROL LOOP #
; #####################
Func _ControlMainLoop()
   _LoadSettings()
   _WinMainInit()
   GUISetState(@SW_UNLOCK, $hWinMain)

   $sFromOrg = GUICtrlRead($idFromOrg)
   If Not StringCompare($sFromOrg, "Gemeente") Then
	  GUICtrlSetData($idFromDepartment, $sGemeente, StringLeft($sGemeente, Not StringInStr($sGemeente, "|")? StringLen($sGemeente) : StringInStr($sGemeente, "|") - 1))
	  GUICtrlSetData($idFromDepartment, IniRead($sSettingsFile, "Afzender", "Dienst", ""))
   ElseIf Not StringCompare($sFromOrg, "OCMW") Then
	  GUICtrlSetData($idFromDepartment, $sOCMW, StringLeft($sOCMW, Not StringInStr($sOCMW, "|")? StringLen($sOCMW) : StringInStr($sOCMW, "|") - 1))
	  GUICtrlSetData($idFromDepartment, IniRead($sSettingsFile, "Afzender", "Dienst", ""))
   ElseIf Not StringCompare($sFromOrg, "AGB") Then
	  GUICtrlSetData($idFromDepartment, $sAGB, StringLeft($sAGB, Not StringInStr($sAGB, "|")? StringLen($sAGB) : StringInStr($sAGB, "|") - 1))
	  GUICtrlSetData($idFromDepartment, IniRead($sSettingsFile, "Afzender", "Dienst", ""))
   EndIf
   ;GUICtrlSetData($idFromRef, IniRead($sSettingsFile, "Afkortingen" & GUICtrlRead($idFromOrg), GUICtrlRead($idFromDepartment), "___") & "/" & StringLower(StringCompare(GUICtrlRead($idFromFirstName), "") ? GUICtrlRead($idFromFirstName) : "___") & "/" & "#ID")

   While True
	  $iMsg = GUIGetMsg()
	  If $iMsg Then
		 Switch $iMsg
			Case $idFontSizeNormal
			   $sFontSize = "Normaal"
			   IniWriteSection($sSettingsFile, "Lettergrootte", "    =" & $sFontSize)
			   _WinMainSetFont()
			Case $idFontSizeLarge
			   $sFontSize = "Groot"
			   IniWriteSection($sSettingsFile, "Lettergrootte", "    =" & $sFontSize)
			   _WinMainSetFont()
			Case $idFontSizeXL
			   $sFontSize = "XL"
			   IniWriteSection($sSettingsFile, "Lettergrootte", "    =" & $sFontSize)
			   _WinMainSetFont()
			;Case $idColorStandard
			   ;$sTheme = "Standard"
			   ;IniWrite($sSettingsFile, "View", "Theme", $sTheme)
			   ;_WinMainColor()
			;Case $idColorDark
			   ;$sTheme = "Dark"
			   ;IniWrite($sSettingsFile, "View", "Theme", $sTheme)
			   ;_WinMainColor()
			Case $idAbout
			   _WinAbout()
			Case $idToPostcode
			   If Not StringCompare(GUICtrlRead($idToCity), "") And Not StringCompare(GUICtrlRead($idToCountry), "België") Then
				  $iIndex = _ArraySearch($aPostcodes, GUICtrlRead($idToPostcode), Default, Default, Default, Default, Default, 0)
				  If $iIndex >= 0 Then GUICtrlSetData($idToCity, $aPostcodes[$iIndex][1])
				  EndIf
			Case $idFromOrg
			   $sFromOrg = GUICtrlRead($idFromOrg)
			   GUICtrlSetData($idFromDepartment, "", "")
			   If Not StringCompare($sFromOrg, "Gemeente") Then
				  GUICtrlSetData($idFromDepartment, $sGemeente, StringLeft($sGemeente, Not StringInStr($sGemeente, "|")? StringLen($sGemeente) : StringInStr($sGemeente, "|") - 1))
			   ElseIf Not StringCompare($sFromOrg, "OCMW") Then
				  GUICtrlSetData($idFromDepartment, $sOCMW, StringLeft($sOCMW, Not StringInStr($sOCMW, "|")? StringLen($sOCMW) : StringInStr($sOCMW, "|") - 1))
			   ElseIf Not StringCompare($sFromOrg, "AGB") Then
				  GUICtrlSetData($idFromDepartment, $sAGB, StringLeft($sAGB, Not StringInStr($sAGB, "|")? StringLen($sAGB) : StringInStr($sAGB, "|") - 1))
			   EndIf
			   ;GUICtrlSetData($idFromRef, IniRead($sSettingsFile, "Afkortingen" & GUICtrlRead($idFromOrg), GUICtrlRead($idFromDepartment), "___") & "/" & StringLower(StringCompare(GUICtrlRead($idFromFirstName), "") ? GUICtrlRead($idFromFirstName) : "___") & "/" & "#ID")
			;Case $idFromFirstName
			   ;GUICtrlSetData($idFromRef, IniRead($sSettingsFile, "Afkortingen" & GUICtrlRead($idFromOrg), GUICtrlRead($idFromDepartment), "___") & "/" & StringLower(StringCompare(GUICtrlRead($idFromFirstName), "") ? GUICtrlRead($idFromFirstName) : "___") & "/" & "#ID")
			;Case $idFromDepartment
			   ;GUICtrlSetData($idFromRef, IniRead($sSettingsFile, "Afkortingen" & GUICtrlRead($idFromOrg), GUICtrlRead($idFromDepartment), "___") & "/" & StringLower(StringCompare(GUICtrlRead($idFromFirstName), "") ? GUICtrlRead($idFromFirstName) : "___") & "/" & "#ID")
			Case $idFromTel
			   $sFromTel = GUICtrlRead($idFromTel)
			   $iFromTel = StringLen($sFromTel)
			   $sTelNew = ""
			   If $iFromTel > 0 And (StringIsDigit(StringMid($sFromTel, 1, 1)) Or Not StringCompare(StringMid($sFromTel, 1, 1), "+")) Then $sTelNew &= StringMid($sFromTel, 1, 1)
			   For $i = 2 To $iFromTel
				  If StringIsDigit(StringMid($sFromTel, $i, 1)) Then $sTelNew &= StringMid($sFromTel, $i, 1)
			   Next
			   If StringLen($sTelNew) = 9 And _
				  (Not StringCompare(stringmid($sTelNew, 1, 2), "02") Or Not StringCompare(stringmid($sTelNew, 1, 2), "03") Or _
				   Not StringCompare(stringmid($sTelNew, 1, 2), "04") Or Not StringCompare(stringmid($sTelNew, 1, 2), "09")) Then
				  GUICtrlSetData($idFromTel, StringMid($sTelNew, 1, 2) & " " & StringMid($sTelNew, 3, 3) & " " & StringMid($sTelNew, 6, 2) & " " & StringMid($sTelNew, 8, 2))
			   ElseIf StringLen($sTelNew) = 9 Then
				  GUICtrlSetData($idFromTel, StringMid($sTelNew, 1, 3) & " " & StringMid($sTelNew, 4, 2) & " " & StringMid($sTelNew, 6, 2) & " " & StringMid($sTelNew, 8, 2))
			   ElseIf StringLen($sTelNew) = 10 Then
				  GUICtrlSetData($idFromTel, StringMid($sTelNew, 1, 4) & " " & StringMid($sTelNew, 5, 2) & " " & StringMid($sTelNew, 7, 2) & " " & StringMid($sTelNew, 9, 2))
			   Else
				  GUICtrlSetData($idFromTel, $sTelNew)
			   EndIf
			Case $idOptionsExplorer
			   $sWorkingDir = @WorkingDir
			   GUICtrlSetData($idOptionsAnnex, FileOpenDialog("Bijlage kiezen", "::{450D8FBA-AD25-11D0-98A8-0800361B1103}", "PDF (*.pdf)", $FD_FILEMUSTEXIST + $FD_PATHMUSTEXIST, "", $hWinMain))
			   FileChangeDir($sWorkingDir)
			Case $idConfirm
			   _WinMainRead()
			   If _FormCheck() Then
				  ;If Not StringInStr($sSalutions, $sToSalution, $STR_CASESENSE) Then
					 ;$sSalutions &= "|" & $sToSalution
					 ;IniWriteSection($sSettingsFile, "Aansprekingen", "    =" & StringReplace($sSalutions, "|", @LF & "    ="))
				  ;EndIf
				  ;If Not StringCompare($sFromOrg, "Gemeente") And Not StringInStr($sGemeente, $sFromDepartment, $STR_CASESENSE) Then
					 ;$sGemeente &= "|" & $sFromDepartment
					 ;IniWriteSection($sSettingsFile, "Gemeente", "    =" & StringReplace($sGemeente, "|", @LF & "    ="))
				  ;ElseIf Not StringCompare($sFromOrg, "OCMW") And Not StringInStr($sOCMW, $sFromDepartment, $STR_CASESENSE) Then
					 ;$sOCMW &= "|" & $sFromDepartment
					 ;IniWriteSection($sSettingsFile, "OCMW", "    =" & StringReplace($sOCMW, "|", @LF & "    ="))
				  ;ElseIf Not StringCompare($sFromOrg, "AGB") And Not StringInStr($sAGB, $sFromDepartment, $STR_CASESENSE) Then
					 ;$sAGB &= "|" & $sFromDepartment
					 ;IniWriteSection($sSettingsFile, "AGB", "    =" & StringReplace($sAGB, "|", @LF & "    ="))
				  ;EndIf
				  IniWrite($sSettingsFile, "Afzender", "Organisatie", $sFromOrg)
				  IniWrite($sSettingsFile, "Afzender", "Voornaam", $sFromFirstName)
				  IniWrite($sSettingsFile, "Afzender", "Achternaam", $sFromSurname)
				  IniWrite($sSettingsFile, "Afzender", "Dienst", $sFromDepartment)
				  IniWrite($sSettingsFile, "Afzender", "Telefoon", $sFromTel)
				  IniWrite($sSettingsFile, "Afzender", "Email", $sFromEmail)
				  If _LetterAdd() Then _WinMainReset()
			   EndIf
			Case $idLettersSelectAll
			   For $i = 1 To $aLetters[0][0]
				  GUICtrlSetState($aLetters[$i][0], $GUI_CHECKED)
			   Next
			Case $idLettersSelectNone
			   For $i = 1 To $aLetters[0][0]
				  GUICtrlSetState($aLetters[$i][0], $GUI_UNCHECKED)
			   Next
			Case $idListLetters
			   _GUICtrlListView_SortItems($idListLetters, GUICtrlGetState($idListLetters))
			Case $idLettersOpen
			   $iAmount = 0
			   $bContinue = True
			   $bError = False
			   For $i = 1 To $aLetters[0][0]
				  If BitAND(GUICtrlRead($aLetters[$i][0], 1), $GUI_CHECKED) Then $iAmount += 1
			   Next
			   If $iAmount > 3 Then
				  If MsgBox($MB_ICONWARNING + $MB_YESNO + $MB_DEFBUTTON2, "Waarschuwing", "U staat op het punt om " & $iAmount & " brieven tegelijk te openen. Wilt u doorgaan?") <> $IDYES Then $bContinue = False
			   EndIf
			   If $bContinue Then
				  For $i = 1 To $aLetters[0][0]
					 If BitAND(GUICtrlRead($aLetters[$i][0], 1), $GUI_CHECKED) Then _LetterOpen(@WorkingDir & "\" & $sLettersDir & $aLetters[$i][1])
					 If @error Then $bError = True
				  Next
			   EndIf
			   If $bError Then MsgBox(48, "Waarschuwing", "Een of meerdere brieven konden niet worden geopend. Mogelijke oorzaken:" & @CRLF & @CRLF & "- Microsoft Word 2003 of hoger is niet geïnstalleerd of werkt niet correct." & @CRLF & "- De bestanden zijn beschadigd of ontbreken.")
			Case $idLettersRemove
			   If MsgBox($MB_ICONWARNING + $MB_YESNO + $MB_DEFBUTTON2, "Waarschuwing", "Bent u zeker dat u de geselecteerde brieven wilt verwijderen?") = $IDYES Then
				  For $i = 1 To $aLetters[0][0]
					 If BitAND(GUICtrlRead($aLetters[$i][0], 1), $GUI_CHECKED) Then _LetterRemove($aLetters[$i][1])
				  Next
				  _WinMainLoadLetters()
			   EndIf
			Case $idLettersSend
			   If MsgBox($MB_ICONWARNING + $MB_YESNO + $MB_DEFBUTTON2, "Waarschuwing", "U kunt ingeleverde brieven niet meer bewerken. Wilt u doorgaan?") = $IDYES Then
				  _DatabaseSend()
				  _WinMainLoadLetters()
			   EndIf
			Case $idListHistory
			   _GUICtrlListView_SortItems($idListHistory, GUICtrlGetState($idListHistory))
			Case $GUI_EVENT_MAXIMIZE
			   _WinMainDraw()
			   GUISetState(@SW_UNLOCK, $hWinMain)
			Case $GUI_EVENT_RESTORE
			   _WinMainDraw()
			   GUISetState(@SW_UNLOCK, $hWinMain)
			Case $GUI_EVENT_RESIZED
			   _WinMainDraw()
			   GUISetState(@SW_UNLOCK, $hWinMain)
			Case $GUI_EVENT_CLOSE
			   ExitLoop
		 EndSwitch
		 If Not StringCompare($sPrint, "Ja") Then
			Switch $iMsg
			   Case $idDatabaseDate
				  GUICtrlSetData($idDatabaseCount, "Aantal brieven: aan het tellen...")
				  If Not _DatabaseLoadLetters() Then
					 GUICtrlSetData($idDatabaseCount, "Aantal brieven: " & $aDatabaseLetters[0][0])
				  Else
					 GUICtrlSetData($idDatabaseCount, "Aantal brieven: onbekend")
				  EndIf
			   Case $idDatabaseReload
				  GUICtrlSetData($idDatabaseCount, "Aantal brieven: aan het tellen...")
				  If Not _DatabaseLoadLetters() Then
					 GUICtrlSetData($idDatabaseCount, "Aantal brieven: " & $aDatabaseLetters[0][0])
				  Else
					 GUICtrlSetData($idDatabaseCount, "Aantal brieven: onbekend")
				  EndIf
			   Case $idDatabaseSelect
				  $sDatabaseSelect = GUICtrlRead($idDatabaseSelect)
				  If Not StringCompare($sDatabaseSelect, "Niets selecteren") Then
					 For $i = 1 To $aDatabaseLetters[0][0]
						GUICtrlSetState($aDatabaseLetters[$i][0], $GUI_UNCHECKED)
					 Next
				  ElseIf Not StringCompare($sDatabaseSelect, "Alles selecteren") Then
					 For $i = 1 To $aDatabaseLetters[0][0]
						GUICtrlSetState($aDatabaseLetters[$i][0], $GUI_CHECKED)
					 Next
				  ElseIf Not StringCompare($sDatabaseSelect, "Reeds afgedrukt") Then
					 For $i = 1 To $aDatabaseLetters[0][0]
						If Not StringCompare($aDatabaseLettersInfo[$i][17], "Ja") Then
						   GUICtrlSetState($aDatabaseLetters[$i][0], $GUI_CHECKED)
						Else
						   GUICtrlSetState($aDatabaseLetters[$i][0], $GUI_UNCHECKED)
						EndIf
					 Next
				  ElseIf Not StringCompare($sDatabaseSelect, "Nog niet afgedrukt") Then
					 For $i = 1 To $aDatabaseLetters[0][0]
						If Not StringCompare($aDatabaseLettersInfo[$i][17], "Nee") Then
						   GUICtrlSetState($aDatabaseLetters[$i][0], $GUI_CHECKED)
						Else
						   GUICtrlSetState($aDatabaseLetters[$i][0], $GUI_UNCHECKED)
						EndIf
					 Next
				  ElseIf Not StringCompare($sDatabaseSelect, "Reeds gehandtekend door bestuur") Then
					 For $i = 1 To $aDatabaseLetters[0][0]
						If Not StringCompare($aDatabaseLettersInfo[$i][18], "Ja") Then
						   GUICtrlSetState($aDatabaseLetters[$i][0], $GUI_CHECKED)
						Else
						   GUICtrlSetState($aDatabaseLetters[$i][0], $GUI_UNCHECKED)
						EndIf
					 Next
				  ElseIf Not StringCompare($sDatabaseSelect, "Nog handtekening door bestuur vereist") Then
					 For $i = 1 To $aDatabaseLetters[0][0]
						If Not StringCompare($aDatabaseLettersInfo[$i][18], "Nee") Then
						   GUICtrlSetState($aDatabaseLetters[$i][0], $GUI_CHECKED)
						Else
						   GUICtrlSetState($aDatabaseLetters[$i][0], $GUI_UNCHECKED)
						EndIf
					 Next
				  ElseIf Not StringCompare($sDatabaseSelect, "Reeds verzonden") Then
					 For $i = 1 To $aDatabaseLetters[0][0]
						If Not StringCompare($aDatabaseLettersInfo[$i][19], "Ja") Then
						   GUICtrlSetState($aDatabaseLetters[$i][0], $GUI_CHECKED)
						Else
						   GUICtrlSetState($aDatabaseLetters[$i][0], $GUI_UNCHECKED)
						EndIf
					 Next
				  ElseIf Not StringCompare($sDatabaseSelect, "Nog niet verzonden") Then
					 For $i = 1 To $aDatabaseLetters[0][0]
						If Not StringCompare($aDatabaseLettersInfo[$i][19], "Nee") Then
						   GUICtrlSetState($aDatabaseLetters[$i][0], $GUI_CHECKED)
						Else
						   GUICtrlSetState($aDatabaseLetters[$i][0], $GUI_UNCHECKED)
						EndIf
					 Next
				  EndIf
			   Case $idDatabaseLettersCustomize
				  If _WinOrder() Then
					 _WinMainDraw()
					 _DatabaseDrawLetters()
					 GUISetState(@SW_UNLOCK, $hWinMain)
				  EndIf
			   Case $idDatabaseLetters
				  _GUICtrlListView_SortItems($idDatabaseLetters, GUICtrlGetState($idDatabaseLetters))
			   Case $idDatabaseOpen
				  $iAmount = 0
				  $bContinue = True
				  For $i = 1 To $aDatabaseLetters[0][0]
					 If BitAND(GUICtrlRead($aDatabaseLetters[$i][0], 1), $GUI_CHECKED) Then $iAmount += 1
				  Next
				  If $iAmount > 3 Then
					 If MsgBox($MB_ICONWARNING + $MB_YESNO + $MB_DEFBUTTON2, "Waarschuwing", "U staat op het punt om " & $iAmount & " brieven tegelijk te openen. Wilt u doorgaan?") <> $IDYES Then $bContinue = False
				  EndIf
				  If $bContinue Then _DatabaseOpenLetters()
			   Case $idDatabasePrint
				  $sTemp = GUICtrlRead($idDatabaseCount)
				  GUICtrlSetData($idDatabaseCount, "Aan het afdrukken, even geduld...")
				  _Print()
				  GUICtrlSetData($idDatabaseCount, $sTemp)
			   Case $idDatabaseMarkSigned
				  _DatabaseSetSigned()
			   Case $idDatabaseMarkSent
				  _DatabaseSetSent()
			EndSwitch
		 EndIf
		 $bLettersSelected = False
		 $i = 1
		 While $i <= $aLetters[0][0] And Not $bLettersSelected
			;If $iMsg = $aLetters[$i][0] Then
			   ;If BitAND(GUICtrlRead($aLetters[$i][0], 1), $GUI_CHECKED) Then
				  ;GUICtrlSetState($aLetters[$i][0], $GUI_UNCHECKED)
			   ;Else
				  ;GUICtrlSetState($aLetters[$i][0], $GUI_CHECKED)
			   ;EndIf
			;EndIf
			If BitAND(GUICtrlRead($aLetters[$i][0], 1), $GUI_CHECKED) Then $bLettersSelected = True
			$i += 1
		 WEnd
		 If $bLettersSelected Then
			GUICtrlSetState($idLettersOpen, $GUI_ENABLE)
			GUICtrlSetState($idLettersRemove, $GUI_ENABLE)
			GUICtrlSetState($idLettersSend, $GUI_ENABLE)
		 Else
			GUICtrlSetState($idLettersOpen, $GUI_DISABLE)
			GUICtrlSetState($idLettersRemove, $GUI_DISABLE)
			GUICtrlSetState($idLettersSend, $GUI_DISABLE)
		 EndIf
		 If Not StringCompare($sPrint, "Ja") Then
			$bLettersSelected = False
			$i = 1
			While $i <= $aDatabaseLetters[0][0] And Not $bLettersSelected
			   ;If $iMsg = $aLetters[$i][0] Then
				  ;If BitAND(GUICtrlRead($aLetters[$i][0], 1), $GUI_CHECKED) Then
					 ;GUICtrlSetState($aLetters[$i][0], $GUI_UNCHECKED)
				  ;Else
					 ;GUICtrlSetState($aLetters[$i][0], $GUI_CHECKED)
				  ;EndIf
			   ;EndIf
			   If BitAND(GUICtrlRead($aDatabaseLetters[$i][0], 1), $GUI_CHECKED) Then $bLettersSelected = True
			   $i += 1
			WEnd
			If $bLettersSelected Then
			   GUICtrlSetState($idDatabaseOpen, $GUI_ENABLE)
			   GUICtrlSetState($idDatabasePrint, $GUI_ENABLE)
			   GUICtrlSetState($idDatabaseMarkSigned, $GUI_ENABLE)
			   GUICtrlSetState($idDatabaseMarkSent, $GUI_ENABLE)
			Else
			   GUICtrlSetState($idDatabaseOpen, $GUI_DISABLE)
			   GUICtrlSetState($idDatabasePrint, $GUI_DISABLE)
			   GUICtrlSetState($idDatabaseMarkSigned, $GUI_DISABLE)
			   GUICtrlSetState($idDatabaseMarkSent, $GUI_DISABLE)
			EndIf
		 EndIf
	  EndIf
   WEnd

   _WinMainDelete()
EndFunc