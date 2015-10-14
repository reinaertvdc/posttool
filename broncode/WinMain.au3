; ##########################
; # INITIALIZE MAIN WINDOW #
; ##########################
Func _WinMainInit()
   ; ### COLOR ###
   ;Global $iWinBKColor
   ;Global $iControlColor
   ;Global $iControlBKColor
   ;Global $iLabelColor
   ;Global $iLabelBKColor
   ;Global $iGroupColor
   ;Global $iGroupBkColor

   ; ### FONT ###
   Global $iFontSize
   Global $iFontName        = "Gill Sans MT"
   Global $iNormalFontWeight  = 400
   Global $iGroupFontWeight = 560

   ; ### WINDOW ###
   ; Window size
   Global $iWinMainWidth  = 640
   Global $iWinMainHeight = 690

   ; ### GROUPS ###
   ; Horizontal margins.
   Global $iGroupMarginLeft      = 10
   Global $iGroupMarginRight     = $iGroupMarginLeft
   Global $iGroupMarginMiddleHor = $iGroupMarginLeft
   ; Vertical margins.
   Global $iGroupMarginTop       = 30
   Global $iGroupMarginBottom    = 10
   Global $iGroupMarginMiddleVer = 10
   ; Number of groups per line.
   Global $iGroupsPerLine = 2
   ; Group size
   Global $iGroupWidth1
   Global $iGroupHeight1
   Global $iGroupWidth2
   Global $iGroupHeight2

   ; ### LABELS ###
   ; Horizontal margins.
   Global $iLabelMarginLeft  = 12
   Global $iLabelMarginRight = 10
   ; Vertical margins.
   Global $iLabelMarginTop    = 30
   Global $iLabelMarginBottom = $iLabelMarginTop
   ; Label size.
   Global $iLabelWidth = 110
   Global $iLabelHeight

   ; ### GENERAL CONTROLS ###
   ; Horizontal margins.
   Global $iCtrlMarginLeft  = 9
   Global $iCtrlMarginRight = $iCtrlMarginLeft
   Global $iCtrlMarginMiddleHor = $iCtrlMarginLeft
   ; Vertical margins.
   Global $iCtrlMarginTop       = $iLabelMarginTop - 3
   Global $iCtrlMarginBottom    = $iCtrlMarginTop
   Global $iCtrlMarginMiddleVer = 10
   ; Margin between a control and its label.
   Global $iMarginBetweenCtrlAndLabel = ($iLabelMarginRight > $iCtrlMarginLeft) ? $iLabelMarginRight : $iCtrlMarginLeft
   ; Control size.
   Global $iCtrlWidth1
   Global $iCtrlWidth2
   Global $iCtrlWidth3 = 200
   Global $iCtrlHeight
   Global $iLineHeight

   ; ### CONFIRM BUTTON ###
   ; Vertical margins.
   Global $iConfirmMarginTop = 20
   Global $iConfirmMarginBottom = $iConfirmMarginTop
   ; Button size.
   Global $iConfirmWidth = 200
   Global $iConfirmHeight = 35

   ; ### CREATE MAIN WINDOW ###
   _WinMainCreate()
   _WinMainCalibrate()
   _WinMainSetFont()
   ;_WinMainColor()
   _WinMainDraw()
   $aiWinMainSize = WinGetClientSize($hWinMain)
   WinMove($hWinMain, "", Default, (@DesktopHeight - $aiWinMainSize[1]) / 2)
   GUISetState(@SW_SHOW, $hWinMain)
EndFunc





; #########################
; # CALIBRATE MAIN WINDOW #
; #########################
Func _WinMainCalibrate()
   ; Get window size.
   $aiWinMainSize = WinGetClientSize($hWinMain)
   $iWinMainWidth = $aiWinMainSize[0]
   $iWinMainHeight = $aiWinMainSize[1]
   ; Calibrate font.
   ;If Not StringCompare($sTheme, "Donker") Then
	  ;$iWinBKColor     = 0x000000
	  ;$iControlColor   = 0xFFFFFF
	  ;$iControlBKColor = 0x000000
	  ;$iLabelColor     = $iControlColor
	  ;$iLabelBKColor   = $iWinBKColor
	  ;$iGroupColor     = $iLabelBKColor
	  ;$iGroupBkColor   = $iLabelColor
   ;Else
	  ;$iWinBKColor     = 0xF0F0F0
	  ;$iControlColor   = 0x000000
	  ;$iControlBKColor = 0xFFFFFF
	  ;$iLabelColor     = $iControlColor
	  ;$iLabelBKColor   = $iWinBKColor
	  ;$iGroupColor     = $iLabelColor
	  ;$iGroupBkColor   = $iLabelBKColor
   ;EndIf
   ; Calibrate size of font and related controls.
   If Not StringCompare($sFontSize, "Groot") Then
	  $iFontSize = 12
	  $iCtrlHeight = 31
   ElseIf Not StringCompare($sFontSize, "XL") Then
	  $iFontSize = 14
	  $iCtrlHeight = 35
   Else
	  $iFontSize = 11
	  $iCtrlHeight = 29
   EndIf
   $iLineHeight = $iCtrlHeight + $iCtrlMarginMiddleVer
   $iGroupWidth1 = ($iWinMainWidth - $iGroupMarginLeft - $iGroupMarginRight - $iGroupMarginMiddleHor * ($iGroupsPerLine - 1)) / $iGroupsPerLine
   $iGroupHeight1 = $iLineHeight * 12 + $iCtrlMarginTop
   $iGroupWidth2 = $iWinMainWidth - $iGroupMarginLeft - $iGroupMarginRight
   $iGroupHeight2 = $iLineHeight * 3 + $iCtrlMarginTop
   $iCtrlWidth1 = $iGroupWidth1 - $iLabelWidth - $iLabelMarginLeft - $iMarginBetweenCtrlAndLabel - $iCtrlMarginRight
   $iCtrlWidth2 = $iGroupWidth2 - $iCtrlWidth3 - $iLabelWidth - $iLabelMarginLeft * 2 - $iMarginBetweenCtrlAndLabel - $iCtrlMarginRight
EndFunc





; ######################
; # CREATE MAIN WINDOW #
; ######################
Func _WinMainCreate()
   Global $aidGroups[] = [0]
   Global $aidLabels[] = [0]
   Global $aidCheckboxes[] = [0]
   Global $aidControls[] = [0]

   Global $hWinMain = GUICreate("Verwerking uitgaande post", $iWinMainWidth, $iWinMainHeight, Default, Default, $WS_MINIMIZEBOX + $WS_CAPTION + $WS_POPUP + $WS_SYSMENU + $WS_MAXIMIZEBOX + $WS_SIZEBOX)
   _GUIScrollBars_EnableScrollBar($hWinMain, $SB_VERT, $ESB_ENABLE_BOTH)
	  GUISetIcon($sProgramLogo)
	  Opt("GUIResizeMode", $GUI_DOCKBORDERS)
	  Global $idMenuView = GUICtrlCreateMenu("Beeld")
		 Global $idMenuFontSize = GUICtrlCreateMenu("Lettergrootte", $idMenuView)
			Global $idFontSizeNormal = GUICtrlCreateMenuItem("Normaal", $idMenuFontSize, 0, 1)
			If Not StringCompare($sFontSize, "Normaal") Then GUICtrlSetState(Default, $GUI_CHECKED)
			Global $idFontSizeLarge = GUICtrlCreateMenuItem("Groot", $idMenuFontSize, 1, 1)
			If Not StringCompare($sFontSize, "Groot") Then GUICtrlSetState(Default, $GUI_CHECKED)
			Global $idFontSizeXL = GUICtrlCreateMenuItem("XL", $idMenuFontSize, 2, 1)
			If Not StringCompare($sFontSize, "XL") Then GUICtrlSetState(Default, $GUI_CHECKED)
		 ;Global $idMenuColor = GUICtrlCreateMenu("Kleur", $idMenuView)
			;Global $idColorStandard = GUICtrlCreateMenuItem("Standaard", $idMenuColor, 0, 1)
			;If Not StringCompare($sTheme, "Standard") Then GUICtrlSetState(Default, $GUI_CHECKED)
			;Global $idColorDark = GUICtrlCreateMenuItem("Donker", $idMenuColor, 1, 1)
			;If Not StringCompare($sTheme, "Dark") Then GUICtrlSetState(Default, $GUI_CHECKED)
	  Global $idMenuHelp = GUICtrlCreateMenu("Help")
		 Global $idAbout = GUICtrlCreateMenuItem("Over Postcentrale", $idMenuHelp)
	  Global $idTab = GUICtrlCreateTab(0, 0, $iWinMainWidth + 2, $iWinMainHeight - 19)
		 Opt("GUIResizeMode", $GUI_DOCKALL)
		 GUICtrlCreateTabItem("Aanmaken")
			; 1 To
			Global $idGroupTo = GUICtrlCreateGroup("Bestemmeling", Default, Default)
			_ArrayAdd($aidGroups, $idGroupTo)
			   ; 1 Org
			   _ArrayAdd($aidLabels, GUICtrlCreateLabel("Organisatie", Default, Default, Default, Default, $SS_RIGHT))
			   Global $idToOrg = GUICtrlCreateInput("", Default, Default)
			   _ArrayAdd($aidControls, $idToOrg)
			   ; 2 Salutation
			   _ArrayAdd($aidLabels, GUICtrlCreateLabel("Aanspreking", Default, Default, Default, Default, $SS_RIGHT))
			   Global $idToSalution = GUICtrlCreateCombo("", Default, Default, Default, Default, $CBS_DROPDOWNLIST)
			   _ArrayAdd($aidControls, $idToSalution)
			   GUICtrlSetData($idToSalution, $sSalutions, StringLeft($sSalutions, Not StringInStr($sSalutions, "|")? StringLen($sSalutions) : StringInStr($sSalutions, "|") - 1))
			   ; 3 FirstName
			   _ArrayAdd($aidLabels, GUICtrlCreateLabel("Voornaam", Default, Default, Default, Default, $SS_RIGHT))
			   Global $idToFirstName = GUICtrlCreateInput("", Default, Default)
			   _ArrayAdd($aidControls, $idToFirstName)
			   ; 4 Surname
			   _ArrayAdd($aidLabels, GUICtrlCreateLabel("Achternaam", Default, Default, Default, Default, $SS_RIGHT))
			   Global $idToSurname = GUICtrlCreateInput("", Default, Default)
			   _ArrayAdd($aidControls, $idToSurname)
			   ; 5 Street
			   _ArrayAdd($aidLabels, GUICtrlCreateLabel("Straatnaam", Default, Default, Default, Default, $SS_RIGHT))
			   Global $idToStreet = GUICtrlCreateInput("", Default, Default)
			   _ArrayAdd($aidControls, $idToStreet)
			   ; 6 Number
			   _ArrayAdd($aidLabels, GUICtrlCreateLabel("Huisnummer", Default, Default, Default, Default, $SS_RIGHT))
			   Global $idToNumber = GUICtrlCreateInput("", Default, Default)
			   _ArrayAdd($aidControls, $idToNumber)
			   ; 7 Bus
			   _ArrayAdd($aidLabels, GUICtrlCreateLabel("Bus", Default, Default, Default, Default, $SS_RIGHT))
			   Global $idToBus = GUICtrlCreateInput("", Default, Default)
			   _ArrayAdd($aidControls, $idToBus)
			   ; 8 Postcode
			   _ArrayAdd($aidLabels, GUICtrlCreateLabel("Postcode", Default, Default, Default, Default, $SS_RIGHT))
			   Global $idToPostcode = GUICtrlCreateInput("", Default, Default)
			   _ArrayAdd($aidControls, $idToPostcode)
			   ; 9 City
			   _ArrayAdd($aidLabels, GUICtrlCreateLabel("Gemeente", Default, Default, Default, Default, $SS_RIGHT))
			   Global $idToCity = GUICtrlCreateInput("", Default, Default)
			   _ArrayAdd($aidControls, $idToCity)
			   ; 10 Country
			   _ArrayAdd($aidLabels, GUICtrlCreateLabel("Land", Default, Default, Default, Default, $SS_RIGHT))
			   Global $idToCountry = GUICtrlCreateCombo("", Default, Default, Default, Default, $CBS_DROPDOWNLIST)
			   _ArrayAdd($aidControls, $idToCountry)
			   GUICtrlSetData($idToCountry, $sCountries, "België")
			   ; 11 Ref
			   _ArrayAdd($aidLabels, GUICtrlCreateLabel("Uw kenmerk", Default, Default, Default, Default, $SS_RIGHT))
			   Global $idToRef = GUICtrlCreateInput("", Default, Default)
			   _ArrayAdd($aidControls, $idToRef)
			; 2 From
			Global $idGroupContact = GUICtrlCreateGroup("Afzender", Default, Default)
			_ArrayAdd($aidGroups, $idGroupContact)
			   ; 12 Org
			   _ArrayAdd($aidLabels, GUICtrlCreateLabel("Organisatie", Default, Default, Default, Default, $SS_RIGHT))
			   Global $idFromOrg = GUICtrlCreateCombo("", Default, Default, Default, Default, $CBS_DROPDOWNLIST)
			   _ArrayAdd($aidControls, $idFromOrg)
			   GUICtrlSetData($idFromOrg, $sOrgs, IniRead($sSettingsFile, "Afzender", "Organisatie", ""))
			   ; 13 FirstName
			   _ArrayAdd($aidLabels, GUICtrlCreateLabel("Voornaam", Default, Default, Default, Default, $SS_RIGHT))
			   Global $idFromFirstName = GUICtrlCreateInput(IniRead($sSettingsFile, "Afzender", "Voornaam", ""), Default, Default)
			   _ArrayAdd($aidControls, $idFromFirstName)
			   ; 14 Surname
			   _ArrayAdd($aidLabels, GUICtrlCreateLabel("Achternaam", Default, Default, Default, Default, $SS_RIGHT))
			   Global $idFromSurname = GUICtrlCreateInput(IniRead($sSettingsFile, "Afzender", "Achternaam", ""), Default, Default)
			   _ArrayAdd($aidControls, $idFromSurname)
			   ; 15 Department
			   _ArrayAdd($aidLabels, GUICtrlCreateLabel("Dienst", Default, Default, Default, Default, $SS_RIGHT))
			   Global $idFromDepartment = GUICtrlCreateCombo("", Default, Default, Default, Default, $CBS_DROPDOWNLIST)
			   _ArrayAdd($aidControls, $idFromDepartment)
			   ; 16 Tel
			   _ArrayAdd($aidLabels, GUICtrlCreateLabel("Telefoon", Default, Default, Default, Default, $SS_RIGHT))
			   Global $idFromTel = GUICtrlCreateInput(IniRead($sSettingsFile, "Afzender", "Telefoon", ""), Default, Default)
			   _ArrayAdd($aidControls, $idFromTel)
			   ; 17 Email
			   _ArrayAdd($aidLabels, GUICtrlCreateLabel("E-mailadres", Default, Default, Default, Default, $SS_RIGHT))
			   Global $idFromEmail = GUICtrlCreateInput(IniRead($sSettingsFile, "Afzender", "Email", ""), Default, Default)
			   _ArrayAdd($aidControls, $idFromEmail)
			   ; 18 Subject
			   _ArrayAdd($aidLabels, GUICtrlCreateLabel("Onderwerp", Default, Default, Default, Default, $SS_RIGHT))
			   Global $idFromSubject = GUICtrlCreateInput("", Default, Default)
			   _ArrayAdd($aidControls, $idFromSubject)
			   ; 19 Date
			   _ArrayAdd($aidLabels, GUICtrlCreateLabel("Datum", Default, Default, Default, Default, $SS_RIGHT))
			   Global $idFromDate = GUICtrlCreateDate("", Default, Default, Default, Default, $DTS_SHORTDATEFORMAT)
			   _ArrayAdd($aidControls, $idFromDate)
			   ; 20 Ref
			   _ArrayAdd($aidLabels, GUICtrlCreateLabel("Ons kenmerk", Default, Default, Default, Default, $SS_RIGHT))
			   Global $idFromRef = GUICtrlCreateInput("#ID", Default, Default)
			   _ArrayAdd($aidControls, $idFromRef)
			; 3 Options
			Global $idGroupOptions = GUICtrlCreateGroup("Opties", Default, Default)
			_ArrayAdd($aidGroups, $idGroupOptions)
			   ; 4 Registered
			   Global $idOptionsRegistered = GUICtrlCreateCheckbox("Aangetekend verzenden", Default, Default)
			   _ArrayAdd($aidCheckboxes, $idOptionsRegistered)
			   ; 5 Confidential
			   Global $idOptionsConfidential = GUICtrlCreateCheckbox("Vertrouwelijk", Default, Default)
			   _ArrayAdd($aidCheckboxes, $idOptionsConfidential)
			   ; 23 Annex
			   _ArrayAdd($aidLabels, GUICtrlCreateLabel("Bijlage", Default, Default, Default, Default, $SS_RIGHT))
			   Global $idOptionsAnnex = GUICtrlCreateInput("", Default, Default)
			   _ArrayAdd($aidControls, $idOptionsAnnex)
			   Global $idOptionsExplorer = GUICtrlCreateButton("...", Default, Default)
			   ; 24 Ref
			   _ArrayAdd($aidLabels, GUICtrlCreateLabel("Opmerking", Default, Default, Default, Default, $SS_RIGHT))
			   Global $idOptionsRemark = GUICtrlCreateInput("", Default, Default)
			   _ArrayAdd($aidControls, $idOptionsRemark)
			   ; 24 Signed
			   Global $idOptionsSigned = GUICtrlCreateCheckbox("Handtekening bestuur", Default, Default)
			   _ArrayAdd($aidCheckboxes, $idOptionsSigned)
			   GuiCtrlSetState($idOptionsSigned, $GUI_CHECKED)
			; 25 Confirm
			Global $idConfirm = GUICtrlCreateButton("Bevestigen", Default, Default)
			_ArrayAdd($aidControls, $idConfirm)
		 GUICtrlCreateTabItem("Bewerken")
			Global $idLettersSelectAll = GUICtrlCreateButton("Alles selecteren", Default, Default)
			_ArrayAdd($aidControls, $idLettersSelectAll)
			Global $idLettersSelectNone = GUICtrlCreateButton("Niets selecteren", Default, Default)
			_ArrayAdd($aidControls, $idLettersSelectNone)
			Global $idListLetters = GUICtrlCreateListView("Laatst bewerkt|Organisatie|Onderwerp|Bestemmeling", Default, Default, Default, Default, Default, $LVS_EX_FULLROWSELECT + $WS_EX_CLIENTEDGE + $LVS_EX_CHECKBOXES)
			_ArrayAdd($aidControls, $idListLetters)
			_GUICtrlListView_RegisterSortCallBack ($idListLetters)
			_GUICtrlListView_SetColumnWidth($idListLetters, 0, 130)
			_GUICtrlListView_SetColumnWidth($idListLetters, 1, 150)
			_GUICtrlListView_SetColumnWidth($idListLetters, 2, 150)
			_GUICtrlListView_SetColumnWidth($idListLetters, 3, 150)
			Global $aLetters[][3] = [[0, 0, 0]]
			Global $idLettersOpen = GUICtrlCreateButton("Openen", Default, Default)
			_ArrayAdd($aidControls, $idLettersOpen)
			Global $idLettersRemove = GUICtrlCreateButton("Verwijderen", Default, Default)
			_ArrayAdd($aidControls, $idLettersRemove)
			Global $idLettersSend = GUICtrlCreateButton("Inleveren", Default, Default)
			_ArrayAdd($aidControls, $idLettersSend)
			_WinMainLoadLetters()
		 GUICtrlCreateTabItem("Geschiedenis")
			Global $idListHistory = GUICtrlCreateListView("Ingediend|Registratienummer|Organisatie|Onderwerp|Bestemmeling", Default, Default, Default, Default, Default, $LVS_EX_FULLROWSELECT + $WS_EX_CLIENTEDGE)
			_ArrayAdd($aidControls, $idListHistory)
			_GUICtrlListView_RegisterSortCallBack ($idListHistory)
			_GUICtrlListView_SetColumnWidth($idListHistory, 0, 120)
			_GUICtrlListView_SetColumnWidth($idListHistory, 1, 130)
			_GUICtrlListView_SetColumnWidth($idListHistory, 2, 120)
			_GUICtrlListView_SetColumnWidth($idListHistory, 3, 120)
			_WinMainLoadHistory()
		 If Not StringCompare($sPrint, "Ja") Then
			GUICtrlCreateTabItem("Printen")
			   _ArrayAdd($aidLabels, GUICtrlCreateLabel("Datum", Default, Default, Default, Default, $SS_RIGHT))
			   Global $idDatabaseDate = GUICtrlCreateDate("", Default, Default, Default, Default, $DTS_SHORTDATEFORMAT)
			   _ArrayAdd($aidControls, $idDatabaseDate)
			   Global $idDatabaseReload = GUICtrlCreateButton("Vernieuwen", Default, Default)
			   _ArrayAdd($aidControls, $idDatabaseReload)
			   Global $idDatabaseCount = GUICtrlCreateLabel("Aantal brieven: onbekend", Default, Default)
			   _ArrayAdd($aidLabels, $idDatabaseCount)
			   Global $idDatabaseSelect = GUICtrlCreateCombo("", Default, Default, Default, Default, $CBS_DROPDOWNLIST)
			   _ArrayAdd($aidControls, $idDatabaseSelect)
			   GUICtrlSetData($idDatabaseSelect, "Niets selecteren|Alles selecteren|Reeds afgedrukt|Nog niet afgedrukt|Reeds gehandtekend door bestuur|Nog handtekening door bestuur vereist|Reeds verzonden|Nog niet verzonden", "Niets selecteren")
			   Global $idDatabaseLettersCustomize = GUICtrlCreateButton("Kolommen in de lijst aanpassen", Default, Default)
			   _ArrayAdd($aidControls, $idDatabaseLettersCustomize)
			   Global $idDatabaseLetters = GUICtrlCreateListView("", Default, Default, Default, Default, Default, $LVS_EX_FULLROWSELECT + $WS_EX_CLIENTEDGE + $LVS_EX_CHECKBOXES)
			   _GUICtrlListView_RegisterSortCallBack ($idDatabaseLetters)
			   _ArrayAdd($aidControls, $idDatabaseLetters)
			   Global $aDatabaseLetters[][7] = [[0, 0, 0, 0, 0, 0, 0]]
			   Global $aDatabaseLettersInfo[][20] = [["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]]
			   Global $idDatabaseOpen = GUICtrlCreateButton("Openen", Default, Default)
			   _ArrayAdd($aidControls, $idDatabaseOpen)
			   Global $idDatabasePrint = GUICtrlCreateButton("Printen", Default, Default)
			   _ArrayAdd($aidControls, $idDatabasePrint)
			   Global $idDatabaseMarkSigned = GUICtrlCreateButton("Is ondertekend", Default, Default)
			   _ArrayAdd($aidControls, $idDatabaseMarkSigned)
			   Global $idDatabaseMarkSent = GUICtrlCreateButton("Is verzonden", Default, Default)
			   _ArrayAdd($aidControls, $idDatabaseMarkSent)
		 EndIf

   $aidGroups[0] = UBound($aidGroups) - 1
   $aidLabels[0] = UBound($aidLabels) - 1
   $aidCheckboxes[0] = UBound($aidCheckboxes) - 1
   $aidControls[0] = UBound($aidControls) - 1
EndFunc





; ########################
; # SET MAIN WINDOW FONT #
; ########################
Func _WinMainSetFont()
   _WinMainDraw()
   GUISetState(@SW_LOCK, $hWinMain)
   $aiWinMainSize = WinGetClientSize($hWinMain)
   $aiWinMainSizeBorder = WinGetPos($hWinMain)
   $iWinMainHeightNew = $iGroupMarginTop + $iGroupHeight1 + $iGroupHeight2 + $iGroupMarginMiddleVer + $iConfirmMarginTop + $iConfirmHeight + $iConfirmMarginBottom + $aiWinMainSizeBorder[3] - $aiWinMainSize[1]
   If $iWinMainHeightNew > $iWinMainHeight Then
	  $iWinMainHeight = $iWinMainHeightNew
	  WinMove($hWinMain, "", Default, Default, Default, $iWinMainHeight)
   EndIf
   ; Set group Font;
   For $i = 1 To $aidGroups[0]
	  GUICtrlSetFont($aidGroups[$i], $iFontSize, $iGroupFontWeight, Default, $iFontName)
   Next
   ; Set label Font;
   For $i = 1 To $aidLabels[0]
	  GUICtrlSetFont($aidLabels[$i], $iFontSize, $iNormalFontWeight, Default, $iFontName)
   Next
   ; Set checkbox Font;
   For $i = 1 To $aidCheckboxes[0]
	  GUICtrlSetFont($aidCheckboxes[$i], $iFontSize, $iNormalFontWeight, Default, $iFontName)
   Next
   ; Set control Font;
   For $i = 1 To $aidControls[0]
	  GUICtrlSetFont($aidControls[$i], $iFontSize, $iNormalFontWeight, Default, $iFontName)
   Next
   GUISetState(@SW_UNLOCK, $hWinMain)
EndFunc





; #####################
; # COLOR MAIN WINDOW #
; #####################
#CS
Func _WinMainColor()
   ; Load the new color settings.
   _WinMainCalibrate()
   ; Set window color;
   GUISetBkColor($iWinBKColor)
   ; Set group colors;
   GUISetState(@SW_LOCK, $hWinMain)
   GUICtrlSetColor($idTab, $iLabelColor)
   GUICtrlSetBkColor($idTab, $iLabelBKColor)
   For $i = 1 To $aidGroups[0]
	  GUICtrlSetColor($aidGroups[$i], $iGroupColor)
	  GUICtrlSetBkColor($aidGroups[$i], $iGroupBKColor)
   Next
   ; Set label colors;
   For $i = 1 To $aidLabels[0]
	  GUICtrlSetColor($aidLabels[$i], $iLabelColor)
	  GUICtrlSetBkColor($aidLabels[$i], $iLabelBKColor)
   Next
   ; Set checkbox colors;
   For $i = 1 To $aidCheckboxes[0]
	  GUICtrlSetColor($aidCheckboxes[$i], $iGroupColor)
	  GUICtrlSetBkColor($aidCheckboxes[$i], $iGroupBKColor)
   Next
   ; Set control colors;
   For $i = 1 To $aidControls[0]
	  GUICtrlSetColor($aidControls[$i], $iControlColor)
	  GUICtrlSetBkColor($aidControls[$i], $iControlBKColor)
   Next
   GUISetState(@SW_UNLOCK, $hWinMain)
EndFunc
#CE





; ####################
; # DRAW MAIN WINDOW #
; ####################
Func _WinMainDraw()
   _WinMainCalibrate()
   GUISetState(@SW_LOCK, $hWinMain)
   GUICtrlSetPos($idGroupTo, $iGroupMarginLeft, $iGroupMarginTop, $iGroupWidth1, $iGroupHeight1)
	  $iLabelOffsetLeft = $iGroupMarginLeft + $iLabelMarginLeft
	  $iLabelOffsetTop = $iGroupMarginTop + $iLabelMarginTop
	  $iCtrlOffsetLeft = $iLabelOffsetLeft + $iLabelWidth + $iMarginBetweenCtrlAndLabel
	  $iCtrlOffsetTop = $iGroupMarginTop + $iCtrlMarginTop
	  $iLine = 0
	  GUICtrlSetPos($aidLabels[1], $iLabelOffsetLeft, $iLabelOffsetTop + $iLineHeight * $iLine, $iLabelWidth, $iCtrlHeight)
	  GUICtrlSetPos($aidControls[1], $iCtrlOffsetLeft, $iCtrlOffsetTop + $iLineHeight * $iLine, $iCtrlWidth1, $iCtrlHeight)
	  $iLine += 1
	  GUICtrlSetPos($aidLabels[2], $iLabelOffsetLeft, $iLabelOffsetTop + $iLineHeight * $iLine, $iLabelWidth, $iCtrlHeight)
	  GUICtrlSetPos($aidControls[2], $iCtrlOffsetLeft, $iCtrlOffsetTop + $iLineHeight * $iLine, $iCtrlWidth1, $iCtrlHeight)
	  $iLine += 1
	  GUICtrlSetPos($aidLabels[3], $iLabelOffsetLeft, $iLabelOffsetTop + $iLineHeight * $iLine, $iLabelWidth, $iCtrlHeight)
	  GUICtrlSetPos($aidControls[3], $iCtrlOffsetLeft, $iCtrlOffsetTop + $iLineHeight * $iLine, $iCtrlWidth1, $iCtrlHeight)
	  $iLine += 1
	  GUICtrlSetPos($aidLabels[4], $iLabelOffsetLeft, $iLabelOffsetTop + $iLineHeight * $iLine, $iLabelWidth, $iCtrlHeight)
	  GUICtrlSetPos($aidControls[4], $iCtrlOffsetLeft, $iCtrlOffsetTop + $iLineHeight * $iLine, $iCtrlWidth1, $iCtrlHeight)
	  $iLine += 1
	  GUICtrlSetPos($aidLabels[5], $iLabelOffsetLeft, $iLabelOffsetTop + $iLineHeight * $iLine, $iLabelWidth, $iCtrlHeight)
	  GUICtrlSetPos($aidControls[5], $iCtrlOffsetLeft, $iCtrlOffsetTop + $iLineHeight * $iLine, $iCtrlWidth1, $iCtrlHeight)
	  $iLine += 1
	  GUICtrlSetPos($aidLabels[6], $iLabelOffsetLeft, $iLabelOffsetTop + $iLineHeight * $iLine, $iLabelWidth, $iCtrlHeight)
	  GUICtrlSetPos($aidControls[6], $iCtrlOffsetLeft, $iCtrlOffsetTop + $iLineHeight * $iLine, $iCtrlWidth1, $iCtrlHeight)
	  $iLine += 1
	  GUICtrlSetPos($aidLabels[7], $iLabelOffsetLeft, $iLabelOffsetTop + $iLineHeight * $iLine, $iLabelWidth, $iCtrlHeight)
	  GUICtrlSetPos($aidControls[7], $iCtrlOffsetLeft, $iCtrlOffsetTop + $iLineHeight * $iLine, $iCtrlWidth1, $iCtrlHeight)
	  $iLine += 1
	  GUICtrlSetPos($aidLabels[8], $iLabelOffsetLeft, $iLabelOffsetTop + $iLineHeight * $iLine, $iLabelWidth, $iCtrlHeight)
	  GUICtrlSetPos($aidControls[8],  $iCtrlOffsetLeft, $iCtrlOffsetTop + $iLineHeight * $iLine, $iCtrlWidth1, $iCtrlHeight)
	  $iLine += 1
	  GUICtrlSetPos($aidLabels[9], $iLabelOffsetLeft, $iLabelOffsetTop + $iLineHeight * $iLine, $iLabelWidth, $iCtrlHeight)
	  GUICtrlSetPos($aidControls[9], $iCtrlOffsetLeft, $iCtrlOffsetTop + $iLineHeight * $iLine, $iCtrlWidth1, $iCtrlHeight)
	  $iLine += 1
	  GUICtrlSetPos($aidLabels[10], $iLabelOffsetLeft, $iLabelOffsetTop + $iLineHeight * $iLine, $iLabelWidth, $iCtrlHeight)
	  GUICtrlSetPos($aidControls[10], $iCtrlOffsetLeft, $iCtrlOffsetTop + $iLineHeight * $iLine, $iCtrlWidth1, $iCtrlHeight)
	  $iLine += 2
	  GUICtrlSetPos($aidLabels[11], $iLabelOffsetLeft, $iLabelOffsetTop + $iLineHeight * $iLine, $iLabelWidth, $iCtrlHeight)
	  GUICtrlSetPos($aidControls[11], $iCtrlOffsetLeft, $iCtrlOffsetTop + $iLineHeight * $iLine, $iCtrlWidth1, $iCtrlHeight)
   GUICtrlSetPos($idGroupContact, $iGroupMarginLeft + $iGroupMarginMiddleHor + $iGroupWidth1, $iGroupMarginTop, $iGroupWidth1, $iGroupHeight1)
	  $iLabelOffsetLeft = $iGroupMarginLeft + $iGroupWidth1 + $iGroupMarginMiddleHor + $iLabelMarginLeft
	  $iLabelOffsetTop = $iGroupMarginTop + $iLabelMarginTop
	  $iCtrlOffsetLeft = $iLabelOffsetLeft + $iLabelWidth + $iMarginBetweenCtrlAndLabel
	  $iCtrlOffsetTop = $iGroupMarginTop + $iCtrlMarginTop
	  $iLine = 0
	  GUICtrlSetPos($aidLabels[12], $iLabelOffsetLeft, $iLabelOffsetTop + $iLineHeight * $iLine, $iLabelWidth, $iCtrlHeight)
	  GUICtrlSetPos($aidControls[12], $iCtrlOffsetLeft, $iCtrlOffsetTop + $iLineHeight * $iLine, $iCtrlWidth1, $iCtrlHeight)
	  $iLine += 2
	  GUICtrlSetPos($aidLabels[13], $iLabelOffsetLeft, $iLabelOffsetTop + $iLineHeight * $iLine, $iLabelWidth, $iCtrlHeight)
	  GUICtrlSetPos($aidControls[13], $iCtrlOffsetLeft, $iCtrlOffsetTop + $iLineHeight * $iLine, $iCtrlWidth1, $iCtrlHeight)
	  $iLine += 1
	  GUICtrlSetPos($aidLabels[14], $iLabelOffsetLeft, $iLabelOffsetTop + $iLineHeight * $iLine, $iLabelWidth, $iCtrlHeight)
	  GUICtrlSetPos($aidControls[14], $iCtrlOffsetLeft, $iCtrlOffsetTop + $iLineHeight * $iLine, $iCtrlWidth1, $iCtrlHeight)
	  $iLine += 1
	  GUICtrlSetPos($aidLabels[15], $iLabelOffsetLeft, $iLabelOffsetTop + $iLineHeight * $iLine, $iLabelWidth, $iCtrlHeight)
	  GUICtrlSetPos($aidControls[15], $iCtrlOffsetLeft, $iCtrlOffsetTop + $iLineHeight * $iLine, $iCtrlWidth1, $iCtrlHeight)
	  $iLine += 1
	  GUICtrlSetPos($aidLabels[16], $iLabelOffsetLeft, $iLabelOffsetTop + $iLineHeight * $iLine, $iLabelWidth, $iCtrlHeight)
	  GUICtrlSetPos($aidControls[16], $iCtrlOffsetLeft, $iCtrlOffsetTop + $iLineHeight * $iLine, $iCtrlWidth1, $iCtrlHeight)
	  $iLine += 1
	  GUICtrlSetPos($aidLabels[17], $iLabelOffsetLeft, $iLabelOffsetTop + $iLineHeight * $iLine, $iLabelWidth, $iCtrlHeight)
	  GUICtrlSetPos($aidControls[17], $iCtrlOffsetLeft, $iCtrlOffsetTop + $iLineHeight * $iLine, $iCtrlWidth1, $iCtrlHeight)
	  $iLine += 2
	  GUICtrlSetPos($aidLabels[18], $iLabelOffsetLeft, $iLabelOffsetTop + $iLineHeight * $iLine, $iLabelWidth, $iCtrlHeight)
	  GUICtrlSetPos($aidControls[18], $iCtrlOffsetLeft, $iCtrlOffsetTop + $iLineHeight * $iLine, $iCtrlWidth1, $iCtrlHeight)
	  $iLine += 1
	  GUICtrlSetPos($aidLabels[19], $iLabelOffsetLeft, $iLabelOffsetTop + $iLineHeight * $iLine, $iLabelWidth, $iCtrlHeight)
	  GUICtrlSetPos($aidControls[19], $iCtrlOffsetLeft, $iCtrlOffsetTop + $iLineHeight * $iLine, $iCtrlWidth1, $iCtrlHeight)
	  $iLine += 2
	  GUICtrlSetPos($aidLabels[20], $iLabelOffsetLeft, $iLabelOffsetTop + $iLineHeight * $iLine, $iLabelWidth, $iCtrlHeight)
	  GUICtrlSetPos($aidControls[20], $iCtrlOffsetLeft, $iCtrlOffsetTop + $iLineHeight * $iLine, $iCtrlWidth1, $iCtrlHeight)
   GUICtrlSetPos($idGroupOptions, $iGroupMarginLeft, $iGroupHeight1 + $iGroupMarginMiddleVer + $iGroupMarginTop, $iGroupWidth2, $iGroupHeight2)
	  $iLabelOffsetLeft = $iGroupMarginLeft + $iLabelMarginLeft
	  $iLabelOffsetTop = $iGroupMarginTop + $iGroupHeight1 + $iGroupMarginMiddleVer + $iLabelMarginTop
	  $iCtrlOffsetLeft = $iLabelOffsetLeft + $iLabelWidth + $iMarginBetweenCtrlAndLabel
	  $iCtrlOffsetTop = $iGroupMarginTop + $iGroupHeight1 + $iGroupMarginMiddleVer + $iCtrlMarginTop
	  $iLine = 0
	  GUICtrlSetPos($aidCheckboxes[1], $iLabelOffsetLeft, $iLabelOffsetTop + $iLineHeight * $iLine, $iCtrlWidth3, $iCtrlHeight)
	  $iLine += 1
	  GUICtrlSetPos($aidCheckboxes[2], $iLabelOffsetLeft, $iLabelOffsetTop + $iLineHeight * $iLine, $iCtrlWidth3, $iCtrlHeight)
	  $iLine += 1
	  GUICtrlSetPos($aidCheckboxes[3], $iLabelOffsetLeft, $iLabelOffsetTop + $iLineHeight * $iLine, $iCtrlWidth3, $iCtrlHeight)
	  $iLabelOffsetLeft = $iGroupMarginLeft + $iLabelMarginLeft * 2 + $iCtrlWidth3
	  $iLabelOffsetTop = $iGroupMarginTop + $iGroupHeight1 + $iGroupMarginMiddleVer + $iLabelMarginTop
	  $iCtrlOffsetLeft = $iLabelOffsetLeft + $iLabelWidth + $iMarginBetweenCtrlAndLabel
	  $iCtrlOffsetTop = $iGroupMarginTop + $iGroupHeight1 + $iGroupMarginMiddleVer + $iCtrlMarginTop
	  $iLine = 0
	  GUICtrlSetPos($aidLabels[21], $iLabelOffsetLeft, $iLabelOffsetTop + $iLineHeight * $iLine, $iLabelWidth, $iCtrlHeight)
	  GUICtrlSetPos($aidControls[21], $iCtrlOffsetLeft, $iCtrlOffsetTop + $iLineHeight * $iLine, $iCtrlWidth2 - $iCtrlHeight - $iCtrlMarginRight, $iCtrlHeight)
	  GUICtrlSetPos($idOptionsExplorer, $iCtrlOffsetLeft + $iCtrlWidth2 - $iCtrlHeight - $iCtrlMarginRight / 2 + 1, $iCtrlOffsetTop + $iLineHeight * $iLine, $iCtrlHeight + $iCtrlMarginRight / 2, $iCtrlHeight)
	  $iLine += 1
	  GUICtrlSetPos($aidLabels[22], $iLabelOffsetLeft, $iLabelOffsetTop + $iLineHeight * $iLine, $iLabelWidth, $iCtrlHeight)
	  GUICtrlSetPos($aidControls[22], $iCtrlOffsetLeft, $iCtrlOffsetTop + $iLineHeight * $iLine, $iCtrlWidth2, $iCtrlHeight)
   GUICtrlSetPos($idConfirm, ($iWinMainWidth - $iConfirmWidth) / 2, $iGroupMarginTop + $iGroupHeight1 + $iGroupHeight2 +  $iGroupMarginMiddleVer + $iConfirmMarginTop, $iConfirmWidth, $iCtrlHeight)
   GUICtrlSetPos($idLettersSelectAll, $iCtrlMarginLeft, $iGroupMarginTop, ($iWinMainWidth - $iCtrlMarginLeft - $iCtrlMarginRight - $iCtrlMarginMiddleHor) / 2, $iCtrlHeight)
   GUICtrlSetPos($idLettersSelectNone, $iCtrlMarginLeft + $iCtrlMarginMiddleHor + ($iWinMainWidth - $iCtrlMarginLeft - $iCtrlMarginRight - $iCtrlMarginMiddleHor) / 2, $iGroupMarginTop, ($iWinMainWidth - $iCtrlMarginLeft - $iCtrlMarginRight - $iCtrlMarginMiddleHor) / 2, $iCtrlHeight)
   GUICtrlSetPos($idListLetters, $iGroupMarginLeft, $iGroupMarginTop + $iCtrlHeight + $iGroupMarginBottom, $iWinMainWidth - $iGroupMarginLeft - $iGroupMarginRight, $iWinMainHeight - $iGroupMarginTop - $iCtrlHeight * 2 - $iGroupMarginMiddleVer - $iGroupMarginBottom * 2)
   GUICtrlSetPos($idLettersOpen, $iCtrlMarginLeft, $iWinMainHeight - $iCtrlHeight - $iGroupMarginBottom, ($iWinMainWidth - $iCtrlMarginLeft - $iCtrlMarginRight - $iCtrlMarginMiddleHor * 2) / 3, $iCtrlHeight)
   GUICtrlSetPos($idLettersRemove, $iCtrlMarginLeft + $iCtrlMarginMiddleHor + ($iWinMainWidth - $iCtrlMarginLeft - $iCtrlMarginRight - $iCtrlMarginMiddleHor * 2) / 3, $iWinMainHeight - $iCtrlHeight - $iGroupMarginBottom, ($iWinMainWidth - $iCtrlMarginLeft - $iCtrlMarginRight - $iCtrlMarginMiddleHor * 2) / 3, $iCtrlHeight)
   GUICtrlSetPos($idLettersSend, $iCtrlMarginLeft + $iCtrlMarginMiddleHor * 2 + ($iWinMainWidth - $iCtrlMarginLeft - $iCtrlMarginRight - $iCtrlMarginMiddleHor * 2) / 3 * 2, $iWinMainHeight - $iCtrlHeight - $iGroupMarginBottom, ($iWinMainWidth - $iCtrlMarginLeft - $iCtrlMarginRight - $iCtrlMarginMiddleHor * 2) / 3, $iCtrlHeight)
   GUICtrlSetState($idLettersOpen, $GUI_DISABLE)
   GUICtrlSetState($idLettersRemove, $GUI_DISABLE)
   GUICtrlSetState($idLettersSend, $GUI_DISABLE)
   GUICtrlSetPos($idListHistory, $iGroupMarginLeft, $iGroupMarginTop, $iWinMainWidth - $iGroupMarginLeft - $iGroupMarginRight, $iWinMainHeight - $iGroupMarginTop - $iGroupMarginBottom)
   If not StringCompare($sPrint, "Ja") Then
	  GUICtrlSetPos($aidLabels[23], $iGroupMarginLeft, $iGroupMarginTop + $iLabelMarginTop - $iCtrlMarginTop, $iLabelWidth, $iCtrlHeight)
	  GUICtrlSetPos($idDatabaseDate, $iGroupMarginLeft + $iLabelWidth + $iMarginBetweenCtrlAndLabel, $iGroupMarginTop, ($iWinMainWidth - $iGroupMarginLeft - $iGroupMarginRight - $iCtrlMarginMiddleHor) / 2 - $iLabelWidth - $iMarginBetweenCtrlAndLabel, $iCtrlHeight)
	  GUICtrlSetPos($idDatabaseReload, $iCtrlMarginLeft + $iCtrlMarginMiddleHor + ($iWinMainWidth - $iCtrlMarginLeft - $iCtrlMarginRight - $iCtrlMarginMiddleHor) / 2, $iGroupMarginTop, ($iWinMainWidth - $iCtrlMarginLeft - $iCtrlMarginRight - $iCtrlMarginMiddleHor) / 2, $iCtrlHeight)
	  GUICtrlSetPos($idDatabaseCount, $iGroupMarginLeft, $iGroupMarginTop + $iCtrlHeight + $iCtrlMarginMiddleVer + $iLabelMarginTop - $iCtrlMarginTop, ($iWinMainWidth - $iCtrlMarginLeft - $iCtrlMarginRight - $iCtrlMarginMiddleHor) / 2, $iCtrlHeight)
	  GUICtrlSetPos($idDatabaseSelect, $iCtrlMarginLeft + $iCtrlMarginMiddleHor + ($iWinMainWidth - $iCtrlMarginLeft - $iCtrlMarginRight - $iCtrlMarginMiddleHor) / 2, $iGroupMarginTop + $iLineHeight, ($iWinMainWidth - $iCtrlMarginLeft - $iCtrlMarginRight - $iCtrlMarginMiddleHor) / 2, $iCtrlHeight)
	  GUICtrlSetPos($idDatabaseLettersCustomize, $iGroupMarginLeft, $iGroupMarginTop + $iCtrlHeight * 2 + $iGroupMarginBottom + $iGroupMarginMiddleVer * 1, $iWinMainWidth - $iGroupMarginLeft - $iGroupMarginRight, $iCtrlHeight)
	  GUICtrlSetPos($idDatabaseLetters, $iGroupMarginLeft, $iGroupMarginTop + $iCtrlHeight * 3 + $iGroupMarginBottom + $iGroupMarginMiddleVer * 2, $iWinMainWidth - $iGroupMarginLeft - $iGroupMarginRight, $iWinMainHeight - $iGroupMarginTop - $iCtrlHeight * 4 - $iGroupMarginMiddleVer * 4 - $iGroupMarginBottom)
	  GUICtrlSetPos($idDatabaseOpen, $iCtrlMarginLeft, $iWinMainHeight - $iCtrlHeight - $iGroupMarginBottom, ($iWinMainWidth - $iCtrlMarginLeft - $iCtrlMarginRight - $iCtrlMarginMiddleHor * 4) / 4, $iCtrlHeight)
	  GUICtrlSetPos($idDatabasePrint, $iCtrlMarginLeft + $iCtrlMarginMiddleHor + ($iWinMainWidth - $iCtrlMarginLeft - $iCtrlMarginRight - $iCtrlMarginMiddleHor * 3) / 4, $iWinMainHeight - $iCtrlHeight - $iGroupMarginBottom, ($iWinMainWidth - $iCtrlMarginLeft - $iCtrlMarginRight - $iCtrlMarginMiddleHor * 3) / 4, $iCtrlHeight)
	  GUICtrlSetPos($idDatabaseMarkSigned, $iCtrlMarginLeft + $iCtrlMarginMiddleHor * 2 + ($iWinMainWidth - $iCtrlMarginLeft - $iCtrlMarginRight - $iCtrlMarginMiddleHor * 3) / 4 * 2, $iWinMainHeight - $iCtrlHeight - $iGroupMarginBottom, ($iWinMainWidth - $iCtrlMarginLeft - $iCtrlMarginRight - $iCtrlMarginMiddleHor * 3) / 4, $iCtrlHeight)
	  GUICtrlSetPos($idDatabaseMarkSent, $iCtrlMarginLeft + $iCtrlMarginMiddleHor * 3 + ($iWinMainWidth - $iCtrlMarginLeft - $iCtrlMarginRight - $iCtrlMarginMiddleHor * 3) / 4 * 3, $iWinMainHeight - $iCtrlHeight - $iGroupMarginBottom, ($iWinMainWidth - $iCtrlMarginLeft - $iCtrlMarginRight - $iCtrlMarginMiddleHor * 3) / 4, $iCtrlHeight)
	  GUICtrlSetState($idDatabaseOpen, $GUI_DISABLE)
	  GUICtrlSetState($idDatabasePrint, $GUI_DISABLE)
   EndIf
EndFunc





; ######################
; # DELETE MAIN WINDOW #
; ######################
Func _WinMainDelete()
   GUIDelete($hWinMain)
EndFunc





; ################
; # LOAD LETTERS #
; ################
Func _WinMainLoadLetters()
   Global $sLettersDir = "brieven\"
   For $i = 1 To $aLetters[0][0]
	  GUICtrlDelete($aLetters[$i][0])
   Next
   ReDim $aLetters[1][3]
   $aLetters[0][0] = 0
   $asLetters = _FileListToArray($sLettersDir, "*.txt", $FLTA_FILES)
   If @error Then
	  DirCreate(".\" & $sLettersDir)
   Else
	  For $i = 1 To $asLetters[0]
		 $hFile = FileOpen($sLettersDir & $asLetters[$i], $FO_UNICODE)
			$asLastAccessed = FileGetTime($sLettersDir & StringTrimRight($asLetters[$i], 4) & ".doc")
			$sLastAccessed = $asLastAccessed[2] & "/" & $asLastAccessed[1] & " " & $asLastAccessed[3] & ":" & $asLastAccessed[4]
			$sToOrg = FileReadLine($hFile, 7)
			$sFromSubject = FileReadLine($hFile, 1)
			$sToName = FileReadLine($hFile, 9) & " " & FileReadLine($hFile, 10)
			ReDim $aLetters[$i + 1][3]
			$aLetters[$i][1] = StringTrimRight($asLetters[$i], 4)
			$aLetters[$i][2] = $sLastAccessed & "|" & $sToOrg & "|" & $sFromSubject & "|" & $sToName
			$aLetters[$i][0] = GUICtrlCreateListViewItem($aLetters[$i][2], $idListLetters)
			$aLetters[0][0] += 1
		 FileClose($hFile)
	  Next
   EndIf
EndFunc





; ################
; # LOAD HISTORY #
; ################
Func _WinMainLoadHistory()
   If Not FileExists(@ScriptDir & "\geschiedenis.txt") Then _FileCreate(@ScriptDir & "\geschiedenis.txt")
   $hHistory = FileOpen(@ScriptDir & "\geschiedenis.txt", $FO_UNICODE)
   If Not @error Then
	  While 1
		 $sTemp = FileReadLine($hHistory)
		 If @error Then
			ExitLoop
		 Else
			GUICtrlCreateListViewItem($sTemp, $idListHistory)
		 EndIf
	  WEnd
	  FileClose($hHistory)
   EndIf
EndFunc





; ####################
; # READ MAIN WINDOW #
; ####################
Func _WinMainRead()
   Global $sToOrg = GUICtrlRead($idToOrg)
   Global $sToSalution = GUICtrlRead($idToSalution)
   Global $sToFirstName = GUICtrlRead($idToFirstName)
   Global $sToSurname = GUICtrlRead($idToSurname)
   Global $sToStreet = GUICtrlRead($idToStreet)
   Global $sToNumber = GUICtrlRead($idToNumber)
   Global $sToBus = GUICtrlRead($idToBus)
   Global $sToPostcode = GUICtrlRead($idToPostcode)
   Global $sToCity = GUICtrlRead($idToCity)
   Global $sToCountry = GUICtrlRead($idToCountry)
   Global $sToRef = GUICtrlRead($idToRef)
   Global $sFromOrg = GUICtrlRead($idFromOrg)
   Global $sFromFirstName = GUICtrlRead($idFromFirstName)
   Global $sFromSurname = GUICtrlRead($idFromSurname)
   Global $sFromDepartment = GUICtrlRead($idFromDepartment)
   Global $sFromTel = GUICtrlRead($idFromTel)
   Global $sFromEmail = GUICtrlRead($idFromEmail)
   Global $sFromSubject = GUICtrlRead($idFromSubject)
   Global $sFromDate = GUICtrlRead($idFromDate)
   If StringLen($sFromDate) = 9 Then $sFromDate = "0" & $sFromDate
   Global $sFromRef = GUICtrlRead($idFromRef)
   Global $bOptionsRegistered = BitAND(GUICtrlRead($idOptionsRegistered), $GUI_CHECKED)
   Global $bOptionsConfidential = BitAND(GUICtrlRead($idOptionsConfidential), $GUI_CHECKED)
   Global $sOptionsAnnex = GUICtrlRead($idOptionsAnnex)
   Global $sOptionsRemark = GUICtrlRead($idOptionsRemark)
   Global $bOptionsSigned = BitAND(GUICtrlRead($idOptionsSigned), $GUI_CHECKED)
EndFunc





; #####################
; # RESET MAIN WINDOW #
; #####################
Func _WinMainReset()
   GuiCtrlSetData($idToOrg, "")
   GuiCtrlSetData($idToSalution, "", "")
   GuiCtrlSetData($idToSalution, $sSalutions, StringLeft($sSalutions, Not StringInStr($sSalutions, "|")? StringLen($sSalutions) : StringInStr($sSalutions, "|") - 1))
   GuiCtrlSetData($idToFirstName, "")
   GuiCtrlSetData($idToSurname, "")
   GuiCtrlSetData($idToStreet, "")
   GuiCtrlSetData($idToNumber, "")
   GuiCtrlSetData($idToBus, "")
   GuiCtrlSetData($idToPostcode, "")
   GuiCtrlSetData($idToCity, "")
   GUICtrlSetData($idToCountry, $sCountries, "België")
   GuiCtrlSetData($idToRef, "")
   ;GUICtrlSetData($idFromOrg, "", "")
   ;GUICtrlSetData($idFromOrg, $sOrgs, "")
   ;GuiCtrlSetData($idFromFirstName, "")
   ;GuiCtrlSetData($idFromSurname, "")
   ;GuiCtrlSetData($idFromDepartment, "", "")
   ;GuiCtrlSetData($idFromTel, "")
   ;GuiCtrlSetData($idFromEmail, "")
   GuiCtrlSetData($idFromSubject, "")
   GuiCtrlSetData($idFromDate, @YEAR & "/" & @MON & "/" & @MDAY)
   ;GUICtrlSetData($idFromRef, IniRead($sSettingsFile, "Afkortingen" & GUICtrlRead($idFromOrg), GUICtrlRead($idFromDepartment), "___") & "/" & StringLower(StringCompare(GUICtrlRead($idFromFirstName), "") ? GUICtrlRead($idFromFirstName) : "___") & "/" & "#ID")
   GUICtrlSetData($idFromRef, "#ID")
   GUICtrlSetState($idOptionsRegistered, $GUI_UNCHECKED)
   GuiCtrlSetState($idOptionsConfidential, $GUI_UNCHECKED)
   GuiCtrlSetData($idOptionsAnnex, "")
   GuiCtrlSetData($idOptionsRemark, "")
   GUICtrlSetState($idOptionsSigned, $GUI_CHECKED)
EndFunc