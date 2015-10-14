; ################
; # ORDER WINDOW #
; ################
Func _WinOrder()

   Local $asTempTabPrintShown[UBound($asTabPrintShown)]
   Local $asTempTabPrintHidden[UBound($asTabPrintHidden)]

   For $i = 0 To UBound($asTabPrintShown) - 1
	  $asTempTabPrintShown[$i] = $asTabPrintShown[$i]
   Next

   For $i = 0 To UBound($asTabPrintHidden) - 1
	  $asTempTabPrintHidden[$i] = $asTabPrintHidden[$i]
   Next

   GUISetState(@SW_DISABLE, $hWinMain)

   $iWinOrderWidth = 400
   $iWinOrderHeight = 400
   $hWinOrder = GUICreate("Lijst aanpassen", $iWinOrderWidth, $iWinOrderHeight, Default, Default, $WS_POPUP + $WS_CAPTION + $WS_SYSMENU)

	  GUISetIcon($sProgramLogo)

	  GUICtrlCreateLabel("Getoond:", 15, 15, $iWinOrderWidth / 2 - 40, 20)
	  $idOrderShown = GUICtrlCreateList("", 15, 35, $iWinOrderWidth / 2 - 40, $iWinOrderHeight - 120, BitAND($WS_BORDER, $WS_VSCROLL))
	  $idOrderShownUp = GUICtrlCreateButton("Omhoog", 15, $iWinOrderHeight - 80, ($iWinOrderWidth / 2 - 40) / 2 - 3, 25)
	  $idOrderShownDown = GUICtrlCreateButton("Omlaag", ($iWinOrderWidth / 2 - 40) / 2 + 18, $iWinOrderHeight - 80, ($iWinOrderWidth / 2 - 40) / 2 - 3, 25)

	  $idOrderAdd = GUICtrlCreateButton("<<", $iWinOrderWidth / 2 - 15, ($iWinOrderHeight - 120) / 2, 30, 30)
	  $idOrderRemove = GUICtrlCreateButton(">>", $iWinOrderWidth / 2 - 15, ($iWinOrderHeight - 120) / 2 + 40, 30, 30)

	  GUICtrlCreateLabel("Niet Getoond:", $iWinOrderWidth / 2 + 25, 15, $iWinOrderWidth / 2 - 40, 20)
	  $idOrderHidden = GUICtrlCreateList("", $iWinOrderWidth / 2 + 25, 35, $iWinOrderWidth / 2 - 40, $iWinOrderHeight - 120, BitAND($WS_BORDER, $WS_VSCROLL))
	  $idOrderHiddenUp = GUICtrlCreateButton("Omhoog", $iWinOrderWidth / 2 + 25, $iWinOrderHeight - 80, ($iWinOrderWidth / 2 - 40) / 2 - 3, 25)
	  $idOrderHiddenDown = GUICtrlCreateButton("Omlaag", ($iWinOrderWidth / 2 - 40) / 2 + $iWinOrderWidth / 2 + 28, $iWinOrderHeight - 80, ($iWinOrderWidth / 2 - 40) / 2 - 3, 25)

	  Local $idCancel = GUICtrlCreateButton("Annuleren", $iWinOrderWidth - 211, $iWinOrderHeight - 30, 100, 25)
	  Local $idConfirm = GUICtrlCreateButton("Bevestigen", $iWinOrderWidth - 105, $iWinOrderHeight - 30, 100, 25)

	  If UBound($asTempTabPrintShown) > 0 Then
		 GUICtrlSetData($idOrderShown, _ArrayToString($asTempTabPrintShown))
	  EndIf
	  If UBound($asTempTabPrintHidden) > 0 Then
		 GUICtrlSetData($idOrderHidden, _ArrayToString($asTempTabPrintHidden))
	  EndIf

   GUISetState(@SW_SHOW, $hWinOrder)

   While True
	  Switch GUIGetMsg()
		 Case $idOrderAdd
			$sTemp = GUICtrlRead($idOrderHidden)
			If Not $sTemp = 0 Then
			   For $i = 0 To UBound($asTempTabPrintHidden)
				  If Not StringCompare($sTemp, $asTempTabPrintHidden[$i]) Then
					 ReDim $asTempTabPrintShown[UBound($asTempTabPrintShown) + 1]
					 $asTempTabPrintShown[UBound($asTempTabPrintShown) - 1] = $asTempTabPrintHidden[$i]
					 For $k = $i + 1 To UBound($asTempTabPrintHidden) - 1
						$asTempTabPrintHidden[$k - 1] = $asTempTabPrintHidden[$k]
					 Next
					 ReDim $asTempTabPrintHidden[UBound($asTempTabPrintHidden) - 1]
					 GUICtrlDelete($idOrderShown)
					 $idOrderShown = GUICtrlCreateList("", 15, 35, $iWinOrderWidth / 2 - 40, $iWinOrderHeight - 120, BitAND($WS_BORDER, $WS_VSCROLL))
					 GUICtrlSetData($idOrderShown, _ArrayToString($asTempTabPrintShown))
					 GUICtrlDelete($idOrderHidden)
					 $idOrderHidden = GUICtrlCreateList("", $iWinOrderWidth / 2 + 25, 35, $iWinOrderWidth / 2 - 40, $iWinOrderHeight - 120, BitAND($WS_BORDER, $WS_VSCROLL))
					 If UBound($asTempTabPrintHidden) > 0 Then
						GUICtrlSetData($idOrderHidden, _ArrayToString($asTempTabPrintHidden))
					 EndIf
					 ExitLoop
				  EndIf
			   Next
			EndIf
		 Case $idOrderRemove
			$sTemp = GUICtrlRead($idOrderShown)
			If Not $sTemp = 0 Then
			   For $i = 0 To UBound($asTempTabPrintShown)
				  If Not StringCompare($sTemp, $asTempTabPrintShown[$i]) Then
					 ReDim $asTempTabPrintHidden[UBound($asTempTabPrintHidden) + 1]
					 $asTempTabPrintHidden[UBound($asTempTabPrintHidden) - 1] = $asTempTabPrintShown[$i]
					 For $k = $i + 1 To UBound($asTempTabPrintShown) - 1
						$asTempTabPrintShown[$k - 1] = $asTempTabPrintShown[$k]
					 Next
					 ReDim $asTempTabPrintShown[UBound($asTempTabPrintShown) - 1]
					 GUICtrlDelete($idOrderShown)
					 $idOrderShown = GUICtrlCreateList("", 15, 35, $iWinOrderWidth / 2 - 40, $iWinOrderHeight - 120, BitAND($WS_BORDER, $WS_VSCROLL))
					 If UBound($asTempTabPrintShown) > 0 Then
						GUICtrlSetData($idOrderShown, _ArrayToString($asTempTabPrintShown))
					 EndIf
					 GUICtrlDelete($idOrderHidden)
					 $idOrderHidden = GUICtrlCreateList("", $iWinOrderWidth / 2 + 25, 35, $iWinOrderWidth / 2 - 40, $iWinOrderHeight - 120, BitAND($WS_BORDER, $WS_VSCROLL))
					 GUICtrlSetData($idOrderHidden, _ArrayToString($asTempTabPrintHidden))
					 ExitLoop
				  EndIf
			   Next
			EndIf
		 Case $idOrderShownUp
			$sTemp = GUICtrlRead($idOrderShown)
			If Not $sTemp = 0 Then
			   For $i = 0 To UBound($asTempTabPrintShown)
				  If Not StringCompare($sTemp, $asTempTabPrintShown[$i]) Then
					 If $i > 0 Then
						$sTemp = $asTempTabPrintShown[$i - 1]
						$asTempTabPrintShown[$i - 1] = $asTempTabPrintShown[$i]
						$asTempTabPrintShown[$i] = $sTemp
						GUICtrlDelete($idOrderShown)
						$idOrderShown = GUICtrlCreateList("", 15, 35, $iWinOrderWidth / 2 - 40, $iWinOrderHeight - 120, BitAND($WS_BORDER, $WS_VSCROLL))
						GUICtrlSetData($idOrderShown, _ArrayToString($asTempTabPrintShown))
						ControlFocus($hWinOrder, "", $idOrderShown)
						GUISetState(@SW_LOCK, $hWinOrder)
						For $j = 0 To $i - 1
						   Send("{DOWN}")
						Next
						GUISetState(@SW_UNLOCK, $hWinOrder)
					 EndIf
					 ExitLoop
				  EndIf
			   Next
			EndIf
		 Case $idOrderShownDown
			$sTemp = GUICtrlRead($idOrderShown)
			If Not $sTemp = 0 Then
			   For $i = 0 To UBound($asTempTabPrintShown)
				  If Not StringCompare($sTemp, $asTempTabPrintShown[$i]) Then
					 If $i < UBound($asTempTabPrintShown) - 1 Then
						$sTemp = $asTempTabPrintShown[$i + 1]
						$asTempTabPrintShown[$i + 1] = $asTempTabPrintShown[$i]
						$asTempTabPrintShown[$i] = $sTemp
						GUICtrlDelete($idOrderShown)
						$idOrderShown = GUICtrlCreateList("", 15, 35, $iWinOrderWidth / 2 - 40, $iWinOrderHeight - 120, BitAND($WS_BORDER, $WS_VSCROLL))
						GUICtrlSetData($idOrderShown, _ArrayToString($asTempTabPrintShown))
						ControlFocus($hWinOrder, "", $idOrderShown)
						GUISetState(@SW_LOCK, $hWinOrder)
						For $j = 0 To $i + 1
						   Send("{DOWN}")
						Next
						GUISetState(@SW_UNLOCK, $hWinOrder)
					 EndIf
					 ExitLoop
				  EndIf
			   Next
			EndIf
		 Case $idOrderHiddenUp
			$sTemp = GUICtrlRead($idOrderHidden)
			If Not $sTemp = 0 Then
			   For $i = 0 To UBound($asTempTabPrintHidden)
				  If Not StringCompare($sTemp, $asTempTabPrintHidden[$i]) Then
					 If $i > 0 Then
						$sTemp = $asTempTabPrintHidden[$i - 1]
						$asTempTabPrintHidden[$i - 1] = $asTempTabPrintHidden[$i]
						$asTempTabPrintHidden[$i] = $sTemp
						GUICtrlDelete($idOrderHidden)
						$idOrderHidden = GUICtrlCreateList("", $iWinOrderWidth / 2 + 25, 35, $iWinOrderWidth / 2 - 40, $iWinOrderHeight - 120, BitAND($WS_BORDER, $WS_VSCROLL))
						GUICtrlSetData($idOrderHidden, _ArrayToString($asTempTabPrintHidden))
						ControlFocus($hWinOrder, "", $idOrderHidden)
						GUISetState(@SW_LOCK, $hWinOrder)
						For $j = 0 To $i - 1
						   Send("{DOWN}")
						Next
						GUISetState(@SW_UNLOCK, $hWinOrder)
					 EndIf
					 ExitLoop
				  EndIf
			   Next
			EndIf
		 Case $idOrderHiddenDown
			$sTemp = GUICtrlRead($idOrderHidden)
			If Not $sTemp = 0 Then
			   For $i = 0 To UBound($asTempTabPrintHidden)
				  If Not StringCompare($sTemp, $asTempTabPrintHidden[$i]) Then
					 If $i < UBound($asTempTabPrintHidden) - 1 Then
						$sTemp = $asTempTabPrintHidden[$i + 1]
						$asTempTabPrintHidden[$i + 1] = $asTempTabPrintHidden[$i]
						$asTempTabPrintHidden[$i] = $sTemp
						GUICtrlDelete($idOrderHidden)
						$idOrderHidden = GUICtrlCreateList("", $iWinOrderWidth / 2 + 25, 35, $iWinOrderWidth / 2 - 40, $iWinOrderHeight - 120, BitAND($WS_BORDER, $WS_VSCROLL))
						GUICtrlSetData($idOrderHidden, _ArrayToString($asTempTabPrintHidden))
						ControlFocus($hWinOrder, "", $idOrderHidden)
						GUISetState(@SW_LOCK, $hWinOrder)
						For $j = 0 To $i + 1
						   Send("{DOWN}")
						Next
						GUISetState(@SW_UNLOCK, $hWinOrder)
					 EndIf
					 ExitLoop
				  EndIf
			   Next
			EndIf
		 Case $idCancel
			GUISetState(@SW_ENABLE, $hWinMain)
			GUIDelete($hWinOrder)
			Return False
		 Case $idConfirm
			IniWriteSection($sSettingsFile, "PrintTabKolommen", "    =" & StringReplace(_ArrayToString($asTempTabPrintShown), "|", @LF & "    ="))
			$asTabPrintShown = $asTempTabPrintShown
			$asTabPrintHidden = $asTempTabPrintHidden
			GUISetState(@SW_ENABLE, $hWinMain)
			GUIDelete($hWinOrder)
			Return True
		 Case $GUI_EVENT_CLOSE
			GUISetState(@SW_ENABLE, $hWinMain)
			GUIDelete($hWinOrder)
			Return False
	  EndSwitch
   WEnd
EndFunc