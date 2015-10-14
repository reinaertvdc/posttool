; #########
; # PRINT #
; #########
Func _Print()
   $oWord = _Word_Create(False)
   If @error Then
	  MsgBox(48, "Waarschuwing", "De geselecteerde brieven kunnen geopend worden. Controleer of Microsoft Word 2003 of hoger is geïnstalleerd en correct werkt.")
	  return False
   EndIf

   ; Create PRINTDLG structure and set initial values for the number of copies, starting, and ending page
   Local $tPRINTDLG = DllStructCreate($tagPRINTDLG)
   DllStructSetData($tPRINTDLG, 'Size', DllStructGetSize($tPRINTDLG))
   DllStructSetData($tPRINTDLG, 'Flags', $PD_PAGENUMS)
   DllStructSetData($tPRINTDLG, 'FromPage', 1)
   DllStructSetData($tPRINTDLG, 'ToPage', 1)
   DllStructSetData($tPRINTDLG, 'MinPage', 1)
   DllStructSetData($tPRINTDLG, 'MaxPage', 1)
   DllStructSetData($tPRINTDLG, 'Copies', 1)

   ; Create Print dialog box

   If Not _WinAPI_PrintDlg($tPRINTDLG) Then
	  Return False
   EndIf

   ; Show results
   Local $hDevNames = DllStructGetData($tPRINTDLG, 'hDevNames')
   Local $pDevNames = _MemGlobalLock($hDevNames)
   Local $tDEVNAMES = DllStructCreate($tagDEVNAMES, $pDevNames)
   $iCopies = DllStructGetData($tPRINTDLG, 'Copies')
   $sPrinter = _WinAPI_GetString($pDevNames + 2 * DllStructGetData($tDEVNAMES, 'DeviceOffset'))
   If DllStructGetData($tDEVNAMES, 'Default') Then
	   $sPrinter = ""
   EndIf

   $bError = False
   For $j = 1 To $iCopies
	  For $i = 1 To $aDatabaseLetters[0][0]
		 If BitAND(GUICtrlRead($aDatabaseLetters[$i][0], 1), $GUI_CHECKED) Then
			$oDoc = _Word_DocOpen($oWord, $sDatabase & "\" & $aDatabaseLetters[$i][1] & ".doc")
			If @error Then
			   $bError = True
			Else
			   _Word_DocPrint($oDoc, True, 1, -1, True, $sPrinter)
			   _Word_DocClose($oDoc)
			   If $aDatabaseLetters[$i][6] Then
				  If Not StringCompare($sPrinter, "") Then
					 RunWait(@ComSpec & " /c " & 'SumatraPDF -print-to-default "' & $sDatabase & "\" & $aDatabaseLetters[$i][1] & '.pdf"', @ScriptDir, @SW_HIDE)
				  Else
					 RunWait(@ComSpec & " /c " & 'SumatraPDF -print-to "' & $sPrinter & '" "' & $sDatabase & "\" & $aDatabaseLetters[$i][1] & '.pdf"', @ScriptDir, @SW_HIDE)
				  EndIf
				  If @error Then $bError = True
				  Sleep($iPrintPause)
			   EndIf
			   If $aDatabaseLetters[$i][4] Then
				  $oDoc = _Word_DocOpen($oWord, $sDatabase & "\" & $aDatabaseLetters[$i][1] & "_aangetekend.doc")
				  If Not @error Then
					 _Word_DocPrint($oDoc, True, 1, -1, True, $sPrinter)
					 _Word_DocClose($oDoc)
				  EndIf
			   EndIf
			EndIf
		 EndIf
	  Next
   Next
   If $bError Then
	  MsgBox(48, "Waarschuwing", "Een of meerdere brieven konden niet worden afgedrukt. Mogelijke oorzaken:" & @CRLF & @CRLF & "- Microsoft Word 2003 of hoger is niet geïnstalleerd of werkt niet correct." & @CRLF & "- De bestanden zijn beschadigd of ontbreken.")
   Else
	  _DatabaseSetPrinted()
   EndIf

   ; Free global memory objects that contains a DEVMODE and DEVNAMES structures
   _MemGlobalFree(DllStructGetData($tPRINTDLG, 'hDevMode'))
   _MemGlobalFree($hDevNames)
EndFunc