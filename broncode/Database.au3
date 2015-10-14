; ########
; # SEND #
; ########
Func _DatabaseSend()
   DirCreate($sDatabase)
   If FileExists($sDatabase & "\bezet") Then
	  MsgBox($MB_ICONINFORMATION, "Even geduld", "De database is momenteel door iemand anders in gebruik, probeer het over enkele seconden opnieuw." & @CRLF & @CRLF & "Indien u meer dan een paar minuten moet wachten, vraag een systeembeheerder om hulp.")
	  Return False
   EndIf
   _FileCreate($sDatabase & "\bezet")
   $oExcel = _Excel_Open(False)
   If @error Then
	  FileDelete($sDatabase & "\bezet")
	  If $bError Then MsgBox($MB_ICONERROR, "Database error", "Het registerbestand kon niet worden geopend. Controleer of Microsoft Excel 2003 of hoger is geïnstalleerd en correct werkt.")
	  Return False
   EndIf
   $asRegisters = _FileListToArray($sDatabase, "register.*", $FLTA_FILES)
   Local $asHeading[][] = [["Registratienummer", _
							"Onderwerp", _
							"Vertrouwelijk", _
							"Aangetekend", _
							"Handtekening bestuur", _
							"Bijlage", _
							"Opmerking", _
							"Organisatie bestemmeling", _
							"Aanspreking", _
							"Voornaam bestemmeling", _
							"Achternaam bestemmeling", _
							"Straatnaam", _
							"Huisnummer", _
							"Bus", _
							"Postcode", _
							"Gemeente", _
							"Land", _
							"Uw kenmerk", _
							"Organisatie afzender", _
							"Voornaam afzender", _
							"Achternaam afzender", _
							"Dienst", _
							"Telefoon", _
							"E-mailadres", _
							"Datum", _
							"Ons kenmerk", _
							"Afgedrukt", _
							"Ondertekend", _
							"Verzonden"]]
   If @error Then
	  If @error = 4 Then
		 $oWorkbook = _Excel_BookNew($oExcel, 1)
		 _Excel_RangeWrite($oWorkbook, Default, $asHeading)
		 $oExcel.ActiveWorkbook.Sheets(1).Range("A1:AC1").Font.Bold = TRUE
		 _Excel_BookSaveAs($oWorkbook, $sDatabase & "\register.xls", $xlExcel8)
		 If @error Then
			_Excel_BookClose($oWorkbook, False)
			_Excel_Close($oExcel, False)
			FileDelete($sDatabase & "\bezet")
			MsgBox($MB_ICONERROR, "Database error", "Het programma kan geen nieuw register aanmaken. Suggesties:" & _
			   @CRLF & @CRLF & "- Controleer of de juiste map is ingesteld." & _
			   @CRLF & "- Controleer of de map bereikbaar is." & _
			   @CRLF & "- Controleer of het programma de juiste rechten heeft." & _
			   @CRLF & "- Controleer of de harde schijf niet vol is." & _
			   @CRLF & "- Controleer of Microsoft Excel 2003 of hoger is geïnstalleerd en correct werkt." & _
			   @CRLF & "- Probeer het later opnieuw.")
			Return False
		 EndIf
	  Else
		 FileDelete($sDatabase & "\bezet")
		 _Excel_Close($oExcel, False)
		 MsgBox($MB_ICONERROR, "Database error", "Het programma kan niet in de database schrijven. Suggesties:" & _
			@CRLF & @CRLF & "- Controleer of de juiste map is ingesteld." & _
			@CRLF & "- Controleer of de map bereikbaar is." & _
			@CRLF & "- Controleer of het programma de juiste rechten heeft." & _
			@CRLF & "- Probeer het later opnieuw.")
		 Return False
	  EndIf
   Else
	  $oWorkbook = _Excel_BookOpen($oExcel, $sDatabase & "\" & $asRegisters[1], Default, Default, True)
	  If @error Then
		 FileDelete($sDatabase & "\bezet")
		 _Excel_Close($oExcel, False)
		 MsgBox($MB_ICONERROR, "Database error", "Het programma kan het register niet openen. Suggesties:" & _
			@CRLF & @CRLF & "- Controleer of de juiste map is ingesteld." & _
			@CRLF & "- Controleer of de map bereikbaar is." & _
			@CRLF & "- Controleer of het programma de juiste rechten heeft." & _
			@CRLF & "- Controleer of Microsoft Excel 2003 of hoger is geïnstalleerd en correct werkt." & _
			@CRLF & "- Probeer het later opnieuw.")
		 Return False
	  EndIf
   EndIf
   $oSheet = $oWorkbook.Worksheets(1)
   $oSheet.Activate
   Local $aEntry[1][29]
   $bLettersMissing = False
   $bWritingFailed = False
   $bCopyFailed = False
   $asHeader = _Excel_RangeRead($oWorkbook, Default, "A1:AC1")
   Local $aiMapping[29] = [-1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1]
   For $i = 0 To UBound($asHeading, $UBOUND_COLUMNS) - 1
	  For $j = 0 To UBound($asHeader, $UBOUND_COLUMNS) - 1
		 If Not StringCompare($asHeading[0][$i], $asHeader[0][$j]) Then
			$aiMapping[$j] = $i
			$asHeading[0][$i] = ""
		 EndIf
	  Next
   Next
   For $i = 0 To UBound($aiMapping) - 1
	  If $aiMapping[$i] < 0 Then
		 For $j = 0 To UBound($asHeading, $UBOUND_COLUMNS) - 1
			If StringCompare($asHeading[0][$j], "") Then
			   $aiMapping[$i] = $j
			   $asHeading[0][$j] = ""
			   ExitLoop
			EndIf
		 Next
	  EndIf
   Next
   $oWord = _Word_Create(False)
   If @error Then
	  MsgBox(48, "Waarschuwing", "Het sjabloon kon niet worden geopend. Controleer of Microsoft Word 2003 of hoger is geïnstalleerd en correct werkt.")
	  Return False
   EndIf
   For $i = 1 To $aLetters[0][0]
	  If BitAND(GUICtrlRead($aLetters[$i][0], 1), $GUI_CHECKED) Then
		 $hFile = FileOpen(@ScriptDir & "\" & $sLettersDir & $aLetters[$i][1] & ".txt", $FO_UNICODE)
		 If @error Then
			$bLettersMissing = True
		 Else
			For $j = 1 To 25
			   $aEntry[0][$aiMapping[$j]] = FileReadLine($hFile)
			Next
			FileClose($hFile)
			If Int($aEntry[0][$aiMapping[2]]) Then
			   $aEntry[0][$aiMapping[2]] = "vertrouwelijk"
			Else
			   $aEntry[0][$aiMapping[2]] = ""
			EndIf
			If Int($aEntry[0][$aiMapping[3]]) Then
			   $aEntry[0][$aiMapping[3]] = "aangetekend"
			Else
			   $aEntry[0][$aiMapping[3]] = ""
			EndIf
			If Int($aEntry[0][$aiMapping[4]]) Then
			   $aEntry[0][$aiMapping[4]] = "handtekening bestuur"
			   $aEntry[0][$aiMapping[27]] = "Nee"
			Else
			   $aEntry[0][$aiMapping[4]] = ""
			   $aEntry[0][$aiMapping[27]] = "NVT"
			EndIf
			$aEntry[0][$aiMapping[26]] = "Nee"
			$aEntry[0][$aiMapping[28]] = "Nee"
			$oRange = $oSheet.UsedRange
			$oRange.SpecialCells($xlCellTypeLastCell).Activate
			$iRow = $oExcel.ActiveCell.Row + 1
			If $aiMapping[0] <= 26 Then
			   $sCol = Chr(65 + $aiMapping[0])
			Else
			   $sCol = "A" & Chr(39 + $aiMapping[0])
			EndIf
			$sIDPrev = _Excel_RangeRead($oWorkbook, Default, $sCol & ($iRow - 1))
			$sID = @YEAR & @MON & @MDAY
			$oExcel.Range("A" & $iRow).Activate
			If Not StringIsDigit($sIDPrev) Or StringCompare(StringLeft($sIDPrev, 8), $sID) Then
			   $sID &= "0001"
			Else
			   $iIDPrevNumber = 0
			   $iIDPrevNumber = Int(StringRight($sIDPrev, 4)) + 1
			   If $iIDPrevNumber < 10 Then
				  $sID &= "000" & $iIDPrevNumber
			   ElseIf $iIDPrevNumber < 100 Then
				  $sID &= "00" & $iIDPrevNumber
			   ElseIf $iIDPrevNumber < 1000 Then
				  $sID &= "0" & $iIDPrevNumber
			   Else
				  $sID &= $iIDPrevNumer
			   EndIf
			EndIf
			$aEntry[0][$aiMapping[0]] = $sID
			$oDoc = _Word_DocOpen($oWord, @WorkingDir & "\" & $sLettersDir & $aLetters[$i][1] & "_aangetekend.doc")
			If Not @error Then
			   ; find and replace in body
			   $oDoc.Range.Select()
			   _Word_DocFindReplace($oDoc, "#ID", $sID, $WdReplaceAll, 0, True, True)
			   $count_sections=$oDoc.Sections.Count()
			   For $sec_i=1 To $count_sections
				  ; select one header section
				  $oDoc.Sections($sec_i).Headers(1).Range.Select()
				  ; replace in header
				  _Word_DocFindReplace($oDoc, "#ID", $sID, $WdReplaceAll, -1, True, True)
			   Next
			   _Word_DocClose($oDoc, $WdSaveChanges)
			EndIf
			$oDoc = _Word_DocOpen($oWord, @WorkingDir & "\" & $sLettersDir & $aLetters[$i][1] & ".doc")
			If Not @error Then
			   ; find and replace in body
			   $oDoc.Range.Select()
			   _Word_DocFindReplace($oDoc, "#ID", $sID, $WdReplaceAll, 0, True, True)
			   $count_sections=$oDoc.Sections.Count()
			   For $sec_i=1 To $count_sections
				  ; select one header section
				  $oDoc.Sections($sec_i).Headers(1).Range.Select()
				  ; replace in header
				  _Word_DocFindReplace($oDoc, "#ID", $sID, $WdReplaceAll, -1, True, True)
			   Next
			   _Word_DocClose($oDoc, $WdSaveChanges)
			   _Excel_RangeWrite($oWorkbook, Default, $aEntry, "A" & $iRow)
			   If @error Then
				  $bWritingFailed = True
			   Else
				  If Not StringCompare($aEntry[0][2], "vertrouwelijk") Then $aEntry[0][2] = "_vertrouwelijk"
				  FileMove(@ScriptDir & "\" & $sLettersDir & $aLetters[$i][1] & ".doc", $sDatabase & "\" & StringLower($aEntry[0][18]) & "\" & @year & "\" & $sID & $aEntry[0][2] & ".doc", $FC_OVERWRITE + $FC_CREATEPATH)
				  FileMove(@ScriptDir & "\" & $sLettersDir & $aLetters[$i][1] & ".pdf", $sDatabase & "\" & StringLower($aEntry[0][18]) & "\" & @year & "\" & $sID & $aEntry[0][2] & ".pdf", $FC_OVERWRITE + $FC_CREATEPATH)
				  FileMove(@ScriptDir & "\" & $sLettersDir & $aLetters[$i][1] & "_aangetekend.doc", $sDatabase & "\" & StringLower($aEntry[0][18]) & "\" & @year & "\" & $sID & $aEntry[0][2] & "_aangetekend.doc", $FC_OVERWRITE + $FC_CREATEPATH)
				  FileDelete(@WorkingDir & "\" & $sLettersDir & $aLetters[$i][1] & ".txt")
				  $hTemp = FileOpen(@ScriptDir & "\geschiedenis.txt", $FO_APPEND + $FO_UNICODE)
					 FileWriteLine($hTemp, @MDAY & "/" & @MON & " " & @HOUR & ":" & @MIN & "|" & $aEntry[0][$aiMapping[0]] & "|" & $aEntry[0][$aiMapping[7]] & "|" & $aEntry[0][$aiMapping[1]] & "|" & $aEntry[0][$aiMapping[9]] & " " & $aEntry[0][$aiMapping[10]])
				  FileClose($hTemp)
			   EndIf
			Else
			   $bWritingFailed = True
			EndIf
		 EndIf
	  EndIf
   Next
   _Word_Quit($oWord)
   $oExcel.ActiveSheet.Columns("A:AC").AutoFit
   _Excel_BookSave($oWorkbook)
   _Excel_BookClose($oWorkbook, False)
   _Excel_Close($oExcel)
   FileDelete($sDatabase & "\bezet")
   If $bLettersMissing Then
	  MsgBox($MB_ICONERROR, "Database error", "De gegevens van een of meerdere brieven ontbreken, deze zijn daarom niet naar de centrale database verstuurd. Vraag hulp aan een systeembeheerder om dit probleem op te lossen.")
   EndIf
   If $bWritingFailed Then
	  MsgBox($MB_ICONERROR, "Database error", "Om een onbekende reden konden een of meerdere brieven niet aan de database worden toegevoegd. Vraag hulp aan een systeembeheerder om dit probleem op te lossen.")
   EndIf
   If $bCopyFailed Then
	  MsgBox($MB_ICONERROR, "Database error", "Een of meerdere brieven zijn wel aan het register toegevoed maar niet naar de database verplaatst. Vraag hulp aan een systeembeheerder om dit probleem op te lossen.")
   EndIf
   _WinMainLoadHistory()
EndFunc





; ################
; # DRAW LETTERS #
; ################
Func _DatabaseDrawLetters()
   For $i = 0 To UBound($asTabPrintShown) + UBound($asTabPrintHidden)
	  _GUICtrlListView_DeleteColumn($idDatabaseLetters, 0)
   Next
   For $i = 0 To UBound($asTabPrintShown) - 1
	  _GUICtrlListView_AddColumn($idDatabaseLetters, $asTabPrintShown[$i])
   Next
   For $i = 0 To UBound($asTabPrintShown) - 1
	  _GUICtrlListView_SetColumnWidth($idDatabaseLetters, $i, $LVSCW_AUTOSIZE_USEHEADER)
   Next
   For $i = 1 To $aDatabaseLetters[0][0]
	  $aDatabaseLetters[$i][2] = ""
	  For $j = 0 To UBound($asTabPrintShown) - 1
		 If Not StringCompare($asTabPrintShown[$j], "Registratienummer") Then $aDatabaseLetters[$i][2] = $aDatabaseLetters[$i][2] & $aDatabaseLettersInfo[$i][0] & "|"
		 If Not StringCompare($asTabPrintShown[$j], "Organisatie") Then $aDatabaseLetters[$i][2] = $aDatabaseLetters[$i][2] & $aDatabaseLettersInfo[$i][1] & "|"
		 If Not StringCompare($asTabPrintShown[$j], "Bestemmeling") Then $aDatabaseLetters[$i][2] = $aDatabaseLetters[$i][2] & $aDatabaseLettersInfo[$i][2] & "|"
		 If Not StringCompare($asTabPrintShown[$j], "Adres") Then $aDatabaseLetters[$i][2] = $aDatabaseLetters[$i][2] & $aDatabaseLettersInfo[$i][3] & "|"
		 If Not StringCompare($asTabPrintShown[$j], "Uw kenmerk") Then $aDatabaseLetters[$i][2] = $aDatabaseLetters[$i][2] & $aDatabaseLettersInfo[$i][4] & "|"
		 If Not StringCompare($asTabPrintShown[$j], "Gemeente/OCMW/AGB") Then $aDatabaseLetters[$i][2] = $aDatabaseLetters[$i][2] & $aDatabaseLettersInfo[$i][5] & "|"
		 If Not StringCompare($asTabPrintShown[$j], "Afzender") Then $aDatabaseLetters[$i][2] = $aDatabaseLetters[$i][2] & $aDatabaseLettersInfo[$i][6] & "|"
		 If Not StringCompare($asTabPrintShown[$j], "Dienst") Then $aDatabaseLetters[$i][2] = $aDatabaseLetters[$i][2] & $aDatabaseLettersInfo[$i][7] & "|"
		 If Not StringCompare($asTabPrintShown[$j], "Onderwerp") Then $aDatabaseLetters[$i][2] = $aDatabaseLetters[$i][2] & $aDatabaseLettersInfo[$i][8] & "|"
		 If Not StringCompare($asTabPrintShown[$j], "Datum") Then $aDatabaseLetters[$i][2] = $aDatabaseLetters[$i][2] & $aDatabaseLettersInfo[$i][9] & "|"
		 If Not StringCompare($asTabPrintShown[$j], "Ons kenmerk") Then $aDatabaseLetters[$i][2] = $aDatabaseLetters[$i][2] & $aDatabaseLettersInfo[$i][10] & "|"
		 If Not StringCompare($asTabPrintShown[$j], "Aangetekend") Then $aDatabaseLetters[$i][2] = $aDatabaseLetters[$i][2] & $aDatabaseLettersInfo[$i][11] & "|"
		 If Not StringCompare($asTabPrintShown[$j], "Vertrouwelijk") Then $aDatabaseLetters[$i][2] = $aDatabaseLetters[$i][2] & $aDatabaseLettersInfo[$i][12] & "|"
		 If Not StringCompare($asTabPrintShown[$j], "Handtekening bestuur") Then $aDatabaseLetters[$i][2] = $aDatabaseLetters[$i][2] & $aDatabaseLettersInfo[$i][13] & "|"
		 If Not StringCompare($asTabPrintShown[$j], "Bijlage") Then $aDatabaseLetters[$i][2] = $aDatabaseLetters[$i][2] & $aDatabaseLettersInfo[$i][14] & "|"
		 If Not StringCompare($asTabPrintShown[$j], "Opmerking") Then $aDatabaseLetters[$i][2] = $aDatabaseLetters[$i][2] & $aDatabaseLettersInfo[$i][15] & "|"
		 If Not StringCompare($asTabPrintShown[$j], "Pagina's") Then $aDatabaseLetters[$i][2] = $aDatabaseLetters[$i][2] & $aDatabaseLettersInfo[$i][16] & "|"
		 If Not StringCompare($asTabPrintShown[$j], "Afgedrukt") Then $aDatabaseLetters[$i][2] = $aDatabaseLetters[$i][2] & $aDatabaseLettersInfo[$i][17] & "|"
		 If Not StringCompare($asTabPrintShown[$j], "Ondertekend") Then $aDatabaseLetters[$i][2] = $aDatabaseLetters[$i][2] & $aDatabaseLettersInfo[$i][18] & "|"
		 If Not StringCompare($asTabPrintShown[$j], "Verzonden") Then $aDatabaseLetters[$i][2] = $aDatabaseLetters[$i][2] & $aDatabaseLettersInfo[$i][19] & "|"
		 Next
	  GUICtrlSetData($aDatabaseLetters[$i][0], $aDatabaseLetters[$i][2])
   Next
EndFunc





; ################
; # LOAD LETTERS #
; ################
Func _DatabaseLoadLetters()
   If FileExists($sDatabase & "\bezet") Then
	  MsgBox($MB_ICONINFORMATION, "Even geduld", "De database is momenteel door iemand anders in gebruik, probeer het over enkele seconden opnieuw." & @CRLF & @CRLF & "Indien u meer dan een paar minuten moet wachten, vraag een systeembeheerder om hulp.")
	  Return False
   EndIf
   _FileCreate($sDatabase & "\bezet")
   $oExcel = _Excel_Open(False)
   If @error Then
	  FileDelete($sDatabase & "\bezet")
	  If $bError Then MsgBox($MB_ICONERROR, "Database error", "Het registerbestand kon niet worden geopend. Controleer of Microsoft Excel 2003 of hoger is geïnstalleerd en correct werkt.")
	  Return False
   EndIf
   $asRegisters = _FileListToArray($sDatabase, "register.*", $FLTA_FILES)
   If @error Then
	  If @error <> 4 Then
		 FileDelete($sDatabase & "\bezet")
		 _Excel_Close($oExcel, False)
		 MsgBox($MB_ICONERROR, "Database error", "Het register kon niet gevonden worden. Suggesties:" & _
			@CRLF & @CRLF & "- Controleer of de juiste map is ingesteld." & _
			@CRLF & "- Controleer of de map bereikbaar is." & _
			@CRLF & "- Controleer of het programma de juiste rechten heeft." & _
			@CRLF & "- Probeer het later opnieuw.")
		 Return False
	  Else
		 _Excel_Close($oExcel, False)
		 ReDim $aDatabaseLetters[1][7]
		 $aDatabaseLetters[0][0] = 0
		 Return True
	  EndIf
   Else
	  $oWorkbook = _Excel_BookOpen($oExcel, $sDatabase & "\" & $asRegisters[1], Default, Default, True)
	  If @error Then
		 FileDelete($sDatabase & "\bezet")
		 _Excel_Close($oExcel, False)
		 MsgBox($MB_ICONERROR, "Database error", "Het programma kan het register niet openen. Suggesties:" & _
			@CRLF & @CRLF & "- Controleer of de juiste map is ingesteld." & _
			@CRLF & "- Controleer of de map bereikbaar is." & _
			@CRLF & "- Controleer of het programma de juiste rechten heeft." & _
			@CRLF & "- Controleer of Microsoft Excel 2003 of hoger is geïnstalleerd en correct werkt." & _
			@CRLF & "- Probeer het later opnieuw.")
		 Return False
	  EndIf
   EndIf
   $oSheet = $oWorkbook.Worksheets(1)
   $oSheet.Activate
   $oRange = $oSheet.UsedRange
   $oRange.SpecialCells($xlCellTypeLastCell).Activate
   $iMin = 2
   $iMax = $oExcel.ActiveCell.Row
   $sFoundID = ""
   $sSearchID = GUICtrlRead($idDatabaseDate)
   If StringLen($sSearchID) = 9 Then $sSearchID = "0" & $sSearchID
   $sSearchID = StringMid($sSearchID, 7, 4) & StringMid($sSearchID, 4, 2) & StringMid($sSearchID, 1, 2) & "0001"
   Local $asHeading[][] = [["Registratienummer", _
							"Onderwerp", _
							"Vertrouwelijk", _
							"Aangetekend", _
							"Handtekening bestuur", _
							"Bijlage", _
							"Opmerking", _
							"Organisatie bestemmeling", _
							"Aanspreking", _
							"Voornaam bestemmeling", _
							"Achternaam bestemmeling", _
							"Straatnaam", _
							"Huisnummer", _
							"Bus", _
							"Postcode", _
							"Gemeente", _
							"Land", _
							"Uw kenmerk", _
							"Organisatie afzender", _
							"Voornaam afzender", _
							"Achternaam afzender", _
							"Dienst", _
							"Telefoon", _
							"E-mailadres", _
							"Datum", _
							"Ons kenmerk", _
							"Afgedrukt", _
							"Ondertekend", _
							"Verzonden"]]
   $asHeader = _Excel_RangeRead($oWorkbook, Default, "A1:AC1")
   Local $aiMapping[29] = [-1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1]
   For $i = 0 To UBound($asHeading, $UBOUND_COLUMNS) - 1
	  For $j = 0 To UBound($asHeader, $UBOUND_COLUMNS) - 1
		 If Not StringCompare($asHeading[0][$j], $asHeader[0][$i]) Then
			$aiMapping[$i] = $j
			$asHeading[0][$i] = "-"
		 EndIf
	  Next
   Next
   For $i = 0 To UBound($aiMapping) - 1
	  If $aiMapping[$i] < 0 Then
		 For $j = 0 To UBound($aiMapping) - 1
			 $bFound = False
			 For $k = 0 To UBound($aiMapping) - 1
				 If $aiMapping[$k] =  $j then
					 $bFound = True
					 Exitloop
				 Endif
			 Next
			 If Not $bFound Then
				 $aiMapping[$i] = $j
				 ExitLoop
			 Endif
		 Next
	  EndIf
   Next
   If $aiMapping[0] <= 26 Then
	  $sCol = Chr(65 + $aiMapping[0])
   Else
	  $sCol = "A" & Chr(39 + $aiMapping[0])
   EndIf
   While $iMin < $iMax And StringCompare($sFoundID, $sSearchID)
	  $iPos = $iMin + Floor(($iMax - $iMin + 1) / 2)
	  $sFoundID = _Excel_RangeRead($oWorkbook, Default, $sCol & $iPos)
	  If StringCompare($sFoundID, $sSearchID) > 0 Or Not StringCompare($sFoundID, "") Then
		 $iMax = $iPos - 1
	  ElseIf StringCompare($sFoundID, $sSearchID) < 0 Then
		 $iMin = $iPos + 1
	  Else
		 ExitLoop
	  EndIf
   WEnd
   $iPos = $iMin + Floor(($iMax - $iMin + 1) / 2)
   $sFoundID = _Excel_RangeRead($oWorkbook, Default, $sCol & $iPos)
   $sSearchID = StringTrimRight($sSearchID, 4)
   For $i = 1 To $aDatabaseLetters[0][0]
	  GUICtrlDelete($aDatabaseLetters[$i][0])
   Next
   ReDim $aDatabaseLetters[1][7]
   $aDatabaseLetters[0][0] = 0
   ReDim $aDatabaseLettersInfo[1][20]
   $sFoundID = _Excel_RangeRead($oWorkbook, Default, $sCol & $iPos)
   While Not StringCompare(StringLeft($sFoundID, 8), $sSearchID)
	  $aDatabaseLetters[0][0] += 1
	  ReDim $aDatabaseLetters[$aDatabaseLetters[0][0] + 1][7]
	  ReDim $aDatabaseLettersInfo[$aDatabaseLetters[0][0] + 1][20]
	  $aDatabaseLetters[$aDatabaseLetters[0][0]][1] = $sFoundID
	  $asCurrentEntry = _Excel_RangeRead($oWorkbook, Default, "A" & $iPos & ":AC" & $iPos)
	  If Not @error Then
		 Local $asTempEntry[1][UBound($aiMapping)]
		 For $i = 0 To UBound($aiMapping) - 1
			$asTempEntry[0][$aiMapping[$i]] = $asCurrentEntry[0][$i]
		 Next
		 $asCurrentEntry = $asTempEntry
		 If Not StringCompare($asCurrentEntry[0][2], "vertrouwelijk") Then $aDatabaseLetters[$aDatabaseLetters[0][0]][1] &= "_vertrouwelijk"
		 $aDatabaseLetters[$aDatabaseLetters[0][0]][1] = $asCurrentEntry[0][18] & "\" & StringLeft($sFoundID, 4) & "\" & $aDatabaseLetters[$aDatabaseLetters[0][0]][1]
		 $aDatabaseLetters[$aDatabaseLetters[0][0]][2] = ""
		 $aDatabaseLetters[$aDatabaseLetters[0][0]][0] = GUICtrlCreateListViewItem("", $idDatabaseLetters)
		 $aDatabaseLetters[$aDatabaseLetters[0][0]][3] = StringCompare($asCurrentEntry[0][2], "") ? True : False  ; vertrouwelijk
		 $aDatabaseLetters[$aDatabaseLetters[0][0]][4] = StringCompare($asCurrentEntry[0][3], "") ? True : False  ; aangetekend
		 $aDatabaseLetters[$aDatabaseLetters[0][0]][5] = StringCompare($asCurrentEntry[0][4], "") ? True : False  ; handtekening bestuur
		 $aDatabaseLetters[$aDatabaseLetters[0][0]][6] = StringCompare($asCurrentEntry[0][5], "0") ? True : False ; bijlage
		 If StringCompare($asCurrentEntry[0][13], "") Then $asCurrentEntry[0][13] = ", bus " & $asCurrentEntry[0][13]
		 If Not StringCompare($asCurrentEntry[0][16], "België") Then
			$asCurrentEntry[0][16] = ""
		 Else
			$asCurrentEntry[0][16] = ", " & $asCurrentEntry[0][16]
		 EndIf
		 $asCurrentEntry[0][24] = StringMid($asCurrentEntry[0][24], 7, 2) & "/" & StringMid($asCurrentEntry[0][24], 5, 2) & "/" & StringMid($asCurrentEntry[0][24], 1, 4)
		 $aDatabaseLettersInfo[$aDatabaseLetters[0][0]][0] = $asCurrentEntry[0][0]																	;  0 Registratienummer
		 $aDatabaseLettersInfo[$aDatabaseLetters[0][0]][1] = $asCurrentEntry[0][7]																	;  1 Organisatie
		 $aDatabaseLettersInfo[$aDatabaseLetters[0][0]][2] = $asCurrentEntry[0][9] & " " & $asCurrentEntry[0][10]									;  2 Bestemmeling
		 $aDatabaseLettersInfo[$aDatabaseLetters[0][0]][3] = $asCurrentEntry[0][11] & " " & $asCurrentEntry[0][12] & $asCurrentEntry[0][13] & _
															 ", " & $asCurrentEntry[0][14] & " " & $asCurrentEntry[0][15] & $asCurrentEntry[0][16]	;  3 Adres
		 $aDatabaseLettersInfo[$aDatabaseLetters[0][0]][4] = $asCurrentEntry[0][17]																	;  4 Uw kenmerk
		 $aDatabaseLettersInfo[$aDatabaseLetters[0][0]][5] = $asCurrentEntry[0][18]																	;  5 Gemeente/OCMW/AGB
		 $aDatabaseLettersInfo[$aDatabaseLetters[0][0]][6] = $asCurrentEntry[0][19] & " " & $asCurrentEntry[0][20]									;  6 Afzender
		 $aDatabaseLettersInfo[$aDatabaseLetters[0][0]][7] = $asCurrentEntry[0][21]																	;  7 Dienst
		 $aDatabaseLettersInfo[$aDatabaseLetters[0][0]][8] = $asCurrentEntry[0][1]																	;  8 Onderwerp
		 $aDatabaseLettersInfo[$aDatabaseLetters[0][0]][9] = $asCurrentEntry[0][24]																	;  9 Datum
		 $aDatabaseLettersInfo[$aDatabaseLetters[0][0]][10] = $asCurrentEntry[0][25]																; 10 Ons kenmerk
		 $aDatabaseLettersInfo[$aDatabaseLetters[0][0]][11] = $asCurrentEntry[0][3]																	; 11 Aangetekend
		 $aDatabaseLettersInfo[$aDatabaseLetters[0][0]][12] = $asCurrentEntry[0][2]																	; 12 Vertrouwelijk
		 $aDatabaseLettersInfo[$aDatabaseLetters[0][0]][13] = $asCurrentEntry[0][4]																	; 13 Handtekening bestuur
		 $aDatabaseLettersInfo[$aDatabaseLetters[0][0]][14] = $asCurrentEntry[0][5]																	; 14 Bijlage
		 $aDatabaseLettersInfo[$aDatabaseLetters[0][0]][15] = $asCurrentEntry[0][6]																	; 15 Opmerking
		 $aDatabaseLettersInfo[$aDatabaseLetters[0][0]][16] = Int($asCurrentEntry[0][5]) + 1														; 16 Pagina's
		 $aDatabaseLettersInfo[$aDatabaseLetters[0][0]][17] = $asCurrentEntry[0][26]																; 17 Afgedrukt
		 $aDatabaseLettersInfo[$aDatabaseLetters[0][0]][18] = $asCurrentEntry[0][27]																; 18 Ondertekend
		 $aDatabaseLettersInfo[$aDatabaseLetters[0][0]][19] = $asCurrentEntry[0][28]																; 19 Verzonden
	  EndIf
	  $iPos += 1
	  $sFoundID = _Excel_RangeRead($oWorkbook, Default, $sCol & $iPos)
   WEnd
   _Excel_BookClose($oWorkbook, False)
   _Excel_Close($oExcel)
   FileDelete($sDatabase & "\bezet")
   _DatabaseDrawLetters()
EndFunc





; ################
; # SET PRINTED #
; ################
Func _DatabaseSetPrinted()
   If FileExists($sDatabase & "\bezet") Then
	  MsgBox($MB_ICONINFORMATION, "Even geduld", "De database is momenteel door iemand anders in gebruik, probeer het over enkele seconden opnieuw." & @CRLF & @CRLF & "Indien u meer dan een paar minuten moet wachten, vraag een systeembeheerder om hulp.")
	  Return False
   EndIf
   _FileCreate($sDatabase & "\bezet")
   $oExcel = _Excel_Open(False)
   If @error Then
	  FileDelete($sDatabase & "\bezet")
	  If $bError Then MsgBox($MB_ICONERROR, "Database error", "Het registerbestand kon niet worden geopend. Controleer of Microsoft Excel 2003 of hoger is geïnstalleerd en correct werkt.")
	  Return False
   EndIf
   $asRegisters = _FileListToArray($sDatabase, "register.*", $FLTA_FILES)
   If @error Then
	  If @error <> 4 Then
		 FileDelete($sDatabase & "\bezet")
		 _Excel_Close($oExcel, False)
		 MsgBox($MB_ICONERROR, "Database error", "Het register kon niet gevonden worden. Suggesties:" & _
			@CRLF & @CRLF & "- Controleer of de juiste map is ingesteld." & _
			@CRLF & "- Controleer of de map bereikbaar is." & _
			@CRLF & "- Controleer of het programma de juiste rechten heeft." & _
			@CRLF & "- Probeer het later opnieuw.")
		 Return False
	  Else
		 _Excel_Close($oExcel, False)
		 ReDim $aDatabaseLetters[1][7]
		 $aDatabaseLetters[0][0] = 0
		 Return True
	  EndIf
   Else
	  $oWorkbook = _Excel_BookOpen($oExcel, $sDatabase & "\" & $asRegisters[1], Default, Default, True)
	  If @error Then
		 FileDelete($sDatabase & "\bezet")
		 _Excel_Close($oExcel, False)
		 MsgBox($MB_ICONERROR, "Database error", "Het programma kan het register niet openen. Suggesties:" & _
			@CRLF & @CRLF & "- Controleer of de juiste map is ingesteld." & _
			@CRLF & "- Controleer of de map bereikbaar is." & _
			@CRLF & "- Controleer of het programma de juiste rechten heeft." & _
			@CRLF & "- Controleer of Microsoft Excel 2003 of hoger is geïnstalleerd en correct werkt." & _
			@CRLF & "- Probeer het later opnieuw.")
		 Return False
	  EndIf
   EndIf
   $oSheet = $oWorkbook.Worksheets(1)
   $oSheet.Activate
   $oRange = $oSheet.UsedRange
   $oRange.SpecialCells($xlCellTypeLastCell).Activate
   $iMin = 2
   $iMax = $oExcel.ActiveCell.Row
   $sFoundID = ""
   $sSearchID = GUICtrlRead($idDatabaseDate)
   If StringLen($sSearchID) = 9 Then $sSearchID = "0" & $sSearchID
   $sSearchID = StringMid($sSearchID, 7, 4) & StringMid($sSearchID, 4, 2) & StringMid($sSearchID, 1, 2) & "0001"
   Local $asHeading[][] = [["Registratienummer", _
							"Onderwerp", _
							"Vertrouwelijk", _
							"Aangetekend", _
							"Handtekening bestuur", _
							"Bijlage", _
							"Opmerking", _
							"Organisatie bestemmeling", _
							"Aanspreking", _
							"Voornaam bestemmeling", _
							"Achternaam bestemmeling", _
							"Straatnaam", _
							"Huisnummer", _
							"Bus", _
							"Postcode", _
							"Gemeente", _
							"Land", _
							"Uw kenmerk", _
							"Organisatie afzender", _
							"Voornaam afzender", _
							"Achternaam afzender", _
							"Dienst", _
							"Telefoon", _
							"E-mailadres", _
							"Datum", _
							"Ons kenmerk", _
							"Afgedrukt", _
							"Ondertekend", _
							"Verzonden"]]
   $asHeader = _Excel_RangeRead($oWorkbook, Default, "A1:AC1")
   Local $aiMapping[29] = [-1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1]
   For $i = 0 To UBound($asHeading, $UBOUND_COLUMNS) - 1
	  For $j = 0 To UBound($asHeader, $UBOUND_COLUMNS) - 1
		 If Not StringCompare($asHeading[0][$j], $asHeader[0][$i]) Then
			$aiMapping[$i] = $j
			$asHeading[0][$i] = "-"
		 EndIf
	  Next
   Next
   For $i = 0 To UBound($aiMapping) - 1
	  If $aiMapping[$i] < 0 Then
		 For $j = 0 To UBound($aiMapping) - 1
			 $bFound = False
			 For $k = 0 To UBound($aiMapping) - 1
				 If $aiMapping[$k] =  $j then
					 $bFound = True
					 Exitloop
				 Endif
			 Next
			 If Not $bFound Then
				 $aiMapping[$i] = $j
				 ExitLoop
			 Endif
		 Next
	  EndIf
   Next
   If $aiMapping[0] <= 26 Then
	  $sColID = Chr(65 + $aiMapping[0])
   Else
	  $sColID = "A" & Chr(39 + $aiMapping[0])
   EndIf
   If $aiMapping[26] < 26 Then
	  $sColPrinted = Chr(65 + $aiMapping[26])
   Else
	  $sColPrinted = "A" & Chr(39 + $aiMapping[26])
   EndIf
   While $iMin < $iMax And StringCompare($sFoundID, $sSearchID)
	  $iPos = $iMin + Floor(($iMax - $iMin + 1) / 2)
	  $sFoundID = _Excel_RangeRead($oWorkbook, Default, $sColID & $iPos)
	  If StringCompare($sFoundID, $sSearchID) > 0 Or Not StringCompare($sFoundID, "") Then
		 $iMax = $iPos - 1
	  ElseIf StringCompare($sFoundID, $sSearchID) < 0 Then
		 $iMin = $iPos + 1
	  Else
		 ExitLoop
	  EndIf
   WEnd
   $iPos = $iMin + Floor(($iMax - $iMin + 1) / 2)
   $sSearchID = StringTrimRight($sSearchID, 4)
   $sFoundID = _Excel_RangeRead($oWorkbook, Default, $sColID & $iPos)
   While Not StringCompare(StringLeft($sFoundID, 8), $sSearchID)
	  $asCurrentEntry = _Excel_RangeRead($oWorkbook, Default, "A" & $iPos & ":AC" & $iPos)
	  For $i = 1 To $aDatabaseLetters[0][0]
		 If BitAND(GUICtrlRead($aDatabaseLetters[$i][0], 1), $GUI_CHECKED) And Not StringCompare($asCurrentEntry[0][$aiMapping[0]], $aDatabaseLettersInfo[$i][0]) Then
			_Excel_RangeWrite($oWorkbook, Default, "Ja", $sColPrinted & $iPos)
			$aDatabaseLettersInfo[$i][17] = "Ja"
			ExitLoop
		 EndIf
	  Next
	  $iPos += 1
	  $sFoundID = _Excel_RangeRead($oWorkbook, Default, $sColID & $iPos)
   WEnd
   _Excel_BookClose($oWorkbook, True)
   _Excel_Close($oExcel)
   FileDelete($sDatabase & "\bezet")
   _DatabaseDrawLetters()
EndFunc





; ############
; # SET SENT #
; ############
Func _DatabaseSetSent()
   If FileExists($sDatabase & "\bezet") Then
	  MsgBox($MB_ICONINFORMATION, "Even geduld", "De database is momenteel door iemand anders in gebruik, probeer het over enkele seconden opnieuw." & @CRLF & @CRLF & "Indien u meer dan een paar minuten moet wachten, vraag een systeembeheerder om hulp.")
	  Return False
   EndIf
   _FileCreate($sDatabase & "\bezet")
   $oExcel = _Excel_Open(False)
   If @error Then
	  FileDelete($sDatabase & "\bezet")
	  If $bError Then MsgBox($MB_ICONERROR, "Database error", "Het registerbestand kon niet worden geopend. Controleer of Microsoft Excel 2003 of hoger is geïnstalleerd en correct werkt.")
	  Return False
   EndIf
   $asRegisters = _FileListToArray($sDatabase, "register.*", $FLTA_FILES)
   If @error Then
	  If @error <> 4 Then
		 FileDelete($sDatabase & "\bezet")
		 _Excel_Close($oExcel, False)
		 MsgBox($MB_ICONERROR, "Database error", "Het register kon niet gevonden worden. Suggesties:" & _
			@CRLF & @CRLF & "- Controleer of de juiste map is ingesteld." & _
			@CRLF & "- Controleer of de map bereikbaar is." & _
			@CRLF & "- Controleer of het programma de juiste rechten heeft." & _
			@CRLF & "- Probeer het later opnieuw.")
		 Return False
	  Else
		 _Excel_Close($oExcel, False)
		 ReDim $aDatabaseLetters[1][7]
		 $aDatabaseLetters[0][0] = 0
		 Return True
	  EndIf
   Else
	  $oWorkbook = _Excel_BookOpen($oExcel, $sDatabase & "\" & $asRegisters[1], Default, Default, True)
	  If @error Then
		 FileDelete($sDatabase & "\bezet")
		 _Excel_Close($oExcel, False)
		 MsgBox($MB_ICONERROR, "Database error", "Het programma kan het register niet openen. Suggesties:" & _
			@CRLF & @CRLF & "- Controleer of de juiste map is ingesteld." & _
			@CRLF & "- Controleer of de map bereikbaar is." & _
			@CRLF & "- Controleer of het programma de juiste rechten heeft." & _
			@CRLF & "- Controleer of Microsoft Excel 2003 of hoger is geïnstalleerd en correct werkt." & _
			@CRLF & "- Probeer het later opnieuw.")
		 Return False
	  EndIf
   EndIf
   $oSheet = $oWorkbook.Worksheets(1)
   $oSheet.Activate
   $oRange = $oSheet.UsedRange
   $oRange.SpecialCells($xlCellTypeLastCell).Activate
   $iMin = 2
   $iMax = $oExcel.ActiveCell.Row
   $sFoundID = ""
   $sSearchID = GUICtrlRead($idDatabaseDate)
   If StringLen($sSearchID) = 9 Then $sSearchID = "0" & $sSearchID
   $sSearchID = StringMid($sSearchID, 7, 4) & StringMid($sSearchID, 4, 2) & StringMid($sSearchID, 1, 2) & "0001"
   Local $asHeading[][] = [["Registratienummer", _
							"Onderwerp", _
							"Vertrouwelijk", _
							"Aangetekend", _
							"Handtekening bestuur", _
							"Bijlage", _
							"Opmerking", _
							"Organisatie bestemmeling", _
							"Aanspreking", _
							"Voornaam bestemmeling", _
							"Achternaam bestemmeling", _
							"Straatnaam", _
							"Huisnummer", _
							"Bus", _
							"Postcode", _
							"Gemeente", _
							"Land", _
							"Uw kenmerk", _
							"Organisatie afzender", _
							"Voornaam afzender", _
							"Achternaam afzender", _
							"Dienst", _
							"Telefoon", _
							"E-mailadres", _
							"Datum", _
							"Ons kenmerk", _
							"Afgedrukt", _
							"Ondertekend", _
							"Verzonden"]]
   $asHeader = _Excel_RangeRead($oWorkbook, Default, "A1:AC1")
   Local $aiMapping[29] = [-1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1]
   For $i = 0 To UBound($asHeading, $UBOUND_COLUMNS) - 1
	  For $j = 0 To UBound($asHeader, $UBOUND_COLUMNS) - 1
		 If Not StringCompare($asHeading[0][$j], $asHeader[0][$i]) Then
			$aiMapping[$i] = $j
			$asHeading[0][$i] = "-"
		 EndIf
	  Next
   Next
   For $i = 0 To UBound($aiMapping) - 1
	  If $aiMapping[$i] < 0 Then
		 For $j = 0 To UBound($aiMapping) - 1
			 $bFound = False
			 For $k = 0 To UBound($aiMapping) - 1
				 If $aiMapping[$k] =  $j then
					 $bFound = True
					 Exitloop
				 Endif
			 Next
			 If Not $bFound Then
				 $aiMapping[$i] = $j
				 ExitLoop
			 Endif
		 Next
	  EndIf
   Next
   If $aiMapping[0] <= 26 Then
	  $sColID = Chr(65 + $aiMapping[0])
   Else
	  $sColID = "A" & Chr(39 + $aiMapping[0])
   EndIf
   If $aiMapping[28] < 26 Then
	  $sColSent = Chr(65 + $aiMapping[28])
   Else
	  $sColSent = "A" & Chr(39 + $aiMapping[28])
   EndIf
   While $iMin < $iMax And StringCompare($sFoundID, $sSearchID)
	  $iPos = $iMin + Floor(($iMax - $iMin + 1) / 2)
	  $sFoundID = _Excel_RangeRead($oWorkbook, Default, $sColID & $iPos)
	  If StringCompare($sFoundID, $sSearchID) > 0 Or Not StringCompare($sFoundID, "") Then
		 $iMax = $iPos - 1
	  ElseIf StringCompare($sFoundID, $sSearchID) < 0 Then
		 $iMin = $iPos + 1
	  Else
		 ExitLoop
	  EndIf
   WEnd
   $iPos = $iMin + Floor(($iMax - $iMin + 1) / 2)
   $sSearchID = StringTrimRight($sSearchID, 4)
   $sFoundID = _Excel_RangeRead($oWorkbook, Default, $sColID & $iPos)
   While Not StringCompare(StringLeft($sFoundID, 8), $sSearchID)
	  $asCurrentEntry = _Excel_RangeRead($oWorkbook, Default, "A" & $iPos & ":AC" & $iPos)
	  For $i = 1 To $aDatabaseLetters[0][0]
		 If BitAND(GUICtrlRead($aDatabaseLetters[$i][0], 1), $GUI_CHECKED) And Not StringCompare($asCurrentEntry[0][$aiMapping[0]], $aDatabaseLettersInfo[$i][0]) Then
			_Excel_RangeWrite($oWorkbook, Default, "Ja", $sColSent & $iPos)
			$aDatabaseLettersInfo[$i][19] = "Ja"
			ExitLoop
		 EndIf
	  Next
	  $iPos += 1
	  $sFoundID = _Excel_RangeRead($oWorkbook, Default, $sColID & $iPos)
   WEnd
   _Excel_BookClose($oWorkbook, True)
   _Excel_Close($oExcel)
   FileDelete($sDatabase & "\bezet")
   _DatabaseDrawLetters()
EndFunc





; ##############
; # SET SIGNED #
; ##############
Func _DatabaseSetSigned()
   If FileExists($sDatabase & "\bezet") Then
	  MsgBox($MB_ICONINFORMATION, "Even geduld", "De database is momenteel door iemand anders in gebruik, probeer het over enkele seconden opnieuw." & @CRLF & @CRLF & "Indien u meer dan een paar minuten moet wachten, vraag een systeembeheerder om hulp.")
	  Return False
   EndIf
   _FileCreate($sDatabase & "\bezet")
   $oExcel = _Excel_Open(False)
   If @error Then
	  FileDelete($sDatabase & "\bezet")
	  If $bError Then MsgBox($MB_ICONERROR, "Database error", "Het registerbestand kon niet worden geopend. Controleer of Microsoft Excel 2003 of hoger is geïnstalleerd en correct werkt.")
	  Return False
   EndIf
   $asRegisters = _FileListToArray($sDatabase, "register.*", $FLTA_FILES)
   If @error Then
	  If @error <> 4 Then
		 FileDelete($sDatabase & "\bezet")
		 _Excel_Close($oExcel, False)
		 MsgBox($MB_ICONERROR, "Database error", "Het register kon niet gevonden worden. Suggesties:" & _
			@CRLF & @CRLF & "- Controleer of de juiste map is ingesteld." & _
			@CRLF & "- Controleer of de map bereikbaar is." & _
			@CRLF & "- Controleer of het programma de juiste rechten heeft." & _
			@CRLF & "- Probeer het later opnieuw.")
		 Return False
	  Else
		 _Excel_Close($oExcel, False)
		 ReDim $aDatabaseLetters[1][7]
		 $aDatabaseLetters[0][0] = 0
		 Return True
	  EndIf
   Else
	  $oWorkbook = _Excel_BookOpen($oExcel, $sDatabase & "\" & $asRegisters[1], Default, Default, True)
	  If @error Then
		 FileDelete($sDatabase & "\bezet")
		 _Excel_Close($oExcel, False)
		 MsgBox($MB_ICONERROR, "Database error", "Het programma kan het register niet openen. Suggesties:" & _
			@CRLF & @CRLF & "- Controleer of de juiste map is ingesteld." & _
			@CRLF & "- Controleer of de map bereikbaar is." & _
			@CRLF & "- Controleer of het programma de juiste rechten heeft." & _
			@CRLF & "- Controleer of Microsoft Excel 2003 of hoger is geïnstalleerd en correct werkt." & _
			@CRLF & "- Probeer het later opnieuw.")
		 Return False
	  EndIf
   EndIf
   $oSheet = $oWorkbook.Worksheets(1)
   $oSheet.Activate
   $oRange = $oSheet.UsedRange
   $oRange.SpecialCells($xlCellTypeLastCell).Activate
   $iMin = 2
   $iMax = $oExcel.ActiveCell.Row
   $sFoundID = ""
   $sSearchID = GUICtrlRead($idDatabaseDate)
   If StringLen($sSearchID) = 9 Then $sSearchID = "0" & $sSearchID
   $sSearchID = StringMid($sSearchID, 7, 4) & StringMid($sSearchID, 4, 2) & StringMid($sSearchID, 1, 2) & "0001"
   Local $asHeading[][] = [["Registratienummer", _
							"Onderwerp", _
							"Vertrouwelijk", _
							"Aangetekend", _
							"Handtekening bestuur", _
							"Bijlage", _
							"Opmerking", _
							"Organisatie bestemmeling", _
							"Aanspreking", _
							"Voornaam bestemmeling", _
							"Achternaam bestemmeling", _
							"Straatnaam", _
							"Huisnummer", _
							"Bus", _
							"Postcode", _
							"Gemeente", _
							"Land", _
							"Uw kenmerk", _
							"Organisatie afzender", _
							"Voornaam afzender", _
							"Achternaam afzender", _
							"Dienst", _
							"Telefoon", _
							"E-mailadres", _
							"Datum", _
							"Ons kenmerk", _
							"Afgedrukt", _
							"Ondertekend", _
							"Verzonden"]]
   $asHeader = _Excel_RangeRead($oWorkbook, Default, "A1:AC1")
   Local $aiMapping[29] = [-1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1]
   For $i = 0 To UBound($asHeading, $UBOUND_COLUMNS) - 1
	  For $j = 0 To UBound($asHeader, $UBOUND_COLUMNS) - 1
		 If Not StringCompare($asHeading[0][$j], $asHeader[0][$i]) Then
			$aiMapping[$i] = $j
			$asHeading[0][$i] = "-"
		 EndIf
	  Next
   Next
   For $i = 0 To UBound($aiMapping) - 1
	  If $aiMapping[$i] < 0 Then
		 For $j = 0 To UBound($aiMapping) - 1
			 $bFound = False
			 For $k = 0 To UBound($aiMapping) - 1
				 If $aiMapping[$k] =  $j then
					 $bFound = True
					 Exitloop
				 Endif
			 Next
			 If Not $bFound Then
				 $aiMapping[$i] = $j
				 ExitLoop
			 Endif
		 Next
	  EndIf
   Next
   If $aiMapping[0] < 26 Then
	  $sColID = Chr(65 + $aiMapping[0])
   Else
	  $sColID = "A" & Chr(39 + $aiMapping[0])
   EndIf
   If $aiMapping[27] <= 26 Then
	  $sColSigned = Chr(65 + $aiMapping[27])
   Else
	  $sColSigned = "A" & Chr(39 + $aiMapping[27])
   EndIf
   While $iMin < $iMax And StringCompare($sFoundID, $sSearchID)
	  $iPos = $iMin + Floor(($iMax - $iMin + 1) / 2)
	  $sFoundID = _Excel_RangeRead($oWorkbook, Default, $sColID & $iPos)
	  If StringCompare($sFoundID, $sSearchID) > 0 Or Not StringCompare($sFoundID, "") Then
		 $iMax = $iPos - 1
	  ElseIf StringCompare($sFoundID, $sSearchID) < 0 Then
		 $iMin = $iPos + 1
	  Else
		 ExitLoop
	  EndIf
   WEnd
   $iPos = $iMin + Floor(($iMax - $iMin + 1) / 2)
   $sSearchID = StringTrimRight($sSearchID, 4)
   $sFoundID = _Excel_RangeRead($oWorkbook, Default, $sColID & $iPos)
   While Not StringCompare(StringLeft($sFoundID, 8), $sSearchID)
	  $asCurrentEntry = _Excel_RangeRead($oWorkbook, Default, "A" & $iPos & ":AC" & $iPos)
	  For $i = 1 To $aDatabaseLetters[0][0]
		 If BitAND(GUICtrlRead($aDatabaseLetters[$i][0], 1), $GUI_CHECKED) And Not StringCompare($asCurrentEntry[0][$aiMapping[0]], $aDatabaseLettersInfo[$i][0]) Then
			If Not StringCompare($aDatabaseLettersInfo[$i][18], "Nee") Then
			   _Excel_RangeWrite($oWorkbook, Default, "Ja", $sColSigned & $iPos)
			   $aDatabaseLettersInfo[$i][18] = "Ja"
			EndIf
			ExitLoop
		 EndIf
	  Next
	  $iPos += 1
	  $sFoundID = _Excel_RangeRead($oWorkbook, Default, $sColID & $iPos)
   WEnd
   _Excel_BookClose($oWorkbook, True)
   _Excel_Close($oExcel)
   FileDelete($sDatabase & "\bezet")
   _DatabaseDrawLetters()
EndFunc





; ################
; # OPEN LETTERS #
; ################
Func _DatabaseOpenLetters()
   $oWord = _Word_Create()
   If @error Then
	  MsgBox(48, "Waarschuwing", "De geselecteerde brieven konden niet worden geopend. Controleer of Microsoft Word 2003 of hoger is geïnstalleerd en correct werkt.")
	  return False
   EndIf
   $bError = False
   For $i = 1 To $aDatabaseLetters[0][0]
	  If BitAND(GUICtrlRead($aDatabaseLetters[$i][0], 1), $GUI_CHECKED) Then
		 $oDoc = _Word_DocOpen($oWord, $sDatabase & "\" & $aDatabaseLetters[$i][1] & ".doc")
		 If @error Then
			$bError = True
		 Else
			$oDoc.Activewindow.View.Type = 3
		 EndIf
		 If $aDatabaseLetters[$i][6] Then Run(@comspec & ' /c start ' & $sDatabase & "\" & $aDatabaseLetters[$i][1] & ".pdf", "", @SW_HIDE)
		 If $aDatabaseLetters[$i][4] Then $oDoc = _Word_DocOpen($oWord, $sDatabase & "\" & $aDatabaseLetters[$i][1] & "_aangetekend.doc")
		 If @error Then
			$bError = True
		 Else
			$oDoc.Activewindow.View.Type = 3
		 EndIf
	  EndIf
   Next
   If $bError Then MsgBox(48, "Waarschuwing", "Een of meerdere brieven konden niet worden geopend. Mogelijke oorzaken:" & @CRLF & @CRLF & "- Microsoft Word 2003 of hoger is niet geïnstalleerd of werkt niet correct." & @CRLF & "- De bestanden zijn beschadigd of ontbreken.")
EndFunc