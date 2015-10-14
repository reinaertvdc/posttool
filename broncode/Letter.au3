; ###############
; # OPEN LETTER #
; ###############
Func _LetterOpen($iID)
   $oWord = _Word_Create()
   If @error Then Return False
   $oDoc = _Word_DocOpen($oWord, $iID & ".doc")
   If @error Then
	  _Word_Quit($oWord)
	  Return False
   Else
	  $oDoc.Activewindow.View.Type = 3
   EndIf
   Return True
EndFunc





; #################
; # REMOVE LETTER #
; #################
Func _LetterRemove($iID)
   FileDelete(@WorkingDir & "\" & $sLettersDir & $iID & ".doc")
   FileDelete(@WorkingDir & "\" & $sLettersDir & $iID & ".txt")
   FileDelete(@WorkingDir & "\" & $sLettersDir & $iID & ".pdf")
EndFunc





; ##############
; # ADD LETTER #
; ##############
Func _LetterAdd()
   $iID = 1

   While FileExists($sLettersDir & $iID & ".txt")
	  $iID += 1
   WEnd

   If Not StringCompare($sFromOrg, "AGB") Then
	  $sTemplate = "agb"
	  $sRegistered = "agb"
   ElseIf Not StringCompare($sFromOrg, "OCMW") Then
	  If Not StringCompare($sFromDepartment, "woonzorgcentrum") Then
		 $sTemplate = "wzc"
		 $sRegistered = "ocmw"
	  Else
		 $sTemplate = "ocmw"
		 $sRegistered = "ocmw"
	  EndIf
   Else
	  If Not StringCompare($sFromDepartment, "dienst toerisme") Then
		 $sTemplate = "toerisme"
		 $sRegistered = "gemeente"
	  Else
		 $sTemplate = "gemeente"
		 $sRegistered = "gemeente"
	  EndIf
   EndIf

   $bSuccess = False
   If $bOptionsSigned Then
	  $bSuccess = FileCopy("sjablonen\" & $sTemplate & "_ondertekend.doc", $sLettersDir, $FC_OVERWRITE + $FC_CREATEPATH)
	  If $bOptionsRegistered Then $bSuccess = _Min($bSuccess, FileCopy("sjablonen\" & $sRegistered & "_aangetekend.doc", $sLettersDir, $FC_OVERWRITE + $FC_CREATEPATH))
	  If $bSuccess Then
		 $bSuccess = _Min($bSuccess, FileMove($sLettersDir & $sTemplate & "_ondertekend.doc", $sLettersDir & $iID & ".doc", $FC_OVERWRITE + $FC_CREATEPATH))
		 If $bOptionsRegistered Then $bSuccess = _Min($bSuccess, FileMove($sLettersDir & $sRegistered & "_aangetekend.doc", $sLettersDir & $iID & "_aangetekend.doc", $FC_OVERWRITE + $FC_CREATEPATH))
		 If Not $bSuccess Then
			FileDelete($sLettersDir & $iID & ".doc")
			FileDelete($sLettersDir & $iID & "_aangetekend.doc")
		 EndIf
	  EndIf
   Else
	  $bSuccess = FileCopy("sjablonen\" & $sTemplate & ".doc", $sLettersDir, $FC_OVERWRITE + $FC_CREATEPATH)
	  If $bOptionsRegistered Then $bSuccess = _Min($bSuccess, FileCopy("sjablonen\" & $sRegistered & "_aangetekend.doc", $sLettersDir, $FC_OVERWRITE + $FC_CREATEPATH))
	  If $bSuccess Then
		 $bSuccess = FileMove($sLettersDir & $sTemplate & ".doc", $sLettersDir & $iID & ".doc", $FC_OVERWRITE + $FC_CREATEPATH)
		 If $bOptionsRegistered Then $bSuccess = _Min($bSuccess, FileMove($sLettersDir & $sRegistered & "_aangetekend.doc", $sLettersDir & $iID & "_aangetekend.doc", $FC_OVERWRITE + $FC_CREATEPATH))
		 If Not $bSuccess Then
			FileDelete($sLettersDir & $iID & ".doc")
			FileDelete($sLettersDir & $iID & "_aangetekend.doc")
		 EndIf
	  EndIf
   EndIf
   If Not $bSuccess Then
	  MsgBox(48, "Waarschuwing", "Het sjabloon kon niet worden gekopieerd. Mogelijke oorzaken:" & @CRLF & @CRLF & "- Het sjabloon ontbreekt." & @CRLF & "- Het programma heeft onvoldoende rechten." & @CRLF & "- De harde schijf is vol.")
	  Return False
   Else
	  $iAnnex = 0
	  If StringCompare($sOptionsAnnex, "") Then
		 $bSuccess = FileCopy($sOptionsAnnex, $sLettersDir, $FC_OVERWRITE + $FC_CREATEPATH)
		 If $bSuccess Then
			$bSuccess = FileMove($sLettersDir & StringMid($sOptionsAnnex, StringInStr($sOptionsAnnex, "\", 0, -1) + 1, StringLen($sOptionsAnnex) - StringInStr($sOptionsAnnex, "\", 0, -1)), $sLettersDir & $iID & ".pdf", $FC_OVERWRITE + $FC_CREATEPATH)
			If Not $bSuccess Then FileDelete($sLettersDir & StringMid($sOptionsAnnex, StringInStr($sOptionsAnnex, "\", 0, -1) + 1, StringLen($sOptionsAnnex) - StringInStr($sOptionsAnnex, "\", 0, -1)))
		 EndIf
		 If Not $bSuccess Then
			MsgBox(48, "Waarschuwing", "De bijlage kon niet worden toegevoegd. Mogelijke oorzaken:" & @CRLF & @CRLF & "- De bijlage ontbreekt, heeft een nieuwe naam of is verplaatst." & @CRLF & "- De harde schijf is vol.")
			FileDelete($sLettersDir & $iID & ".doc")
			Return False
		 EndIf
		 $c = 0
		 $iAnnex = -1
		 ;'OPEN the PDF file.
		 ;$hPDF = FileOpen($sLettersDir & $iID & ".pdf", 0)
		 ;If @ERROR Then
			;MsgBox(48, "Waarschuwing", "De bijlage kon niet worden geopend. Controleer of de bijlage een geldig en onbeschadigd PDF bestand is.")
			;FileDelete($sLettersDir & $iID & ".doc")
			;FileDelete($sLettersDir & $iID & ".pdf")
			;return False
		 ;EndIf
		 $iAnnex = 1
		 #CS
		 ; ###############################
		 ; # WE CAN DO BETTER THAN THIS! #
		 ; ###############################
		 ;Get the data from the file.
		 Do
			$s = FileReadLine($hPDF)
			$iEOF = @ERROR
			$c += 1
			;Look within the top 10 lines for /N.
			If $c <= 10 Then
			   If StringInStr($s, "/N") > 0 Then
				  $array1 = StringRegExp($s, "/N\h(\d+)", 3)
				  If @error Then
					 $iAnnex = 1
				  Else
					 $iAnnex = Number($array1[0])
				  EndIf
				  ExitLoop
			   EndIf
			EndIf
			;Check every line for /count.
			If StringInStr($s, "/count") > 0 Then
			   $array2 = StringRegExp($s, "/count\h(\d+)", 3)
			   If @error Then
				  MsgBox(0, "", @error)
				  $iAnnex = 1
			   Else
				  $iAnnex = Number($array2[0])
			   EndIf
			   ExitLoop
			EndIf
		 Until $iEOF = -1
		 FileClose($hPDF)
		 If $iAnnex < 0 Then $iAnnex = 1
		 #CE
		 ;FileClose($hPDF)
	  EndIf

	  $hFile = FileOpen($sLettersDir & $iID & ".txt", $FO_OVERWRITE + $FO_CREATEPATH + $FO_UNICODE)
	  If $hFile = -1 Then
		 FileDelete($sLettersDir & $iID & ".doc")
		 FileDelete($sLettersDir & $iID & ".pdf")
		 MsgBox(48, "Waarschuwing", "Een systeembestand kon niet worden aangemaakt. Controleer of de harde schijf niet vol is.")
		 Return False
	  EndIf
		 FileWriteLine($hFile, $sFromSubject)
		 FileWriteLine($hFile, $bOptionsConfidential)
		 FileWriteLine($hFile, $bOptionsRegistered)
		 FileWriteLine($hFile, $bOptionsSigned)
		 FileWriteLine($hFile, $iAnnex)
		 FileWriteLine($hFile, $sOptionsRemark)
		 FileWriteLine($hFile, $sToOrg)
		 FileWriteLine($hFile, $sToSalution)
		 FileWriteLine($hFile, $sToFirstName)
		 FileWriteLine($hFile, $sToSurname)
		 FileWriteLine($hFile, $sToStreet)
		 FileWriteLine($hFile, $sToNumber)
		 FileWriteLine($hFile, $sToBus)
		 FileWriteLine($hFile, $sToPostcode)
		 FileWriteLine($hFile, $sToCity)
		 FileWriteLine($hFile, $sToCountry)
		 If Not StringCompare($sToCountry, "België") Then $sToCountry = ""
		 FileWriteLine($hFile, $sToRef)
		 FileWriteLine($hFile, $sFromOrg)
		 FileWriteLine($hFile, $sFromFirstName)
		 FileWriteLine($hFile, $sFromSurname)
		 FileWriteLine($hFile, $sFromDepartment)
		 FileWriteLine($hFile, $sFromTel)
		 FileWriteLine($hFile, $sFromEmail)
		 FileWriteLine($hFile, $sFromDate)
		 FileWriteLine($hFile, $sFromRef)
	  FileClose($hFile)

	  $oWord = _Word_Create(False)
	  If @error Then
		 FileDelete($sLettersDir & $iID & ".doc")
		 FileDelete($sLettersDir & $iID & ".pdf")
		 FileDelete($sLettersDir & $iID & ".txt")
		 MsgBox(48, "Waarschuwing", "Het sjabloon kon niet worden geopend. Controleer of Microsoft Word 2003 of hoger is geïnstalleerd en correct werkt.")
		 Return False
	  EndIf
	  $oDoc = _Word_DocOpen($oWord, @WorkingDir & "\" & $sLettersDir & $iID & ".doc")
	  If @error Then
		 _Word_Quit($oWord)
		 FileDelete($sLettersDir & $iID & ".doc")
		 FileDelete($sLettersDir & $iID & ".pdf")
		 FileDelete($sLettersDir & $iID & ".txt")
		 If $bError Then MsgBox(48, "Waarschuwing", "Het sjabloon kon niet worden geopend. Mogelijke oorzaken:" & @CRLF & @CRLF & "- Microsoft Word 2003 of hoger is niet geïnstalleerd of werkt niet correct." & @CRLF & "- Het bestand is beschadigd of ontbreken.")
		 Return False
	  EndIf

	  ; find and replace in body
	  $oDoc.Range.Select()
	  _Word_DocFindReplace($oDoc, "#ToOrg", $sToOrg, $WdReplaceAll, 0, True, True)
	  _Word_DocFindReplace($oDoc, "#ToSalution", $sToSalution, $WdReplaceAll, 0, True, True)
	  _Word_DocFindReplace($oDoc, "#ToFirstName", $sToFirstName, $WdReplaceAll, 0, True, True)
	  _Word_DocFindReplace($oDoc, "#ToSurname", $sToSurname, $WdReplaceAll, 0, True, True)
	  _Word_DocFindReplace($oDoc, "#ToStreet", $sToStreet, $WdReplaceAll, 0, True, True)
	  _Word_DocFindReplace($oDoc, "#ToNumber", $sToNumber, $WdReplaceAll, 0, True, True)
	  _Word_DocFindReplace($oDoc, "#ToBus", $sToBus, $WdReplaceAll, 0, True, True)
	  _Word_DocFindReplace($oDoc, "#ToPostcode", $sToPostcode, $WdReplaceAll, 0, True, True)
	  _Word_DocFindReplace($oDoc, "#ToCity", $sToCity, $WdReplaceAll, 0, True, True)
	  _Word_DocFindReplace($oDoc, "#ToCountry", $sToCountry, $WdReplaceAll, 0, True, True)
	  _Word_DocFindReplace($oDoc, "#ToRef", $sToRef, $WdReplaceAll, 0, True, True)
	  _Word_DocFindReplace($oDoc, "#FromOrg", $sFromOrg, $WdReplaceAll, 0, True, True)
	  _Word_DocFindReplace($oDoc, "#FromFirstName", $sFromFirstName, $WdReplaceAll, 0, True, True)
	  _Word_DocFindReplace($oDoc, "#FromSurname", $sFromSurname, $WdReplaceAll, 0, True, True)
	  _Word_DocFindReplace($oDoc, "#FromDepartment", $sFromDepartment, $WdReplaceAll, 0, True, True)
	  _Word_DocFindReplace($oDoc, "#FromTel", $sFromTel, $WdReplaceAll, 0, True, True)
	  _Word_DocFindReplace($oDoc, "#FromEmail", $sFromEmail, $WdReplaceAll, 0, True, True)
	  _Word_DocFindReplace($oDoc, "#FromSubject", $sFromSubject, $WdReplaceAll, 0, True, True)
	  _Word_DocFindReplace($oDoc, "#FromDate", $sFromDate, $WdReplaceAll, 0, True, True)
	  _Word_DocFindReplace($oDoc, "#FromRef", $sFromRef, $WdReplaceAll, 0, True, True)
	  _Word_DocFindReplace($oDoc, "#Annex", $iAnnex, $WdReplaceAll, 0, True, True)
	  If $bOptionsRegistered Then
		 $sOptionsRegistered = "Aangetekend"
	  Else
		 $sOptionsRegistered = ""
	  EndIf
	  _Word_DocFindReplace($oDoc, "#OptionsRegistered", $sOptionsRegistered, $WdReplaceAll, 0, True, True)

	  $count_sections=$oDoc.Sections.Count()
	  For $sec_i=1 To $count_sections
		 ; select one header section
		 $oDoc.Sections($sec_i).Headers(1).Range.Select()
		 ; replace in header
		 _Word_DocFindReplace($oDoc, "#FromDepartment", $sFromDepartment, $WdReplaceAll, -1, True, True)
	  Next
	  _Word_DocClose($oDoc, $WdSaveChanges)

	  If $bOptionsRegistered Then
		 $oDoc = _Word_DocOpen($oWord, @WorkingDir & "\" & $sLettersDir & $iID & "_aangetekend.doc")
		 If @error Then
			_Word_Quit($oWord)
			Return False
		 EndIf

		 ; find and replace in body
		 $oDoc.Range.Select()
		 _Word_DocFindReplace($oDoc, "#ToOrg", $sToOrg, $WdReplaceAll, 0, True, True)
		 _Word_DocFindReplace($oDoc, "#ToSalution", $sToSalution, $WdReplaceAll, 0, True, True)
		 _Word_DocFindReplace($oDoc, "#ToFirstName", $sToFirstName, $WdReplaceAll, 0, True, True)
		 _Word_DocFindReplace($oDoc, "#ToSurname", $sToSurname, $WdReplaceAll, 0, True, True)
		 _Word_DocFindReplace($oDoc, "#ToStreet", $sToStreet, $WdReplaceAll, 0, True, True)
		 _Word_DocFindReplace($oDoc, "#ToNumber", $sToNumber, $WdReplaceAll, 0, True, True)
		 _Word_DocFindReplace($oDoc, "#ToBus", $sToBus, $WdReplaceAll, 0, True, True)
		 _Word_DocFindReplace($oDoc, "#ToPostcode", $sToPostcode, $WdReplaceAll, 0, True, True)
		 _Word_DocFindReplace($oDoc, "#ToCity", $sToCity, $WdReplaceAll, 0, True, True)
		 _Word_DocFindReplace($oDoc, "#ToCountry", $sToCountry, $WdReplaceAll, 0, True, True)
		 _Word_DocFindReplace($oDoc, "#ToRef", $sToRef, $WdReplaceAll, 0, True, True)
		 _Word_DocFindReplace($oDoc, "#FromOrg", $sFromOrg, $WdReplaceAll, 0, True, True)
		 _Word_DocFindReplace($oDoc, "#FromFirstName", $sFromFirstName, $WdReplaceAll, 0, True, True)
		 _Word_DocFindReplace($oDoc, "#FromSurname", $sFromSurname, $WdReplaceAll, 0, True, True)
		 _Word_DocFindReplace($oDoc, "#FromDepartment", $sFromDepartment, $WdReplaceAll, 0, True, True)
		 _Word_DocFindReplace($oDoc, "#FromTel", $sFromTel, $WdReplaceAll, 0, True, True)
		 _Word_DocFindReplace($oDoc, "#FromEmail", $sFromEmail, $WdReplaceAll, 0, True, True)
		 _Word_DocFindReplace($oDoc, "#FromSubject", $sFromSubject, $WdReplaceAll, 0, True, True)
		 _Word_DocFindReplace($oDoc, "#FromDate", $sFromDate, $WdReplaceAll, 0, True, True)
		 _Word_DocFindReplace($oDoc, "#FromRef", $sFromRef, $WdReplaceAll, 0, True, True)
		 _Word_DocFindReplace($oDoc, "#Annex", $iAnnex, $WdReplaceAll, 0, True, True)
		 If $bOptionsRegistered Then
			$sOptionsRegistered = "Aangetekend"
		 Else
			$sOptionsRegistered = ""
		 EndIf
		 _Word_DocFindReplace($oDoc, "#OptionsRegistered", $sOptionsRegistered, $WdReplaceAll, 0, True, True)

		 $count_sections=$oDoc.Sections.Count()
		 For $sec_i=1 To $count_sections
			; select one header section
			$oDoc.Sections($sec_i).Headers(1).Range.Select()
			; replace in header
			_Word_DocFindReplace($oDoc, "#FromDepartment", $sFromDepartment, $WdReplaceAll, -1, True, True)
		 Next
		 _Word_DocClose($oDoc, $WdSaveChanges)
	  EndIf

	  _Word_Quit($oWord)
	  _WinMainLoadLetters()
	  If Not _LetterOpen(@WorkingDir & "\" & $sLettersDir & $iID) Then MsgBox(48, "Waarschuwing", "De brief is aangemaakt maar kon niet automatisch worden geopend. Mogelijke oorzaken:" & @CRLF & @CRLF & "- Microsoft Word werkt niet correct." & @CRLF & "- Het bestand is beschadigd of ontbreekt.")
   EndIf
   return True
EndFunc