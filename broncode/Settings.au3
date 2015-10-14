; #################
; # LOAD SETTINGS #
; #################
Func _LoadSettings()
   ; Name of the file to read settings from.
   Global $sSettingsFile = "instellingen.ini"
   ; Load organisations.
   Global $sOrgs = "Gemeente|OCMW|AGB"
   ; Read settings from file. If a setting is missing, fill in the default value and use that.
   _ReadSettings()
   ; Load the postcodes and their corresponding areas.
   _LoadPostcodes()
   ; Load the countries.
   _LoadCountries()
EndFunc





; #################
; # READ SETTINGS #
; #################
Func _ReadSettings()
   ; Read database location.
   Global $sDatabase = _ArrayToString(IniReadSection($sSettingsFile, "Database"), "", 1, 0, "|", 1, 1)
   If @error Or Not StringCompare($sDatabase, "") Then
	  MsgBox($MB_ICONWARNING, "Database locatie", "Geen database locatie gevonden. Selecteer een map in het volgende menu.")
	  $sDatabase = FileSelectFolder("Database locatie selecteren", "")
	  IniWriteSection($sSettingsFile, "Database", "    =" & $sDatabase)
   EndIf
   $sDatabase = StringInStr($sDatabase, "|")? StringLeft($sDatabase, StringInStr($sDatabase, "|") - 1) : $sDatabase
   ; Check for newer version of program.
   If Not FileExists($sDatabase & "\patches") Then
	  DirCreate($sDatabase & "\patches")
	  FileSetAttrib($sDatabase & "\patches", "+H")
   EndIf
   If @Compiled Then
	  If FileExists(@ScriptDir & "\~" & @ScriptName) Then
		 FileDelete(@ScriptDir & "\~" & @ScriptName)
	  ElseIf FileExists($sDatabase & "\patches\Verwerking uitgaande post.exe") And _StringCompareVersions(FileGetVersion($sDatabase & "\patches\Verwerking uitgaande post.exe", "ProductVersion"), $sVersion) > 0 Then
		 FileMove(@ScriptFullPath, @ScriptDir & "\~" & @ScriptName)
		 FileCopy($sDatabase & "\patches\Verwerking uitgaande post.exe", @ScriptFullPath)
		 Run(@ScriptFullPath)
		 Exit 0
	  EndIf
   EndIf
   ; Check for new templates.
   If StringCompare(FileGetTime($sDatabase & "\patches\wzc_ondertekend.doc", $FT_MODIFIED, 1), FileGetTime("sjablonen\wzc_ondertekend.doc", $FT_MODIFIED, 1)) > 0 Then
	  FileDelete("sjablonen\wzc_ondertekend.doc")
	  FileCopy($sDatabase & "\patches\wzc_ondertekend.doc", "sjablonen\wzc_ondertekend.doc")
   EndIf
   If StringCompare(FileGetTime($sDatabase & "\patches\wzc.doc", $FT_MODIFIED, 1), FileGetTime("sjablonen\wzc.doc", $FT_MODIFIED, 1)) > 0 Then
	  FileDelete("sjablonen\wzc.doc")
	  FileCopy($sDatabase & "\patches\wzc.doc", "sjablonen\wzc.doc")
   EndIf
   If StringCompare(FileGetTime($sDatabase & "\patches\toerisme_ondertekend.doc", $FT_MODIFIED, 1), FileGetTime("sjablonen\toerisme_ondertekend.doc", $FT_MODIFIED, 1)) > 0 Then
	  FileDelete("sjablonen\toerisme_ondertekend.doc")
	  FileCopy($sDatabase & "\patches\toerisme_ondertekend.doc", "sjablonen\toerisme_ondertekend.doc")
   EndIf
   If StringCompare(FileGetTime($sDatabase & "\patches\toerisme.doc", $FT_MODIFIED, 1), FileGetTime("sjablonen\toerisme.doc", $FT_MODIFIED, 1)) > 0 Then
	  FileDelete("sjablonen\toerisme.doc")
	  FileCopy($sDatabase & "\patches\toerisme.doc", "sjablonen\toerisme.doc")
   EndIf
   If StringCompare(FileGetTime($sDatabase & "\patches\ocmw_ondertekend.doc", $FT_MODIFIED, 1), FileGetTime("sjablonen\ocmw_ondertekend.doc", $FT_MODIFIED, 1)) > 0 Then
	  FileDelete("sjablonen\ocmw_ondertekend.doc")
	  FileCopy($sDatabase & "\patches\ocmw_ondertekend.doc", "sjablonen\ocmw_ondertekend.doc")
   EndIf
   If StringCompare(FileGetTime($sDatabase & "\patches\ocmw_aangetekend.doc", $FT_MODIFIED, 1), FileGetTime("sjablonen\ocmw_aangetekend.doc", $FT_MODIFIED, 1)) > 0 Then
	  FileDelete("sjablonen\ocmw_aangetekend.doc")
	  FileCopy($sDatabase & "\patches\ocmw_aangetekend.doc", "sjablonen\ocmw_aangetekend.doc")
   EndIf
   If StringCompare(FileGetTime($sDatabase & "\patches\ocmw.doc", $FT_MODIFIED, 1), FileGetTime("sjablonen\ocmw.doc", $FT_MODIFIED, 1)) > 0 Then
	  FileDelete("sjablonen\ocmw.doc")
	  FileCopy($sDatabase & "\patches\ocmw.doc", "sjablonen\ocmw.doc")
   EndIf
   If StringCompare(FileGetTime($sDatabase & "\patches\gemeente_ondertekend.doc", $FT_MODIFIED, 1), FileGetTime("sjablonen\gemeente_ondertekend.doc", $FT_MODIFIED, 1)) > 0 Then
	  FileDelete("sjablonen\gemeente_ondertekend.doc")
	  FileCopy($sDatabase & "\patches\gemeente_ondertekend.doc", "sjablonen\gemeente_ondertekend.doc")
   EndIf
   If StringCompare(FileGetTime($sDatabase & "\patches\gemeente_aangetekend.doc", $FT_MODIFIED, 1), FileGetTime("sjablonen\gemeente_aangetekend.doc", $FT_MODIFIED, 1)) > 0 Then
	  FileDelete("sjablonen\gemeente_aangetekend.doc")
	  FileCopy($sDatabase & "\patches\gemeente_aangetekend.doc", "sjablonen\gemeente_aangetekend.doc")
   EndIf
   If StringCompare(FileGetTime($sDatabase & "\patches\gemeente.doc", $FT_MODIFIED, 1), FileGetTime("sjablonen\gemeente.doc", $FT_MODIFIED, 1)) > 0 Then
	  FileDelete("sjablonen\gemeente.doc")
	  FileCopy($sDatabase & "\patches\gemeente.doc", "sjablonen\gemeente.doc")
   EndIf
   If StringCompare(FileGetTime($sDatabase & "\patches\agb_ondertekend.doc", $FT_MODIFIED, 1), FileGetTime("sjablonen\agb_ondertekend.doc", $FT_MODIFIED, 1)) > 0 Then
	  FileDelete("sjablonen\agb_ondertekend.doc")
	  FileCopy($sDatabase & "\patches\agb_ondertekend.doc", "sjablonen\agb_ondertekend.doc")
   EndIf
   If StringCompare(FileGetTime($sDatabase & "\patches\agb_aangetekend.doc", $FT_MODIFIED, 1), FileGetTime("sjablonen\agb_aangetekend.doc", $FT_MODIFIED, 1)) > 0 Then
	  FileDelete("sjablonen\agb_aangetekend.doc")
	  FileCopy($sDatabase & "\patches\agb_aangetekend.doc", "sjablonen\agb_aangetekend.doc")
   EndIf
   If StringCompare(FileGetTime($sDatabase & "\patches\agb.doc", $FT_MODIFIED, 1), FileGetTime("sjablonen\agb.doc", $FT_MODIFIED, 1)) > 0 Then
	  FileDelete("sjablonen\agb.doc")
	  FileCopy($sDatabase & "\patches\agb.doc", "sjablonen\agb.doc")
   EndIf
   ; Read print.
   Global $sPrint = _ArrayToString(IniReadSection($sSettingsFile, "Print"), "", 1, 0, "|", 1, 1)
   If @error Or Not StringCompare($sPrint, "") Then
	  $sPrint = "Nee"
	  IniWriteSection($sSettingsFile, "Print", "    =" & $sPrint)
   EndIf
   $sPrint = StringInStr($sPrint, "|")? StringLeft($sPrint, StringInStr($sPrint, "|") - 1) : $sPrint
   ; Read pause between prints.
   Global $iPrintPause = _ArrayToString(IniReadSection($sSettingsFile, "PrintPauze"), "", 1, 0, "|", 1, 1)
   If @error Or Not StringCompare($iPrintPause, "") Then
	  $iPrintPause = "3000"
	  IniWriteSection($sSettingsFile, "PrintPauze", "    =" & $iPrintPause)
   EndIf
   $iPrintPause = Int(StringInStr($iPrintPause, "|")? StringLeft($iPrintPause, StringInStr($iPrintPause, "|") - 1) : $iPrintPause)
   ; Read font size.
   Global $sFontSize = _ArrayToString(IniReadSection($sSettingsFile, "Lettergrootte"), "", 1, 0, "|", 1, 1)
   If @error Or Not StringCompare($sFontSize, "") Then
	  $sFontSize = "Normaal"
	  IniWriteSection($sSettingsFile, "Lettergrootte", "    =" & $sFontSize)
   EndIf
   $sFontSize = StringInStr($sFontSize, "|")? StringLeft($sFontSize, StringInStr($sFontSize, "|") - 1) : $sFontSize
   ; Read color theme.
   ;Global $sTheme = IniRead($sSettingsFile, "Beeld", "Thema", "")
   ;If Not StringCompare($sTheme, "") Then
	  ;$sTheme = "Standaard"
	  ;IniWrite($sSettingsFile, "Beeld", "Thema", $sTheme)
   ;EndIf
   ; Read salutions.
   Global $sSalutions = _ArrayToString(IniReadSection($sSettingsFile, "Aansprekingen"), "", 1, 0, "|", 1, 1)
   If @error Or Not StringCompare($sSalutions, "") Then
	  $sSalutions = "geachte heer|geachte mevrouw|geachte"
	  IniWriteSection($sSettingsFile, "Aansprekingen", "    =" & StringReplace($sSalutions, "|", @LF & "    ="))
   EndIf
   ; Read 'Gemeente' departments.
   Global $sGemeente = _ArrayToString(IniReadSection($sSettingsFile, "Gemeente"), "", 1, 0, "|", 1, 1)
   If @error Or Not StringCompare($sGemeente, "") Then
	  $sGemeente = "archief|" & _
				   "burgemeester|" & _
				   "burgerzaken|" & _
				   "college van burgemeester en schepenen|" & _
				   "communicatiedienst|" & _
				   "cultuurdienst|" & _
				   "de kinderclub|" & _
				   "dienst financiën|" & _
				   "dienst ict|" & _
				   "dienst lokale economie|" & _
				   "dienst onderwijs|" & _
				   "dienst ruimtelijke ordening|" & _
				   "dienst toerisme|" & _
				   "financieel beheerder|" & _
				   "gemeenteraad|" & _
				   "jeugddienst|" & _
				   "kopiedienst|" & _
				   "milieudienst|" & _
				   "personeelsdienst|" & _
				   "preventiedienst|" & _
				   "secretariaat|" & _
				   "secretaris gemeente en ocmw|" & _
				   "sportdienst|" & _
				   "technisch centrum|" & _
				   "technische dienst|" & _
				   "technische dienst mobiliteit|" & _
				   "technische dienst overheidsopdrachten|" & _
				   "werelddienst"
	  IniWriteSection($sSettingsFile, "Gemeente", "    =" & StringReplace($sGemeente, "|", @LF & "    ="))
   EndIf
   ; Read 'OCMW' departments.
   Global $sOCMW = _ArrayToString(IniReadSection($sSettingsFile, "OCMW"), "", 1, 0, "|", 1, 1)
   If @error Or Not StringCompare($sOCMW, "") Then
	  $sOCMW = "de dreef|" & _
			   "dienst financiën|" & _
			   "dienst rechtshulp|" & _
			   "facilitaire dienst|" & _
			   "financieel beheerder|" & _
			   "lokaal dienstencentrum TerHarte|" & _
			   "personeelsdienst|" & _
			   "secretariaat|" & _
			   "secretaris gemeente en ocmw|" & _
			   "sociale dienst|" & _
			   "technische dienst|" & _
			   "woonzorgcentrum"
	  IniWriteSection($sSettingsFile, "OCMW", "    =" & StringReplace($sOCMW, "|", @LF & "    ="))
   EndIf
   ; Read 'AGB' departments.
   Global $sAGB = _ArrayToString(IniReadSection($sSettingsFile, "AGB"), "", 1, 0, "|", 1, 1)
   If @error Or Not StringCompare($sAGB, "") Then
	  $sAGB = "directiecomité|" & _
			  "raad van bestuur"
	  IniWriteSection($sSettingsFile, "AGB", "    =" & StringReplace($sAGB, "|", @LF & "    ="))
   EndIf
   #cs
   ; Read 'Gemeente' abbreviations.
   Global $asGemAbbr = IniReadSection($sSettingsFile, "AfkortingenGemeente")
   If @error Then
	  Global $asGemAbbr[][] = [["", ""], _
							   ["archief", "DOC"], _
							   ["burgemeester", "CBS"], _
							   ["burgerzaken", "BUR"], _
							   ["college van burgemeester en schepenen", "CBS"], _
							   ["communicatiedienst", "COM"], _
							   ["cultuurdienst", "CUL"], _
							   ["de kinderclub", "BKO"], _
							   ["dienst financiën", "FIN"], _
							   ["dienst ict", "ICT"], _
							   ["dienst lokale economie", "COM"], _
							   ["dienst onderwijs", "OND"], _
							   ["dienst ruimtelijke ordening", "ROL"], _
							   ["dienst toerisme", "TOE"], _
							   ["financieel beheerder", "FIN"], _
							   ["gemeenteraad", "GR"], _
							   ["jeugddienst", "CUL"], _
							   ["kopiedienst", "DOC"], _
							   ["milieudienst", "ROL"], _
							   ["personeelsdienst", "PER"], _
							   ["preventiedienst", "PRE"], _
							   ["secretariaat", "SEC"], _
							   ["secretaris gemeente en ocmw", "SEC"], _
							   ["sportdienst", "SPO"], _
							   ["technisch centrum", "TD"], _
							   ["technische dienst", "TD"], _
							   ["technische dienst mobiliteit", "MOB"], _
							   ["technische dienst overheidsopdrachten", "OO"], _
							   ["werelddienst", "CUL"]]
	  IniWriteSection($sSettingsFile, "AfkortingenGemeente", $asGemAbbr)
   EndIf
   ; Read 'OCMW' abbreviations.
   Global $asGemAbbr = IniReadSection($sSettingsFile, "AfkortingenOCMW")
   If @error Then
	  Global $asOCMWAbbr[][] = [["", ""], _
							    ["de dreef", "PHD"], _
							    ["dienst financiën", "FIN"], _
							    ["dienst rechtshulp", "JUR"], _
							    ["facilitaire dienst", "FAC"], _
							    ["financieel beheerder", "FIN"], _
							    ["lokaal dienstencentrum TerHarte", "LDC"], _
							    ["personeelsdienst", "PER"], _
							    ["secretariaat", "SEC"], _
							    ["secretaris gemeente en ocmw", "SEC"], _
							    ["sociale dienst", "SOC"], _
							    ["technische dienst", "TD"], _
							    ["woonzorgcentrum", "WZC"]]
	  IniWriteSection($sSettingsFile, "AfkortingenOCMW", $asOCMWAbbr)
   EndIf
   ; Read 'AGB' abbreviations.
   Global $asGemAbbr = IniReadSection($sSettingsFile, "AfkortingenAGB")
   If @error Then
	  Global $asAGBAbbr[][] = [["", ""], _
							   ["directiecomité", "DIR"], _
							   ["raad van bestuur", "RVB"]]
	  IniWriteSection($sSettingsFile, "AfkortingenAGB", $asAGBAbbr)
   EndIf
   #ce
   ; Read Print tab columns.
   Local $asTemp = IniReadSection($sSettingsFile, "PrintTabKolommen")
   If @error Then
	  Global $asTabPrintShown = ["Registratienummer", _
								 "Pagina's", _
								 "Handtekening bestuur", _
								 "Opmerking", _
								 "Afgedrukt", _
								 "Ondertekend", _
								 "Verzonden"]
	  Global $asTabPrintHidden = ["Organisatie", _
								  "Bestemmeling", _
								  "Adres", _
								  "Uw kenmerk", _
								  "Gemeente/OCMW/AGB", _
								  "Afzender", _
								  "Dienst", _
								  "Onderwerp", _
								  "Datum", _
								  "Ons kenmerk", _
								  "Aangetekend", _
								  "Vertrouwelijk", _
								  "Bijlage"]
	  IniWriteSection($sSettingsFile, "PrintTabKolommen", "    =" & StringReplace(_ArrayToString($asTabPrintShown), "|", @LF & "    ="))
   Else
	  $iTemp = UBound($asTemp) - 1
	  Global $asTabPrintShown[$iTemp]
	  For $i = 1 To $iTemp
		 $asTabPrintShown[$i - 1] = $asTemp[$i][1]
	  Next
	  Global $asTabPrintHidden = ["Registratienummer", _
								  "Organisatie", _
								  "Bestemmeling", _
								  "Adres", _
								  "Uw kenmerk", _
								  "Gemeente/OCMW/AGB", _
								  "Afzender", _
								  "Dienst", _
								  "Onderwerp", _
								  "Datum", _
								  "Ons kenmerk", _
								  "Aangetekend", _
								  "Vertrouwelijk", _
								  "Handtekening bestuur", _
								  "Bijlage", _
								  "Opmerking", _
								  "Pagina's", _
								  "Afgedrukt", _
								  "Ondertekend", _
								  "Verzonden"]
	  $iTabPrintShown = UBound($asTabPrintShown) - 1
	  $iTabPrintHidden = UBound($asTabPrintHidden) - 1
	  $i = 0
	  While $i <= $iTabPrintShown
		 $bMatch = False
		 $j = 0
		 While $j <= $iTabPrintHidden
			If Not StringCompare($asTabPrintShown[$i], $asTabPrintHidden[$j]) Then
			   $asTabPrintShown[$i] = $asTabPrintHidden[$j]
			   For $k = $j + 1 To $iTabPrintHidden
				  $asTabPrintHidden[$k - 1] = $asTabPrintHidden[$k]
			   Next
			   ReDim $asTabPrintHidden[$iTabPrintHidden]
			   $iTabPrintHidden -= 1
			   $bMatch = True
			   ExitLoop
			Else
			   $j += 1
			EndIf
		 WEnd
		 If Not $bMatch Then
			For $k = $i + 1 To $iTabPrintShown
			   $asTabPrintShown[$k - 1] = $asTabPrintShown[$k]
			Next
			ReDim $asTabPrintShown[$iTabPrintShown]
			$iTabPrintShown -= 1
		 Else
			$i += 1
		 EndIf
	  WEnd
   EndIf
EndFunc





Func _StringCompareVersions($s_Version1, $s_Version2 = "0.0.0.0")

; Confirm strings are of correct basic format. Set @error to 1,2 or 3 if not.
    SetError((StringIsDigit(StringReplace($s_Version1, ".", ""))=0) + 2 * (StringIsDigit(StringReplace($s_Version2, ".", ""))=0))
    If @error>0 Then Return 0; Ought to Return something!

    Local $i_Index, $i_Result, $ai_Version1, $ai_Version2

; Split into arrays by the "." separator
    $ai_Version1 = StringSplit($s_Version1, ".")
    $ai_Version2 = StringSplit($s_Version2, ".")
    $i_Result = 0; Assume strings are equal

; Ensure strings are of the same (correct) format:
;  Short strings are padded with 0s. Extraneous components of long strings are ignored. Values are Int.
    If $ai_Version1[0] <> 4 Then ReDim $ai_Version1[5]
    For $i_Index = 1 To 4
        $ai_Version1[$i_Index] = Int($ai_Version1[$i_Index])
    Next

    If $ai_Version2[0] <> 4 Then ReDim $ai_Version2[5]
    For $i_Index = 1 To 4
        $ai_Version2[$i_Index] = Int($ai_Version2[$i_Index])
    Next

    For $i_Index = 1 To 4
        If $ai_Version1[$i_Index] < $ai_Version2[$i_Index] Then; Version1 older than Version2
            $i_Result = -1
        ElseIf $ai_Version1[$i_Index] > $ai_Version2[$i_Index] Then; Version1 newer than Version2
            $i_Result = 1
        EndIf
   ; Bail-out if they're not equal
        If $i_Result <> 0 Then ExitLoop
    Next

    Return $i_Result

EndFunc