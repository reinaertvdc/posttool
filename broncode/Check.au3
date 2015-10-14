; ###################
; # CHECK MAIN FORM #
; ###################
Func _FormCheck()
   $iErrorStyle = $MB_ICONERROR
   $sErrorTitle = "Ongeldige waarde"
   $iWarningStyle = $MB_ICONWARNING + $MB_YESNO + $MB_DEFBUTTON2
   $sWarningTitle = "Waarschuwing"
   $bValid = True

   ; ### TO ###
   ; Org
	  ; Always valid
   ; Salutation
   If $bValid And Not StringCompare($sToSalution, "") Then
	  MsgBox($iErrorStyle, $sErrorTitle, "Aanspreking van de bestemmeling vereist.")
	  $bValid = False
   EndIf
   ; FirstName
	  ; Always valid
   ; Surname
   If $bValid And Not StringCompare($sToSurname, "") Then
	  MsgBox($iErrorStyle, $sErrorTitle, "Achternaam van de bestemmeling vereist.")
	  $bValid = False
   EndIf
   ; Street
   If $bValid And Not StringCompare($sToStreet, "") Then
	  MsgBox($iErrorStyle, $sErrorTitle, "Straatnaam van de bestemmeling vereist.")
	  $bValid = False
   EndIf
   ; Number
   If $bValid And Not StringCompare($sToNumber, "") Then
	  MsgBox($iErrorStyle, $sErrorTitle, "Huisnummer van de bestemmeling vereist.")
	  $bValid = False
   EndIf
   ; Bus
	  ; Always valid
   ; Postcode
	  ; Always valid
   ; City
   If $bValid And Not StringCompare($sToCity, "") Then
	  MsgBox($iErrorStyle, $sErrorTitle, "Gemeente van de bestemmeling vereist.")
	  $bValid = False
   EndIf
   If $bValid And StringCompare($sToPostcode, "") And Not StringCompare($sToCountry, "België") Then
	  $bMatch = False
	  $iPostcodes = UBound($aPostcodes) - 1
	  For $i = 0 To $iPostcodes
		 If Not StringCompare($sToPostcode, $aPostcodes[$i][0]) And Not StringCompare($sToCity, $aPostcodes[$i][1]) Then $bMatch = True
	  Next
	  If Not $bMatch Then
		 If MsgBox($iWarningStyle, $sWarningTitle, "De gekozen postcode en gemeente komen mogelijk niet overeen. Wilt u toch doorgaan?") <> $IDYES Then $bValid = False
	  EndIf
   EndIf
   ; Country
   If $bValid And Not StringCompare($sToNumber, "") Then
	  MsgBox($iErrorStyle, $sErrorTitle, "Land van de bestemmeling vereist.")
	  $bValid = False
   EndIf
   ; Ref
	  ; Always valid
   ; ### FROM ###
   ; Org
   If $bValid And Not StringCompare($sFromOrg, "") Then
	  MsgBox($iErrorStyle, $sErrorTitle, "Organisatie van de afzender vereist.")
	  $bValid = False
   EndIf
   ; FirstName
   If $bValid And Not StringCompare($sFromFirstName, "") Then
	  MsgBox($iErrorStyle, $sErrorTitle, "Voornaam van de afzender vereist.")
	  $bValid = False
   EndIf
   ; Surname
   If $bValid And Not StringCompare($sFromSurname, "") Then
	  MsgBox($iErrorStyle, $sErrorTitle, "Achternaam van de afzender vereist.")
	  $bValid = False
   EndIf
   ; Department
   If $bValid And Not StringCompare($sFromDepartment, "") Then
	  MsgBox($iErrorStyle, $sErrorTitle, "Dienst van de afzender vereist.")
	  $bValid = False
   EndIf
   ; Tel
   If $bValid And Not StringCompare($sFromTel, "") Then
	  MsgBox($iErrorStyle, $sErrorTitle, "Telefoonnummer van de afzender vereist.")
	  $bValid = False
   EndIf
   If $bValid And StringLen($sFromTel) < 7 Then
	  MsgBox($iErrorStyle, $sErrorTitle, "Geldig telefoonnummer van de afzender vereist." & @CRLF & @CRLF & "Geldige formaten:" & @CRLF & "        014 54 06 12" & @CRLF& "        02 454 06 12" & @CRLF & "        0475 80 42 61" & @CRLF & "        +32475804261" & @CRLF & "        0032475804261")
	  $bValid = False
   EndIf
   ; Email
   If $bValid And Not StringCompare($sFromEmail, "") Then
	  MsgBox($iErrorStyle, $sErrorTitle, "E-mailadres van de afzender vereist.")
	  $bValid = False
   EndIf
   If $bValid And (StringInStr($sFromEmail, "@", 0, 1) < 2 Or StringInStr($sFromEmail, "@", 0, 2) > 0 Or _
	  StringInStr($sFromEmail, ".", 0, -1) < StringInStr($sFromEmail, "@", 0, 1) Or StringInStr($sFromEmail, ".", 0, -1) = 1 Or _
	  StringInStr($sFromEmail, ".", 0, -1) = StringLen($sFromEmail) Or StringInStr($sFromEmail, "..") Or _
	  StringInStr($sFromEmail, "@.") Or StringInStr($sFromEmail, ".@")) Then
	  MsgBox($iErrorStyle, $sErrorTitle, "Geldig e-mailadres van de afzender vereist." & @CRLF & @CRLF & "Formaat: 'naam@voorbeeld.be'")
	  $bValid = False
   EndIf
   ; Subject
   If $bValid And Not StringCompare($sFromSubject, "") Then
	  MsgBox($iErrorStyle, $sErrorTitle, "Onderwerp vereist.")
	  $bValid = False
   EndIf
   ; Date
   If $bValid And Not StringCompare($sFromDate, "") Then
	  MsgBox($iErrorStyle, $sErrorTitle, "Datum vereist.")
	  $bValid = False
   EndIf
   If $bValid And (Not StringLen($sFromDate) = 10 Or Not StringIsDigit(StringMid($sFromDate, 1, 2)) Or StringCompare(StringMid($sFromDate, 3, 1), "/") Or _
	  Not StringIsDigit(StringMid($sFromDate, 4, 2)) Or StringCompare(StringMid($sFromDate, 6, 1), "/") Or _
	  Not StringIsDigit(StringMid($sFromDate, 7, 4))) Then
	  MsgBox($iErrorStyle, $sErrorTitle, "Geldige Datum vereist." & @CRLF & @CRLF & "Formaat: 7/03/2014")
	  $bValid = False
   EndIf
   ; Ref
	  ; Always valid
   ; ### OPTIONS ###
   ; Annex
   If $bValid And StringCompare($sOptionsAnnex, "") Then
	  If Not FileExists($sOptionsAnnex) Then
		 MsgBox($iErrorStyle, $sErrorTitle, "De gekozen bijlage bestaat niet.")
		 $bValid = False
	  ElseIf StringCompare(StringRight($sOptionsAnnex, 4), ".pdf") Then
		 MsgBox($iErrorStyle, $sErrorTitle, "De gekozen bijlage is geen PDF-bestand.")
		 $bValid = False
	  EndIf
   EndIf

   Return $bValid
EndFunc