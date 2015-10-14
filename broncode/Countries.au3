Func _LoadCountries()
   Global $sCountries = _
		"|Afghanistan" _
	  & "|Albanië" _
	  & "|Algerije" _
	  & "|Amerikaans-Samoa" _
	  & "|Amerikaanse Maagdeneilanden" _
	  & "|Andorra" _
	  & "|Angola" _
	  & "|Anguilla" _
	  & "|Antigua en Barbuda" _
	  & "|Argentinië" _
	  & "|Armenië" _
	  & "|Aruba" _
	  & "|Australië" _
	  & "|Azerbeidzjan" _
	  & "|Bahama's" _
	  & "|Bahrein" _
	  & "|Bangladesh" _
	  & "|Barbados" _
	  & "|Belarus" _
	  & "|Belau" _
	  & "|België" _
	  & "|Belize" _
	  & "|Benin" _
	  & "|Bermuda" _
	  & "|Bhutan" _
	  & "|Bolivië" _
	  & "|Bolivia" _
	  & "|Bosnië en Herzegovina" _
	  & "|Botswana" _
	  & "|Brazilië" _
	  & "|Brits Indische Oceaanterritorium" _
	  & "|Britse Maagdeneilanden" _
	  & "|Brunei" _
	  & "|Bulgarije" _
	  & "|Burkina Faso" _
	  & "|Burundi" _
	  & "|Cambodja" _
	  & "|Canada" _
	  & "|Caymaneilanden" _
	  & "|Centraal-Afrikaanse Republiek" _
	  & "|Chili" _
	  & "|China" _
	  & "|Christmaseiland" _
	  & "|Colombia" _
	  & "|Comoren" _
	  & "|Congo(-Brazzaville)" _
	  & "|Congo(-Kinshasa)" _
	  & "|Cookeilanden" _
	  & "|Costa Rica" _
	  & "|Cuba" _
	  & "|Curaçao" _
	  & "|Cyprus" _
	  & "|Denemarken" _
	  & "|Djibouti" _
	  & "|Dominica" _
	  & "|Dominicaanse Republiek" _
	  & "|Duitsland" _
	  & "|Ecuador" _
	  & "|Egypte" _
	  & "|El Salvador" _
	  & "|Equatoriaal-Guinea" _
	  & "|Eritrea" _
	  & "|Estland" _
	  & "|Ethiopië" _
	  & "|Faeröer" _
	  & "|Falklandeilanden" _
	  & "|Fiji" _
	  & "|Filipijnen" _
	  & "|Finland" _
	  & "|Frankrijk" _
	  & "|Frans-Guyana" _
	  & "|Frans-Polynesië" _
	  & "|Gabon" _
	  & "|Gambia" _
	  & "|Georgië" _
	  & "|Ghana" _
	  & "|Gibraltar" _
	  & "|Grenada" _
	  & "|Griekenland" _
	  & "|Groenland" _
	  & "|Groot-Brittannië (en Noord-Ierland)" _
	  & "|Guadeloupe" _
	  & "|Guam" _
	  & "|Guatemala" _
	  & "|Guinee" _
	  & "|Guinee-Bissau" _
	  & "|Guyana" _
	  & "|Haïti" _
	  & "|Honduras" _
	  & "|Hongarije" _
	  & "|Hongkong" _
	  & "|Ierland" _
	  & "|IJsland" _
	  & "|India" _
	  & "|Indonesië" _
	  & "|Irak" _
	  & "|Iran" _
	  & "|Israël" _
	  & "|Italië" _
	  & "|Ivoorkust" _
	  & "|Jamaica" _
	  & "|Japan" _
	  & "|Jemen" _
	  & "|Jordanië" _
	  & "|Kaaimaneilanden" _
	  & "|Kaapverdië" _
	  & "|Kameroen" _
	  & "|Kazachstan" _
	  & "|Kenia" _
	  & "|Kenya" _
	  & "|Kirgistan" _
	  & "|Kirgizië" _
	  & "|Kiribati" _
	  & "|Koeweit" _
	  & "|Kosovo" _
	  & "|Kroatië" _
	  & "|Laos" _
	  & "|Lesotho" _
	  & "|Letland" _
	  & "|Libanon" _
	  & "|Liberia" _
	  & "|Libië" _
	  & "|Liechtenstein" _
	  & "|Litouwen" _
	  & "|Luxemburg" _
	  & "|Macau" _
	  & "|Macedonië" _
	  & "|Madagaskar" _
	  & "|Malawi" _
	  & "|Maldiven" _
	  & "|Malediven" _
	  & "|Maleisië" _
	  & "|Mali" _
	  & "|Malta" _
	  & "|Marokko" _
	  & "|Marshalleilanden" _
	  & "|Martinique" _
	  & "|Mauritanië" _
	  & "|Mauritius" _
	  & "|Mexico" _
	  & "|Micronesia" _
	  & "|Moldavië" _
	  & "|Monaco" _
	  & "|Mongolië" _
	  & "|Montenegro" _
	  & "|Montserrat" _
	  & "|Mozambique" _
	  & "|Myanmar" _
	  & "|Namibië" _
	  & "|Nauru" _
	  & "|Nederland" _
	  & "|Nederlandse Antillen" _
	  & "|Nepal" _
	  & "|Nicaragua" _
	  & "|Nieuw-Caledonië" _
	  & "|Nieuw-Zeeland" _
	  & "|Niger" _
	  & "|Nigeria" _
	  & "|Niue" _
	  & "|Noord-Korea" _
	  & "|Noordelijke Marianen" _
	  & "|Noorwegen" _
	  & "|Norfolk" _
	  & "|Oeganda" _
	  & "|Oekraïne" _
	  & "|Oezbekistan" _
	  & "|Oman" _
	  & "|Oost-Timor" _
	  & "|Oostenrijk" _
	  & "|Pakistan" _
	  & "|Palau" _
	  & "|Palestina" _
	  & "|Panama" _
	  & "|Papoea-Nieuw-Guinea" _
	  & "|Paraguay" _
	  & "|Peru" _
	  & "|Pitcairneilanden" _
	  & "|Polen" _
	  & "|Porto Rico" _
	  & "|Portugal" _
	  & "|Puerto Rico" _
	  & "|Qatar" _
	  & "|Réunion" _
	  & "|Roemenië" _
	  & "|Rusland" _
	  & "|Rwanda" _
	  & "|Saint Kitts en Nevis" _
	  & "|Saint Lucia" _
	  & "|Saint Vincent en de Grenadines" _
	  & "|Saint-Pierre en Miquelon" _
	  & "|Salomonseilanden" _
	  & "|Samoa" _
	  & "|San Marino" _
	  & "|Sao Tomé en Principe" _
	  & "|Saoedi-Arabië" _
	  & "|Saudi-Arabië" _
	  & "|Senegal" _
	  & "|Servië" _
	  & "|Seychellen" _
	  & "|Sierra Leone" _
	  & "|Singapore" _
	  & "|Sint-Helena" _
	  & "|Sint-Maarten" _
	  & "|Slovakije" _
	  & "|Slovenië" _
	  & "|Slowakije" _
	  & "|Soedan" _
	  & "|Somalië" _
	  & "|Spanje" _
	  & "|Sri Lanka" _
	  & "|Sudan" _
	  & "|Suriname" _
	  & "|Swaziland" _
	  & "|Syrië" _
	  & "|Tadzjikistan" _
	  & "|Taiwan" _
	  & "|Tanzania" _
	  & "|Thailand" _
	  & "|Togo" _
	  & "|Tokelau" _
	  & "|Tonga" _
	  & "|Trinidad en Tobago" _
	  & "|Tsjaad" _
	  & "|Tsjechië" _
	  & "|Tunesië" _
	  & "|Turkije" _
	  & "|Turkmenistan" _
	  & "|Turks- en Caicoseilanden" _
	  & "|Tuvalu" _
	  & "|Uganda" _
	  & "|Uruguay" _
	  & "|Vanuatu" _
	  & "|Vaticaanstad" _
	  & "|Venezuela" _
	  & "|Verenigd Koninkrijk" _
	  & "|Verenigde Arabische Emiraten" _
	  & "|Verenigde Staten" _
	  & "|Vietnam" _
	  & "|Wallis en Futuna" _
	  & "|Wit-Rusland" _
	  & "|Zambia" _
	  & "|Zimbabwe" _
	  & "|Zuid-Afrika" _
	  & "|Zuid-Georgia en de Zuidelijke Sandwicheilanden" _
	  & "|Zuid-Korea" _
	  & "|Zuid-Soedan" _
	  & "|Zuid-Sudan" _
	  & "|Zweden" _
	  & "|Zwitserland"
EndFunc