Func _LoadCountries()
   Global $sCountries = _
		"|Afghanistan" _
	  & "|Albani�" _
	  & "|Algerije" _
	  & "|Amerikaans-Samoa" _
	  & "|Amerikaanse Maagdeneilanden" _
	  & "|Andorra" _
	  & "|Angola" _
	  & "|Anguilla" _
	  & "|Antigua en Barbuda" _
	  & "|Argentini�" _
	  & "|Armeni�" _
	  & "|Aruba" _
	  & "|Australi�" _
	  & "|Azerbeidzjan" _
	  & "|Bahama's" _
	  & "|Bahrein" _
	  & "|Bangladesh" _
	  & "|Barbados" _
	  & "|Belarus" _
	  & "|Belau" _
	  & "|Belgi�" _
	  & "|Belize" _
	  & "|Benin" _
	  & "|Bermuda" _
	  & "|Bhutan" _
	  & "|Bolivi�" _
	  & "|Bolivia" _
	  & "|Bosni� en Herzegovina" _
	  & "|Botswana" _
	  & "|Brazili�" _
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
	  & "|Cura�ao" _
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
	  & "|Ethiopi�" _
	  & "|Faer�er" _
	  & "|Falklandeilanden" _
	  & "|Fiji" _
	  & "|Filipijnen" _
	  & "|Finland" _
	  & "|Frankrijk" _
	  & "|Frans-Guyana" _
	  & "|Frans-Polynesi�" _
	  & "|Gabon" _
	  & "|Gambia" _
	  & "|Georgi�" _
	  & "|Ghana" _
	  & "|Gibraltar" _
	  & "|Grenada" _
	  & "|Griekenland" _
	  & "|Groenland" _
	  & "|Groot-Brittanni� (en Noord-Ierland)" _
	  & "|Guadeloupe" _
	  & "|Guam" _
	  & "|Guatemala" _
	  & "|Guinee" _
	  & "|Guinee-Bissau" _
	  & "|Guyana" _
	  & "|Ha�ti" _
	  & "|Honduras" _
	  & "|Hongarije" _
	  & "|Hongkong" _
	  & "|Ierland" _
	  & "|IJsland" _
	  & "|India" _
	  & "|Indonesi�" _
	  & "|Irak" _
	  & "|Iran" _
	  & "|Isra�l" _
	  & "|Itali�" _
	  & "|Ivoorkust" _
	  & "|Jamaica" _
	  & "|Japan" _
	  & "|Jemen" _
	  & "|Jordani�" _
	  & "|Kaaimaneilanden" _
	  & "|Kaapverdi�" _
	  & "|Kameroen" _
	  & "|Kazachstan" _
	  & "|Kenia" _
	  & "|Kenya" _
	  & "|Kirgistan" _
	  & "|Kirgizi�" _
	  & "|Kiribati" _
	  & "|Koeweit" _
	  & "|Kosovo" _
	  & "|Kroati�" _
	  & "|Laos" _
	  & "|Lesotho" _
	  & "|Letland" _
	  & "|Libanon" _
	  & "|Liberia" _
	  & "|Libi�" _
	  & "|Liechtenstein" _
	  & "|Litouwen" _
	  & "|Luxemburg" _
	  & "|Macau" _
	  & "|Macedoni�" _
	  & "|Madagaskar" _
	  & "|Malawi" _
	  & "|Maldiven" _
	  & "|Malediven" _
	  & "|Maleisi�" _
	  & "|Mali" _
	  & "|Malta" _
	  & "|Marokko" _
	  & "|Marshalleilanden" _
	  & "|Martinique" _
	  & "|Mauritani�" _
	  & "|Mauritius" _
	  & "|Mexico" _
	  & "|Micronesia" _
	  & "|Moldavi�" _
	  & "|Monaco" _
	  & "|Mongoli�" _
	  & "|Montenegro" _
	  & "|Montserrat" _
	  & "|Mozambique" _
	  & "|Myanmar" _
	  & "|Namibi�" _
	  & "|Nauru" _
	  & "|Nederland" _
	  & "|Nederlandse Antillen" _
	  & "|Nepal" _
	  & "|Nicaragua" _
	  & "|Nieuw-Caledoni�" _
	  & "|Nieuw-Zeeland" _
	  & "|Niger" _
	  & "|Nigeria" _
	  & "|Niue" _
	  & "|Noord-Korea" _
	  & "|Noordelijke Marianen" _
	  & "|Noorwegen" _
	  & "|Norfolk" _
	  & "|Oeganda" _
	  & "|Oekra�ne" _
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
	  & "|R�union" _
	  & "|Roemeni�" _
	  & "|Rusland" _
	  & "|Rwanda" _
	  & "|Saint Kitts en Nevis" _
	  & "|Saint Lucia" _
	  & "|Saint Vincent en de Grenadines" _
	  & "|Saint-Pierre en Miquelon" _
	  & "|Salomonseilanden" _
	  & "|Samoa" _
	  & "|San Marino" _
	  & "|Sao Tom� en Principe" _
	  & "|Saoedi-Arabi�" _
	  & "|Saudi-Arabi�" _
	  & "|Senegal" _
	  & "|Servi�" _
	  & "|Seychellen" _
	  & "|Sierra Leone" _
	  & "|Singapore" _
	  & "|Sint-Helena" _
	  & "|Sint-Maarten" _
	  & "|Slovakije" _
	  & "|Sloveni�" _
	  & "|Slowakije" _
	  & "|Soedan" _
	  & "|Somali�" _
	  & "|Spanje" _
	  & "|Sri Lanka" _
	  & "|Sudan" _
	  & "|Suriname" _
	  & "|Swaziland" _
	  & "|Syri�" _
	  & "|Tadzjikistan" _
	  & "|Taiwan" _
	  & "|Tanzania" _
	  & "|Thailand" _
	  & "|Togo" _
	  & "|Tokelau" _
	  & "|Tonga" _
	  & "|Trinidad en Tobago" _
	  & "|Tsjaad" _
	  & "|Tsjechi�" _
	  & "|Tunesi�" _
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