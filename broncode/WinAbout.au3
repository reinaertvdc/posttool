; ################
; # ABOUT WINDOW #
; ################
Func _WinAbout()
   GUISetState(@SW_DISABLE, $hWinMain)

   $sLicense = 'The MIT License (MIT)' & @CRLF & _
			   @CRLF & _
			   'Copyright (c) 2014-2016 Reinaert Van de Cruys' & @CRLF & _
			   @CRLF & _
			   'Permission is hereby granted, free of charge, to any person obtaining a copy ' & _
			   'of this software and associated documentation files (the "Software"), to deal ' & _
			   'in the Software without restriction, including without limitation the rights ' & _
			   'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell ' & _
			   'copies of the Software, and to permit persons to whom the Software is ' & _
			   'furnished to do so, subject to the following conditions:' & @CRLF & _
			   @CRLF & _
			   'The above copyright notice and this permission notice shall be included in all ' & _
			   'copies or substantial portions of the Software.' & @CRLF & _
			   @CRLF & _
			   'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR ' & _
			   'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, ' & _
			   'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE ' & _
			   'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER ' & _
			   'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, ' & _
			   'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE ' & _
			   'SOFTWARE.'

   $hWinAbout = GUICreate("Over Postcentrale", 400, 400, Default, Default, $WS_POPUP + $WS_CAPTION + $WS_SYSMENU)
	  GUISetIcon($sProgramLogo)
	  Opt("GUICoordMode", 0)
	  GUICtrlCreateLabel("Postcentrale", 10, 10, 380, 20, $SS_CENTER)
	  GUICtrlSetFont(Default, 16, 700)
	  GUICtrlCreateLabel("Versie " & $sVersion, 0, 28, 380, 15, $SS_CENTER)
	  GUICtrlCreateLabel("Auteur: Reinaert Van de Cruys", 0, 20, 380, 15, $SS_CENTER)
	  GUICtrlCreateLabel("Gemaakt met AutoIt v3, http://www.autoitscript.com/", 0, 32, 380, 15, $SS_CENTER)
	  Opt("GUICoordMode", 1)
	  GUICtrlCreateIcon($sAutoITLogo, Default, 10, 10, 62, 62)
	  GUICtrlCreateGroup("Licentie", 10, 120, 380, 225)
	  GUICtrlCreateEdit($sLicense, 20, 140, 360, 195, $WS_VSCROLL + $ES_MULTILINE + $ES_READONLY)
	  Local $idOK = GUICtrlCreateButton("OK", 150, 360, 100, 25)
   GUISetState(@SW_SHOW, $hWinAbout)

   While True
	  Switch GUIGetMsg()
		 Case $idOK
			GUISetState(@SW_ENABLE, $hWinMain)
			GUIDelete($hWinAbout)
			ExitLoop
		 Case $GUI_EVENT_CLOSE
			GUISetState(@SW_ENABLE, $hWinMain)
			GUIDelete($hWinAbout)
			ExitLoop
	  EndSwitch
   WEnd
EndFunc