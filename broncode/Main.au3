#CS ###########################################################################

The MIT License (MIT)

Copyright (c) 2014-2015 Reinaert Van de Cruys

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

#CE ###########################################################################





; ####################
; # COMPILER OPTIONS #
; ####################
#pragma compile(Out,                  "Verwerking uitgaande post.exe"                                                )
#pragma compile(Icon,                 "resources\postcentrale.ico"                                                   )
#pragma compile(ExecLevel,            highestavailable                                                               )
#pragma compile(UPX,                  false                                                                          )
#pragma compile(AutoItExecuteAllowed, true                                                                           )
#pragma compile(Console,              false                                                                          )
#pragma compile(Compression,          3                                                                              )
#pragma compile(Compatibility,        vista                                                                          )
#pragma compile(x64,                  false                                                                          )
#pragma compile(inputboxres,          false                                                                          )
#pragma compile(Comments,             ""                                                                             )
#pragma compile(CompanyName,          ""                                                                             )
#pragma compile(FileDescription,      ""                                                                             )
#pragma compile(FileVersion,          ""                                                                             )
#pragma compile(InternalName,         ""                                                                             )
#pragma compile(LegalCopyright,       "Licensed under the MIT License, Copyright (c) 2014-2015 Reinaert Van de Cruys")
#pragma compile(LegalTrademarks,      ""                                                                             )
#pragma compile(OriginalFilename,     ""                                                                             )
#pragma compile(ProductName,          "Postcentrale"                                                                 )
#pragma compile(ProductVersion,       "0.13"                                                                         )

Global $sProgramLogo = _TempFile()
FileInstall("resources\postcentrale.ico", $sProgramLogo)
Global $sAutoITLogo = _TempFile()
FileInstall("resources\autoit.ico", $sAutoITLogo)
$sVersion = FileGetVersion(@ScriptFullPath, "ProductVersion")





; ############
; # INCLUDES #
; ############
#include <APIDlgConstants.au3>
#include <Array.au3>
#include <ComboConstants.au3>
#include <Constants.au3>
#include <DateTimeConstants.au3>
#include <EditConstants.au3>
#include <Excel.au3>
#include <ExcelConstants.au3>
#include <File.au3>
#include <FileConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiListView.au3>
#include <GuiScrollBars.au3>
#include <ListViewConstants.au3>
#include <Math.au3>
#include <Memory.au3>
#include <MsgBoxConstants.au3>
#include <StaticConstants.au3>
#include <WinAPIDlg.au3>
#include <WinAPIFiles.au3>
#include <WindowsConstants.au3>
#include <Word.au3>

#include "GUIScrollbars_Ex.au3"

#include "Check.au3"
#include "Control.au3"
#include "Countries.au3"
#include "Database.au3"
#include "Letter.au3"
#include "Postcodes.au3"
#include "Print.au3"
#include "Settings.au3"
#include "WinAbout.au3"
#include "WinMain.au3"
#include "WinOrder.au3"





; #################
; # MAIN FUNCTION #
; #################
Func _Main()
   _ControlMainLoop()
   Exit 0
EndFunc





; ##############
; # ENTRYPOINT #
; ##############
_Main()