#include <Misc.au3>

HotKeySet("^{F1}", "Available")
HotKeySet("^{F2}", "NotAvailable")
HotKeySet("^{F3}", "AtLunch")
HotKeySet("^{F4}", "GoneHome")
HotKeySet("^+{F1}", "AutoChangeOn")
HotKeySet("^+{F2}", "AutoChangeOff")

; Set Title Match Mode to only require part of the title
Opt("WinTitleMatchMode", 2)

; Get the position and size of the entire desktop (all screens)
$aPos = WinGetPos("Program Manager")

While (True)
   Sleep(10)
WEnd

Func Available()
   WaitForRelease("11") ;CTRL
   RightClickIcon()
   Send("{DOWN 7}{ENTER}")
EndFunc

Func NotAvailable()
   WaitForRelease("11") ;CTRL
   RightClickIcon()
   Send("{DOWN 8}{ENTER}")
EndFunc

Func AtLunch()
   WaitForRelease("11") ;CTRL
   RightClickIcon()
   Send("{DOWN 6}{ENTER}")
EndFunc

Func GoneHome()
   WaitForRelease("11") ;CTRL
   RightClickIcon()
   Send("{DOWN 16}{ENTER}")
EndFunc

Func AutoChangeOn()
   WaitForRelease("10") ;SHIFT
   WaitForRelease("11") ;CTRL
   If (WinExists("Interaction Client: .NET Edition")) Then
	  $aININPos = OpenININConfig()
	  $aMPos = MouseGetPos()
	  MouseClick("left", $aININPos[0] + $aININPos[2] - 402, $aININPos[1] + 182, 1, 0)
	  MouseMove($aMPos[0], $aMPos[1], 0)
	  Send("{ENTER}")
   EndIf
EndFunc

Func AutoChangeOff()
   WaitForRelease("10") ;SHIFT
   WaitForRelease("11") ;CTRL
   If (WinExists("Interaction Client: .NET Edition")) Then
	  $aININPos = OpenININConfig()
	  $aMPos = MouseGetPos()
	  MouseClick("left", $aININPos[0] + $aININPos[2] - 402, $aININPos[1] + 156, 1, 0)
	  MouseMove($aMPos[0], $aMPos[1], 0)
	  Send("{ENTER}")
   EndIf
EndFunc

Func OpenININConfig()
   WinActivate("Interaction Client: .NET Edition")
   Send("!oc")
   Sleep(100)
   Send("{TAB 2}")
   Sleep(100)
   Send("au")
   Return WinGetPos("Configuration")
EndFunc

Func WaitForRelease($hKey)
   While (_IsPressed($hKey))
	  Sleep(10)
   WEnd
EndFunc

Func RightClickIcon()
   $aMPos = MouseGetPos()
   MouseClick("right", $aPos[0] + $aPos[2] / 3 * 2 - 155, $aPos[1] + $aPos[3] - 20, 1, 0)
   MouseMove($aMPos[0], $aMPos[1], 0)
EndFunc