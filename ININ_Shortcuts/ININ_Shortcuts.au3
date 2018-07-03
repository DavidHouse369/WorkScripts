#include <Misc.au3>
#include <TrayConstants.au3>
#include <Array.au3>
#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>

If _Singleton("ININ_Shortcuts", 1) == 0 Then
   MsgBox(64, "Warning", "Only once instance of ININ_Shortcuts can be running at once." & @CRLF & "(This window will close automatically after two seconds)", 2)
   Exit
EndIf

Opt("TrayMenuMode", 3)
TraySetToolTip("ININ Shortcuts")

; Create dictionary object
$oHotKeys = ObjCreate("Scripting.Dictionary")

If @error Then
   MsgBox(0, "", "Error creating the dictionary object")
Else
   $oHotKeys.Add("^{F1}", "7") ; Available
   $oHotKeys.Add("^{F2}", "8") ; Available, No ACD
   $oHotKeys.Add("^{F3}", "6") ; At Lunch
   $oHotKeys.Add("^{F4}", "16") ; Gone Home
;~    $oHotKeys.Add("^{F5}", "#") ; Unassigned
;~    $oHotKeys.Add("^{F6}", "#") ; Unassigned
;~    $oHotKeys.Add("^{F7}", "#") ; Unassigned
;~    $oHotKeys.Add("^{F8}", "#") ; Unassigned

   $oHotKeys.Add("+^{F1}", "3") ; Admin Duties
   $oHotKeys.Add("+^{F2}", "9") ; Away from desk
   $oHotKeys.Add("+^{F3}", "11") ; Customer Assist Away From Desk
;~    $oHotKeys.Add("+^{F4}", "#") ; Unassigned
;~    $oHotKeys.Add("+^{F5}", "#") ; Unassigned
;~    $oHotKeys.Add("+^{F6}", "#") ; Unassigned
;~    $oHotKeys.Add("+^{F7}", "#") ; Unassigned
;~    $oHotKeys.Add("+^{F8}", "#") ; Unassigned

   HotKeySet("^{F1}", "HotKeyPressed")
   HotKeySet("^{F2}", "HotKeyPressed")
   HotKeySet("^{F3}", "HotKeyPressed")
   HotKeySet("^{F4}", "HotKeyPressed")
;~    HotKeySet("^{F5}", "HotKeyPressed")
;~    HotKeySet("^{F6}", "HotKeyPressed")
;~    HotKeySet("^{F7}", "HotKeyPressed")
;~    HotKeySet("^{F8}", "HotKeyPressed")

   HotKeySet("+^{F1}", "HotKeyPressed")
   HotKeySet("+^{F2}", "HotKeyPressed")
   HotKeySet("+^{F3}", "HotKeyPressed")
;~    HotKeySet("+^{F4}", "HotKeyPressed")
;~    HotKeySet("+^{F5}", "HotKeyPressed")
;~    HotKeySet("+^{F6}", "HotKeyPressed")
;~    HotKeySet("+^{F7}", "HotKeyPressed")
;~    HotKeySet("+^{F8}", "HotKeyPressed")

   Local $idShowShortcuts = TrayCreateItem("Show Shortcuts")
   Local $idCreateShortcut = TrayCreateItem("Create Shortcut")
   Local $idExit = TrayCreateItem("Exit")
   TraySetState($TRAY_ICONSTATE_SHOW)
   TrayTip("ININ_Shortcuts Script Running", "ININ_Shortcuts script is now running in the background.", 10, 1)
EndIf

;-------------------------------------------------------;
; Hotkey Release Codes									;
; WaitForRelease("10") ;SHIFT							;
; WaitForRelease("11") ;CTRL							;
; WaitForRelease("70") ;F1								;
; WaitForRelease("71") ;F2								;
; WaitForRelease("72") ;F3								;
; WaitForRelease("73") ;F4								;
; WaitForRelease("74") ;F5								;
; WaitForRelease("75") ;F6								;
; WaitForRelease("76") ;F7								;
; WaitForRelease("77") ;F8								;
;-------------------------------------------------------;

; Set Title Match Mode to only require part of the title
Opt("WinTitleMatchMode", 2)

; Get the position and size of the entire desktop (all screens)
$aPos = WinGetPos("Program Manager")

While (True)
   Switch TrayGetMsg()
	  Case $idShowShortcuts
		 Local $arrShortcuts = $oHotKeys.Keys
		 For $i = 0 To UBound($arrShortcuts) - 1
			$arrShortcuts[$i] &= " : " & $oHotKeys.Item($oHotKeys.Keys[$i])
		 Next
		 _ArrayDisplay($arrShortcuts)
	  Case $idCreateShortcut
		 CreateShortcut()
	  Case $idExit
		 ExitLoop
   EndSwitch
WEnd

;-------------------------------------------------------;
; Shortcuts Function									;
;-------------------------------------------------------;

Func HotKeyPressed()
   For $vKey In $oHotKeys
	  If @HotKeyPressed == $vKey Then ; We've found the hotkey
		 ; Wait for the keys to be released before moving forward
		 If StringInStr($vKey, "+") Then WaitForRelease("10") ;SHIFT
		 If StringInStr($vKey, "^") Then WaitForRelease("11") ;CTRL
		 If StringInStr($vKey, "{F1}") Then WaitForRelease("70") ;F1
		 If StringInStr($vKey, "{F2}") Then WaitForRelease("71") ;F2
		 If StringInStr($vKey, "{F3}") Then WaitForRelease("72") ;F3
		 If StringInStr($vKey, "{F4}") Then WaitForRelease("73") ;F4
		 If StringInStr($vKey, "{F5}") Then WaitForRelease("74") ;F5
		 If StringInStr($vKey, "{F6}") Then WaitForRelease("75") ;F6
		 If StringInStr($vKey, "{F7}") Then WaitForRelease("76") ;F7
		 If StringInStr($vKey, "{F8}") Then WaitForRelease("77") ;F8

		 ; Block user input while moving the mouse and clicking to prevent clicking the wrong items
		 BlockInput(1)
;~ 		 MsgBox(0, "", $vKey)
		 $aMPos = RightClickIcon()
		 $sendString = "{DOWN " & $oHotKeys.Item($vKey) & "}{ENTER}"
		 Send($sendString)
		 MouseMove($aMPos[0], $aMPos[1], 0)

		 ; Re-enable user input as we are done
		 BlockInput(0)
	  EndIf
   Next
EndFunc

;-------------------------------------------------------;
; Helper Functions										;
;-------------------------------------------------------;

Func CreateShortcut()
   ; Create a GUI with various controls.
    Local $hGUI = GUICreate("Create a shortcut", 140, 270)

    ; Create Checkboxes for modifier keys
    Local $idModifierLabel = GUICtrlCreateLabel("Modifier Key[s]: ", 10, 15, 75, 30)
    Local $idShift = GUICtrlCreateCheckbox("Shift", 85, 10, 50, 25)
    Local $idCtrl = GUICtrlCreateCheckbox("Ctrl", 85, 35, 50, 25)

    ; Create Radios for hotkeys
    Local $idHotkeysLabel = GUICtrlCreateLabel("Hotkey: ", 10, 75, 40, 30)
	Local $idF1 = GUICtrlCreateRadio("F1", 50, 70, 40, 25)
	Local $idF2 = GUICtrlCreateRadio("F2", 50, 95, 40, 25)
	Local $idF3 = GUICtrlCreateRadio("F3", 50, 120, 40, 25)
	Local $idF4 = GUICtrlCreateRadio("F4", 50, 145, 40, 25)
	Local $idF5 = GUICtrlCreateRadio("F5", 90, 70, 40, 25)
	Local $idF6 = GUICtrlCreateRadio("F6", 90, 95, 40, 25)
	Local $idF7 = GUICtrlCreateRadio("F7", 90, 120, 40, 25)
	Local $idF8 = GUICtrlCreateRadio("F8", 90, 145, 40, 25)
	GUICtrlSetState($idF1, $GUI_CHECKED)

    ; Create Combo for status ids
    Local $idStatusIDLabel = GUICtrlCreateLabel("Status ID: ", 10, 198, 50, 25)
    Local $idStatusID = GUICtrlCreateCombo("", 60, 195, 60, 25)
    GUICtrlSetData(-1, "1|2|3|4|5|6|7|8|9|10|11|12|13|14|15|16|17|18|19|20|21|22|23|24|25|26|27|28|29|30|31")

    Local $idButton_Create = GUICtrlCreateButton("Create", 60, 230, 50, 25)

    ; Display the GUI.
    GUISetState(@SW_SHOW, $hGUI)

    ; Loop until the user exits.
    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE
                ExitLoop

			Case $idButton_Create
			   $vNewKey = ""
			   If _IsChecked($idShift) Then $vNewKey &= "+"
			   If _IsChecked($idCtrl) Then $vNewKey &= "^"
			   If _IsChecked($idF1) Then $vNewKey &= "{F1}"
			   If _IsChecked($idF2) Then $vNewKey &= "{F2}"
			   If _IsChecked($idF3) Then $vNewKey &= "{F3}"
			   If _IsChecked($idF4) Then $vNewKey &= "{F4}"
			   If _IsChecked($idF5) Then $vNewKey &= "{F5}"
			   If _IsChecked($idF6) Then $vNewKey &= "{F6}"
			   If _IsChecked($idF7) Then $vNewKey &= "{F7}"
			   If _IsChecked($idF8) Then $vNewKey &= "{F8}"

			   $vStatusID = GUICtrlRead($idStatusID)
			   $oHotKeys.Item($vNewKey) = GUICtrlRead($idStatusID)

			   If $vStatusID == "" Then
				  HotKeySet($vNewKey, "")
			   Else
				  HotKeySet($vNewKey, "HotKeyPressed")
			   EndIf

			   ExitLoop
        EndSwitch
    WEnd

    ; Delete the previous GUI and all controls.
    GUIDelete($hGUI)
EndFunc

Func WaitForRelease($hKey)
   While (_IsPressed($hKey))
	  Sleep(10)
   WEnd
EndFunc

Func _IsChecked($idControlID)
    Return BitAND(GUICtrlRead($idControlID), $GUI_CHECKED) = $GUI_CHECKED
EndFunc

Func RightClickIcon()
   $aMPos = MouseGetPos()
   MouseClick("right", $aPos[0] + $aPos[2] / 3 * 2 - 155, $aPos[1] + $aPos[3] - 20, 1, 0)
   Return $aMPos
EndFunc