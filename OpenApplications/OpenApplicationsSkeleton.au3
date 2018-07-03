#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>

Func WaitForWindow($winString, $needActive=False, $timeOut=-1, $showFailMsg=False)
   $exists = WinExists($winString)
   While (Not $exists) And ($timeOut <> 0)
	  Sleep(1000)
	  $timeOut -= 1
	  $exists = WinExists($winString)
   WEnd
   If $exists Then
	  If $needActive Then
		 WinActivate($winString)
	  EndIf
   Else
	  If $showFailMsg Then MsgBox(64, "Failed To Open", '"' & $winString & '" has failed to open' & @CRLF & '(This window will close after 2 seconds)', 2)
   EndIf
   Return $exists
EndFunc

; Wait 5 seconds to allow login to complete
Sleep(5000)

; Set Title Match Mode to only require part of the title
Opt("WinTitleMatchMode", 2)

; Get the position and size of the entire desktop (all screens)
$aPos = WinGetPos("Program Manager")


; Run ECM and wait for sign-in
Run(@ComSpec & " /c " & '"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\TechnologyOne ECM\ECM SmartClient.url"', @TempDir, @SW_HIDE)
While Not WaitForWindow("TCC ECM Service Desk Officer", False, 20, False)
   If MsgBox($MB_YesNo, 'Failed To Open', '"TCC ECM Service Desk Officer" has failed to open so far.' & @CRLF & 'Are you still entering details?') == $IDNO Then
	  ExitLoop
   EndIf
WEnd

; Open Outlook
Run(@ComSpec & " /c " & '"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Outlook 2010.lnk"', @TempDir, @SW_HIDE)

; Open Snagit
Run("C:\Program Files (x86)\TechSmith\Snagit 11\Snagit32")

; Open DameWare remote access with admin rights
Run(@ComSpec & " /c " & '""D:\Users\dsh\Documents\Admin Tools\DameWare Mini Remote.lnk""', @TempDir, @SW_HIDE)

; Open Account Lockout Tool with admin rights
Run(@ComSpec & " /c " & '""D:\Users\dsh\Documents\Admin Tools\Account Lockout Tool.lnk""', @TempDir, @SW_HIDE)

; Open Internet Explorer
Run("C:\Program Files\Internet Explorer\iexplore.exe","",@SW_MAXIMIZE)

; Open Chrome with starting tabs
Run("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe","",@SW_MAXIMIZE)

; Open ININ
Run("C:\Program Files (x86)\ININ Call Centre User (VMware ThinApp)\Interaction Client.exe")

; Open AD with admin rights
Run(@ComSpec & " /c " & '""D:\Users\dsh\Documents\Admin Tools\AD Users & Computers.lnk""', @TempDir, @SW_HIDE)

; Open Notepad++
Run("D:\Users\dsh\Documents\Notepad++\notepad++.exe")

; Open CES and P&R
Run(@ComSpec & " /c " & '"D:\Users\dsh\Desktop\TechOneCES.lnk"', @TempDir, @SW_HIDE)
While Not WaitForWindow("TCC Service Desk Officer - TechnologyOne Enterprise Suite", False, 15, False)
   If MsgBox($MB_YesNo, 'Failed To Open', '"TCC Service Desk Officer - TechnologyOne Enterprise Suite" has failed to open so far.' & @CRLF & 'Are you still entering details?') == $IDNO Then
	  ExitLoop
   EndIf
WEnd

Run(@ComSpec & " /c " & '"D:\Users\dsh\Desktop\TechOneP&R.lnk"', @TempDir, @SW_HIDE)

; OrganiseApplications
; Position Outlook
If WinExists("Microsoft Outlook") Then
   WinMove("Microsoft Outlook", "", $aPos[0], $aPos[1], $aPos[2] / 3 - 655, $aPos[3])
EndIf

; Position Notepad++
If WinExists("Notepad++") Then
   WinMove("Notepad++", "", $aPos[0] + $aPos[2] / 3 * 2, $aPos[1], 570, $aPos[3])
EndIf

; Position ININ
If WinExists("Interaction Client: .NET Edition") Then
   WinMove("Interaction Client: .NET Edition", "", $aPos[0] + $aPos[2] /3 - 655, $aPos[1], 655, $aPos[3])
EndIf

; Run ININ Shortcuts script
Run(@ComSpec & " /c " & '""D:\Users\dsh\Desktop\ININ_Shortcuts.lnk""', "", @SW_HIDE)

; Run Text Replacements script
Run(@ComSpec & " /c " & '""D:\Users\dsh\Desktop\TextReplacements.lnk""', "", @SW_HIDE)

MsgBox(64, "Success", "OpenApplications script ran successfully." & @CRLF & "(This window will close after 2 seconds)", 2)