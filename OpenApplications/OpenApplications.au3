#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>

; Wait 5 seconds to allow login to complete
Sleep(5000)

; Set Title Match Mode to only require part of the title
Opt("WinTitleMatchMode", 2)

; Get the position and size of the entire desktop (all screens)
$aPos = WinGetPos("Program Manager")

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

; Open Outlook
Run(@ComSpec & " /c " & '"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Outlook 2010.lnk"', @TempDir, @SW_HIDE)

; Open ECM
Run(@ComSpec & " /c " & '"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\TechnologyOne ECM\ECM SmartClient.url"', @TempDir, @SW_HIDE)
While Not WaitForWindow("TCC ECM Service Desk Officer", False, 30, False)
   If MsgBox($MB_YesNo, 'Failed To Open', '"TCC ECM Service Desk Officer" has failed to open so far.' & @CRLF & 'Are you still entering details?') == $IDNO Then
	  ExitLoop
   EndIf
WEnd


; Open CES and P&R
Run(@ComSpec & " /c " & '"D:\Users\dsh\Desktop\TechOneCES.lnk"', @TempDir, @SW_HIDE)
Run(@ComSpec & " /c " & '"D:\Users\dsh\Desktop\TechOneP&R.lnk"', @TempDir, @SW_HIDE)

; Open Snagit
Run("C:\Program Files (x86)\TechSmith\Snagit 11\Snagit32")

; Open DameWare remote access with admin rights
Run(@ComSpec & " /c " & '""D:\Users\dsh\Documents\Admin Tools\DameWare Mini Remote.lnk""', @TempDir, @SW_HIDE)

; Open Account Lockout Tool with admin rights
Run(@ComSpec & " /c " & '""D:\Users\dsh\Documents\Admin Tools\Account Lockout Tool.lnk""', @TempDir, @SW_HIDE)

; Open Internet Explorer, open a new tab and load the OrgChart from favourites
Run("C:\Program Files\Internet Explorer\iexplore.exe","",@SW_MAXIMIZE)
If WaitForWindow(" - Internet Explorer", True, 15, True) Then
   Send("^t")
   Sleep(300)
   Send("!a")
   Sleep(300)
   Send("f")
   Sleep(300)
   Send("o")
EndIf

; Open Chrome with starting tabs
Run("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe","",@SW_MAXIMIZE)
If WaitForWindow(" - Google Chrome", True, 15, True) Then
   Send("^{TAB}")
EndIf

; Open ININ and hit Logon
Run("C:\Program Files (x86)\ININ Call Centre User (VMware ThinApp)\Interaction Client.exe")
If WaitForWindow("Interaction Client", True, 30, True) Then
   Send("{ENTER}")
   If WaitForWindow("Interaction Client: .NET Edition", True, 30, True) Then
	  WinMove("Interaction Client: .NET Edition", "", $aPos[0] + $aPos[2] /3 - 655, $aPos[1], 655, $aPos[3])
   EndIf
EndIf

; Open AD with admin rights and open the search window
Run(@ComSpec & " /c " & '""D:\Users\dsh\Documents\Admin Tools\AD Users & Computers.lnk""', @TempDir, @SW_HIDE)
If WaitForWindow("Active Directory Users and Computers", True, 15, True) Then
   Send("c")
   Sleep(300)
   Send("!a")
   Sleep(300)
   Send("i")
EndIf

; Open Notepad++, position and resize it, and run the "NewDay" script
Run("D:\Users\dsh\Documents\Notepad++\notepad++.exe")
If WaitForWindow("Notepad++", True, 15, True) Then
   Send("^{END}")
   Sleep(300)
   Send("{ENTER 4}")
   Send("+{HOME}")
   Send("{BACKSPACE}")
   Sleep(300)
   Send("^-")
   Send("{ENTER}")
   Sleep(300)
   Send("^{NUMPADDOT}")
EndIf

; Navigate to correct page in ECM
If WaitForWindow("TCC ECM Service Desk Officer", True, 0, True) Then
   WinSetState("TCC ECM Service Desk Officer", "", @SW_RESTORE)
   WinMove("TCC ECM Service Desk Officer", "", $aPos[0] + $aPos[2] /3 * 2 + 570, $aPos[1], $aPos[2] /3 - 570, $aPos[3])
   Send("!v")
   Sleep(300)
   Send("w")
   Sleep(300)
   Send("s")
   Sleep(3000)
   Send("{TAB}{ENTER}")
EndIf

; Navigate to correct page in CES
If WaitForWindow("TCC Service Desk Officer - TechnologyOne Enterprise Suite", True, 0, True) Then
   WinSetState("TCC Service Desk Officer - TechnologyOne Enterprise Suite", "", @SW_RESTORE)
   WinMove("TCC Service Desk Officer - TechnologyOne Enterprise Suite", "", $aPos[0] + $aPos[2] /3 * 2 + 570, $aPos[1], $aPos[2] /3 - 570, $aPos[3])
   Send("!v")
   Sleep(300)
   Send("w")
   Sleep(300)
   Send("u{ENTER}")
   Sleep(300)
   Send("!a")
   Sleep(300)
   Send("h{ENTER}")
   Sleep(4000)
   Send("{ENTER}")
EndIf

; Navigate to correct page in P&R
If WaitForWindow("Service Desk Officer - TechnologyOne Property", True, 0, True) Then
   WinSetState("Service Desk Officer - TechnologyOne Property", "", @SW_RESTORE)
   WinMove("Service Desk Officer - TechnologyOne Property", "", $aPos[0] + $aPos[2] /3 * 2 + 570, $aPos[1], $aPos[2] /3 - 570, $aPos[3])
   Sleep(1000)
   Send("!v")
   Sleep(300)
   Send("w")
   Sleep(300)
   Send("u")
EndIf

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
MsgBox(64, "ININ Shortcuts", 'Click Ok to run "ININ Shortcuts" script.')
Run(@ComSpec & " /c " & '""D:\Users\dsh\Desktop\ININ_Shortcuts.lnk""', "", @SW_HIDE)

; Run Text Replacements script
MsgBox(64, "Text Replacements", 'Click Ok to run "Text Replacements" script.')
Run(@ComSpec & " /c " & '""D:\Users\dsh\Desktop\TextReplacements.lnk""', "", @SW_HIDE)

MsgBox(64, "Success", "OpenApplications script ran successfully." & @CRLF & "(This window will close after 2 seconds)", 2)