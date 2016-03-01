$hIE = WinGetHandle("View Downloads - Windows Internet Explorer")
$hCtrl = ControlGetHandle($hIE,"","[Class:DirectUIHWND]")
$aPos = ControlGetPos($hIE,"",$hCtrl)
$x = $aPos[2]-250
$y = $aPos[3]-35
;Use
WinActivate($hIE)
;doesn't work in the background
ControlClick($hIE,"",$hCtrl,"primary",1,$x,$y)
;this only gives focus to the save button
ControlSend($hIE,"",$hCtrl,"{Enter}")
Send("{TAB}")
Sleep(1000)
Send("{TAB}")
Sleep(500)
Send("{TAB}")
Sleep(500)
Send("{ENTER}")






