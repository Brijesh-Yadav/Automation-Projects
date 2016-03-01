
;Dim $title = "https://www.google.com.au/url?sa=t&rct=j&q=&esrc=s&frm=1&source=web&cd=17&ved=0CE8QFjAGOAo&url= - Windows Internet Explorer"

For $i = 1 To 10
	$hIE = WinGetHandle("[Class:IEFrame]")
    If WinExists($hIE) Then
;	   MsgBox(0, "path", "Present")
	   $hCtrl = ControlGetHandle($hIE,"","[Class:DirectUIHWND]")
	   $aPos = ControlGetPos($hIE,"",$hCtrl)
       $x = $aPos[2]-150
       $y = $aPos[3]-35
       ;Use
       WinActivate($hIE)
       ;doesn't work in the background
       ControlClick($hIE,"",$hCtrl,"primary",1,$x,$y)
       ;this only gives focus to the save button
       ControlSend($hIE,"",$hCtrl,"{Enter}")
;	   MouseMove($x,$y)
       ExitLoop
	 Else
;       MsgBox(0, "", "The value of $i is: " & $i)
	   Sleep(999)
    EndIf
Next

