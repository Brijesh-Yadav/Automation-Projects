
Dim $title = $CmdLine[1]

For $i = 1 To 10
    If WinExists($title) Then
;	   MsgBox(0, "path", $title)
	   WinActivate($title)
	   Sleep(2000)
	   Send("{ENTER}")
       ExitLoop
	 Else
;     MsgBox(0, "path", "not present")
	Sleep(999)
    EndIf
Next




