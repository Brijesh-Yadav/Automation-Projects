
Dim $filepath = $CmdLine[1]

For $i = 1 To 10
    If WinExists("File Upload") Then
;	   MsgBox(0, "path", $filepath)
	   WinActivate("File Upload")
	   Send($filepath);
	   Send("{ENTER}")
       ExitLoop
	 Else
;(0, "", "The value of $i is: " & $i)
	Sleep(999)
    EndIf
Next




