;---------------------------------------------------------
;~ Save_Dialog_FF.au3
;~ Purpose: To handle the Dowload/save Dialogbox in Firefox
;~ Usage: Save_Dialog_FF.exe "Dialog Title" "Opetaion" "Path"
;----------------------------------------------------------

 ; set the select mode to select using substring

AutoItSetOption("WinTitleMatchMode","2")

ControlCommand ( "Opening Employee Profile Report", "Save", "&Save", "Check" )

; wait until dialog box appears

WinWait($CmdLine[1]) ; match the window with substring

$title = WinGetTitle($CmdLine[1]) ; retrives whole window title

WinActive($title);

; if user choose to save file

If (StringCompare($CmdLine[2],"Save",0) = 0) Then

WinActivate($title)

WinWaitActive($title)

Sleep(10)

; If firefox is set to save the file on some specif location without asking to user.

; It will be save after this point.

Send("{DOWN}")

Send("{ENTER}")

;If Dialog appear prompting for Path to save

WinWait("Enter name")

$title = WinGetTitle("Enter name")

If WinExists($title) Then

;Set path and save file

WinActivate($title)

WinWaitActive($title)

ControlSetText($title, "", "Edit1", $CmdLine[3])

ControlClick($title, "", "Button2")

EndIf

WinWait("Downloads")

If WinExists("Downloads") Then

$title = WinGetTitle("Downloads")

WinActivate($title)

WinWaitActive($title)

Send("{ESCAPE}")

EndIf

Else

;Firefox is configured to save file at specific location

Exit

EndIf

; do not save the file

If (StringCompare("Cancel","Cancel",0) = 0) Then

WinWaitActive($title)

Send("{ESCAPE}")

EndIf
