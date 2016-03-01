
#include-once


#cs ----------------------------------------------------------------------------

    Title:   genericfunctions to automate window based objects.
	Filename:  genericfunctions.au3
	Description: It contains generic functions which help to automate the window based
	             object and it contains also other fnction which help to tackle other problems.
	Author:   Brijesh Yadav
	Version:  genericfunctions-v.01
	Last Update: 01/29/2015

#ce ----------------------------------------------------------------------------


;===============================================================================
;
; Description : All variables for all function is defined below-
; Author(s):    Brijesh Yadav
;===============================================================================

Dim $function_name = $CmdLine[1]
;Dim $function_name = "ie_savefile_2"

;===============================================================================
;
; Description : Function execution based on parameter
; Author(s):    Brijesh Yadav
;===============================================================================

If $function_name = "1" Then
    MsgBox(0, "Path", "function 1")
	 Dim $window_title_name = $CmdLine[2]
	 MsgBox(0, "Path", $window_title_name)
	;getWindowText($window_title_name)

ElseIf $function_name = "ff_openfile" Then
	; MsgBox(0, "Path", "function 2")
	 ff_openfile()

ElseIf $function_name = "ie_openfile" Then
	 ;MsgBox(0, "Path", "function 3")
	 ie_openfile()

ElseIf $function_name = "ff_fileupload" Then
	 ;MsgBox(0, "Path", "function 4")
     ff_fileupload()

ElseIf $function_name = "ie_savefile_1" Then
	 ;MsgBox(0, "Path", "function 5")
	 ie_savefile_1()

ElseIf $function_name = "ie_savefile_2" Then
	; MsgBox(0, "Path", "function 5")
	 ie_savefile_2()

ElseIf $function_name = "ie_cancelfile" Then
	  ;MsgBox(0, "Path", "function 5")
	 ie_cancelfile()

ElseIf $function_name = "close_browser" Then
	; MsgBox(0, "Path", "function 5")
	 Dim $window_title_name = $CmdLine[2]
	 close_browser($window_title_name)

EndIf


;===============================================================================
;
; Function Name:    getWindowText()
; Description:		Obtain the URL references within a frame by 0-based index
; Parameter(s):
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):  On Success - Returns an object variable pointing to
;                   On Failure - 0  and sets @ERROR = 1
; Author(s):        Dale Hohm
;
;===============================================================================

Func getWindowText($window_title_name)
    ; Retrieve the window text of the active window.
    Local $sText = WinGetText($window_title_name)

    ; Display the window text.
      MsgBox(0, "", $sText)

EndFunc

;===============================================================================
;
; Function Name:    ff_openfile()
; Description:		Obtain the URL references within a frame by 0-based index
; Parameter(s):
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):  On Success - Returns an object variable pointing to
;                   On Failure - 0  and sets @ERROR = 1
; Author(s):        Dale Hohm
;
;===============================================================================

Func ff_openfile()
	Dim $wintitle =$CmdLine[2]
    For $i = 1 To 10
    If WinExists($wintitle) Then
;	   MsgBox(0, "path", $title)
	   WinActivate($wintitle)
	   Sleep(2000)
	   Send("{ENTER}")
       ExitLoop
	 Else
;     MsgBox(0, "path", "not present")
	Sleep(999)
    EndIf
    Next

EndFunc

;===============================================================================
;
; Function Name:    ie_openfile()
; Description:		Obtain the URL references within a frame by 0-based index
; Parameter(s):
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):  On Success - Returns an object variable pointing to
;                   On Failure - 0  and sets @ERROR = 1
; Author(s):        Dale Hohm
;
;===============================================================================

Func ie_openfile()

    ;Dim $title = "https://www.google.com.au/url?sa=t&rct=j&q=&esrc=s&frm=1&source=web&cd=17&ved=0CE8QFjAGOAo&url= - Windows Internet Explorer"

For $i = 1 To 10
	$hIE = WinGetHandle("[Class:IEFrame]")
    If WinExists($hIE) Then
;	   MsgBox(0, "path", "Present")
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
;	   MouseMove($x,$y)
       ExitLoop
	 Else
;       MsgBox(0, "", "The value of $i is: " & $i)
	   Sleep(999)
    EndIf
Next

EndFunc

;===============================================================================
;
; Function Name:    ff_fileupload()
; Description:		Obtain the URL references within a frame by 0-based index
; Parameter(s):
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):  On Success - Returns an object variable pointing to
;                   On Failure - 0  and sets @ERROR = 1
; Author(s):        Dale Hohm
;
;===============================================================================

Func ff_fileupload()
	Dim $filepath = $CmdLine[2]

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

EndFunc

;===============================================================================
;
; Function Name:    ie_savefile_1()
; Description:		It saves the file from IE dialog box
; Parameter(s):
; Author(s):        Dale Hohm
;
;===============================================================================

Func ie_savefile_1()

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

EndFunc

;===============================================================================
;
; Function Name:    ie_savefile_2()
; Description:		It saves the file from IE notification bar
; Parameter(s):
; Author(s):        Brijesh Yasdav
;
;===============================================================================

Func ie_savefile_2()

;Dim $title = "https://www.google.com.au/url?sa=t&rct=j&q=&esrc=s&frm=1&source=web&cd=17&ved=0CE8QFjAGOAo&url= - Windows Internet Explorer"

For $i = 1 To 10
	$wtitl = "View Downloads - Windows Internet Explorer"
	$hIE = WinGetHandle($wtitl)
    If WinExists($hIE) Then
;	   MsgBox(0, "path", "Present")
	   $hCtrl = ControlGetHandle($hIE,"",$hIE)
	   $aPos = ControlGetPos($hIE,"",$hCtrl)
       $x = $aPos[2]-250
       $y = $aPos[3]-35
       ;Use
       WinActivate($hIE)
       ;doesn't work in the background
       ControlClick($hIE,"",$hCtrl,"primary",1,$x,$y)
       ;this only gives focus to the save button
       ControlSend($hIE,"",$hCtrl,"{Enter}")
;	   MouseMove($x,$y)
       Send("{TAB}")
	   Sleep(1000)
       Send("{TAB}")
       Sleep(500)
       Send("{TAB}")
       Sleep(500)
       Send("{ENTER}")
       ExitLoop
	 Else
;       MsgBox(0, "", "The value of $i is: " & $i)
	   Sleep(999)
    EndIf
Next

EndFunc

;===============================================================================
;
; Function Name:    ie_cancelfile()
; Description:		Obtain the URL references within a frame by 0-based index
; Parameter(s):
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):  On Success - Returns an object variable pointing to
;                   On Failure - 0  and sets @ERROR = 1
; Author(s):        Dale Hohm
;
;===============================================================================

Func ie_cancelfile()

;Dim $title = "https://www.google.com.au/url?sa=t&rct=j&q=&esrc=s&frm=1&source=web&cd=17&ved=0CE8QFjAGOAo&url= - Windows Internet Explorer"
For $i = 1 To 10
	$hIE = WinGetHandle("[Class:IEFrame]")
    If WinExists($hIE) Then
;	   MsgBox(0, "path", "Present")
	   $hCtrl = ControlGetHandle($hIE,"","[Class:DirectUIHWND]")
	   $aPos = ControlGetPos($hIE,"",$hCtrl)
       $x = $aPos[2]-100
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

EndFunc

;===============================================================================
;
; Function Name:    close_browser()
; Description:		Obtain the URL references within a frame by 0-based index
; Parameter(s):
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):  On Success - Returns an object variable pointing to
;                   On Failure - 0  and sets @ERROR = 1
; Author(s):        Dale Hohm
;
;===============================================================================

Func close_browser($window_title_name)

 For $i = 1 To 10
	If WinExists($window_title_name) Then
;	   MsgBox(0, "path", $title)
	   WinActivate($window_title_name)
	   WinClose($window_title_name)
       if(WinExists("Internet Explorer")) Then
		   WinActivate("Internet Explorer")
		   Send("{TAB}")
		   Send("{ENTER}")
	   ElseIf WinExists("Confirm close") Then
	       WinActivate("Confirm close")
		   Send("{ENTER}")
	   EndIf
	ExitLoop
	Else
;     MsgBox(0, "path", "not present")
	  Sleep(999)
    EndIf
Next

EndFunc

