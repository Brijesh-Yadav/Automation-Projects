
#include-once

#cs ----------------------------------------------------------------------------

    Title:   genericfunctions to automate window based objects.
	Filename:  genericfunctions.au3
	Description: It contains generic functions which help to automate the window based
	             object and it contains also other fnction which help to tackle other problems.
	Author:   Brijesh Yadav
	Version:  genericfunctions-v.02
	Last Update: 01/29/2015

#ce ----------------------------------------------------------------------------

;===============================================================================
;
; Description : All variables for all function is defined below-
; Author(s):    Brijesh Yadav
;===============================================================================

Dim $function_name = $CmdLine[1]
;Dim $function_name = "ie_fileupload"

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

ElseIf $function_name = "MicrosotExcel_formatsupport" Then
	; MsgBox(0, "Path", "function 5")
	 MicrosotExcel_formatsupport()

ElseIf $function_name = "EnterValue" Then
	; MsgBox(0, "Path", "function 5")
	 EnterValue()

ElseIf $function_name = "ie_savefile_dialogwindow" Then
	; MsgBox(0, "Path", "function 5")
	 ie_savefile_dialogwindow()

ElseIf $function_name = "ie_openfile_dialogwindow" Then
	; MsgBox(0, "Path", "function 5")
	 ie_openfile_dialogwindow()

ElseIf $function_name = "SavepdfFile_with_saveAs_option" Then
	; MsgBox(0, "Path", "function 5")
	 SavepdfFile_with_saveAs_option()

ElseIf $function_name = "mouse_handle_ie_notificationbar_openfile" Then
	; MsgBox(0, "Path", "function 5")
	 mouse_handle_ie_notificationbar_openfile()

ElseIf $function_name = "mouse_handle_ie_notificationbar_savefile" Then
	; MsgBox(0, "Path", "function 5")
	 mouse_handle_ie_notificationbar_savefile()

ElseIf $function_name = "mouse_handle_ie_notificationbar_cancel" Then
	; MsgBox(0, "Path", "function 5")
	 mouse_handle_ie_notificationbar_cancel()

ElseIf $function_name = "mouse_handle_ie_notificationbar_Yes" Then
	; MsgBox(0, "Path", "function 5")
	 mouse_handle_ie_notificationbar_Yes()

ElseIf $function_name = "mouse_handle_ie_notificationbar_No" Then
	; MsgBox(0, "Path", "function 5")
	 mouse_handle_ie_notificationbar_No()

ElseIf $function_name = "ie_fileupload" Then
	; MsgBox(0, "Path", "function 5")
	 ie_fileupload()

EndIf


;===============================================================================
;
; Function Name:    getWindowText()
; Description:		It returns the present text on window
; Parameter(s):
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):
; Author(s):        Brijesh Yadav
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
; Description:
; Parameter(s):
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):
; Author(s):    Brijesh Yadav
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
; Description:
; Parameter(s):
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):
; Author(s):
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
; Description:
; Parameter(s):
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):
; Author(s):        Brijesh Yadav
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
; Function Name:    ie_fileupload()
; Description:
; Parameter(s):
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):
; Author(s):        Brijesh Yadav
;
;===============================================================================

Func ie_fileupload()
	Dim $filepath = $CmdLine[2]
	;$filepath = "Test/testfile"
	$tilte = "Choose File to Upload"
For $i = 1 To 10
    If WinExists($tilte) Then
	 ;  MsgBox(0, "path", $filepath)
	   WinActivate($tilte)
	   ControlSetText("","","[CLASS:Edit]",$filepath)
	   ControlClick($tilte,"","[CLASS:Button; INSTANCE:1]")
       ExitLoop
	 Else
	;MsgBox(0, "path", "Not present")
	Sleep(999)
    EndIf
Next

EndFunc


;===============================================================================
;
; Function Name:    ie_savefile_1()
; Description:		It saves the file from IE dialog box
; Parameter(s):
; Author(s):        Brijesh Yadav
;
;===============================================================================

Func ie_savefile_1()

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
; Description:
; Parameter(s):
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):
; Author(s):    Brijesh Yadav
;
;===============================================================================

Func ie_cancelfile()

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
; Description:
; Parameter(s):
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):
; Author(s):    Brijesh Yadav
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

;===============================================================================
;
; Function Name:    MicrosotExcel_formatsupport()
; Description:
; Parameter(s):
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):
; Author(s):    Brijesh Yadav
;
;===============================================================================

Func MicrosotExcel_formatsupport()

$window_title_name = "Microsoft Excel"
Dim $btnname = $CmdLine[2]
Dim $btnvalue

If $btnname = "Yes" Then
    $btnvalue = "&Yes"
ElseIf $btnname = "No" Then
    $btnvalue = "&No"
EndIf

 For $i = 1 To 10
	If WinExists($window_title_name) Then
	   ;MsgBox(0, "path", "Present")
	   WinActivate($window_title_name)
	   Local $sText = WinGetText($window_title_name)
       ;MsgBox(0, "path", $sText)
       ControlClick($window_title_name,"",$btnvalue)
	ExitLoop
	Else
      ;MsgBox(0, "path", "not present")
	  Sleep(999)
    EndIf
Next

EndFunc

;===============================================================================
;
; Function Name:    EnterValue()
; Description:
; Parameter(s):
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):
; Author(s):    Brijesh Yadav
;
;===============================================================================

Func EnterValue()
Dim $value = $CmdLine[2]
Send($value);
EndFunc

;===============================================================================
;
; Function Name:    ie_savefile_dialogwindow()
; Description:
; Parameter(s):
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):
; Author(s):    Brijesh Yadav
;
;===============================================================================

Func ie_savefile_dialogwindow()
$bvalue= "False"
$window_title_name = "Internet Explorer"
 For $i = 1 To 10
	If WinExists($window_title_name) Then
       if(WinExists($window_title_name)) Then
		   WinActivate($window_title_name)
;		   MsgBox(0, "window", "present")
           ControlClick($window_title_name,"","&Save")
;		   Send("{TAB}")
;		   Send("{ENTER}")
;		   Send("{ENTER}")
		   $bvalue="true"
	       ExitLoop
	   EndIf
	Else
;     MsgBox(0, "path", "not present")
	  Sleep(999)
    EndIf
 Next

If $bvalue="False" Then
;    MsgBox(0, "Alert message", "window not present")
    ie_savefile_1()
EndIf

EndFunc

;===============================================================================
;
; Function Name:    ie_openfile_dialogwindow()
; Description:
; Parameter(s): &Open
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):
; Author(s):    Brijesh Yadav
;
;===============================================================================

Func ie_openfile_dialogwindow()
$bvalue= "False"
$window_title_name = "Internet Explorer"
 For $i = 1 To 10
	If WinExists($window_title_name) Then
       if(WinExists($window_title_name)) Then
		   WinActivate($window_title_name)
;		   MsgBox(0, "window", "present")
           ControlClick($window_title_name,"","&Open")
		   $bvalue="true"
	       ExitLoop
	   EndIf
	Else
;     MsgBox(0, "path", "not present")
	  Sleep(999)
    EndIf
 Next

If $bvalue="False" Then
ie_openfile()
EndIf

EndFunc

;===============================================================================
;
; Function Name:    SavepdfFile_with_saveAs_option()
; Description:
; Parameter(s):
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):
; Author(s):    Brijesh Yadav
;
;===============================================================================

Func SavepdfFile_with_saveAs_option()
Dim $value = $CmdLine[2]
$window_title_name = "[CLASS:AcrobatSDIWindow]"
 For $i = 1 To 10
	If WinExists($window_title_name) Then
       if(WinExists($window_title_name)) Then
		   WinActivate($window_title_name)
;		   MsgBox(0, "window", "present")
           Send("{ALT}")
		   Sleep(200)
           Send("{DOWN}")
		   Sleep(200)
		   Send("{DOWN}")
		   Sleep(200)
		   Send("{DOWN}")
		   Sleep(200)
		   Send("{RIGHT}")
		   Sleep(200)
		   Send("{ENTER}")
           Sleep(400)
           ControlSetText("","","[CLASS:Edit]",$value)
		   Sleep(300)
		   Send("{ENTER}")
		   $aletwind = "Confirm Save As"
		   If WinExists($aletwind) Then
;			  MsgBox(0, "window", "present")
			  WinActivate($aletwind)
			  ControlClick($aletwind,"","[CLASS:Button; INSTANCE:1]")
			EndIf
	       ExitLoop
	   EndIf
	Else
;      MsgBox(0, "path", "not present")
	  Sleep(999)
    EndIf
 Next

EndFunc

;===============================================================================
;
; Function Name:    mouse_handle_ie_notificationbar_openfile()
; Description:
; Parameter(s):
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):
; Author(s):    Brijesh Yadav
;
;===============================================================================

Func mouse_handle_ie_notificationbar_openfile()

$window_title_name = "[CLASS:IEFrame]"
 For $i = 1 To 2
	If WinExists($window_title_name) Then
	    WinActivate($window_title_name)
		Local $windHandle=WinGetHandle("[Class:IEFrame]", "")
        Local $winTitle = "[HANDLE:" & $windHandle & "]";
;		MsgBox(0, "TITLE", $winTitle)
;		get coordinates of default HWND
		Local $aPos=WinGetPos($winTitle)
;		MouseMove(100,490)
		; Display the array values returned by WinGetPos.
;		MsgBox(0, "", "X-Pos: " & $aPos[0] & @CRLF & _
;				"Y-Pos: " & $aPos[1] & @CRLF & _
;				"Width: " & $aPos[2] & @CRLF & _
;				"Height: " & $aPos[3])
		Local $objPos=ControlGetPos("","","[CLASS:DirectUIHWND; INSTANCE:1]")
;		MsgBox(0, "", "X-Pos: " & $objPos[0] & @CRLF & _
;				"Y-Pos: " & $objPos[1] & @CRLF & _
;				"Width: " & $objPos[2] & @CRLF & _
;				"Height: " & $objPos[3])

		$x_cord = $aPos[0]+$objPos[0]+$objPos[2]-210
		$y_cord = $aPos[1]+$objPos[1]+$objPos[3]-3

		MouseMove($x_cord,$y_cord)
		Sleep(300)
		MouseClick("left",$x_cord,$y_cord)

        ExitLoop
	Else
;      MsgBox(0, "path", "not present")
	  Sleep(999)
    EndIf
 Next

EndFunc

;===============================================================================
;
; Function Name:    mouse_handle_ie_notificationbar_savefile()
; Description:
; Parameter(s):
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):
; Author(s):    Brijesh Yadav
;
;===============================================================================

Func mouse_handle_ie_notificationbar_savefile()

$window_title_name = "[CLASS:IEFrame]"
 For $i = 1 To 2
	If WinExists($window_title_name) Then
	    WinActivate($window_title_name)
		Local $windHandle=WinGetHandle("[Class:IEFrame]", "")
        Local $winTitle = "[HANDLE:" & $windHandle & "]";
;		MsgBox(0, "TITLE", $winTitle)
;		get coordinates of default HWND
		Local $aPos=WinGetPos($winTitle)
;		MouseMove(100,490)
		; Display the array values returned by WinGetPos.
;		MsgBox(0, "", "X-Pos: " & $aPos[0] & @CRLF & _
;				"Y-Pos: " & $aPos[1] & @CRLF & _
;				"Width: " & $aPos[2] & @CRLF & _
;				"Height: " & $aPos[3])
		Sleep(400)
;		MouseMove($aPos[0],$aPos[1])
		Local $objPos=ControlGetPos("","","[CLASS:DirectUIHWND; INSTANCE:1]")
;		MsgBox(0, "", "X-Pos: " & $objPos[0] & @CRLF & _
;				"Y-Pos: " & $objPos[1] & @CRLF & _
;				"Width: " & $objPos[2] & @CRLF & _
;				"Height: " & $objPos[3])

		$x_cord = $aPos[0]+$objPos[0]+$objPos[2]-135
		$y_cord = $aPos[1]+$objPos[1]+$objPos[3]-5

		MouseMove($x_cord,$y_cord)
		Sleep(300)
		MouseClick("left",$x_cord,$y_cord)
        ExitLoop
	Else
;      MsgBox(0, "path", "not present")
	  Sleep(999)
    EndIf
 Next

EndFunc


;===============================================================================
;
; Function Name:    mouse_handle_ie_notificationbar_cancel()
; Description:
; Parameter(s):
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):
; Author(s):    Brijesh Yadav
;
;===============================================================================

Func mouse_handle_ie_notificationbar_cancel()

$window_title_name = "[CLASS:IEFrame]"
 For $i = 1 To 2
	If WinExists($window_title_name) Then
	    WinActivate($window_title_name)
		Local $windHandle=WinGetHandle("[Class:IEFrame]", "")
        Local $winTitle = "[HANDLE:" & $windHandle & "]";
;		MsgBox(0, "TITLE", $winTitle)
;		get coordinates of default HWND
		Local $aPos=WinGetPos($winTitle)
;		MouseMove(100,490)
		; Display the array values returned by WinGetPos.
;		MsgBox(0, "", "X-Pos: " & $aPos[0] & @CRLF & _
;				"Y-Pos: " & $aPos[1] & @CRLF & _
;				"Width: " & $aPos[2] & @CRLF & _
;				"Height: " & $aPos[3])
		Local $objPos=ControlGetPos("","","[CLASS:DirectUIHWND; INSTANCE:1]")
;		MsgBox(0, "", "X-Pos: " & $objPos[0] & @CRLF & _
;				"Y-Pos: " & $objPos[1] & @CRLF & _
;				"Width: " & $objPos[2] & @CRLF & _
;				"Height: " & $objPos[3])

		$x_cord = $aPos[0]+$objPos[0]+$objPos[2]-55
		$y_cord = $aPos[1]+$objPos[1]+$objPos[3]-3

		MouseMove($x_cord,$y_cord)
		Sleep(300)
		MouseClick("left",$x_cord,$y_cord)

		ExitLoop
	Else
;      MsgBox(0, "path", "not present")
	  Sleep(999)
    EndIf
 Next

EndFunc

;===============================================================================
;
; Function Name:    mouse_handle_ie_notificationbar_Yes()
; Description:
; Parameter(s):
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):
; Author(s):    Brijesh Yadav
;
;===============================================================================

Func mouse_handle_ie_notificationbar_Yes()

$window_title_name = "[CLASS:IEFrame]"
 For $i = 1 To 2
	If WinExists($window_title_name) Then
	    WinActivate($window_title_name)
		Local $windHandle=WinGetHandle("[Class:IEFrame]", "")
        Local $winTitle = "[HANDLE:" & $windHandle & "]";
;		MsgBox(0, "TITLE", $winTitle)
;		get coordinates of default HWND
		Local $aPos=WinGetPos($winTitle)
;		MouseMove(100,490)
		; Display the array values returned by WinGetPos.
;		MsgBox(0, "", "X-Pos: " & $aPos[0] & @CRLF & _
;				"Y-Pos: " & $aPos[1] & @CRLF & _
;				"Width: " & $aPos[2] & @CRLF & _
;				"Height: " & $aPos[3])
		Local $objPos=ControlGetPos("","","[CLASS:DirectUIHWND; INSTANCE:1]")
;		MsgBox(0, "", "X-Pos: " & $objPos[0] & @CRLF & _
;				"Y-Pos: " & $objPos[1] & @CRLF & _
;				"Width: " & $objPos[2] & @CRLF & _
;				"Height: " & $objPos[3])
		MouseMove($aPos[0]+$objPos[0]+$objPos[2]-130,$aPos[1]+$objPos[1]+$objPos[3]-22)
		Sleep(300)
		MouseClick("left",$aPos[0]+$objPos[0]+$objPos[2]-130,$aPos[1]+$objPos[1]+$objPos[3]-22)
        ExitLoop
	Else
;      MsgBox(0, "path", "not present")
	  Sleep(999)
    EndIf
 Next

EndFunc

;===============================================================================
;
; Function Name:    mouse_handle_ie_notificationbar_No()
; Description:
; Parameter(s):
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):
; Author(s):    Brijesh Yadav
;
;===============================================================================

Func mouse_handle_ie_notificationbar_No()

$window_title_name = "[CLASS:IEFrame]"
 For $i = 1 To 2
	If WinExists($window_title_name) Then
	    WinActivate($window_title_name)
		Local $windHandle=WinGetHandle("[Class:IEFrame]", "")
        Local $winTitle = "[HANDLE:" & $windHandle & "]";
;		MsgBox(0, "TITLE", $winTitle)
;		get coordinates of default HWND
		Local $aPos=WinGetPos($winTitle)
;		MouseMove(100,490)
		; Display the array values returned by WinGetPos.
;		MsgBox(0, "", "X-Pos: " & $aPos[0] & @CRLF & _
;				"Y-Pos: " & $aPos[1] & @CRLF & _
;				"Width: " & $aPos[2] & @CRLF & _
;				"Height: " & $aPos[3])
		Local $objPos=ControlGetPos("","","[CLASS:DirectUIHWND; INSTANCE:1]")
;		MsgBox(0, "", "X-Pos: " & $objPos[0] & @CRLF & _
;				"Y-Pos: " & $objPos[1] & @CRLF & _
;				"Width: " & $objPos[2] & @CRLF & _
;				"Height: " & $objPos[3])
		MouseMove($aPos[0]+$objPos[0]+$objPos[2]-70,$aPos[1]+$objPos[1]+$objPos[3]-22)
		Sleep(300)
		MouseClick("left",$aPos[0]+$objPos[0]+$objPos[2]-70,$aPos[1]+$objPos[1]+$objPos[3]-22)
        ExitLoop
	Else
;       MsgBox(0, "path", "not present")
	  Sleep(999)
    EndIf
 Next

EndFunc

;===============================================================================
;
; Function Name:    WriteToLog()
; Description:
; Parameter(s):
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):
; Author(s):    Brijesh Yadav
;
;===============================================================================

Func WriteToLog($Value)

;	WriteToLog("This is test data, only for testing purpose for autoit logger"&@LF&"")
;	read file to get the logfile path
	Local Const $sFilePath = "logfilepath.txt"
	$file = FileOpen($sFilePath, 0)
    $FileContent = FileRead($file)
    MsgBox(0, "Content:", $FileContent)
    FileClose($file)

;	write log into file
	$FileHandle = FileOpen($FileContent, 1) ; 1 = append mode
	If $FileHandle <> -1 Then
		FileWrite($FileHandle, $Value)
	EndIf
	FileClose($FileHandle)

EndFunc

