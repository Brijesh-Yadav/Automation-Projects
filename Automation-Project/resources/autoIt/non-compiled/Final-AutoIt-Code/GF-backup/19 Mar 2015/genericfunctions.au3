
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
;Dim $function_name = "SavepdfFile_with_saveAs_option"

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
; Function Name:    ie_savefile_1()
; Description:		It saves the file from IE dialog box
; Parameter(s):
; Author(s):        Brijesh Yadav
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
; Description:
; Parameter(s):
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):
; Author(s):    Brijesh Yadav
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

