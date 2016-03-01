
#include-once

#cs ----------------------------------------------------------------------------

    Title:   eIPE autoit function
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
 ;Dim $function_name = "accept_message"

;===============================================================================
;
; Description : Function execution based on parameter
; Author(s):    Brijesh Yadav
;===============================================================================

If $function_name = "accept_message" Then
    ;MsgBox(0, "Path", "function 1")
	accept_message()

EndIf

;===============================================================================
;
; Function Name:    select_group_bench_report()
; Description:		Obtain the URL references within a frame by 0-based index
; Parameter(s):
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):  On Success - Returns an object variable pointing to
;                   On Failure - 0  and sets @ERROR = 1
; Author(s):        Dale Hohm
;
;===============================================================================

Func accept_message()
	Dim $title = "Message from webpage"
    For $i = 1 To 10
    If WinExists($title) Then
	  ; MsgBox(0, "path", "window present!!")
	   WinActivate($title)
	   ;Send("{TAB}")
	   Send("{ENTER}")
       ExitLoop
	 Else
;     MsgBox(0, "path", "not present")
	Sleep(999)
    EndIf
Next

EndFunc




