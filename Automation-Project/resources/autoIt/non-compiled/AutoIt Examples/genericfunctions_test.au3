
#include-once
#include "IE.au3"

#cs ----------------------------------------------------------------------------

    Title:   genericfunctions to automate window based objects.
	Filename:  genericfunctions.au3
	Description: It contains generic functions which help to automate the window based
	             objects and also other complex window related objects.
	Author:   Brijesh Yadav
	Version:  genericfunctions-v.01
	Last Update: 01/29/2015

#ce ----------------------------------------------------------------------------


;===============================================================================
;
; Description : All variables for all functions are defined below-
; Author(s):    Brijesh Yadav
;===============================================================================

Dim $function_name = "5"; ; function name which has to be executed
Dim $window_title_name = "Function WinKill - Mozilla Firefox";
Dim $return_window_txt ; return window text contains all control

;===============================================================================
;
; Description : Function execution based on parameter
; Author(s):    Brijesh Yadav
;===============================================================================

If $function_name = "1" Then
    MsgBox(0, "Path", "function 1")
    getWindowText()

ElseIf $function_name = "2" Then
	 MsgBox(0, "Path", "function 2")
    SelectValuefromComboxbox()

ElseIf $function_name = "3" Then
	 MsgBox(0, "Path", "function 3")
	 SelectValuefromComboxbox2()

ElseIf $function_name = "4" Then
	 MsgBox(0, "Path", "function 4")
	 SelectValuefromComboxbox3()

ElseIf $function_name = "5" Then
	 MsgBox(0, "Path", "function 4")
	 close_browser($window_title_name)

EndIf


;===============================================================================
;
; Function Name:    _IEFrameGetSrcByIndex()
; Description:		Obtain the URL references within a frame by 0-based index
; Parameter(s):
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):  On Success - Returns an object variable pointing to
;                   On Failure - 0  and sets @ERROR = 1
; Author(s):        Dale Hohm
;
;===============================================================================
;
Func getWindowText()
    ; Retrieve the window text of the active window.
    Local $sText = WinGetText($window_title_name)
	$return_window_txt = $sText
    ; Display the window text.
    MsgBox(0, "", $return_window_txt)

EndFunc

Func SelectValuefromComboxbox()
	 WinActivate($window_title_name)
	; Local $oIE = _IE_Example($window_title_name)
     Local $oForm = _IEFormGetObjByName($window_title_name, "aspnetForm")
	  MsgBox(0, "Path", $oForm)
     Local $oText = _IEFormElementGetObjByName($oForm, "search-input")
    _IEFormElementSetValue($oText, "Hey! This works!")

EndFunc

; Open a browser with the form example, get reference to form, get reference
; to select multiple element, cycle 5 times selecting and then deselecting
; options byValue, byText and byIndex.

Func SelectValuefromComboxbox2()

Local $oIE = _IE_Example("form")
Local $oForm = _IEFormGetObjByName($oIE, "myForm")
Local $oSelect = _IEFormElementGetObjByName($oForm, "multipleSelectExample")
MsgBox(0, "Path", $oSelect)
_IEAction($oSelect, "focus")
For $i = 1 To 3
    _IEFormElementOptionSelect($oSelect, "Carlos", 1, "byText")
    Sleep(1000)
    _IEFormElementOptionSelect($oSelect, "Name2", 1, "byValue")
    Sleep(1000)
    _IEFormElementOptionSelect($oSelect, 5, 1, "byIndex")
    Sleep(1000)
    _IEFormElementOptionSelect($oSelect, "Carlos", 0, "byText")
    Sleep(1000)
    _IEFormElementOptionSelect($oSelect, "Name2", 0, "byValue")
    Sleep(1000)
    _IEFormElementOptionSelect($oSelect, 5, 0, "byIndex")
    Sleep(1000)
Next
;_IEQuit($oIE)

EndFunc


Func SelectValuefromComboxbox3()

Local $oIE = _IECreate("http://www.indianrail.gov.in/between_Imp_Stations.html")
;MsgBox(0, "Path", $oIE)
Local $oForm = _IEFormGetObjByName($oIE, "myForm")
Local $oSelect = _IEFormElementGetObjByName($oForm, "lccp_src_stncode")
_IEAction($oSelect, "focus")
Sleep(1000)
_IEFormElementOptionSelect($oSelect, "ADI", 1, "byValue")
Local $oSelect2 = _IEFormElementGetObjByName($oForm, "lccp_dstn_stncode")
_IEAction($oSelect2, "focus")
Sleep(1000)
_IEFormElementOptionSelect($oSelect2, "AWY", 1, "byValue")
_IEQuit($oIE)

EndFunc

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
