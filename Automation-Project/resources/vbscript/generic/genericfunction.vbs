
'VBSTART
Set objArgs = Wscript.Arguments
function_name = objArgs(0)
'function_name="1"
Dim IE     ' IE browser instance
Dim checkval   ' validing of ie object
urlcsvfilepath = currentDirectory & "IEurls.csv" ' file path for url
Dim urladdrss
Dim urllist
Dim valopr   ' validing operation


'Attaches to an existing instance of IE with matching URL
Sub GetIE(URL)
  Dim objInstances, objIE
  Set objInstances = CreateObject("Shell.Application").windows
  If objInstances.Count > 0 Then '/// make sure we have instances open.
    For Each objIE In objInstances
     If InStr(objIE.LocationURL,URL) > 0 then
       Set IE = objIE
       Set wshShell = CreateObject("WScript.Shell")
       wshShell.AppActivate(IE)
	   checkval = 1
       'msgbox("running")
     End if
    Next
  End if
End Sub

'===============================================================================
' Function Name:    selectValueFromDropdown()
' Description:		Obtain the URL references within a frame by 0-based index
' Parameter(s):
' Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
' Return Value(s):  On Success - Returns an object variable pointing to
'                   On Failure - 0  and sets @ERROR = 1
' Author(s):        Brijesh yadav
'===============================================================================

Sub selectValueFromDropdown(index)
Do
    WScript.Sleep 100
Loop While IE.ReadyState < 4 And IE.Busy

' get first HTMLSelectElement object:
Set e = IE.Document.getElementsByTagName("select")(0)

' just for undestanding...
'MsgBox e.Options(e.selectedIndex).Value '-> "D"
'MsgBox e.Options(e.selectedIndex).Text  '-> "Audi"

'select first option:
e.selectedIndex = index
End Sub

'===============================================================================
' Function Name:    clickOnObject()
' Description:		Obtain the URL references within a frame by 0-based index
' Parameter(s):
' Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
' Return Value(s):  On Success - Returns an object variable pointing to
'                   On Failure - 0  and sets @ERROR = 1
' Author(s):        Brijesh yadav
'===============================================================================

Sub clickOnObject(name)
 IE.Document.getElementsByName(name).Item(0).Click
End Sub

'===============================================================================
' Function Name:    closeIE()
' Description:		Obtain the URL references within a frame by 0-based index
' Parameter(s):
' Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
' Return Value(s):  On Success - Returns an object variable pointing to
'                   On Failure - 0  and sets @ERROR = 1
' Author(s):        Brijesh yadav
'===============================================================================

Sub closeIE()
  IE.Quit
End Sub

'===============================================================================
' Function Name:    sleep()
' Description:		Obtain the URL references within a frame by 0-based index
' Parameter(s):
' Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
' Return Value(s):  On Success - Returns an object variable pointing to
'                   On Failure - 0  and sets @ERROR = 1
' Author(s):        Brijesh yadav
'===============================================================================

Sub sleep(time)
  WScript.Sleep 100*time
End Sub

'===============================================================================
' Function Name:    deleteFile()
' Description:		Obtain the URL references within a frame by 0-based index
' Parameter(s):
' Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
' Return Value(s):  On Success - Returns an object variable pointing to
'                   On Failure - 0  and sets @ERROR = 1
' Author(s):        Brijesh yadav
'===============================================================================

Sub deleteFile(filename)
   Set fso = createobject("Scripting.FileSystemObject")
   if fso.FileExists(filename) then
        fso.DeleteFile filename 
   Else 
    '   Wscript.Echo "File does not exist!!"
   end if
End Sub

'===============================================================================
' Function Name:    returnUrl()
' Description:		Obtain the URL references within a frame by 0-based index
' Parameter(s):
' Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
' Return Value(s):  On Success - Returns an object variable pointing to
'                   On Failure - 0  and sets @ERROR = 1
' Author(s):        Brijesh yadav
'===============================================================================

Sub returnUrl(index)

deleteFile(urlcsvfilepath)    ' delete existing file 
sleep(3)
Set objShell = CreateObject("Shell.Application")
Set objShellWindows = objShell.Windows
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set urllist = CreateObject("System.Collections.ArrayList")

If objShellWindows.Count = 0 Then 
    WScript.Echo "Can not find any opend tabs in Internet Explorer." 
Else 
    Set objExportFile = objFSO.CreateTextFile(currentDirectory & "IEurls.csv", ForWriting, True) 
    objExportFile.WriteLine "URL" & "," & "Title" 
     
    For Each objItem In objShellWindows 
        FileFullName = objItem.FullName 
        objFile = objFSO.GetFile(objItem.FullName) 
        objFileName = objFSO.GetFileName(objFile) 
         
        If LCase(objFileName) = "iexplore.exe" Then 
            LocationURL = objItem.LocationURL 
            LocationName = objItem.LocationName 
            objExportFile.WriteLine LocationURL & "," & LocationName 
            'urladdrss = LocationURL
			urllist.Add LocationURL
        End If 
    Next 
    ' wscript.echo join(urllist.ToArray(), ", ")
	'  wscript.Echo "Index count : " & urllist.count
	If index <= urllist.count  Then
	   valopr = 1
       urladdrss = urllist.Item(index-1)
    Else 
	   wscript.Echo "Specified index for browser instance not found"
	End IF
     
    'WScript.Echo "Successfully generated IEurls.csv on " & currentDirectory 
End If
End Sub

'===============================================================================
' Function Name:    arraylistExample()
' Description:		Obtain the URL references within a frame by 0-based index
' Parameter(s):
' Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
' Return Value(s):  On Success - Returns an object variable pointing to
'                   On Failure - 0  and sets @ERROR = 1
' Author(s):        Brijesh yadav
'===============================================================================

Sub arraylistExample()
dim list
Set list = CreateObject("System.Collections.ArrayList")
list.Add "Banana"
list.Add "Apple"
list.Add "Pear"
list.Sort
list.Reverse
wscript.echo list.Count                 ' --> 3
wscript.echo list.Item(0)               ' --> Pear
wscript.echo list.IndexOf("Apple", 0)   ' --> 2
wscript.echo join(list.ToArray(), ", ") ' --> Pear, Banana, Apple
End Sub

'===============================================================================
' Function Name:    ClickLink()
' Description:		Obtain the URL references within a frame by 0-based index
' Parameter(s):
' Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
' Return Value(s):  On Success - Returns an object variable pointing to
'                   On Failure - 0  and sets @ERROR = 1
' Author(s):        Brijesh yadav
'===============================================================================

'clicks specified link
Sub ClickLink(linktext)
  Dim anchors
  Dim ItemNr

  Set anchors = IE.document.getElementsbyTagname("a")
  For ItemNr = 0 to anchors.length - 1
    If anchors.Item(ItemNr).innertext = linktext Then
     anchors.Item(ItemNr).click
   End If
  next

  do while IE.Busy
  loop
End Sub

'VBEND


'===============================================================================
' Code Execution start here
' WScript.Echo urlcsvfilepath

'===============================================================================



If function_name=1 Then
    returnUrl(1)
	If valopr=1 Then 
	   GetIE(urladdrss)
       If checkval=1 Then 
     'select country
      selectValueFromDropdown(21)
      'click on select country button
      clickOnObject("load")
      'sleep for 60 sec
      sleep(50)
      'closing ie instance
       closeIE()
       End If
   Else 
    'wscript.echo "required IE instance not found!!"
   End If
   
ElseIf function_name=2 Then

    returnUrl(1)
	If valopr=1 Then 
	   GetIE(urladdrss)
       If checkval=1 Then 
          ClickLink("WBEG-Online")
       End If
   Else 
    'wscript.echo "required IE instance not found!!"
   End If

End If 




