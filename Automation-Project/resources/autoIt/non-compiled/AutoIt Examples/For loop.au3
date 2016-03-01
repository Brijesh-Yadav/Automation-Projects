
For $i = 1 To 10
    If $i = 7 Then
        ContinueLoop ; Skip displaying the message box when $i is equal to 7.
    EndIf
    MsgBox(0, "", "The value of $i is: " & $i)
Next




