Dim mainString, subString

mainString = "Kwaku is a good boy"
subString = "boy"

Set regEx = new RegExp
regEx.Pattern = subString

Set match = regEx.Execute(mainString)

If match.Count > 0 Then
  MsgBox "The substring is found within the mainstring"
Else
  MsgBox "The substring is not found within the mainstring"
End If