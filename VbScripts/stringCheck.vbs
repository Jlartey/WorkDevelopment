Dim mainString, subString

mainString = "This is the main string"
subString = "@gra.gov.gh"

Set regex = New RegExp
regex.Pattern = subString

Set match = regex.Execute(mainString)

If match.Count > 0 Then
    MsgBox "The main string contains the substring."
Else
    MsgBox "The main string does not contain the substring."
End If

' The id needed: tdLabelInpClearanceMemo2
' email = Trim(Request("InpClearanceMemo2"))
