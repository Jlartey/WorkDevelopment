Dim num1, num2, result, operator

num1 = InputBox("Enter the first number:")

operator = InputBox("Enter the operator (+, -, *, /):")

num2 = InputBox("Enter the second number:")

num1 = CInt(num1)
num2 = Cint(num2)

Select Case operator
  Case "+"
    result = num1 + num2
  Case "-"
    result = num1 = num2
  Case "*"
    result = num1 * num2
  Case "/"
    If num2 <> 0 Then
      result = num1 / num2
    Else
      MsgBox "Error: Invalid operator!"
      WScript.Quit
    End If
  Case Else
    MsgBox "Error: Invalid operator!"
    WScript.Quit
End Select

' Display result 
MsgBox "Result: " & num1 & " " & operator & " " & num2 & " is " & result 