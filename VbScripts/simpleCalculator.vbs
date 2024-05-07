Dim num1, num2, result, operator

' Prompt the user to enter the first number
num1 = InputBox("Enter the first number:")

' Prompt the user to enter the operator (+, -, *, /)
operator = InputBox("Enter the operator (+, -, *, /):")

' Prompt the user to enter the second number
num2 = InputBox("Enter the second number:")

' Convert the input strings to numbers
num1 = CDbl(num1)
num2 = CDbl(num2)

' Perform the calculation based on the operator
Select Case operator
    Case "+"
        result = num1 + num2
    Case "-"
        result = num1 - num2
    Case "*"
        result = num1 * num2
    Case "/"
        If num2 <> 0 Then
            result = num1 / num2
        Else
            MsgBox "Error: Division by zero!"
            WScript.Quit
        End If
    Case Else
        MsgBox "Error: Invalid operator!"
        WScript.Quit
End Select

' Display the result
MsgBox "Result: " & num1 & " " & operator & " " & num2 & " = " & result
