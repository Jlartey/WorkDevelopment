Dim num1, num2, sum

' Prompt the user to enter the first number
num1 = InputBox("Enter the first number:")

' Prompt the user to enter the second number
num2 = InputBox("Enter the second number:")

' Convert the input strings to numbers
num1 = CDbl(num1)
num2 = CDbl(num2)

' Add the numbers
sum = num1 + num2

' Display the result
MsgBox "The sum of " & num1 & " and " & num2 & " is: " & sum

