Dim num1, num2, num3, sum

' Prompt the user to enter the first number
num1 = InputBox("Enter the first number:")

' Prompt the user to enter the second number  
num2 = InputBox("Enter the second number")

' Prompt the user to enter the third number  
num3 = InputBox("Enter the third number")

num1 = CDbl(num1)
num2 = CDbl(num2)
num3 = CDbl(num3)

sum = num1 + num2 + num3

MsgBox "The sum of " & num1 & " and " & num2 & " and " & num3 & " is: " & sum
