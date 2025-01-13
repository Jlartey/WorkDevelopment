Dim otherDate, today, difference

' Specify the other date
otherDate = #11/20/2024# ' Format: mm/dd/yyyy

' Get today's date
today = Date()

' Calculate the difference in days
difference = DateDiff("d", otherDate, today)

' Output the result
If difference > 0 Then
    WScript.Echo "The date " & otherDate & " was " & difference & " days ago."
ElseIf difference < 0 Then
    WScript.Echo "The date " & otherDate & " is " & Abs(difference) & " days from today."
Else
    WScript.Echo "The date " & otherDate & " is today!"
End If
