
Response.Buffer = True
For i = 1 to 10
    Response.Write "Counter: " & i & "<br>"
    Response.Flush
    ' Simulate some processing delay
    WScript.Sleep 1000  ' Sleep for 1 second
Next

