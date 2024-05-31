 Sub greet()
        response.write "Hello World!!<br/>"
    End Sub
    
    Function greeting(name)
        greeting = "Hello " & name & "!!"
    End Function
    
    response.write "<img src=>"
    
    greet
    
    response.write greeting("Joseph")
    
    Sub C20240410()
    Dim rst, sql
    
    Set rst = CreateObject("ADODB.RecordSet")
    'patId' = ""
    
    sql = "SELECT TOP 50 * FROM Patient " 'Where PatientID='" & patId & "' "
    
    
    With rst
    .Open sql, conn, 3, 4
      If .RecordCount > 0 Then
      .movefirst
      
      response.write "<table>"
        response.write "<tr>"
        response.write "<td>##</td>"
        response.write "<td>Name</td>"
        response.write "</tr>"
      
      Do Until .EOF
        response.write "<tr>"
        response.write "<td> & </td>"
        response.write "<td>"
        response.write "<"
      
      
    
    End Sub
