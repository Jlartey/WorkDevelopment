'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim specialistGroup
specialistGroup = Trim(Request.querystring("PrintFilter"))

Styles
displayPhoneNumbers

Sub displayPhoneNumbers()
    Dim sql, rst
    
    Set rst = CreateObject("ADODB.Recordset")
    
    sql = "SELECT DISTINCT "
    sql = sql & "STRING_AGG(CAST(CASE WHEN LEN(ResidencePhone) = 10 THEN ResidencePhone "
    sql = sql & "ELSE NULL END AS NVARCHAR(MAX)), ', ') AS PhoneNumbers "
    sql = sql & "From Patient "
    sql = sql & "JOIN Visitation ON Patient.PatientID = Visitation.PatientID "
    sql = sql & "WHERE SpecialistGroupID = '" & specialistGroup & "'"

    With rst
        .open sql, conn, 3, 4
        
        If .fields("PhoneNumbers") <> "NULL" Then
            
            response.write "<div class='container'>"
             response.write .fields("PhoneNumbers")
            response.write "</div>"

        Else
            response.write "<h1>No records found</h1>"
        End If
    End With
    rst.Close
    Set rst = Nothing
End Sub


Sub Styles()
    response.write "<style>"
        response.write ".container, h1 {"
        response.write "    width: 65vw;"
        response.write "    font-family: Arial, sans-serif;"
        response.write "}"
        
response.write "</style>"

End Sub



'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
