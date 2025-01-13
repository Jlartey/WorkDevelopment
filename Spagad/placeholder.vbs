 If (periodStart <> "" And periodEnd <> "") Then
    sql = sql & " BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    Else
        sql = sql & " BETWEEN '2018-01-10'AND '2022-12-31'"
    End If