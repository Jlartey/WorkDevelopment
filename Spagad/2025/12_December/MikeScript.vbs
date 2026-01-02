'Call IMAHMSUtils.GenerateSMSTransactions("", "")
'Call mNotifySendAll
'
'Sub mNotifySendAll()
'    Dim sql, rst
'
'    Set rst = CreateObject("ADODB.RecordSet")
'
'    sql = "select SMSTransactionID from SMSTransaction where SMSSentStatusID in ('S001', 'S004') "
'    sql = sql & " and cast(RequestDate as date)=cast(getdate() as date)"
'    rst.open qryPro.FltQry(sql), conn, 3, 4
'    If rst.recordCount > 0 Then
'        rst.MoveFirst
'        Do While Not rst.EOF
'            IMAHMSUtils.mNotifySend (rst.fields("SMSTransactionID"))
'            rst.MoveNext
'        Loop
'    End If
'    rst.Close
'    Set rst = Nothing
'
'End Sub
'Sub UpdateTableValues()
'    Dim sql
'
'    sql = "alter table SMSTransaction alter column Description nvarchar(3000)"
'    conn.execute qryPro.FltQry(sql)
'
'    sql = "update Appointment set Appointment.KeyPrefix=Patient.ResidencePhone"
'    sql = sql & " from Appointment inner join Patient on Patient.PatientID=Appointment.PatientID"
'    sql = sql & "   and len(Appointment.KeyPrefix)=0 and len(Patient.ResidencePhone)>0"
'    sql = sql & "   and Patient.PatientID not in ('P1', 'P2', 'P3', 'P4') "
'    conn.execute qryPro.FltQry(sql)
'
'    sql = "update LabRequest set LabRequest.ContactNo=Patient.ResidencePhone"
'    sql = sql & " from LabRequest inner join Patient on Patient.PatientID=LabRequest.PatientID"
'    sql = sql & "   and len(LabRequest.ContactNo)=0 and len(Patient.ResidencePhone)>0"
'    sql = sql & "   and Patient.PatientID not in ('P1', 'P2', 'P3', 'P4') "
'    conn.execute qryPro.FltQry(sql)
'End Sub