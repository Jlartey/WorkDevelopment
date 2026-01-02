Dim vst
vst = Trim(request("inpVisitationID"))

If checkDoubleAdmission(vst) Then
  If objPage.rtnHdlProcessPoint Then
     objPage.hdlProcessPoint = False
  End If
End If


Function checkDoubleAdmission(vst)
Dim rst, sql, vld
Set rst = CreateObject("ADODB.RecordSet")
vld = False
sql = "SELECT * FROM Admission where AdmissionStatusId = 'A001' and visitationID = '" & vst & "'"
With rst
    .open sql, conn, 3, 4
        If .RecordCount > 0 Then
            vld = True
            SetPageMessages "This patient is already on admission"
        End If
    .Close
End With

checkDoubleAdmission = vld
End Function
