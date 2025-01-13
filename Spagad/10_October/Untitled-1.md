wpgPrtPrintLayoutAll.asp?PositionForTableName=WorkingDay&WorkingDayID=DAY20160401&PrintLayoutName=EMRSelector&CompTableKeyID=EMRComponentID&VisitationID=" &vst & "&EMRDataID=RES018&InvestDays=&ModuleManagerName=Research&PositionForCtxTableName=Visitation&SectionType=EMR&WorkFlowNav=POP

Do While Not .EOF
cnt = cnt + 1
PatientID = .fields("PatientID")
EMRRequestID = .fields("EMRRequestID")
vst = .fields("visitationID")
hrf = "wpgPrtPrintLayoutAll.asp?PositionForTableName=WorkingDay&WorkingDayID=DAY20160401&PrintLayoutName=EMRSelector&CompTableKeyID=EMRComponentID&VisitationID=" & vst & "&EMRDataID=RES018&InvestDays=&ModuleManagerName=Research&PositionForCtxTableName=Visitation&SectionType=EMR&WorkFlowNav=POP"
response.write " <tr>"
response.write " <a href='" & hrf & "' target='_blank'> "
response.write " <td>" & cnt & "</td>"
response.write " <td>" & PatientID & "</td>"
response.write " <td>" & GetComboName("Patient", PatientID) & "</td>"
response.write " <td>" & getSRS_Score(EMRRequestID) & "</td>"
response.write " </a>"
response.write " </tr>"
.MoveNext
Loop
