VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PRTLYO_MedicalExamHBReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
response.Clear
response.write Glob_GetBootstrap5()


Dim dpt, secid
 dpt = Trim(Request.QueryString("dptID"))
'  secid = Trim(request.QueryString("secID"))
  genID = Trim(Request.QueryString("genID"))
  branid = Trim(Request.QueryString("branID"))
  yrid = Trim(Request.QueryString("yrID"))
  
  currentYr = "YRS" & Year(Date)

   If Len(yrid) < 1 Then
   yrid = currentYr
   Else
   yrid = yrid
   End If
'   If Len(branid) < 1 Then
'   branid = "B002"
'   Else
'   branid = branid
'   End If
'   If Len(genid) < 1 Then
'   genid = "GEN02"
'   Else
'   genid = genid
'   End If
'   If Len(secid) < 1 Then
'   secid = "L0460"
'   Else
'   secid = secid
'   End If
   

'   MedicalExaminationLabResultsAllFBC dpt
'   MedicalExaminationLabResultsAllLFT dpt
'MedicalExaminationLabResultsAllBUE dpt
'MedicalExaminationLabResultsAllLIPID dpt
'MedicalExaminationLabResultsAllESR dpt
'MedicalExaminationLabResultsAllPSA dpt

MedicalExaminationLabResultsAllHB dpt
'MedicalExaminationLabResultsAllMPS dpt
'MedicalExaminationLabResultsAllUR dpt
   
'   MedicalExaminationLabResultsOne dpt, secid
 
   
   
  
Sub MedicalExaminationLabResultsAllFBC(testId)

Dim rst, sql

Set rst = CreateObject("ADODB.Recordset")


  sql = ""
sql = sql & " WITH LabReportRange AS ( "
sql = sql & " SELECT i.LabRequestID, i.LabTestID, i.RequestDate, i.VisitationID, i.PatientID, i.GenderID, i.BranchID, [L0460].Column2 AS WBC, [L0468].Column2 AS RBC, [L0467].Column2 AS HB, [L0469].Column2 AS HCT, [L0470].Column2 AS MCV, [L0471].Column2 AS MCH, [L0472].Column2 AS MCHC, [L0475].Column2 AS PLT, [L0200].Column2 AS RDWSD, [L0201].Column2 AS RDWCV, [L0202].Column2 AS PDW, [L0476].Column2 AS MPV, [L0204].Column2 AS PLCR, [L0323].Column2 AS PCT, [L0199].Column2 AS NEUTGRAN, [L0196].Column2 AS [NEUTGRANN], [L0461].Column2 AS [LYM], [L0464].Column2 AS [LYMPH], [L0317].Column2 AS [MONO], [L0318].Column2 AS [MONOO], [L0319].Column2 AS [EO], [L0320].Column2 AS [EOO], [L0321].Column2 AS [BASO], [L0322].Column2 AS [BASOO], [L0719].Column2 AS [MID], [L0720].Column2 AS [MIDD] FROM Investigation i "
sql = sql & " LEFT JOIN Visitation v ON i.VisitationID = v.VisitationID "
    sql = sql & " LEFT JOIN LabResults AS [L0460] ON i.LabRequestID = [L0460].LabRequestID "
        sql = sql & " AND [L0460].TestComponentID = 'L0460' "
    sql = sql & " LEFT JOIN LabResults AS [L0468] ON i.LabRequestID = [L0468].LabRequestID "
        sql = sql & " AND [L0468].TestComponentID = 'L0468' "
    sql = sql & " LEFT JOIN LabResults AS [L0467] ON i.LabRequestID = [L0467].LabRequestID "
        sql = sql & " AND [L0467].TestComponentID = 'L0467' "
    sql = sql & " LEFT JOIN LabResults AS [L0469] ON i.LabRequestID = [L0469].LabRequestID "
        sql = sql & " AND [L0469].TestComponentID = 'L0469' "
    sql = sql & " LEFT JOIN LabResults AS [L0470] ON i.LabRequestID = [L0470].LabRequestID "
        sql = sql & " AND [L0470].TestComponentID = 'L0470' "
    sql = sql & " LEFT JOIN LabResults AS [L0471] ON i.LabRequestID = [L0471].LabRequestID "
        sql = sql & " AND [L0471].TestComponentID = 'L0471' "
    sql = sql & " LEFT JOIN LabResults AS [L0472] ON i.LabRequestID = [L0472].LabRequestID "
        sql = sql & " AND [L0472].TestComponentID = 'L0472' "
   sql = sql & " LEFT JOIN LabResults AS [L0475] ON i.LabRequestID = [L0475].LabRequestID "
        sql = sql & " AND [L0475].TestComponentID = 'L0475' "
   sql = sql & " LEFT JOIN LabResults AS [L0200] ON i.LabRequestID = [L0200].LabRequestID "
        sql = sql & " AND [L0200].TestComponentID = 'L0200' "
   sql = sql & " LEFT JOIN LabResults AS [L0201] ON i.LabRequestID = [L0201].LabRequestID "
        sql = sql & " AND [L0201].TestComponentID = 'L0201' "
   sql = sql & " LEFT JOIN LabResults AS [L0202] ON i.LabRequestID = [L0202].LabRequestID "
        sql = sql & " AND [L0202].TestComponentID = 'L0202' "
   sql = sql & " LEFT JOIN LabResults AS [L0476] ON i.LabRequestID = [L0476].LabRequestID "
        sql = sql & " AND [L0476].TestComponentID = 'L0476' "
   sql = sql & " LEFT JOIN LabResults AS [L0204] ON i.LabRequestID = [L0204].LabRequestID "
        sql = sql & " AND [L0204].TestComponentID = 'L0204' "
   sql = sql & " LEFT JOIN LabResults AS [L0323] ON i.LabRequestID = [L0323].LabRequestID "
        sql = sql & " AND [L0323].TestComponentID = 'L0323' "
   sql = sql & " LEFT JOIN LabResults AS [L0199] ON i.LabRequestID = [L0199].LabRequestID "
        sql = sql & " AND [L0199].TestComponentID = 'L0199' "
   sql = sql & " LEFT JOIN LabResults AS [L0196] ON i.LabRequestID = [L0196].LabRequestID "
        sql = sql & " AND [L0196].TestComponentID = 'L0196' "
   sql = sql & " LEFT JOIN LabResults AS [L0461] ON i.LabRequestID = [L0461].LabRequestID "
        sql = sql & " AND [L0461].TestComponentID = 'L0461' "
   sql = sql & " LEFT JOIN LabResults AS [L0464] ON i.LabRequestID = [L0464].LabRequestID "
        sql = sql & " AND [L0464].TestComponentID = 'L0464' "
   sql = sql & " LEFT JOIN LabResults AS [L0317] ON i.LabRequestID = [L0317].LabRequestID "
        sql = sql & " AND [L0317].TestComponentID = 'L0317' "
   sql = sql & " LEFT JOIN LabResults AS [L0318] ON i.LabRequestID = [L0318].LabRequestID "
        sql = sql & " AND [L0318].TestComponentID = 'L0318' "
   sql = sql & " LEFT JOIN LabResults AS [L0319] ON i.LabRequestID = [L0319].LabRequestID "
        sql = sql & " AND [L0319].TestComponentID = 'L0319' "
   sql = sql & " LEFT JOIN LabResults AS [L0320] ON i.LabRequestID = [L0320].LabRequestID "
        sql = sql & " AND [L0320].TestComponentID = 'L0320' "
   sql = sql & " LEFT JOIN LabResults AS [L0321] ON i.LabRequestID = [L0321].LabRequestID "
        sql = sql & " AND [L0321].TestComponentID = 'L0321' "
   sql = sql & " LEFT JOIN LabResults AS [L0322] ON i.LabRequestID = [L0322].LabRequestID "
        sql = sql & " AND [L0322].TestComponentID = 'L0322' "
   sql = sql & " LEFT JOIN LabResults AS [L0719] ON i.LabRequestID = [L0719].LabRequestID "
        sql = sql & " AND [L0719].TestComponentID = 'L0719' "
   sql = sql & " LEFT JOIN LabResults AS [L0720] ON i.LabRequestID = [L0720].LabRequestID "
        sql = sql & " AND [L0720].TestComponentID = 'L0720' "
    sql = sql & " Where i.LabTestID = '" & dpt & "' "
        If Len(branid) > 0 Then
        sql = sql & " AND i.BranchID='" & branid & "' "
    End If
    If Len(genID) > 0 Then
        sql = sql & " AND i.GenderID='" & genID & "' "
    End If
        sql = sql & " AND i.SponsorID IN ('S571','S592','S593', 'S594','S657','VRA') "
        sql = sql & " AND v.SpecialistTypeID IN ('B001-0200','B002-0208','B003-0208') "
        If Len(yrid) > 0 Then
        sql = sql & "  AND i.WorkingYearID = '" & yrid & "' "
    End If
    sql = sql & " Union All "
    sql = sql & " SELECT i2.LabRequestID, i2.LabTestID, i2.RequestDate, i2.VisitationID, i2.PatientID, i2.GenderID, i2.BranchID, [L0460].Column2 AS WBC, [L0468].Column2 AS RBC, [L0467].Column2 AS HB, [L0469].Column2 AS HCT, [L0470].Column2 AS MCV, [L0471].Column2 AS MCH, [L0472].Column2 AS MCHC, [L0475].Column2 AS PLT, [L0200].Column2 AS RDWSD, [L0201].Column2 AS RDWCV, [L0202].Column2 AS PDW, [L0476].Column2 AS MPV, [L0204].Column2 AS PLCR, [L0323].Column2 AS PCT, [L0199].Column2 AS NEUTGRAN, [L0196].Column2 AS [NEUTGRANN], [L0461].Column2 AS [LYM], [L0464].Column2 AS [LYMPH], [L0317].Column2 AS [MONO], [L0318].Column2 AS [MONOO], [L0319].Column2 AS [EO], [L0320].Column2 AS [EOO], [L0321].Column2 AS [BASO], [L0322].Column2 AS [BASOO], [L0719].Column2 AS [MID], [L0720].Column2 AS [MIDD] FROM Investigation2 i2 "
    sql = sql & " LEFT JOIN Visitation v ON i2.VisitationID = v.VisitationID "
     sql = sql & " LEFT JOIN LabResults AS [L0460] ON i2.LabRequestID = [L0460].LabRequestID "
        sql = sql & " AND [L0460].TestComponentID = 'L0460' "
    sql = sql & " LEFT JOIN LabResults AS [L0468] ON i2.LabRequestID = [L0468].LabRequestID "
        sql = sql & " AND [L0468].TestComponentID = 'L0468' "
    sql = sql & " LEFT JOIN LabResults AS [L0467] ON i2.LabRequestID = [L0467].LabRequestID "
        sql = sql & " AND [L0467].TestComponentID = 'L0467' "
    sql = sql & " LEFT JOIN LabResults AS [L0469] ON i2.LabRequestID = [L0469].LabRequestID "
        sql = sql & " AND [L0469].TestComponentID = 'L0469' "
    sql = sql & " LEFT JOIN LabResults AS [L0470] ON i2.LabRequestID = [L0470].LabRequestID "
        sql = sql & " AND [L0470].TestComponentID = 'L0470' "
    sql = sql & " LEFT JOIN LabResults AS [L0471] ON i2.LabRequestID = [L0471].LabRequestID "
        sql = sql & " AND [L0471].TestComponentID = 'L0471' "
    sql = sql & " LEFT JOIN LabResults AS [L0472] ON i2.LabRequestID = [L0472].LabRequestID "
        sql = sql & " AND [L0472].TestComponentID = 'L0472' "
   sql = sql & " LEFT JOIN LabResults AS [L0475] ON i2.LabRequestID = [L0475].LabRequestID "
        sql = sql & " AND [L0475].TestComponentID = 'L0475' "
   sql = sql & " LEFT JOIN LabResults AS [L0200] ON i2.LabRequestID = [L0200].LabRequestID "
        sql = sql & " AND [L0200].TestComponentID = 'L0200' "
   sql = sql & " LEFT JOIN LabResults AS [L0201] ON i2.LabRequestID = [L0201].LabRequestID "
        sql = sql & " AND [L0201].TestComponentID = 'L0201' "
   sql = sql & " LEFT JOIN LabResults AS [L0202] ON i2.LabRequestID = [L0202].LabRequestID "
        sql = sql & " AND [L0202].TestComponentID = 'L0202' "
   sql = sql & " LEFT JOIN LabResults AS [L0476] ON i2.LabRequestID = [L0476].LabRequestID "
        sql = sql & " AND [L0476].TestComponentID = 'L0476' "
   sql = sql & " LEFT JOIN LabResults AS [L0204] ON i2.LabRequestID = [L0204].LabRequestID "
        sql = sql & " AND [L0204].TestComponentID = 'L0204' "
   sql = sql & " LEFT JOIN LabResults AS [L0323] ON i2.LabRequestID = [L0323].LabRequestID "
        sql = sql & " AND [L0323].TestComponentID = 'L0323' "
   sql = sql & " LEFT JOIN LabResults AS [L0199] ON i2.LabRequestID = [L0199].LabRequestID "
        sql = sql & " AND [L0199].TestComponentID = 'L0199' "
   sql = sql & " LEFT JOIN LabResults AS [L0196] ON i2.LabRequestID = [L0196].LabRequestID "
        sql = sql & " AND [L0196].TestComponentID = 'L0196' "
   sql = sql & " LEFT JOIN LabResults AS [L0461] ON i2.LabRequestID = [L0461].LabRequestID "
        sql = sql & " AND [L0461].TestComponentID = 'L0461' "
   sql = sql & " LEFT JOIN LabResults AS [L0464] ON i2.LabRequestID = [L0464].LabRequestID "
        sql = sql & " AND [L0464].TestComponentID = 'L0464' "
   sql = sql & " LEFT JOIN LabResults AS [L0317] ON i2.LabRequestID = [L0317].LabRequestID"
        sql = sql & " AND [L0317].TestComponentID = 'L0317' "
   sql = sql & " LEFT JOIN LabResults AS [L0318] ON i2.LabRequestID = [L0318].LabRequestID "
        sql = sql & " AND [L0318].TestComponentID = 'L0318' "
   sql = sql & " LEFT JOIN LabResults AS [L0319] ON i2.LabRequestID = [L0319].LabRequestID "
        sql = sql & " AND [L0319].TestComponentID = 'L0319' "
   sql = sql & " LEFT JOIN LabResults AS [L0320] ON i2.LabRequestID = [L0320].LabRequestID "
        sql = sql & " AND [L0320].TestComponentID = 'L0320' "
   sql = sql & " LEFT JOIN LabResults AS [L0321] ON i2.LabRequestID = [L0321].LabRequestID "
        sql = sql & " AND [L0321].TestComponentID = 'L0321' "
   sql = sql & " LEFT JOIN LabResults AS [L0322] ON i2.LabRequestID = [L0322].LabRequestID "
        sql = sql & " AND [L0322].TestComponentID = 'L0322' "
   sql = sql & " LEFT JOIN LabResults AS [L0719] ON i2.LabRequestID = [L0719].LabRequestID "
        sql = sql & " AND [L0719].TestComponentID = 'L0719' "
   sql = sql & " LEFT JOIN LabResults AS [L0720] ON i2.LabRequestID = [L0720].LabRequestID "
        sql = sql & " AND [L0720].TestComponentID = 'L0720' "
    sql = sql & " Where i2.LabTestID = '" & testId & "' "
    If Len(branid) > 0 Then
        sql = sql & " AND i2.BranchID='" & branid & "' "
    End If
    If Len(genID) > 0 Then
        sql = sql & " AND i2.GenderID='" & genID & "' "
    End If
        
        
        sql = sql & " AND i2.SponsorID IN ('S571','S592','S593', 'S594','S657','VRA') "
        sql = sql & " AND v.SpecialistTypeID IN ('B001-0200','B002-0208','B003-0208') "
          If Len(yrid) > 0 Then
           sql = sql & "  AND i2.WorkingYearID = '" & yrid & "' "
         End If
sql = sql & " ) "
sql = sql & " SELECT LabRequestID, LabTestID,BranchID, RequestDate,VisitationID,PatientID, GenderID,WBC,RBC,HB,HCT, MCV,MCH,MCHC,PLT,RDWSD, RDWCV,PDW, MPV,PLCR,PCT,NEUTGRAN, [NEUTGRANN],[LYM],[LYMPH],[MONO],[MONOO], [EO], [EOO], [BASO],[BASOO],[MID],[MIDD] FROM LabReportRange ORDER BY RequestDate  "

'response.write "<!DOCTYPE html>"
'response.write "<html lang=""en"">"
'
'response.write "<head> "
'response.write "    <title>AME " & GetComboName("WorkingYear", yrid) & " Lab Tests Records</title>"
'response.write "    <title>AME " & currentYr & " Employee Records</title>"

response.write "<link href=""https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600&display=swap"" rel=""stylesheet"">"
response.write "<link href=""https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css"" rel=""stylesheet"">"
response.write "<link href=""https://cdn.datatables.net/buttons/2.4.1/css/buttons.dataTables.min.css"" rel=""stylesheet"">"
response.write "<link rel=""stylesheet"" href=""https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/css/bootstrap.min.css"" integrity=""sha384-rbsA2VBKQhggwzxH7pPCaAqO46MgnOM80zW1RWuH61DGLwZJEdK2Kadq2F9CUG65"" crossorigin=""anonymous"">"
response.write "<link href=""https://cdn.datatables.net/v/dt/dt-1.13.6/datatables.min.css"" rel=""stylesheet"">"
response.write "<link href=""https://cdn.datatables.net/v/bs5/jszip-3.10.1/dt-1.13.6/b-2.4.1/b-html5-2.4.1/b-print-2.4.1/datatables.min.css"" rel=""stylesheet"">"

AddCss
'response.write "</head>"

BuildComboFilter
response.write "<a class=""btn btn-primary1"" href=""wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamStatRCP&PositionForTableName=WorkingDay&WorkingDayID="">Go Home</a>"

response.write "<a class=""btn btn-primary"" href=""wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamLabResultRCP&PositionForTableName=WorkingDay&dptID=" & dpt & "&secID=" & secid & "&genID=" & genID & "&branID=" & branid & "&yrID=" & yrid & " "">Refresh</a>"
response.write "        </td>"
response.write "</table>"
response.write "</div>"
response.write "                    <table id=""example"" class=""table table-striped table-bordered""  width=""100%"">"
response.write "                        <thead class=""beautiful-green-header"">"
response.write "<tr>"
response.write "<th>#</th>"
response.write "<th>LabRequestID</th>"
response.write "<th>LabTestID</th>"
response.write "<th>Facility</th>"
response.write "<th>RequestDate</th>"
response.write "<th>VisitationID</th>"
response.write "<th>PatientID</th>"
response.write "<th>GenderID</th>"
response.write "<th>WBC</th>"
response.write "<th>RBC</th>"
response.write "<th>HB</th>"
response.write "<th>HCT</th>"
response.write "<th>MCV</th>"
response.write "<th>MCH</th>"
response.write "<th>MCHC</th>"
response.write "<th>PLT</th>"
response.write "<th>RDW_SD</th>"
response.write "<th>RDW_CV</th>"
response.write "<th>PDW</th>"
response.write "<th>MPV</th>"
response.write "<th>P_LCR</th>"
response.write "<th>PCT</th>"
response.write "<th>[NEUT#_GRAN#]</th>"
response.write "<th>[NEUT%_GRAN%]</th>"
response.write "<th>[LYM#]</th>"
response.write "<th>[LYMPH%]</th>"
response.write "<th>[MONO#]</th>"
response.write "<th>[MONO%]</th>"
response.write "<th>[EO#]</th>"
response.write "<th>[EO%]</th>"
response.write "<th>[BASO#]</th>"
response.write "<th>[BASO%]</th>"
response.write "<th>[MID#]</th>"
response.write "<th>[MID%]</th>"
response.write "</tr>"
response.write "</thead>"
response.write "<tbody>"

count = 0

With rst
    rst.open qryPro.FltQry(sql), conn, 3, 4

    If rst.RecordCount > 0 Then
        rst.MoveFirst

        Do While Not rst.EOF

            count = count + 1
   
 
  
            response.write "<tr>"
            response.write "<td>" & count & "</td>"
            response.write "<td>" & rst.fields("LabRequestID") & "</td>"
            response.write "<td>" & GetComboName("LabTest", .fields("LabTestID")) & "</td>"
            response.write "<td>" & GetComboName("Branch", .fields("BranchID")) & "</td>"
            response.write "<td>" & FormatDate(rst.fields("RequestDate")) & "</td>"
            response.write "<td>" & rst.fields("visitationID") & "</td>"
            response.write "<td>" & rst.fields("patientID") & "</td>"
            response.write "<td>" & GetComboName("Gender", .fields("GenderID")) & "</td>"
            response.write "<td>" & rst.fields("WBC") & "</td>"
            response.write "<td>" & rst.fields("RBC") & "</td>"
            response.write "<td>" & rst.fields("HB") & "</td>"
            response.write "<td>" & rst.fields("HCT") & "</td>"
            response.write "<td>" & rst.fields("MCV") & "</td>"
            response.write "<td>" & rst.fields("MCH") & "</td>"
            response.write "<td>" & rst.fields("MCHC") & "</td>"
            response.write "<td>" & rst.fields("PLT") & "</td>"
            response.write "<td>" & rst.fields("RDWSD") & "</td>"
            response.write "<td>" & rst.fields("RDWCV") & "</td>"
            response.write "<td>" & rst.fields("PDW") & "</td>"
            response.write "<td>" & rst.fields("MPV") & "</td>"
            response.write "<td>" & rst.fields("PLCR") & "</td>"
            response.write "<td>" & rst.fields("PCT") & "</td>"
            response.write "<td>" & rst.fields("NEUTGRAN") & "</td>"
            response.write "<td>" & rst.fields("NEUTGRANN") & "</td>"
            response.write "<td>" & rst.fields("LYM") & "</td>"
            response.write "<td>" & rst.fields("LYMPH") & "</td>"
            response.write "<td>" & rst.fields("MONO") & "</td>"
            response.write "<td>" & rst.fields("MONOO") & "</td>"
            response.write "<td>" & rst.fields("EO") & "</td>"
            response.write "<td>" & rst.fields("EOO") & "</td>"
            response.write "<td>" & rst.fields("BASO") & "</td>"
            response.write "<td>" & rst.fields("BASOO") & "</td>"
            response.write "<td>" & rst.fields("MID") & "</td>"
            response.write "<td>" & rst.fields("MIDD") & "</td>"
            response.write "</tr>"
            rst.MoveNext
            response.flush
        Loop
    End If
    rst.Close
End With
Set rst = Nothing

response.write "                        </tbody>"
response.write "                    </table>"
response.write "                </div>"
response.write "            </div>"
response.write "        </div>"
response.write "    </div>"




AddJS

'response.write "</html>"



response.write "<script>"
response.write "function reload2() {"
response.write "    let selected = document.getElementById('cID').value;"
response.write "    let url = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamFBCReport';"
response.write "    url += '&PositionForTableName=WorkingDay&dptID=' + selected;"
response.write "    window.location.href = url;"
response.write "}"

response.write "</script>"
response.write "<script>"
response.write "function reload3() {"
response.write "    let selected = document.getElementById('cID').value;"
'response.write "    let selected2 = document.getElementById('cID2').value;"
response.write "    let selected3 = document.getElementById('cID3').value;"
response.write "    let selected4 = document.getElementById('cID4').value;"
response.write "    let selected5 = document.getElementById('cID5').value;"
'response.write "    let sat = document.getElementById('statusFilter').value;"
response.write "    let url = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamFBCReport';"
response.write "    url += '&PositionForTableName=WorkingDay&dptID=' + selected + '&genID=' + selected3 + '&branID=' + selected4 + '&yrID=' + selected5;"
response.write "    window.location.href = url;"
response.write "}"



response.write "</script>"

response.write "<script>"

response.write "        var dataTable = $('#example').DataTable({"
response.write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, ""All""]],"
response.write "        dom: 'lBfrtip',"
response.write "            pageLength: 10,"
response.write "            paging: true, "
response.write "            pagingType: 'simple', "
response.write "            search: {"
response.write "                smart: true"
response.write "                    },"
response.write "            buttons: ["
response.write "                {"
response.write "                    extend: 'csv',"
response.write "                    text: 'CSV',"
response.write "                    title: '" & dpt & " Labtest-" & FormatDate(Now) & "' "
response.write "                },"
response.write "                {"
response.write "                    extend: 'excel',"
response.write "                    text: 'EXCEL',"
response.write "                    title: '" & dpt & " Labtest-" & FormatDate(Now) & "'"
response.write "                },"
response.write "                {"
response.write "                    extend: 'pdf',"
response.write "                    text: 'PDF',"
response.write "                    title: '" & dpt & " Labtest-" & FormatDate(Now) & "'"
response.write "                },"

response.write "                {"
response.write "                    extend: 'print',"
response.write "                    text: 'PRINT',"
response.write "                    title: '" & dpt & " Labtest-" & FormatDate(Now) & "'"
response.write "                }"

response.write "            ],"

response.write "        columnDefs: ["
response.write "        {"
response.write "            target: 7,"
response.write "            visible: true,"
response.write "            searchable: true"
response.write "        }"
response.write "    ]"

response.write "        });"
response.write "</script>"




End Sub
Sub MedicalExaminationLabResultsAllLFT(testId)

Dim rst, sql

Set rst = CreateObject("ADODB.Recordset")


  sql = ""
sql = sql & " WITH LabReportRange AS ( "
sql = sql & " SELECT i.LabRequestID, i.LabTestID, i.RequestDate, i.VisitationID, i.PatientID, i.GenderID, i.BranchID, [L0164].Column2 AS Total_Bilirubin, [L0152].Column2 AS Direct_Bilirubin, [L0398].Column2 AS Indirect_Bilirubin, [L0012].Column2 AS SGOT_AST, [L0006].Column2 AS SGPT_ALT, [L0005].Column2 AS Alkaline_Phosphatase, [L0055].Column2 AS GGT, [L0153].Column2 AS Total_Protein, [L0154].Column2 AS Albumin, [L0399].Column2 AS Globulin  FROM Investigation i "
sql = sql & " LEFT JOIN Visitation v ON i.VisitationID = v.VisitationID "
    sql = sql & " LEFT JOIN LabResults AS [L0164] ON i.LabRequestID = [L0164].LabRequestID "
        sql = sql & " AND [L0164].TestComponentID = 'L0164' "
    sql = sql & " LEFT JOIN LabResults AS [L0152] ON i.LabRequestID = [L0152].LabRequestID "
        sql = sql & " AND [L0152].TestComponentID = 'L0152' "
    sql = sql & " LEFT JOIN LabResults AS [L0398] ON i.LabRequestID = [L0398].LabRequestID "
        sql = sql & " AND [L0398].TestComponentID = 'L0398' "
    sql = sql & " LEFT JOIN LabResults AS [L0012] ON i.LabRequestID = [L0012].LabRequestID "
        sql = sql & " AND [L0012].TestComponentID = 'L0012' "
    sql = sql & " LEFT JOIN LabResults AS [L0006] ON i.LabRequestID = [L0006].LabRequestID "
        sql = sql & " AND [L0006].TestComponentID = 'L0006' "
    sql = sql & " LEFT JOIN LabResults AS [L0005] ON i.LabRequestID = [L0005].LabRequestID "
        sql = sql & " AND [L0005].TestComponentID = 'L0005' "
    sql = sql & " LEFT JOIN LabResults AS [L0055] ON i.LabRequestID = [L0055].LabRequestID "
        sql = sql & " AND [L0055].TestComponentID = 'L0055' "
   sql = sql & " LEFT JOIN LabResults AS [L0153] ON i.LabRequestID = [L0153].LabRequestID "
        sql = sql & " AND [L0153].TestComponentID = 'L0153' "
   sql = sql & " LEFT JOIN LabResults AS [L0154] ON i.LabRequestID = [L0154].LabRequestID "
        sql = sql & " AND [L0154].TestComponentID = 'L0154' "
   sql = sql & " LEFT JOIN LabResults AS [L0399] ON i.LabRequestID = [L0399].LabRequestID "
        sql = sql & " AND [L0399].TestComponentID = 'L0399' "
    sql = sql & " Where i.LabTestID = '" & testId & "' "
        If Len(branid) > 0 Then
        sql = sql & " AND i.BranchID= '" & branid & "' "
    End If
    If Len(genID) > 0 Then
        sql = sql & " AND i.GenderID= '" & genID & "' "
    End If
        sql = sql & " AND i.SponsorID IN ('S571','S592','S593', 'S594','S657','VRA') "
        sql = sql & " AND v.SpecialistTypeID IN ('B001-0200','B002-0208','B003-0208') "
        If Len(yrid) > 0 Then
        sql = sql & "  AND i.WorkingYearID = '" & yrid & "' "
    End If
    sql = sql & " Union All "
    sql = sql & " SELECT i2.LabRequestID, i2.LabTestID, i2.RequestDate, i2.VisitationID, i2.PatientID, i2.GenderID, i2.BranchID, [L0164].Column2 AS Total_Bilirubin, [L0152].Column2 AS Direct_Bilirubin, [L0398].Column2 AS Indirect_Bilirubin, [L0012].Column2 AS SGOT_AST, [L0006].Column2 AS SGPT_ALT, [L0005].Column2 AS Alkaline_Phosphatase, [L0055].Column2 AS GGT, [L0153].Column2 AS Total_Protein, [L0154].Column2 AS Albumin, [L0399].Column2 AS Globulin FROM Investigation2 i2 "
    sql = sql & " LEFT JOIN Visitation v ON i2.VisitationID = v.VisitationID "
    sql = sql & " LEFT JOIN LabResults AS [L0164] ON i2.LabRequestID = [L0164].LabRequestID "
        sql = sql & " AND [L0164].TestComponentID = 'L0164' "
    sql = sql & " LEFT JOIN LabResults AS [L0152] ON i2.LabRequestID = [L0152].LabRequestID "
        sql = sql & " AND [L0152].TestComponentID = 'L0152' "
    sql = sql & " LEFT JOIN LabResults AS [L0398] ON i2.LabRequestID = [L0398].LabRequestID "
        sql = sql & " AND [L0398].TestComponentID = 'L0398' "
    sql = sql & " LEFT JOIN LabResults AS [L0012] ON i2.LabRequestID = [L0012].LabRequestID "
        sql = sql & " AND [L0012].TestComponentID = 'L0012' "
    sql = sql & " LEFT JOIN LabResults AS [L0006] ON i2.LabRequestID = [L0006].LabRequestID "
        sql = sql & " AND [L0006].TestComponentID = 'L0006' "
    sql = sql & " LEFT JOIN LabResults AS [L0005] ON i2.LabRequestID = [L0005].LabRequestID "
        sql = sql & " AND [L0005].TestComponentID = 'L0005' "
    sql = sql & " LEFT JOIN LabResults AS [L0055] ON i2.LabRequestID = [L0055].LabRequestID "
        sql = sql & " AND [L0055].TestComponentID = 'L0055' "
   sql = sql & " LEFT JOIN LabResults AS [L0153] ON i2.LabRequestID = [L0153].LabRequestID "
        sql = sql & " AND [L0153].TestComponentID = 'L0153' "
   sql = sql & " LEFT JOIN LabResults AS [L0154] ON i2.LabRequestID = [L0154].LabRequestID "
        sql = sql & " AND [L0154].TestComponentID = 'L0154' "
   sql = sql & " LEFT JOIN LabResults AS [L0399] ON i2.LabRequestID = [L0399].LabRequestID "
        sql = sql & " AND [L0399].TestComponentID = 'L0399' "
    sql = sql & " Where i2.LabTestID = '" & testId & "' "
    If Len(branid) > 0 Then
        sql = sql & " AND i2.BranchID='" & branid & "' "
    End If
    If Len(genID) > 0 Then
        sql = sql & " AND i2.GenderID='" & genID & "' "
    End If
        
        
        sql = sql & " AND i2.SponsorID IN ('S571','S592','S593', 'S594','S657','VRA') "
        sql = sql & " AND v.SpecialistTypeID IN ('B001-0200','B002-0208','B003-0208') "
          If Len(yrid) > 0 Then
           sql = sql & "  AND i2.WorkingYearID = '" & yrid & "' "
         End If
sql = sql & " ) "
sql = sql & " SELECT LabRequestID, LabTestID,BranchID, RequestDate,VisitationID,PatientID, GenderID,Total_Bilirubin,Direct_Bilirubin,Indirect_Bilirubin,SGOT_AST, SGPT_ALT,Alkaline_Phosphatase,GGT,Total_Protein,Albumin, Globulin FROM LabReportRange ORDER BY RequestDate  "

response.write "<!DOCTYPE html>"
response.write "<html lang=""en"">"

response.write "<head> "
response.write "    <title>AME " & GetComboName("WorkingYear", yrid) & " Lab Tests Records</title>"
'response.write "    <title>AME " & currentYr & " Employee Records</title>"

response.write "<link href=""https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600&display=swap"" rel=""stylesheet"">"
response.write "<link href=""https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css"" rel=""stylesheet"">"
response.write "<link href=""https://cdn.datatables.net/buttons/2.4.1/css/buttons.dataTables.min.css"" rel=""stylesheet"">"
response.write "<link rel=""stylesheet"" href=""https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/css/bootstrap.min.css"" integrity=""sha384-rbsA2VBKQhggwzxH7pPCaAqO46MgnOM80zW1RWuH61DGLwZJEdK2Kadq2F9CUG65"" crossorigin=""anonymous"">"
response.write "<link href=""https://cdn.datatables.net/v/dt/dt-1.13.6/datatables.min.css"" rel=""stylesheet"">"
response.write "<link href=""https://cdn.datatables.net/v/bs5/jszip-3.10.1/dt-1.13.6/b-2.4.1/b-html5-2.4.1/b-print-2.4.1/datatables.min.css"" rel=""stylesheet"">"

AddCss
response.write "</head>"

BuildComboFilter


response.write "<a class=""btn btn-primary1"" href=""wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamStatRCP&PositionForTableName=WorkingDay&WorkingDayID="">Go Home</a>"

response.write "<a class=""btn btn-primary"" href=""wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamLabResultRCP&PositionForTableName=WorkingDay&dptID=" & dpt & "&secID=" & secid & "&genID=" & genID & "&branID=" & branid & "&yrID=" & yrid & " "">Refresh</a>"
response.write "        </td>"
response.write "</table>"
response.write "</div>"
response.write "                    <table id=""example"" class=""table table-striped table-bordered""  width=""100%"">"
response.write "                        <thead class=""beautiful-green-header"">"
response.write "<tr>"
response.write "<th>#</th>"
response.write "<th>LabRequestID</th>"
response.write "<th>LabTestID</th>"
response.write "<th>Facility</th>"
response.write "<th>RequestDate</th>"
response.write "<th>VisitationID</th>"
response.write "<th>PatientID</th>"
response.write "<th>GenderID</th>"
response.write "<th>Total_Bilirubin</th>"
response.write "<th>Direct_Bilirubin</th>"
response.write "<th>Indirect_Bilirubin</th>"
response.write "<th>SGOT_AST</th>"
response.write "<th>SGPT_ALT</th>"
response.write "<th>Alkaline_Phosphatase</th>"
response.write "<th>GGT</th>"
response.write "<th>Total_Protein</th>"
response.write "<th>Albumin</th>"
response.write "<th>Globulin</th>"
response.write "</tr>"
response.write "</thead>"
response.write "<tbody>"

count = 0

With rst
    rst.open qryPro.FltQry(sql), conn, 3, 4

    If rst.RecordCount > 0 Then
        rst.MoveFirst

        Do While Not rst.EOF

            count = count + 1
   
 
  
            response.write "<tr>"
            response.write "<td>" & count & "</td>"
            response.write "<td>" & rst.fields("LabRequestID") & "</td>"
            response.write "<td>" & GetComboName("LabTest", .fields("LabTestID")) & "</td>"
            response.write "<td>" & GetComboName("Branch", .fields("BranchID")) & "</td>"
            response.write "<td>" & FormatDate(rst.fields("RequestDate")) & "</td>"
            response.write "<td>" & rst.fields("visitationID") & "</td>"
            response.write "<td>" & rst.fields("patientID") & "</td>"
            response.write "<td>" & GetComboName("Gender", .fields("GenderID")) & "</td>"
            response.write "<td>" & rst.fields("Total_Bilirubin") & "</td>"
            response.write "<td>" & rst.fields("Direct_Bilirubin") & "</td>"
            response.write "<td>" & rst.fields("Indirect_Bilirubin") & "</td>"
            response.write "<td>" & rst.fields("SGOT_AST") & "</td>"
            response.write "<td>" & rst.fields("SGPT_ALT") & "</td>"
            response.write "<td>" & rst.fields("Alkaline_Phosphatase") & "</td>"
            response.write "<td>" & rst.fields("GGT") & "</td>"
            response.write "<td>" & rst.fields("Total_Protein") & "</td>"
            response.write "<td>" & rst.fields("Albumin") & "</td>"
            response.write "<td>" & rst.fields("Globulin") & "</td>"
            response.write "</tr>"
            rst.MoveNext
            response.flush
        Loop
    End If
    rst.Close
End With
Set rst = Nothing

response.write "                        </tbody>"
response.write "                    </table>"
response.write "                </div>"
response.write "            </div>"
response.write "        </div>"
response.write "    </div>"




AddJS

response.write "</html>"



response.write "<script>"
response.write "function reload2() {"
response.write "    let selected = document.getElementById('cID').value;"
response.write "    let url = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamLFTReport';"
response.write "    url += '&PositionForTableName=WorkingDay&dptID=' + selected;"
response.write "    window.location.href = url;"
response.write "}"

response.write "</script>"
response.write "<script>"
response.write "function reload3() {"
response.write "    let selected = document.getElementById('cID').value;"
'response.write "    let selected2 = document.getElementById('cID2').value;"
response.write "    let selected3 = document.getElementById('cID3').value;"
response.write "    let selected4 = document.getElementById('cID4').value;"
response.write "    let selected5 = document.getElementById('cID5').value;"
'response.write "    let sat = document.getElementById('statusFilter').value;"
response.write "    let url = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamLFTReport';"
response.write "    url += '&PositionForTableName=WorkingDay&dptID=' + selected +  '&genID=' + selected3 + '&branID=' + selected4 + '&yrID=' + selected5;"
response.write "    window.location.href = url;"
response.write "}"


    response.write " function openPopup1(linkUrl, width, height) {"
    response.write " var left = (window.innerWidth - width) / 2;"
    response.write " var top = (window.innerHeight - height) / 2;"
    response.write " window.open(linkUrl, 'PopupWindow', 'width=' + width + ', height=' + height + ', left=' + left + ', top=' + top);"
    response.write " }"
response.write "</script>"

response.write "<script>"

response.write "        var dataTable = $('#example').DataTable({"
response.write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, ""All""]],"
response.write "        dom: 'lBfrtip',"
response.write "            pageLength: 10,"
response.write "            paging: true, "
response.write "            pagingType: 'simple', "
response.write "            search: {"
response.write "                smart: true"
response.write "                    },"
response.write "            buttons: ["
response.write "                {"
response.write "                    extend: 'csv',"
response.write "                    text: 'CSV',"
response.write "                    title: '" & dpt & " Labtest-" & FormatDate(Now) & "' "
response.write "                },"
response.write "                {"
response.write "                    extend: 'excel',"
response.write "                    text: 'EXCEL',"
response.write "                    title: '" & dpt & " Labtest-" & FormatDate(Now) & "'"
response.write "                },"
response.write "                {"
response.write "                    extend: 'pdf',"
response.write "                    text: 'PDF',"
response.write "                    title: '" & dpt & " Labtest-" & FormatDate(Now) & "'"
response.write "                },"

response.write "                {"
response.write "                    extend: 'print',"
response.write "                    text: 'PRINT',"
response.write "                    title: '" & dpt & " Labtest-" & FormatDate(Now) & "'"
response.write "                }"

response.write "            ],"

response.write "        columnDefs: ["
response.write "        {"
response.write "            target: 7,"
response.write "            visible: true,"
response.write "            searchable: true"
response.write "        }"
response.write "    ]"

response.write "        });"
response.write "</script>"





End Sub
Sub MedicalExaminationLabResultsAllBUE(testId)

Dim rst, sql

Set rst = CreateObject("ADODB.Recordset")


  sql = ""
sql = sql & " WITH LabReportRange AS ( "
sql = sql & " SELECT i.LabRequestID, i.LabTestID, i.RequestDate, i.VisitationID, i.PatientID, i.GenderID, i.SponsorID, [L0097].Column2 AS Potassium, [L0112].Column2 AS Sodium, [L0543].Column2 AS Chloride_Ion, [L0159].Column2 AS Urea, [L0036].Column2 AS Creatinine, [L0340].Column2 AS eGFR FROM Investigation i "
sql = sql & " LEFT JOIN Visitation v ON i.VisitationID = v.VisitationID "
    sql = sql & " LEFT JOIN LabResults AS [L0097] ON i.LabRequestID = [L0097].LabRequestID "
        sql = sql & " AND [L0097].TestComponentID = 'L0097' "
    sql = sql & " LEFT JOIN LabResults AS [L0112] ON i.LabRequestID = [L0112].LabRequestID "
        sql = sql & " AND [L0112].TestComponentID = 'L0112' "
    sql = sql & " LEFT JOIN LabResults AS [L0543] ON i.LabRequestID = [L0543].LabRequestID "
        sql = sql & " AND [L0543].TestComponentID = 'L0543' "
    sql = sql & " LEFT JOIN LabResults AS [L0159] ON i.LabRequestID = [L0159].LabRequestID "
        sql = sql & " AND [L0159].TestComponentID = 'L0159' "
    sql = sql & " LEFT JOIN LabResults AS [L0036] ON i.LabRequestID = [L0036].LabRequestID "
        sql = sql & " AND [L0036].TestComponentID = 'L0036' "
    sql = sql & " LEFT JOIN LabResults AS [L0340] ON i.LabRequestID = [L0340].LabRequestID "
        sql = sql & " AND [L0340].TestComponentID = 'L0340' "
    sql = sql & " Where i.LabTestID = '" & testId & "' "
        If Len(branid) > 0 Then
        sql = sql & " AND i.BranchID='" & branid & "' "
    End If
    If Len(genID) > 0 Then
        sql = sql & " AND i.GenderID='" & genID & "' "
    End If
        sql = sql & " AND i.SponsorID IN ('S571','S592','S593', 'S594','S657','VRA') "
        sql = sql & " AND v.SpecialistTypeID IN ('B001-0200','B002-0208','B003-0208') "
        If Len(yrid) > 0 Then
        sql = sql & "  AND i.WorkingYearID = '" & yrid & "' "
    End If
    sql = sql & " Union All "
    sql = sql & " SELECT i2.LabRequestID, i2.LabTestID, i2.RequestDate, i2.VisitationID, i2.PatientID, i2.GenderID, i2.SponsorID, [L0097].Column2 AS Potassium, [L0112].Column2 AS Sodium, [L0543].Column2 AS Chloride_Ion, [L0159].Column2 AS Urea, [L0036].Column2 AS Creatinine, [L0340].Column2 AS eGFR FROM Investigation2 i2 "
    sql = sql & " LEFT JOIN Visitation v ON i2.VisitationID = v.VisitationID "
    sql = sql & " LEFT JOIN LabResults AS [L0097] ON i2.LabRequestID = [L0097].LabRequestID "
        sql = sql & " AND [L0097].TestComponentID = 'L0097' "
    sql = sql & " LEFT JOIN LabResults AS [L0112] ON i2.LabRequestID = [L0112].LabRequestID "
        sql = sql & " AND [L0112].TestComponentID = 'L0112' "
    sql = sql & " LEFT JOIN LabResults AS [L0543] ON i2.LabRequestID = [L0543].LabRequestID "
        sql = sql & " AND [L0543].TestComponentID = 'L0543' "
    sql = sql & " LEFT JOIN LabResults AS [L0159] ON i2.LabRequestID = [L0159].LabRequestID "
        sql = sql & " AND [L0159].TestComponentID = 'L0159' "
    sql = sql & " LEFT JOIN LabResults AS [L0036] ON i2.LabRequestID = [L0036].LabRequestID "
        sql = sql & " AND [L0036].TestComponentID = 'L0036' "
    sql = sql & " LEFT JOIN LabResults AS [L0340] ON i2.LabRequestID = [L0340].LabRequestID "
        sql = sql & " AND [L0340].TestComponentID = 'L0340' "
    sql = sql & " Where i2.LabTestID = '" & dpt & "' "
    If Len(branid) > 0 Then
        sql = sql & " AND i2.BranchID='" & branid & "' "
    End If
    If Len(genID) > 0 Then
        sql = sql & " AND i2.GenderID='" & genID & "' "
    End If
        
        
        sql = sql & " AND i2.SponsorID IN ('S571','S592','S593', 'S594','S657','VRA') "
        sql = sql & " AND v.SpecialistTypeID IN ('B001-0200','B002-0208','B003-0208') "
          If Len(yrid) > 0 Then
           sql = sql & "  AND i2.WorkingYearID = '" & yrid & "' "
         End If
sql = sql & " ) "
sql = sql & " SELECT LabRequestID, LabTestID,SponsorID, RequestDate,VisitationID,PatientID, GenderID,Potassium,Sodium, Chloride_Ion, Urea, Creatinine, eGFR FROM LabReportRange ORDER BY RequestDate  "

response.write "<!DOCTYPE html>"
response.write "<html lang=""en"">"

response.write "<head> "
response.write "    <title>AME " & GetComboName("WorkingYear", yrid) & " Lab Tests Records</title>"
'response.write "    <title>AME " & currentYr & " Employee Records</title>"

response.write "<link href=""https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600&display=swap"" rel=""stylesheet"">"
response.write "<link href=""https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css"" rel=""stylesheet"">"
response.write "<link href=""https://cdn.datatables.net/buttons/2.4.1/css/buttons.dataTables.min.css"" rel=""stylesheet"">"
response.write "<link rel=""stylesheet"" href=""https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/css/bootstrap.min.css"" integrity=""sha384-rbsA2VBKQhggwzxH7pPCaAqO46MgnOM80zW1RWuH61DGLwZJEdK2Kadq2F9CUG65"" crossorigin=""anonymous"">"
response.write "<link href=""https://cdn.datatables.net/v/dt/dt-1.13.6/datatables.min.css"" rel=""stylesheet"">"
response.write "<link href=""https://cdn.datatables.net/v/bs5/jszip-3.10.1/dt-1.13.6/b-2.4.1/b-html5-2.4.1/b-print-2.4.1/datatables.min.css"" rel=""stylesheet"">"

AddCss
response.write "</head>"

BuildComboFilter
response.write "<a class=""btn btn-primary1"" href=""wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamStatRCP&PositionForTableName=WorkingDay&WorkingDayID="">Go Home</a>"

response.write "<a class=""btn btn-primary"" href=""wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamLabResultRCP&PositionForTableName=WorkingDay&dptID=" & dpt & "&secID=" & secid & "&genID=" & genID & "&branID=" & branid & "&yrID=" & yrid & " "">Refresh</a>"
response.write "        </td>"
response.write "</table>"
response.write "</div>"
response.write "                    <table id=""example"" class=""table table-striped table-bordered""  width=""100%"">"
response.write "                        <thead class=""beautiful-green-header"">"
response.write "<tr>"
response.write "<th>#</th>"
response.write "<th>LabRequestID</th>"
response.write "<th>LabTestID</th>"
response.write "<th>SponsorID</th>"
response.write "<th>RequestDate</th>"
response.write "<th>VisitationID</th>"
response.write "<th>PatientID</th>"
response.write "<th>GenderID</th>"
response.write "<th>Potassium</th>"
response.write "<th>Sodium</th>"
response.write "<th>Chloride_Ion</th>"
response.write "<th>Urea</th>"
response.write "<th>Creatinine</th>"
response.write "<th>eGFR</th>"
response.write "</tr>"
response.write "</thead>"
response.write "<tbody>"

count = 0

With rst
    rst.open qryPro.FltQry(sql), conn, 3, 4

    If rst.RecordCount > 0 Then
        rst.MoveFirst

        Do While Not rst.EOF

            count = count + 1
   
 
  
            response.write "<tr>"
            response.write "<td>" & count & "</td>"
            response.write "<td>" & rst.fields("LabRequestID") & "</td>"
            response.write "<td>" & GetComboName("LabTest", .fields("LabTestID")) & "</td>"
            response.write "<td>" & GetComboName("Sponsor", .fields("SponsorID")) & "</td>"
            response.write "<td>" & FormatDate(rst.fields("RequestDate")) & "</td>"
            response.write "<td>" & rst.fields("visitationID") & "</td>"
            response.write "<td>" & rst.fields("patientID") & "</td>"
            response.write "<td>" & GetComboName("Gender", .fields("GenderID")) & "</td>"
            response.write "<td>" & rst.fields("Potassium") & "</td>"
            response.write "<td>" & rst.fields("Sodium") & "</td>"
            response.write "<td>" & rst.fields("Chloride_Ion") & "</td>"
            response.write "<td>" & rst.fields("Urea") & "</td>"
            response.write "<td>" & rst.fields("Creatinine") & "</td>"
            response.write "<td>" & rst.fields("eGFR") & "</td>"
           
            response.write "</tr>"
            rst.MoveNext
            response.flush
        Loop
    End If
    rst.Close
End With
Set rst = Nothing

response.write "                        </tbody>"
response.write "                    </table>"
response.write "                </div>"
response.write "            </div>"
response.write "        </div>"
response.write "    </div>"




AddJS

response.write "</html>"



response.write "<script>"
response.write "function reload2() {"
response.write "    let selected = document.getElementById('cID').value;"
response.write "    let url = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamBUEReport';"
response.write "    url += '&PositionForTableName=WorkingDay&dptID=' + selected;"
response.write "    window.location.href = url;"
response.write "}"

response.write "</script>"
response.write "<script>"
response.write "function reload3() {"
response.write "    let selected = document.getElementById('cID').value;"
'response.write "    let selected2 = document.getElementById('cID2').value;"
response.write "    let selected3 = document.getElementById('cID3').value;"
response.write "    let selected4 = document.getElementById('cID4').value;"
response.write "    let selected5 = document.getElementById('cID5').value;"
'response.write "    let sat = document.getElementById('statusFilter').value;"
response.write "    let url = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamBUEReport';"
response.write "    url += '&PositionForTableName=WorkingDay&dptID=' + selected + '&genID=' + selected3 + '&branID=' + selected4 + '&yrID=' + selected5;"
response.write "    window.location.href = url;"
response.write "}"



response.write "</script>"
response.write "<script>"

response.write "        var dataTable = $('#example').DataTable({"
response.write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, ""All""]],"
response.write "        dom: 'lBfrtip',"
response.write "            pageLength: 10,"
response.write "            paging: true, "
response.write "            pagingType: 'simple', "
response.write "            search: {"
response.write "                smart: true"
response.write "                    },"
response.write "            buttons: ["
response.write "                {"
response.write "                    extend: 'csv',"
response.write "                    text: 'CSV',"
response.write "                    title: '" & dpt & " Labtest-" & FormatDate(Now) & "' "
response.write "                },"
response.write "                {"
response.write "                    extend: 'excel',"
response.write "                    text: 'EXCEL',"
response.write "                    title: '" & dpt & " Labtest-" & FormatDate(Now) & "'"
response.write "                },"
response.write "                {"
response.write "                    extend: 'pdf',"
response.write "                    text: 'PDF',"
response.write "                    title: '" & dpt & " Labtest-" & FormatDate(Now) & "'"
response.write "                },"

response.write "                {"
response.write "                    extend: 'print',"
response.write "                    text: 'PRINT',"
response.write "                    title: '" & dpt & " Labtest-" & FormatDate(Now) & "'"
response.write "                }"

response.write "            ],"

response.write "        columnDefs: ["
response.write "        {"
response.write "            target: 7,"
response.write "            visible: true,"
response.write "            searchable: true"
response.write "        }"
response.write "    ]"

response.write "        });"
response.write "</script>"



End Sub
Sub MedicalExaminationLabResultsAllLIPID(testId)

Dim rst, sql

Set rst = CreateObject("ADODB.Recordset")


  sql = ""
sql = sql & " WITH LabReportRange AS ( "
sql = sql & " SELECT i.LabRequestID, i.LabTestID, i.RequestDate, i.VisitationID, i.PatientID, i.GenderID, i.SponsorID, [L0155].Column2 AS Total_Chol, [L0156].Column2 AS Triglyceride, [L0160].Column2 AS Direct_HDL, [L0031].Column2 AS LDL FROM Investigation i "
sql = sql & " LEFT JOIN Visitation v ON i.VisitationID = v.VisitationID "
    sql = sql & " LEFT JOIN LabResults AS [L0155] ON i.LabRequestID = [L0155].LabRequestID "
        sql = sql & " AND [L0155].TestComponentID = 'L0155' "
    sql = sql & " LEFT JOIN LabResults AS [L0156] ON i.LabRequestID = [L0156].LabRequestID "
        sql = sql & " AND [L0156].TestComponentID = 'L0156' "
    sql = sql & " LEFT JOIN LabResults AS [L0160] ON i.LabRequestID = [L0160].LabRequestID "
        sql = sql & " AND [L0160].TestComponentID = 'L0160' "
    sql = sql & " LEFT JOIN LabResults AS [L0031] ON i.LabRequestID = [L0031].LabRequestID "
        sql = sql & " AND [L0031].TestComponentID = 'L0031' "
    sql = sql & " Where i.LabTestID = '" & testId & "' "
        If Len(branid) > 0 Then
        sql = sql & " AND i.BranchID='" & branid & "' "
    End If
    If Len(genID) > 0 Then
        sql = sql & " AND i.GenderID='" & genID & "' "
    End If
        sql = sql & " AND i.SponsorID IN ('S571','S592','S593', 'S594','S657','VRA') "
        sql = sql & " AND v.SpecialistTypeID IN ('B001-0200','B002-0208','B003-0208') "
        If Len(yrid) > 0 Then
        sql = sql & "  AND i.WorkingYearID = '" & yrid & "' "
    End If
    sql = sql & " Union All "
    sql = sql & " SELECT i2.LabRequestID, i2.LabTestID, i2.RequestDate, i2.VisitationID, i2.PatientID, i2.GenderID, i2.SponsorID, [L0155].Column2 AS Total_Chol, [L0156].Column2 AS Triglyceride, [L0160].Column2 AS Direct_HDL, [L0031].Column2 AS LDL FROM Investigation2 i2 "
    sql = sql & " LEFT JOIN Visitation v ON i2.VisitationID = v.VisitationID "
    sql = sql & " LEFT JOIN LabResults AS [L0155] ON i2.LabRequestID = [L0155].LabRequestID "
        sql = sql & " AND [L0155].TestComponentID = 'L0155' "
    sql = sql & " LEFT JOIN LabResults AS [L0156] ON i2.LabRequestID = [L0156].LabRequestID "
        sql = sql & " AND [L0156].TestComponentID = 'L0156' "
    sql = sql & " LEFT JOIN LabResults AS [L0160] ON i2.LabRequestID = [L0160].LabRequestID "
        sql = sql & " AND [L0160].TestComponentID = 'L0160' "
    sql = sql & " LEFT JOIN LabResults AS [L0031] ON i2.LabRequestID = [L0031].LabRequestID "
        sql = sql & " AND [L0031].TestComponentID = 'L0031' "
    sql = sql & " Where i2.LabTestID = '" & dpt & "' "
    If Len(branid) > 0 Then
        sql = sql & " AND i2.BranchID='" & branid & "' "
    End If
    If Len(genID) > 0 Then
        sql = sql & " AND i2.GenderID='" & genID & "' "
    End If
        
        
        sql = sql & " AND i2.SponsorID IN ('S571','S592','S593', 'S594','S657','VRA') "
        sql = sql & " AND v.SpecialistTypeID IN ('B001-0200','B002-0208','B003-0208') "
          If Len(yrid) > 0 Then
           sql = sql & "  AND i2.WorkingYearID = '" & yrid & "' "
         End If
sql = sql & " ) "
sql = sql & " SELECT LabRequestID, LabTestID,SponsorID, RequestDate,VisitationID,PatientID, GenderID,Total_Chol,Triglyceride,Direct_HDL,LDL FROM LabReportRange ORDER BY RequestDate  "

response.write "<!DOCTYPE html>"
response.write "<html lang=""en"">"

response.write "<head> "
response.write "    <title>AME " & GetComboName("WorkingYear", yrid) & " Lab Tests Records</title>"
'response.write "    <title>AME " & currentYr & " Employee Records</title>"

response.write "<link href=""https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600&display=swap"" rel=""stylesheet"">"
response.write "<link href=""https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css"" rel=""stylesheet"">"
response.write "<link href=""https://cdn.datatables.net/buttons/2.4.1/css/buttons.dataTables.min.css"" rel=""stylesheet"">"
response.write "<link rel=""stylesheet"" href=""https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/css/bootstrap.min.css"" integrity=""sha384-rbsA2VBKQhggwzxH7pPCaAqO46MgnOM80zW1RWuH61DGLwZJEdK2Kadq2F9CUG65"" crossorigin=""anonymous"">"
response.write "<link href=""https://cdn.datatables.net/v/dt/dt-1.13.6/datatables.min.css"" rel=""stylesheet"">"
response.write "<link href=""https://cdn.datatables.net/v/bs5/jszip-3.10.1/dt-1.13.6/b-2.4.1/b-html5-2.4.1/b-print-2.4.1/datatables.min.css"" rel=""stylesheet"">"

AddCss
response.write "</head>"

BuildComboFilter
response.write "<a class=""btn btn-primary1"" href=""wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamStatRCP&PositionForTableName=WorkingDay&WorkingDayID="">Go Home</a>"

response.write "<a class=""btn btn-primary"" href=""wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamLabResultRCP&PositionForTableName=WorkingDay&dptID=" & dpt & "&secID=" & secid & "&genID=" & genID & "&branID=" & branid & "&yrID=" & yrid & " "">Refresh</a>"
response.write "        </td>"
response.write "</table>"
response.write "</div>"
response.write "                    <table id=""example"" class=""table table-striped table-bordered""  width=""100%"">"
response.write "                        <thead class=""beautiful-green-header"">"
response.write "<tr>"
response.write "<th>#</th>"
response.write "<th>LabRequestID</th>"
response.write "<th>LabTestID</th>"
response.write "<th>SponsorID</th>"
response.write "<th>RequestDate</th>"
response.write "<th>VisitationID</th>"
response.write "<th>PatientID</th>"
response.write "<th>GenderID</th>"
response.write "<th>Total_Chol</th>"
response.write "<th>Triglyceride</th>"
response.write "<th>Direct_HDL</th>"
response.write "<th>LDL</th>"
response.write "</tr>"
response.write "</thead>"
response.write "<tbody>"

count = 0

With rst
    rst.open qryPro.FltQry(sql), conn, 3, 4

    If rst.RecordCount > 0 Then
        rst.MoveFirst

        Do While Not rst.EOF

            count = count + 1
   
 
  
            response.write "<tr>"
            response.write "<td>" & count & "</td>"
            response.write "<td>" & rst.fields("LabRequestID") & "</td>"
            response.write "<td>" & GetComboName("LabTest", .fields("LabTestID")) & "</td>"
            response.write "<td>" & GetComboName("Sponsor", .fields("SponsorID")) & "</td>"
            response.write "<td>" & FormatDate(rst.fields("RequestDate")) & "</td>"
            response.write "<td>" & rst.fields("visitationID") & "</td>"
            response.write "<td>" & rst.fields("patientID") & "</td>"
            response.write "<td>" & GetComboName("Gender", .fields("GenderID")) & "</td>"
            response.write "<td>" & rst.fields("Total_Chol") & "</td>"
            response.write "<td>" & rst.fields("Triglyceride") & "</td>"
            response.write "<td>" & rst.fields("Direct_HDL") & "</td>"
            response.write "<td>" & rst.fields("LDL") & "</td>"
      
           
            response.write "</tr>"
            rst.MoveNext
            response.flush
        Loop
    End If
    rst.Close
End With
Set rst = Nothing

response.write "                        </tbody>"
response.write "                    </table>"
response.write "                </div>"
response.write "            </div>"
response.write "        </div>"
response.write "    </div>"




AddJS

response.write "</html>"



response.write "<script>"
response.write "function reload2() {"
response.write "    let selected = document.getElementById('cID').value;"
response.write "    let url = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamLIPIDReport';"
response.write "    url += '&PositionForTableName=WorkingDay&dptID=' + selected;"
response.write "    window.location.href = url;"
response.write "}"

response.write "</script>"
response.write "<script>"
response.write "function reload3() {"
response.write "    let selected = document.getElementById('cID').value;"
'response.write "    let selected2 = document.getElementById('cID2').value;"
response.write "    let selected3 = document.getElementById('cID3').value;"
response.write "    let selected4 = document.getElementById('cID4').value;"
response.write "    let selected5 = document.getElementById('cID5').value;"
'response.write "    let sat = document.getElementById('statusFilter').value;"
response.write "    let url = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamLIPIDReport';"
response.write "    url += '&PositionForTableName=WorkingDay&dptID=' + selected +  '&genID=' + selected3 + '&branID=' + selected4 + '&yrID=' + selected5;"
response.write "    window.location.href = url;"
response.write "}"



response.write "</script>"
response.write "<script>"

response.write "        var dataTable = $('#example').DataTable({"
response.write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, ""All""]],"
response.write "        dom: 'lBfrtip',"
response.write "            pageLength: 10,"
response.write "            paging: true, "
response.write "            pagingType: 'simple', "
response.write "            search: {"
response.write "                smart: true"
response.write "                    },"
response.write "            buttons: ["
response.write "                {"
response.write "                    extend: 'csv',"
response.write "                    text: 'CSV',"
response.write "                    title: '" & dpt & " Labtest-" & FormatDate(Now) & "' "
response.write "                },"
response.write "                {"
response.write "                    extend: 'excel',"
response.write "                    text: 'EXCEL',"
response.write "                    title: '" & dpt & " Labtest-" & FormatDate(Now) & "'"
response.write "                },"
response.write "                {"
response.write "                    extend: 'pdf',"
response.write "                    text: 'PDF',"
response.write "                    title: '" & dpt & " Labtest-" & FormatDate(Now) & "'"
response.write "                },"

response.write "                {"
response.write "                    extend: 'print',"
response.write "                    text: 'PRINT',"
response.write "                    title: '" & dpt & " Labtest-" & FormatDate(Now) & "'"
response.write "                }"

response.write "            ],"

response.write "        columnDefs: ["
response.write "        {"
response.write "            target: 7,"
response.write "            visible: true,"
response.write "            searchable: true"
response.write "        }"
response.write "    ]"

response.write "        });"
response.write "</script>"



End Sub
Sub MedicalExaminationLabResultsAllESR(testId)

Dim rst, sql

Set rst = CreateObject("ADODB.Recordset")


  sql = ""
sql = sql & " WITH LabReportRange AS ( "
sql = sql & " SELECT i.LabRequestID, i.LabTestID, i.RequestDate, i.VisitationID, i.PatientID, i.GenderID, i.SponsorID, [L0043].Column2 AS ESR FROM Investigation i "
sql = sql & " LEFT JOIN Visitation v ON i.VisitationID = v.VisitationID "
    sql = sql & " LEFT JOIN LabResults AS [L0043] ON i.LabRequestID = [L0043].LabRequestID "
        sql = sql & " AND [L0043].TestComponentID = 'L0043' "
    
    sql = sql & " Where i.LabTestID = '" & testId & "' "
        If Len(branid) > 0 Then
        sql = sql & " AND i.BranchID='" & branid & "' "
    End If
    If Len(genID) > 0 Then
        sql = sql & " AND i.GenderID='" & genID & "' "
    End If
        sql = sql & " AND i.SponsorID IN ('S571','S592','S593', 'S594','S657','VRA') "
        sql = sql & " AND v.SpecialistTypeID IN ('B001-0200','B002-0208','B003-0208') "
        If Len(yrid) > 0 Then
        sql = sql & "  AND i.WorkingYearID = '" & yrid & "' "
    End If
    sql = sql & " Union All "
    sql = sql & " SELECT i2.LabRequestID, i2.LabTestID, i2.RequestDate, i2.VisitationID, i2.PatientID, i2.GenderID, i2.SponsorID, [L0043].Column2 AS ESR FROM Investigation2 i2 "
    sql = sql & " LEFT JOIN Visitation v ON i2.VisitationID = v.VisitationID "
    sql = sql & " LEFT JOIN LabResults AS [L0043] ON i2.LabRequestID = [L0043].LabRequestID "
        sql = sql & " AND [L0043].TestComponentID = 'L0043' "
    sql = sql & " Where i2.LabTestID = '" & dpt & "' "
    If Len(branid) > 0 Then
        sql = sql & " AND i2.BranchID='" & branid & "' "
    End If
    If Len(genID) > 0 Then
        sql = sql & " AND i2.GenderID='" & genID & "' "
    End If
        
        
        sql = sql & " AND i2.SponsorID IN ('S571','S592','S593', 'S594','S657','VRA') "
        sql = sql & " AND v.SpecialistTypeID IN ('B001-0200','B002-0208','B003-0208') "
          If Len(yrid) > 0 Then
           sql = sql & "  AND i2.WorkingYearID = '" & yrid & "' "
         End If
sql = sql & " ) "
sql = sql & " SELECT LabRequestID, LabTestID,SponsorID, RequestDate,VisitationID,PatientID, GenderID,ESR FROM LabReportRange ORDER BY RequestDate  "

response.write "<!DOCTYPE html>"
response.write "<html lang=""en"">"

response.write "<head> "
response.write "    <title>AME " & GetComboName("WorkingYear", yrid) & " Lab Tests Records</title>"
'response.write "    <title>AME " & currentYr & " Employee Records</title>"

response.write "<link href=""https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600&display=swap"" rel=""stylesheet"">"
response.write "<link href=""https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css"" rel=""stylesheet"">"
response.write "<link href=""https://cdn.datatables.net/buttons/2.4.1/css/buttons.dataTables.min.css"" rel=""stylesheet"">"
response.write "<link rel=""stylesheet"" href=""https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/css/bootstrap.min.css"" integrity=""sha384-rbsA2VBKQhggwzxH7pPCaAqO46MgnOM80zW1RWuH61DGLwZJEdK2Kadq2F9CUG65"" crossorigin=""anonymous"">"
response.write "<link href=""https://cdn.datatables.net/v/dt/dt-1.13.6/datatables.min.css"" rel=""stylesheet"">"
response.write "<link href=""https://cdn.datatables.net/v/bs5/jszip-3.10.1/dt-1.13.6/b-2.4.1/b-html5-2.4.1/b-print-2.4.1/datatables.min.css"" rel=""stylesheet"">"

AddCss
response.write "</head>"

BuildComboFilter
response.write "<a class=""btn btn-primary1"" href=""wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamStatRCP&PositionForTableName=WorkingDay&WorkingDayID="">Go Home</a>"

response.write "<a class=""btn btn-primary"" href=""wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamLabResultRCP&PositionForTableName=WorkingDay&dptID=" & dpt & "&secID=" & secid & "&genID=" & genID & "&branID=" & branid & "&yrID=" & yrid & " "">Refresh</a>"
response.write "        </td>"
response.write "</table>"
response.write "</div>"
response.write "                    <table id=""example"" class=""table table-striped table-bordered""  width=""100%"">"
response.write "                        <thead class=""beautiful-green-header"">"
response.write "<tr>"
response.write "<th>#</th>"
response.write "<th>LabRequestID</th>"
response.write "<th>LabTestID</th>"
response.write "<th>SponsorID</th>"
response.write "<th>RequestDate</th>"
response.write "<th>VisitationID</th>"
response.write "<th>PatientID</th>"
response.write "<th>GenderID</th>"
response.write "<th>ESR</th>"
response.write "</tr>"
response.write "</thead>"
response.write "<tbody>"

count = 0

With rst
    rst.open qryPro.FltQry(sql), conn, 3, 4

    If rst.RecordCount > 0 Then
        rst.MoveFirst

        Do While Not rst.EOF

            count = count + 1
   
 
  
            response.write "<tr>"
            response.write "<td>" & count & "</td>"
            response.write "<td>" & rst.fields("LabRequestID") & "</td>"
            response.write "<td>" & GetComboName("LabTest", .fields("LabTestID")) & "</td>"
            response.write "<td>" & GetComboName("Sponsor", .fields("SponsorID")) & "</td>"
            response.write "<td>" & FormatDate(rst.fields("RequestDate")) & "</td>"
            response.write "<td>" & rst.fields("visitationID") & "</td>"
            response.write "<td>" & rst.fields("patientID") & "</td>"
            response.write "<td>" & GetComboName("Gender", .fields("GenderID")) & "</td>"
            response.write "<td>" & rst.fields("ESR") & "</td>"
          
      
           
            response.write "</tr>"
            rst.MoveNext
            response.flush
        Loop
    End If
    rst.Close
End With
Set rst = Nothing

response.write "                        </tbody>"
response.write "                    </table>"
response.write "                </div>"
response.write "            </div>"
response.write "        </div>"
response.write "    </div>"




AddJS

response.write "</html>"



response.write "<script>"
response.write "function reload2() {"
response.write "    let selected = document.getElementById('cID').value;"
response.write "    let url = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamESRReport';"
response.write "    url += '&PositionForTableName=WorkingDay&dptID=' + selected;"
response.write "    window.location.href = url;"
response.write "}"

response.write "</script>"
response.write "<script>"
response.write "function reload3() {"
response.write "    let selected = document.getElementById('cID').value;"
'response.write "    let selected2 = document.getElementById('cID2').value;"
response.write "    let selected3 = document.getElementById('cID3').value;"
response.write "    let selected4 = document.getElementById('cID4').value;"
response.write "    let selected5 = document.getElementById('cID5').value;"
'response.write "    let sat = document.getElementById('statusFilter').value;"
response.write "    let url = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamESRReport';"
response.write "    url += '&PositionForTableName=WorkingDay&dptID=' + selected +  '&genID=' + selected3 + '&branID=' + selected4 + '&yrID=' + selected5;"
response.write "    window.location.href = url;"
response.write "}"



response.write "</script>"
response.write "<script>"

response.write "        var dataTable = $('#example').DataTable({"
response.write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, ""All""]],"
response.write "        dom: 'lBfrtip',"
response.write "            pageLength: 10,"
response.write "            paging: true, "
response.write "            pagingType: 'simple', "
response.write "            search: {"
response.write "                smart: true"
response.write "                    },"
response.write "            buttons: ["
response.write "                {"
response.write "                    extend: 'csv',"
response.write "                    text: 'CSV',"
response.write "                    title: '" & dpt & " Labtest-" & FormatDate(Now) & "' "
response.write "                },"
response.write "                {"
response.write "                    extend: 'excel',"
response.write "                    text: 'EXCEL',"
response.write "                    title: '" & dpt & " Labtest-" & FormatDate(Now) & "'"
response.write "                },"
response.write "                {"
response.write "                    extend: 'pdf',"
response.write "                    text: 'PDF',"
response.write "                    title: '" & dpt & " Labtest-" & FormatDate(Now) & "'"
response.write "                },"

response.write "                {"
response.write "                    extend: 'print',"
response.write "                    text: 'PRINT',"
response.write "                    title: '" & dpt & " Labtest-" & FormatDate(Now) & "'"
response.write "                }"

response.write "            ],"

response.write "        columnDefs: ["
response.write "        {"
response.write "            target: 7,"
response.write "            visible: true,"
response.write "            searchable: true"
response.write "        }"
response.write "    ]"

response.write "        });"
response.write "</script>"



End Sub
Sub MedicalExaminationLabResultsAllUR(testId)

Dim rst, sql

Set rst = CreateObject("ADODB.Recordset")



  sql = ""
sql = sql & " WITH LabReportRange AS ( "
sql = sql & " SELECT i.LabRequestID, i.LabTestID, i.RequestDate, i.VisitationID, i.PatientID, i.GenderID, i.SponsorID, [L0133].Column2 AS Appearance, [cl].Column4 AS Colour,[L0504].Column2 AS UROBILINOGEN,[L0505].Column2 AS BILIRUBIN,[L0506].Column2 AS KETONES,[L0507].Column2 AS BLOOD,[L0509].Column2 AS NITRITES,[L0508].Column2 AS PROTEIN,[L0510].Column2 AS LEUCOCYTES,[L0511].Column2 AS GLUCOSE,[L0512].Column2 AS SPECIFIC_GRAVITY,[L0513].Column2 AS P_H,[L0514].Column2 AS MICROALBUMIN,[L0178].Column2 AS Pus_Cells_per_HPF,[yeast].Column4 AS Yeast_like_Cells,[L0179].Column2 AS Epithelial_Cells_per_HPF,[haema].Column4 AS S_haematobium,[L0180].Column2 AS RBC_per_HPF,[vaginalis].Column4 AS T_vaginalis,[L0181].Column4 AS Spermatozoa,[Bacteria].Column4 AS Bacteria,[L0242].Column1 AS Crystals,[CrystalsP].Column2 AS Crystals_Present,[L0245].Column1 AS Cast,[CastP].Column2 AS Cast_Present FROM Investigation i "
sql = sql & " LEFT JOIN Visitation v ON i.VisitationID = v.VisitationID "
    sql = sql & "  LEFT JOIN LabResults AS [L0133] ON i.LabRequestID = [L0133].LabRequestID "
        sql = sql & " AND [L0133].TestComponentID = 'L0133' "
    sql = sql & "  LEFT JOIN LabResults AS [cl] ON i.LabRequestID = [cl].LabRequestID "
        sql = sql & " AND [cl].TestComponentID = 'L0133' "
    sql = sql & "  LEFT JOIN LabResults AS [L0504] ON i.LabRequestID = [L0504].LabRequestID "
        sql = sql & " AND [L0504].TestComponentID = 'L0504' "
    sql = sql & "  LEFT JOIN LabResults AS [L0505] ON i.LabRequestID = [L0505].LabRequestID "
        sql = sql & " AND [L0505].TestComponentID = 'L0505' "
    sql = sql & "  LEFT JOIN LabResults AS [L0506] ON i.LabRequestID = [L0506].LabRequestID "
        sql = sql & " AND [L0506].TestComponentID = 'L0506' "
    sql = sql & "  LEFT JOIN LabResults AS [L0507] ON i.LabRequestID = [L0507].LabRequestID "
        sql = sql & " AND [L0507].TestComponentID = 'L0507' "
    sql = sql & "  LEFT JOIN LabResults AS [L0509] ON i.LabRequestID = [L0509].LabRequestID "
        sql = sql & " AND [L0509].TestComponentID = 'L0509' "
    sql = sql & "  LEFT JOIN LabResults AS [L0508] ON i.LabRequestID = [L0508].LabRequestID "
        sql = sql & " AND [L0508].TestComponentID = 'L0508' "
    sql = sql & "  LEFT JOIN LabResults AS [L0510] ON i.LabRequestID = [L0510].LabRequestID "
        sql = sql & " AND [L0510].TestComponentID = 'L0510' "
    sql = sql & "  LEFT JOIN LabResults AS [L0511] ON i.LabRequestID = [L0511].LabRequestID "
        sql = sql & " AND [L0511].TestComponentID = 'L0511' "
    sql = sql & "  LEFT JOIN LabResults AS [L0512] ON i.LabRequestID = [L0512].LabRequestID "
        sql = sql & " AND [L0512].TestComponentID = 'L0512' "
    sql = sql & "  LEFT JOIN LabResults AS [L0513] ON i.LabRequestID = [L0513].LabRequestID "
        sql = sql & " AND [L0513].TestComponentID = 'L0513' "
    sql = sql & "  LEFT JOIN LabResults AS [L0514] ON i.LabRequestID = [L0514].LabRequestID "
        sql = sql & " AND [L0514].TestComponentID = 'L0514' "
    sql = sql & "  LEFT JOIN LabResults AS [L0178] ON i.LabRequestID = [L0178].LabRequestID "
        sql = sql & " AND [L0178].TestComponentID = 'L0178' "
    sql = sql & "  LEFT JOIN LabResults AS [yeast] ON i.LabRequestID = [yeast].LabRequestID "
        sql = sql & " AND [yeast].TestComponentID = 'L0178' "
    sql = sql & "  LEFT JOIN LabResults AS [L0179] ON i.LabRequestID = [L0179].LabRequestID "
        sql = sql & " AND [L0179].TestComponentID = 'L0179' "
    sql = sql & "  LEFT JOIN LabResults AS [haema] ON i.LabRequestID = [haema].LabRequestID "
        sql = sql & " AND [haema].TestComponentID = 'L0179' "
    sql = sql & "  LEFT JOIN LabResults AS [L0180] ON i.LabRequestID = [L0180].LabRequestID "
        sql = sql & " AND [L0180].TestComponentID = 'L0180' "
    sql = sql & "  LEFT JOIN LabResults AS [vaginalis] ON i.LabRequestID = [vaginalis].LabRequestID "
        sql = sql & " AND [vaginalis].TestComponentID = 'L0180' "
    sql = sql & "  LEFT JOIN LabResults AS [L0181] ON i.LabRequestID = [L0181].LabRequestID "
        sql = sql & " AND [L0181].TestComponentID = 'L0181' "
    sql = sql & "  LEFT JOIN LabResults AS [Bacteria] ON i.LabRequestID = [Bacteria].LabRequestID "
        sql = sql & " AND [Bacteria].TestComponentID = 'L0181' "
    sql = sql & "  LEFT JOIN LabResults AS [L0242] ON i.LabRequestID = [L0242].LabRequestID "
        sql = sql & " AND [L0242].TestComponentID = 'L0242' "
    sql = sql & "  LEFT JOIN LabResults AS [CrystalsP] ON i.LabRequestID = [CrystalsP].LabRequestID "
        sql = sql & " AND [CrystalsP].TestComponentID = 'L0242' "
    sql = sql & "  LEFT JOIN LabResults AS [L0245] ON i.LabRequestID = [L0245].LabRequestID "
        sql = sql & " AND [L0245].TestComponentID = 'L0245' "
    sql = sql & "  LEFT JOIN LabResults AS [CastP] ON i.LabRequestID = [CastP].LabRequestID "
        sql = sql & " AND [CastP].TestComponentID = 'L0245' "
    
    sql = sql & " Where i.LabTestID = '" & testId & "' "
        If Len(branid) > 0 Then
        sql = sql & " AND i.BranchID='" & branid & "' "
    End If
    If Len(genID) > 0 Then
        sql = sql & " AND i.GenderID='" & genID & "' "
    End If
        sql = sql & " AND i.SponsorID IN ('S571','S592','S593', 'S594','S657','VRA') "
        sql = sql & " AND v.SpecialistTypeID IN ('B001-0200','B002-0208','B003-0208') "
        If Len(yrid) > 0 Then
        sql = sql & "  AND i.WorkingYearID = '" & yrid & "' "
    End If
    sql = sql & " Union All "
    sql = sql & " SELECT i2.LabRequestID, i2.LabTestID, i2.RequestDate, i2.VisitationID, i2.PatientID, i2.GenderID, i2.SponsorID, [L0133].Column2 AS Appearance, [cl].Column4 AS Colour,[L0504].Column2 AS UROBILINOGEN,[L0505].Column2 AS BILIRUBIN,[L0506].Column2 AS KETONES,[L0507].Column2 AS BLOOD,[L0509].Column2 AS NITRITES,[L0508].Column2 AS PROTEIN,[L0510].Column2 AS LEUCOCYTES,[L0511].Column2 AS GLUCOSE,[L0512].Column2 AS SPECIFIC_GRAVITY,[L0513].Column2 AS P_H,[L0514].Column2 AS MICROALBUMIN,[L0178].Column2 AS Pus_Cells_per_HPF,[yeast].Column4 AS Yeast_like_Cells,[L0179].Column2 AS Epithelial_Cells_per_HPF,[haema].Column4 AS S_haematobium,[L0180].Column2 AS RBC_per_HPF,[vaginalis].Column4 AS T_vaginalis,[L0181].Column4 AS Spermatozoa,[Bacteria].Column4 AS Bacteria,[L0242].Column1 AS Crystals,[CrystalsP].Column2 AS Crystals_Present,[L0245].Column1 AS Cast,[CastP].Column2 AS Cast_Present FROM Investigation2 i2 "
    sql = sql & " LEFT JOIN Visitation v ON i2.VisitationID = v.VisitationID "
    sql = sql & "  LEFT JOIN LabResults AS [L0133] ON i2.LabRequestID = [L0133].LabRequestID "
        sql = sql & " AND [L0133].TestComponentID = 'L0133' "
    sql = sql & "  LEFT JOIN LabResults AS [cl] ON i2.LabRequestID = [cl].LabRequestID "
        sql = sql & " AND [cl].TestComponentID = 'L0133' "
     sql = sql & "  LEFT JOIN LabResults AS [L0504] ON i2.LabRequestID = [L0504].LabRequestID "
        sql = sql & " AND [L0504].TestComponentID = 'L0504' "
    sql = sql & "  LEFT JOIN LabResults AS [L0505] ON i2.LabRequestID = [L0505].LabRequestID "
        sql = sql & " AND [L0505].TestComponentID = 'L0505' "
    sql = sql & "  LEFT JOIN LabResults AS [L0506] ON i2.LabRequestID = [L0506].LabRequestID "
        sql = sql & " AND [L0506].TestComponentID = 'L0506' "
    sql = sql & "  LEFT JOIN LabResults AS [L0507] ON i2.LabRequestID = [L0507].LabRequestID "
        sql = sql & " AND [L0507].TestComponentID = 'L0507' "
    sql = sql & "  LEFT JOIN LabResults AS [L0509] ON i2.LabRequestID = [L0509].LabRequestID "
        sql = sql & " AND [L0509].TestComponentID = 'L0509' "
    sql = sql & "  LEFT JOIN LabResults AS [L0508] ON i2.LabRequestID = [L0508].LabRequestID "
        sql = sql & " AND [L0508].TestComponentID = 'L0508' "
    sql = sql & "  LEFT JOIN LabResults AS [L0510] ON i2.LabRequestID = [L0510].LabRequestID "
        sql = sql & " AND [L0510].TestComponentID = 'L0510' "
    sql = sql & "  LEFT JOIN LabResults AS [L0511] ON i2.LabRequestID = [L0511].LabRequestID "
        sql = sql & " AND [L0511].TestComponentID = 'L0511' "
    sql = sql & "  LEFT JOIN LabResults AS [L0512] ON i2.LabRequestID = [L0512].LabRequestID "
        sql = sql & " AND [L0512].TestComponentID = 'L0512' "
    sql = sql & "  LEFT JOIN LabResults AS [L0513] ON i2.LabRequestID = [L0513].LabRequestID "
        sql = sql & " AND [L0513].TestComponentID = 'L0513' "
    sql = sql & "  LEFT JOIN LabResults AS [L0514] ON i2.LabRequestID = [L0514].LabRequestID "
        sql = sql & " AND [L0514].TestComponentID = 'L0514' "
    sql = sql & "  LEFT JOIN LabResults AS [L0178] ON i2.LabRequestID = [L0178].LabRequestID "
        sql = sql & " AND [L0178].TestComponentID = 'L0178' "
    sql = sql & "  LEFT JOIN LabResults AS [yeast] ON i2.LabRequestID = [yeast].LabRequestID "
        sql = sql & " AND [yeast].TestComponentID = 'L0178' "
    sql = sql & "  LEFT JOIN LabResults AS [L0179] ON i2.LabRequestID = [L0179].LabRequestID "
        sql = sql & " AND [L0179].TestComponentID = 'L0179' "
    sql = sql & "  LEFT JOIN LabResults AS [haema] ON i2.LabRequestID = [haema].LabRequestID "
        sql = sql & " AND [haema].TestComponentID = 'L0179' "
    sql = sql & "  LEFT JOIN LabResults AS [L0180] ON i2.LabRequestID = [L0180].LabRequestID "
        sql = sql & " AND [L0180].TestComponentID = 'L0180' "
    sql = sql & "  LEFT JOIN LabResults AS [vaginalis] ON i2.LabRequestID = [vaginalis].LabRequestID "
        sql = sql & " AND [vaginalis].TestComponentID = 'L0180' "
    sql = sql & "  LEFT JOIN LabResults AS [L0181] ON i2.LabRequestID = [L0181].LabRequestID "
        sql = sql & " AND [L0181].TestComponentID = 'L0181' "
    sql = sql & "  LEFT JOIN LabResults AS [Bacteria] ON i2.LabRequestID = [Bacteria].LabRequestID "
        sql = sql & " AND [Bacteria].TestComponentID = 'L0181' "
    sql = sql & "  LEFT JOIN LabResults AS [L0242] ON i2.LabRequestID = [L0242].LabRequestID "
        sql = sql & " AND [L0242].TestComponentID = 'L0242' "
    sql = sql & "  LEFT JOIN LabResults AS [CrystalsP] ON i2.LabRequestID = [CrystalsP].LabRequestID "
        sql = sql & " AND [CrystalsP].TestComponentID = 'L0242' "
    sql = sql & "  LEFT JOIN LabResults AS [L0245] ON i2.LabRequestID = [L0245].LabRequestID "
        sql = sql & " AND [L0245].TestComponentID = 'L0245' "
    sql = sql & "  LEFT JOIN LabResults AS [CastP] ON i2.LabRequestID = [CastP].LabRequestID "
        sql = sql & " AND [CastP].TestComponentID = 'L0245' "
    sql = sql & " Where i2.LabTestID = '" & dpt & "' "
    If Len(branid) > 0 Then
        sql = sql & " AND i2.BranchID='" & branid & "' "
    End If
    If Len(genID) > 0 Then
        sql = sql & " AND i2.GenderID='" & genID & "' "
    End If
        
        
        sql = sql & " AND i2.SponsorID IN ('S571','S592','S593', 'S594','S657','VRA') "
        sql = sql & " AND v.SpecialistTypeID IN ('B001-0200','B002-0208','B003-0208') "
          If Len(yrid) > 0 Then
           sql = sql & "  AND i2.WorkingYearID = '" & yrid & "' "
         End If
sql = sql & " ) "
sql = sql & " SELECT LabRequestID, LabTestID,SponsorID, RequestDate,VisitationID,PatientID, GenderID,Appearance, Colour,UROBILINOGEN,BILIRUBIN,KETONES,BLOOD,NITRITES,PROTEIN,LEUCOCYTES,GLUCOSE,SPECIFIC_GRAVITY,P_H,MICROALBUMIN,Pus_Cells_per_HPF,Yeast_like_Cells,Epithelial_Cells_per_HPF,S_haematobium,RBC_per_HPF,T_vaginalis,Spermatozoa,Bacteria,Crystals,Crystals_Present,Cast,Cast_Present  FROM LabReportRange ORDER BY RequestDate  "

response.write "<!DOCTYPE html>"
response.write "<html lang=""en"">"

response.write "<head> "
response.write "    <title>AME " & GetComboName("WorkingYear", yrid) & " Lab Tests Records</title>"
'response.write "    <title>AME " & currentYr & " Employee Records</title>"

response.write "<link href=""https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600&display=swap"" rel=""stylesheet"">"
response.write "<link href=""https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css"" rel=""stylesheet"">"
response.write "<link href=""https://cdn.datatables.net/buttons/2.4.1/css/buttons.dataTables.min.css"" rel=""stylesheet"">"
response.write "<link rel=""stylesheet"" href=""https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/css/bootstrap.min.css"" integrity=""sha384-rbsA2VBKQhggwzxH7pPCaAqO46MgnOM80zW1RWuH61DGLwZJEdK2Kadq2F9CUG65"" crossorigin=""anonymous"">"
response.write "<link href=""https://cdn.datatables.net/v/dt/dt-1.13.6/datatables.min.css"" rel=""stylesheet"">"
response.write "<link href=""https://cdn.datatables.net/v/bs5/jszip-3.10.1/dt-1.13.6/b-2.4.1/b-html5-2.4.1/b-print-2.4.1/datatables.min.css"" rel=""stylesheet"">"

AddCss
response.write "</head>"

BuildComboFilter
response.write "<a class=""btn btn-primary1"" href=""wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamStatRCP&PositionForTableName=WorkingDay&WorkingDayID="">Go Home</a>"

response.write "<a class=""btn btn-primary"" href=""wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamLabResultRCP&PositionForTableName=WorkingDay&dptID=" & dpt & "&secID=" & secid & "&genID=" & genID & "&branID=" & branid & "&yrID=" & yrid & " "">Refresh</a>"
response.write "        </td>"
response.write "</table>"
response.write "</div>"
response.write "                    <table id=""example"" class=""table table-striped table-bordered""  width=""100%"">"
response.write "                        <thead class=""beautiful-green-header"">"
response.write "<tr>"
response.write "<th>#</th>"
response.write "<th>LabRequestID</th>"
response.write "<th>LabTestID</th>"
response.write "<th>SponsorID</th>"
response.write "<th>RequestDate</th>"
response.write "<th>VisitationID</th>"
response.write "<th>PatientID</th>"
response.write "<th>GenderID</th>"
response.write "<th>Appearance</th>"
response.write "<th>Colour</th>"
response.write "<th>UROBILINOGEN</th>"
response.write "<th>BILIRUBIN</th>"
response.write "<th>KETONES</th>"
response.write "<th>BLOOD</th>"
response.write "<th>NITRITES</th>"
response.write "<th>PROTEIN</th>"
response.write "<th>LEUCOCYTES</th>"
response.write "<th> GLUCOSE</th>"
response.write "<th>SPECIFIC_GRAVITY</th>"
response.write "<th>P_H</th>"
response.write "<th>MICROALBUMIN</th>"
response.write "<th>Pus_Cells_per_HPF</th>"
response.write "<th>Yeast_like_Cells</th>"
response.write "<th>Epithelial_Cells_per_HPF</th>"
response.write "<th>S_haematobium</th>"
response.write "<th>RBC_per_HPF</th>"
response.write "<th>T_vaginalis</th>"
response.write "<th>Spermatozoa</th>"
response.write "<th>Bacteria</th>"
response.write "<th>Crystals</th>"
response.write "<th>Crystals Presence</th>"
response.write "<th>Cast</th>"
response.write "<th>Cast Presence</th>"
response.write "</tr>"
response.write "</thead>"
response.write "<tbody>"
                    
count = 0

With rst
    rst.open qryPro.FltQry(sql), conn, 3, 4

    If rst.RecordCount > 0 Then
        rst.MoveFirst

        Do While Not rst.EOF

            count = count + 1
   
 
  
            response.write "<tr>"
            response.write "<td>" & count & "</td>"
            response.write "<td>" & rst.fields("LabRequestID") & "</td>"
            response.write "<td>" & GetComboName("LabTest", .fields("LabTestID")) & "</td>"
            response.write "<td>" & GetComboName("Sponsor", .fields("SponsorID")) & "</td>"
            response.write "<td>" & FormatDate(rst.fields("RequestDate")) & "</td>"
            response.write "<td>" & rst.fields("visitationID") & "</td>"
            response.write "<td>" & rst.fields("patientID") & "</td>"
            response.write "<td>" & GetComboName("Gender", .fields("GenderID")) & "</td>"
            response.write "<td>" & GetComboName("T_Appearance", rst.fields("Appearance")) & "</td>"
            response.write "<td>" & GetComboName("T_Colour", rst.fields("Colour")) & "</td>"
            response.write "<td>" & GetComboName("T_UROBILINOGEN", rst.fields("UROBILINOGEN")) & "</td>"
            response.write "<td>" & GetComboName("T_NegPosExtra", rst.fields("BILIRUBIN")) & "</td>"
            response.write "<td>" & GetComboName("T_NegPosExtra", rst.fields("KETONES")) & "</td>"
            response.write "<td>" & GetComboName("T_NegPosExtra", rst.fields("BLOOD")) & "</td>"
            response.write "<td>" & GetComboName("T_NegPosExtra", rst.fields("NITRITES")) & "</td>"
            response.write "<td>" & GetComboName("T_NegPosExtra", rst.fields("PROTEIN")) & "</td>"
            response.write "<td>" & GetComboName("T_NegPosExtra", rst.fields("LEUCOCYTES")) & "</td>"
            response.write "<td>" & GetComboName("T_NegPosExtra", rst.fields("GLUCOSE")) & "</td>"
            response.write "<td>" & GetComboName("T_SpGravity", rst.fields("SPECIFIC_GRAVITY")) & "</td>"
            response.write "<td>" & GetComboName("T_PH", rst.fields("P_H")) & "</td>"
            response.write "<td>" & GetComboName("T_NegPosExtra", rst.fields("MICROALBUMIN")) & "</td>"
            response.write "<td>" & rst.fields("Pus_Cells_per_HPF") & "</td>"
            response.write "<td>" & GetComboName("T_Present", rst.fields("Yeast_like_Cells")) & "</td>"
            response.write "<td>" & rst.fields("Epithelial_Cells_per_HPF") & "</td>"
            response.write "<td>" & GetComboName("T_Present", rst.fields("S_haematobium")) & "</td>"
            response.write "<td>" & rst.fields("RBC_per_HPF") & "</td>"
            response.write "<td>" & GetComboName("T_Present", rst.fields("T_vaginalis")) & "</td>"
            response.write "<td>" & GetComboName("T_Present", rst.fields("Spermatozoa")) & "</td>"
            response.write "<td>" & GetComboName("T_Present", rst.fields("Bacteria")) & "</td>"
            response.write "<td>" & GetComboName("T_Crystals", rst.fields("Crystals")) & "</td>"
            response.write "<td>" & GetComboName("T_Present", rst.fields("Crystals_Present")) & "</td>"
            response.write "<td>" & GetComboName("T_Cast", rst.fields("Cast")) & "</td>"
            response.write "<td>" & GetComboName("T_Present", rst.fields("Cast_Present")) & "</td>"
          

            response.write "</tr>"
            rst.MoveNext
            response.flush
        Loop
    End If
    rst.Close
End With
Set rst = Nothing

response.write "                        </tbody>"
response.write "                    </table>"
response.write "                </div>"
response.write "            </div>"
response.write "        </div>"
response.write "    </div>"




AddJS

response.write "</html>"



response.write "<script>"
response.write "function reload2() {"
response.write "    let selected = document.getElementById('cID').value;"
response.write "    let url = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamLabResultRCP';"
response.write "    url += '&PositionForTableName=WorkingDay&dptID=' + selected;"
response.write "    window.location.href = url;"
response.write "}"

response.write "</script>"
response.write "<script>"
response.write "function reload3() {"
response.write "    let selected = document.getElementById('cID').value;"
response.write "    let selected2 = document.getElementById('cID2').value;"
response.write "    let selected3 = document.getElementById('cID3').value;"
response.write "    let selected4 = document.getElementById('cID4').value;"
response.write "    let selected5 = document.getElementById('cID5').value;"
'response.write "    let sat = document.getElementById('statusFilter').value;"
response.write "    let url = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamLabResultRCP';"
response.write "    url += '&PositionForTableName=WorkingDay&dptID=' + selected + '&secID=' + selected2 + '&genID=' + selected3 + '&branID=' + selected4 + '&yrID=' + selected5;"
response.write "    window.location.href = url;"
response.write "}"



response.write "</script>"



End Sub
Sub MedicalExaminationLabResultsAllPSA(testId)

Dim rst, sql

Set rst = CreateObject("ADODB.Recordset")


  sql = ""
sql = sql & " WITH LabReportRange AS ( "
sql = sql & " SELECT i.LabRequestID, i.LabTestID, i.RequestDate, i.VisitationID, i.PatientID, i.GenderID, i.SponsorID, [L0125].Column2 AS PSA FROM Investigation i "
sql = sql & " LEFT JOIN Visitation v ON i.VisitationID = v.VisitationID "
    sql = sql & " LEFT JOIN LabResults AS [L0125] ON i.LabRequestID = [L0125].LabRequestID "
        sql = sql & " AND [L0125].TestComponentID = 'L0125' "
    
    sql = sql & " Where i.LabTestID = '" & testId & "' "
        If Len(branid) > 0 Then
        sql = sql & " AND i.BranchID='" & branid & "' "
    End If
    If Len(genID) > 0 Then
        sql = sql & " AND i.GenderID='" & genID & "' "
    End If
        sql = sql & " AND i.SponsorID IN ('S571','S592','S593', 'S594','S657','VRA') "
        sql = sql & " AND v.SpecialistTypeID IN ('B001-0200','B002-0208','B003-0208') "
        If Len(yrid) > 0 Then
        sql = sql & "  AND i.WorkingYearID = '" & yrid & "' "
    End If
    sql = sql & " Union All "
    sql = sql & " SELECT i2.LabRequestID, i2.LabTestID, i2.RequestDate, i2.VisitationID, i2.PatientID, i2.GenderID, i2.SponsorID, [L0125].Column2 AS PSA FROM Investigation2 i2 "
    sql = sql & " LEFT JOIN Visitation v ON i2.VisitationID = v.VisitationID "
     sql = sql & " LEFT JOIN LabResults AS [L0125] ON i2.LabRequestID = [L0125].LabRequestID "
        sql = sql & " AND [L0125].TestComponentID = 'L0125' "
    
    sql = sql & " Where i2.LabTestID = '" & dpt & "' "
    If Len(branid) > 0 Then
        sql = sql & " AND i2.BranchID='" & branid & "' "
    End If
    If Len(genID) > 0 Then
        sql = sql & " AND i2.GenderID='" & genID & "' "
    End If
        
        
        sql = sql & " AND i2.SponsorID IN ('S571','S592','S593', 'S594','S657','VRA') "
        sql = sql & " AND v.SpecialistTypeID IN ('B001-0200','B002-0208','B003-0208') "
          If Len(yrid) > 0 Then
           sql = sql & "  AND i2.WorkingYearID = '" & yrid & "' "
         End If
sql = sql & " ) "
sql = sql & " SELECT LabRequestID, LabTestID,SponsorID, RequestDate,VisitationID,PatientID, GenderID,PSA FROM LabReportRange ORDER BY RequestDate  "

response.write "<!DOCTYPE html>"
response.write "<html lang=""en"">"

response.write "<head> "
response.write "    <title>AME " & GetComboName("WorkingYear", yrid) & " Lab Tests Records</title>"
'response.write "    <title>AME " & currentYr & " Employee Records</title>"

response.write "<link href=""https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600&display=swap"" rel=""stylesheet"">"
response.write "<link href=""https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css"" rel=""stylesheet"">"
response.write "<link href=""https://cdn.datatables.net/buttons/2.4.1/css/buttons.dataTables.min.css"" rel=""stylesheet"">"
response.write "<link rel=""stylesheet"" href=""https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/css/bootstrap.min.css"" integrity=""sha384-rbsA2VBKQhggwzxH7pPCaAqO46MgnOM80zW1RWuH61DGLwZJEdK2Kadq2F9CUG65"" crossorigin=""anonymous"">"
response.write "<link href=""https://cdn.datatables.net/v/dt/dt-1.13.6/datatables.min.css"" rel=""stylesheet"">"
response.write "<link href=""https://cdn.datatables.net/v/bs5/jszip-3.10.1/dt-1.13.6/b-2.4.1/b-html5-2.4.1/b-print-2.4.1/datatables.min.css"" rel=""stylesheet"">"

AddCss
response.write "</head>"

BuildComboFilter
response.write "<a class=""btn btn-primary1"" href=""wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamStatRCP&PositionForTableName=WorkingDay&WorkingDayID="">Go Home</a>"

response.write "<a class=""btn btn-primary"" href=""wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamLabResultRCP&PositionForTableName=WorkingDay&dptID=" & dpt & "&secID=" & secid & "&genID=" & genID & "&branID=" & branid & "&yrID=" & yrid & " "">Refresh</a>"
response.write "        </td>"
response.write "</table>"
response.write "</div>"
response.write "                    <table id=""example"" class=""table table-striped table-bordered""  width=""100%"">"
response.write "                        <thead class=""beautiful-green-header"">"
response.write "<tr>"
response.write "<th>#</th>"
response.write "<th>LabRequestID</th>"
response.write "<th>LabTestID</th>"
response.write "<th>SponsorID</th>"
response.write "<th>RequestDate</th>"
response.write "<th>VisitationID</th>"
response.write "<th>PatientID</th>"
response.write "<th>GenderID</th>"
response.write "<th>PSA</th>"
response.write "</tr>"
response.write "</thead>"
response.write "<tbody>"

count = 0

With rst
    rst.open qryPro.FltQry(sql), conn, 3, 4

    If rst.RecordCount > 0 Then
        rst.MoveFirst

        Do While Not rst.EOF

            count = count + 1
   
 
  
            response.write "<tr>"
            response.write "<td>" & count & "</td>"
            response.write "<td>" & rst.fields("LabRequestID") & "</td>"
            response.write "<td>" & GetComboName("LabTest", .fields("LabTestID")) & "</td>"
            response.write "<td>" & GetComboName("Sponsor", .fields("SponsorID")) & "</td>"
            response.write "<td>" & FormatDate(rst.fields("RequestDate")) & "</td>"
            response.write "<td>" & rst.fields("visitationID") & "</td>"
            response.write "<td>" & rst.fields("patientID") & "</td>"
            response.write "<td>" & GetComboName("Gender", .fields("GenderID")) & "</td>"
            response.write "<td>" & rst.fields("PSA") & "</td>"
          
      
           
            response.write "</tr>"
            rst.MoveNext
            response.flush
        Loop
    End If
    rst.Close
End With
Set rst = Nothing

response.write "                        </tbody>"
response.write "                    </table>"
response.write "                </div>"
response.write "            </div>"
response.write "        </div>"
response.write "    </div>"




AddJS

response.write "</html>"



response.write "<script>"
response.write "function reload2() {"
response.write "    let selected = document.getElementById('cID').value;"
response.write "    let url = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamLabResultRCP';"
response.write "    url += '&PositionForTableName=WorkingDay&dptID=' + selected;"
response.write "    window.location.href = url;"
response.write "}"

response.write "</script>"
response.write "<script>"
response.write "function reload3() {"
response.write "    let selected = document.getElementById('cID').value;"
response.write "    let selected2 = document.getElementById('cID2').value;"
response.write "    let selected3 = document.getElementById('cID3').value;"
response.write "    let selected4 = document.getElementById('cID4').value;"
response.write "    let selected5 = document.getElementById('cID5').value;"
'response.write "    let sat = document.getElementById('statusFilter').value;"
response.write "    let url = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamLabResultRCP';"
response.write "    url += '&PositionForTableName=WorkingDay&dptID=' + selected + '&secID=' + selected2 + '&genID=' + selected3 + '&branID=' + selected4 + '&yrID=' + selected5;"
response.write "    window.location.href = url;"
response.write "}"



response.write "</script>"



End Sub
Sub MedicalExaminationLabResultsAllHB(testId)

Dim rst, sql

Set rst = CreateObject("ADODB.Recordset")


  sql = ""
sql = sql & " WITH LabReportRange AS ( "
sql = sql & " SELECT i.LabRequestID, i.LabTestID, i.RequestDate, i.VisitationID, i.PatientID, i.GenderID, i.SponsorID FROM Investigation i "
sql = sql & " LEFT JOIN Visitation v ON i.VisitationID = v.VisitationID "
    
    sql = sql & " Where i.LabTestID = '" & testId & "' "
        If Len(branid) > 0 Then
        sql = sql & " AND i.BranchID='" & branid & "' "
    End If
    If Len(genID) > 0 Then
        sql = sql & " AND i.GenderID='" & genID & "' "
    End If
        sql = sql & " AND i.SponsorID IN ('S571','S592','S593', 'S594','S657','VRA') "
        sql = sql & " AND v.SpecialistTypeID IN ('B001-0200','B002-0208','B003-0208','B005-0200','B006-0200') "
        If Len(yrid) > 0 Then
        sql = sql & "  AND i.WorkingYearID = '" & yrid & "' "
    End If
    sql = sql & " Union All "
    sql = sql & " SELECT i2.LabRequestID, i2.LabTestID, i2.RequestDate, i2.VisitationID, i2.PatientID, i2.GenderID, i2.SponsorID FROM Investigation2 i2 "
    sql = sql & " LEFT JOIN Visitation v ON i2.VisitationID = v.VisitationID "
    sql = sql & " Where i2.LabTestID = '" & dpt & "' "
    If Len(branid) > 0 Then
        sql = sql & " AND i2.BranchID='" & branid & "' "
    End If
    If Len(genID) > 0 Then
        sql = sql & " AND i2.GenderID='" & genID & "' "
    End If
        
        
        sql = sql & " AND i2.SponsorID IN ('S571','S592','S593', 'S594','S657','VRA') "
        sql = sql & " AND v.SpecialistTypeID IN ('B001-0200','B002-0208','B003-0208','B005-0200','B006-0200') "
          If Len(yrid) > 0 Then
           sql = sql & "  AND i2.WorkingYearID = '" & yrid & "' "
         End If
sql = sql & " ) "
sql = sql & " SELECT LabRequestID, LabTestID,SponsorID, RequestDate,VisitationID,PatientID, GenderID FROM LabReportRange ORDER BY RequestDate  "

response.write "<!DOCTYPE html>"
response.write "<html lang=""en"">"

response.write "<head> "
response.write "    <title>AME " & GetComboName("WorkingYear", yrid) & " Lab Tests Records</title>"
'response.write "    <title>AME " & currentYr & " Employee Records</title>"

response.write "<link href=""https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600&display=swap"" rel=""stylesheet"">"
response.write "<link href=""https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css"" rel=""stylesheet"">"
response.write "<link href=""https://cdn.datatables.net/buttons/2.4.1/css/buttons.dataTables.min.css"" rel=""stylesheet"">"
response.write "<link rel=""stylesheet"" href=""https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/css/bootstrap.min.css"" integrity=""sha384-rbsA2VBKQhggwzxH7pPCaAqO46MgnOM80zW1RWuH61DGLwZJEdK2Kadq2F9CUG65"" crossorigin=""anonymous"">"
response.write "<link href=""https://cdn.datatables.net/v/dt/dt-1.13.6/datatables.min.css"" rel=""stylesheet"">"
response.write "<link href=""https://cdn.datatables.net/v/bs5/jszip-3.10.1/dt-1.13.6/b-2.4.1/b-html5-2.4.1/b-print-2.4.1/datatables.min.css"" rel=""stylesheet"">"

AddCss
response.write "</head>"

BuildComboFilter
response.write "<a class=""btn btn-primary1"" href=""wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamLabResultRCP&PositionForTableName=WorkingDay&WorkingDayID=&WorkFlowNav=POP"">Go Home</a>"

response.write "<a class=""btn btn-primary"" href=""wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamHBReport&PositionForTableName=WorkingDay&dptID=" & dpt & "&secID=" & secid & "&genID=" & genID & "&branID=" & branid & "&yrID=" & yrid & " "">Refresh</a>"
response.write "        </td>"
response.write "</table>"
response.write "</div>"
response.write "                    <table id=""example"" class=""table table-striped table-bordered""  width=""100%"">"
response.write "                        <thead class=""beautiful-green-header"">"
response.write "<tr>"
response.write "<th>#</th>"
response.write "<th>LabRequestID</th>"
response.write "<th>LabTestID</th>"
response.write "<th>SponsorID</th>"
response.write "<th>RequestDate</th>"
response.write "<th>VisitationID</th>"
response.write "<th>PatientID</th>"
response.write "<th>GenderID</th>"
response.write "<th>Hepatitis_B_Surface_Antigen</th>"
response.write "</tr>"
response.write "</thead>"
response.write "<tbody>"

count = 0

With rst
    rst.open qryPro.FltQry(sql), conn, 3, 4

    If rst.RecordCount > 0 Then
        rst.MoveFirst

        Do While Not rst.EOF

            count = count + 1
   
              requestId = rst.fields("LabRequestID")
              testId = rst.fields("LabTestID")
              
              results = getLabResultsHB(requestId, testId)
        
  
            response.write "<tr>"
            response.write "<td>" & count & "</td>"
            response.write "<td>" & rst.fields("LabRequestID") & "</td>"
            response.write "<td>" & GetComboName("LabTest", .fields("LabTestID")) & "</td>"
            response.write "<td>" & GetComboName("Sponsor", .fields("SponsorID")) & "</td>"
            response.write "<td>" & FormatDate(rst.fields("RequestDate")) & "</td>"
            response.write "<td>" & rst.fields("visitationID") & "</td>"
            response.write "<td>" & rst.fields("patientID") & "</td>"
            response.write "<td>" & GetComboName("Gender", .fields("GenderID")) & "</td>"
            response.write "<td>" & results & "</td>"
          
      
           
            response.write "</tr>"
            rst.MoveNext
            response.flush
        Loop
    End If
    rst.Close
End With
Set rst = Nothing

response.write "                        </tbody>"
response.write "                    </table>"
response.write "                </div>"
response.write "            </div>"
response.write "        </div>"
response.write "    </div>"




AddJS

response.write "</html>"



response.write "<script>"
response.write "function reload2() {"
response.write "    let selected = document.getElementById('cID').value;"
response.write "    let url = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamHBReport';"
response.write "    url += '&PositionForTableName=WorkingDay&dptID=' + selected;"
response.write "    window.location.href = url;"
response.write "}"

response.write "</script>"
response.write "<script>"
response.write "function reload3() {"
response.write "    let selected = document.getElementById('cID').value;"
'response.write "    let selected2 = document.getElementById('cID2').value;"
response.write "    let selected3 = document.getElementById('cID3').value;"
response.write "    let selected4 = document.getElementById('cID4').value;"
response.write "    let selected5 = document.getElementById('cID5').value;"
'response.write "    let sat = document.getElementById('statusFilter').value;"
response.write "    let url = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamHBReport';"
response.write "    url += '&PositionForTableName=WorkingDay&dptID=' + selected +  '&genID=' + selected3 + '&branID=' + selected4 + '&yrID=' + selected5;"
response.write "    window.location.href = url;"
response.write "}"



response.write "</script>"
response.write "<script>"

response.write "        var dataTable = $('#example').DataTable({"
response.write "        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, ""All""]],"
response.write "        dom: 'lBfrtip',"
response.write "            pageLength: 10,"
response.write "            paging: true, "
response.write "            pagingType: 'simple', "
response.write "            search: {"
response.write "                smart: true"
response.write "                    },"
response.write "            buttons: ["
response.write "                {"
response.write "                    extend: 'csv',"
response.write "                    text: 'CSV',"
response.write "                    title: '" & dpt & " Labtest-" & FormatDate(Now) & "' "
response.write "                },"
response.write "                {"
response.write "                    extend: 'excel',"
response.write "                    text: 'EXCEL',"
response.write "                    title: '" & dpt & " Labtest-" & FormatDate(Now) & "'"
response.write "                },"
response.write "                {"
response.write "                    extend: 'pdf',"
response.write "                    text: 'PDF',"
response.write "                    title: '" & dpt & " Labtest-" & FormatDate(Now) & "'"
response.write "                },"

response.write "                {"
response.write "                    extend: 'print',"
response.write "                    text: 'PRINT',"
response.write "                    title: '" & dpt & " Labtest-" & FormatDate(Now) & "'"
response.write "                }"

response.write "            ],"

response.write "        columnDefs: ["
response.write "        {"
response.write "            target: 7,"
response.write "            visible: true,"
response.write "            searchable: true"
response.write "        }"
response.write "    ]"

response.write "        });"
response.write "</script>"



End Sub
Sub MedicalExaminationLabResultsAllMPS(testId)

Dim rst, sql

Set rst = CreateObject("ADODB.Recordset")


  sql = ""
sql = sql & " WITH LabReportRange AS ( "
sql = sql & " SELECT i.LabRequestID, i.LabTestID, i.RequestDate, i.VisitationID, i.PatientID, i.GenderID, i.SponsorID FROM Investigation i "
sql = sql & " LEFT JOIN Visitation v ON i.VisitationID = v.VisitationID "
    
    sql = sql & " Where i.LabTestID = '" & testId & "' "
        If Len(branid) > 0 Then
        sql = sql & " AND i.BranchID='" & branid & "' "
    End If
    If Len(genID) > 0 Then
        sql = sql & " AND i.GenderID='" & genID & "' "
    End If
        sql = sql & " AND i.SponsorID IN ('S571','S592','S593', 'S594','S657','VRA') "
        sql = sql & " AND v.SpecialistTypeID IN ('B001-0200','B002-0208','B003-0208') "
        If Len(yrid) > 0 Then
        sql = sql & "  AND i.WorkingYearID = '" & yrid & "' "
    End If
    sql = sql & " Union All "
    sql = sql & " SELECT i2.LabRequestID, i2.LabTestID, i2.RequestDate, i2.VisitationID, i2.PatientID, i2.GenderID, i2.SponsorID FROM Investigation2 i2 "
    sql = sql & " LEFT JOIN Visitation v ON i2.VisitationID = v.VisitationID "
    sql = sql & " Where i2.LabTestID = '" & dpt & "' "
    If Len(branid) > 0 Then
        sql = sql & " AND i2.BranchID='" & branid & "' "
    End If
    If Len(genID) > 0 Then
        sql = sql & " AND i2.GenderID='" & genID & "' "
    End If
        
        
        sql = sql & " AND i2.SponsorID IN ('S571','S592','S593', 'S594','S657','VRA') "
        sql = sql & " AND v.SpecialistTypeID IN ('B001-0200','B002-0208','B003-0208') "
          If Len(yrid) > 0 Then
           sql = sql & "  AND i2.WorkingYearID = '" & yrid & "' "
         End If
sql = sql & " ) "
sql = sql & " SELECT LabRequestID, LabTestID,SponsorID, RequestDate,VisitationID,PatientID, GenderID FROM LabReportRange ORDER BY RequestDate  "

response.write "<!DOCTYPE html>"
response.write "<html lang=""en"">"

response.write "<head> "
response.write "    <title>AME " & GetComboName("WorkingYear", yrid) & " Lab Tests Records</title>"
'response.write "    <title>AME " & currentYr & " Employee Records</title>"

response.write "<link href=""https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600&display=swap"" rel=""stylesheet"">"
response.write "<link href=""https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css"" rel=""stylesheet"">"
response.write "<link href=""https://cdn.datatables.net/buttons/2.4.1/css/buttons.dataTables.min.css"" rel=""stylesheet"">"
response.write "<link rel=""stylesheet"" href=""https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/css/bootstrap.min.css"" integrity=""sha384-rbsA2VBKQhggwzxH7pPCaAqO46MgnOM80zW1RWuH61DGLwZJEdK2Kadq2F9CUG65"" crossorigin=""anonymous"">"
response.write "<link href=""https://cdn.datatables.net/v/dt/dt-1.13.6/datatables.min.css"" rel=""stylesheet"">"
response.write "<link href=""https://cdn.datatables.net/v/bs5/jszip-3.10.1/dt-1.13.6/b-2.4.1/b-html5-2.4.1/b-print-2.4.1/datatables.min.css"" rel=""stylesheet"">"

AddCss
response.write "</head>"

BuildComboFilter
response.write "<a class=""btn btn-primary1"" href=""wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamStatRCP&PositionForTableName=WorkingDay&WorkingDayID="">Go Home</a>"

response.write "<a class=""btn btn-primary"" href=""wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamLabResultRCP&PositionForTableName=WorkingDay&dptID=" & dpt & "&secID=" & secid & "&genID=" & genID & "&branID=" & branid & "&yrID=" & yrid & " "">Refresh</a>"
response.write "        </td>"
response.write "</table>"
response.write "</div>"
response.write "                    <table id=""example"" class=""table table-striped table-bordered""  width=""100%"">"
response.write "                        <thead class=""beautiful-green-header"">"
response.write "<tr>"
response.write "<th>#</th>"
response.write "<th>LabRequestID</th>"
response.write "<th>LabTestID</th>"
response.write "<th>SponsorID</th>"
response.write "<th>RequestDate</th>"
response.write "<th>VisitationID</th>"
response.write "<th>PatientID</th>"
response.write "<th>GenderID</th>"
response.write "<th>BLOOD_FILM_FOR MPS</th>"
response.write "</tr>"
response.write "</thead>"
response.write "<tbody>"

count = 0

With rst
    rst.open qryPro.FltQry(sql), conn, 3, 4

    If rst.RecordCount > 0 Then
        rst.MoveFirst

        Do While Not rst.EOF

            count = count + 1
   
              requestId = rst.fields("LabRequestID")
              testId = rst.fields("LabTestID")
              
              results = getLabResultsMPS(requestId, testId)
        
  
            response.write "<tr>"
            response.write "<td>" & count & "</td>"
            response.write "<td>" & rst.fields("LabRequestID") & "</td>"
            response.write "<td>" & GetComboName("LabTest", .fields("LabTestID")) & "</td>"
            response.write "<td>" & GetComboName("Sponsor", .fields("SponsorID")) & "</td>"
            response.write "<td>" & FormatDate(rst.fields("RequestDate")) & "</td>"
            response.write "<td>" & rst.fields("visitationID") & "</td>"
            response.write "<td>" & rst.fields("patientID") & "</td>"
            response.write "<td>" & GetComboName("Gender", .fields("GenderID")) & "</td>"
            response.write "<td>" & results & "</td>"
          
      
           
            response.write "</tr>"
            rst.MoveNext
            response.flush
        Loop
    End If
    rst.Close
End With
Set rst = Nothing

response.write "                        </tbody>"
response.write "                    </table>"
response.write "                </div>"
response.write "            </div>"
response.write "        </div>"
response.write "    </div>"




AddJS

response.write "</html>"



response.write "<script>"
response.write "function reload2() {"
response.write "    let selected = document.getElementById('cID').value;"
response.write "    let url = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamLabResultRCP';"
response.write "    url += '&PositionForTableName=WorkingDay&dptID=' + selected;"
response.write "    window.location.href = url;"
response.write "}"

response.write "</script>"
response.write "<script>"
response.write "function reload3() {"
response.write "    let selected = document.getElementById('cID').value;"
response.write "    let selected2 = document.getElementById('cID2').value;"
response.write "    let selected3 = document.getElementById('cID3').value;"
response.write "    let selected4 = document.getElementById('cID4').value;"
response.write "    let selected5 = document.getElementById('cID5').value;"
'response.write "    let sat = document.getElementById('statusFilter').value;"
response.write "    let url = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamLabResultRCP';"
response.write "    url += '&PositionForTableName=WorkingDay&dptID=' + selected + '&secID=' + selected2 + '&genID=' + selected3 + '&branID=' + selected4 + '&yrID=' + selected5;"
response.write "    window.location.href = url;"
response.write "}"



response.write "</script>"



End Sub
Sub MedicalExaminationLabResultsOne(testId, tcompid)

Dim rst, sql

Set rst = CreateObject("ADODB.Recordset")


 sql = ""
sql = sql & " WITH LabReportRange AS ("
    sql = sql & " SELECT"
        sql = sql & " i.LabRequestID,"
        sql = sql & " i.LabTestID,"
        sql = sql & " i.RequestDate,"
        sql = sql & " i.VisitationID,"
        sql = sql & " i.PatientID,"
        sql = sql & " i.GenderID,"
        sql = sql & " i.SponsorID,"
        sql = sql & " [L0460].TestComponentID,"
        sql = sql & " [L0460].Column1 AS FLAG,"
        sql = sql & " [L0460].Column2 AS WBC,"
        sql = sql & " [L0460].Column3 AS UNITS"
    sql = sql & " FROM Investigation i"
    sql = sql & " LEFT JOIN Visitation v ON i.VisitationID = v.VisitationID"
     If Len(tcompid) > 0 Then
        sql = sql & " LEFT JOIN LabResults AS [L0460] ON i.LabRequestID = [L0460].LabRequestID"
        sql = sql & " AND [L0460].TestComponentID = '" & tcompid & "' "
    End If
    sql = sql & " Where i.LabTestID = '" & testId & "' "
        
         If Len(branid) > 0 Then
        sql = sql & " AND i.BranchID='" & branid & "' "
        End If
        If Len(genID) > 0 Then
        sql = sql & " AND i.GenderID='" & genID & "' "
        End If
        sql = sql & " AND i.SponsorID IN ('S571','S592','S593', 'S594','S657','VRA')"
        sql = sql & " AND v.SpecialistTypeID IN ('B001-0200','B002-0208','B003-0208')"
        If Len(yrid) > 0 Then
        sql = sql & "  AND i.WorkingYearID = '" & yrid & "' "
    End If

    sql = sql & " Union All"
    sql = sql & " SELECT"
        sql = sql & " i2.LabRequestID,"
        sql = sql & " i2.LabTestID,"
        sql = sql & " i2.RequestDate,"
        sql = sql & " i2.VisitationID,"
        sql = sql & " i2.PatientID,"
        sql = sql & " i2.GenderID,"
        sql = sql & " i2.SponsorID,"
        sql = sql & " [L0460].TestComponentID,"
        sql = sql & " [L0460].Column1 AS FLAG,"
        sql = sql & " [L0460].Column2 AS WBC,"
        sql = sql & " [L0460].Column3 AS UNITS"
    sql = sql & " FROM Investigation2 i2"
    sql = sql & " LEFT JOIN Visitation v ON i2.VisitationID = v.VisitationID"
    If Len(tcompid) > 0 Then
        sql = sql & " LEFT JOIN LabResults AS [L0460] ON i2.LabRequestID = [L0460].LabRequestID"
        sql = sql & " AND [L0460].TestComponentID = '" & tcompid & "' "
    End If
    
    sql = sql & " Where i2.LabTestID = '" & testId & "' "
        
        If Len(branid) > 0 Then
        sql = sql & " AND i2.BranchID='" & branid & "' "
        End If
        If Len(genID) > 0 Then
        sql = sql & " AND i2.GenderID='" & genID & "' "
        End If
        sql = sql & " AND i2.SponsorID IN ('S571','S592','S593', 'S594','S657','VRA')"
        sql = sql & " AND v.SpecialistTypeID IN ('B001-0200','B002-0208','B003-0208')"
        If Len(yrid) > 0 Then
        sql = sql & "  AND i2.WorkingYearID = '" & yrid & "' "
    End If
sql = sql & " )"
sql = sql & " SELECT"
    sql = sql & " LabRequestID,"
    sql = sql & " LabTestID,"
    sql = sql & " SponsorID,"
    sql = sql & " RequestDate,"
    sql = sql & " VisitationID,"
    sql = sql & " PatientID,"
    sql = sql & " GenderID,"
    sql = sql & " TestComponentID,"
    sql = sql & " FLAG,"
    sql = sql & " WBC,"
    sql = sql & " UNITS"
sql = sql & " FROM LabReportRange;"


response.write "<!DOCTYPE html>"
response.write "<html lang=""en"">"

response.write "<head> "
response.write "    <title>AME " & GetComboName("WorkingYear", yrid) & " Lab Tests Records</title>"
'response.write "    <title>AME " & currentYr & " Employee Records</title>"

response.write "<link href=""https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600&display=swap"" rel=""stylesheet"">"
response.write "<link href=""https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css"" rel=""stylesheet"">"
response.write "<link href=""https://cdn.datatables.net/buttons/2.4.1/css/buttons.dataTables.min.css"" rel=""stylesheet"">"
response.write "<link rel=""stylesheet"" href=""https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/css/bootstrap.min.css"" integrity=""sha384-rbsA2VBKQhggwzxH7pPCaAqO46MgnOM80zW1RWuH61DGLwZJEdK2Kadq2F9CUG65"" crossorigin=""anonymous"">"
response.write "<link href=""https://cdn.datatables.net/v/dt/dt-1.13.6/datatables.min.css"" rel=""stylesheet"">"
response.write "<link href=""https://cdn.datatables.net/v/bs5/jszip-3.10.1/dt-1.13.6/b-2.4.1/b-html5-2.4.1/b-print-2.4.1/datatables.min.css"" rel=""stylesheet"">"

AddCss
response.write "</head>"

BuildComboFilter
response.write "<a class=""btn btn-primary1"" href=""wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamStatRCP&PositionForTableName=WorkingDay&WorkingDayID="">Go Home</a>"

response.write "<a class=""btn btn-primary"" href=""wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamLabResultRCP&PositionForTableName=WorkingDay&dptID=" & dpt & "&secID=" & secid & "&genID=" & genID & "&branID=" & branid & "&yrID=" & yrid & " "">Refresh</a>"
response.write "        </td>"
response.write "</table>"
response.write "</div>"
response.write "                    <table id=""TestComp"" class=""table table-striped table-bordered""  width=""100%"">"
response.write "                        <thead class=""beautiful-green-header"">"
response.write "<tr>"
response.write "<th>#</th>"
response.write "<th>LabRequestID</th>"
response.write "<th>LabTestID</th>"
response.write "<th>RequestDate</th>"
response.write "<th>VisitationID</th>"
response.write "<th>PatientID</th>"
response.write "<th>GenderID</th>"
response.write "<th>Parameter</th>"
response.write "<th>Results</th>"
response.write "<th>Flag</th>"
response.write "<th>Unit</th>"
response.write "</tr>"
response.write "</thead>"
response.write "<tbody>"

count = 0

With rst
    rst.open qryPro.FltQry(sql), conn, 3, 4

    If rst.RecordCount > 0 Then
        rst.MoveFirst

        Do While Not rst.EOF

            count = count + 1
   
 
  
            response.write "<tr>"
            response.write "<td>" & count & "</td>"
            response.write "<td>" & rst.fields("LabRequestID") & "</td>"
            response.write "<td>" & GetComboName("LabTest", .fields("LabTestID")) & "</td>"
            response.write "<td>" & FormatDate(rst.fields("RequestDate")) & "</td>"
            response.write "<td>" & rst.fields("visitationID") & "</td>"
            response.write "<td>" & rst.fields("patientID") & "</td>"
            response.write "<td>" & GetComboName("Gender", .fields("GenderID")) & "</td>"
            response.write "<td>" & GetComboName("TestComponent", .fields("TestComponentID")) & "</td>"
            response.write "<td>" & rst.fields("WBC") & "</td>"
            response.write "<td>" & rst.fields("FLAG") & "</td>"
            response.write "<td>" & rst.fields("UNITS") & "</td>"
            response.write "</tr>"
            rst.MoveNext
            response.flush
        Loop
    End If
    rst.Close
End With
Set rst = Nothing

response.write "                        </tbody>"
response.write "                    </table>"
response.write "                </div>"
response.write "            </div>"
response.write "        </div>"
response.write "    </div>"




AddJS

response.write "</html>"



response.write "<script>"
response.write "function reload2() {"
response.write "    let selected = document.getElementById('cID').value;"
response.write "    let url = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamLabResultRCP';"
response.write "    url += '&PositionForTableName=WorkingDay&dptID=' + selected;"
response.write "    window.location.href = url;"
response.write "}"

response.write "</script>"
response.write "<script>"
response.write "function reload3() {"
response.write "    let selected = document.getElementById('cID').value;"
response.write "    let selected2 = document.getElementById('cID2').value;"
response.write "    let selected3 = document.getElementById('cID3').value;"
response.write "    let selected4 = document.getElementById('cID4').value;"
response.write "    let selected5 = document.getElementById('cID5').value;"
'response.write "    let sat = document.getElementById('statusFilter').value;"
response.write "    let url = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=MedicalExamLabResultRCP';"
response.write "    url += '&PositionForTableName=WorkingDay&dptID=' + selected + '&secID=' + selected2 + '&genID=' + selected3 + '&branID=' + selected4 + '&yrID=' + selected5;"
response.write "    window.location.href = url;"
response.write "}"



response.write "</script>"



End Sub



Sub BuildComboFilter()


response.write "<header class=""header_part"">"
response.write "<h4>AME " & GetComboName("WorkingYear", yrid) & " LAB TEST RECORDS &gt;&gt;&gt;HEPATITIS B SURFACE ANTIGEN SCREEN [HBsAg]</h4>"

response.write "        <p id=""date-time"" style=""text-align: right;color:#ffffff;""></p>"
response.write "    </header>"
response.write "    <div class=""container"">"
response.write "        <div class=""row"">"
response.write "            <div class=""col-12"">"

'DEPARTMENT COMBO STARTS

Dim rst6, rst7, rst8, rst9, rst10, sql6, sql7, sql8, sql9, sql10

         Set rst6 = CreateObject("ADODB.Recordset")
         Set rst7 = CreateObject("ADODB.Recordset")
         Set rst8 = CreateObject("ADODB.Recordset")
         Set rst9 = CreateObject("ADODB.Recordset")
         Set rst10 = CreateObject("ADODB.Recordset")
sql6 = ""
sql6 = sql6 & " SELECT LabTestID, LabTestName FROM LabTest WHERE LabTestID = 'LL3021'  "

sql7 = ""
sql7 = sql7 & " SELECT DISTINCT(TestComponentID) FROM LabResults WHERE LabTestID='" & dpt & "'  "

'sql7 = ""
'sql7 = sql7 & " SELECT DISTINCT(SponsorID) FROM Visitation WHERE SponsorID IN ('S571','S592','S593', 'S594','S657','VRA')  "

sql8 = ""
sql8 = sql8 & " SELECT GenderID FROM Gender "

sql9 = ""
sql9 = sql9 & " SELECT BranchID FROM Branch WHERE BranchID IN ('B001','B002','B003','B004','B005','B006') "

sql10 = ""
'sql10 = sql10 & " SELECT WorkingYearID FROM WorkingYear "
sql10 = sql10 & " SELECT DISTINCT(WorkingYearID) FROM Visitation where WorkingYearID >= 'YRS2022' "




        
response.write "<div>"
response.write "<p><b>Showing Results for&gt;&gt;&gt;</b><b>Lab Test:</b> " & GetComboName("LabTest", dpt) & ", <b>Parameter:</b> " & GetComboName("TestComponent", secid) & ", <b>Gender:</b> " & GetComboName("Gender", genID) & ", <b>Facility:</b> " & GetComboName("Branch", branid) & ", <b>Year:</b> " & GetComboName("WorkingYear", yrid) & "</p>"
response.write "</div>"

response.write "<div id='buildCombos'>"
response.write "<table>"
response.write "        <tr>"
response.write "            <td class=""department-label""></td>"
response.write "        <td>"



'DEPARTMENT COMBO STARTS

response.write "<select  id='cID'>"
response.write "    <option value="""">Select Labtest</option>"
With rst6
    .open qryPro.FltQry(sql6), conn, 3, 4
    Do While Not .EOF
        If UCase(.fields("LabTestID")) = UCase(dpt) Then
            response.write "    <option value=""" & .fields("LabTestID") & """ Selected>" & GetComboName("LabTest", .fields("LabTestID")) & "</option>"
        Else
           response.write "    <option value=""" & .fields("LabTestID") & """ >" & GetComboName("LabTest", .fields("LabTestID")) & "</option>"
        End If
       
        .MoveNext
    Loop
End With
rst6.Close
response.write "</select>"

'Section COMBO STARTS

'response.write "<select  id='cID2'>"
'response.write "    <option value="""">Select Parameter</option>"
'With rst7
'    .open qryPro.FltQry(sql7), conn, 3, 4
'    Do While Not .EOF
'        If UCase(.fields("TestComponentID")) = UCase(secid) Then
'            response.write "    <option value=""" & .fields("TestComponentID") & """ Selected>" & GetComboName("TestComponent", .fields("TestComponentID")) & "</option>"
'        Else
'           response.write "    <option value=""" & .fields("TestComponentID") & """ >" & GetComboName("TestComponent", .fields("TestComponentID")) & "</option>"
'        End If
'
'        .MoveNext
'    Loop
'End With
'rst7.Close
'response.write "</select>"


'Gender COMBO STARTS

response.write "<select  id='cID3'>"
response.write "    <option value="""">Select Gender</option>"
With rst8
    .open qryPro.FltQry(sql8), conn, 3, 4
    Do While Not .EOF
        If UCase(.fields("GenderID")) = UCase(genID) Then
            response.write "    <option value=""" & .fields("GenderID") & """ Selected>" & GetComboName("Gender", .fields("GenderID")) & "</option>"
        Else
           response.write "    <option value=""" & .fields("GenderID") & """ >" & GetComboName("Gender", .fields("GenderID")) & "</option>"
        End If
       
        .MoveNext
    Loop
End With
rst8.Close
response.write "</select>"


'Facility COMBO STARTS
response.write "<select  id='cID4'>"
response.write "    <option value="""">Select Facility</option>"
With rst9
    .open qryPro.FltQry(sql9), conn, 3, 4
    Do While Not .EOF
        If UCase(.fields("BranchID")) = UCase(branid) Then
            response.write "    <option value=""" & .fields("BranchID") & """ Selected>" & GetComboName("Branch", .fields("BranchID")) & "</option>"
        Else
           response.write "    <option value=""" & .fields("BranchID") & """ >" & GetComboName("Branch", .fields("BranchID")) & "</option>"
        End If
       
        .MoveNext
    Loop
End With
rst9.Close
response.write "</select>"

'Year COMBO STARTS
response.write "<select  id='cID5'>"
response.write "    <option value="""">Select Year</option>"
With rst10
    .open qryPro.FltQry(sql10), conn, 3, 4
    Do While Not .EOF
        If UCase(.fields("WorkingYearID")) = UCase(yrid) Then
            response.write "    <option value=""" & .fields("WorkingYearID") & """ Selected>" & GetComboName("WorkingYear", .fields("WorkingYearID")) & "</option>"
        Else
           response.write "    <option value=""" & .fields("WorkingYearID") & """ >" & GetComboName("WorkingYear", .fields("WorkingYearID")) & "</option>"
        End If
       
        .MoveNext
    Loop
End With
rst10.Close
response.write "</select>"



response.write "<a class=""btn btn-primary"" onClick=""reload3()"">Apply Filter(s)</a>"
response.write "        </td>"
'DEPARTMENT COMBO ENDS

End Sub






Sub AddCss()

response.write "<style>"
response.write "* {"
response.write "    margin: 0;"
response.write "    padding: 0;"
response.write "    outline: 0;"
response.write "}"
response.write ""
response.write "html {"
response.write "    scroll-behavior: smooth;"
response.write "}"
response.write "table {"
response.write "    width:100%;"
response.write "}"
response.write ""
response.write "html ::-webkit-scrollbar {"
response.write "    width: 10px;"
response.write "    background-color: #111d2c;"
response.write "    cursor: pointer;"
response.write "}"
response.write ""
response.write "html ::-webkit-scrollbar-thumb {"
response.write "    background: linear-gradient(160deg, #3c76ba, #1e428b);"
response.write "    width: 10px;"
response.write "    border-radius: 5px;"
response.write "}"
response.write ""
response.write "html ::-webkit-scrollbar-thumb:hover {"
response.write "    background: #379dc3;"
response.write "    cursor: pointer;"
response.write "}"
response.write ""
response.write "body {"
response.write "    background: #eee;"
response.write "    font-family: 'Raleway', sans-serif;"
response.write "    font-family: 'Roboto', sans-serif;"
response.write "    font-family: 'Poppins', sans-serif;"
response.write "    font-family: 'Rubik', sans-serif;"
response.write "    padding: 100px 0;"
response.write "}"
response.write ""
response.write "h1,"
response.write "h2,"
response.write "h3,"
response.write "h4,"
response.write "h5,"
response.write "h6,"
response.write "p,"
response.write "a,"
response.write "ol,"
response.write "ul,"
response.write "li,"
response.write "img,"
response.write "span,"
response.write "input,"
response.write "form,"
response.write "label,"
response.write ".btn,"
response.write ".navbar,"
response.write ".navbar-brand {"
response.write "    margin: 0;"
response.write "    padding: 0;"
response.write "}"
response.write ""
response.write "h1::selection,"
response.write "h2::selection,"
response.write "h3::selection,"
response.write "h4::selection,"
response.write "h5::selection,"
response.write "h6::selection,"
response.write "p::selection,"
response.write "a::selection,"
response.write "ol::selection,"
response.write "ul::selection,"
response.write "li::selection,"
response.write "img::selection,"
response.write "span::selection,"
response.write "input::selection,"
response.write "form::selection,"
response.write "label::selection,"
response.write ".btn::selection,"
response.write ".navbar::selection,"
response.write ".navbar-brand::selection {"
response.write "    color: #fff;"
response.write "    background: #187fcc;"
response.write "    text-shadow: 1px 1px 1px #333;"
response.write "}"
response.write ""
response.write "ol,"
response.write "ul,"
response.write "li {"
response.write "    list-style-type: none;"
response.write "}"
response.write ""
response.write "a,"
response.write "a:hover {"
response.write "    text-decoration: none;"
response.write "    color: #333;"
response.write "    -webkit-transition: 0.2s;"
response.write "    transition: 0.2s;"
response.write "}"
response.write ""
response.write "h1 {"
response.write "    font-size: 50px;"
response.write "    color: #111;"
response.write "    font-weight: 600;"
response.write "}"
response.write ""
response.write "h2 {"
response.write "    font-size: 40px;"
response.write "    color: #111;"
response.write "    font-weight: 600;"
response.write "}"
response.write ""
response.write "h3 {"
response.write "    font-size: 35px;"
response.write "    color: #111;"
response.write "    font-weight: 600;"
response.write "}"
response.write ""
response.write "h4 {"
response.write "    font-size: 30px;"
response.write "    color: #111;"
response.write "    font-weight: 600;"
response.write "}"
response.write ""
response.write "h5 {"
response.write "    font-size: 25px;"
response.write "    color: #111;"
response.write "    font-weight: 600;"
response.write "}"
response.write ""
response.write "h6 {"
response.write "    font-size: 18px;"
response.write "    color: #111;"
response.write "}"
response.write ""
response.write "p {"
response.write "    font-size: 14px;"
response.write "    color: #111;"
response.write "}"
response.write ""
response.write ".ptb-50 {"
response.write "    padding: 50px 0px;"
response.write "}"
response.write ""
response.write ".mtb-15 {"
response.write "    margin: 15px 0;"
response.write "}"
response.write ""
response.write ""
response.write "/* ==== Scoll  = Start ===== */"
response.write ".totop {"
response.write "    position: fixed;"
response.write "    bottom: 0px;"
response.write "    right: 10px;"
response.write "    height: 45px;"
response.write "    width: 45px;"
response.write "    cursor: pointer;"
response.write "    display: none;"
response.write "    background: rgba(0, 0, 0, 0.3);"
response.write "    color: #fff;"
response.write "    border-radius: 5px 5px 0 0;"
response.write "    line-height: 47px;"
response.write "    font-size: 20px;"
response.write "    text-align: center;"
response.write "}"
response.write ""
response.write ".totop:hover {"
response.write "    line-height: 40px;"
response.write "    color: #fff;"
response.write "    background: #194274;"
response.write "    text-shadow: 0 0 5px #000;"
response.write "}"
response.write ""
response.write "/* ==== Scoll  = End ===== */"
response.write ""
response.write ""
response.write ""
response.write ""
response.write "/* ==== Header  = Start ===== */"
'Response.Write ".header_part {"
'Response.Write "    background: #2ecc71;"
'Response.Write "    padding: 5px 80px;"
'Response.Write "    position: fixed;"
'Response.Write "    top: 0;"
'Response.Write "    width: 100%;"
'Response.Write "    text-align: center;"
'Response.Write "    z-index: 9999;"
'Response.Write "}"
'#ADD8E6


response.write "    .header_part {"
response.write "        background: #663396;"
response.write "        padding: 5px 80px;"
response.write "        position: fixed;"
response.write "        top: 0;"
response.write "        width: 100%;"
response.write "        text-align: center;"
response.write "        z-index: 9999;"
response.write "        box-shadow: 0px 2px 10px rgba(0, 0, 0, 0.2); "
response.write "        transition: transform 0.3s ease; "
response.write "    }"
response.write ""
response.write "    .header_part:hover {"
response.write "        transform: translateY(-5px); "
response.write "    }"

response.write ""
response.write ".header_part h4 {"
response.write "    color: #fff;"
response.write "    padding: 10px;"
response.write "}"
response.write ""
response.write ".header_part img {"
response.write "    float: left;"
response.write "    height: 60px;"
response.write "    border-radius: 50%;"
response.write "}"
response.write ""
response.write "/* ==== Header  = End ===== */"
response.write ""
response.write ""
response.write ""
response.write ""
response.write ""
response.write ""
response.write "/* ==== Data table  = Start ===== */"
response.write ""
response.write ".data_table{"
response.write "   background: #fff;"
response.write "    padding: 15px;"
response.write "    box-shadow: 1px 3px 5px #aaa;"
response.write "    border-radius: 5px;"
response.write "}"
response.write ""
response.write ".data_table .btn{"
response.write "    padding: 5px 10px;"
response.write "    margin: 10px 3px 10px 0;"
response.write "}"


response.write "    #previous, #next {"
response.write "        float: right;"
response.write "    }"

response.write "        #filterButton, #resetButton, #previous, #next {"
response.write "            border-radius: 5px; "
response.write "            background-color: #007bff; "
response.write "            color: #ffffff; "
response.write "            padding: 8px 12px; "
response.write "            margin-right: 10px; "
response.write "            border: none;"
response.write "            cursor: pointer;"
response.write "        }"
response.write "        #filterButton:hover, #resetButton:hover, #previous:hover, #next:hover {"
response.write "            background-color: #0056b3;"
response.write "        }"

response.write ".styled-input {"
response.write "    border: 1px solid #ccc;"
response.write "    border-radius: 8px;"
response.write "    padding: 8px;"
response.write "    margin-right: 10px;"
response.write "    font-size: 14px;"
response.write "    outline: none;"
response.write "}"
response.write ""
response.write ".styled-button {"
response.write "    background-color: #007bff;"
response.write "    color: white;"
response.write "    border: none;"
response.write "    border-radius: 8px;"
response.write "    padding: 8px 16px;"
response.write "    cursor: pointer;"
response.write "    font-size: 14px;"
response.write "}"
response.write ""
response.write ".styled-button:hover {"
response.write "    background-color: #0056b3;"
response.write "}"

response.write "    .beautiful-green-table tbody tr:nth-child(odd) {"
response.write "        background-color: #e0f2e9; "
response.write "    }"
response.write "    .beautiful-green-header {"
response.write "        background-color: #663396; "
response.write "        color: white; "
response.write "    }"

response.write "  .department-label {"
response.write "    color: black;"
response.write "    font-weight: bold;"
response.write "    text-align: center;"
response.write "  }"
response.write "  .btn-Primary {"
response.write "    color: white;"
response.write "        background: #0DCAF0;"
response.write "  }"
response.write "  .btn-Primary1 {"
response.write "    color: white;"
response.write "        background: #663396;"
response.write "  }"
response.write "  .btn-Primary1:hover {"
response.write "    color: white;"
response.write "        background: #663396;"
response.write "  }"
response.write "  #cID {"
response.write "    width: 200px; "
response.write "    padding: 5px;"
response.write "    border: 1px solid #ccc;"
response.write "    border-radius: 4px;"
response.write "    font-size: 16px;"
response.write "    margin-left: 10%; "
response.write "  }"
response.write "  #cID2 {"
response.write "    width: 200px; "
response.write "    padding: 5px;"
response.write "    border: 1px solid #ccc;"
response.write "    border-radius: 4px;"
response.write "    font-size: 16px;"
response.write "    margin-left: 1%; "
response.write "  }"
response.write "  #cID3 {"
response.write "    width: 200px; "
response.write "    padding: 5px;"
response.write "    border: 1px solid #ccc;"
response.write "    border-radius: 4px;"
response.write "    font-size: 16px;"
response.write "    margin-left: 1%; "
response.write "  }"
response.write "  #cID4 {"
response.write "    width: 200px; "
response.write "    padding: 5px;"
response.write "    border: 1px solid #ccc;"
response.write "    border-radius: 4px;"
response.write "    font-size: 16px;"
response.write "    margin-left: 1%; "
response.write "  }"
response.write "  #cID5 {"
response.write "    width: 200px; "
response.write "    padding: 5px;"
response.write "    border: 1px solid #ccc;"
response.write "    border-radius: 4px;"
response.write "    font-size: 16px;"
response.write "    margin-left: 1%; "
response.write "  }"
response.write "  #cID option {"
response.write "    font-size: 16px;"
response.write "    background-color: #fff; "
response.write "  }"
response.write "  #buildCombos {"
response.write "    margin-top: 1%; "
response.write "    padding-bottom: 20px; "
response.write "    background-color: #6C757D; "
response.write "  }"

response.write "  #example {"
response.write "    margin-top: 20%; "
response.write "    padding-top: 20px; "
response.write "  }"

response.write "</style>"
 
'#2ecc71
End Sub
Sub AddJS()


response.write "<script src=""https://code.jquery.com/jquery-3.6.0.min.js""></script>"
response.write "<script src=""https://cdn.datatables.net/buttons/2.4.1/js/dataTables.buttons.min.js""></script>"
response.write "<script src=""https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js""></script>"

response.write "<script src=""https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/js/bootstrap.bundle.min.js"" integrity=""sha384-kenU1KFdBIe4zVF0s0G1M5b4hcpxyD9F7jL+jjXkk+Q2h455rYXK/7HAuoJl+0I4"" crossorigin=""anonymous""></script>"
response.write "<script src=""https://cdn.datatables.net/v/dt/dt-1.13.6/datatables.min.js""></script>"
response.write "<script src=""https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/pdfmake.min.js""></script>"
response.write "<script src=""https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/vfs_fonts.js""></script>"
response.write "<script src=""https://cdn.datatables.net/v/bs5/jszip-3.10.1/dt-1.13.6/b-2.4.1/b-html5-2.4.1/b-print-2.4.1/datatables.min.js""></script>"
response.write "<script src=""https://cdn.datatables.net/buttons/2.4.1/js/buttons.html5.min.js""></script>"
response.write "<script src=""https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.1/xlsx.full.min.js""></script>"






End Sub

Function getLabResultsMPS(requestId, testId)
    Dim rst2, sql2, testResult

    Set rst2 = CreateObject("ADODB.Recordset")

    sql2 = " SELECT Column1 FROM LabResults lr "
    sql2 = sql2 & " WHERE LabRequestID = '" & requestId & "' "
    sql2 = sql2 & " AND LabTestID = '" & testId & "' "
    
        
        '           BLOOD FILM FOR MPS
        If testId = "LL4002" Then
        sql2 = sql2 & " AND TestComponentID = 'L0750' "
    
    
        End If
    
  

    With rst2
        .open qryPro.FltQry(sql2), conn, 3, 4

        If .RecordCount > 0 Then
            .MoveFirst

            Do While Not .EOF
                testResult = .fields("Column1")

                    
'                          BLOOD FILM FOR MPS
                If testId = "LL4002" Then
                    getLabResultsMPS = GetComboName("TestVar6", testResult)
                End If

                .MoveNext
            Loop
        End If

        .Close
    End With

    Set rst2 = Nothing
End Function

Function getLabResultsHB(requestId, testId)
    Dim rst2, sql2, testResult

    Set rst2 = CreateObject("ADODB.Recordset")

    sql2 = " SELECT Column1 FROM LabResults lr "
    sql2 = sql2 & " WHERE LabRequestID = '" & requestId & "' "
    sql2 = sql2 & " AND LabTestID = '" & testId & "' "
    
    If testId = "LL3021" Then
        sql2 = sql2 & " AND TestComponentID = 'L0500' "
    
    End If
    
  

    With rst2
        .open qryPro.FltQry(sql2), conn, 3, 4

        If .RecordCount > 0 Then
            .MoveFirst

            Do While Not .EOF
                testResult = .fields("Column1")

                If testId = "LL3021" Then
                    getLabResultsHB = GetComboName("T_NegPos", testResult)
                End If

                .MoveNext
            Loop
        End If

        .Close
    End With

    Set rst2 = Nothing
End Function

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
