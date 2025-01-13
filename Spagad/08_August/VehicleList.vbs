'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim requested, approved, available, issued, vehicleNo
SetPageVariable "AutoHidePrintControl", "Yes"

'calling alert
response.write "<meta http-equiv=""refresh"" content=""600"">"



If (jschd = "M10A") Then
strHd = "display:block"
Else
vehicleBooking
strHd = "display:none"
End If

response.write "<!DOCTYPE html>"
response.write "<html lang=""en"">"

response.write "<head> "
response.write "  <title>Vehicle Management</title>"
 response.write "<link href=""https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600&display=swap"" rel=""stylesheet"">"
 response.write "<link rel=""stylesheet"" href=""https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/css/bootstrap.min.css"" integrity=""sha384-rbsA2VBKQhggwzxH7pPCaAqO46MgnOM80zW1RWuH61DGLwZJEdK2Kadq2F9CUG65"" crossorigin=""anonymous"">"
 response.write "<link href=""https://cdn.datatables.net/v/dt/dt-1.13.6/datatables.min.css"" rel=""stylesheet"">"
 response.write "<link href=""https://cdn.datatables.net/v/bs5/jszip-3.10.1/dt-1.13.6/b-2.4.1/b-html5-2.4.1/b-print-2.4.1/datatables.min.css"" rel=""stylesheet"">"
  
  response.write Glob_GetBootstrap5()
  response.write Glob_GetIconFontAwesome()
  addCSS
 response.write "</head>"
  InitPageScript
  If (jschd = "M10A") Then

response.write "<div>"
response.write "<nav class=""navbar navbar-expand-lg navbar-dark"">"
response.write "  <button class=""navbar-toggler "" type=""button"" data-toggle=""collapse"" data-target=""#navbarNav"" aria-controls=""navbarNav"" aria-expanded=""false"" aria-label=""Toggle navigation"" >"
response.write "    <span class=""navbar-toggler-icon""></span>"
response.write "  </button>"
response.write "  <div class=""collapse navbar-collapse"" id=""navbarNav"">"
response.write "    <ul class=""navbar-nav"">"
response.write "      <li class=""nav-item"" >"
response.write "        <a class=""nav-link btn text-dark ml-2"" href=""#"" data-content=""home"" style = 'border:1px solid black; " & strHd & " '>Vehicle Management</a>"
response.write "      </li>"
response.write "      <li class=""nav-item"">"
response.write "        <a class=""nav-link btn  mr-2 text-dark"" href=""#"" data-content=""services"" style = 'border:1px solid black;margin-left:5px '>Vehicle Booking</a>"
response.write "      </li>"
response.write "      <li class=""nav-item"">"
response.write "        <a class=""nav-link btn  mr-2 text-dark"" href=""#"" data-content=""maintain"" style = 'border:1px solid black; " & strHd & "'>Maintenace Management</a>"
response.write "      </li>"
response.write "      <li class=""nav-item"">"
response.write "        <a class=""nav-link btn  mr-2 text-dark"" href=""#"" data-content=""about"" style = 'border:1px solid black; " & strHd & "'>Fault Reporting</a>"
response.write "      </li>"
response.write "      <li class=""nav-item"">"
response.write "        <a class=""nav-link btn text-dark"" href=""#"" data-content=""contact"" style = 'border:1px solid black;" & strHd & "'>Report Incident</a>"
response.write "      </li>"
response.write "    </ul>"
response.write "  </div>"
response.write "</nav>"
response.write ""

response.write "<div id=""home"" class=""content"">"

vehicleManagement

response.write "</div>"

response.write "<div id=""maintain"" class=""content container-fluid"">"
maintenaceMgt
response.write "</div>"
response.write ""
response.write "<div id=""about"" class=""content container-fluid"">"
 faultMgt
response.write "</div>"
response.write ""
response.write "<div id=""services"" class=""content"">"

    vehicleBooking

response.write "</div>"
response.write ""
response.write "<div id=""contact"" class=""content container-fluid"">"
ComplainMgt
response.write "</div>"
response.write "</div>"
End If

response.write "<script src=""https://code.jquery.com/jquery-3.6.0.min.js""></script>"
response.write "<script src=""https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/js/bootstrap.bundle.min.js"" integrity=""sha384-kenU1KFdBIe4zVF0s0G1M5b4hcpxyD9F7jL+jjXkk+Q2h455rYXK/7HAuoJl+0I4"" crossorigin=""anonymous""></script>"
response.write "<script src=""https://cdn.datatables.net/v/dt/dt-1.13.6/datatables.min.js""></script>"
response.write "<script src=""https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/pdfmake.min.js""></script>"
response.write "<script src=""https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/vfs_fonts.js""></script>"
response.write "<script src=""https://cdn.datatables.net/v/bs5/jszip-3.10.1/dt-1.13.6/b-2.4.1/b-html5-2.4.1/b-print-2.4.1/datatables.min.js""></script>"



response.write "</html>"
   
response.write "<script>"
response.write "        var table = $('#vehicleTable').DataTable({"
response.write "        });"
response.write "        var table = $('#incidentTable').DataTable({"
response.write "        });"
response.write "        var table = $('#faultTable').DataTable({"
response.write "        });"
response.write "        var table = $('#maintenanceTable').DataTable({"
response.write "        });"
response.write ""
response.write "        table.buttons().container().appendTo('#vehicleTable_wrapper .col-md-6:eq(0)');"

response.write "</script>"
   
   
   
   
 Sub vehicleManagement()

    Dim rst, sql, count, FinalAmount

    Set rst = CreateObject("ADODB.Recordset")

    
    

        response.write "<table class='table table-bordered' table-responsive border=""0"" cellpadding=""0"" cellspacing=""0"" style=""font-size:12pt"" width=""100%"">"
        response.write "<table class=""table table-responsive  table-hover"" cellpadding=""2"" border=""1"" cellspacing=""0"" style=""font-size:10pt"">"
        response.write "<tr>"
                                '<---- getting counts values for cards --------->
                                
                ambCnt = getTotalAmbulance()
                vhlCnt = getTotalVehicle()
                                'getting available ambulance
                                 avAmbl = getTotalVehicleByStatus("=", "A008")
                                 'getting available vehicle
                                 avVhl = getTotalVehicleByStatus("<>", "A008")
                                 'getting in-use ambulance
                                 inusAmbl = getTotalVehicleByStatus("=", "A001")
                                 'getting in-use vehicle
                                 insVhl = getTotalVehicleByStatus("<>", "A001")
                                 'getting on Maintenance ambulance
                                 mntAmbl = getTotalVehicleByStatus("=", "A003")
                                  'getting on Maintenance vehicle
                                 mntVhl = getTotalVehicleByStatus("<>", "A003")
                                '<---- End --------->
                response.write "<td>"
                response.write "<div class=""card text-white bg-info mb-3"">"
                response.write "    <div class=""card-body"">"
                response.write "      <h3 class=""card-title"">Total</h3>"
                response.write "        <i class=""fas fa-Book fa-2x""></i> "
              response.write "        <span class=""count fa-1x ""><b>Ambulance: " & ambCnt & "</b></span>"
                response.write "        <span class=""count fa-1x ""><b>Others: " & vhlCnt & "</b></span>"
                response.write "      </div>"
                response.write "    </div>"
            response.write "</td>"
            
            response.write "<td>"
                response.write "<div class=""card text-white bg-success mb-3"" >"
                response.write "    <div class=""card-body"" >"
                response.write "      <h3 class=""card-title"">Available</h3>"
                response.write "<div>"
                response.write "        <i class=""fas fa-car fa-2x""></i> "
                response.write "        <span class=""count fa-1x ""><b>Ambulance: " & avAmbl & "</b></span>"
                response.write "        <span class=""count fa-1x ""><b>Others: " & avVhl & "</b></span>"
                response.write "      </div>"
                response.write "    </div>"
                
               
                response.write "<td>"
                response.write "<div class=""card text-white bg-warning mb-3 "">"
                response.write "    <div class=""card-body"">"
                response.write "      <h3 class=""card-title"">In Use</h3>"
                response.write "        <i class=""fas fa-compass fa-2x""></i> "
                response.write "        <span class=""count fa-1x ""><b>Ambulance: " & inusAmbl & "</b></span>"
                 response.write "        <span class=""count fa-1x ""><b>Others: " & insVhl & "</b></span>"
                response.write "      </div>"
                response.write "    </div>"
                response.write "<td>"
                response.write "        <div class=""card text-white bg-danger mb-3"">"
                response.write "    <div class=""card-body"">"
                response.write "      <h3 class=""card-title"">Maintenance</h3>"
                response.write "        <i class=""fas fa-cog fa-2x""></i> "
                 response.write "        <span class=""count fa-1x ""><b>Ambulance: " & mntAmbl & "</b></span>"
                 response.write "        <span class=""count fa-1x ""><b>Others: " & mntVhl & "</b></span>"
                response.write "      </div>"
                response.write "    </div>"
            response.write "<td>"

            response.write "</tr>"
            response.write "<tr>"

                                response.write "<td colspan ='6'>"
                          
                                  response.write " <div class='float-end mr-2' style='margin-left:8px'> "
                                response.write "  <a data-href='wpgAssetIncident.asp?PageMode=AddNew' onclick= 'openPopup(this)' class='btn btn-primary float-end mr-2'>"
                                response.write "    <i class=""fa fa-cog""></i> Report Fault"
                                response.write "  </a>"
                                 response.write " </div> "
                                    response.write " <div class='float-end mr-2' style='margin-left:8px'> "
                                response.write "  <a data-href='wpgAssetMaintain.asp?PageMode=AddNew' onclick= 'openPopup(this)' class='btn btn-primary float-end mr-2'>"
                                response.write "    <i class=""fa fa-cog""></i> Maintenace"
                                response.write "  </a>"
                                 response.write " </div> "
                                response.write " <div class='float-end mr-2 'style='margin-left:8px'> "
                                response.write "  <a data-href='wpgItemRequest2.asp?PageMode=AddNew&ItemCategoryID=ITC002' onclick= 'openPopup(this)' class='btn btn-primary float-end mr-2'>"
                                response.write "    <i class=""fas fa-book""></i> Book Vehicle"
                                response.write "  </a>"
                                response.write " </div> "
                                response.write " <div class='float-end mr-2'style='margin-left:8px'> "
                                response.write "  <a data-href='wpgIncident.asp?PageMode=AddNew' onclick= 'openPopup(this)' class='btn btn-primary float-end mr-2'>"
                                response.write "    <i class=""fas fa-flag""></i> Report Incident"
                                response.write "  </a>"
                                response.write " </div> "
                                response.write " <div class='float-end mr-2' style='margin-left:15%' > "
                                response.write "  <a data-href='wpgAssetPurchase.asp?PageMode=AddNew' onclick= 'openPopup(this)' class='btn btn-primary float-end mr-2'>"
                                response.write "    <i class=""fas fa-plus""></i> Add Vehicle"
                                response.write "  </a>"
                                response.write " </div> "
                                response.write "</td>"
                                                    

                                response.write "<td></div>"
                                response.write "<td></div>"
            
            response.write "</tr>"
            response.write "</table> "
                        'response.write "            <h5 class = 'text-center'>LIST OF VEHICLES AND AMBULANCE</h5>"
                       ' response.write " <center> <input type='text' id='myinput' onkeyup='tableSearch()' placeholder='search' >"
                        response.write "<table id=""vehicleTable"" class=""table  table-bordered my-custom-table "" style='font-size:13px'>"
                        response.write "            <thead >"
                        response.write "              <tr style='background-color:#d2d2ed'>"
                        response.write "                <th style='background-color:#d2d2ed'>*</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Vehicle Number</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Vehicle Type</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Vehicle Capacity</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Status</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Vehicle Model</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Action</th>"
                        response.write "              </tr>"
                        response.write "            </thead>"
                        response.write "            <tbody>"
                brnch = "B001" ' for testing purpose
                sql = "SELECT * FROM AssetPurchase WHERE AssetCategoryID = 'ITC002' "
                cnt = 0
                        With rst
                        rst.open qryPro.FltQry(sql), conn, 3, 4
                        If rst.RecordCount > 0 Then
                                rst.movefirst

               Do While Not rst.EOF
               cnt = cnt + 1
               status = rst.fields("AssetPurStatusID")
               If (status = "A008") Then
               clr = "#02db3840"
               dis = ""
               ElseIf (status = "A001") Then
               clr = "#ffcb1b42"
               dis = "disable-link"
               ElseIf (status = "A003") Then
               clr = "#ff01274f"
               dis = "disable-link"
               End If
               aspID = rst.fields("AssetPurchaseID")
                response.write "<tr>"
            
                         response.write "                    <td style='background-color:" & clr & "'>" & cnt & "</td>"
                         response.write "                    <td>" & rst.fields("AssetPurchaseID") & "</td>"
                         response.write "                    <td>" & GetComboName("AssetType", rst.fields("AssetTypeID")) & "</td>"
                         response.write "                    <td>" & rst.fields("SerialNumber") & "</td>"
                         response.write "                    <td><span>" & GetComboName("AssetPurStatus", rst.fields("AssetPurStatusID")) & "</span></td>"
                         response.write "                    <td>" & rst.fields("AssetPurchaseName") & "</td>"
                         response.write "                    <td>"
                         response.write "                    <div class=""d-flex "">"
                          response.write "                                       <a data-href='wpgPrtPrintLayoutAll.asp?PrintLayoutName=AssetPurchaseRCP&PositionForTableName=AssetPurchase&AssetPurchaseID=" & aspID & " ' onclick=""openPopup(this)"" class=""action-icon "" "" ><i class=""fas fa-eye fa-1x "" Title='View Vehicle'></i></a>&nbsp;&nbsp;&nbsp"
                        ' response.write "                                <a data-href='wpgAssetPurchase.asp?PageMode=ProcessSelect&AssetPurchaseID=" & aspID & " ' onclick=""openPopup(this)"" class=""action-icon "" "" id = " & dis & "><i class=""fas fa-calendar-plus fa-1x "" Title='Book Vehicle'></i></a>&nbsp;&nbsp;&nbsp"
                         response.write "          <a data-href='wpgAssetPurchase.asp?PageMode=ProcessSelect&AssetPurchaseID=" & aspID & " ' onclick=""openPopup(this)"" class=""action-icon "" Title='Edit Vehicle'><i class=""fas fa-edit fa-1x""></i></a>&nbsp;&nbsp;&nbsp"
                         response.write "                                       <a data-href='wpgAssetmaintain.asp?PageMode=AddNew&PullUpData=AssetpurchaseID||" & aspID & "; ' onclick=""openPopup(this)"" class=""action-icon "" "" ><i class=""fas fa-cog fa-lx "" Title='Request Maintenance'></i></a>&nbsp;&nbsp;"
                         response.write "        </div>"
                         response.write "                    </td>"
                response.write "</tr>"
                rst.MoveNext
                response.Flush
            Loop
           
        End If

        rst.Close
    End With

    Set rst = Nothing
    response.write "              </tbody>"
    response.write "          </table>"
          

        response.write "</table>"
End Sub


Sub vehicleBooking()

        
Dim pat, patNm, dur, bDt, gen, genNm, sltHt1, sltHt2, dyHt, modMgr, cDt, wkDy, wkDyNm, cWkDy
Dim cnt, vDt1, vDt2, rst, pCnt, num, sql2, htStr, dyNm, prtUrl, patVdt, dCnt
Dim lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, tb, tbKy, tbNm, mdDet, wdNm
Dim recKy, hasPrt, vst, spTyp, spTypNm, vDt, lnkCnt, nDt, nDt2, prevDys, ordByTyp, sql0, startDy, endDy, dispTyp
Dim sDt, eDt, cMth, mth, clDpt, patDet, patAg, patAgDet, patBdt, patDys, spTypDet, sp, spNm, md, mdNm, mdOutWhCls
Dim labDet, radDet, pharmDet, theaDet, currSp, otSummary, lstWhCls2

    Set rst = CreateObject("ADODB.Recordset")
    lnkCnt = 0
    prevDys = 365
    ordByTyp = Trim(Request("OrderByType"))
    dispTyp = Trim(Request("DisplayType"))
    currSp = Trim(Request("Specialist"))
    currMs = Trim(Request("MedicalService"))
    If dur = Null Then
     dur = "MTH202310"
     Else
      dur = Trim(Request.QueryString("NoOfDays"))
    End If
   
    ' If Len(dur)<>9 Then
    '   dur = FormatWorkingMonth(Now())
    ' End If
    If IsEmpty(currMs) Then
      currMs = "M001"
    End If

    LoadCSS
    InitPageScript
    SetListWhCls2

    prtUrl = "wpgItemRequest2.asp?PageMode=ProcessSelect&ItemCategoryID=ITC002"
 
    cnt = 0
    cnt = cnt + 1
    nDt = Now()
    pCnt = 0

    cMth = ""
    mth = ""
    cWkDy = ""
    pCnt = 0
    dCnt = 0
    currSto = Glob_GetUserItemStore(jschd)
    response.write "<div class='container-fluid'>"

    response.write "<table class=""table table-striped table-bordered cmpTdSty"" cellpadding=""2"" border=""1"" cellspacing=""0"" width=""100%"" style=""font-size:10pt"">"

    response.write "<tr><td align=""left"" width=""100%"" valign=""top"" colspan=""10"">"
      response.write "<table border=""0"" cellpadding=""3"" cellspacing=""0"" style=""font-size:12pt"">"
       response.write "<tr><td colspan=""2"" align=""center"">"

        response.write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""font-size:12pt"" width=""100%"">"
            response.write "<tr><td class=""cpHdrTd2"" style=""color:" & Glob_BrandingColor("sage") & """>&emsp;<u>AMBULANCE&nbsp;/VEHICLE&nbsp;&nbsp;BOOKING&nbsp;</u>&emsp;</td>"
            dTyp = GetDispType2(jschd)
         response.write "<td>&emsp;&emsp;<b>Month:&nbsp;</b></td>"

         response.write "<td>"
         ' SetPrescriptionDays prevDys, nDt, nDt2, dur
         SetRequisitionMonth currSto, dur
         response.write "</td>"
        response.write "<td style=""" & Glob_BrandingColor("sage") & """ class=""cpHdrTd2"">&nbsp;&nbsp;<u>As&nbsp;At&nbsp;:&nbsp;&nbsp;" & FormatDateDetail(Now()) & "</u>&emsp;&emsp;&nbsp;</td>"

        lnkCnt = "5"
        lnkID = "trslt||lnk" & CStr(lnkCnt)
        response.write "<td onclick=""RefreshPage()"" class=""btn_"" style=""color:#8888ee"" id=""" & lnkID & """ onmouseover=""DoOnMouseOverNav5 ('" & lnkID & "')"" onmouseout=""DoOnMouseOutNav5 ('" & lnkID & "')"">"
        response.write "<b>Refresh&nbsp;</b></td>"

    
        response.write "<td valign=""top"">"
        sTb2 = "ItemRequest2"
        If HasAccessRight(uName, "frm" & sTb2, "New") Then
        'Clickable Url Link
        lnkCnt = lnkCnt + 1
        lnkID = "lnk" & CStr(lnkCnt)
        lnkText = "<b>Make New Request</b>"
        lnkUrl = "wpgItemRequest2.asp?PageMode=AddNew&ItemCategoryID=ITC002" '' & "&ItemRequest2ID=" & vst
        navPop = "POP"
        inout = "IN"
        fntSize = "10"
        fntColor = "#444488"
        bgColor = ""
        wdth = ""
        Glob_AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
        End If
        response.write "</td>"
        response.write "</tr>"
        response.write "</table>"

       response.write "</td></tr>"
      response.write "</table>"
    response.write "</td></tr>"


    sql = "select ItemRequest2.*, Staff.StaffName "
    sql = sql & " From ItemRequest2, SystemUser, Staff Where ItemRequest2.SystemUserID=SystemUser.SystemUserID And SystemUser.StaffID=Staff.StaffID  "
    ' sql = sql & " And ItemRequest2.WorkingMonthID='" & dur & "' And ItemRequest2.BranchID='" & brnch & "' "
    sql = sql & " And ItemRequest2.WorkingMonthID='" & dur & "' And ItemRequest2.ItemCategoryID = 'ITC002' "
    If Glob_HasTransProcessAccess2("ItemRequest2Pro", uName) Then
        ' sql = sql & " And (ItemRequest2.ItemStoreID='" & currSto & "' Or ItemRequest2.ItemRequestStoreID='" & currSto & "')  "
    ElseIf Trim(currSto) = "" Then
        sql = sql & " And ItemRequest2.BranchID='" & brnch & "' "
    ElseIf Trim(currSto) <> "" Then
        sql = sql & " And (ItemRequest2.ItemStoreID='" & currSto & "' Or ItemRequest2.ItemRequestStoreID='" & currSto & "')  "
        sql = sql & " And ItemRequest2.BranchID='" & brnch & "' "
    End If
    sql = sql & " order by ItemRequest2.RequestDate desc "
    ' response.write sql
    With rst
        '.maxrecords = 50
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            .movefirst
            wkDyNm = GetComboName("WorkingMonth", dur)
            response.write "<tr style=""font-weight:bold;font-size:12pt"" bgcolor=""ffffff""><td colspan=""100"" align=""left"" valign=""top"">"
            response.write "<b>" & wkDyNm & "</b>&emsp;->&emsp; " & rst.RecordCount & " Requests "
            response.write "&emsp;&emsp;"
                response.write "My Store: " & GetComboName("JobSchedule", currSto)
            response.write "</td></tr>"

            response.write "<tr style=""font-weight:bold;font-size:12pt"" bgcolor=""#eeeeee"">"
            response.write "<td valign=""top"" align=""center"">No.</td>"
            response.write "<td valign=""top"">Request&nbsp;Details</td>"
            response.write "<td valign=""top"">Request&nbsp;Vehicle</td>"
            response.write "<td valign=""top"">Request&nbsp;Description</td>"
            response.write "<td valign=""top"">Approval&nbsp;Details</td>"
            ' response.write "<td valign=""top"">Status</td>"
            response.write "<td valign=""top"">Issuance</td>"
            response.write "<td valign=""top"">Acceptance</td>"
    
            response.write "</tr>"
            Do While Not .EOF
                vDt = ""
                patDet = ""
                clr = "#0fff0045" ''Final, Green
                TransProcessStatID = rst.fields("TransProcessStatID")
                dtReq = .fields("RequestDate")
                tmAgo = Glob_GetHowLong(dtReq, Now())
                spTypDet = ""
                jbNm = GetComboName("JobSchedule", .fields("JobScheduleID"))
                reqDrg = .fields("ItemRequest2ID")
                drgSto = .fields("ItemStoreID")
                reqNm = .fields("ItemRequest2Name")
                spTypNm = .fields("StaffName")
                md = .fields("ItemRequestStoreID")
                dscr = .fields("remarks")
                dest = .fields("requestinfo1")
                trpDt = .fields("ClosedDate")
                durt = .fields("ClosedBy")
                
                pCnt = pCnt + 1
                If UCase(TransProcessStatID) = UCase("T001") Then ''Initial, red
                    clr = "#ff000045"
                ElseIf UCase(TransProcessStatID) = UCase("T002") Then ''Authorize, yellow
                    clr = "#ffff0045"
                End If
                If UCase(currSto) = UCase(md) Then
                        tmAgo = tmAgo & "<br><b>Incoming Request</b>"
                ElseIf UCase(currSto) = UCase(drgSto) Then
                        tmAgo = tmAgo & "<br><b>Outgoing Request</b>"
                Else ''If UCase(currSto)=UCase(drgSto) Then
                        ' tmAgo = tmAgo & "<br><b>Approve/Authorization</b>"
                End If


                'Requisition
                spTypDet = "<b>" & spTypNm & "<br>" & jbNm & "</b><br>"
                patDet = "<br>" & reqNm & "<br>No:&nbsp;<b>" & reqDrg & "</b><br>" & tmAgo
                ' patDet = Replace(patDet, " ", "&nbsp;")

                response.write "<tr>"
                response.write "<td valign=""top"" align=""center"" style=""background-color:" & clr & ";"">" & CStr(pCnt) & "</td>"

                ' response.write "<td valign=""top"">" & patDet & "</td>"
                response.write "<td valign=""top"">"
                response.write spTypDet & patDet & "<br>"
                'Clickable Url Link
                lnkCnt = lnkCnt + 1
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = "<b>Adjust Request</b>"
                lnkUrl = prtUrl & "&ItemRequest2ID=" & reqDrg
                navPop = "POP"
                inout = "IN"
                fntSize = "10"
                fntColor = "#3f8a00"
                bgColor = ""
                wdth = ""
                Glob_AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
                response.write "</td>"
                
                 
                
                response.write "<td valign=""top"">"
                DisplayRequestedItems reqDrg
                response.write "</td>"
                
                'response.write "<td valign=""top"">" & dscr & "</td>"
                
                response.write "<td valign=""top"">"
                response.write "<table class='table table-bordered'>"
                response.write "<tr>"
                response.write "<th>Trip Date/Time</th>"
                response.write "<th>Duration</th>"
                response.write "<th>Destination</th>"
                response.write "</tr>"
                response.write "<tr>"
                response.write "<td>" & trpDt & "</td>"
                response.write "<td>" & durt & "hr(s)</td>"
                response.write "<td>" & dest & "</td>"
                response.write "</tr>"
                response.write "<tr><td colspan='3' >" & dscr & "</td></tr>"
                response.write "</table>"
                response.write "</td>"

                    
                response.write "<td valign=""top"">"
                DisplayApprovals reqDrg
                response.write "</td>"
                
               
               
                response.write "<td valign=""top"">"
                reqIss = DisplayIssuedItems(reqDrg)
                response.write "</td>"

                response.write "<td valign=""top"">"
                If (jschd = "M10A") Then
                DisplayCompletedTrip reqDrg
                response.write "  <a data-href='wpgPrtPrintLayoutAll.asp?PrintLayoutName=BookingCheckList&PositionForTableName=WorkingDay&WorkingDayID=&requested=" & requested & "&approved=" & approved & "&availabilty=" & available & "&Issued=" & issued & "&tripID=" & reqDrg & "&tripdate=" & trpDt & "&destination=" & dest & "&vehicleNo=" & vehicleNo & " ' onclick= 'openPopup(this)' class='btn btn-outline-success float-end'>"
                response.write "  <i class=""fas fa-print""></i> CheckList"
                response.write "  </a>"
                Else
                End If
                response.write "</td>"

                response.write "</tr>"
                .MoveNext
            Loop
        Else
            response.write "<tr><td colspan=""100"">No Requisition for this Month and Facility</td></tr>"
        End If
        .Close
    End With

    response.write "</table>"
    Set rst = Nothing

    response.Flush
    response.write "</div>"
End Sub

Sub faultMgt()
    Dim rst, sql, count, FinalAmount

    Set rst = CreateObject("ADODB.Recordset")

            response.write "<table class='table table-bordered' table-responsive border=""0"" style=""font-size:12pt"" width=""100%"">"
            response.write "<table class=""table table-responsive  table-hover"" cellpadding=""2"" border=""1"" cellspacing=""0"" style=""font-size:10pt"">"
                response.write "<tr>"
                response.write "</tr>"
                response.write "<tr>"
            response.write "<tr>"

          
               response.write "<td colspan ='6'>"
                          
                                  response.write " <div class='float-end mr-2' style='margin-left:8px'> "
                                response.write "  <a data-href='wpgAssetIncident.asp?PageMode=AddNew' onclick= 'openPopup(this)' class='btn btn-primary float-end mr-2'>"
                                response.write "    <i class=""fa fa-cog""></i> Report Fault"
                                response.write "  </a>"
                                 response.write " </div> "
                                     response.write " <div class='float-end mr-2' style='margin-left:8px'> "
                                response.write "  <a data-href='wpgAssetMaintain.asp?PageMode=AddNew' onclick= 'openPopup(this)' class='btn btn-primary float-end mr-2'>"
                                response.write "    <i class=""fa fa-cog""></i>Maintenace"
                                response.write "  </a>"
                                 response.write " </div> "
                                response.write " <div class='float-end mr-2 'style='margin-left:8px'> "
                                response.write "  <a data-href='wpgItemRequest2.asp?PageMode=AddNew&ItemCategoryID=ITC002 ' onclick= 'openPopup(this)' class='btn btn-primary float-end mr-2'>"
                                response.write "    <i class=""fas fa-book""></i> Book Vehicle"
                                response.write "  </a>"
                                response.write " </div> "
                                response.write " <div class='float-end mr-2'style='margin-left:8px'> "
                                response.write "  <a data-href='wpgIncident.asp?PageMode=AddNew' onclick= 'openPopup(this)' class='btn btn-primary float-end mr-2'>"
                                response.write "    <i class=""fas fa-flag""></i> Report Incident"
                                response.write "  </a>"
                                response.write " </div> "
                                response.write " <div class='float-end mr-2' style='margin-left:15%' > "
                                response.write "  <a data-href='wpgAssetPurchase.asp?PageMode=AddNew' onclick= 'openPopup(this)' class='btn btn-primary float-end mr-2'>"
                                response.write "    <i class=""fas fa-plus""></i> Add Vehicle"
                                response.write "  </a>"
                                response.write " </div> "
                                response.write "</td>"
            response.write "</td>"

            response.write "<td></div>"
            response.write "<td></div>"
            
            response.write "</tr>"
            response.write "</table> "
           ' response.write "            <h5 class = 'text-center'>LIST OF FAULT REQUESTS</h5>"
                        response.write "        <table id=""faultTable"" class=""table  table-bordered my-custom-table "" style='font-size:13px'>"
                        response.write "            <thead>"
                        response.write "              <tr>"
                        response.write "                <th style='background-color:#d2d2ed'>#</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Inc. Number</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Title</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Vehicle</th>"
                        'response.write "                <th style='background-color:#d2d2ed'>Vehicle Parts</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Priority</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Status</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Details</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Entry Date</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Assigned To</th>"
                        response.write "                <th style='background-color:#d2d2ed'>End Date</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Entered By</th>"
                       ' response.write "                <th style='background-color:#d2d2ed'>Approvals</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Comments</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Actions</th>"
                        response.write "              </tr>"
                        response.write "            </thead>"
                        response.write "            <tbody>"
    brnch = "B001" ' for testing purpose
    sql = "SELECT Assetincidentid,assetincidentName,assetpurchaseid,incidentcategoryid,incidentmodeid,incidentdetail,incidentdate1,incidentpriorityid,incidentstatusid,startdate,enddate,assignedtoid,assigneddate,remarks,jobscheduleid,systemuserid,entrydate FROM assetincident where directorytypeid = 'C001' ORDER BY Assetincidentid DESC "
    
    cnt = 0
    With rst
        rst.open qryPro.FltQry(sql), conn, 3, 4

        If rst.RecordCount > 0 Then
            rst.movefirst

               Do While Not rst.EOF
               cnt = cnt + 1
               status = rst.fields("incidentstatusid")
                If (status = "I001") Then
               clr = "#ffcb1996"
               ElseIf (status = "I002") Then
               clr = "#ffcb1996"
               ElseIf (status = "I003") Then
               clr = "#ff012766"
                ElseIf (status = "I004") Then
               clr = "#02db3878"
                ElseIf (status = "I005") Then
               clr = "#ffffff"
               End If
                 rqtID = rst.fields("Assetincidentid")
                response.write "<tr>"
                          
                         response.write "                  <td style='background-color:" & clr & "'>" & cnt & "</td>"
                         response.write "                    <td>" & rst.fields("Assetincidentid") & "</td>"
                         response.write "                    <td>" & rst.fields("assetincidentName") & "</td>"
                         response.write "                    <td>" & rst.fields("assetpurchaseid") & "</td>"
                                                             ' getAssetPartsList rqtID
                         response.write "                    <td>" & GetComboName("incidentpriority", rst.fields("incidentpriorityid")) & "</td>"
                         response.write "                    <td>" & GetComboName("incidentstatus", rst.fields("incidentstatusid")) & "</td>"
                         response.write "                    <td>" & rst.fields("incidentdetail") & "</td>"
                         response.write "                    <td>" & rst.fields("entrydate") & "</td>"
                         response.write "                    <td><span>" & GetComboName("assignedto", rst.fields("assignedtoid")) & "</span></td>"
                         response.write "                    <td>" & rst.fields("startdate") & "</td>"
                         response.write "                    <td><span>" & GetComboName("systemuser", rst.fields("systemuserid")) & "</span></td>"
                        ' response.write "                    <td style = 'color:red'>Not Approved</td>"
                         response.write "                    <td>" & rst.fields("remarks") & "</td>"
                         response.write "                    <td>"
                         response.write "                               <div class=""d-flex "">"
                         'response.write "                                       <a href=""#"" class=""action-icon""><i class=""fa fa-thumbs-up fa-2x "" Title='Approve Request'></i></a>&nbsp;&nbsp;"
                         response.write "          <a data-href='wpgAssetIncident.asp?PageMode=ProcessSelect&AssetIncidentID=" & rqtID & " ' onclick=""openPopup(this)"" class=""action-icon "" Title='Edit Request'><i style='font-size:16px' class=""fas fa-edit ""></i></a>&nbsp;&nbsp;"
                         'response.write "          <a href=""#"" class=""action-icon "" Title='Assign to'><i class=""fas fa-user fa-lx""></i></a>"
                                                 response.write "                               </div>"
                         response.write "                    </td>"
                response.write "</tr>"
                rst.MoveNext
                response.Flush
            Loop
           
        End If

        rst.Close
    End With

    Set rst = Nothing
    response.write "              </tbody>"
    response.write "          </table>"
    
    response.write "</table>"

End Sub

'Maintenace management starts here

Sub maintenaceMgt()
    Dim rst, sql, count, FinalAmount

    Set rst = CreateObject("ADODB.Recordset")

            response.write "<table class='table table-bordered' table-responsive border=""0"" style=""font-size:12pt"" width=""100%"">"
            response.write "<table class=""table table-responsive  table-hover"" cellpadding=""2"" border=""1"" cellspacing=""0"" style=""font-size:10pt"">"
                response.write "<tr>"
                response.write "</tr>"
                response.write "<tr>"
            response.write "<tr>"

          
               response.write "<td colspan ='6'>"
                          
                                  response.write " <div class='float-end mr-2' style='margin-left:8px'> "
                                response.write "  <a data-href='wpgAssetIncident.asp?PageMode=AddNew' onclick= 'openPopup(this)' class='btn btn-primary float-end mr-2'>"
                                response.write "    <i class=""fa fa-cog""></i> Report Fault"
                                response.write "  </a>"
                                 response.write " </div> "
                                     response.write " <div class='float-end mr-2' style='margin-left:8px'> "
                                response.write "  <a data-href='wpgAssetMaintain.asp?PageMode=AddNew' onclick= 'openPopup(this)' class='btn btn-primary float-end mr-2'>"
                                response.write "    <i class=""fa fa-cog""></i>Maintenace"
                                response.write "  </a>"
                                 response.write " </div> "
                                response.write " <div class='float-end mr-2 'style='margin-left:8px'> "
                                response.write "  <a data-href='wpgItemRequest2.asp?PageMode=AddNew&ItemCategoryID=ITC002' onclick= 'openPopup(this)' class='btn btn-primary float-end mr-2'>"
                                response.write "    <i class=""fas fa-book""></i> Book Vehicle"
                                response.write "  </a>"
                                response.write " </div> "
                                response.write " <div class='float-end mr-2'style='margin-left:8px'> "
                                response.write "  <a data-href='wpgIncident.asp?PageMode=AddNew' onclick= 'openPopup(this)' class='btn btn-primary float-end mr-2'>"
                                response.write "    <i class=""fas fa-flag""></i> Report Incident"
                                response.write "  </a>"
                                response.write " </div> "
                                response.write " <div class='float-end mr-2' style='margin-left:15%' > "
                                response.write "  <a data-href='wpgAssetPurchase.asp?PageMode=AddNew' onclick= 'openPopup(this)' class='btn btn-primary float-end mr-2'>"
                                response.write "    <i class=""fas fa-plus""></i> Add Vehicle"
                                response.write "  </a>"
                                response.write " </div> "
                                response.write "</td>"
            response.write "</td>"

            response.write "<td></div>"
            response.write "<td></div>"
            
            response.write "</tr>"
            response.write "</table> "
           ' response.write "            <h5 class = 'text-center'>LIST OF MAINTENACE REQUESTS</h5>"
                        response.write "        <table id=""maintenanceTable"" class=""table  table-bordered my-custom-table "" style='font-size:13px'>"
                        response.write "            <thead>"
                        response.write "              <tr>"
                        response.write "                <th style='background-color:#d2d2ed'>#</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Routine. Number</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Vehicle</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Vehicle Type</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Details</th>"
                         response.write "                <th style='background-color:#d2d2ed'>Maintenance Category</th>"
                         response.write "                <th style='background-color:#d2d2ed'>Maintenance Mode</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Priority</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Status</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Date</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Requested By</th>"
                       ' response.write "                <th style='background-color:#d2d2ed'>Approvals</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Actions</th>"
                        response.write "              </tr>"
                        response.write "            </thead>"
                        response.write "            <tbody>"
    brnch = "B001" ' for testing purpose
    sql = "SELECT Assetmaintainid,assetpurchaseid,assetcategoryid,assettypeid,maintainmodeid,systemuserid,maintaincategoryid,maintaintypeid,maintaindetail,entrydate,maintainpriorityid,maintainstatusid FROM AssetMaintain WHERE directoryTypeID = 'C001'"

    cnt = 0
    With rst
        rst.open qryPro.FltQry(sql), conn, 3, 4

        If rst.RecordCount > 0 Then
            rst.movefirst

               Do While Not rst.EOF
               cnt = cnt + 1
               status = rst.fields("maintainstatusid")
                If (status = "I001") Then
               clr = "#ffcb1996"
               ElseIf (status = "I002") Then
               clr = "#ffcb1996"
               ElseIf (status = "I003") Then
               clr = "#ff012766"
                ElseIf (status = "I004") Then
               clr = "#02db3878"
                ElseIf (status = "I005") Then
               clr = "#ffffff"
               End If
                 rqtID = rst.fields("Assetmaintainid")
                response.write "<tr>"
                          
                         response.write "                  <td style='background-color:" & clr & "'>" & cnt & "</td>"
                         response.write "                    <td>" & rst.fields("Assetmaintainid") & "</td>"
                         response.write "                    <td>" & rst.fields("assetpurchaseid") & "</td>"
                                                             ' getAssetPartsList rqtID
                         response.write "                    <td>" & GetComboName("assettype", rst.fields("assettypeid")) & "</td>"
                         response.write "                    <td>" & rst.fields("maintaindetail") & "</td>"
                         response.write "                    <td>" & GetComboName("maintaincategory", rst.fields("maintaincategoryid")) & "</td>"
                          response.write "                    <td>" & GetComboName("maintainmode", rst.fields("maintainmodeid")) & "</td>"
                           response.write "                    <td>" & GetComboName("maintainpriority", rst.fields("maintainpriorityid")) & "</td>"
                           response.write "                    <td>" & GetComboName("maintainstatus", rst.fields("maintainstatusid")) & "</td>"
                         response.write "                    <td>" & rst.fields("entrydate") & "</td>"
                         response.write "                    <td><span>" & GetComboName("systemuser", rst.fields("systemuserid")) & "</span></td>"
                         'response.write "                    <td style = 'color:red'>Not Approved</td>"
                         response.write "                    <td>"
                         response.write "                               <div class=""d-flex "">"
                         'response.write "                                       <a href=""#"" class=""action-icon""><i class=""fa fa-thumbs-up fa-1x "" Title='Approve Request'></i></a>&nbsp;&nbsp;"
                         response.write "          <a data-href='wpgAssetmaintain.asp?PageMode=ProcessSelect&AssetmaintainID=" & rqtID & " ' onclick=""openPopup(this)"" class=""action-icon "" Title='Edit Request'><i style='font-size:16px' class=""fas fa-edit fa-1x""></i></a>&nbsp;&nbsp;"
                         'response.write "          <a href=""#"" class=""action-icon "" Title='Assign to'><i class=""fas fa-user fa-lx""></i></a>"
                                                 response.write "                               </div>"
                         response.write "                    </td>"
                response.write "</tr>"
                rst.MoveNext
                response.Flush
            Loop
           
        End If

        rst.Close
    End With

    Set rst = Nothing
    response.write "              </tbody>"
    response.write "          </table>"
    
    response.write "</table>"

End Sub



Sub ComplainMgt()
    Dim rst, sql, count, FinalAmount

    Set rst = CreateObject("ADODB.Recordset")

        response.write "<table class='table table-bordered' table-responsive border=""0"" cellpadding=""0"" cellspacing=""0"" style=""font-size:12pt"" width=""100%"">"
        response.write "<table class=""table table-responsive  table-hover"" cellpadding=""2"" border=""1"" cellspacing=""0"" style=""font-size:10pt"">"
        response.write "<tr>"
        response.write "</tr>"
        response.write "<tr>"
        

            response.write "<tr>"
'            response.write "<td class='text-center'>"
'            response.write "<span class=' d-flex'> Month: &nbsp;"
'            SetComplaintMonth
'            response.write "</span>"
'
'            response.write "</td>"
'            response.write "<td class='text-center'>"
'            response.write "<span class=' d-flex'> Priority: &nbsp;"
'            SetComplaintPriority
'            response.write "</span>"
'
'            response.write "</td>"
'            response.write "<td class='text-center'>"
'            response.write "<span class=' d-flex'> Status: &nbsp;"
'            SetComplaintStatus
'            response.write "</span>"
'
'            response.write "</td>"
            
           response.write "<td colspan ='6'>"
                          
                                  response.write " <div class='float-end mr-2' style='margin-left:8px'> "
                                response.write "  <a data-href='wpgAssetIncident.asp?PageMode=AddNew' onclick= 'openPopup(this)' class='btn btn-primary float-end mr-2'>"
                                response.write "    <i class=""fa fa-cog""></i> Report Fault"
                                response.write "  </a>"
                                 response.write " </div> "
                                     response.write " <div class='float-end mr-2' style='margin-left:8px'> "
                                response.write "  <a data-href='wpgAssetMaintain.asp?PageMode=AddNew' onclick= 'openPopup(this)' class='btn btn-primary float-end mr-2'>"
                                response.write "    <i class=""fa fa-cog""></i> Maintenance"
                                response.write "  </a>"
                                 response.write " </div> "
                                response.write " <div class='float-end mr-2 'style='margin-left:8px'> "
                                response.write "  <a data-href='wpgItemRequest2.asp?PageMode=AddNew&ItemCategoryID=ITC002' onclick= 'openPopup(this)' class='btn btn-primary float-end mr-2'>"
                                response.write "    <i class=""fas fa-book""></i> Book Vehicle"
                                response.write "  </a>"
                                response.write " </div> "
                                response.write " <div class='float-end mr-2'style='margin-left:8px'> "
                                response.write "  <a data-href='wpgIncident.asp?PageMode=AddNew' onclick= 'openPopup(this)' class='btn btn-primary float-end mr-2'>"
                                response.write "    <i class=""fas fa-flag""></i> Report Incident"
                                response.write "  </a>"
                                response.write " </div> "
                                response.write " <div class='float-end mr-2' style='margin-left:15%' > "
                                response.write "  <a data-href='wpgAssetPurchase.asp?PageMode=AddNew' onclick= 'openPopup(this)' class='btn btn-primary float-end mr-2'>"
                                response.write "    <i class=""fas fa-plus""></i> Add Vehicle"
                                response.write "  </a>"
                                response.write " </div> "
                                response.write "</td>"
            response.write "</td>"
            response.write "<td></div>"
            response.write "<td></div>"
            
            response.write "</tr>"
            response.write "</table> "
                      '  response.write "            <h5 class = 'text-center'>LIST OF REPORTED INCIDENTS</h5>"
                        response.write "    <table id=""incidentTable"" class=""table  table-bordered my-custom-table "" style='font-size:13px'>"
                        response.write "            <thead>"
                        response.write "              <tr>"
                        response.write "                <th style='background-color:#d2d2ed'>#</th>"
                        response.write "                <th style='background-color:#d2d2ed'>ID</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Title</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Department</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Category</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Details</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Priority</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Status</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Reported By</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Date</th>"
                        'response.write "                <th style='background-color:#d2d2ed'>Approval</th>"
                        response.write "                <th style='background-color:#d2d2ed'>Actions</th>"
                        response.write "              </tr>"
                        response.write "            </thead>"
                        response.write "            <tbody>"
    brnch = "B001" ' for testing purpose
    sql = "SELECT incidentid,incidentname,incidentcategoryid,incidenttypeid,incidentdetail,incidentpriorityid,incidentstatusid,startdate,enddate,Assignedtoid,assignedDate,clienttypeid,SystemUserid,entrydate FROM incident WHERE clienttypeid = 'C011' ORDER by entrydate Desc"
    cnt = 0
    With rst
        rst.open qryPro.FltQry(sql), conn, 3, 4

        If rst.RecordCount > 0 Then
            rst.movefirst

               Do While Not rst.EOF
               cnt = cnt + 1
               status = rst.fields("incidentstatusid")
               If (status = "I001") Then
               clr = "#ffcb1996"
               ElseIf (status = "I002") Then
               clr = "#ffcb1996"
               ElseIf (status = "I003") Then
               clr = "#ff012766"
                ElseIf (status = "I004") Then
               clr = "#02db3878"
                ElseIf (status = "I005") Then
               clr = "#ffffff"
               End If

               incID = rst.fields("incidentid")
                response.write "<tr>"
            
                         response.write "                  <td style='background-color:" & clr & "'>" & cnt & "</td>"
                         response.write "                    <td>" & rst.fields("incidentid") & "</td>"
                          response.write "                    <td>" & rst.fields("incidentname") & "</td>"
                        response.write "                    <td>" & GetComboName("clienttype", rst.fields("clienttypeid")) & "</td>"
                         response.write "                    <td>" & GetComboName("incidentcategory", rst.fields("incidentcategoryid")) & "</td>"
                        response.write "                    <td>" & rst.fields("incidentdetail") & "</td>"
                         response.write "                    <td><span><b>" & GetComboName("incidentpriority", rst.fields("incidentpriorityid")) & "</b></span></td>"
                         response.write "                    <td><span>" & GetComboName("incidentstatus", rst.fields("incidentstatusid")) & "</span></td>"
                          response.write "                    <td><span>" & GetComboName("SystemUser", rst.fields("SystemUserid")) & "</span></td>"
                         response.write "                    <td>" & rst.fields("entrydate") & "</td>"
                        ' response.write "                    <td><h6>Not Approved</h6></td>"
                         response.write "                    <td>"
                         response.write "         <div class=""d-flex "">"
                         'response.write "          <a data-href='wpgIncidentpro.asp?PageMode=AddNew&PullupData=IncidentID||" & incID & " ' onclick=""openPopup(this)"" class=""action-icon""><i class=""fa fa-thumbs-up fa-1x "" Title='Approve Request'></i></a>&nbsp;&nbsp;"
                         response.write "          <a data-href='wpgIncident.asp?PageMode=processselect&IncidentID=" & incID & " ' onclick=""openPopup(this)"" class=""action-icon "" Title='Edit Request'><i style='font-size:16px' class=""fas fa-edit fa-1x""></i></a>&nbsp;&nbsp;"
                         'response.write "                  <a href=""#"" class=""action-icon "" Title='Assign to'><i class=""fas fa-user fa-lx""></i></a>"
                                                 response.write "        </div>"
                         response.write "                    </td>"
                response.write "</tr>"
                rst.MoveNext
                response.Flush
            Loop
           
        End If

        rst.Close
    End With

    Set rst = Nothing

    response.write "              </tbody>"
    response.write "          </table>"
    
response.write "</table>"
End Sub

Sub getPassList()
    response.write "<td> "
    response.write " <table class ='table table-bordered table-striped' style='font-size:12px'>"
    response.write "    <tr>"
    response.write "        <td>1</td>"
    response.write "        <td>James Offei</td>"
    response.write "    </tr>"
    response.write "    <tr>"
    response.write "        <td>2</td>"
    response.write "        <td>Ransford Quaye</td>"
    response.write "    </tr>"
    response.write "    <tr>"
    response.write "        <td>3</td>"
    response.write "        <td>John Abeka</td>"
    response.write "    </tr>"
    response.write "</table>"
    response.write "</td> "
End Sub

Sub SetVehicleStatus(emrDataID)
    Set rst = CreateObject("ADODB.Recordset")
    dyHt = "<select class='form-select' size=""1"" name=""emrdata"" id=""emrdata"" onchange=""emrdataOnchange()"">"
    dyHt = dyHt & "<option value=""""></option>"
    sql0 = "select DISTINCT AssetPurStatusID from AssetPurchase WHERE AssetDepartmentid = 'A003'"
 

    With rst
      .open qryPro.FltQry(sql0), conn, 3, 4
      If .RecordCount > 0 Then
        .movefirst
        Do While Not .EOF
          emrid = Trim(.fields("AssetPurStatusID"))
          emrName = GetComboName("AssetPurStatus", Trim(.fields("AssetPurStatusID"))) '' & " -> " & GetComboName("WorkingYear", yr)
          
          If UCase(CStr(emrid)) = UCase(AssetPurStatusID) Then
           dyHt = dyHt & "<option value=""" & CStr(emrid) & """ Selected >" & emrName & "</option>"
        Else
           dyHt = dyHt & "<option value=""" & CStr(emrid) & """ >" & emrName & "</option>"
        End If
        rst.MoveNext
            response.Flush
        Loop
      End If
      .Close
    End With
    dyHt = dyHt & "</select>"
    response.write dyHt
    Set rst = Nothing
End Sub

Sub SetVehicleType(emrDataID)
    Set rst = CreateObject("ADODB.Recordset")
    dyHt = "<select class='form-select' size=""1"" name=""emrdata"" id=""emrdata"" onchange=""emrdataOnchange()"">"
    dyHt = dyHt & "<option value=""""></option>"
    sql0 = "select DISTINCT AssetTypeID from AssetPurchase WHERE AssetDepartmentid = 'A003'"
 

    With rst
      .open qryPro.FltQry(sql0), conn, 3, 4
      If .RecordCount > 0 Then
        .movefirst
        Do While Not .EOF
          emrid = Trim(.fields("AssetTypeID"))
          emrName = GetComboName("AssetType", Trim(.fields("AssetTypeID"))) '' & " -> " & GetComboName("WorkingYear", yr)
          
          If UCase(CStr(emrid)) = UCase(AssetTypeID) Then
           dyHt = dyHt & "<option value=""" & CStr(emrid) & """ Selected >" & emrName & "</option>"
        Else
           dyHt = dyHt & "<option value=""" & CStr(emrid) & """ >" & emrName & "</option>"
        End If
        rst.MoveNext
            response.Flush
        Loop
      End If
      .Close
    End With
    dyHt = dyHt & "</select>"
    response.write dyHt
    Set rst = Nothing
End Sub

'code to get complaint priority
Sub SetComplaintPriority()
    Set rst = CreateObject("ADODB.Recordset")
    dyHt = "<select class='form-select' size=""1"" name=""emrdata"" id=""emrdata"" onchange=""emrdataOnchange()"">"
    dyHt = dyHt & "<option value=""""></option>"
    sql0 = "select DISTINCT incidentpriorityid from Incident WHERE clienttypeid = 'C011'"
 

    With rst
      .open qryPro.FltQry(sql0), conn, 3, 4
      If .RecordCount > 0 Then
        .movefirst
        Do While Not .EOF
          emrid = Trim(.fields("incidentpriorityid"))
          emrName = GetComboName("incidentpriority", Trim(.fields("incidentpriorityid"))) '' & " -> " & GetComboName("WorkingYear", yr)
          
          If UCase(CStr(emrid)) = UCase(incidentpriorityid) Then
           dyHt = dyHt & "<option value=""" & CStr(emrid) & """ Selected >" & emrName & "</option>"
        Else
           dyHt = dyHt & "<option value=""" & CStr(emrid) & """ >" & emrName & "</option>"
        End If
        rst.MoveNext
            response.Flush
        Loop
      End If
      .Close
    End With
    dyHt = dyHt & "</select>"
    response.write dyHt
    Set rst = Nothing
End Sub


Sub SetComplaintMonth()
    Set rst = CreateObject("ADODB.Recordset")
    dyHt = "<select class='form-select' size=""1"" name=""emrdata"" id=""emrdata"" onchange=""emrdataOnchange()"">"
    dyHt = dyHt & "<option value=""""></option>"
    sql0 = "select DISTINCT WorkingMonthID,WorkingYearID from Incident WHERE clienttypeid = 'C011'"
 

    With rst
      .open qryPro.FltQry(sql0), conn, 3, 4
      If .RecordCount > 0 Then
        .movefirst
        Do While Not .EOF
          emrid = Trim(.fields("WorkingMonthID"))
          emrName = GetComboName("WorkingMonth", Trim(.fields("WorkingMonthID"))) '' & " -> " & GetComboName("WorkingYear", yr)
          
          If UCase(CStr(emrid)) = UCase(WorkingMonthID) Then
           dyHt = dyHt & "<option value=""" & CStr(emrid) & """ Selected >" & emrName & "</option>"
        Else
           dyHt = dyHt & "<option value=""" & CStr(emrid) & """ >" & emrName & "</option>"
        End If
        rst.MoveNext
            response.Flush
        Loop
      End If
      .Close
    End With
    dyHt = dyHt & "</select>"
    response.write dyHt
    Set rst = Nothing
End Sub

Sub SetComplaintStatus()
    Set rst = CreateObject("ADODB.Recordset")
    dyHt = "<select class='form-select' size=""1"" name=""emrdata"" id=""emrdata"" onchange=""emrdataOnchange()"">"
    dyHt = dyHt & "<option value=""""></option>"
    sql0 = "select DISTINCT IncidentStatusID from Incident WHERE clienttypeid = 'C011'"
 

    With rst
      .open qryPro.FltQry(sql0), conn, 3, 4
      If .RecordCount > 0 Then
        .movefirst
        Do While Not .EOF
          emrid = Trim(.fields("IncidentStatusID"))
          emrName = GetComboName("IncidentStatus", Trim(.fields("IncidentStatusID"))) '' & " -> " & GetComboName("WorkingYear", yr)
          
          If UCase(CStr(emrid)) = UCase(IncidentStatusID) Then
           dyHt = dyHt & "<option value=""" & CStr(emrid) & """ Selected >" & emrName & "</option>"
        Else
           dyHt = dyHt & "<option value=""" & CStr(emrid) & """ >" & emrName & "</option>"
        End If
        rst.MoveNext
            response.Flush
        Loop
      End If
      .Close
    End With
    dyHt = dyHt & "</select>"
    response.write dyHt
    Set rst = Nothing
End Sub

'getting the parts requested for repairs
Sub getAssetPartsList(rqtID)
 Dim rst, sql, count, total
   cnt = 0
   response.write "<td> "
    response.write " <table class ='table table-bordered ' style='font-size:12px'>"
    response.write "            <thead>"
    response.write "              <tr>"
    response.write "                <th>#</th>"
    response.write "                <th>Description</th>"
    response.write "                <th>Cost</th>"
    'response.write "                <th>Department</th>"

    response.write "              </tr>"
    response.write "            </thead>"
    Set rst = CreateObject("ADODB.Recordset")
    sql = "select Assetpartid,finalAmt from assetincidentpart WHERE assetincidentid = '" & rqtID & "'"
  With rst
        rst.open qryPro.FltQry(sql), conn, 3, 4

        If rst.RecordCount > 0 Then
            rst.movefirst
               Do While Not rst.EOF
               cnt = cnt + 1
               amt = rst.fields("finalAmt")
               total = total + amt
    response.write "    <tr>"
    response.write "       <td>" & cnt & "</td>"
    response.write "       <td><span>" & GetComboName("Assetpart", rst.fields("Assetpartid")) & "</span></td>"
    response.write "       <td><span>" & rst.fields("finalAmt") & "</span></td>"
    response.write "    </tr>"
    rst.MoveNext
    response.Flush
    Loop
    response.write "              <tr>"
     response.write "              <td colspan = '2'><b><span>Grand Total</span></b></td>"
      response.write "              <td><b><span>" & total & "</span></b></td>"
    response.write "              </tr>"
      End If
      .Close
    End With
    Set rst = Nothing
    response.write "</td> "
     response.write "</table>"
End Sub

'getting total Ambulance in the facility
Function getTotalAmbulance()
   Dim rst, sql, count, total
   ot = 0
   Set rst = CreateObject("ADODB.Recordset")
   sql = "SELECT COUNT(AssetPurchaseid) AS count FROM AssetPurchase WHERE AssetDepartmentid = 'A003' AND assettypeid = 'ITT003'"
    With rst
        rst.open qryPro.FltQry(sql), conn, 3, 4

        If rst.RecordCount > 0 Then
                ot = rst.fields("count")
         End If
      .Close
    End With
    Set rst = Nothing
   
        getTotalAmbulance = ot
End Function
'getting total vehicles in the facility
Function getTotalVehicle()
   Dim rst, sql, count, total
   ot = 0
   Set rst = CreateObject("ADODB.Recordset")
   sql = "SELECT COUNT(AssetPurchaseid) AS count FROM AssetPurchase WHERE AssetDepartmentid = 'A003' AND assettypeid <> 'ITT003'"
    With rst
        rst.open qryPro.FltQry(sql), conn, 3, 4

        If rst.RecordCount > 0 Then
                ot = rst.fields("count")
         End If
      .Close
    End With
    Set rst = Nothing
   
        getTotalVehicle = ot
End Function

Function getTotalVehicleByStatus(cmd, id)
   Dim rst, sql, count, total
   ot = 0
   Set rst = CreateObject("ADODB.Recordset")
   sql = "SELECT COUNT(AssetPurchaseid) AS count FROM AssetPurchase WHERE AssetDepartmentid = 'A003' AND assettypeid " & cmd & " 'ITT003'  and assetpurstatusid = '" & id & "'"
    
    With rst
        rst.open qryPro.FltQry(sql), conn, 3, 4

        If rst.RecordCount > 0 Then
                ot = rst.fields("count")
         End If
      .Close
    End With
    Set rst = Nothing
   
        getTotalVehicleByStatus = ot
End Function




'/////////////////////////////////////////////////////////////////////

Sub DisplayRequestedItems(reqID)
    Dim rst, sql
    Set rst = CreateObject("ADODB.Recordset")
    sql = "select dr.*, d.ItemName, u.UnitOfMeasureName from ItemRequestItems dr,Items d, UnitOfMeasure u Where d.ItemID=dr.ItemID And u.UnitOfMeasureID=d.UnitOfMeasureID "
    sql = sql & " And dr.ItemRequest2ID='" & reqID & "' order by d.ItemName "
    response.write "<table class='table table-responsive table-striped table-hover' style='font-size:12px'>"
    With rst
        rst.open qryPro.FltQry(sql), conn, 3, 4
        If rst.RecordCount > 0 Then
            requested = True
            vehicleNo = .fields("ItemID")
            rst.movefirst
            response.write "<thead><tr>"
            response.write "<th>#</th>"
            response.write "<th>Code</th>"
            response.write "<th>Item / Description</th>"
            'response.write "<td align='Right'><b>Qty (Req.)</b></td>"
           ' response.write "<td align='Right'><b>Qty (Appr.)</b></td>"
            'response.write "<th>UOM</th>"
            response.write "</tr></thead>"
            Do While Not rst.EOF
                response.write "<tr>"
                response.write "<td>" & rst.AbsolutePosition & "</td>"
                response.write "<td>" & rst.fields("ItemID") & "</td>"
                response.write "<td>" & rst.fields("ItemName") & "</td>"
                'response.write "<td align='Right'>" & FormatNumber(rst.fields("RequestValue1"), 1) & "</td>"
                'response.write "<td align='Right'>" & FormatNumber(rst.fields("RequestedQty"), 1) & "</td>"
                'response.write "<td>" & rst.fields("UnitOfMeasureName") & "</td>"
                response.write "</tr>"
                response.Flush
                rst.MoveNext
            Loop
        Else
          requested = False
        End If
        rst.Close
    End With
    response.write "</table>"
    Set rst = Nothing
End Sub

Sub DisplayApprovals(reqID)
    Dim rst, sql, apprLevel1, apprLevel2, lnkCnt
    Set rst = CreateObject("ADODB.Recordset")
    lnkCnt = 0
    apprLevel1 = False
    apprLevel2 = False
    available = apprLevel2
    approved = apprLevel1
    sql = "select * from ItemRequest2Pro Where TransProcessTblID='ItemRequest2Pro' And ItemRequest2ID='" & reqID & "' "
    sql = sql & " order by TransProcessDate1 "
    response.write "<table class='table table-responsive table-striped table-hover style='font-size:12px''>"
    With rst
        rst.open qryPro.FltQry(sql), conn, 3, 4
        If rst.RecordCount > 0 Then
            rst.movefirst
            Do While Not rst.EOF
           
                response.write "<tr>"
                If UCase(rst.fields("TransProcessVal2ID")) = UCase("ItemRequest2Pro-T002") Then
                    apprLevel1 = True
                    approved = apprLevel1
                End If
                If UCase(rst.fields("TransProcessVal2ID")) = UCase("ItemRequest2Pro-T003") Then
                    apprLevel1 = True
                    apprLevel2 = True
                    available = apprLevel2
                End If
                response.write "<td><p style='font-size:12px'>#" & rst.AbsolutePosition & ". <b>" & GetComboName("TransProcessVal2", rst.fields("TransProcessVal2ID")) & "</b><br>"
                response.write "By " & Glob_FormatName2(rst.fields("SystemUserID")) & "<br><em>[" & GetComboName("JobSchedule", rst.fields("JobScheduleID")) & "]</em>"
                response.write " on " & FormatDate(rst.fields("TransProcessDate1")) & "<br>"
                response.write "" & rst.fields("TransProcessDetail")
                response.write "</p></td>"
                response.write "</tr>"
                response.Flush
                rst.MoveNext
            Loop
        Else
            response.write "<tr><th colspan='100' style='color:red;font-style:italic;'>No Approval</th></tr>"
        End If

        sTb2 = "ItemRequest2Pro"
        If Not apprLevel1 Then
            response.write "<tr><td>"
            ' If HasAccessRight(uName, "frm" & sTb2, "New") Then
            
           
            If HasAccessRight(uName, "frm" & sTb2, "New") And (Glob_HasTransProcessAccess(sTb2, uName, "T001", "T002") Or Glob_HasTransProcessAccess(sTb2, uName, "T002", "T002")) Then
                'Clickable Url Link
                lnkCnt = Int((10000 * Rnd) + 1)
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = "<b>1. Authorize Req.</b>"
                lnkUrl = "wpg" & sTb2 & ".asp?PageMode=AddNew&TransProcessVal2ID=ItemRequest2Pro-T002&PullupData=ItemRequest2ID||" & reqID
                navPop = "POP"
                inout = "IN"
                fntSize = "8"
                fntColor = "#444488"
                bgColor = ""
                wdth = ""
                Glob_AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
            End If
            response.write "</td></tr>"
        End If
        If Not apprLevel2 Then
            response.write "<tr><td>"
            ' If HasAccessRight(uName, "frm" & sTb2, "New") Then
            If HasAccessRight(uName, "frm" & sTb2, "New") And (Glob_HasTransProcessAccess(sTb2, uName, "T001", "T012") Or Glob_HasTransProcessAccess(sTb2, uName, "T002", "T012")) Then
                'Clickable Url Link
                lnkCnt = Int((10000 * Rnd) + 1)
               'lnkCnt = "50"
                
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = "<b>2. Confirm Avail.</b>"
                lnkUrl = "wpg" & sTb2 & ".asp?PageMode=AddNew&TransProcessVal2ID=ItemRequest2Pro-T012&PullupData=ItemRequest2ID||" & reqID
                navPop = "POP"
                inout = "IN"
                fntSize = "8"
                fntColor = "#444488"
                bgColor = ""
                wdth = ""
                Glob_AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
            End If
            response.write "</td></tr>"
        End If
        rst.Close
    End With
    response.write "</table>"
    Set rst = Nothing
End Sub

Function DisplayIssuedItems(reqID)
    Dim rst, sql, ot
    Set rst = CreateObject("ADODB.Recordset")
    sql = "select di.*, d.ItemName from ItemIssueItems di, Items d where d.ItemID=di.ItemID and di.ItemRequest2ID='" & reqID & "' "
    sql = sql & " order by d.ItemName "
    ot = ""
    response.write "<table class='table table-responsive table-striped table-hover' style='font-size:12px'>"
    With rst
        rst.open qryPro.FltQry(sql), conn, 3, 4
        If rst.RecordCount > 0 Then
            issued = True
            rst.movefirst
            response.write "<thead><tr><th colspan='100'>"
            response.write "Issued By: " & Glob_FormatName2(rst.fields("SystemUserID")) & " @ " & FormatDate(rst.fields("IssuedDate1"))
            response.write "</th></tr></thead>"

            response.write "<thead><tr>"
            response.write "<th>#</th>"
            response.write "<th>Code</th>"
            response.write "<th>Vehicle / Desc</th>"
            response.write "<td align='Right'><b>Qty</b></td>"
            response.write "</tr></thead>"
            Do While Not rst.EOF
                response.write "<tr>"
                response.write "<td>" & rst.AbsolutePosition & "</td>"
                response.write "<td>" & rst.fields("ItemID") & "</td>"
                response.write "<td>" & rst.fields("ItemName") & "</td>"
                response.write "<td align='Right'>" & FormatNumber(rst.fields("IssuedQty"), 1) & "</td>"
                response.write "</tr>"
                ot = rst.fields("ItemIssueID")
                response.Flush
                rst.MoveNext
            Loop
        Else
            issued = False
            response.write "<tr><th colspan='100'>No Issued Items</th></tr>"
            response.write "<tr><td>"
            sTb2 = "ItemIssue"
            If HasAccessRight(uName, "frm" & sTb2, "New") Then
                lnkCnt = Int((1000 * Rnd) + 1)
                lnkID = "lnk" & CStr(lnkCnt)
                lnkText = "<b>&nbsp;&nbsp;Issue Request</b>"
                lnkUrl = "wpg" & sTb2 & ".asp?PageMode=AddNew&PullupData=ItemRequest2ID||" & reqID
                navPop = "POP"
                inout = "IN"
                fntSize = "10"
                fntColor = "#444488"
                bgColor = ""
                wdth = ""
                Glob_AddUrlLink lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth
            End If
            response.write "</td></tr>"
        End If
        rst.Close
    End With
    response.write "</table>"
    Set rst = Nothing
    DisplayIssuedItems = ot
End Function

Sub DisplayCompletedTrip(reqID)
    Dim rst, sql
    Set rst = CreateObject("ADODB.Recordset")
    sql = "SELECT PreviousRemarks From ItemRequestItems  WHERE PreviousRemarks LIKE '%completed%' AND ItemRequest2id = '" & reqID & "'"


    response.write "<table class='table table-responsive table-striped table-hover'>"
    With rst
        rst.open qryPro.FltQry(sql), conn, 3, 4
        If rst.RecordCount > 0 Then
        rmk = rst.fields("PreviousRemarks")
        
        rmkArr = Split(rmk, "||")
        sts = rmkArr(0)
        usName = rmkArr(1)
        dt = rmkArr(2)
        response.write "<tr>"
           response.write "<td  style='font-size:16px;color:Green' font-weight:bold; text-align:center><span style='font-size:22px;font-weight:bolder'>Completed</span><br> By: " & usName & " <br> Date: " & dt & "</td>"
         response.write "</tr>"
        Else
             response.write "<tr><th colspan='100' style='color:red;font-style:italic;'>Trip Not Completed</th></tr>"
           
            sTb2 = "ItemAccept"
            If HasAccessRight(uName, "frm" & sTb2, "New") Then

                response.write "<tr><td>"
               If (jschd = "M10A") Then
                response.write " <div class='float-end mr-2'style='margin-left:8px'> "
               response.write "  <a  onclick='DoAjaxUpdate(""" & reqID & """)' class='btn btn-outline-success float-end mr-2'>"
                response.write "    <i class=""fas fa-edit""></i> Complete"
                response.write "  </a>"
                response.write " </div> "
                End If
                 response.write "</td></tr>"
            End If
           
        End If
        rst.Close
    End With
    response.write "</table>"
    Set rst = Nothing
End Sub


Sub AddUrlLink(lnkID, lnkText, lnkUrl, navPop, inout, fntSize, fntColor, bgColor, wdth)
  Dim plusMinus, imgName, lnkOpClNavPop, align
   plusMinus = ""
   imgName = ""
   align = ""
   lnkOpClNavPop = inout & "||" & navPop & "||800||600||CLOSE"
  AddPrtNavLink lnkID, plusMinus, imgName, lnkText, lnkUrl, lnkOpClNavPop, fntSize, fntColor, bgColor, align, wdth
End Sub

'ExtractDates
Sub ExtractDates(inFlt, outDt1, outDt2)
  Dim arr, ul, num, dat1, dat2
  dat1 = ""
  dat2 = ""
  arr = Split(inFlt, "||")
  ul = UBound(arr)
  If ul >= 0 Then
    For num = 0 To ul
      If num = 0 Then
        dat1 = Trim(arr(0))
      ElseIf num = 1 Then
        dat2 = Trim(arr(1))
      End If
    Next
    If IsDate(dat1) Then
      If IsDate(dat2) Then
      Else 'No Dat2
        dat2 = FormatDate(CDate(dat1)) & " 23:59:59"
        dat1 = FormatDate(CDate(dat1)) & " 00:00:00"
      End If
    Else 'No Dat1
      If IsDate(dat2) Then
        dat1 = FormatDate(CDate(dat2)) & " 0:00:00"
        dat2 = FormatDate(CDate(dat2)) & " 23:59:59"
      Else 'No Dat2
      End If
    End If
  End If
  outDt1 = dat1
  outDt2 = dat2
End Sub

Function ExtractWorkingDate(wkDay)
    Dim str
    ExtractWorkingDate = Null
    str = Trim(wkDay)
    If Len(str) = 11 Then
      If UCase(Left(str, 3)) = "DAY" Then
        ExtractWorkingDate = CDate(mid(str, 10, 2) & " " & monthName(CInt(mid(str, 8, 2)), 1) & " " & mid(str, 4, 4))
      End If
    End If
End Function

Function HasPrintOutAccess(jb, prt)
  Dim rstTblSql, sql, ot
  ot = False
  Set rstTblSql = CreateObject("ADODB.Recordset")
  With rstTblSql
    sql = "select JobScheduleID from printoutalloc "
    sql = sql & " where printlayoutid='" & prt & "' and jobscheduleid='" & jb & "'"
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .movefirst
      ot = True
    End If
    .Close
  End With
  HasPrintOutAccess = ot
  Set rstTblSql = Nothing
End Function

Function HasModuleMgrAccess(jb, tb)
  Dim rstTblSql, sql, ot
  ot = ""
  Set rstTblSql = CreateObject("ADODB.Recordset")
  With rstTblSql
    sql = "select ModuleManagerID from ModuleManageralloc "
    sql = sql & " where tableid='" & tb & "' and jobscheduleid='" & jb & "' order by ModuleManagerID"
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .movefirst
      ot = .fields("ModuleManagerID")
    End If
    .Close
  End With
  HasModuleMgrAccess = ot
  Set rstTblSql = Nothing
End Function

Sub SetListWhCls2()
  Dim jb
  jb = Trim(jschd)
  dispTyp = GetDispType2(jb)

  ' If dispTyp = "LAB" Then
  '   lstWhCls2 = " and (LabByDoctor.TestCategoryID='B13' Or LabByDoctor.TestGroupID='B13')"
  ' ElseIf dispTyp = "IMAGING" Then
  '   lstWhCls2 = " and (LabByDoctor.TestCategoryID='B19' Or LabByDoctor.TestGroupID='B19')"
  ' End If
  ' lstWhCls2 = lstWhCls2 & " And LabByDoctor.WorkingDayID >= 'DAY20220501' "
End Sub


Function GetDispType2(jb)
  Dim ot
  ot = ""
  If UCase(Left(jb, 3)) = "M07" Then
    ot = "IMAGING"
  ElseIf UCase(Left(jb, 3)) = "M05" Then
    ot = "LAB"
  End If
  GetDispType2 = ot
End Function

Sub SetRequisitionMonth(store, currMth)
    Set rst = CreateObject("ADODB.Recordset")
    dyHt = "<select size=""1"" name=""NoOfDays"" id=""NoOfDays"" onchange=""NoOfDaysOnchange()"">"
    dyHt = dyHt & "<option value=""""></option>"

    cYr = ""
    yr = ""
    dtDply = CDate(Glob_DeploymentDate())
    lstWhCls2 = " And dr.WorkingMonthID>='" & FormatWorkingMonth(dtDply) & "' "
    sql0 = "select distinct dr.WorkingMonthID, wm.WorkingMonthName, wm.WorkingYearID from ItemRequest2 dr, WorkingMonth wm where dr.WorkingMonthID=wm.WorkingMonthID "
    sql0 = sql0 & " And dr.BranchID='" & brnch & "' " & lstWhCls2
    ' If Trim(store)<>"" Or Not Glob_HasTransProcessAccess2("ItemRequest2Pro", uName) Then
    If Not (Glob_HasStaffLevel(uName) Or Glob_HasTransProcessAccess2("ItemRequest2Pro", uName)) Then
        sql = sql & " And (ItemRequest2.ItemStoreID='" & store & "' Or ItemRequest2.ItemRequestStoreID='" & store & "')  "
    End If
    sql0 = sql0 & " order by dr.WorkingMonthID desc"
    '' response.write sql0

    With rst
      .open qryPro.FltQry(sql0), conn, 3, 4
      If .RecordCount > 0 Then
        .movefirst
        Do While Not .EOF
          wkMth = Trim(.fields("WorkingMonthID"))
          dyNm = Trim(.fields("WorkingMonthName")) '' & " -> " & GetComboName("WorkingYear", yr)
          yr = Trim(.fields("WorkingYearID"))
          If UCase(cYr) <> UCase(yr) Then
             dyHt = dyHt & "<optGroup label=""" & GetComboName("WorkingYear", yr) & """>"
             cYr = yr
          End If
          If UCase(CStr(currMth)) = UCase(wkMth) Then
            dyHt = dyHt & "<option value=""" & CStr(wkMth) & """ selected>" & dyNm & "</option>"
          Else
             dyHt = dyHt & "<option value=""" & CStr(wkMth) & """>" & dyNm & "</option>"
          End If
          .MoveNext
        Loop
      End If
      .Close
    End With
    dyHt = dyHt & "</select>"
    response.write dyHt
End Sub

Sub SetMedicalService(br, vDt1, vDt2, currMs)
  Dim rs, ot, sql, sp
  Set rs = CreateObject("ADODB.Recordset")
  sql = "select distinct MedicalServiceID from Visitation where BranchID='" & br & "' and VisitDate between '" & vDt1 & "' and '" & vDt2 & "'"
  sql = sql & " And MedicalServiceID IN ('M001','M003') order by MedicalServiceID"

  ot = "<select id=""MedicalService"" name=""MedicalService"" onchange=""MedicalServiceOnchange()"">"
  ot = ot & "<option></option>"
  With rs
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .movefirst
      Do While Not .EOF
        sp = .fields("MedicalServiceID")
        If UCase(sp) = UCase(currMs) Then
          ot = ot & "<option value=""" & sp & """ selected>" & GetComboName("MedicalService", sp) & "</option>"
        Else
          ot = ot & "<option value=""" & sp & """>" & GetComboName("MedicalService", sp) & "</option>"
        End If
        .MoveNext
      Loop
    End If
    .Close
  End With
  ot = ot & "</select>"
  response.write ot
  Set rs = Nothing
End Sub

Sub SetItemStore(br, jb, currSto)
  Dim rs, ot, sql, dSto
  Set rs = CreateObject("ADODB.Recordset")
  ' sql = "select  distinct ds.ItemStoreID, ds.ItemStoreName from ItemStore ds, ItemStore2 ds2 where ds.JobScheduleID=ds2.JobScheduleID "
  ' sql = sql & " And ds.BranchID='" & br & "' and ds.ItemStoreID IN ('M0601','M0602','M0603','M0604','M0605','M0612') "
  ' sql = sql & " order by ds.ItemStoreID "

  sql = "select distinct ds.ItemStoreID, ds.ItemStoreName from ItemStore ds Where ds.BranchID='" & br & "' "
  sql = sql & " And ds.ItemStoreID IN ('M0601','M0602','M0603','M0604','M0605','M0612') "
  ' sql = sql & " UNION "
  ' sql = sql & " select  distinct ds.ItemStoreID, ds.ItemStoreName from ItemStore ds, ItemStore2 ds2 where ds.JobScheduleID=ds2.JobScheduleID "
  ' sql = sql & " And ds.BranchID='" & br & "' and ds.ItemStoreID IN ('M0601','M0602','M0603','M0604','M0605','M0612') "
  sql = sql & " order by ds.ItemStoreID "

  ot = "<select id=""ItemStore"" name=""ItemStore"" onchange=""ItemStoreOnchange()"">"
  ot = ot & "<option></option>"
  With rs
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .movefirst
      Do While Not .EOF
        dSto = .fields("ItemStoreID")
        If UCase(dSto) = UCase(currSto) Then
          ot = ot & "<option value=""" & dSto & """ selected>" & .fields("ItemStoreName") & "</option>"
        Else
          ot = ot & "<option value=""" & dSto & """>" & .fields("ItemStoreName") & "</option>"
        End If
        .MoveNext
      Loop
    End If
    .Close
  End With
  ot = ot & "</select>"
  response.write ot
  Set rs = Nothing
End Sub

Sub SetItemStoreIC(br, jb, currSto)
  Dim rs, ot, sql, dSto
  Set rs = CreateObject("ADODB.Recordset")
  sql = "select distinct js.JobScheduleID, js.JobScheduleName from JobSchedule js "
  sql = sql & " Where js.JobScheduleID IN (SELECT JobScheduleID From ItemStore2 Where JobScheduleID IN   "
  sql = sql & "  ('" & jb & "','M06IC','M0601IC','M0602IC','M0603IC','M0604IC','M0605IC','M0612IC') and BranchID='" & br & "'  "
  sql = sql & " )  "
  sql = sql & " order BY js.Jobscheduleid "

  ot = "<select id=""ItemStore"" name=""ItemStore"" onchange=""ItemStoreOnchange()"">"
  ot = ot & "<option></option>"
  With rs
    .open qryPro.FltQry(sql), conn, 3, 4
    If .RecordCount > 0 Then
      .movefirst
      Do While Not .EOF
        dSto = .fields("JobScheduleID")
        ' If UCase(dSto) = UCase(currSto) Then
        If UCase(dSto) = UCase(jb) Then
          ot = ot & "<option value=""" & dSto & """ selected>" & .fields("JobScheduleName") & "</option>"
        Else
          ot = ot & "<option value=""" & dSto & """>" & .fields("JobScheduleName") & "</option>"
        End If
        .MoveNext
      Loop
    End If
    .Close
  End With
  ot = ot & "</select>"
  response.write ot
  Set rs = Nothing
End Sub

Function GetItemStore(jb)
    Dim ot
    Set rst = CreateObject("ADODB.Recordset")
    ot = GetComboNameFld("ItemStore", jb, "JobScheduleID")
    If Len(Trim(ot)) > 0 Then
        ot = Trim(ot)
    Else
        ot = ""
        sql = "select top 1 * from ItemStore2 Where JobScheduleID='" & jb & "' "
        With rst
            rst.open qryPro.FltQry(sql), conn, 3, 4
            If rst.RecordCount > 0 Then
                rst.movefirst
                ot = rst.fields("ItemStoreID")
            End If
            rst.Close
        End With
    End If
    Set rst = Nothing
    GetItemStore = ot
End Function


Sub LoadCSS()
  Dim str
  str = ""
  str = str & "<style type='text/css' id=""styPrt"">"
  str = str & ".cpHdrTd{font-size:14pt;font-weight:bold}"
  str = str & ".cpHdrTr{background-color:#eeeeee}"
  str = str & ".cpHdrTd2{font-size:12pt;font-weight:bold}"
  str = str & ".cpHdrTr2{background-color:#eeeeee}" 'fafafa
  str = str & ".table{font-size:14px;}"
  str = str & "</style>"
  response.write str

  response.write "<style>"
  response.write ".cmpTdSty {"
  response.write "border:1px solid #d0d0d0;"
  response.write "border-collapse: collapse;"
  response.write "}"
  response.write "</style>"
End Sub





If (jschd = "M10A") Then
SetVehicleAlerts
'SetPurchaseOrderAlerts
End If

Sub SetVehicleAlerts()
    Set rst = CreateObject("ADODB.Recordset")
    response.Flush
 
    
    dtNow = Now()
    'Recent Vehicle Request
    minsAgo = 90
    vDt1 = FormatDateDetail(DateAdd("n", (-1 * minsAgo), dtNow))
    vDt2 = FormatDateDetail(dtNow)
    sql = " SELECT SUM(ItemRequestCount) AS count FROM ("
    sql = sql & " SELECT COUNT(ItemRequest2id) AS ItemRequestCount From ItemRequest2Pro "
    sql = sql & "   WHERE itemrequeststoreid = 'M10A'  GROUP BY ItemRequest2id"
    sql = sql & "  Having count(ItemRequest2id) = 1  ) AS SubqueryResult;"
    

'    response.write Glob_GetBootstrapToastAlertHeader("")


    With rst
        rst.open qryPro.FltQry(sql), conn, 3, 4
        cnt = rst.fields("count")
        If cnt > 0 Then
            rst.movefirst
             response.write Glob_GetBootstrapToastAlertHeader("")
             Do While Not rst.EOF
                  Set tOption = server.CreateObject("Scripting.Dictionary")
                  alertText = cnt & " Recent Vehicle Request " '& GetComboName("Ward", rst.fields("WardID"))
'                   tOption.Add "close", False
                  tOption.Add "close", True
                  tOption.Add "icon", True
                  tOption.Add "delay", 60 * 2 ''total seconds

                  tOption.Add "title", "New Vehicle Request "
                  tOption.Add "subtitle", "Within last " & minsAgo & " minutes"
'                  tOption.Add "button1", "View List"
                  tOption.Add "button1Url", "wpgPrtPrintLayoutAll.asp?PrintLayoutName=MaintenanceDetails&PositionForTableName=WorkingDay&MedicalService=&WorkingDayID=DAY20160401&NoOfDays=&OrderByType=&Specialist=&TransProcessVal=AdmissionPro-T013&DisplayType=Prescription"
                   tOption.Add "button2", "See User Details"
                   tOption.Add "button2Url", lnkUrl
                  lnkCnt = lnkCnt + 1
                  response.write Glob_GetBootstrapToastAlert("Danger", alertText, tOption, lnkCnt)
              response.Flush
              Set tOption = Nothing
               rst.MoveNext
             Loop
             response.write Glob_GetBootstrapToastAlertFooter()
'             Else
'             response.write Glob_GetBootstrapToastAlertFooter()
        End If
        rst.Close
    End With
    
    
    dtNow = Now()
    ''Recent Fuel
    minsAgo = 90
    vDt1 = FormatDateDetail(DateAdd("n", (-1 * minsAgo), dtNow))
    vDt2 = FormatDateDetail(dtNow)
    sql = "SELECT COUNT(AssetmaintainID) as count FROM AssetMaintain"
    sql = sql & " Where 1=1"
    sql = sql & " AND DateDiff(day, getDate(), endDate) <= 14 And DateDiff(day, getDate(), endDate) > -2"
    sql = sql & " AND MaintainCategoryID = 'M008'"

    With rst
        rst.open qryPro.FltQry(sql), conn, 3, 4
        cnt = rst.fields("count")
        If cnt > 0 Then

            rst.movefirst
             response.write Glob_GetBootstrapToastAlertHeader("")
            ' Do While Not rst.EOF
                  Set tOption = server.CreateObject("Scripting.Dictionary")
                  alertText = cnt & " Recent Fuel Alert " '& GetComboName("Ward", rst.fields("WardID"))
                  ' tOption.Add "close", False
                  tOption.Add "close", True
                  tOption.Add "icon", True
                  tOption.Add "delay", 60 * 2 ''total seconds

                  tOption.Add "title", "New Maintainance Alert "
                  tOption.Add "subtitle", "Within last " & minsAgo & " minutes"
                  tOption.Add "button1", "See List"
                  tOption.Add "button1Url", "wpgPrtPrintLayoutAll.asp?PrintLayoutName=MaintenanceDetails&PositionForTableName=WorkingDay&MedicalService=&WorkingDayID=DAY20160401&NoOfDays=&OrderByType=&Specialist=&TransProcessVal=AdmissionPro-T013&DisplayType=Prescription&MaintainCategoryID=" & M008 & ""
                   tOption.Add "button2", "See Details"
                   tOption.Add "button2Url", lnkUrl
                  lnkCnt = lnkCnt + 1
                  response.write Glob_GetBootstrapToastAlert("Danger", alertText, tOption, lnkCnt)
              response.Flush
              Set tOption = Nothing
              ' rst.MoveNext
            ' Loop
             'response.write Glob_GetBootstrapToastAlertFooter()
             Else
             'response.write Glob_GetBootstrapToastAlertFooter()
        End If
        rst.Close
    End With
    
    
    dtNow = Now()
    ''Recent Roadworthy
    minsAgo = 90
    vDt1 = FormatDateDetail(DateAdd("n", (-1 * minsAgo), dtNow))
    vDt2 = FormatDateDetail(dtNow)
    sql = "SELECT COUNT(AssetmaintainID) as count FROM AssetMaintain"
    sql = sql & " Where 1=1"
    sql = sql & " AND DateDiff(day, getDate(), endDate) <= 14 And DateDiff(day, getDate(), endDate) > -2"
    sql = sql & " AND MaintainCategoryID = 'M009'"

    With rst
        rst.open qryPro.FltQry(sql), conn, 3, 4
        cnt = rst.fields("count")
        If cnt > 0 Then
          
            rst.movefirst
'             response.write Glob_GetBootstrapToastAlertHeader("")
            ' Do While Not rst.EOF
                  Set tOption = server.CreateObject("Scripting.Dictionary")
                  alertText = cnt & " Recent Roadworthy Alert " '& GetComboName("Ward", rst.fields("WardID"))
                  ' tOption.Add "close", False
                  tOption.Add "close", True
                  tOption.Add "icon", True
                  tOption.Add "delay", 60 * 2 ''total seconds

                  tOption.Add "title", "New Maintainance Alert "
                  tOption.Add "subtitle", "Within last " & minsAgo & " minutes"
                  tOption.Add "button1", "See List"
                  tOption.Add "button1Url", "wpgPrtPrintLayoutAll.asp?PrintLayoutName=MaintenanceDetails&PositionForTableName=WorkingDay&MedicalService=&WorkingDayID=DAY20160401&NoOfDays=&OrderByType=&Specialist=&TransProcessVal=AdmissionPro-T013&DisplayType=Prescription&MaintainCategoryID=" & M009 & ""
                   tOption.Add "button2", "See Details"
                   tOption.Add "button2Url", lnkUrl
                  lnkCnt = lnkCnt + 1
                  response.write Glob_GetBootstrapToastAlert("Danger", alertText, tOption, lnkCnt)
              response.Flush
              Set tOption = Nothing
              ' rst.MoveNext
            ' Loop
             response.write Glob_GetBootstrapToastAlertFooter()
'             Else
             'response.write Glob_GetBootstrapToastAlertFooter()
        End If
        rst.Close
    End With
    
    
    dtNow = Now()
    ''Recent Maintainance
    minsAgo = 90
    vDt1 = FormatDateDetail(DateAdd("n", (-1 * minsAgo), dtNow))
    vDt2 = FormatDateDetail(dtNow)
    sql = "SELECT COUNT(AssetmaintainID) as count FROM AssetMaintain"
    sql = sql & " Where 1=1"
    sql = sql & " AND DateDiff(day, getDate(), endDate) <= 14 And DateDiff(day, getDate(), endDate) > -2"
    sql = sql & " AND MaintainCategoryID = 'M002'"

    With rst
        rst.open qryPro.FltQry(sql), conn, 3, 4
        cnt = rst.fields("count")
        If cnt > 0 Then
          
            rst.movefirst
'             response.write Glob_GetBootstrapToastAlertHeader("")
            ' Do While Not rst.EOF
                  Set tOption = server.CreateObject("Scripting.Dictionary")
                  alertText = cnt & " Recent Maintainance Alert " '& GetComboName("Ward", rst.fields("WardID"))
                  ' tOption.Add "close", False
                  tOption.Add "close", True
                  tOption.Add "icon", True
                  tOption.Add "delay", 60 * 2 ''total seconds

                  tOption.Add "title", "New Maintainance Alert "
                  tOption.Add "subtitle", "Within last " & minsAgo & " minutes"
                  tOption.Add "button1", "See List"
                  tOption.Add "button1Url", "wpgPrtPrintLayoutAll.asp?PrintLayoutName=MaintenanceDetails&PositionForTableName=WorkingDay&MedicalService=&WorkingDayID=DAY20160401&NoOfDays=&OrderByType=&Specialist=&TransProcessVal=AdmissionPro-T013&DisplayType=Prescription&MaintainCategoryID=" & M002 & ""
                   tOption.Add "button2", "See Details"
                   tOption.Add "button2Url", lnkUrl
                  lnkCnt = lnkCnt + 1
                  response.write Glob_GetBootstrapToastAlert("Danger", alertText, tOption, lnkCnt)
              response.Flush
              Set tOption = Nothing
              ' rst.MoveNext
            ' Loop
             response.write Glob_GetBootstrapToastAlertFooter()
             'Else
             'response.write Glob_GetBootstrapToastAlertFooter()
        End If
        rst.Close
    End With

     
    dtNow = Now()
    wkDay = FormatWorkingDay(dtNow)
    ''Recent Prescriptions
    minsAgo = 2
    vDt1 = FormatDateDetail(DateAdd("n", (-1 * minsAgo), dtNow))
    vDt2 = FormatDateDetail(dtNow)
    sql = " SELECT COUNT(IncomingDrugID) as count FROM IncomingDrug"
     sql = sql & " WHERE workingdayid = '" & wkDay & "'"
    sql = sql & " AND DATEDIFF(HOUR, entrydate, GETDATE()) <= 2 "
    'response.write Glob_GetBootstrapToastAlertHeader("")

    'response.write sql


    With rst
        rst.open qryPro.FltQry(sql), conn, 3, 4
        cnt = rst.fields("count")

        If cnt > 0 Then
            ' response.write Glob_GetBootstrapToastAlertHeader("")
            ' Do While Not rst.EOF
                  Set tOption = server.CreateObject("Scripting.Dictionary")
                  alertText = cnt & " Recent Items Arrival " '& GetComboName("Ward", rst.fields("WardID"))
                  ' tOption.Add "close", False
                  tOption.Add "close", True
                  tOption.Add "icon", True
                  tOption.Add "delay", 60 * 2 ''total seconds

                  tOption.Add "title", "New Purchase Order "
                  tOption.Add "subtitle", "Within last " & minsAgo & " Hours"
                  tOption.Add "button1", "See List"
                  tOption.Add "button1Url", "wpgPrtPrintLayoutAll.asp?PrintLayoutName=PurchaseOrderList&PositionForTableName=WorkingDay&MedicalService=&WorkingDayID=DAY20160401&NoOfDays=&OrderByType=&Specialist=&TransProcessVal=AdmissionPro-T013&DisplayType=Prescription"
                  ' tOption.Add "button2", "See User Details"
                  ' tOption.Add "button2Url", lnkUrl
                  lnkCnt = lnkCnt + 1
                  response.write Glob_GetBootstrapToastAlert("warning", alertText, tOption, lnkCnt)
              response.Flush
              Set tOption = Nothing
              ' rst.MoveNext
            ' Loop
            response.write Glob_GetBootstrapToastAlertFooter()
            Else
            response.write Glob_GetBootstrapToastAlertFooter()
        End If
        rst.Close
    End With
 

End Sub








Sub InitPageScript()
  SetPageVariable "AutoHidePrintControl", "Yes"

 Dim htStr
  'Client Script
  htStr = ""
  htStr = htStr & "<script id=""scptPrintLayoutExtraScript"" LANGUAGE=""javascript"">" & vbCrLf
  htStr = htStr & vbCrLf
  'RefreshPage()
  htStr = htStr & "function RefreshPage(){" & vbCrLf
  htStr = htStr & "window.location.reload();" & vbCrLf
  htStr = htStr & "}" & vbCrLf


  htStr = htStr & "$(document).ready(function () {" & vbCrLf
  htStr = htStr & "  $('.nav-link').click(function (e) {" & vbCrLf
  htStr = htStr & "    e.preventDefault();" & vbCrLf
  htStr = htStr & " $('.nav-link').removeClass('active');"
  htStr = htStr & "    $(this).addClass('active');"
  htStr = htStr & "    var contentId = $(this).data('content');" & vbCrLf
  htStr = htStr & "    $('.content').hide();" & vbCrLf
  htStr = htStr & "    $('#' + contentId).show();" & vbCrLf
  htStr = htStr & "    // Store the selected content in local storage" & vbCrLf
  htStr = htStr & "    localStorage.setItem('selectedContent', contentId);" & vbCrLf
  htStr = htStr & "  });" & vbCrLf
  htStr = htStr & "  var selectedContent = localStorage.getItem('selectedContent');" & vbCrLf
  htStr = htStr & "  if (selectedContent) {" & vbCrLf
  htStr = htStr & "    $('#' + selectedContent).show();" & vbCrLf
  
   htStr = htStr & "  $('.nav-link[data-content=' + selectedContent + ']').addClass('active');" & vbCrLf
   'Add the "active" class to the corresponding link
      
  htStr = htStr & "  }" & vbCrLf
  htStr = htStr & "});" & vbCrLf
  
    htStr = htStr & " function openPopup(anc){" & vbCrLf
    htStr = htStr & "     let win=window.open(anc.dataset.href, '_blank', 'resizeable=yes,scrollbars=yes,width=820,height=560,status=yes  ');" & vbCrLf
    htStr = htStr & "     "
    htStr = htStr & "     let intvl = setInterval(function(){" & vbCrLf
    htStr = htStr & "         if(win.closed !== false){" & vbCrLf
    htStr = htStr & "             clearInterval(intvl);" & vbCrLf
    htStr = htStr & "             window.location.reload();" & vbCrLf
    htStr = htStr & "          }" & vbCrLf
    htStr = htStr & "     }, 200);" & vbCrLf
    htStr = htStr & "}" & vbCrLf
    
   
    htStr = htStr & "$(document).ready(function(){" & vbCrLf
    htStr = htStr & "    var table = $('#VehicleTable').DataTable({" & vbCrLf
    htStr = htStr & "        buttons:['copy', 'csv', 'excel', 'pdf', 'print']" & vbCrLf
    htStr = htStr & "    });" & vbCrLf
    htStr = htStr & "    table.buttons().container()" & vbCrLf
    htStr = htStr & "    .appendTo('#vehicleTable_wrapper .col-md-6:eq(0)');" & vbCrLf
    htStr = htStr & "});" & vbCrLf
    
  'NoOfDaysOnchange
  htStr = htStr & "function NoOfDaysOnchange(){" & vbCrLf
  htStr = htStr & "var ur,dy,sp,ordByTyp,ms;" & vbCrLf
  htStr = htStr & "dy=GetEleVal('NoOfDays');" & vbCrLf
  'htStr = htStr & "sp=GetEleVal('Specialist');" & vbCrLf
 'htStr = htStr & "ms=GetEleVal('MedicalService');" & vbCrLf
  'htStr = htStr & "ordByTyp=GetCheckedRadio('inpOrderByType');" & vbCrLf
  htStr = htStr & "ur='wpgPrtPrintLayoutAll.asp?PrintLayoutName=VehicleList&PositionForTableName=WorkingDay';" & vbCrLf
  htStr = htStr & "ur=ur + '&WorkingDayID=DAY20160401&NoOfDays=' + dy ;" & vbCrLf
  htStr = htStr & "window.location.href=processurl(ur);" & vbCrLf
  htStr = htStr & "}" & vbCrLf
    
  htStr = htStr & "function PLExtraScriptOnLoad(){" & vbCrLf
  htStr = htStr & "window.onresize=windowOnresize;" & vbCrLf
  htStr = htStr & "HideEle(""trPrintControl"");" & vbCrLf
  htStr = htStr & "windowOnresize();" & vbCrLf
  htStr = htStr & "}" & vbCrLf

  htStr = htStr & "function windowOnresize(){" & vbCrLf
  htStr = htStr & " var ht,ele;" & vbCrLf
  htStr = htStr & " ht=window.innerHeight;" & vbCrLf
  htStr = htStr & " if (Helpers.isnumeric(ht)){" & vbCrLf
  htStr = htStr & "ele = document.getElementById('iFrm1');" & vbCrLf

  If UCase(fullScrn) = "NO" Then 'No Full Screen
  htStr = htStr & "if (ele) {" & vbCrLf
  htStr = htStr & " ele.height=Helpers.cstr(Helpers.cint(ht)-80);" & vbCrLf
  htStr = htStr & "}" & vbCrLf
  htStr = htStr & "ele = document.getElementById('iFrm2');" & vbCrLf
  htStr = htStr & "if (ele) {" & vbCrLf
  htStr = htStr & " ele.height=Helpers.cstr(Helpers.cint(ht)-90);" & vbCrLf
  htStr = htStr & "}" & vbCrLf
  Else
  htStr = htStr & "if (ele) {" & vbCrLf
  htStr = htStr & " ele.height=Helpers.cstr(Helpers.cint(ht));" & vbCrLf
  htStr = htStr & "}" & vbCrLf
  htStr = htStr & "ele = document.getElementById('iFrm2');" & vbCrLf
  htStr = htStr & "if (ele) {" & vbCrLf
  htStr = htStr & " ele.height=Helpers.cstr(Helpers.cint(ht));" & vbCrLf
  htStr = htStr & "}" & vbCrLf
  End If
  htStr = htStr & "}" & vbCrLf
  htStr = htStr & "}" & vbCrLf
  
  htStr = htStr & "function formatwinposprt(wd, ht) {" & vbCrLf
  htStr = htStr & "var lft, tp;" & vbCrLf
  htStr = htStr & "var ot;" & vbCrLf
  htStr = htStr & "lft = Helpers.cstr((screen.availWidth - Helpers.cint(wd)) / 2);" & vbCrLf
  htStr = htStr & "tp = Helpers.cstr((screen.availHeight - Helpers.cint(ht)) / 2);" & vbCrLf
  htStr = htStr & "if (Helpers.cint(lft)<0){" & vbCrLf
  htStr = htStr & "lft=""0""" & vbCrLf
  htStr = htStr & "}" & vbCrLf
  htStr = htStr & "if (Helpers.cint(tp)<0){" & vbCrLf
  htStr = htStr & "tp=""0""" & vbCrLf
  htStr = htStr & "}" & vbCrLf
  htStr = htStr & "ot = ""top="" + tp + "",left="" + lft + "",height="" + ht + "",width="" + wd + "",status=no,toolbar=no,menubar=no,location=no,resizable=yes,scrollbars=yes"";" & vbCrLf
  htStr = htStr & "return ot;" & vbCrLf
  htStr = htStr & "}" & vbCrLf
  
  
  
  'function to update vehicle status to complete

    htStr = htStr & "function DoAjaxUpdate(id) {" & vbCrLf
    htStr = htStr & "  if (confirm('Are you sure you want to update?')) {" & vbCrLf
    htStr = htStr & "    UpdateTripStatus(id);" & vbCrLf
    htStr = htStr & "  } else {" & vbCrLf
    htStr = htStr & "    alert('Update canceled.');" & vbCrLf
    htStr = htStr & "  }" & vbCrLf
    htStr = htStr & "}" & vbCrLf
    htStr = htStr & "function UpdateTripStatus(id) {" & vbCrLf
    htStr = htStr & "  let url, getStr;" & vbCrLf
    htStr = htStr & "  if (Helpers.len(id) > 0) {" & vbCrLf
    htStr = htStr & "    getStr = 'ProcedureName=completeTrip&reqID=' + id;" & vbCrLf
    htStr = htStr & "    url = 'wpgXmlHttp.' + appfilext + '?' + getStr;" & vbCrLf
    htStr = htStr & "    console.log(url);" & vbCrLf
    htStr = htStr & "        alert('Trip Completed successful!');"
    htStr = htStr & "    Helpers.xmlhttprequest(Helpers.ucase('GET'), url, UpdateTripStatusCount);" & vbCrLf
    htStr = htStr & "  } else {" & vbCrLf
    htStr = htStr & "    alert('Please Select a Trip!!');" & vbCrLf
    htStr = htStr & "  }" & vbCrLf
    htStr = htStr & "}" & vbCrLf
    htStr = htStr & "function UpdateTripStatusCount(readyState, responseText) {" & vbCrLf
    htStr = htStr & "  var arr, ul, num, str, rec;" & vbCrLf
    htStr = htStr & "  if (readyState == 4) {" & vbCrLf
    htStr = htStr & "    str = ReplaceXmlHttpComment(responseText);" & vbCrLf
    htStr = htStr & "    arr = Split(str, delim(1));" & vbCrLf
    htStr = htStr & "    ul = UBound(arr);" & vbCrLf
    htStr = htStr & "    if (ul >= 0) {" & vbCrLf
    htStr = htStr & "      rec = Helpers.trim(arr(0));" & vbCrLf
    htStr = htStr & "      if (Helpers.ucase(rec) == Helpers.ucase('True')) {" & vbCrLf
    htStr = htStr & "        if (ele) {" & vbCrLf
    htStr = htStr & "          //ele.setAttribute('class', 'light');" & vbCrLf
    htStr = htStr & "        }" & vbCrLf
    htStr = htStr & "        ele = null;" & vbCrLf
    htStr = htStr & "      } else {" & vbCrLf
    htStr = htStr & "        //ele.setAttribute('class', 'warning');" & vbCrLf
    htStr = htStr & "      }" & vbCrLf
    htStr = htStr & "    }" & vbCrLf
    htStr = htStr & "  }" & vbCrLf
    htStr = htStr & "}" & vbCrLf


  htStr = htStr & "</script>"
  response.write htStr
  js = js & "<script>" & vbCrLf
  js = js & "  " & vbCrLf
  js = js & "  " & vbCrLf
  js = js & "</script>"
  response.write js
  
End Sub

Sub addCSS()
    response.write "<style>"
    response.write "    .tabs {"
    response.write "      display: flex;"
    response.write "    }"
    response.write "    .tab {"
   ' response.write "      padding: 10px 15px;"
    response.write "      background-color: #f1f1f1;"
    response.write "      cursor: pointer;"
    response.write "    }"
    response.write "    .tab.active {"
    response.write "      background-color: #ccc;"
    response.write "    }"
    response.write "    .tab-content {"
    response.write "      display: none;"
    response.write "      padding: 15px;"
    response.write "      border: 1px solid #ccc;"
    response.write "    }"
    response.write "    .tab-content.active {"
    response.write "      display: block;"
    response.write "    }"
    response.write ".popup {"
    response.write "      display: none;"
    response.write "      position: fixed;"
    response.write "      top: 50%;"
    response.write "      left: 50%;"
    response.write "      transform: translate(-50%, -50%);"
    'response.write "      padding: 20px;"
    response.write "      border: 1px solid #ccc;"
    response.write "      background-color: #fff;"
    response.write "      z-index: 9999;"
    response.write "    }"
    response.write ""
    response.write "    /* Overlay to darken the background when the ""popup"" is open */"
    response.write "    .overlay {"
    response.write "      display: none;"
    response.write "      position: fixed;"
    response.write "      top: 0;"
    response.write "      left: 0;"
    response.write "      width: 100%;"
    response.write "      height: 100%;"
    response.write "      background-color: rgba(0, 0, 0, 0.5);"
    response.write "      z-index: 9998;"
    response.write "    }"
        
        response.write " #vehicleTable {"
        response.write "      width: 100%;"
        response.write "   }"
        response.write "   "
        response.write "    #vehicleTable thead th {"
        response.write "      font-weight: bold;"
        response.write "      background-color: #f8f9fa;"
        response.write "      border-top: none;"
        response.write "    }"
        response.write "    #vehicleTable tbody td {"
        response.write "      border-top: none;"
        response.write "    }"
        response.write ""
        response.write "    div.dataTables_wrapper div.dataTables_filter input {"
        response.write "      width: 200px;"
        response.write "    }"
        response.write ""
        response.write "    /* Styling for pagination */"
        response.write "    .dataTables_wrapper .dataTables_paginate .paginate_button {"
        response.write "      margin: 0 5px;"
        response.write "      padding: 6px 10px;"
        response.write "      border-radius: 4px;"
        response.write "      color: #007bff;"
        response.write "      background-color: #f8f9fa;"
        response.write "      border: 1px solid #dee2e6;"
        response.write "      cursor: pointer;"
        response.write "    }"
        response.write ""
        response.write "    .dataTables_wrapper .dataTables_paginate .paginate_button:hover {"
        response.write "      background-color: #007bff;"
        response.write "      color: #fff;"
        response.write "    }"
        response.write ""
        response.write "    .dataTables_wrapper .dataTables_paginate .paginate_button.current {"
        response.write "      background-color: #007bff;"
        response.write "      color: #fff;"
        response.write "    }"


        response.write ".navbar {"
        response.write "      background-color: #ddd9d9;"
        response.write "       margin-top:15px; "
                response.write "       margin-inline:15px; "
                response.write "       border-radius:4px; "
        response.write "    }"
        response.write "#navbarNav {"
        'response.write "      background-color: #F1F1F1;"
        'response.write "   margin-top:10px; "
        response.write "    }"
        response.write "    .navbar-brand {"
        response.write "      color: #aaa8a8;"
        response.write "      font-size: 24px;"
        response.write "    }"
        response.write "    .nav-link {"
        response.write "      color:#ffc107 ;"
        response.write "      font-size: 16px;"
        response.write "      margin-right: 5px;"
        response.write "      border:1px solid #a59f9f;"
        response.write "    }"
        response.write "    .nav-link:hover {"
        response.write "      color: #ffc107; /* Change color on hover */"
        response.write "      background-color: #666161;"
        response.write "    }"
        response.write "    /* Custom styling for the content area */"
        response.write "    .content {"
        response.write "      display: none;"
       ' response.write "      padding: 20px;"
        response.write "      background-color: #f8f9fa;"
        response.write "      min-height: 250px;"
        response.write "    }"
        response.write ".nav-link.active {"
        response.write "    color: #ffffff; /* Change color for the active link */"
        response.write "    background-color: #004ced42;"
        response.write "  }"
        
        response.write " #theader{"
        response.write "    background-color: #d2d2ed;"
        response.write "  }"
                
        response.write "#trPrintControl{"
        response.write "   display:none ;"
        
        response.write "  }"
                
                 response.write "#disable-link{"
        response.write "  color: #999 ;"
                response.write "  pointer-events: none;"
                response.write "  cursor: default ;"
                response.write "  text-decoration: none ;"
        
        response.write "  }"
                
                
                

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
