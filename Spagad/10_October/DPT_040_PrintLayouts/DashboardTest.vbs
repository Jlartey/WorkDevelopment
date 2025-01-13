'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

response.write " <script src='https://cdn.jsdelivr.net/npm/chart.js'></script>"

response.write " <script src='https://cdn.plot.ly/plotly-2.24.1.min.js'></script>"

response.write " <script src='https://cdnjs.cloudflare.com/ajax/libs/apexcharts/3.35.5/apexcharts.min.js'></script>"

response.write " <link href='https://fonts.googleapis.com/css2?family=Montserrat:wght@100;200;300;400;500;600;700;800;900&display=swap' rel='stylesheet'>"

response.write " <link href='https://fonts.googleapis.com/icon?family=Material+Icons+Outlined' rel='stylesheet'>"

StyleCSS
AddContentJS

Dim staffPat, periodStart, periodEnd, dt, totalBill, finalAmt, finalAmt1, finalAmt3, paidBill
Dim hrf1, hrf2, hrf3, hrf4, hrf5, hrf6, hrf7, hrf8, hrf9, hrf10, hrf11, hrf12, hrf13, hrf14, hrf15, hrf16, hrf17

If Len(Trim(Request.QueryString("selectedValue"))) > 1 Then
    periodStart = Trim(Request.QueryString("selectedValue"))
    periodEnd = Trim(Request.QueryString("selectedValue1"))
    
    periodStart = FormatDate(periodStart)
    periodEnd = FormatDate(periodEnd)
Else
    periodStart = FormatDate("2022-07-01")
    periodEnd = FormatDate(Now)
End If

totalPatients = GetTotalPat()
totalPatientsAd = GetTotalPatAd()
totalVisit = GetTotalVisit()
totalSpecialist = GetTotalSpec()
totalStaff = GetTotalStaff()
MortalityRate = GetMortalityR()
bedOccupancy = GetBedOcc()
totalCancelledApp = GetTotalCancelled()
pharmWaitAvg = GetPharmWaitTime()
opdWaitTimeAvg = GetOpdAvgWaitTime()

staffPat = totalStaff / totalPatients
Patstaff = totalPatients / totalStaff

hrf1 = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=AverageStayByDpt&PositionForTableName=WorkingDay&WorkingDayID=&printfilter=" & periodStart & "&printfilter1=" & periodEnd
hrf2 = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=AverageStayByWard&PositionForTableName=WorkingDay&WorkingDayID=&printfilter=" & periodStart & "&printfilter1=" & periodEnd
hrf3 = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=drugCostPerStay&PositionForTableName=WorkingDay&WorkingDayID=&printfilter=" & periodStart & "&printfilter1=" & periodEnd
hrf4 = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=AppointmentsRpt&PositionForTableName=WorkingDay&WorkingDayID=&printfilter=" & periodStart & "&printfilter1=" & periodEnd
hrf5 = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=&PositionForTableName=WorkingDay&WorkingDayID=&printfilter=" & periodStart & "&printfilter1=" & periodEnd
hrf6 = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=Top10Suppliers&PositionForTableName=WorkingDay&WorkingDayID=&printfilter=" & periodStart & "&printfilter1=" & periodEnd
hrf7 = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=IncomingDrugsDistribution&PositionForTableName=WorkingDay&WorkingDayID=&printfilter=" & periodStart & "&printfilter1=" & periodEnd
hrf8 = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=PatientsonAdmission&PositionForTableName=WorkingDay&WorkingDayID=&printfilter=" & periodStart & "&printfilter1=" & periodEnd
hrf9 = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=&PositionForTableName=WorkingDay&WorkingDayID=&printfilter=" & periodStart & "&printfilter1=" & periodEnd
hrf10 = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=top10MostSoldDrugsBySale&PositionForTableName=WorkingDay&WorkingDayID=&printfilter=" & periodStart & "&printfilter1=" & periodEnd
hrf11 = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=top10MartItems&PositionForTableName=WorkingDay&WorkingDayID=&printfilter=" & periodStart & "&printfilter1=" & periodEnd
hrf12 = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=top10MostVisitedDpt&PositionForTableName=WorkingDay&WorkingDayID=&printfilter=" & periodStart & "&printfilter1=" & periodEnd
hrf13 = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=top10MostLabTestsBySale&PositionForTableName=WorkingDay&WorkingDayID=&printfilter=" & periodStart & "&printfilter1=" & periodEnd
hrf14 = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=AverageLabWaitTime&PositionForTableName=WorkingDay&WorkingDayID=&printfilter=" & periodStart & "&printfilter1=" & periodEnd
hrf15 = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=AverageDptWaitTime&PositionForTableName=WorkingDay&WorkingDayID=&printfilter=" & periodStart & "&printfilter1=" & periodEnd
hrf16 = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=MostPrescribedDrugs&PositionForTableName=WorkingDay&WorkingDayID=&printfilter=" & periodStart & "&printfilter1=" & periodEnd
hrf18 = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=MostPrescribedDrugByQuantity&PositionForTableName=WorkingDay&WorkingDayID=&printfilter=" & periodStart & "&printfilter1=" & periodEnd
hrf17 = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=MostPrescribedLab&PositionForTableName=WorkingDay&WorkingDayID=&printfilter=" & periodStart & "&printfilter1=" & periodEnd
hrf18 = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=ReAdmissionRate&PositionForTableName=WorkingDay&PrintFilter=&PrintFilter=" & periodStart & "&printfilter1=" & periodEnd

finalAmt = GetDrugSaleAmt()
finalAmt1 = GetLabRequestAmt()
finalAmt2 = GetTreatAmt()
finalAmt3 = GetVstCost()
paidBill = GetPaidAmt()

totalBill = (finalAmt + finalAmt1 + finalAmt2 + CDbl(finalAmt3))
AvgtotalBill = totalBill / totalPatients

response.write "<div id='container'>"
response.write "    <div class='grid-container'>"
response.write "      <header class=""header"">"
response.write "        <div class=""menu-icon"" onclick=""openSidebar()"">"
response.write "          <span class=""material-icons-outlined""></span>"
response.write "        </div>"
response.write "        <div class=""header-left"">"
response.write "          <span class=""material-icons-outlined""></span>"
response.write "        </div>"
response.write "        <div class=""header-right"">"
response.write "          <span class=""material-icons-outlined""></span>"
response.write "          <span class=""material-icons-outlined""></span>"
response.write "          <span class=""material-icons-outlined""></span>"
response.write "        </div>"
response.write "      </header>"
response.write "      <!-- End Header -->"

response.write "      <!-- Sidebar -->"
response.write "      <aside id=""sidebar"">"
response.write "        <div class=""sidebar-title"">"
response.write "          <div class=""sidebar-brand"">"
response.write "            <span class=""material-icons-outlined"">groups</span> RMC KPIs"
response.write "          </div>"
response.write "          <span class=""material-icons-outlined"" onclick=""closeSidebar()"">close</span>"
response.write "        </div>"
response.write ""
response.write "        <ul class=""sidebar-list"">"
response.write "          <li class=""sidebar-list-item"">"
response.write "            <a href=""#"">"
response.write "              <span class=""material-icons-outlined"">dashboard</span><span>Dashboard</span>"
response.write "            </a>"
response.write "          </li>"
' response.write "          <li class=""sidebar-list-item"">"
' response.write "           <a href='#' />"
' response.write "              <span class=""material-icons-outlined"">groups</span><span>Visitations</span>"
' response.write "            </a>"
' response.write "          </li>"
response.write "          <li class=""sidebar-list-item"">"
response.write "            <a href=""wpgPrtPrintLayoutAll.asp?PrintLayoutName=GetPatientRmcKPI&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=2022-01-01&selectedValue1=2023-08-24"" target=""_blank"">"
response.write "              <span class=""material-icons-outlined"">poll</span><span>Patients</span>"
response.write "            </a>"
response.write "          </li>"
' response.write "          <li class=""sidebar-list-item"" >"
' response.write "           <a href='#'>"
' response.write "              <span class=""material-icons-outlined"">poll</span><span>SPecialists KPI</span>"
' response.write "            </a>"
' response.write "          </li>"
' response.write "          <li class=""sidebar-list-item"">"
' response.write "            <a href=""#"" target=""_blank"">"
' response.write "              <span class=""material-icons-outlined"">groups</span><span>Staffs KPI</span>"
' response.write "            </a>"
' response.write "          </li>"
response.write "          <li class=""sidebar-list-item"">"
response.write "            <a href=""wpgPrtPrintLayoutAll.asp?PrintLayoutName=GetSupplierRmcKPI&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=2022-01-01&selectedValue1=2023-08-24"" target=""_blank"">"
response.write "              <span class=""material-icons-outlined"">poll</span><span>Suppliers</span>"
response.write "            </a>"
response.write "          </li>"
response.write "        </ul>"
response.write "      </aside>"
response.write "      <!-- End Sidebar -->"

response.write "      <!-- Main -->"
response.write "      <main class=""main-container"">"
response.write "            <div class=""date_filter"">"
response.write "                <label for=""fromDate"">From: </label>"
response.write "                <input type=""date"" id=""fromDate"" class=""styled-input"">"
response.write "                <label for=""toDate"">To: </label>"
response.write "                <input type=""date"" id=""toDate"" class=""styled-input"">"
response.write "                <button id=""filterButton"" class=""styled-button"">Apply Filter</button>"
response.write "            </div>"
response.write "            <div>"
response.write "            <h4>Records Ranging from [" & periodStart & "] to [" & periodEnd & "] </h4>"
response.write "            </div>"
response.write "      <div class=""main-title"">"
response.write "        <h2>DASHBOARD</h2>"
response.write "      </div>"

response.write "      <div class=""main-cards"">"
response.write "        <div class=""card"">"
response.write "          <div class=""card-inner"">"
response.write "            <h3>Total Patients</h3>"
response.write "            <span class=""material-icons-outlined"">groups</span>"
response.write "          </div>"
response.write "          <h1>" & FormatNumber(totalPatients, 0) & "</h1>"
response.write "        </div>"
response.write ""
response.write "        <div class=""card"">"
response.write "          <div class=""card-inner"">"
response.write "            <h3>Total Patients Admitted</h3>"
response.write "            <span class=""material-icons-outlined"">groups</span>"
response.write "          </div>"
response.write "          <h1>" & FormatNumber(totalPatientsAd, 0) & "</h1>"
response.write "        </div>"
response.write ""
response.write "        <div class=""card"">"
response.write "          <div class=""card-inner"">"
response.write "            <h3>Total Visits</h3>"
response.write "            <span class=""material-icons-outlined"">groups</span>"
response.write "          </div>"
response.write "          <h1>" & FormatNumber(totalVisit, 0) & "</h1>"
response.write "        </div>"
response.write ""
response.write "        <div class=""card"">"
response.write "          <div class=""card-inner"">"
response.write "            <h3>Total Treatment Cost</h3>"
response.write "            <span class=""material-icons-outlined"">money</span>"
response.write "          </div>"
response.write "          <h1>" & FormatNumber(totalBill, 0) & "</h1>"
response.write "        </div>"

response.write "      </div>"

response.write "      <div class=""main-cards"">"
response.write "        <div class=""card2"">"
response.write "          <div class=""card-inner"">"
response.write "            <h3>Total Staffs</h3>"
response.write "            <span class=""material-icons-outlined"">groups</span>"
response.write "          </div>"
response.write "          <h1>" & FormatNumber(totalStaff, 0) & "</h1>"
response.write "        </div>"
response.write ""
response.write "        <div class=""card2"">"
response.write "          <div class=""card-inner"">"
response.write "            <h3>Staff-To-Patient Ratio</h3>"
response.write "            <span class=""material-icons-outlined"">groups</span>"
response.write "          </div>"
response.write "          <h1>" & FormatNumber(staffPat, 5) & "</h1>"
response.write "        </div>"
response.write ""
response.write "        <div class=""card2"">"
response.write "          <div class=""card-inner"">"
response.write "            <h3>Patient-To-Staff Ratio</h3>"
response.write "            <span class=""material-icons-outlined"">groups</span>"
response.write "          </div>"
response.write "          <h1>" & FormatNumber(Patstaff, 2) & "</h1>"
response.write "        </div>"
response.write ""
response.write "        <div class=""card2"">"
response.write "          <div class=""card-inner"">"
response.write "            <h3>Mortality Rate</h3>"
response.write "            <span class=""material-icons-outlined"">groups</span>"
response.write "          </div>"
response.write "          <h1 id='mortalityRate'></h1>"
response.write "        </div>"
' response.write ""
' response.write "        <div class=""card2"">"
' response.write "          <div class=""card-inner"">"
' response.write "            <h3>Bed Occupancy Rate</h3>"
' response.write "            <span class=""material-icons-outlined"">groups</span>"
' response.write "          </div>"
' response.write "          <h1>" & FormatNumber(bedOccupancy, 2) & "%</h1>"
' response.write "        </div>"

response.write "      </div>"
response.write ""
response.write "      <div class=""main-cards"">"
response.write "        <div class=""card3"">"
response.write "          <div class=""card-inner"">"
response.write "            <h3>Total Cancelled Appointments</h3>"
response.write "            <span class=""material-icons-outlined"">groups</span>"
response.write "          </div>"
response.write "          <h1>" & FormatNumber(totalCancelledApp, 0) & "</h1>"
response.write "        </div>"
response.write ""
response.write "        <div class=""card3"">"
response.write "          <div class=""card-inner"">"
response.write "            <h3>Pharmacy Avg Waiting Time</h3>"
response.write "            <span class=""material-icons-outlined"">groups</span>"
response.write "          </div>"
response.write "          <h1> " & FormatNumber(pharmWaitAvg, 0) & " Mins</h1>"
response.write "        </div>"

response.write "        <div class=""card3"">"
response.write "          <div class=""card-inner"">"
response.write "            <h3>OPD Avg Waiting Time</h3>"
response.write "            <span class=""material-icons-outlined"">groups</span>"
response.write "          </div>"
response.write "          <h1>" & Round(opdWaitTimeAvg / 60, 2) & " Hours</h1>"
response.write "        </div>"
response.write ""
response.write "        <div class=""card3"">"
response.write "          <div class=""card-inner"">"
response.write "            <h3>Bed Occupancy Rate</h3>"
response.write "            <span class=""material-icons-outlined"">groups</span>"
response.write "          </div>"
response.write "          <h1>" & FormatNumber(bedOccupancy, 2) & "%</h1>"
response.write "        </div>"
response.write "      </div>"

response.write "      <div class=""charts"">"
response.write "        <div class=""charts-card"">"
response.write "          <h2 class=""chart-title"">Readmission Rate</h2>"
response.write "            <canvas id=""readmissionRate"" width=""300"" height=""300""></canvas>"
response.write "          <button type = ""button"" class = ""btn3""> "
response.write "            <a href='" & hrf18 & "' target='_blank'>View More</a>"
response.write "          </button>"
response.write "        </div>"
response.write "        <div class=""charts-card"">"
response.write "          <h2 class=""chart-title"">Average Hospital Stay At Department By Days</h2>"
response.write "            <canvas id=""pieChart2"" width=""300"" height=""300""></canvas>"
response.write "          <button type = ""button"" class = ""btn2""> "
response.write "            <a href = '" & hrf1 & "' target = '_blank'>View More</a>"
response.write "          </button>"
response.write "        </div>"
response.write "        <div class=""charts-card"">"
response.write "          <h2 class=""chart-title"">Average Hospital Stay At Ward By Days</h2>"
response.write "            <canvas id=""pieChart3"" width=""300"" height=""300""></canvas>"
response.write "          <button type = ""button"" class = ""btn2""> "
response.write "            <a href = '" & hrf2 & "' target = '_blank'>View More</a>"
response.write "          </button>"
response.write "        </div>"
response.write "        <div class=""charts-card"">"
response.write "          <h2 class=""chart-title"">Average Drug Cost Per Stay By Departments</h2>"
response.write "            <canvas id=""pieChart4"" width=""300"" height=""300""></canvas>"
response.write "          <button type = ""button"" class = ""btn3""> "
response.write "            <a href='" & hrf3 & "' target='_blank'>View More</a>"
response.write "          </button>"
response.write "        </div>"
response.write "        <div class=""charts-card"">"
response.write "          <h2 class=""chart-title"">Appointment Status</h2>"
response.write "            <canvas id=""CancelChart"" ></canvas>"
response.write "          <button type = ""button"" class = ""btn3""> "
response.write "            <a href='" & hrf4 & "' target='_blank'>View More</a>"
response.write "          </button>"
response.write "        </div>"
response.write "        <div class=""charts-card"">"
response.write "          <h2 class=""chart-title"">Gender</h2>"
response.write "            <canvas id=""pieChart5"" ></canvas>"
response.write "          <button type = ""button"" class = ""btn3""> "
response.write "            <a href='" & hrf5 & "' target='_blank'>View More</a>"
response.write "          </button>"
response.write "        </div>"
response.write "        <div class=""charts-card"">"
response.write "          <h2 class=""chart-title"">Top 10 Suppliers By Goods Received</h2>"
response.write "            <canvas id=""pieChart6"" width=""300"" height=""300""></canvas>"
response.write "          <button type = ""button"" class = ""btn3""> "
response.write "            <a href='" & hrf6 & "' target='_blank'>View More</a>"
response.write "          </button>"
response.write "        </div>"
response.write "        <div class=""charts-card"">"
response.write "          <h2 class=""chart-title"">Incoming Drugs Distribution By Suppliers</h2>"
response.write "           <canvas id=""suppDg-chart"" width=""300"" height=""300""></canvas>"
response.write "          <button type = ""button"" class = ""btn3""> "
response.write "            <a href='" & hrf7 & "' target='_blank'>View More</a>"
response.write "          </button>"
response.write "        </div>"
response.write "        <div class=""charts-card"">"
response.write "          <h2 class=""chart-title"">Total Patients on Admission Over Time</h2>"
response.write "           <canvas id=""Admission-chart"" width=""300"" height=""300""></canvas>"
response.write "          <button type = ""button"" class = ""btn3""> "
response.write "            <a href='" & hrf8 & "' target='_blank'>View More</a>"
response.write "          </button>"
response.write "        </div>"
response.write "        <div class=""charts-card"">"
response.write "          <h2 class=""chart-title"">Total Patients on Admission Over Time</h2>"
response.write "           <canvas id=""AdScPat"" width=""300"" height=""300""></canvas>"
response.write "          <button type = ""button"" class = ""btn3""> "
response.write "            <a href='" & hrf8 & "' target='_blank'>View More</a>"
response.write "          </button>"
response.write "        </div>"
response.write "        <div class=""charts-card"">"
response.write "          <h2 class=""chart-title"">Top 10 Most Visited Departments</h2>"
response.write "          <canvas id=""visitedDpt"" width=""300"" height=""300""></canvas>"
response.write "          <button type = ""button"" class = ""btn3""> "
response.write "            <a href='" & hrf12 & "' target='_blank'>View More</a>"
response.write "          </button>"
response.write "        </div>"
response.write "        <div class=""charts-card"">"
response.write "          <h2 class=""chart-title"">Top 10 Most Sold Drugs By Quantity</h2>"
response.write "           <canvas id=""prescribedDrug"" width=""300"" height=""300""></canvas>"
response.write "          <button type = ""button"" class = ""btn3""> "
response.write "            <a href='" & hrf16 & "' target='_blank'>View More</a>"
response.write "          </button>"
response.write "        </div>"
response.write "        <div class=""charts-card"">"
response.write "          <h2 class=""chart-title"">Top 10 Most Sold Drugs By Amount</h2>"
response.write "           <canvas id=""soldDrugs"" width=""300"" height=""300""></canvas>"
response.write "          <button type = ""button"" class = ""btn3""> "
response.write "            <a href='" & hrf10 & "' target='_blank'>View More</a>"
response.write "          </button>"
response.write "        </div>"
response.write "        <div class=""charts-card"">"
response.write "          <h2 class=""chart-title"">Top 10 Most Prescribed Drugs</h2>"
response.write "           <canvas id=""prescribedDrugByQuantity"" width=""300"" height=""300""></canvas>"
response.write "          <button type = ""button"" class = ""btn3""> "
response.write "            <a href='" & hrf18 & "' target='_blank'>View More</a>"
response.write "          </button>"
response.write "        </div>"
response.write "        <div class=""charts-card"">"
response.write "          <h2 class=""chart-title"">TOP 10 Most Requested Labs</h2>"
response.write "         <canvas id=""prescribedLab"" width=""300"" height=""300""></canvas>"
response.write "          <button type = ""button"" class = ""btn3""> "
response.write "            <a href='" & hrf17 & "' target='_blank'>View More</a>"
response.write "          </button>"
response.write "        </div>"
response.write "        <div class=""charts-card"">"
response.write "          <h2 class=""chart-title"">TOP 10 LABTESTS BY SALES</h2>"
response.write "         <canvas id=""MostLabTest"" width=""300"" height=""300""></canvas>"
response.write "          <button type = ""button"" class = ""btn3""> "
response.write "            <a href='" & hrf13 & "' target='_blank'>View More</a>"
response.write "          </button>"
response.write "        </div>"
response.write "        <div class=""charts-card"">"
response.write "          <h2 class=""chart-title"">TOP 10 MART ITEMS BY SALES</h2>"
response.write "          <canvas id=""MartItems"" width=""300"" height=""300""></canvas>"
response.write "          <button type = ""button"" class = ""btn3""> "
response.write "            <a href='" & hrf11 & "' target='_blank'>View More</a>"
response.write "          </button>"
response.write "        </div>"
response.write "        <div class=""charts-card"">"
response.write "          <h2 class=""chart-title"">TOP 10 Laboratory Average Waiting Time</h2>"
response.write "            <canvas id=""waitChart"" width=""300"" height=""300""></canvas>"
response.write "          <button type = ""button"" class = ""btn3""> "
response.write "            <a href='" & hrf14 & "' target='_blank'>View More</a>"
response.write "          </button>"
response.write "        </div>"
response.write "        <div class=""charts-card"">"
response.write "          <h2 class=""chart-title"">TOP 10 Departments Average Waiting Time</h2>"
response.write "            <canvas id=""dptWaitTime"" width=""300"" height=""300""></canvas>"
response.write "          <button type = ""button"" class = ""btn3""> "
response.write "            <a href='" & hrf15 & "' target='_blank'>View More</a>"
response.write "          </button>"
response.write "        </div>"


response.write "      </div>"


response.write "      </main>"
response.write "      <!-- End Main -->"
response.write "    </div>"
response.write "</div>"

response.write "<script>"
response.write "function formatMortalityRate(value) {"
response.write "    var formattedValue = (+value).toPrecision(2);"
response.write "    return formattedValue;"
response.write "}"
response.write ""
response.write "document.addEventListener(""DOMContentLoaded"", function() {"
'response.write "document.getElementById('fromDate').value = '" & defaultStartDate & "';"
'response.write "document.getElementById('toDate').value = '" & defaultEndDate & "';"
response.write "    var mortalityRate = '" & MortalityRate & "';"
response.write "    var formattedMortalityRate = formatMortalityRate(mortalityRate);"
response.write "    "
response.write "    document.getElementById('mortalityRate').textContent = formattedMortalityRate;"
response.write "});"

response.write "document.getElementById('filterButton').addEventListener('click', function() {"
response.write "    var fromDate = document.getElementById('fromDate').value;"
response.write "    var toDate = document.getElementById('toDate').value;"
response.write "    let url = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=DashboardTest&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=' + fromDate + '&selectedValue1=' + toDate;"
response.write "    window.location.href = url;"
response.write "    });"
response.write "</script>"

AvgDptJS
AvgWdJS
DrugCostPerStayDptJS
ReadmissionRate
AppChartJS
SupplierChartJS
SupplierDrugsJS
TotalPatByAdJS
TotalPatByAdJS2
mostVisitedDpt
mostPrescribedDrug
mostPrescribedDrugByQuantity
MostPrescribedLab
mostSoldDrugs
MostLabTestBySale
mostSoldDMartItems
GenderChartJS
AvgWaitingTime
GetDptWaitTimeJS

'GetOpdWaitTimeJS

Sub StyleCSS()

    response.write "<style>"

    response.write "#container {"
    response.write "  width: 100vw; "
    response.write "  margin: 0;"
    response.write "  padding: 0;"
    response.write "  background-color: #1d2634;"
    response.write "  color: #9e9ea4;"
    response.write "  font-family: 'Montserrat', sans-serif;"
    response.write "}"
    response.write "body { "
    response.write "  overflow: hidden; "
    response.write "  width: 100% "
    response.write "}"
    response.write ".material-icons-outlined {"
    response.write "  vertical-align: middle;"
    response.write "  line-height: 1px;"
    response.write "  font-size: 35px;"
    response.write "  margin-inline: 15px;"
    response.write "}"

    response.write "#trPrintControl {"
    response.write "  display: none;"
    response.write "}"

    response.write ".grid-container {"
    response.write "  display: grid;"
    response.write "  grid-template-columns: 260px 1fr 1fr 1fr;"
    response.write "  grid-template-rows: 0.2fr 3fr;"
    response.write "  grid-template-areas:"
    response.write "    'sidebar header header header'"
    response.write "    'sidebar main main main';"
    response.write "  height: 100vh;"
    response.write "}"

    response.write "/* ---------- HEADER ---------- */"
    response.write ".header {"
    response.write "  grid-area: header;"
    response.write "  height: 70px;"
    ' response.write "  display: flex;"
    response.write "  display: none;"
    response.write "  align-items: center;"
    response.write "  justify-content: space-between;"
    response.write "  padding: 0 30px 0 30px;"
    response.write "  box-shadow: 0 6px 7px -3px rgba(0, 0, 0, 0.35);"
    response.write "}"

    response.write ".menu-icon {"
    response.write "  display: none;"
    response.write "}"
    response.write ""
    response.write "/* ---------- SIDEBAR ---------- */"

    response.write "#sidebar {"
    response.write "  grid-area: sidebar;"
    response.write "  height: 100%;"
    response.write "  background-color: #263043;"
    response.write "  overflow-y: auto;"
    response.write "  transition: all 0.5s;"
    response.write "  -webkit-transition: all 0.5s;"
    response.write "}"

    response.write ".sidebar-title {"
    response.write "  display: flex;"
    response.write "  justify-content: space-between;"
    response.write "  align-items: center;"
    response.write "  padding: 30px 30px 30px 30px;"
    response.write "  margin-bottom: 30px;"
    response.write "}"

    response.write ".sidebar-list-item a { "
    response.write "   display: flex;"
    response.write "   align-items: center;"
    ' response.write "   justify-content: space-around;"
    response.write "}"


    response.write ".sidebar-title > span {"
    response.write "  display: none;"
    response.write "}"
    response.write ""
    response.write ".sidebar-brand {"
    response.write "  margin-top: 15px;"
    response.write "  font-size: 20px;"
    response.write "  font-weight: 700;"
    response.write "}"

    response.write ".sidebar-list {"
    response.write "  padding: 0;"
    response.write "  margin-top: 15px;"
    response.write "  list-style-type: none;"
    response.write "}"

    response.write ".sidebar-list-item {"
    response.write "  padding: 20px 20px 20px 20px;"
    response.write "  font-size: 18px;"
    response.write "}"

    response.write ".sidebar-list-item:hover {"
    response.write "  background-color: rgba(255, 255, 255, 0.2);"
    response.write "  cursor: pointer;"
    response.write "}"

    response.write ".sidebar-list-item > a {"
    response.write "  text-decoration: none;"
    response.write "  color: #9e9ea4;"
    response.write "}"

    response.write ".sidebar-responsive {"
    response.write "  display: inline !important;"
    response.write "  position: absolute;"
    response.write "  /*"
    response.write "    the z-index of the ApexCharts is 11"
    response.write "    we want the z-index of the sidebar higher so that"
    response.write "    the charts are not showing over the sidebar "
    response.write "    on small screens"
    response.write "  */"
    response.write "  z-index: 12 !important;"
    response.write "}"
    response.write ""
    response.write "/* ---------- MAIN ---------- */"

    response.write ".main-container {"
    response.write "  grid-area: main;"
    response.write "  overflow-y: auto;"
    response.write "  padding: 0px 20px;"
    response.write "  color: rgba(255, 255, 255, 0.95);"
    response.write "}"

    response.write ".main-title {"
    response.write "  display: flex;"
    response.write "  justify-content: space-between;"
    response.write "}"

    response.write ".main-cards {"
    response.write "  display: grid;"
    response.write "  grid-template-columns: 1fr 1fr 1fr 1fr;"
    response.write "  gap: 20px;"
    response.write "  margin: 20px 0;"
    response.write "}"

    response.write ".card {"
    response.write "  display: flex;"
    response.write "  flex-direction: column;"
    response.write "  justify-content: space-around;"
    response.write "  padding: 25px;"
    response.write "  border-radius: 5px;"
    response.write "}"

    response.write ".btn1, .btn2, .btn3, .btn4, .btn5, .btn6 {"
    response.write "   position: absolute;"
    response.write "   top: 10px;"
    response.write "   right: 10px;"
    response.write "   background-color: #263055;"
    response.write "   color: #fff;"
    response.write "   border-radius: 5px;"
    response.write "}"

    response.write "a { "
    response.write "    color: #fff;"
    response.write "    text-decoration: none;"
    response.write "}"


    response.write ".card:first-child {"
    response.write "  background-color: #2962ff;"
    response.write "}"
    response.write ".card:nth-child(2) {"
    response.write "  background-color: #ff6d00;"
    response.write "}"

    response.write ".card:nth-child(3) {"
    response.write "  background-color: #2e7d32;"
    response.write "}"
    response.write ""
    response.write ".card:nth-child(4) {"
    response.write "  background-color: #d50000;"
    response.write "}"
    response.write ""
    response.write ".card2 {"
    response.write "  display: flex;"
    response.write "  flex-direction: column;"
    response.write "  justify-content: space-around;"
    response.write "  padding: 25px;"
    response.write "  border-radius: 5px;"
    response.write "}"
    response.write ""
    response.write ".card2:first-child {"
    response.write "  background-color:#d50000;"
    response.write "}"
    response.write ""
    response.write ".card2:nth-child(2) {"
    response.write "  background-color:#2e7d32 ;"
    response.write "}"
    response.write ""
    response.write ".card2:nth-child(3) {"
    response.write "  background-color:  #ff6d00;"
    response.write "}"
    response.write ""
    response.write ".card2:nth-child(4) {"
    response.write "  background-color: #2962ff;"
    response.write "}"

    response.write ""
    response.write ".card3 {"
    response.write "  display: flex;"
    response.write "  flex-direction: column;"
    response.write "  justify-content: space-around;"
    response.write "  padding: 25px;"
    response.write "  border-radius: 5px;"
    response.write "}"
    response.write ""
    response.write ".card3:first-child {"
    response.write "  background-color:#2962ff;"
    response.write "}"
    response.write ""
    response.write ".card3:nth-child(2) {"
    response.write "  background-color:#2e7d32 ;"
    response.write "}"
    response.write ""
    response.write ".card3:nth-child(3) {"
    response.write "  background-color:  #ff6d00;"
    response.write "}"
    response.write ""
    response.write ".card3:nth-child(4) {"
    response.write "  background-color:  #d50000;"
    response.write "}"
    response.write ""
    response.write ".card-inner {"
    response.write "  display: flex;"
    response.write "  align-items: center;"
    response.write "  justify-content: space-between;"
    response.write "}"
    response.write ""
    response.write ".card-inner > .material-icons-outlined {"
    response.write "  font-size: 45px;"
    response.write "}"
    response.write ""
    response.write ".charts {"
    response.write "  display: grid;"
    response.write "  grid-template-columns: 1fr 1fr;"
    response.write "  gap: 20px;"
    response.write "  margin-top: 60px;"
    response.write "}"
    response.write ""
    response.write ".charts-card {"
    response.write "  background-color: #263043;"
    response.write "  margin-bottom: 20px;"
    response.write "  padding: 25px;"
    response.write "  box-sizing: border-box;"
    response.write "  -webkit-column-break-inside: avoid;"
    response.write "  border-radius: 5px;"
    response.write "  position: relative;"
    response.write "  box-shadow: 0 6px 7px -4px rgba(0, 0, 0, 0.2);"
    response.write "}"
    response.write ""
    response.write ".chart-title {"
    response.write "  display: flex;"
    response.write "  align-items: center;"
    response.write "  justify-content: center;"
    response.write "}"
    response.write ""
    response.write "/* ---------- MEDIA QUERIES ---------- */"
    response.write ""
    response.write "/* Medium <= 992px */"
    response.write ""
    response.write "@media screen and (max-width: 992px) {"
    response.write "  .grid-container {"
    response.write "    grid-template-columns: 1fr;"
    response.write "    grid-template-rows: 0.2fr 3fr;"
    response.write "    grid-template-areas:"
    response.write "      'header'"
    response.write "      'main';"
    response.write "  }"
    response.write ""
    response.write "  #sidebar {"
    response.write "    display: none;"
    response.write "  }"
    response.write ""
    response.write "  .menu-icon {"
    response.write "    display: inline;"
    response.write "  }"
    response.write ""
    response.write "  .sidebar-title > span {"
    response.write "    display: inline;"
    response.write "  }"
    response.write "}"
    response.write ""
    response.write "/* Small <= 768px */"
    response.write ""
    response.write "@media screen and (max-width: 768px) {"
    response.write "  .main-cards {"
    response.write "    grid-template-columns: 1fr;"
    response.write "    gap: 10px;"
    response.write "    margin-bottom: 0;"
    response.write "  }"

    response.write "  .charts {"
    response.write "    grid-template-columns: 1fr;"
    response.write "    margin-top: 30px;"
    response.write "  }"
    response.write "}"
    response.write ""
    response.write "/* Extra Small <= 576px */"
    response.write ""
    response.write "@media screen and (max-width: 576px) {"
    response.write "  .hedaer-left {"
    response.write "    display: none;"
    response.write "  }"
    response.write "}"
    
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
    response.write "</style>"

End Sub

Sub AddContentJS()
    response.write "<script>"
    response.write "var sidebarOpen = false;"
    response.write "var sidebar = document.getElementById('sidebar');"
    response.write "function openSidebar() {"
    response.write "  if (!sidebarOpen) {"
    response.write "    sidebar.classList.add('sidebar-responsive');"
    response.write "    sidebarOpen = true;"
    response.write "  }"
    response.write "}"

    response.write "function closeSidebar() {"
    response.write "  if (sidebarOpen) {"
    response.write "    sidebar.classList.remove('sidebar-responsive');"
    response.write "    sidebarOpen = false;"
    response.write "  }"
    response.write "}"
    response.write "</script>"
End Sub

Sub AvgWaitingTime()

 Dim sql, rst
    
    Set rst = CreateObject("ADODB.recordset")
    
    sql = " WITH CombinedResults AS ( "
    sql = sql & " SELECT "
    sql = sql & " Investigation.LabTestID,"
    sql = sql & " LabTest.LabTestName,"
    sql = sql & " MAX(requestdate) AS max_requestdate,"
    sql = sql & " MAX(requestdate1) AS max_requestdate1,"
    sql = sql & " CASE"
    sql = sql & "   WHEN MAX(DATEDIFF(Hour, requestdate, requestdate1)) < 0 THEN 0"
    sql = sql & "    Else MAX (DateDiff(Hour, requestdate, requestdate1))"
    sql = sql & " END As Waiting_Time"
    sql = sql & " From"
    sql = sql & "   Investigation"
    sql = sql & "  LEFT JOIN LabTest ON Investigation.LabTestID = LabTest.LabTestID"
    sql = sql & " Where"
    sql = sql & " Investigation.billgroupid = 'BG002'"

    If (periodStart <> "") And (periodEnd <> "") Then
    sql = sql & " AND Investigation.RequestDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If

    sql = sql & " Group By"
    sql = sql & "   Investigation.labTestID , LabTest.LabTestName"
    sql = sql & " Union All"
    sql = sql & " SELECT"
    sql = sql & "   Investigation2.LabTestID,"
    sql = sql & "    LabTest.LabTestName,"
    sql = sql & "   MAX(requestdate) AS max_requestdate,"
    sql = sql & "    MAX(requestdate1) AS max_requestdate1,"
    sql = sql & "    CASE"
    sql = sql & "    WHEN MAX(DATEDIFF(Hour, requestdate, requestdate1)) < 0 THEN 0 "
    sql = sql & "        Else MAX (DateDiff(Hour, requestdate, requestdate1))"
    sql = sql & "    END As Waiting_Time "
    sql = sql & " From "
    sql = sql & "    Investigation2 "
    sql = sql & " LEFT JOIN LabTest ON Investigation2.LabTestID = LabTest.LabTestID "
    sql = sql & " Where "
    sql = sql & "   Investigation2.billgroupid = 'BG002'"
    
    If (periodStart <> "") And (periodEnd <> "") Then
    sql = sql & " AND Investigation2.RequestDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If

    sql = sql & " Group By"
    sql = sql & "    Investigation2.labTestID , LabTest.LabTestName"
    sql = sql & " )"
    sql = sql & " SELECT"
    sql = sql & "    TOP 10 "
    sql = sql & "    LabTestName,"
    sql = sql & "    AVG(Waiting_Time) As Average_Waiting_Time "
    sql = sql & " From "
    sql = sql & "  CombinedResults "
    sql = sql & " Group By"
    sql = sql & "    LabTestName "
    sql = sql & " Order By "
    sql = sql & "    Average_Waiting_Time desc"
    
    rst.open qryPro.FltQry(sql), conn, 3, 4
    
    Dim departmentsArray, stayArray
    departmentsArray = ""
    stayArray = ""
    
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            If departmentsArray <> "" Then departmentsArray = departmentsArray & ","
            If stayArray <> "" Then stayArray = stayArray & ","
            
            departmentsArray = departmentsArray & "'" & rst.fields("LabTestName") & "'"
            stayArray = stayArray & rst.fields("Average_Waiting_Time")
            
            rst.MoveNext
        Loop
        
        response.write "  <script>"
        response.write "    var labels = [" & departmentsArray & "];"
        response.write "    var values = [" & stayArray & "];"
        response.write "    var total = values.reduce((acc, val) => acc + val, 0);"
        response.write "    var percentages = values.map(val => ((val / total) * 100).toFixed(2) + '%');"
        response.write "    var ctx = document.getElementById('waitChart').getContext('2d');"
        response.write "    var pieChart = new Chart(ctx, {"
        response.write "      type: 'bar',"
        response.write "      data: {"
        response.write "        labels: labels,"
        response.write "        datasets: [{"
        response.write "          data: values,"
        response.write "          backgroundColor: ['#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff'],"
        response.write "        }]"
        response.write "      },"
        response.write "      options: {"
        response.write "        scales: {"
        response.write "          x: {"
        response.write "            ticks: {"
        response.write "              color: '#fff'"
        response.write "            }"
        response.write "          },"
        response.write "          y: {"
        response.write "            title: {"
        response.write "            display: true,"
        response.write "            text: 'Time(Hour)',"
        response.write "            color: '#fff'"
        response.write "          },"
        response.write "            ticks: {"
        response.write "              color: '#fff'"
        response.write "            }"
        response.write "          }"
        response.write "        },"
        response.write "        plugins: {"
        response.write "          legend: {"
        response.write "            display: false "
        response.write "          },"
        response.write "          tooltips: {"
        response.write "            callbacks: {"
        response.write "              label: (tooltipItem, data) => {"
        response.write "                var dataIndex = tooltipItem.index;"
        response.write "                return `${data.labels[dataIndex]}: ${data.datasets[0].data[dataIndex]} (${percentages[dataIndex]})`;"
        response.write "              }"
        response.write "            }"
        response.write "          }"
        response.write "        }"
        response.write "      }"
        response.write "    });"
        response.write "  </script>"
    End If
    
    rst.Close
    Set rst = Nothing

End Sub

Sub GetDptWaitTimeJS()

 Dim sql, rst
    
    Set rst = CreateObject("ADODB.recordset")
    
    sql = " SELECT TOP 10 "
    sql = sql & " SpecialistGroup.SpecialistGroupName,"
    sql = sql & " AVG(DateDiff(Hour, visitDate, EMRRequestItems.emrDate)) As avg_wait_time"
    sql = sql & " From"
    sql = sql & " Visitation"
    sql = sql & " Left Join"
    sql = sql & " SpecialistGroup ON Visitation.SpecialistGroupID = SpecialistGroup.SpecialistGroupID"
    sql = sql & " Left Join"
    sql = sql & " EMRRequestItems ON Visitation.VisitationID = EMRRequestItems.VisitationID"
    sql = sql & " Where"
    sql = sql & " EMRRequestItems.EMRDataID = 'TH060'"

    If (periodStart <> "") And (periodEnd <> "") Then
    sql = sql & " AND Visitation.VisitDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If

    sql = sql & " Group By"
    sql = sql & " SpecialistGroup.SpecialistGroupName"
    sql = sql & " ORDER BY avg_wait_time DESC"
                    
    rst.open qryPro.FltQry(sql), conn, 3, 4
    
    Dim appointArray, totalpArray
    appointArray = ""
    totalpArray = ""
        
        If rst.RecordCount > 0 Then
    rst.MoveFirst
    Do While Not rst.EOF
        If appointArray <> "" Then appointArray = appointArray & ","
        If totalpArray <> "" Then totalpArray = totalpArray & ","
        
        appointArray = appointArray & "'" & rst.fields("SpecialistGroupName") & "'"
        totalpArray = totalpArray & rst.fields("avg_wait_time")
        
        rst.MoveNext
    Loop
    
    response.write "  <script>"
    response.write "    var labels = [" & appointArray & "];"
    response.write "    var values = [" & totalpArray & "];"
    response.write "    var total = values.reduce((acc, val) => acc + val, 0);"
    response.write "    var percentages = values.map(val => ((val / total) * 100).toFixed(2) + '%');"
    response.write "    var ctx = document.getElementById('dptWaitTime').getContext('2d');"
    response.write "    var pieChart = new Chart(ctx, {"
    response.write "      type: 'bar',"
    response.write "      data: {"
    response.write "        labels: labels,"
    response.write "        datasets: [{"
    response.write "          data: values,"
    response.write "          backgroundColor: ['#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff'],"
    response.write "        }]"
    response.write "      },"
    response.write "      options: {"
    response.write "        scales: {"
    response.write "          x: {"
    response.write "            ticks: {"
    response.write "              color: '#fff'"
    response.write "            }"
    response.write "          },"
    response.write "          y: {"
    response.write "            title: {"
    response.write "            display: true,"
    response.write "            text: 'Time(Hour)',"
    response.write "            color: '#fff'"
    response.write "          },"
    response.write "            ticks: {"
    response.write "              color: '#fff'"
    response.write "            }"
    response.write "          }"
    response.write "        },"
    response.write "        plugins: {"
    response.write "          legend: {"
    response.write "            display: false "
    response.write "          },"
    response.write "          tooltips: {"
    response.write "            callbacks: {"
    response.write "              label: (tooltipItem, data) => {"
    response.write "                var dataIndex = tooltipItem.index;"
    response.write "                return `${data.labels[dataIndex]}: ${data.datasets[0].data[dataIndex]} (${percentages[dataIndex]})`;"
    response.write "              }"
    response.write "            }"
    response.write "          }"
    response.write "        }"
    response.write "      }"
    response.write "    });"
    response.write "  </script>"
 End If

    rst.Close
    Set rst = Nothing
   
End Sub

Sub AvgDptJS()


 Dim sql, rst
    
    Set rst = CreateObject("ADODB.recordset")
    
    sql = "SELECT SpecialistGroup.SpecialistGroupName AS Departments, AVG(DATEDIFF(DAY, AdmissionDate, DischargeDate)) AS average_stay "
    sql = sql & "FROM Admission "
    sql = sql & "JOIN Visitation ON Visitation.VisitationID = Admission.VisitationID "
    sql = sql & "JOIN SpecialistGroup ON SpecialistGroup.SpecialistGroupID = Visitation.SpecialistGroupID "
    If (periodStart <> "") And (periodEnd <> "") Then
        sql = sql & " AND Admission.AdmissionDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If
    sql = sql & "WHERE DischargeDate Is Not Null "
    sql = sql & "GROUP BY SpecialistGroup.SpecialistGroupName"
    sql = sql & " ORDER BY average_stay DESC "
    
    rst.open qryPro.FltQry(sql), conn, 3, 4
    
    Dim departmentsArray, stayArray
    departmentsArray = ""
    stayArray = ""
    
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            If departmentsArray <> "" Then departmentsArray = departmentsArray & ","
            If stayArray <> "" Then stayArray = stayArray & ","
            
            departmentsArray = departmentsArray & "'" & rst.fields("Departments") & "'"
            stayArray = stayArray & rst.fields("average_stay")
            
            rst.MoveNext
        Loop
        
        response.write "  <script>"
        response.write "    var labels = [" & departmentsArray & "];"
        response.write "    var values = [" & stayArray & "];"
        response.write "    var total = values.reduce((acc, val) => acc + val, 0);"
        response.write "    var percentages = values.map(val => ((val / total) * 100).toFixed(2) + '%');"
        response.write "    var ctx = document.getElementById('pieChart2').getContext('2d');"
        response.write "    var pieChart = new Chart(ctx, {"
        response.write "      type: 'bar',"
        response.write "      data: {"
        response.write "        labels: labels,"
        response.write "        datasets: [{"
        response.write "          data: values,"
        response.write "          backgroundColor: ['#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff'],"
        response.write "        }]"
        response.write "      },"
        response.write "      options: {"
        response.write "        scales: {"
        response.write "          x: {"
        response.write "            ticks: {"
        response.write "              color: '#fff'"
        response.write "            }"
        response.write "          },"
        response.write "          y: {"
        response.write "title: {"
        response.write "            display: true,"
        response.write "            text: 'Days',"
        response.write "            color: '#fff'"
        response.write "          },"
        response.write "            ticks: {"
        response.write "              color: '#fff'"
        response.write "            }"
        response.write "          }"
        response.write "        },"
        response.write "        plugins: {"
        response.write "          legend: {"
        response.write "            display: false "
        response.write "          },"
        response.write "          tooltips: {"
        response.write "            callbacks: {"
        response.write "              label: (tooltipItem, data) => {"
        response.write "                var dataIndex = tooltipItem.index;"
        response.write "                return `${data.labels[dataIndex]}: ${data.datasets[0].data[dataIndex]} (${percentages[dataIndex]})`;"
        response.write "              }"
        response.write "            }"
        response.write "          }"
        response.write "        }"
        response.write "      }"
        response.write "    });"
        response.write "  </script>"
    End If
    
    rst.Close
    Set rst = Nothing


End Sub

Sub AvgWdJS()

 Dim sql, rst
    
    Set rst = CreateObject("ADODB.recordset")
    
        sql = " SELECT wardname, AVG(days) AS Average_Stay"
        sql = sql & " FROM( "
        sql = sql & " SELECT wardid, visitationid, DATEDIFF(DAY, min(Admission.admissiondate), max(Admission.dischargedate)) AS days "
        sql = sql & " From Admission "
        
        If (periodStart <> "") And (periodEnd <> "") Then
        sql = sql & " WHERE AdmissionDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
        End If

        sql = sql & " GROUP BY wardid, visitationid "
        sql = sql & " )"
        sql = sql & " avgdays INNER JOIN Ward ON Ward.wardid=avgdays.wardid "
        sql = sql & " GROUP BY Ward.wardname "
        sql = sql & " ORDER BY average_stay DESC"
    
    rst.open qryPro.FltQry(sql), conn, 3, 4
    
    Dim wardssArray, stayArray
    wardssArray = ""
    stayArray = ""
    
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            If wardssArray <> "" Then wardssArray = wardssArray & ","
            If stayArray <> "" Then stayArray = stayArray & ","
            
            wardssArray = wardssArray & "'" & rst.fields("wardname") & "'"
            stayArray = stayArray & rst.fields("average_stay")
            
            rst.MoveNext
        Loop
        response.write "  <script>"
        response.write "    var labels = [" & wardssArray & "];"
        response.write "    var values = [" & stayArray & "];"
        response.write "    var total = values.reduce((acc, val) => acc + val, 0);"
        response.write "    var percentages = values.map(val => ((val / total) * 100).toFixed(2) + '%');"
        response.write "    var ctx = document.getElementById('pieChart3').getContext('2d');"
        response.write "    var pieChart = new Chart(ctx, {"
        response.write "      type: 'bar',"
        response.write "      data: {"
        response.write "        labels: labels,"
        response.write "        datasets: [{"
        response.write "          data: values,"
        response.write "          backgroundColor: ['#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff'],"
        response.write "        }]"
        response.write "      },"
        response.write "      options: {"
        response.write "        scales: {"
        response.write "          x: {"
        response.write "            ticks: {"
        response.write "              color: '#fff'"
        response.write "            }"
        response.write "          },"
        response.write "          y: {"
        response.write "            title: {"
        response.write "            display: true,"
        response.write "            text: 'Days',"
        response.write "            color: '#fff'"
        response.write "          },"
        response.write "            ticks: {"
        response.write "              color: '#fff'"
        response.write "            }"
        response.write "          }"
        response.write "        },"
        response.write "        plugins: {"
        response.write "          legend: {"
        response.write "            display: false "
        response.write "          },"
        response.write "          tooltips: {"
        response.write "            callbacks: {"
        response.write "              label: (tooltipItem, data) => {"
        response.write "                var dataIndex = tooltipItem.index;"
        response.write "                return `${data.labels[dataIndex]}: ${data.datasets[0].data[dataIndex]} (${percentages[dataIndex]})`;"
        response.write "              }"
        response.write "            }"
        response.write "          }"
        response.write "        }"
        response.write "      }"
        response.write "    });"
        response.write "  </script>"
    End If
    
    rst.Close
    Set rst = Nothing


End Sub

Sub DrugCostPerStayDptJS()


 Dim sql, rst
    
    Set rst = CreateObject("ADODB.recordset")
    
     sql = " SELECT SpecialistGroupName As Departments, SUM(totalcost) AS totalcost"
     sql = sql & " FROM ( "
     sql = sql & " SELECT SpecialistGroup.SpecialistGroupName, SUM(FinalAmt) AS totalcost "
     sql = sql & " From DrugSaleItems "
     sql = sql & " LEFT JOIN Visitation ON DrugSaleItems.VisitationID = Visitation.VisitationID "
     sql = sql & " LEFT JOIN SpecialistGroup ON SpecialistGroup.SpecialistGroupID = Visitation.SpecialistGroupID "
     sql = sql & " LEFT JOIN Admission ON Admission.VisitationID = Visitation.VisitationID "
        
    If (periodStart <> "") And (periodEnd <> "") Then
        sql = sql & " WHERE DrugSaleItems.DispenseDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If

     sql = sql & " GROUP BY SpecialistGroup.SpecialistGroupName "
     sql = sql & " Union All "
     sql = sql & " SELECT SpecialistGroup.SpecialistGroupName, SUM(DispenseAmt2) AS totalcost "
     sql = sql & " From DrugSaleItems2 "
     sql = sql & " LEFT JOIN Visitation ON DrugSaleItems2.VisitationID = Visitation.VisitationID "
     sql = sql & " LEFT JOIN SpecialistGroup ON SpecialistGroup.SpecialistGroupID = Visitation.SpecialistGroupID "
     sql = sql & " LEFT JOIN Admission ON Admission.VisitationID = Visitation.VisitationID "
        
    If (periodStart <> "") And (periodEnd <> "") Then
        sql = sql & " WHERE DrugSaleItems2.DispenseDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If

     sql = sql & " GROUP BY SpecialistGroup.SpecialistGroupName "
     sql = sql & " ) AS subquery "
     sql = sql & " GROUP BY SpecialistGroupName "
     sql = sql & " ORDER BY totalcost DESC "
    
    rst.open qryPro.FltQry(sql), conn, 3, 4
    
    Dim departmentsArray, totalcostArray
    departmentsArray = ""
    totalcostArray = ""
    
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            If departmentsArray <> "" Then departmentsArray = departmentsArray & ","
            If totalcostArray <> "" Then totalcostArray = totalcostArray & ","
            
            departmentsArray = departmentsArray & "'" & rst.fields("Departments") & "'"
            totalcostArray = totalcostArray & rst.fields("totalcost")
            
            rst.MoveNext
        Loop
        
        response.write "  <script>"
        response.write "    var labels = [" & departmentsArray & "];"
        response.write "    var values = [" & totalcostArray & "];"
        response.write "    var total = values.reduce((acc, val) => acc + val, 0);"
        response.write "    var percentages = values.map(val => ((val / total) * 100).toFixed(2) + '%');"
        response.write "    var ctx = document.getElementById('pieChart4').getContext('2d');"
        response.write "    var pieChart = new Chart(ctx, {"
        response.write "      type: 'bar',"
        response.write "      data: {"
        response.write "        labels: labels,"
        response.write "        datasets: [{"
        response.write "          data: values,"
        response.write "          backgroundColor: ['#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff'],"
        response.write "        }]"
        response.write "      },"
        response.write "      options: {"
        response.write "        scales: {"
        response.write "          x: {"
        response.write "            ticks: {"
        response.write "              color: '#fff'"
        response.write "            }"
        response.write "          },"
        response.write "          y: {"
        response.write "            title: {"
        response.write "            display: true,"
        response.write "            text: 'Total Cost',"
        response.write "            color: '#fff'"
        response.write "          },"
        response.write "            ticks: {"
        response.write "              color: '#fff'"
        response.write "            }"
        response.write "          }"
        response.write "        },"
        response.write "        plugins: {"
        response.write "          legend: {"
        response.write "            display: false "
        response.write "          },"
        response.write "          tooltips: {"
        response.write "            callbacks: {"
        response.write "              label: (tooltipItem, data) => {"
        response.write "                var dataIndex = tooltipItem.index;"
        response.write "                return `${data.labels[dataIndex]}: ${data.datasets[0].data[dataIndex]} (${percentages[dataIndex]})`;"
        response.write "              }"
        response.write "            }"
        response.write "          }"
        response.write "        }"
        response.write "      }"
        response.write "    });"
        response.write "  </script>"
    End If
    
    rst.Close
    Set rst = Nothing


End Sub

Sub GenderChartJS()

 Dim sql, rst

    Set rst = CreateObject("ADODB.recordset")

     sql = "SELECT Gender.GenderName, COUNT(PatientID) AS totPat FROM Patient"
     sql = sql & " INNER JOIN Gender ON Patient.GenderID = Gender.GenderID"
 
     If (periodStart <> "") And (periodEnd <> "") Then
         sql = sql & " WHERE Patient.FirstVisitDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
     End If
    
    sql = sql & " GROUP BY Gender.GenderName"

    rst.open qryPro.FltQry(sql), conn, 3, 4

    Dim genderArray, totalgArray
    genderArray = ""
    totalgArray = ""

    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            If genderArray <> "" Then genderArray = genderArray & ","
            If totalgArray <> "" Then totalgArray = totalgArray & ","

            genderArray = genderArray & "'" & rst.fields("GenderName") & "'"
           totalgArray = totalgArray & rst.fields("totPat")

            rst.MoveNext
        Loop

        response.write "  <script>"
        response.write "    var labels = [" & genderArray & "];"
        response.write "    var values = [" & totalgArray & "];"
        response.write "    var total = values.reduce((acc, val) => acc + val, 0);"
        response.write "    var percentages = values.map(val => ((val / total) * 100).toFixed(2) + '%');"
        response.write "    var ctx = document.getElementById('pieChart5').getContext('2d');"
        response.write "    var pieChart = new Chart(ctx, {"
        response.write "      type: 'pie',"
        response.write "      data: {"
        response.write "        labels: labels,"
        response.write "        datasets: [{"
        response.write "          data: values,"
        response.write "          backgroundColor: ['#FF5733', '#36A2EB', '#FFC300', '#d50000', '#FF5733'],"
        response.write "        }]"
        response.write "      },"
        response.write "      options: {"
        response.write "        tooltips: {"
        response.write "          callbacks: {"
        response.write "            label: (tooltipItem, data) => {"
        response.write "              var dataIndex = tooltipItem.index;"
        response.write "              return `${data.labels[dataIndex]}: ${data.datasets[0].data[dataIndex]} (${percentages[dataIndex]})`;"
        response.write "            }"
        response.write "          },"
        response.write "          backgroundColor: 'rgba(0, 0, 0, 0.7)',"  ' Background color of tooltip
        response.write "          titleFontColor: '#fff',"  ' Font color of tooltip title
        response.write "          bodyFontColor: '#fff',"   ' Font color of tooltip content
        response.write "          footerFontColor: '#fff',"  ' Font color of tooltip footer
        response.write "        },"
        response.write "        legend: {"
        response.write "          labels: {"
        response.write "            fontColor: '#fff'"
        response.write "          }"
        response.write "        }"
        response.write "      }"
        response.write "    });"
        response.write "  </script>"
    End If

    rst.Close
    Set rst = Nothing


End Sub

Sub AppChartJS()

 Dim sql, rst
    
    Set rst = CreateObject("ADODB.recordset")
    
    sql = " SELECT  AppointmentStatus.AppointmentStatusName, COUNT(appointment.patientID) AS totalpatients"
    sql = sql & " From Appointment"
    sql = sql & " left JOIN AppointmentStatus ON AppointmentStatus.AppointmentStatusID = Appointment.AppointmentStatusID"

    If (periodStart <> "") And (periodEnd <> "") Then
    sql = sql & " WHERE Appointment.AppointDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If
    
    sql = sql & " GROUP BY AppointmentStatus.AppointmentStatusName"
    sql = sql & " ORDER BY totalpatients DESC "
                    
    rst.open qryPro.FltQry(sql), conn, 3, 4
    
    Dim appointArray, totalpArray
    appointArray = ""
    totalpArray = ""
    
        If rst.RecordCount > 0 Then
         rst.MoveFirst
         Do While Not rst.EOF
             If appointArray <> "" Then appointArray = appointArray & ","
             If totalpArray <> "" Then totalpArray = totalpArray & ","
             
             appointArray = appointArray & "'" & rst.fields("AppointmentStatusName") & "'"
             totalpArray = totalpArray & rst.fields("totalpatients")
             
             rst.MoveNext
         Loop
            
            
            response.write "  <script>"
            response.write "    var labels = [" & appointArray & "];"
            response.write "    var values = [" & totalpArray & "];"
            response.write "    var total = values.reduce((acc, val) => acc + val, 0);"
            response.write "    var percentages = values.map(val => ((val / total) * 100).toFixed(2) + '%');"
            response.write "    var ctx = document.getElementById('CancelChart').getContext('2d');"
            response.write "    var pieChart = new Chart(ctx, {"
            response.write "      type: 'pie',"
            response.write "      data: {"
            response.write "        labels: labels,"
            response.write "        datasets: [{"
            response.write "          data: values,"
            response.write "          backgroundColor: ['#FF5733', '#36A2EB', '#FFC300', '#d50000', '#FF5733'],"
            response.write "        }]"
            response.write "      },"
            response.write "      options: {"
            response.write "        tooltips: {"
            response.write "          callbacks: {"
            response.write "            label: (tooltipItem, data) => {"
            response.write "              var dataIndex = tooltipItem.index;"
            response.write "              return `${data.labels[dataIndex]}: ${data.datasets[0].data[dataIndex]} (${percentages[dataIndex]})`;"
            response.write "            }"
            response.write "          },"
            response.write "          backgroundColor: 'rgba(0, 0, 0, 0.7)',"  ' Background color of tooltip
            response.write "          titleFontColor: '#fff',"  ' Font color of tooltip title
            response.write "          bodyFontColor: '#fff',"   ' Font color of tooltip content
            response.write "          footerFontColor: '#fff',"  ' Font color of tooltip footer
            response.write "        },"
            response.write "        legend: {"
            response.write "          labels: {"
            response.write "            fontColor: '#fff'"
            response.write "          }"
            response.write "        }"
            response.write "      }"
            response.write "    });"
            response.write "  </script>"
    
     End If
        
        rst.Close
        Set rst = Nothing
            


End Sub

Sub ReadmissionRate()

 Dim sql, rst
    
    Set rst = CreateObject("ADODB.recordset")
    
    sql = "WITH myTable AS ("
    sql = sql & " select "
    sql = sql & " dbo.GetAgeLabel(V.PatientAge) AS AgeGroup,"
    sql = sql & " DATENAME(MONTH,a.AdmissionDate) [Month],"
    sql = sql & " MONTH(a.AdmissionDate) [MonthNum],"
    sql = sql & " count(dbo.GetAgeLabel(V.PatientAge)) frequency"
    sql = sql & " from admission a join visitation v"
    sql = sql & " on v.VisitationID = a.VisitationID"

    If (periodStart <> "") And (periodEnd <> "") Then
        sql = sql & " where a.AdmissionDate between '" & periodStart & "' and '" & periodEnd & "'"
    End If

    sql = sql & " group by dbo.GetAgeLabel(V.PatientAge),"
    sql = sql & " DATENAME(MONTH,a.AdmissionDate),"
    sql = sql & " MONTH(a.AdmissionDate)"
    'sql = sql & " --order by dbo.GetAgeLabel(V.PatientAge) DESC,"
    'sql = sql & " --MONTH(a.AdmissionDate) asc"
    sql = sql & " )"
    sql = sql & " SELECT [Month], SUM(frequency) AS Frequency"
    sql = sql & " FROM myTable"
    sql = sql & " GROUP BY [Month], [MonthNum]"
    sql = sql & " ORDER BY [MonthNum]  "
            
    rst.open qryPro.FltQry(sql), conn, 3, 4
    
    Dim SuppArray, totalAmtArray
    SuppArray = ""
    totalAmtArray = ""
    
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            If SuppArray <> "" Then SuppArray = SuppArray & ","
            If totalAmtArray <> "" Then totalAmtArray = totalAmtArray & ","
            
            SuppArray = SuppArray & "'" & rst.fields("Month") & "'"
            totalAmtArray = totalAmtArray & rst.fields("Frequency")
            
            rst.MoveNext
        Loop

        response.write "  <script>"
        response.write "    var labels = [" & SuppArray & "];"
        response.write "    var values = [" & totalAmtArray & "];"
        response.write "    var total = values.reduce((acc, val) => acc + val, 0);"
        response.write "    var percentages = values.map(val => ((val / total) * 100).toFixed(2) + '%');"
        response.write "    var ctx = document.getElementById('readmissionRate').getContext('2d');"
        response.write "    var pieChart = new Chart(ctx, {"
        response.write "      type: 'bar',"
        response.write "      data: {"
        response.write "        labels: labels,"
        response.write "        datasets: [{"
        response.write "          data: values,"
        response.write "          backgroundColor: ['#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff'],"
        response.write "        }]"
        response.write "      },"
        response.write "      options: {"
        response.write "        scales: {"
        response.write "          x: {"
        response.write "title: {"
        response.write "            display: true,"
        response.write "            text: 'Month',"
        response.write "            color: '#fff'"
        response.write "          },"
        response.write "            ticks: {"
        response.write "              color: '#fff'"
        response.write "            }"
        response.write "          },"
        response.write "          y: {"
        response.write "title: {"
        response.write "            display: true,"
        response.write "            text: 'Frequency',"
        response.write "            color: '#fff'"
        response.write "          },"
        response.write "            ticks: {"
        response.write "              color: '#fff'"
        response.write "            }"
        response.write "          }"
        response.write "        },"
        response.write "        plugins: {"
        response.write "          legend: {"
        response.write "            display: false "
        response.write "          },"
        response.write "          tooltips: {"
        response.write "            callbacks: {"
        response.write "              label: (tooltipItem, data) => {"
        response.write "                var dataIndex = tooltipItem.index;"
        response.write "                return `${data.labels[dataIndex]}: ${data.datasets[0].data[dataIndex]} (${percentages[dataIndex]})`;"
        response.write "              }"
        response.write "            }"
        response.write "          }"
        response.write "        }"
        response.write "      }"
        response.write "    });"
        response.write "  </script>"
 End If
    
    rst.Close
    Set rst = Nothing


End Sub

Sub SupplierChartJS()

 Dim sql, rst
    
    Set rst = CreateObject("ADODB.recordset")
    
    sql = " SELECT TOP 10  s.SupplierName, sum(TotalCost) AS amt"
    sql = sql & " From IncomingDrugItems "
    sql = sql & " LEFT JOIN Supplier s ON IncomingDrugItems.SupplierID = s.SupplierID "

    If (periodStart <> "") And (periodEnd <> "") Then
        sql = sql & " WHERE IncomingDrugItems.EntryDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If
    
    sql = sql & " GROUP BY s.SupplierID, s.SupplierName "
    sql = sql & " ORDER BY amt desc "
            
    rst.open qryPro.FltQry(sql), conn, 3, 4
    
    Dim SuppArray, totalAmtArray
    SuppArray = ""
    totalAmtArray = ""
    
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            If SuppArray <> "" Then SuppArray = SuppArray & ","
            If totalAmtArray <> "" Then totalAmtArray = totalAmtArray & ","
            
            SuppArray = SuppArray & "'" & rst.fields("SupplierName") & "'"
            totalAmtArray = totalAmtArray & rst.fields("amt")
            
            rst.MoveNext
        Loop

        response.write "  <script>"
        response.write "    var labels = [" & SuppArray & "];"
        response.write "    var values = [" & totalAmtArray & "];"
        response.write "    var total = values.reduce((acc, val) => acc + val, 0);"
        response.write "    var percentages = values.map(val => ((val / total) * 100).toFixed(2) + '%');"
        response.write "    var ctx = document.getElementById('pieChart6').getContext('2d');"
        response.write "    var pieChart = new Chart(ctx, {"
        response.write "      type: 'bar',"
        response.write "      data: {"
        response.write "        labels: labels,"
        response.write "        datasets: [{"
        response.write "          data: values,"
        response.write "          backgroundColor: ['#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff'],"
        response.write "        }]"
        response.write "      },"
        response.write "      options: {"
        response.write "        scales: {"
        response.write "          x: {"
        response.write "            ticks: {"
        response.write "              color: '#fff'"
        response.write "            }"
        response.write "          },"
        response.write "          y: {"
        response.write "title: {"
        response.write "            display: true,"
        response.write "            text: 'Total Cost',"
        response.write "            color: '#fff'"
        response.write "          },"
        response.write "            ticks: {"
        response.write "              color: '#fff'"
        response.write "            }"
        response.write "          }"
        response.write "        },"
        response.write "        plugins: {"
        response.write "          legend: {"
        response.write "            display: false "
        response.write "          },"
        response.write "          tooltips: {"
        response.write "            callbacks: {"
        response.write "              label: (tooltipItem, data) => {"
        response.write "                var dataIndex = tooltipItem.index;"
        response.write "                return `${data.labels[dataIndex]}: ${data.datasets[0].data[dataIndex]} (${percentages[dataIndex]})`;"
        response.write "              }"
        response.write "            }"
        response.write "          }"
        response.write "        }"
        response.write "      }"
        response.write "    });"
        response.write "  </script>"
 End If
    
    rst.Close
    Set rst = Nothing


End Sub

Sub SupplierDrugsJS()

 Dim sql, rst
    
    Set rst = CreateObject("ADODB.recordset")
    
    sql = " SELECT TOP 10  s.SupplierName, COUNT(DrugID) AS totalDrugs"
    sql = sql & " From IncomingDrugItems "
    sql = sql & " LEFT JOIN Supplier s ON IncomingDrugItems.SupplierID = s.SupplierID "

    If (periodStart <> "") And (periodEnd <> "") Then
        sql = sql & " WHERE IncomingDrugItems.EntryDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If

    sql = sql & " GROUP BY s.SupplierID, s.SupplierName "
    sql = sql & " ORDER BY totalDrugs desc "
            
    rst.open qryPro.FltQry(sql), conn, 3, 4
    
    Dim SuppArray, totalAmtArray
    SuppArray = ""
    totalAmtArray = ""
    
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            If SuppArray <> "" Then SuppArray = SuppArray & ","
            If totalAmtArray <> "" Then totalAmtArray = totalAmtArray & ","
            
            SuppArray = SuppArray & "'" & rst.fields("SupplierName") & "'"
            totalAmtArray = totalAmtArray & rst.fields("totalDrugs")
            
            rst.MoveNext
        Loop
        
        response.write "  <script>"
        response.write "    var labels = [" & SuppArray & "];"
        response.write "    var values = [" & totalAmtArray & "];"
        response.write "    var total = values.reduce((acc, val) => acc + val, 0);"
        response.write "    var percentages = values.map(val => ((val / total) * 100).toFixed(2) + '%');"
        response.write "    var ctx = document.getElementById('suppDg-chart').getContext('2d');"
        response.write "    var pieChart = new Chart(ctx, {"
        response.write "      type: 'bar',"
        response.write "      data: {"
        response.write "        labels: labels,"
        response.write "        datasets: [{"
        response.write "          data: values,"
        response.write "          backgroundColor: ['#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff'],"
        response.write "        }]"
        response.write "      },"
        response.write "      options: {"
        response.write "        scales: {"
        response.write "          x: {"
        response.write "            ticks: {"
        response.write "              color: '#fff'"
        response.write "            }"
        response.write "          },"
        response.write "          y: {"
        response.write "title: {"
        response.write "            display: true,"
        response.write "            text: 'Number of Drugs',"
        response.write "            color: '#fff'"
        response.write "          },"
        response.write "            ticks: {"
        response.write "              color: '#fff'"
        response.write "            }"
        response.write "          }"
        response.write "        },"
        response.write "        plugins: {"
        response.write "          legend: {"
        response.write "            display: false "
        response.write "          },"
        response.write "          tooltips: {"
        response.write "            callbacks: {"
        response.write "              label: (tooltipItem, data) => {"
        response.write "                var dataIndex = tooltipItem.index;"
        response.write "                return `${data.labels[dataIndex]}: ${data.datasets[0].data[dataIndex]} (${percentages[dataIndex]})`;"
        response.write "              }"
        response.write "            }"
        response.write "          }"
        response.write "        }"
        response.write "      }"
        response.write "    });"
        response.write "  </script>"
 End If
    
    rst.Close
    Set rst = Nothing

End Sub

Sub mostSoldDrugs()

 Dim sql, rst
    
    Set rst = CreateObject("ADODB.recordset")
    
    sql = "SELECT TOP 10 DrugID, DrugName, SUM(TotalAmount) AS Amount "
    sql = sql & " FROM ( "
    sql = sql & "     SELECT DrugSaleItems.DrugID, Drug.DrugName, SUM(finalAmt) AS TotalAmount "
    sql = sql & "     FROM DrugSaleItems "
    sql = sql & "     JOIN Drug ON Drug.DrugID = DrugSaleItems.drugID "
    sql = sql & "     WHERE DrugSaleItems.BillGroupID IN ('BG003', 'B15') "

    If (periodStart <> "") And (periodEnd <> "") Then
      sql = sql & " AND DrugSaleItems.DispenseDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If
    
    sql = sql & "     GROUP BY DrugSaleItems.DrugID, Drug.DrugName "
    sql = sql & "     UNION ALL  "
    sql = sql & "     SELECT DrugSaleItems2.DrugID, Drug.DrugName, SUM(finalAmt) AS TotalAmount "
    sql = sql & "     FROM DrugSaleItems2 "
    sql = sql & "     JOIN Drug ON Drug.DrugID = DrugSaleItems2.drugID "
    sql = sql & "     WHERE DrugSaleItems2.BillGroupID IN ('BG003', 'B15') "

    If (periodStart <> "") And (periodEnd <> "") Then
      sql = sql & " AND DrugSaleItems2.DispenseDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If

    sql = sql & "     GROUP BY DrugSaleItems2.DrugID, Drug.DrugName "
    sql = sql & " ) AS subquery  "
    sql = sql & " GROUP BY DrugID, DrugName, TotalAmount "
    sql = sql & " ORDER BY Amount DESC"
            
    rst.open qryPro.FltQry(sql), conn, 3, 4
    
    Dim SuppArray, totalAmtArray
    SuppArray = ""
    totalAmtArray = ""
    
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            If SuppArray <> "" Then SuppArray = SuppArray & ","
            If totalAmtArray <> "" Then totalAmtArray = totalAmtArray & ","
            
            SuppArray = SuppArray & "'" & rst.fields("DrugName") & "'"
            totalAmtArray = totalAmtArray & rst.fields("Amount")
            
            rst.MoveNext
        Loop
        
        response.write "  <script>"
        response.write "    var labels = [" & SuppArray & "];"
        response.write "    var values = [" & totalAmtArray & "];"
        response.write "    var total = values.reduce((acc, val) => acc + val, 0);"
        response.write "    var percentages = values.map(val => ((val / total) * 100).toFixed(2) + '%');"
        response.write "    var ctx = document.getElementById('soldDrugs').getContext('2d');"
        response.write "    var pieChart = new Chart(ctx, {"
        response.write "      type: 'bar',"
        response.write "      data: {"
        response.write "        labels: labels,"
        response.write "        datasets: [{"
        response.write "          data: values,"
        response.write "          backgroundColor: ['#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff'],"
        response.write "        }]"
        response.write "      },"
        response.write "      options: {"
        response.write "        scales: {"
        response.write "          x: {"
        response.write "            ticks: {"
        response.write "              color: '#fff'"
        response.write "            }"
        response.write "          },"
        response.write "          y: {"
        response.write "title: {"
        response.write "            display: true,"
        response.write "            text: 'Amount',"
        response.write "            color: '#fff'"
        response.write "          },"
        response.write "            ticks: {"
        response.write "              color: '#fff'"
        response.write "            }"
        response.write "          }"
        response.write "        },"
        response.write "        plugins: {"
        response.write "          legend: {"
        response.write "            display: false "
        response.write "          },"
        response.write "          tooltips: {"
        response.write "            callbacks: {"
        response.write "              label: (tooltipItem, data) => {"
        response.write "                var dataIndex = tooltipItem.index;"
        response.write "                return `${data.labels[dataIndex]}: ${data.datasets[0].data[dataIndex]} (${percentages[dataIndex]})`;"
        response.write "              }"
        response.write "            }"
        response.write "          }"
        response.write "        }"
        response.write "      }"
        response.write "    });"
        response.write "  </script>"
    End If
    
    rst.Close
    Set rst = Nothing


End Sub

Sub mostPrescribedDrugByQuantity()

 Dim sql, rst
    
    Set rst = CreateObject("ADODB.recordset")

    sql = "SELECT TOP 10 PrescriptionName, COUNT(drugID) AS Quantity"
    sql = sql & " FROM Prescription"

    If (periodStart <> "") And (periodEnd <> "") Then
      sql = sql & " WHERE PrescriptionDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If
    
    sql = sql & " GROUP BY drugID, PrescriptionName"
    sql = sql & " ORDER BY Quantity DESC"
            
    rst.open qryPro.FltQry(sql), conn, 3, 4
    
    Dim SuppArray, totalAmtArray
    SuppArray = ""
    totalAmtArray = ""
    
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            If SuppArray <> "" Then SuppArray = SuppArray & ","
            If totalAmtArray <> "" Then totalAmtArray = totalAmtArray & ","
            
            SuppArray = SuppArray & "'" & rst.fields("PrescriptionName") & "'"
            totalAmtArray = totalAmtArray & rst.fields("Quantity")
            
            rst.MoveNext
        Loop
        
        response.write "  <script>"
        response.write "    var labels = [" & SuppArray & "];"
        response.write "    var values = [" & totalAmtArray & "];"
        response.write "    var total = values.reduce((acc, val) => acc + val, 0);"
        response.write "    var percentages = values.map(val => ((val / total) * 100).toFixed(2) + '%');"
        response.write "    var ctx = document.getElementById('prescribedDrugByQuantity').getContext('2d');"
        response.write "    var pieChart = new Chart(ctx, {"
        response.write "      type: 'bar',"
        response.write "      data: {"
        response.write "        labels: labels,"
        response.write "        datasets: [{"
        response.write "          data: values,"
        response.write "          backgroundColor: ['#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff'],"
        response.write "        }]"
        response.write "      },"
        response.write "      options: {"
        response.write "        scales: {"
        response.write "          x: {"
        response.write "            ticks: {"
        response.write "              color: '#fff'"
        response.write "            }"
        response.write "          },"
        response.write "          y: {"
        response.write "title: {"
        response.write "            display: true,"
        response.write "            text: 'Quantity',"
        response.write "            color: '#fff'"
        response.write "          },"
        response.write "            ticks: {"
        response.write "              color: '#fff'"
        response.write "            }"
        response.write "          }"
        response.write "        },"
        response.write "        plugins: {"
        response.write "          legend: {"
        response.write "            display: false "
        response.write "          },"
        response.write "          tooltips: {"
        response.write "            callbacks: {"
        response.write "              label: (tooltipItem, data) => {"
        response.write "                var dataIndex = tooltipItem.index;"
        response.write "                return `${data.labels[dataIndex]}: ${data.datasets[0].data[dataIndex]} (${percentages[dataIndex]})`;"
        response.write "              }"
        response.write "            }"
        response.write "          }"
        response.write "        }"
        response.write "      }"
        response.write "    });"
        response.write "  </script>"
    End If
    
    rst.Close
    Set rst = Nothing


End Sub

Sub mostPrescribedDrug()

 Dim sql, rst
    
    Set rst = CreateObject("ADODB.recordset")
    
    sql = "SELECT TOP 10 DrugID, DrugName, SUM(TotalAmount) AS Amount "
    sql = sql & " FROM ( "
    sql = sql & "     SELECT DrugSaleItems.DrugID, Drug.DrugName, COUNT(DrugSaleItems.DrugID) AS TotalAmount "
    sql = sql & "     FROM DrugSaleItems "
    sql = sql & "     JOIN Drug ON Drug.DrugID = DrugSaleItems.drugID "
    sql = sql & "     WHERE DrugSaleItems.billGroupID = 'BG003' "

    If (periodStart <> "") And (periodEnd <> "") Then
      sql = sql & " AND DrugSaleItems.DispenseDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If

    sql = sql & "     GROUP BY DrugSaleItems.DrugID, Drug.DrugName "
    sql = sql & "     UNION ALL  "
    sql = sql & "     SELECT DrugSaleItems2.DrugID, Drug.DrugName, COUNT(DrugSaleItems2.DrugID) AS TotalAmount "
    sql = sql & "     FROM DrugSaleItems2 "
    sql = sql & "     JOIN Drug ON Drug.DrugID = DrugSaleItems2.drugID "
    sql = sql & "     WHERE DrugSaleItems2.billGroupID = 'BG003' "

    If (periodStart <> "") And (periodEnd <> "") Then
      sql = sql & " AND DrugSaleItems2.DispenseDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If

    sql = sql & "     GROUP BY DrugSaleItems2.DrugID, Drug.DrugName "
    sql = sql & " ) AS subquery  "
    sql = sql & " GROUP BY DrugID, DrugName "
    sql = sql & " ORDER BY Amount DESC"
            
    rst.open qryPro.FltQry(sql), conn, 3, 4
    
    Dim SuppArray, totalAmtArray
    SuppArray = ""
    totalAmtArray = ""
    
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            If SuppArray <> "" Then SuppArray = SuppArray & ","
            If totalAmtArray <> "" Then totalAmtArray = totalAmtArray & ","
            
            SuppArray = SuppArray & "'" & rst.fields("DrugName") & "'"
            totalAmtArray = totalAmtArray & rst.fields("Amount")
            
            rst.MoveNext
        Loop
        
        response.write "  <script>"
        response.write "    var labels = [" & SuppArray & "];"
        response.write "    var values = [" & totalAmtArray & "];"
        response.write "    var total = values.reduce((acc, val) => acc + val, 0);"
        response.write "    var percentages = values.map(val => ((val / total) * 100).toFixed(2) + '%');"
        response.write "    var ctx = document.getElementById('prescribedDrug').getContext('2d');"
        response.write "    var pieChart = new Chart(ctx, {"
        response.write "      type: 'bar',"
        response.write "      data: {"
        response.write "        labels: labels,"
        response.write "        datasets: [{"
        response.write "          data: values,"
        response.write "          backgroundColor: ['#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff'],"
        response.write "        }]"
        response.write "      },"
        response.write "      options: {"
        response.write "        scales: {"
        response.write "          x: {"
        response.write "            ticks: {"
        response.write "              color: '#fff'"
        response.write "            }"
        response.write "          },"
        response.write "          y: {"
        response.write "title: {"
        response.write "            display: true,"
        response.write "            text: 'Number',"
        response.write "            color: '#fff'"
        response.write "          },"
        response.write "            ticks: {"
        response.write "              color: '#fff'"
        response.write "            }"
        response.write "          }"
        response.write "        },"
        response.write "        plugins: {"
        response.write "          legend: {"
        response.write "            display: false "
        response.write "          },"
        response.write "          tooltips: {"
        response.write "            callbacks: {"
        response.write "              label: (tooltipItem, data) => {"
        response.write "                var dataIndex = tooltipItem.index;"
        response.write "                return `${data.labels[dataIndex]}: ${data.datasets[0].data[dataIndex]} (${percentages[dataIndex]})`;"
        response.write "              }"
        response.write "            }"
        response.write "          }"
        response.write "        }"
        response.write "      }"
        response.write "    });"
        response.write "  </script>"
    End If
    
    rst.Close
    Set rst = Nothing


End Sub

Sub mostSoldDMartItems()

 Dim sql, rst
    
    Set rst = CreateObject("ADODB.recordset")
    
    sql = "SELECT TOP 10 DrugSaleItems.DrugID AS MartID, Drug.DrugName AS Mart_Item, SUM(finalAmt) AS TotalAmount "
    sql = sql & " FROM DrugSaleItems "
    sql = sql & " JOIN Drug ON Drug.DrugID = DrugSaleItems.drugID "
    sql = sql & " WHERE DrugSaleItems.BillGroupID = 'BG008' "

    If (periodStart <> "") And (periodEnd <> "") Then
      sql = sql & " AND DrugSaleItems.DispenseDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If
    
    sql = sql & " GROUP BY DrugSaleItems.DrugID, Drug.DrugName "
    sql = sql & " ORDER BY totalAmount DESC "
            
    rst.open qryPro.FltQry(sql), conn, 3, 4
    
    Dim SuppArray, totalAmtArray
    SuppArray = ""
    totalAmtArray = ""
    
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            If SuppArray <> "" Then SuppArray = SuppArray & ","
            If totalAmtArray <> "" Then totalAmtArray = totalAmtArray & ","
            
            SuppArray = SuppArray & "'" & rst.fields("Mart_Item") & "'"
            totalAmtArray = totalAmtArray & rst.fields("TotalAmount")
            
            rst.MoveNext
        Loop
        
        response.write "  <script>"
        response.write "    var labels = [" & SuppArray & "];"
        response.write "    var values = [" & totalAmtArray & "];"
        response.write "    var total = values.reduce((acc, val) => acc + val, 0);"
        response.write "    var percentages = values.map(val => ((val / total) * 100).toFixed(2) + '%');"
        response.write "    var ctx = document.getElementById('MartItems').getContext('2d');"
        response.write "    var pieChart = new Chart(ctx, {"
        response.write "      type: 'bar',"
        response.write "      data: {"
        response.write "        labels: labels,"
        response.write "        datasets: [{"
        response.write "          data: values,"
        response.write "          backgroundColor: ['#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff'],"
        response.write "        }]"
        response.write "      },"
        response.write "      options: {"
        response.write "        scales: {"
        response.write "          x: {"
        response.write "            ticks: {"
        response.write "              color: '#fff'"
        response.write "            }"
        response.write "          },"
        response.write "          y: {"
        response.write " title: {"
        response.write "            display: true,"
        response.write "            text: 'Total Cost',"
        response.write "            color: '#fff'"
        response.write "          },"
        response.write "            ticks: {"
        response.write "              color: '#fff'"
        response.write "            }"
        response.write "          }"
        response.write "        },"
        response.write "        plugins: {"
        response.write "          legend: {"
        response.write "            display: false "
        response.write "          },"
        response.write "          tooltips: {"
        response.write "            callbacks: {"
        response.write "              label: (tooltipItem, data) => {"
        response.write "                var dataIndex = tooltipItem.index;"
        response.write "                return `${data.labels[dataIndex]}: ${data.datasets[0].data[dataIndex]} (${percentages[dataIndex]})`;"
        response.write "              }"
        response.write "            }"
        response.write "          }"
        response.write "        }"
        response.write "      }"
        response.write "    });"
        response.write "  </script>"
 End If
    
    rst.Close
    Set rst = Nothing


End Sub

Sub mostVisitedDpt()

 Dim sql, rst
    
    Set rst = CreateObject("ADODB.recordset")
    
    sql = "SELECT TOP 10 SpecialistGroup.SpecialistGroupName, count(visitation.visitationID) AS Total_Number "
    sql = sql & " FROM Visitation "
    sql = sql & " JOIN SpecialistGroup ON SpecialistGroup.SpecialistGroupID = Visitation.SpecialistGroupID "

    If (periodStart <> "") And (periodEnd <> "") Then
      sql = sql & " WHERE Visitation.VisitDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If

    sql = sql & " GROUP BY SpecialistGroup.SpecialistGroupName "
    sql = sql & " ORDER BY Total_Number DESC"
            
    rst.open qryPro.FltQry(sql), conn, 3, 4
    
    Dim SuppArray, totalAmtArray
    SuppArray = ""
    totalAmtArray = ""
    
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            If SuppArray <> "" Then SuppArray = SuppArray & ","
            If totalAmtArray <> "" Then totalAmtArray = totalAmtArray & ","
            
            SuppArray = SuppArray & "'" & rst.fields("SpecialistGroupName") & "'"
            totalAmtArray = totalAmtArray & rst.fields("Total_Number")
            
            rst.MoveNext
        Loop
        
        response.write "  <script>"
        response.write "    var labels = [" & SuppArray & "];"
        response.write "    var values = [" & totalAmtArray & "];"
        response.write "    var total = values.reduce((acc, val) => acc + val, 0);"
        response.write "    var percentages = values.map(val => ((val / total) * 100).toFixed(2) + '%');"
        response.write "    var ctx = document.getElementById('visitedDpt').getContext('2d');"
        response.write "    var pieChart = new Chart(ctx, {"
        response.write "      type: 'bar',"
        response.write "      data: {"
        response.write "        labels: labels,"
        response.write "        datasets: [{"
        response.write "          data: values,"
        response.write "          backgroundColor: ['#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff'],"
        response.write "        }]"
        response.write "      },"
        response.write "      options: {"
        response.write "        scales: {"
        response.write "          x: {"
        response.write "            ticks: {"
        response.write "              color: '#fff'"
        response.write "            }"
        response.write "          },"
        response.write "          y: {"
        response.write "title: {"
        response.write "            display: true,"
        response.write "            text: 'Number of Visits',"
        response.write "            color: '#fff'"
        response.write "          },"
        response.write "            ticks: {"
        response.write "              color: '#fff'"
        response.write "            }"
        response.write "          }"
        response.write "        },"
        response.write "        plugins: {"
        response.write "          legend: {"
        response.write "            display: false "
        response.write "          },"
        response.write "          tooltips: {"
        response.write "            callbacks: {"
        response.write "              label: (tooltipItem, data) => {"
        response.write "                var dataIndex = tooltipItem.index;"
        response.write "                return `${data.labels[dataIndex]}: ${data.datasets[0].data[dataIndex]} (${percentages[dataIndex]})`;"
        response.write "              }"
        response.write "            }"
        response.write "          }"
        response.write "        }"
        response.write "      }"
        response.write "    });"
        response.write "  </script>"
 End If
    
    rst.Close
    Set rst = Nothing

End Sub

Sub MostPrescribedLab()

 Dim sql, rst
    
    Set rst = CreateObject("ADODB.recordset")
    
    sql = "SELECT TOP 10 LabTestName, LabTestID, SUM(TotalNumber) AS Number "
    sql = sql & " FROM ( "
    sql = sql & "     SELECT LabTest.LabTestName, investigation.LabTestID, COUNT(investigation.LabTestID) as TotalNumber "
    sql = sql & "     FROM Investigation "
    sql = sql & "     JOIN LabTest ON LabTest.LabTestID = Investigation.LabTestID "

    If (periodStart <> "") And (periodEnd <> "") Then
      sql = sql & " AND Investigation.RequestDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If

    sql = sql & "     GROUP BY Investigation.LabTestID, LabTest.LabTestName"
    sql = sql & "     UNION ALL "
    sql = sql & "     SELECT LabTest.LabTestName, Investigation2.LabTestID, COUNT(investigation2.LabTestID) as TotalNumber "
    sql = sql & "     FROM Investigation2 "
    sql = sql & "     JOIN LabTest ON LabTest.LabTestID = Investigation2.LabTestID "

    If (periodStart <> "") And (periodEnd <> "") Then
      sql = sql & " AND Investigation2.RequestDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If
    
    sql = sql & "     GROUP BY Investigation2.LabTestID, LabTest.LabTestName"
    sql = sql & " ) AS combined_data "
    sql = sql & " GROUP BY LabTestID, LabTestName "
    sql = sql & " ORDER BY Number DESC; "
            
    rst.open qryPro.FltQry(sql), conn, 3, 4
    
    Dim SuppArray, totalAmtArray
    SuppArray = ""
    totalAmtArray = ""
    
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            If SuppArray <> "" Then SuppArray = SuppArray & ","
            If totalAmtArray <> "" Then totalAmtArray = totalAmtArray & ","
            
            SuppArray = SuppArray & "'" & rst.fields("LabTestName") & "'"
            totalAmtArray = totalAmtArray & rst.fields("Number")
            
            rst.MoveNext
        Loop
        
        response.write "  <script>"
        response.write "    var labels = [" & SuppArray & "];"
        response.write "    var values = [" & totalAmtArray & "];"
        response.write "    var total = values.reduce((acc, val) => acc + val, 0);"
        response.write "    var percentages = values.map(val => ((val / total) * 100).toFixed(2) + '%');"
        response.write "    var ctx = document.getElementById('prescribedLab').getContext('2d');"
        response.write "    var pieChart = new Chart(ctx, {"
        response.write "      type: 'bar',"
        response.write "      data: {"
        response.write "        labels: labels,"
        response.write "        datasets: [{"
        response.write "          data: values,"
        response.write "          backgroundColor: ['#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff'],"
        response.write "        }]"
        response.write "      },"
        response.write "      options: {"
        response.write "        scales: {"
        response.write "          x: {"
        response.write "            ticks: {"
        response.write "              color: '#fff'"
        response.write "            }"
        response.write "          },"
        response.write "          y: {"
        response.write "title: {"
        response.write "            display: true,"
        response.write "            text: 'Number',"
        response.write "            color: '#fff'"
        response.write "          },"
        response.write "            ticks: {"
        response.write "              color: '#fff'"
        response.write "            }"
        response.write "          }"
        response.write "        },"
        response.write "        plugins: {"
        response.write "          legend: {"
        response.write "            display: false "
        response.write "          },"
        response.write "          tooltips: {"
        response.write "            callbacks: {"
        response.write "              label: (tooltipItem, data) => {"
        response.write "                var dataIndex = tooltipItem.index;"
        response.write "                return `${data.labels[dataIndex]}: ${data.datasets[0].data[dataIndex]} (${percentages[dataIndex]})`;"
        response.write "              }"
        response.write "            }"
        response.write "          }"
        response.write "        }"
        response.write "      }"
        response.write "    });"
        response.write "  </script>"
 End If
    
    rst.Close
    Set rst = Nothing

End Sub

Sub MostLabTestBySale()

 Dim sql, rst
    
    Set rst = CreateObject("ADODB.recordset")
    
    sql = "SELECT TOP 10 LabTestName, LabTestID, SUM(finalAmt) AS amount "
    sql = sql & " FROM ( "
    sql = sql & "     SELECT LabTest.LabTestName, investigation.LabTestID, finalAmt "
    sql = sql & "     FROM Investigation "
    sql = sql & "     JOIN LabTest ON LabTest.LabTestID = Investigation.LabTestID "

    If (periodStart <> "") And (periodEnd <> "") Then
      sql = sql & " WHERE Investigation.RequestDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If

    sql = sql & "     UNION ALL "
    sql = sql & "     SELECT LabTest.LabTestName, Investigation2.LabTestID, finalAmt "
    sql = sql & "     FROM Investigation2 "
    sql = sql & "     JOIN LabTest ON LabTest.LabTestID = Investigation2.LabTestID "
    sql = sql & " ) AS combined_data "
    sql = sql & " GROUP BY LabTestID, LabTestName "
    sql = sql & " ORDER BY amount DESC; "
            
    rst.open qryPro.FltQry(sql), conn, 3, 4
    
    Dim SuppArray, totalAmtArray
    SuppArray = ""
    totalAmtArray = ""
    
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            If SuppArray <> "" Then SuppArray = SuppArray & ","
            If totalAmtArray <> "" Then totalAmtArray = totalAmtArray & ","
            
            SuppArray = SuppArray & "'" & rst.fields("LabTestName") & "'"
            totalAmtArray = totalAmtArray & rst.fields("Amount")
            
            rst.MoveNext
        Loop
        
        response.write "  <script>"
        response.write "    var labels = [" & SuppArray & "];"
        response.write "    var values = [" & totalAmtArray & "];"
        response.write "    var total = values.reduce((acc, val) => acc + val, 0);"
        response.write "    var percentages = values.map(val => ((val / total) * 100).toFixed(2) + '%');"
        response.write "    var ctx = document.getElementById('MostLabTest').getContext('2d');"
        response.write "    var pieChart = new Chart(ctx, {"
        response.write "      type: 'bar',"
        response.write "      data: {"
        response.write "        labels: labels,"
        response.write "        datasets: [{"
        response.write "          data: values,"
        response.write "          backgroundColor: ['#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff'],"
        response.write "        }]"
        response.write "      },"
        response.write "      options: {"
        response.write "        scales: {"
        response.write "          x: {"
        response.write "            ticks: {"
        response.write "              color: '#fff'"
        response.write "            }"
        response.write "          },"
        response.write "          y: {"
        response.write "title: {"
        response.write "            display: true,"
        response.write "            text: 'Total Cost',"
        response.write "            color: '#fff'"
        response.write "          },"
        response.write "            ticks: {"
        response.write "              color: '#fff'"
        response.write "            }"
        response.write "          }"
        response.write "        },"
        response.write "        plugins: {"
        response.write "          legend: {"
        response.write "            display: false "
        response.write "          },"
        response.write "          tooltips: {"
        response.write "            callbacks: {"
        response.write "              label: (tooltipItem, data) => {"
        response.write "                var dataIndex = tooltipItem.index;"
        response.write "                return `${data.labels[dataIndex]}: ${data.datasets[0].data[dataIndex]} (${percentages[dataIndex]})`;"
        response.write "              }"
        response.write "            }"
        response.write "          }"
        response.write "        }"
        response.write "      }"
        response.write "    });"
        response.write "  </script>"
 End If
    
    rst.Close
    Set rst = Nothing

End Sub

Sub TotalPatByAdJS()

 Dim sql, rst
    
    Set rst = CreateObject("ADODB.recordset")
    
    sql = " SELECT WorkingYear.WorkingYearID, COUNT(PatientID) AS totPat FROM Admission "
    sql = sql & " LEFT JOIN WorkingYear ON Admission.WorkingYearID = WorkingYear.WorkingYearID "

    If (periodStart <> "") And (periodEnd <> "") Then
      sql = sql & " WHERE Admission.AdmissionDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If

    sql = sql & " GROUP BY WorkingYear.WorkingYearID, WorkingYear.WorkingYearName "
    
    rst.open qryPro.FltQry(sql), conn, 3, 4
    
    Dim yearssArray, totPatArray
    yearssArray = ""
    totPatArray = ""
    
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            If yearssArray <> "" Then yearssArray = yearssArray & ","
            If totPatArray <> "" Then totPatArray = totPatArray & ","
            
            yearssArray = yearssArray & "'" & rst.fields("WorkingYearID") & "'"
            totPatArray = totPatArray & rst.fields("totPat")
            
            rst.MoveNext
        Loop
        
        response.write "  <script>"
        response.write "    var labels = [" & yearssArray & "];"
        response.write "    var values = [" & totPatArray & "];"
        response.write "    var ctx = document.getElementById('Admission-chart').getContext('2d');"
        response.write "    var pieChart = new Chart(ctx, {"
        response.write "      type: 'bar',"
        response.write "      data: {"
        response.write "        labels: labels,"
        response.write "        datasets: [{"
        response.write "          data: values,"
        response.write "          backgroundColor: ['#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff'],"
        response.write "        }]"
        response.write "      },"
        response.write "      options: {"
        response.write "        scales: {"
        response.write "          x: {"
        response.write "            ticks: {"
        response.write "              color: '#fff'"
        response.write "            }"
        response.write "          },"
        response.write "          y: {"
        response.write "title: {"
        response.write "            display: true,"
        response.write "            text: 'Number of Patients',"
        response.write "            color: '#fff'"
        response.write "          },"
        response.write "            ticks: {"
        response.write "              color: '#fff'"
        response.write "            }"
        response.write "          }"
        response.write "        },"
        response.write "        plugins: {"
        response.write "          legend: {"
        response.write "            display: false "
        response.write "          },"
        response.write "          tooltips: {"
        response.write "          }"
        response.write "        }"
        response.write "      }"
        response.write "    });"
        response.write "  </script>"
    End If
    
    rst.Close
    Set rst = Nothing

End Sub

Sub TotalPatByAdJS2()

    Dim sql, rst
    
    Set rst = CreateObject("ADODB.recordset")
    
    sql = " SELECT WorkingYear.WorkingYearID, COUNT(PatientID) AS totPat FROM Admission "
    sql = sql & " LEFT JOIN WorkingYear ON Admission.WorkingYearID = WorkingYear.WorkingYearID "

    If (periodStart <> "") And (periodEnd <> "") Then
      sql = sql & " WHERE Admission.AdmissionDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If
    
    sql = sql & " GROUP BY WorkingYear.WorkingYearID, WorkingYear.WorkingYearName "
    
    rst.open qryPro.FltQry(sql), conn, 3, 4
    
    Dim yearssArray, totPatArray
    yearssArray = ""
    totPatArray = ""
    
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            If yearssArray <> "" Then yearssArray = yearssArray & ","
            If totPatArray <> "" Then totPatArray = totPatArray & ","
            
            yearssArray = yearssArray & "'" & rst.fields("WorkingYearID") & "'"
            totPatArray = totPatArray & rst.fields("totPat")
            
            rst.MoveNext
        Loop

        response.write "  <script>"
        response.write "    var labels = [" & yearssArray & "];"
        response.write "    var values = [" & totPatArray & "];"
        response.write "    var ctx = document.getElementById('AdScPat').getContext('2d');"
        response.write "    var pieChart = new Chart(ctx, {"
        response.write "      type: 'line',"
        response.write "      data: {"
        response.write "        labels: labels,"
        response.write "        datasets: [{"
        response.write "          data: values,"
        response.write "          backgroundColor: ['#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff', '#2962ff'],"
        response.write "        }]"
        response.write "      },"
        response.write "      options: {"
        response.write "        scales: {"
        response.write "          x: {"
        response.write "            ticks: {"
        response.write "              color: '#fff'"
        response.write "            }"
        response.write "          },"
        response.write "          y: {"
        response.write "title: {"
        response.write "            display: true,"
        response.write "            text: 'Number of Patients',"
        response.write "            color: '#fff'"
        response.write "          },"
        response.write "            ticks: {"
        response.write "              color: '#fff'"
        response.write "            }"
        response.write "          }"
        response.write "        },"
        response.write "        plugins: {"
        response.write "          legend: {"
        response.write "            display: false "
        response.write "          },"
        response.write "          tooltips: {"
        
        response.write "          }"
        response.write "        }"
        response.write "      }"
        response.write "    });"
        response.write "  </script>"
    End If
    
    rst.Close
    Set rst = Nothing

End Sub

Function GetTotalPat()
    Dim rstPat, sql, finalAmt
    
    Set rstPat = CreateObject("ADODB.recordset")
    
    sql = "SELECT COUNT(PatientID) as totPat FROM Patient"
    If (periodStart <> "") And (periodEnd <> "") Then
      sql = sql & " WHERE FirstVisitDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If

    rstPat.open qryPro.FltQry(sql), conn, 3, 4
    If Not rstPat.EOF Then
        finalNum = rstPat.fields("totPat")
    End If
    rstPat.Close
    
    GetTotalPat = finalNum
End Function

Function GetTotalPatAd()
    Dim rstPat, sql
    
    Set rstPat = CreateObject("ADODB.recordset")
    
    sql = "SELECT COUNT(PatientID) as totPatAd FROM Admission"
    If (periodStart <> "") And (periodEnd <> "") Then
      sql = sql & " WHERE AdmissionDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If

    rstPat.open qryPro.FltQry(sql), conn, 3, 4
    If Not rstPat.EOF Then
            finalNum = rstPat.fields("totPatAd")
    End If
    rstPat.Close
    
    GetTotalPatAd = finalNum
End Function

Function GetTotalVisit()
    Dim rstPat, sql
    
    Set rstPat = CreateObject("ADODB.recordset")
    
    sql = "SELECT COUNT(VisitationID) as totVit FROM Visitation"
    If (periodStart <> "") And (periodEnd <> "") Then
      sql = sql & " WHERE VisitDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If
    rstPat.open qryPro.FltQry(sql), conn, 3, 4
    If Not rstPat.EOF Then
            finalNum = rstPat.fields("totVit")
    End If
    rstPat.Close
    
    GetTotalVisit = finalNum
End Function

Function GetTotalSpec()
    Dim rstPat, sql
    
    Set rstPat = CreateObject("ADODB.recordset")
    
    sql = "SELECT COUNT(SpecialistID) as totSpec FROM Specialist"
    rstPat.open qryPro.FltQry(sql), conn, 3, 4
    If Not rstPat.EOF Then
            finalNum = rstPat.fields("totSpec")
    End If
    rstPat.Close
    
    GetTotalSpec = finalNum
End Function

Function GetTotalStaff()
    Dim rstPat, sql
    
    Set rstPat = CreateObject("ADODB.recordset")
    
    sql = "SELECT COUNT(StaffID) as totStaff FROM Staff"
    sql = sql & " WHERE StaffStatusID = 'S001'"

    rstPat.open qryPro.FltQry(sql), conn, 3, 4
    If Not rstPat.EOF Then
        finalNum = rstPat.fields("totStaff")
    End If
    rstPat.Close
    
    GetTotalStaff = finalNum
End Function

Function GetTotalCancelled()
    Dim rstPat, sql
    
    Set rstPat = CreateObject("ADODB.recordset")

    sql = " SELECT  AppointmentStatus.AppointmentStatusName, COUNT(appointment.patientID) AS totalpatients "
    sql = sql & " From Appointment "
    sql = sql & " left JOIN AppointmentStatus ON AppointmentStatus.AppointmentStatusID = Appointment.AppointmentStatusID "
    sql = sql & " WHERE Appointment.AppointmentStatusID='A004' "

    If (periodStart <> "") And (periodEnd <> "") Then
      sql = sql & " AND Appointment.AppointDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If

    sql = sql & " GROUP BY AppointmentStatus.AppointmentStatusName "
                
    rstPat.open qryPro.FltQry(sql), conn, 3, 4
    If Not rstPat.EOF Then
            finalNum = rstPat.fields("totalpatients")
    End If
    rstPat.Close
    
    GetTotalCancelled = finalNum
End Function

Function GetMortalityR()
    Dim rstPat, sql
    
    Set rstPat = CreateObject("ADODB.recordset")
    
    sql = "  SELECT"
    sql = sql & " COUNT(DISTINCT patientid) AS total_patients, "
    sql = sql & "  COUNT(CASE WHEN MedicalOutcomeID = 'm002' "
    sql = sql & " THEN patientid END) AS patients_with_mortality,"
    sql = sql & "  CAST(COUNT(CASE WHEN MedicalOutcomeID = 'm002'"
    sql = sql & " THEN patientid END) AS FLOAT) / COUNT(DISTINCT patientid) AS mortality_rate "
    sql = sql & " FROM visitation "

    If (periodStart <> "") And (periodEnd <> "") Then
        sql = sql & " WHERE VisitDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If

    rstPat.open qryPro.FltQry(sql), conn, 3, 4
    If Not rstPat.EOF Then
            finalNum = rstPat.fields("mortality_rate")
    End If
    rstPat.Close
    
    GetMortalityR = finalNum
End Function

Function GetBedOcc()
    Dim rstPat, sql
    
    Set rstPat = CreateObject("ADODB.recordset")
    
    sql = " SELECT "
    sql = sql & " occupied_beds AS [total occupied beds], "
    sql = sql & " total_beds,"
    sql = sql & " (CAST(occupied_beds AS FLOAT) / total_beds) * 100 AS Bed_Occupancy_Rate "
    sql = sql & " FROM ( "
    sql = sql & " SELECT COUNT(bedid) AS total_beds FROM Bed "
    sql = sql & " ) AS beds_info "
    sql = sql & " CROSS JOIN ( "
    sql = sql & " SELECT COUNT(bedid) AS occupied_beds FROM Admission "
    sql = sql & " Where AdmissionDate <= getDate() And (DischargeDate > getDate() Or DischargeDate Is Null)"
    sql = sql & " ) AS occupied_beds_info "
           
    rstPat.open qryPro.FltQry(sql), conn, 3, 4
    If Not rstPat.EOF Then
            finalNum = rstPat.fields("Bed_Occupancy_Rate")
    End If
    rstPat.Close
    
    GetBedOcc = finalNum
End Function

Function GetDrugSaleAmt()
    Dim rstAmt1, rstAmt2, sql, finalAmt
    
    Set rstAmt1 = CreateObject("ADODB.recordset")
    Set rstAmt2 = CreateObject("ADODB.recordset")
    
    sql = "SELECT SUM(FinalAmt) AS totAmt FROM DrugSaleItems"
    If (periodStart <> "") And (periodEnd <> "") Then
        sql = sql & " WHERE DispenseDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If
    rstAmt1.open qryPro.FltQry(sql), conn, 3, 4
    If Not rstAmt1.EOF Then
        If Not IsNull(rstAmt1.fields("totAmt")) Then
            finalAmt = rstAmt1.fields("totAmt")
        End If
    End If
    rstAmt1.Close
    
    sql = "SELECT SUM(FinalAmt) AS totAmt1 FROM DrugSaleItems2"
    If (periodStart <> "") And (periodEnd <> "") Then
        sql = sql & " WHERE DispenseDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If
    rstAmt2.open qryPro.FltQry(sql), conn, 3, 4
    If Not rstAmt2.EOF Then
        If Not IsNull(rstAmt2.fields("totAmt1")) Then
            finalAmt = finalAmt + rstAmt2.fields("totAmt1")
        End If
    End If
    rstAmt2.Close
    
    GetDrugSaleAmt = finalAmt
End Function

Function GetLabRequestAmt()
    Dim rstAmt1, rstAmt2, sql, finalAmt
    
    Set rstAmt1 = CreateObject("ADODB.recordset")
    Set rstAmt2 = CreateObject("ADODB.recordset")
    
    sql = "SELECT SUM(FinalAmt) AS totAmt FROM Investigation"
    If (periodStart <> "") And (periodEnd <> "") Then
        sql = sql & " WHERE RequestDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If
    rstAmt1.open qryPro.FltQry(sql), conn, 3, 4
    If Not rstAmt1.EOF Then
        If Not IsNull(rstAmt1.fields("totAmt")) Then
            finalAmt = rstAmt1.fields("totAmt")
        End If
    End If
    rstAmt1.Close
    
    sql = "SELECT SUM(FinalAmt) AS totAmt1 FROM Investigation2"
    If (periodStart <> "") And (periodEnd <> "") Then
        sql = sql & " WHERE RequestDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If
    rstAmt2.open qryPro.FltQry(sql), conn, 3, 4
    If Not rstAmt2.EOF Then
        If Not IsNull(rstAmt2.fields("totAmt1")) Then
            finalAmt = finalAmt + rstAmt2.fields("totAmt1")
        End If
    End If
    rstAmt2.Close
    
    GetLabRequestAmt = finalAmt
End Function

Function GetTreatAmt()
    Dim rstAmt1, sql, finalAmt
    
    Set rstAmt1 = CreateObject("ADODB.recordset")
    
    sql = "SELECT SUM(FinalAmt) as totAmt FROM TreatCharges"
    If (periodStart <> "") And (periodEnd <> "") Then
        sql = sql & " WHERE ConsultReviewDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If
    rstAmt1.open qryPro.FltQry(sql), conn, 3, 4
    If Not rstAmt1.EOF Then
        If Not IsNull(rstAmt1.fields("totAmt")) Then
            finalAmt = rstAmt1.fields("totAmt")
        End If
    End If
    rstAmt1.Close
    
    GetTreatAmt = finalAmt
End Function

Function GetVstCost()
    Dim rstAmt1, sql, finalAmt
    
    Set rstAmt1 = CreateObject("ADODB.recordset")
    
    sql = "SELECT SUM(VisitCost) AS totAmt FROM Visitation"
    If (periodStart <> "") And (periodEnd <> "") Then
        sql = sql & " WHERE VisitDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If
    rstAmt1.open qryPro.FltQry(sql), conn, 3, 4
    If Not rstAmt1.EOF Then
        If Not IsNull(rstAmt1.fields("totAmt")) Then
            finalAmt = rstAmt1.fields("totAmt")
        End If
    End If
    rstAmt1.Close
    
    GetVstCost = finalAmt
End Function

Function GetPaidAmt()
    Dim rstAmt1, sql, finalAmt
    
    Set rstAmt1 = CreateObject("ADODB.recordset")
    
    sql = "SELECT SUM(PaidAmount) AS totAmt FROM PatientReceipt2"
    If (periodStart <> "") And (periodEnd <> "") Then
        sql = sql & " WHERE ReceiptDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If
    rstAmt1.open qryPro.FltQry(sql), conn, 3, 4
    If Not rstAmt1.EOF Then
        If Not IsNull(rstAmt1.fields("totAmt")) Then
            finalAmt = rstAmt1.fields("totAmt")
        End If
    End If
    rstAmt1.Close
    
    GetPaidAmt = finalAmt
End Function

Function GetPharmWaitTime()
    Dim rstPat, sql, finalAmt
    
    Set rstPat = CreateObject("ADODB.recordset")
   
    sql = " SELECT  AVG(DATEDIFF(minute, Prescription.PrescriptionDate, DispenseDate)) AS avg_wait_time FROM DrugSaleItems2"
    sql = sql & " LEFT JOIN Prescription ON DrugSaleItems2.PrescriptionID = Prescription.PrescriptionID"
    If (periodStart <> "") And (periodEnd <> "") Then
        sql = sql & " WHERE DrugSaleItems2.DispenseDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
    End If
    
    rstPat.open qryPro.FltQry(sql), conn, 3, 4
    If Not rstPat.EOF Then
            finalNum = rstPat.fields("avg_wait_time")
    End If
    rstPat.Close
    
    GetPharmWaitTime = finalNum
End Function

Function GetOpdAvgWaitTime()
    Dim rstPat, sql, finalAmt
    
    Set rstPat = CreateObject("ADODB.recordset")
    
        sql = " SELECT"
        sql = sql & " AVG(DateDiff(Minute, visitDate, EMRRequestItems.emrDate)) As avg_wait_time"
        sql = sql & " From"
        sql = sql & " Visitation"
        sql = sql & " Left Join"
        sql = sql & " EMRRequestItems ON Visitation.VisitationID = EMRRequestItems.VisitationID"
        sql = sql & " Where"
        sql = sql & " EMRRequestItems.EMRDataID = 'EMR050'"
        If (periodStart <> "") And (periodEnd <> "") Then
            sql = sql & " AND Visitation.VisitDate BETWEEN '" & periodStart & "' AND '" & periodEnd & "'"
        End If
        
    rstPat.open qryPro.FltQry(sql), conn, 3, 4
    If Not rstPat.EOF Then
            finalNum = rstPat.fields("avg_wait_time")
    End If
    rstPat.Close
    
    GetOpdAvgWaitTime = finalNum
End Function

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
