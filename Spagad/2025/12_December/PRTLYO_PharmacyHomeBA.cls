VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PRTLYO_PharmacyHomeBA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

response.Clear



response.write "<!DOCTYPE html>"
response.write "<html lang=""en"">"
response.write "<head>"
response.write "    <title>Pharmacy Analytics Dashboard </title>"
response.write "    <meta charset=""utf-8"">"
response.write "    <meta http-equiv=""X-UA-Compatible"" content=""IE=Edge"">"
response.write "    <meta name=""viewport"" content=""width=device-width, initial-scale=1, shrink-to-fit=no"">"
' Response.Write "    <meta name=""description"" content=""Responsive Bootstrap 5 admin dashboard template & web App ui kit."">"
' Response.Write "    <meta name=""keyword"" content=""bootstrap admin template"">"
response.write ""
response.write "    <!--[ Favicon]-->"
response.write "    <link rel=""icon"" type=""image/x-icon"" href=""images/banner1.bmp"">"
response.write "    <link rel=""icon"" type=""image/png"" sizes=""16x16"" href=""images/banner1.bmp"">"
response.write "    <link rel=""icon"" type=""image/png"" sizes=""32x32"" href=""images/banner1.bmp"">"
response.write "    <link rel=""apple-touch-icon"" sizes=""180x180"" href=""images/banner1.bmp"">"
' response.write "    <link rel=""icon"" type=""image/x-icon"" href=""BusinessAnalytics/assets/images/favicon.ico"">"
' response.write "    <link rel=""icon"" type=""image/png"" sizes=""16x16"" href=""BusinessAnalytics/assets/images/favicon-16x16.png"">"
' response.write "    <link rel=""icon"" type=""image/png"" sizes=""32x32"" href=""BusinessAnalytics/assets/images/favicon-32x32.png"">"
' response.write "    <link rel=""apple-touch-icon"" sizes=""180x180"" href=""BusinessAnalytics/assets/images/apple-touch-icon.png"">"
response.write ""
response.write "    <link rel=""stylesheet"" href=""BusinessAnalytics/assets/vendor/dataTables.min.css"">"
response.write ""
response.write "    <!--[ Template main css file ]-->"
response.write "    <link rel=""stylesheet"" href=""BusinessAnalytics/assets/css/style.min.css"">"
response.write "<style>"
response.write "    body {"
response.write "        overflow: hidden;" ' Hide the scrollbar for the main page
response.write "    }"
response.write "</style>"
response.write "</head>"

response.write "<body data-theme=""theme-PurpleHeart"" class=""svgstroke-a bg-gradient"">"


response.write "<style>"
response.write "    #spinner {"
response.write "        position: fixed;"
response.write "        top: 0;"
response.write "        left: 0;"
response.write "        width: 100%;"
response.write "        height: 100%;"
' Response.Write "        background: rgba(255, 255, 255, 0.7);"
response.write "        display: none;" ' Initially hidden
response.write "        align-items: center;"
response.write "        justify-content: center;"
response.write "        backdrop-filter: blur(4px);"
response.write "        z-index: 9999;"
response.write "    }"
response.write "    .spinner-border {"
response.write "        width: 3rem;"
response.write "        height: 3rem;"
response.write "    }"
response.write "    .spinner-grow {"
response.write "        width: 3rem;"
response.write "        height: 3rem;"
response.write "        margin-left: 1rem;"
response.write "    }"
response.write "</style>"

response.write "<div id='spinner'>"
response.write "    <div class=""spinner-border"" role=""status"">"
response.write "        <span class=""visually-hidden"">Loading...</span>"
response.write "    </div>"
response.write "    <div class=""spinner-grow"" role=""status"">"
response.write "        <span class=""visually-hidden"">Loading...</span>"
response.write "    </div>"
response.write "</div>"

Dim users




response.write "    <main class=""container-fluid px-0"">"
'    <!-- start: page menu link -->
response.write "        <aside class=""ps-xl-5 ps-lg-4 ps-3 pe-2 py-3 sidebar sticky-top"" data-bs-theme=""none"">"
response.write "            <nav class=""navbar navbar-expand-xl py-0"">"
response.write "                <div class=""offcanvas offcanvas-start"" data-bs-scroll=""true"" tabindex=""-1"" id=""offcanvas_Navbar"">"
response.write "                    <div class=""offcanvas-header"">"
response.write "                        <span class=""fw-bold fs-5 text-gradient"">" & GetComboName("JobSchedule", jSchd) & "</span>"

response.write "                        <button type=""button"" class=""btn-close btn-close-white"" data-bs-dismiss=""offcanvas"" aria-label=""Close""></button>"
response.write "                    </div>"
response.write "                    <div class=""offcanvas-body flex-column custom_scroll ps-4 ps-xl-0"">"
response.write "                        <div class=""d-flex align-items-start mb-4"">"
' response.write "                            <img class=""avatar lg rounded-circle border border-3"" src=""BusinessAnalytics/assets/images/profile_av.png"" alt=""avatar"">"
response.write "                            <img class=""avatar lg rounded-circle border border-3"" src=""images/TESTFOLDER/logo1.gif"" alt=""avatar"">"
' response.write "                            <img class=""avatar lg rounded-circle border border-3"" src=""images/TESTFOLDER/banner1.bmp"" alt=""avatar"">"

response.write "                            <div class=""ms-3 mt-1"">"
' Response.Write "                                <h4 class=""mb-0 text-gradient title-font"">BA01</h4>"
response.write "                                <h4 class=""mb-0 text-gradient title-font"">" & UCase(jSchd) & "</h4>"
        

response.write "                                <span class=""text-muted"">" & GetComboName("JobSchedule", jSchd) & "</span>"
response.write "                            </div>"
response.write "                        </div>"


' ====================================================================================================================================
' ====================================================================================================================================
' ====================================================================================================================================

If UCase(uName) = "BA01" Or UCase(jSchd) = "M0603" Then
    response.write "                        <!-- start: Menu tab -->"
    response.write "                        <ul class=""nav nav-tabs tab-dynamic px-0 mb-4"" role=""tablist"" style=""--dynamic-color: var(--accent-color)"">"
    response.write "                            <li class=""nav-item"" role=""presentation""><a class=""ps-2 pe-3 nav-link active"" data-bs-toggle=""tab"" href=""#tab_analysis"" role=""tab"" aria-selected=""false"" tabindex=""-1"" id =""main_tab_analysis"">ANALYSIS</a></li>"
'     response.write "                            <li class=""nav-item"" role=""presentation""><a class=""ps-2 pe-3 nav-link"" data-bs-toggle=""tab"" href=""#tab_predictions"" role=""tab"" aria-selected=""true"">PREDICTIONS</a></li>"
    ' response.write "                            <li class=""nav-item"" role=""presentation""><a class=""ps-2 pe-3 nav-link"" data-bs-toggle=""tab"" href=""#tab_reports"" role=""tab"" aria-selected=""true"">REPORTS</a></li>"
'    response.write "                            <li class=""nav-item"" role=""presentation""><a class=""ps-2 pe-3 nav-link"" data-bs-toggle=""tab"" href=""#tab_predictions"" role=""tab"" aria-selected=""true"">PREDICTIONS not</a></li>"
    response.write "                        </ul>"
    response.write "                        <div class=""tab-content"">"
    response.write "                            <!-- start: Analytics tab-->"
    response.write "                            <div class=""tab-pane fade active show"" id=""tab_analysis"" role=""tabpanel"">"
    ' Response.Write "                                <h6 class=""fl-title title-font ps-2 small text-uppercase text-muted"" style=""--dynamic-color: var(--theme-color1)"">Universal</h6>"
    response.write "                                <ul class=""list-unstyled mb-4 menu-list"">"
    response.write "                                    <li>"
    response.write "                                        <a class=""active"" onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=DashboardContentBI&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)"" href=""#Dashboard"" aria-label=""Analytics Dashboard"">"
    response.write "                                            <svg class=""svg-stroke"" xmlns=""http://www.w3.org/2000/svg"" width=""24"" viewBox=""0 0 24 24"" stroke=""currentColor"" fill=""none"" stroke-linecap=""round"" stroke-linejoin=""round"">"
    response.write "                                                <path stroke=""none"" d=""M0 0h24v24H0z"" fill=""none""></path>"
    response.write "                                                <path d=""M5 12l-2 0l9 -9l9 9l-2 0""></path>"
    response.write "                                                <path d=""M5 12v7a2 2 0 0 0 2 2h10a2 2 0 0 0 2 -2v-7""></path>"
    response.write "                                                <path d=""M10 12h4v4h-4z""></path>"
    response.write "                                            </svg>"
    response.write "                                            <span class=""mx-3"">Dashboard</span>"
    ' response.write "                                            <span><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=DashboardContentBI&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""Dashboard"">Dashboard</a></span>"
    ' response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=DashboardContentBI&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""Dashboard"">Dashboard</a></li>"
    response.write "                                        </a>"
    response.write "                                    </li>"
    

    response.write "                                    <li>"
    response.write "                                        <a href=""#Drugs"" data-bs-toggle=""collapse"" aria-expanded=""false"" class=""dropdown-toggle"" aria-label=""Users"">"
    response.write "                                            <p> &nbsp;</p>"
    response.write "                                            <svg xmlns=""http://www.w3.org/2000/svg"" width=""16"" height=""16"" fill=""currentColor"" class=""bi bi-capsule-pill"" viewBox=""0 0 16 16"">"
    response.write "                                                <path d=""M11.02 5.364a3 3 0 0 0-4.242-4.243L1.121 6.778a3 3 0 1 0 4.243 4.243l5.657-5.657Zm-6.413-.657 2.878-2.879a2 2 0 1 1 2.829 2.829L7.435 7.536zM12 8a4 4 0 1 1 0 8 4 4 0 0 1 0-8m-.5 1.042a3 3 0 0 0 0 5.917zm1 5.917a3 3 0 0 0 0-5.917z""/>"
    response.write "                                              </svg>"
    response.write "                                            <span class=""mx-3"">Pharmacy & Drugs</span>"
    response.write "                                        </a>"
    response.write "                                        <ul class=""collapse list-unstyled"" id=""Drugs"">"
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=PrescriptionDashboardBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""prescription dashboard"">Prescription Dashboard</a></li>"
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=prescriptionanalysisba&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""prescription analysis"">Prescription Analysis</a></li>"
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=PrescriptionFulfilmentAnalysisBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""Prescription Fulfilment Analysis"">Prescription Fulfilment Analysis</a></li>"
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=PrescriptionByDrugTypeBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""Prescription By Drug Type"">Prescription By Drug Type</a></li>"
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=DoctorPrescriptionHistoryBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""Doctor Prescription History"">Doctor Prescription History</a></li>"
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=PrescriptionByMedicalServicesBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""Prescription By Medical Services"">Prescription By Medical Services</a></li>"
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=PrescriptionPerPatientBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""Prescription Behavior"">Prescription Behaviour</a></li>"
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=ADRReportingBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""ADRReportingBA"">Adverse Reporting</a></li>"
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=ReconciliationAtAdmissionBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""ReconciliationAtAdmissionBA"">Admission Reconciliation</a></li>"
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=CounseledOnDischargeBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""CounseledOnDischargeBA"">Counseled On Discharge</a></li>"
    
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=MostDespensedDrugsBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""MostDespensedDrugsBA"">Most Dispensed Drugs</a></li>"
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=AntimicrobialStewardshipBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""AntimicrobialStewardshipBA"">Antimicrobial Stewardship</a></li>"
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=PrescriptionErrorRatesBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""PrescriptionErrorRatesBA"">Prescription Error Rates</a></li>"
        response.write "                                        <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=DespenseRevenueOnDrugsBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""DespenseRevenueOnDrugsBA"">Revenue on Dispenses </a></li>"
        response.write "                                       <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=RationalUseOfDrugsBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""RationalUseOfDrugsBA"">Rational Use Of Medicines </a></li>"
        response.write "                                        <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=TopLeastDispensedDrugsBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""TopLeastDispensedDrugsBA"">Top Least Dispensed Drugs</a></li>"
     response.write "                                        <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=AntibioticSurvelanceBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""AntibioticSurvelanceBA"">Antibiotic Survelance</a></li>"
     response.write "                                        <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=RestrictedMedicationsTrendsBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""RestrictedMedicationsTrendsBA"">Restricted Medications</a></li>"
    
    
    response.write "                                    <li>"
    response.write "                                        <a href=""#DrugsReport"" data-bs-toggle=""collapse"" aria-expanded=""false"" class=""dropdown-toggle"" aria-label=""Users"">"
    ' response.write "                                            <p> &nbsp;</p>"
    ' response.write "                                            <svg xmlns=""http://www.w3.org/2000/svg"" width=""16"" height=""16"" fill=""currentColor"" class=""bi bi-capsule-pill"" viewBox=""0 0 16 16"">"
    ' response.write "                                                <path d=""M11.02 5.364a3 3 0 0 0-4.242-4.243L1.121 6.778a3 3 0 1 0 4.243 4.243l5.657-5.657Zm-6.413-.657 2.878-2.879a2 2 0 1 1 2.829 2.829L7.435 7.536zM12 8a4 4 0 1 1 0 8 4 4 0 0 1 0-8m-.5 1.042a3 3 0 0 0 0 5.917zm1 5.917a3 3 0 0 0 0-5.917z""/>"
    ' response.write "                                              </svg>"
    response.write "                                            <span class=""mx-3"">Reports</span>"
    response.write "                                        </a>"
    response.write "                                        <ul class=""collapse list-unstyled"" id=""DrugsReport"">"
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=AntibioticsStewardshipAWaReRpt2BA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""Antibiotics Stewardship AWaRe"">Antibiotics Stewardship AWaRe</a></li>"
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=HighVolumeLowValueDrugSalesRpt&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""High/Low Volume DrugSales"">High/Low Volume DrugSales</a></li>"
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=AntibioticDispensationComplainceRpt&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""Antibiotic DispensationC omplaince"">Antibiotic Dispensation Complaince</a></li>"
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=PolypharmacyWithDiagnosisAndDispenseRpt&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""PolypharmacyWithDiagnosisAndDispenseRpt"">Polypharmacy with Diagnosis and Dispense</a></li>"

    response.write "                                        </ul>"
    response.write "                                    </li>"



    ' response.write "                                            <li><a href=""test.html"" aria-label=""Requests"">Requests</a></li>"
    ' response.write "                                            <li><a href=""test.html"" aria-label=""Expired Drugs"">Expired Drugs</a></li>"
    ' response.write "                                            <li><a href=""test.html"" aria-label=""By Prescription"">Prescription</a></li>"
    ' response.write "                                            <li><a href=""test.html"" aria-label=""Dispenses"">Dispenses</a></li>"
    response.write "                                        </ul>"
    response.write "                                    </li>"
    
    



    response.write "                            </div>"


    
    ' -----------------------------------------
    '


    response.write "                            <!-- start: HR tab -->"
    response.write "                            <div class=""tab-pane fade"" id=""tab_predictions"" role=""tabpanel"">"
    ' response.write "                                <h6 class=""fl-title title-font ps-2 small text-uppercase text-muted"" style=""--dynamic-color: var(--theme-color1)"">Usual</h6>"
    response.write "                                <ul class=""list-unstyled mb-4 menu-list"">"
    response.write "                                    <li>"
    response.write "                                        <a  href=""#ForcastAdmissions"" data-bs-toggle=""collapse"" aria-expanded=""false"" class=""dropdown-toggle"" aria-label=""Users"">"
    response.write "                                            <p> &nbsp;</p>"
    response.write "                                            <svg xmlns=""http://www.w3.org/2000/svg"" width=""16"" height=""16"" fill=""currentColor"" class=""bi bi-hospital"" viewBox=""0 0 16 16"">"
    response.write "                                                <path d=""M8.5 5.034v1.1l.953-.55.5.867L9 7l.953.55-.5.866-.953-.55v1.1h-1v-1.1l-.953.55-.5-.866L7 7l-.953-.55.5-.866.953.55v-1.1zM13.25 9a.25.25 0 0 0-.25.25v.5c0 .138.112.25.25.25h.5a.25.25 0 0 0 .25-.25v-.5a.25.25 0 0 0-.25-.25zM13 11.25a.25.25 0 0 1 .25-.25h.5a.25.25 0 0 1 .25.25v.5a.25.25 0 0 1-.25.25h-.5a.25.25 0 0 1-.25-.25zm.25 1.75a.25.25 0 0 0-.25.25v.5c0 .138.112.25.25.25h.5a.25.25 0 0 0 .25-.25v-.5a.25.25 0 0 0-.25-.25zm-11-4a.25.25 0 0 0-.25.25v.5c0 .138.112.25.25.25h.5A.25.25 0 0 0 3 9.75v-.5A.25.25 0 0 0 2.75 9zm0 2a.25.25 0 0 0-.25.25v.5c0 .138.112.25.25.25h.5a.25.25 0 0 0 .25-.25v-.5a.25.25 0 0 0-.25-.25zM2 13.25a.25.25 0 0 1 .25-.25h.5a.25.25 0 0 1 .25.25v.5a.25.25 0 0 1-.25.25h-.5a.25.25 0 0 1-.25-.25z""/>"
    response.write "                                                <path d=""M5 1a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v1a1 1 0 0 1 1 1v4h3a1 1 0 0 1 1 1v7a1 1 0 0 1-1 1H1a1 1 0 0 1-1-1V8a1 1 0 0 1 1-1h3V3a1 1 0 0 1 1-1zm2 14h2v-3H7zm3 0h1V3H5v12h1v-3a1 1 0 0 1 1-1h2a1 1 0 0 1 1 1zm0-14H6v1h4zm2 7v7h3V8zm-8 7V8H1v7z""/>"
    response.write "                                              </svg>"
    response.write "                                            <span class=""mx-3"">Visual Analytics</span>"
    response.write "                                        </a>"
    response.write "                                        <ul class=""collapse list-unstyled"" id=""ForcastAdmissions"">"
    ' Response.Write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=AdmissionTimeSeries&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1')"">Test Time Series Analysis </a></li>"
    ' response.write "                                            <li><a href='#' onclick=""openPage('http://192.168.5.170:8501')"">ML Tools</a></li>"
    
    response.write "                                            <li><a href='#' onclick=""openPage('http://192.168.5.13:8501')"">ML Tools</a></li>"
    
    ' response.write "                                            <li><a href=""test.html"" aria-label=""Admissions"">By Age Range</a></li>"
    ' response.write "                                            <li><a href=""test.html"" aria-label=""Time Series Analysis"">Time Series Analysis</a></li>"
    ' response.write "                                            <li><a href=""test.html"" aria-label=""Length Of Stay (LOS)"">Length Of Stay (LOS)</a></li>"
    response.write "                                        </ul>"
    response.write "                                    </li>"
    response.write "                                </ul>"
    response.write "                            </div>"

    ' PHARMACY section
ElseIf UCase(uName) = "PHABA01" Or UCase(uName) = "PHABA012" Or jSchd = "M0603" Then
    response.write "                        <!-- start: Menu tab -->"
    response.write "                        <ul class=""nav nav-tabs tab-dynamic px-0 mb-4"" role=""tablist"" style=""--dynamic-color: var(--accent-color)"">"
    response.write "                            <li class=""nav-item"" role=""presentation""><a class=""ps-2 pe-3 nav-link active"" data-bs-toggle=""tab"" href=""#tab_analysis"" role=""tab"" aria-selected=""false"" tabindex=""-1"" id =""main_tab_analysis"">ANALYSIS</a></li>"
    ' response.write "                            <li class=""nav-item"" role=""presentation""><a class=""ps-2 pe-3 nav-link"" data-bs-toggle=""tab"" href=""#tab_predictions"" role=""tab"" aria-selected=""true"">PREDICTIONS not</a></li>"
    ' response.write "                            <li class=""nav-item"" role=""presentation""><a class=""ps-2 pe-3 nav-link"" data-bs-toggle=""tab"" href=""#tab_reports"" role=""tab"" aria-selected=""true"">REPORTS</a></li>"
    
'    response.write "                            <li class=""nav-item"" role=""presentation""><a class=""ps-2 pe-3 nav-link"" data-bs-toggle=""tab"" href=""#tab_predictions"" role=""tab"" aria-selected=""true"">PREDICTIONS</a></li>"
    response.write "                        </ul>"
    response.write "                        <div class=""tab-content"">"
    response.write "                            <!-- start: Analytics tab-->"
    response.write "                            <div class=""tab-pane fade active show"" id=""tab_analysis"" role=""tabpanel"">"
    ' Response.Write "                                <h6 class=""fl-title title-font ps-2 small text-uppercase text-muted"" style=""--dynamic-color: var(--theme-color1)"">Universal</h6>"
    response.write "                                <ul class=""list-unstyled mb-4 menu-list"">"
    response.write "                                    <li>"
    response.write "                                        <a class=""active"" onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=PrescriptionDashboardBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)"" href=""#Dashboard"" aria-label=""Analytics Dashboard"">"
    response.write "                                            <svg class=""svg-stroke"" xmlns=""http://www.w3.org/2000/svg"" width=""24"" viewBox=""0 0 24 24"" stroke=""currentColor"" fill=""none"" stroke-linecap=""round"" stroke-linejoin=""round"">"
    response.write "                                                <path stroke=""none"" d=""M0 0h24v24H0z"" fill=""none""></path>"
    response.write "                                                <path d=""M5 12l-2 0l9 -9l9 9l-2 0""></path>"
    response.write "                                                <path d=""M5 12v7a2 2 0 0 0 2 2h10a2 2 0 0 0 2 -2v-7""></path>"
    response.write "                                                <path d=""M10 12h4v4h-4z""></path>"
    response.write "                                            </svg>"
    response.write "                                            <span class=""mx-3"">Dashboard</span>"
    ' response.write "                                            <span><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=PrescriptionDashboardBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""Dashboard"">Dashboard</a></span>"
    ' response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=PrescriptionDashboardBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""Dashboard"">Dashboard</a></li>"
    response.write "                                        </a>"
    response.write "                                    </li>"
    
    response.write "                                    <li>"
    response.write "                                        <a href=""#Drugs"" data-bs-toggle=""collapse"" aria-expanded=""false"" class=""dropdown-toggle"" aria-label=""Users"">"
    response.write "                                            <p> &nbsp;</p>"
    response.write "                                            <svg xmlns=""http://www.w3.org/2000/svg"" width=""16"" height=""16"" fill=""currentColor"" class=""bi bi-capsule-pill"" viewBox=""0 0 16 16"">"
    response.write "                                                <path d=""M11.02 5.364a3 3 0 0 0-4.242-4.243L1.121 6.778a3 3 0 1 0 4.243 4.243l5.657-5.657Zm-6.413-.657 2.878-2.879a2 2 0 1 1 2.829 2.829L7.435 7.536zM12 8a4 4 0 1 1 0 8 4 4 0 0 1 0-8m-.5 1.042a3 3 0 0 0 0 5.917zm1 5.917a3 3 0 0 0 0-5.917z""/>"
    response.write "                                              </svg>"
    response.write "                                            <span class=""mx-3"">Pharmacy & Drugs</span>"
    response.write "                                        </a>"
    response.write "                                        <ul class=""collapse list-unstyled"" id=""Drugs"">"
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=PrescriptionDashboardBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""prescription dashboard"">Prescription Dashboard</a></li>"
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=prescriptionanalysisba&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""prescription analysis"">Prescription Analysis</a></li>"
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=PrescriptionFulfilmentAnalysisBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""Prescription Fulfilment Analysis"">Prescription Fulfilment Analysis</a></li>"
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=PrescriptionByDrugTypeBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""Prescription By Drug Type"">Prescription By Drug Type</a></li>"
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=DoctorPrescriptionHistoryBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""Doctor Prescription History"">Doctor Prescription History</a></li>"
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=PrescriptionByMedicalServicesBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""Prescription By Medical Services"">Prescription By Medical Services</a></li>"
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=PrescriptionPerPatientBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""Prescription Behavior"">Prescription Behaviour</a></li>"
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=ADRReportingBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""ADRReportingBA"">Adverse Reporting</a></li>"
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=ReconciliationAtAdmissionBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""ReconciliationAtAdmissionBA"">Admission Reconciliation</a></li>"
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=CounseledOnDischargeBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""CounseledOnDischargeBA"">Counseled On Discharge</a></li>"
    
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=MostDespensedDrugsBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""MostDespensedDrugsBA"">Most Dispensed Drugs</a></li>"
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=AntimicrobialStewardshipBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""AntimicrobialStewardshipBA"">Antimicrobial Stewardship</a></li>"
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=PrescriptionErrorRatesBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""PrescriptionErrorRatesBA"">Prescription Error Rates</a></li>"
        response.write "                                        <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=DespenseRevenueOnDrugsBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""DespenseRevenueOnDrugsBA"">Revenue on Dispenses </a></li>"
        
        response.write "                                        <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=RationalUseOfDrugsBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""RationalUseOfDrugsBA"">Rational Use Of Medicines </a></li>"
        response.write "                                        <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=TopLeastDispensedDrugsBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""TopLeastDispensedDrugsBA"">Top Least Dispensed Drugs</a></li>"

    response.write "                                    <li>"
    response.write "                                        <a href=""#DrugsReport"" data-bs-toggle=""collapse"" aria-expanded=""false"" class=""dropdown-toggle"" aria-label=""Users"">"
    ' response.write "                                            <p> &nbsp;</p>"
    ' response.write "                                            <svg xmlns=""http://www.w3.org/2000/svg"" width=""16"" height=""16"" fill=""currentColor"" class=""bi bi-capsule-pill"" viewBox=""0 0 16 16"">"
    ' response.write "                                                <path d=""M11.02 5.364a3 3 0 0 0-4.242-4.243L1.121 6.778a3 3 0 1 0 4.243 4.243l5.657-5.657Zm-6.413-.657 2.878-2.879a2 2 0 1 1 2.829 2.829L7.435 7.536zM12 8a4 4 0 1 1 0 8 4 4 0 0 1 0-8m-.5 1.042a3 3 0 0 0 0 5.917zm1 5.917a3 3 0 0 0 0-5.917z""/>"
    ' response.write "                                              </svg>"
    response.write "                                            <span class=""mx-3"">Reports</span>"
    response.write "                                        </a>"
    response.write "                                        <ul class=""collapse list-unstyled"" id=""DrugsReport"">"
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=AntibioticsStewardshipAWaReRpt2BA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""Antibiotics Stewardship AWaRe"">Antibiotics Stewardship AWaRe</a></li>"
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=HighVolumeLowValueDrugSalesRpt&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""High/Low Volume DrugSales"">High/Low Volume DrugSales</a></li>"
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=AntibioticDispensationComplainceRpt&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""Antibiotic DispensationC omplaince"">Antibiotic Dispensation Complaince</a></li>"
    response.write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=PolypharmacyWithDiagnosisAndDispenseRpt&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1', this)""aria-label=""PolypharmacyWithDiagnosisAndDispenseRpt"">Polypharmacy with Diagnosis and Dispense</a></li>"

    response.write "                                        </ul>"
    response.write "                                    </li>"



    ' response.write "                                            <li><a href=""test.html"" aria-label=""Requests"">Requests</a></li>"
    ' response.write "                                            <li><a href=""test.html"" aria-label=""Expired Drugs"">Expired Drugs</a></li>"
    ' response.write "                                            <li><a href=""test.html"" aria-label=""By Prescription"">Prescription</a></li>"
    ' response.write "                                            <li><a href=""test.html"" aria-label=""Dispenses"">Dispenses</a></li>"
    response.write "                                        </ul>"
    response.write "                                    </li>"

    ' response.write "                                            <li><a href=""test.html"" aria-label=""Requests"">Requests</a></li>"
    ' response.write "                                            <li><a href=""test.html"" aria-label=""Expired Drugs"">Expired Drugs</a></li>"
    ' response.write "                                            <li><a href=""test.html"" aria-label=""By Prescription"">Prescription</a></li>"
    ' response.write "                                            <li><a href=""test.html"" aria-label=""Dispenses"">Dispenses</a></li>"
    response.write "                                        </ul>"
    response.write "                                    </li>"
    response.write ""
    response.write "                                </ul>"
    response.write "                            </div>"


    response.write "                            <!-- start: HR tab -->"
    response.write "                            <div class=""tab-pane fade"" id=""tab_predictions"" role=""tabpanel"">"
    ' response.write "                                <h6 class=""fl-title title-font ps-2 small text-uppercase text-muted"" style=""--dynamic-color: var(--theme-color1)"">Usual</h6>"
    response.write "                                <ul class=""list-unstyled mb-4 menu-list"">"
    response.write "                                    <li>"
    response.write "                                        <a  href=""#ForcastAdmissions"" data-bs-toggle=""collapse"" aria-expanded=""false"" class=""dropdown-toggle"" aria-label=""Users"">"
    response.write "                                            <p> &nbsp;</p>"
    response.write "                                            <svg xmlns=""http://www.w3.org/2000/svg"" width=""16"" height=""16"" fill=""currentColor"" class=""bi bi-hospital"" viewBox=""0 0 16 16"">"
    response.write "                                                <path d=""M8.5 5.034v1.1l.953-.55.5.867L9 7l.953.55-.5.866-.953-.55v1.1h-1v-1.1l-.953.55-.5-.866L7 7l-.953-.55.5-.866.953.55v-1.1zM13.25 9a.25.25 0 0 0-.25.25v.5c0 .138.112.25.25.25h.5a.25.25 0 0 0 .25-.25v-.5a.25.25 0 0 0-.25-.25zM13 11.25a.25.25 0 0 1 .25-.25h.5a.25.25 0 0 1 .25.25v.5a.25.25 0 0 1-.25.25h-.5a.25.25 0 0 1-.25-.25zm.25 1.75a.25.25 0 0 0-.25.25v.5c0 .138.112.25.25.25h.5a.25.25 0 0 0 .25-.25v-.5a.25.25 0 0 0-.25-.25zm-11-4a.25.25 0 0 0-.25.25v.5c0 .138.112.25.25.25h.5A.25.25 0 0 0 3 9.75v-.5A.25.25 0 0 0 2.75 9zm0 2a.25.25 0 0 0-.25.25v.5c0 .138.112.25.25.25h.5a.25.25 0 0 0 .25-.25v-.5a.25.25 0 0 0-.25-.25zM2 13.25a.25.25 0 0 1 .25-.25h.5a.25.25 0 0 1 .25.25v.5a.25.25 0 0 1-.25.25h-.5a.25.25 0 0 1-.25-.25z""/>"
    response.write "                                                <path d=""M5 1a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v1a1 1 0 0 1 1 1v4h3a1 1 0 0 1 1 1v7a1 1 0 0 1-1 1H1a1 1 0 0 1-1-1V8a1 1 0 0 1 1-1h3V3a1 1 0 0 1 1-1zm2 14h2v-3H7zm3 0h1V3H5v12h1v-3a1 1 0 0 1 1-1h2a1 1 0 0 1 1 1zm0-14H6v1h4zm2 7v7h3V8zm-8 7V8H1v7z""/>"
    response.write "                                              </svg>"
    response.write "                                            <span class=""mx-3"">Visual Analytics</span>"
    response.write "                                        </a>"
    response.write "                                        <ul class=""collapse list-unstyled"" id=""ForcastAdmissions"">"
    ' Response.Write "                                            <li><a href='#' onclick=""openPage('wpgPrtPrintLayoutAll.asp?PrintLayoutName=AdmissionTimeSeries&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1')"">Test Time Series Analysis </a></li>"
    ' response.write "                                            <li><a href='#' onclick=""openPage('http://192.168.5.170:8501')"">ML Tools</a></li>"
    response.write "                                            <li><a href='#' onclick=""openPage('http://192.168.5.13:8501')"">ML Tools</a></li>"
    ' response.write "                                            <li><a href=""test.html"" aria-label=""Admissions"">By Age Range</a></li>"
    ' response.write "                                            <li><a href=""test.html"" aria-label=""Time Series Analysis"">Time Series Analysis</a></li>"
    ' response.write "                                            <li><a href=""test.html"" aria-label=""Length Of Stay (LOS)"">Length Of Stay (LOS)</a></li>"
    response.write "                                        </ul>"
    response.write "                                    </li>"
    response.write "                                </ul>"
    response.write "                            </div>"
End If


response.write "                        </div>"
response.write "                    </div>"
response.write "                </div>"
response.write "            </nav>"
response.write "        </aside>"
response.write ""
response.write "        <div class=""content"">"
response.write "            <!-- start: page header -->"
response.write "            <header class=""px-xl-5 px-lg-4 px-3 sticky-top"" data-bs-theme=""none"">"
response.write "                <div class=""d-flex justify-content-between align-items-center py-2 w-100"">"
response.write "                    <div class=""d-flex align-items-center pe-4"">"
response.write "                        <button class=""btn d-inline-flex d-xl-none border-0 p-0 pe-2"" type=""button"" data-bs-toggle=""offcanvas"" data-bs-target=""#offcanvas_Navbar"">"
response.write "                            <svg xmlns=""http://www.w3.org/2000/svg"" class=""svg-stroke"" width=""24"" height=""24"" viewBox=""0 0 24 24"" stroke=""currentColor"" fill=""none"" stroke-linecap=""round"" stroke-linejoin=""round""><path stroke=""none"" d=""M0 0h24v24H0z"" fill=""none""/><path d=""M6 21a3 3 0 0 1 -3 -3v-12a3 3 0 0 1 3 -3"" /><path d=""M21 6v12a3 3 0 0 1 -6 0v-12a3 3 0 0 1 6 0z"" fill=""var(--accent-color)"" /><path d=""M15 12h-8"" /><path d=""M10 9l-3 3l3 3"" /></svg>"
response.write "                        </button>"
response.write "                        <button class=""btn d-xl-inline-flex d-none border-0 p-0 pe-2 sidebar-toggle-btn"" type=""button"">"
response.write "                            <svg xmlns=""http://www.w3.org/2000/svg"" class=""svg-stroke"" width=""24"" height=""24"" viewBox=""0 0 24 24"" stroke=""currentColor"" fill=""none"" stroke-linecap=""round"" stroke-linejoin=""round""><path stroke=""none"" d=""M0 0h24v24H0z"" fill=""none""/><path d=""M6 21a3 3 0 0 1 -3 -3v-12a3 3 0 0 1 3 -3"" /><path d=""M21 6v12a3 3 0 0 1 -6 0v-12a3 3 0 0 1 6 0z"" fill=""var(--accent-color)"" /><path d=""M15 12h-8"" /><path d=""M10 9l-3 3l3 3"" /></svg>"
response.write "                        </button>"
response.write "                        <!--[ Start:: Brand Logo icon ]-->"


response.write "                    </div>"
response.write "                    <form class=""dropdown main-search me-md-4 w-100 d-none d-md-inline-flex"">"
response.write "                    </form>"
response.write "                    <ul class=""header-menu flex-grow-1"">"
response.write "                        <!--[ Start:: notification ]-->"
response.write "                        <li class=""nav-item dropdown px-md-1 d-none d-md-inline-flex"">"
response.write "                            <a class=""dropdown-toggle gray-6"" href=""#"" role=""button"" data-bs-toggle=""dropdown"" aria-expanded=""false"" title=""notification"">"
response.write "                                <span class=""bullet-dot bg-primary animation-blink""></span>"
response.write "                                <svg class=""svg-stroke"" xmlns=""http://www.w3.org/2000/svg"" width=""22"" height=""22"" viewBox=""0 0 24 24"" stroke=""currentColor"" fill=""none"" stroke-linecap=""round"" stroke-linejoin=""round"">"
response.write "                                    <path stroke=""none"" d=""M0 0h24v24H0z"" fill=""none""></path>"
response.write "                                    <path d=""M10 5a2 2 0 0 1 4 0a7 7 0 0 1 4 6v3a4 4 0 0 0 2 3h-16a4 4 0 0 0 2 -3v-3a7 7 0 0 1 4 -6""></path>"
response.write "                                    <path d=""M9 17v1a3 3 0 0 0 6 0v-1""></path>"
response.write "                                    <path d=""M21 6.727a11.05 11.05 0 0 0 -2.794 -3.727""></path>"
response.write "                                    <path d=""M3 6.727a11.05 11.05 0 0 1 2.792 -3.727""></path>"
response.write "                                </svg>"
response.write "                            </a>"
response.write "                            <div class=""dropdown-menu shadow rounded-4 notification"" id=""NotificationsDiv"">"
response.write "                                <div class=""card border-0"">"
response.write "                                    <div class=""card-header d-flex justify-content-between align-items-center py-3"">"
response.write "                                        <h4 class=""mb-0 text-gradient title-font"">Notifications</h4>"
response.write "                                        <a href=""#"" class=""btn btn-link"" title=""view all"">View all</a>"
response.write "                                    </div>"
response.write "                                    <ul class=""card-body p-0 list-unstyled mb-0 custom_scroll ps-2"" >"
' Response.Write "                                    <ul class=""card-body p-0 list-unstyled mb-0 custom_scroll ps-2"" style=""height: 400px;"">"
response.write "                                        <li class=""pe-2"">"
response.write "                                            <a class=""d-flex p-lg-3 p-2 rounded-3"" href=""javascript:void(0);"">"
response.write "                                                <div class=""avatar sm"">"
response.write "                                                         <i class=""icon-calendar""></i>"
response.write "                                                </div>"
response.write "                                                <div class=""flex-fill ms-3"">"
response.write "                                                    <span class=""d-flex justify-content-between""><small class=""text-primary"">Password Expiry</small><small class=""text-muted"">11:30 AM Today</small></span>"
response.write "                                                    <p class=""mb-0 mt-1"">Change your Password</p>"
response.write "                                                </div>"
response.write "                                            </a>"
response.write "                                        </li>"
response.write "                                    </ul>"
response.write "                                </div>"
response.write "                            </div>"
response.write "                        </li>"
response.write "                        <li class=""nav-item py-2 py-md-1 col-auto"">"
response.write "                            <div class=""vr d-none d-sm-flex h-100 mx-sm-2""></div>"
response.write "                        </li>"
response.write "                        <!--[ Start:: theme light/dark ]-->"
response.write "                        <li class=""nav-item dropdown px-md-1"">"
response.write "                            <a class=""dropdown-toggle gray-6"" href=""#"" id=""bd-theme"" data-bs-toggle=""dropdown"" aria-expanded=""false"">"
response.write "                                <svg width=""20"" height=""20"" fill=""currentColor"" viewBox=""0 0 24 24"" class=""theme-icon-active""><use href=""#sun-fill""></use></svg>"
response.write "                            </a>"
response.write "                            <ul class=""dropdown-menu dropdown-menu-end p-2 p-xl-3 shadow rounded-4"" aria-labelledby=""bd-theme"">"
response.write "                                <li class=""mb-1""><a class=""dropdown-item rounded-pill"" href=""#"" data-bs-theme-value=""light""  onclick=""ReloadPage()""><svg class=""avatar xs me-2 opacity-50 theme-icon"" fill=""currentColor""><use href=""#sun-fill""></use></svg> Light</a></li>"
response.write "                                <li class=""mb-1""><a class=""dropdown-item rounded-pill active"" href=""#"" data-bs-theme-value=""dark""  onclick=""ReloadPage()""><svg class=""avatar xs me-2 opacity-50 theme-icon"" fill=""currentColor""><use href=""#moon-stars-fill""></use></svg> Dark</a></li>"
response.write "                                <li class=""mb-1""><a class=""dropdown-item rounded-pill"" href=""#"" data-bs-theme-value=""auto""  onclick=""ReloadPage()""><svg class=""avatar xs me-2 opacity-50 theme-icon"" fill=""currentColor""><use href=""#circle-half""></use></svg> Auto</a></li>"
response.write "                            </ul>"
response.write "                            <svg xmlns=""http://www.w3.org/2000/svg"" width=""16"" height=""16"" fill=""currentColor"" viewBox=""0 0 16 16"" style=""display: none;"">"
response.write "                                <symbol id=""sun-fill"" viewBox=""0 0 16 16"">"
response.write "                                    <path d=""M12 8a4 4 0 1 1-8 0 4 4 0 0 1 8 0zM8 0a.5.5 0 0 1 .5.5v2a.5.5 0 0 1-1 0v-2A.5.5 0 0 1 8 0zm0 13a.5.5 0 0 1 .5.5v2a.5.5 0 0 1-1 0v-2A.5.5 0 0 1 8 13zm8-5a.5.5 0 0 1-.5.5h-2a.5.5 0 0 1 0-1h2a.5.5 0 0 1 .5.5zM3 8a.5.5 0 0 1-.5.5h-2a.5.5 0 0 1 0-1h2A.5.5 0 0 1 3 8zm10.657-5.657a.5.5 0 0 1 0 .707l-1.414 1.415a.5.5 0 1 1-.707-.708l1.414-1.414a.5.5 0 0 1 .707 0zm-9.193 9.193a.5.5 0 0 1 0 .707L3.05 13.657a.5.5 0 0 1-.707-.707l1.414-1.414a.5.5 0 0 1 .707 0zm9.193 2.121a.5.5 0 0 1-.707 0l-1.414-1.414a.5.5 0 0 1 .707-.707l1.414 1.414a.5.5 0 0 1 0 .707zM4.464 4.465a.5.5 0 0 1-.707 0L2.343 3.05a.5.5 0 1 1 .707-.707l1.414 1.414a.5.5 0 0 1 0 .708z""/>"
response.write "                                </symbol>"
response.write "                                <symbol id=""moon-stars-fill"" viewBox=""0 0 16 16"">"
response.write "                                    <path d=""M6 .278a.768.768 0 0 1 .08.858 7.208 7.208 0 0 0-.878 3.46c0 4.021 3.278 7.277 7.318 7.277.527 0 1.04-.055 1.533-.16a.787.787 0 0 1 .81.316.733.733 0 0 1-.031.893A8.349 8.349 0 0 1 8.344 16C3.734 16 0 12.286 0 7.71 0 4.266 2.114 1.312 5.124.06A.752.752 0 0 1 6 .278z""/>"
response.write "                                    <path d=""M10.794 3.148a.217.217 0 0 1 .412 0l.387 1.162c.173.518.579.924 1.097 1.097l1.162.387a.217.217 0 0 1 0 .412l-1.162.387a1.734 1.734 0 0 0-1.097 1.097l-.387 1.162a.217.217 0 0 1-.412 0l-.387-1.162A1.734 1.734 0 0 0 9.31 6.593l-1.162-.387a.217.217 0 0 1 0-.412l1.162-.387a1.734 1.734 0 0 0 1.097-1.097l.387-1.162zM13.863.099a.145.145 0 0 1 .274 0l.258.774c.115.346.386.617.732.732l.774.258a.145.145 0 0 1 0 .274l-.774.258a1.156 1.156 0 0 0-.732.732l-.258.774a.145.145 0 0 1-.274 0l-.258-.774a1.156 1.156 0 0 0-.732-.732l-.774-.258a.145.145 0 0 1 0-.274l.774-.258c.346-.115.617-.386.732-.732L13.863.1z""/>"
response.write "                                </symbol>"
response.write "                                <symbol id=""circle-half"" viewBox=""0 0 16 16"">"
response.write "                                    <path d=""M8 15A7 7 0 1 0 8 1v14zm0 1A8 8 0 1 1 8 0a8 8 0 0 1 0 16z""/>"
response.write "                                </symbol>"
response.write "                            </svg>"
response.write "                        </li>"
response.write ""
response.write "                        <li class=""nav-item py-2 py-md-1 col-auto"">"
response.write "                            <div class=""vr d-none d-sm-flex h-100 mx-sm-2""></div>"
response.write "                        </li>"
response.write "                        <!--[ Start:: user detail ]-->"
response.write "                        <li class=""nav-item user ms-3"">"
response.write "                            <a class=""dropdown-toggle gray-6 d-flex text-decoration-none align-items-center lh-sm p-0"" href=""#"" role=""button"" data-bs-toggle=""dropdown"" aria-expanded=""false"" title=""User"" data-bs-auto-close=""outside"">"
' response.write "                                <img class=""avatar rounded-circle border border-3"" src=""BusinessAnalytics/assets/images/profile_av.png"" alt=""avatar"">"
' response.write "                                <img class=""avatar rounded-circle border border-3"" src="""" alt=""avatar"">"
response.write "                                <span class=""ms-2 fs-6 d-none d-sm-inline-flex"">" & UCase(uName) & "</span>"
response.write "                            </a>"
response.write "                            <div class=""dropdown-menu dropdown-menu-end shadow p-2 p-xl-3 rounded-4"">"
response.write "                                <div class=""bg-body p-3 rounded-3"">"
response.write "                                    <h4 class=""mb-1 title-font text-gradient"">" & UCase(uName) & "</h4>"
response.write "                                    <p class=""small text-muted"">" & uName & "@spagad.com</p>"
response.write "                                </div>"
response.write "                                <ul class=""list-unstyled mt-3"">"
response.write "                                    <li><a class=""dropdown-item rounded-pill"" aria-label=""common settings"" >Profile: " & GetComboName("JobSchedule", jSchd) & "</a></li>"
' response.write "                                    <li><a class=""dropdown-item rounded-pill"" aria-label=""common settings"" >Hospital:   TRUST HOSPITAL - OSU</a></li>"
response.write "                                    <li><a class=""dropdown-item rounded-pill"" aria-label=""common settings"">Date: " & GetLoginDate() & "</a></li>"
response.write "                                </ul>"
response.write "                                <a class=""btn py-2 btn-primary w-100 mt-3 rounded-pill"" role=""button""  onclick=""LogoutFunction()"">Logout</a>"
response.write "                                <div class=""mt-3 text-center small"">"
response.write "                                    <!-- <a class=""text-muted me-1"" href=""#!"">Privacy policy</a> <a class=""text-muted mx-1"" href=""#!"">Terms</a> <a class=""text-muted ms-1"" href=""#!"">Cookies</a> -->"
response.write "                                </div>"
response.write "                            </div>"
response.write "                        </li>"
response.write "                    </ul>"
response.write "                </div>"
response.write "            </header>"

response.write "    <section class='main_content' id='mainContent'>"
' response.write "        <div class='container-fluid g-0'>"
' response.write "            <div class='row'>"
' response.write "                <div class='col-12'>"
' response.write "                    <h1>Welcome to Hospital Dashboard</h1>"
' response.write "                    <h1><p>Some cards will be here.</p></h1>"
' response.write "                </div>"
' response.write "            </div>"
' response.write "        </div>"
response.write "    </section>"



response.write ""
response.write "    <script src='BusinessAnalytics/assets/bundles/libscripts.bundle.js'></script>"
response.write "    <script src='BusinessAnalytics/assets/bundles/apexcharts.bundle.js'></script>"
response.write "    <script src='BusinessAnalytics/assets/bundles/dataTables.bundle.js'></script>"
response.write ""
response.write "    <!-- Template page js -->"
response.write "    <script src='BusinessAnalytics/assets/js/main.js'></script>"


response.write "    <script>"
response.write "       function hideRowsInTargetTable() {"
    response.write "         let form = parent.document.getElementById('form1');"
    response.write "         if (form) {"
    response.write "         let tables = form.getElementsByTagName('table');"
    response.write "         Array.from(tables).forEach((tbl) => {"
    response.write "             tbl.style.border = 'none';"
    response.write "         });"
    response.write "         for (let table of tables) {"
    response.write "           if ("
    response.write "             table.getAttribute('width') === '100%' &&"
    response.write "             table.getAttribute('height') === '100%'"
    response.write "           ) {"
    response.write "             let rows = table.getElementsByTagName('tr');"
    response.write "             if (rows.length > 0) {"
    response.write "               rows[0].style.display = 'none';"
    response.write "               let lastRow = rows[rows.length - 1];"
    response.write "               let parentParentTr = lastRow.parentElement;"
    response.write "               if (parentParentTr && parentParentTr.parentElement) {"
    response.write "                 parentParentTr = parentParentTr.parentElement;"
    response.write "                 parentParent1Tr = parentParentTr.parentElement;"
    response.write "                 parentParent2Tr = parentParent1Tr.parentElement;"
    response.write "                 parentParent2Tr.style.display = 'none';"
    response.write "               } else {"
    response.write "                 console.error('Unable to find grandparent TR element.');"
    response.write "               }"
    response.write "             } else {"
    response.write "               console.error('No rows found in the target table.');"
    response.write "             }"
    response.write "             break;"
    response.write "           }"
    response.write "         }"
    response.write "         }"
    response.write "       }"

   

    response.write "       function ReloadPage() {"
    response.write "                window.location.reload(); "
    response.write "       }"

 

 response.write "    const userid  =  '" & UCase(uName) & "';"

response.write "    document.addEventListener('DOMContentLoaded', function() {" ' Run after the DOM is fully loaded

response.write "        if (userid == 'PHABA01' || userid == 'PHABA012' ) {"
response.write "             var defaultPage = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=PrescriptionDashboardBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1';"
response.write "        }else if(userid == 'LABBA01'){"
response.write "              var defaultPage = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=LabolatoryDashboardBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1';"
response.write "         }else {"
response.write "              var defaultPage = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=PrescriptionDashboardBA&PositionForTableName=WorkingDay&WorkingDayID=&selectedValue=" & periodStart & "&selectedValue1=" & periodEnd & "&branID=" & brnchID & "&NoGlobal=1';"
response.write "        }"

response.write "        var defaultLink = document.querySelector('a[href=\""#Dashboard\""]');" ' Find the Dashboard link
response.write "        if (defaultLink) {"
response.write "            defaultLink.classList.add('active');" ' Add active class to the default link
response.write "            openPage(defaultPage, defaultLink);" ' Load the default page
response.write "        }"
response.write "    });"

    response.write "    function openPage(page, linkElement) {"
    response.write "        var spinner = document.getElementById('spinner');"
    response.write "        spinner.style.display = 'flex';" ' Show spinner when function is called
    response.write "        document.querySelectorAll('a.active').forEach(function(link) {"
    response.write "            link.classList.remove('active');"
    response.write "        });"
    response.write "        var main_tab_analysis = document.getElementById('main_tab_analysis');"
    response.write "            main_tab_analysis.classList.add('active');"

    response.write "        if (linkElement) {"
    response.write "            linkElement.classList.add('active');"
    response.write "        }"
    response.write "        var iframe = document.createElement('iframe');"
    response.write "        iframe.src = page;"
    response.write "        iframe.style.width = '100%';"
    response.write "        iframe.style.height = 'calc(100vh - 60px)';"
    response.write "        iframe.style.border = 'none';"
    response.write "        iframe.onload = function() {"
    response.write "            spinner.style.display = 'none';"
    response.write "        };"
    response.write "        var mainContent = document.getElementById('mainContent');"
    response.write "        mainContent.innerHTML = '';"
    response.write "        mainContent.appendChild(iframe);"
    response.write "    }"



    response.write "       function LogoutFunction() {"
    response.write "           let mssg = 'Are you sure you want to Logout of the system?';"
    response.write "           let userConfirmed = confirm(mssg);"
    response.write "           if (userConfirmed) {"
    response.write "               openInSelf('" & processNavUrl("wpgLogout.POSFORFILEEXT") & "');"
    response.write "           } "
    response.write "       }"

    response.write "       function openInSelf(url) {"
    response.write "           window.open(url, '_self');"
    response.write "       }"

    response.write "       hideRowsInTargetTable();"
    
    response.write "    </script>"

response.write "<script>"
response.write "    $(document).ready(function() {"
response.write "        var options = {"
response.write "            chart: {"
response.write "                height: 280,"
response.write "                type: 'line',"
response.write "                toolbar: {"
response.write "                    show: false"
response.write "                }"
response.write "            },"
response.write "            colors: ['var(--theme-color4)', 'var(--theme-color5)'],"
response.write "            series: [{"
response.write "                name: 'Appointment Booking',"
response.write "                type: 'column',"
response.write "                data: [440, 505, 414, 671, 227, 413, 201, 352, 752, 320, 257, 160]"
response.write "            }, {"
response.write "                name: 'Visits Created',"
response.write "                type: 'line',"
response.write "                data: [23, 42, 35, 27, 43, 22, 17, 31, 22, 22, 12, 16]"
response.write "            }],"
response.write "            stroke: {"
response.write "                width: [0, 4]"
response.write "            },"
response.write "            labels: ['01 Jan 2001', '02 Jan 2001', '03 Jan 2001', '04 Jan 2001', '05 Jan 2001', '06 Jan 2001', '07 Jan 2001', '08 Jan 2001', '09 Jan 2001', '10 Jan 2001', '11 Jan 2001', '12 Jan 2001'],"
response.write "            xaxis: {"
response.write "                type: 'datetime'"
response.write "            },"
response.write "            yaxis: [{"
response.write "                title: {"
response.write "                    text: 'Appointment Booking'"
response.write "                }"
response.write "            }, {"
response.write "                opposite: true,"
response.write "                title: {"
response.write "                    text: 'Visits Created'"
response.write "                }"
response.write "            }]"
response.write "        };"
response.write "        var chart = new ApexCharts(document.querySelector('#reservations'), options);"
response.write "        chart.render();"
response.write "    });"
response.write "    $(document).ready(function() {"
response.write "        $('.dataTable')"
response.write "            .addClass('nowrap')"
response.write "            .dataTable({"
response.write "                responsive: true"
response.write "            });"
response.write "    });"
response.write "</script>"

response.write ""
response.write "</body>"
response.write ""
response.write "</html>"

Function processNavUrl(url)
    Dim ot, rstm
    
    ot = url
    kyVal = "asp"

    If InStr(ot, "POSFORFILEEXT") > 0 Then
        ot = Replace(ot, "POSFORFILEEXT", kyVal)
    End If
    
    processNavUrl = ProcessCliUrl(ot)
End Function

'Get Excluded Branch ID
Function GetLoginDate()
    Dim rst, sql, ot
   
    ot = False
    
    sql = sql & " SELECT FORMAT(su.LastLoginDate, 'yyyy-MM-dd HH:mm tt') AS LastLoginDate from SystemUser su "
    sql = sql & " WHERE JobScheduleID = '" & jSchd & "'"
    
    Set rst = CreateObject("ADODB.Recordset")
    With rst
        .open sql, conn, 3, 4
        If .RecordCount > 0 Then
            .MoveFirst
            ot = .Fields("LastLoginDate")
            .Close
        End If
    End With
    
    GetLoginDate = ot
    
    Set rst = Nothing
End Function

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>

