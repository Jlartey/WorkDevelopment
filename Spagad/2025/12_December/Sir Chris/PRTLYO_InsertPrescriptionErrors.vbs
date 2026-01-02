'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Option Explicit

Response.Clear
Response.Charset = "UTF-8"

Dim isPost: isPost = (Request.ServerVariables("REQUEST_METHOD") = "POST")

' --- FORM VARIABLES ---
Dim VisitationID, DrugID, Dose, Frequency, Route
Dim AppDrug, AppDose, AppFreq, AppRoute, TherDup
Dim Comments, InterventionOutcome
Dim errors, successMessage

VisitationID = Request.QueryString("VisitationID")
DrugID = "": Dose = ""
Frequency = "": Route = ""
AppDrug = "": AppDose = "": AppFreq = ""
AppRoute = "": TherDup = ""
Comments = "": InterventionOutcome = ""
errors = "": successMessage = ""

If isPost Then
    'VisitationID = Trim(Request.Form("VisitationID"))
    DrugID = Trim(Request.Form("DrugID"))
    Dose = Trim(Request.Form("Dose"))
    Frequency = Trim(Request.Form("Frequency"))
    Route = Trim(Request.Form("Route"))
    AppDrug = Trim(Request.Form("AppDrug"))
    AppDose = Trim(Request.Form("AppDose"))
    AppFreq = Trim(Request.Form("AppFreq"))
    AppRoute = Trim(Request.Form("AppRoute"))
    TherDup = Trim(Request.Form("TherDup"))
    Comments = Trim(Request.Form("Comments"))
    InterventionOutcome = Trim(Request.Form("InterventionOutcome"))

    'If VisitationID = "" Then errors = errors & "� Visitation ID is required.<br>"
    If DrugID = "" Then errors = errors & "� Drug is required.<br>"
    If AppDrug = "" Then errors = errors & "� Appropriateness of Drug is required.<br>"
    If AppDose = "" Then errors = errors & "� Appropriateness of Dose is required.<br>"
    If AppFreq = "" Then errors = errors & "� Appropriateness of Frequency is required.<br>"
    If AppRoute = "" Then errors = errors & "� Appropriateness of Route is required.<br>"
    If TherDup = "" Then errors = errors & "� Therapeutic Duplication is required.<br>"

    If errors = "" Then
        Dim cmd
        Set cmd = Server.CreateObject("ADODB.Command")
        cmd.ActiveConnection = conn
        cmd.CommandType = 1

        cmd.CommandText = _
            "INSERT INTO TherapyAppropriatenessAudit " & _
            "(VisitationID, DrugID, Dose, Frequency, Route, " & _
            "AppropriatenessOfDrug, AppropriatenessOfDose, " & _
            "AppropriatenessOfFrequency, AppropriatenessOfRoute, " & _
            "TherapeuticDuplication, Comments, interventionOutcome) " & _
            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?)"

        cmd.Parameters.Append cmd.CreateParameter("", 200, 1, 50, VisitationID)
        cmd.Parameters.Append cmd.CreateParameter("", 200, 1, 50, DrugID)
        cmd.Parameters.Append cmd.CreateParameter("", 200, 1, 50, Dose)
        cmd.Parameters.Append cmd.CreateParameter("", 200, 1, 50, Frequency)
        cmd.Parameters.Append cmd.CreateParameter("", 200, 1, 50, Route)
        cmd.Parameters.Append cmd.CreateParameter("", 200, 1, 255, AppDrug)
        cmd.Parameters.Append cmd.CreateParameter("", 200, 1, 255, AppDose)
        cmd.Parameters.Append cmd.CreateParameter("", 200, 1, 255, AppFreq)
        cmd.Parameters.Append cmd.CreateParameter("", 200, 1, 255, AppRoute)
        cmd.Parameters.Append cmd.CreateParameter("", 200, 1, 255, TherDup)
        cmd.Parameters.Append cmd.CreateParameter("", 200, 1, 250, Comments)
        cmd.Parameters.Append cmd.CreateParameter("", 200, 1, 250, InterventionOutcome)

        cmd.Execute
        successMessage = "Therapy Appropriateness Audit saved successfully."

        Set cmd = Nothing
    End If
End If

' -------- DROPDOWN OPTIONS --------
Dim options, freqOptions, routeOptions

options = "<option value="""">-- Select --</option>" & _
          "<option>Appropriate</option>" & _
          "<option>Inappropriate</option>" & _
          "<option>Uncertain</option>" & _
          "<option>Not Applicable</option>"

freqOptions = "<option value="""">-- Select --</option>" & _
              "<option>BID</option>" & _
              "<option>OD</option>" & _
              "<option>PRN</option>" & _
              "<option>TID</option>"

routeOptions = "<option value="""">-- Select --</option>" & _
               "<option>IM</option>" & _
               "<option>IV</option>" & _
               "<option>Oral</option>" & _
               "<option>SC</option>" & _
               "<option>Subcutaneous</option>"

' -------- RENDER PAGE --------
Response.Write "<!DOCTYPE html>"
Response.Write "<html><head>"
Response.Write "<meta charset='UTF-8'>"
Response.Write "<title>Therapy Appropriateness Audit</title>"
Response.Write "<link href='https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css' rel='stylesheet'>"
Response.Write "</head>"

Response.Write "<body class='bg-light p-4'>"
Response.Write "<div class='card mx-auto shadow' style='max-width:800px;'>"
Response.Write "<div class='card-header bg-primary text-white'><h5 class='mb-0'>Therapy Appropriateness Audit</h5></div>"
Response.Write "<div class='card-body'>"

If errors <> "" Then
    Response.Write "<div class='alert alert-danger'>" & errors & "</div>"
ElseIf successMessage <> "" Then
    Response.Write "<div class='alert alert-success'>" & successMessage & "</div>"
End If

Response.Write "<form method='post'><div class='row g-3'>"

Response.Write "<div class='col-md-6'><label class='form-label'>Visitation ID *</label>"
Response.Write "<input class='form-control' name='VisitationID' value='" & Server.HTMLEncode(VisitationID) & "' disabled></div>"

Response.Write "<div class='col-md-6'><label class='form-label'>Drug ID *</label>"
Response.Write "<input class='form-control' name='DrugID' value='" & Server.HTMLEncode(DrugID) & "'></div>"

Response.Write "<div class='col-md-4'><label class='form-label'>Dose</label>"
Response.Write "<input class='form-control' name='Dose' value='" & Server.HTMLEncode(Dose) & "'></div>"

Response.Write "<div class='col-md-4'><label class='form-label'>Frequency</label>"
Response.Write "<select class='form-select' name='Frequency'>" & freqOptions & "</select></div>"

Response.Write "<div class='col-md-4'><label class='form-label'>Route</label>"
Response.Write "<select class='form-select' name='Route'>" & routeOptions & "</select></div>"

Response.Write "<div class='col-md-6'><label class='form-label'>Appropriateness of Drug *</label>"
Response.Write "<select class='form-select' name='AppDrug'>" & options & "</select></div>"

Response.Write "<div class='col-md-6'><label class='form-label'>Appropriateness of Dose *</label>"
Response.Write "<select class='form-select' name='AppDose'>" & options & "</select></div>"

Response.Write "<div class='col-md-6'><label class='form-label'>Appropriateness of Frequency *</label>"
Response.Write "<select class='form-select' name='AppFreq'>" & options & "</select></div>"

Response.Write "<div class='col-md-6'><label class='form-label'>Appropriateness of Route *</label>"
Response.Write "<select class='form-select' name='AppRoute'>" & options & "</select></div>"

Response.Write "<div class='col-md-6'><label class='form-label'>Therapeutic Duplication *</label>"
Response.Write "<select class='form-select' name='TherDup'>"
Response.Write "<option value="""">-- Select --</option>"
Response.Write "<option>Duplicate Therapy</option>"
Response.Write " <option>Potential</option>"
Response.Write "<option>None</option>"
Response.Write "</select></div>"

Response.Write "<div class='col-12'><label class='form-label'>Comments</label>"
Response.Write "<textarea class='form-control' name='Comments'>" & Server.HTMLEncode(Comments) & "</textarea></div>"

Response.Write "<div class='col-12'><label class='form-label'>Intervention Outcome</label>"
Response.Write "<textarea class='form-control' name='InterventionOutcome'>" & Server.HTMLEncode(InterventionOutcome) & "</textarea></div>"

Response.Write "<div class='col-12 text-end'><button class='btn btn-primary'>Save Audit</button></div>"
Response.Write "</div></form></div></div></body></html>"

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>



