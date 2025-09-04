VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "USRPRO_ValidateIncomingDrug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

ValidateIncomingDrug

Sub ValidateIncomingDrug()
    Dim vld, inDrg, purID
    
    vld = True
    
    purID = Trim(Request("inpDrugPurOrderID"))
    
    If Not HasValidApproval(purID) Then
        vld = False
    End If
    
    If Not vld Then
        If objPage.rtnHdlProcessPoint Then
              objPage.hdlProcessPoint = False
        End If
    End If
End Sub
Function HasValidApproval(purID)
    Dim sql, rst, rst2, maxPurLm
    
    ot = False
   
    sql = "select * from DrugPurOrder where DrugPurOrderID='" & purID & "' "
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        'CEO's approval required
        If UCase(rst.fields("TransProcessValID")) = UCase("DrugPurOrderPro-T004") Then
            ot = True
        Else
            SetPageMessages "CEO's approval is required."
        End If
           
    Else
        SetPageMessages "Cannot find the purchase order for this incoming drug."
    End If
    
    rst.Close
    Set rst = Nothing
    
    HasValidApproval = ot
End Function
