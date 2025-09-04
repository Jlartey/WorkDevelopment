'<<--CONFIGURATION SCRIPT-->>
Dim replicateTbl, repWhCls, repPriKyLst, newPriKyVlLst

Dim Request

config
'ConvertNavToMenu3 "s01head"

Sub config()
    Dim sql, rst, source, target, tmp, phoneNo, idx, rst2
    '

    Exit Sub
    GrantAccessToTransition "DPT040", "ReceiptPro", "T001", "T003" ' new to cancel
    GrantAccessToTransition "DPT040", "ReceiptPro", "T004", "T003" 'change owner to cancel
    GrantAccessToTransition "DPT040", "ReceiptPro", "T005", "T003" 'change info to cancel
    
    GrantAccessToTransition "DPT040", "ReceiptPro", "T001", "T004" 'new tp change owner
    GrantAccessToTransition "DPT040", "ReceiptPro", "T005", "T004" 'change info to change owner
    
    GrantAccessToTransition "DPT040", "ReceiptPro", "T001", "T005" 'new to change info
    GrantAccessToTransition "DPT040", "ReceiptPro", "T003", "T005" 'refund to change info
    GrantAccessToTransition "DPT040", "ReceiptPro", "T005", "T005" 'change info to change info
    GrantAccessToTransition "DPT040", "ReceiptPro", "T005", "T002" 'change info, refund
    GrantAccessToTransition "DPT040", "ReceiptPro", "T004", "T004" 'change owner, change owner
    GrantAccessToTransition "DPT040", "ReceiptPro", "T004", "T002" 'change owner, change owner
    
    GrantAccessToTransition "DPT040", "ReceiptPro", "T001", "T002"

    Exit Sub
    
    RepAnyPrintOutAlloc "", "M0603"
    RepAnyPrintOutAlloc "", "M0603"
    RepAnyPrintOutAlloc "", "M0603"
    
    Exit Sub
    ConvertNavToMenu3 "S22A"
    'ReplicateProfile "S22", "S22A", "Pharmacy Technician"
    
    Exit Sub
    ConvertNavToMenu3 "MedicalRecords"
    
    Exit Sub
    sql = "select * from JobSchedule where JobScheduleID like '%headhead%'"
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        Do While Not rst.EOF
            DelJobSchedule rst.fields("JobScheduleID")
            'DeleteTableKey "JobSchedule", rst.fields("JobScheduleID")
            rst.MoveNext
        Loop
    End If
    
    rst.Close
    Exit Sub
    dt = "01 Oct 2022 13:19:00"
    MigrateDrugStock "pharmacy_stock$", "DRG-S00001", dt
    MigrateDrugStock "pharmacy_stock_inj$", "DRG-S00002", dt
    Exit Sub
    
    GrantTblAccessToPrintlayout "WorkingDay", "customFindLargeTableKey"
    RepAnyPrintOutAlloc "KenUpdatePack", "SystemAdmin"
    
    Exit Sub
    dt = "01 Oct 2022 13:19:00"
    MigrateDrugStock "pharmacy_stock$", "DRG-S00001", dt
    MigrateDrugStock "pharmacy_stock_inj$", "DRG-S00002", dt
    
    Exit Sub
    DeleteOldData
    
    Exit Sub
    dt = "01 Oct 2020 08:00:00"
    MigrateDrugStock "pharmacy_stock$", "DRG-S00001", dt
    MigrateDrugStock "pharmacy_stock_inj$", "DRG-S00002", dt
    Exit Sub
    ConvertNavToMenu3 "S22"
    Exit Sub
    AddTransitionProcess "DrugPurOrderPro", "T001", "New Purchase Order"
    UpdTransitionProcessName "DrugPurOrderPro", "T001", "New Purchase Order"
    
    AddTransitionProcess "DrugPurOrderPro", "T002", "Approved By Supply Chain Manager"
    UpdTransitionProcessName "DrugPurOrderPro", "T002", "Approved By Supply Chain Manager"
    
    AddTransitionProcess "DrugPurOrderPro", "T003", "Approved By Admin Manager"
    UpdTransitionProcessName "DrugPurOrderPro", "T003", "Approved By Admin Manager"
    
    AddTransitionProcess "DrugPurOrderPro", "T004", "Approved By CEO"
    UpdTransitionProcessName "DrugPurOrderPro", "T004", "Approved By CEO"
    
    AddTransitionProcess "DrugPurOrderPro", "T005", "Cancel Purchase Order"
    UpdTransitionProcessName "DrugPurOrderPro", "T005", "Cancel Purchase Order"
    
    AddTransitionProcess "DrugPurOrderPro", "T006", "Approved By Finance"
    UpdTransitionProcessName "DrugPurOrderPro", "T006", "Approved By Finance"
    
    AddTransitionProcess "DrugPurOrderPro", "T007", "Return To Supply Chain Manager"
    UpdTransitionProcessName "DrugPurOrderPro", "T007", "Return To Supply Chain Manager"
    
    AddTransitionProcess "DrugPurOrderPro", "T008", "Return To Admin"
    UpdTransitionProcessName "DrugPurOrderPro", "T008", "Return To Admin"
    
    AddTransitionProcess "DrugPurOrderPro", "T009", "Return To Finance"
    UpdTransitionProcessName "DrugPurOrderPro", "T009", "Return To Finance"
    
    AddTransitionProcess "DrugPurOrderPro", "T010", "Return To Issuer"
    UpdTransitionProcessName "DrugPurOrderPro", "T010", "Return To Issuer"
    
    
    'items
    AddTransitionProcess "ItemPurOrderPro", "T001", "New Purchase Order"
    UpdTransitionProcessName "ItemPurOrderPro", "T001", "New Purchase Order"
    
    AddTransitionProcess "ItemPurOrderPro", "T002", "Approved By Supply Chain Manager"
    UpdTransitionProcessName "ItemPurOrderPro", "T002", "Approved By Supply Chain Manager"
    
    AddTransitionProcess "ItemPurOrderPro", "T003", "Approved By Admin Manager"
    UpdTransitionProcessName "ItemPurOrderPro", "T003", "Approved By Admin Manager"
    
    AddTransitionProcess "ItemPurOrderPro", "T004", "Approved By CEO"
    UpdTransitionProcessName "ItemPurOrderPro", "T004", "Approved By CEO"
    
    AddTransitionProcess "ItemPurOrderPro", "T005", "Cancel Purchase Order"
    UpdTransitionProcessName "ItemPurOrderPro", "T005", "Cancel Purchase Order"
    
    AddTransitionProcess "ItemPurOrderPro", "T006", "Approved By Finance"
    UpdTransitionProcessName "ItemPurOrderPro", "T006", "Approved By Finance"
    
    AddTransitionProcess "ItemPurOrderPro", "T007", "Return To Supply Chain Manager"
    UpdTransitionProcessName "ItemPurOrderPro", "T007", "Return To Supply Chain Manager"
    
    AddTransitionProcess "ItemPurOrderPro", "T008", "Return To Admin"
    UpdTransitionProcessName "ItemPurOrderPro", "T008", "Return To Admin"
    
    AddTransitionProcess "ItemPurOrderPro", "T009", "Return To Finance"
    UpdTransitionProcessName "ItemPurOrderPro", "T009", "Return To Finance"
    
    AddTransitionProcess "ItemPurOrderPro", "T010", "Return To Issuer"
    UpdTransitionProcessName "ItemPurOrderPro", "T010", "Return To Issuer"
    
    
    
    
    
    GrantAccessToTable "S01Head", "Treatment", "edit"
    GrantAccessToTable "S01Head", "Treatment", "search"
    
    Exit Sub
    ReplicateAccess "Cashier", "S22"
    Exit Sub
    GrantAccessToTable "S01", "BenefitOption", "new"
    GrantAccessToTable "S01", "BenefitOption", "view"
    Exit Sub
    ConvertNavToMenu3 "S01Head"
    Exit Sub
    DeleteTableKey "DrugSale", "D122090070"
    Exit Sub
    ReplicateAccess "W001", "S01Head"
    Exit Sub
    ReplicateProfile "MedicalRecords", "S01Head", "OPD + Medical Records (Head)"
    ReplicateAccess "DPT010", "S01Head"
    ReplicateAccess "S01", "S01Head"
    
    ReplicateProfile "S22", "M0603", "Pharmacist / Pharmacy Head"
    ReplicateAccess "DPT010", "M0603"
    
    sql = "select * from JobSchedule Where JobScheduleID like 'w00%' "
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            ReplicateProfile rst.fields("JobScheduleID"), rst.fields("JobScheduleID") & "Head", rst.fields("JobScheduleName") & " - HEAD"
            ReplicateAccess "DPT010", rst.fields("JobScheduleID") & "Head"
            rst.MoveNext
        Loop
    End If
    rst.Close
    
    Exit Sub
    DeleteTableKey "DrugSale", "D122090044"

    Exit Sub
    GrantAccessToTable "M0303", "Disease", "new"
    GrantAccessToTable "M0303", "Disease", "view"
    
    GrantAccessToTable "M0303", "Disease", "new"
    GrantAccessToTable "M0303", "Disease", "view"
    
    GrantAccessToTable "M0303", "Disease", "new"
    GrantAccessToTable "M0303", "Disease", "view"
    
    GrantAccessToTable "M0303", "Disease", "new"
    GrantAccessToTable "M0303", "Disease", "view"
    GrantAccessToTable "M0303", "Disease", "new"
    GrantAccessToTable "M0303", "Disease", "view"
    GrantAccessToTable "M0303", "Disease", "new"
    GrantAccessToTable "M0303", "Disease", "view"
    
    GrantAccessToTable "M0303", "Disease", "new"
    GrantAccessToTable "M0303", "Disease", "view"
    Exit Sub
    sql = "select * from Receipt where ReceiptID like 'r1%' and PatientID='P1'"
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    MsgBox rst.RecordCount
    If rst.RecordCount > 0 Then
        Do While Not rst.EOF
            DeleteTableKey "Receipt", rst.fields("ReceiptID")
            rst.MoveNext
        Loop
    End If
    
    Exit Sub
    sql = "select * from Patient where PatientID like 'mr/2209/%'"
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        Do While Not rst.EOF
            DeleteTableKey "Patient", rst.fields("PatientID")
            rst.MoveNext
        Loop
    End If
    Exit Sub
    ReplicateAccess "Cashier", "S22"
    Exit Sub
    ConvertNavToMenu3 "S22"
    Exit Sub
    AddProcessCall "Receipt", "ValidateReceipt", "P024"
    AddProcessCall "Receipt", "ValidateReceipt", "P027"
    
    Exit Sub
    RepAnyPrintOutAlloc "ReceiptSlip2", "S22"
    Exit Sub
    GrantAccessToTable "S22", "Receipt", "View"
    GrantAccessToTable "S22", "Receipt", "New"
    
    'AddProcesscallSibling "GeneratePatientInvoice", ""
    Exit Sub
    GrantAccessToTable "S22", "LabTest", "view"
    GrantAccessToTable "S22", "Investigation", "view"
    GrantAccessToTable "S22", "Investigation2", "view"
    GrantAccessToTable "S22", "LabRequest", "new"
    GrantAccessToTable "S22", "LabRequest", "edit"
    GrantAccessToTable "S22", "LabByDoctor", "View"
    
    Exit Sub
    sql = "select JobScheduleID from JobSchedule where JobScheduleID>='W000' and JobScheduleID<='W015' "
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            GrantAccessToTable rst.fields("JobScheduleID"), "FoodOrder", "view"
            GrantAccessToTable rst.fields("JobScheduleID"), "FoodOrder", "new"
            GrantAccessToTable rst.fields("JobScheduleID"), "FoodOrder", "edit"
            GrantAccessToTable rst.fields("JobScheduleID"), "FoodRecipe", "View"
            GrantAccessToTable rst.fields("JobScheduleID"), "FoodOrderItems", "View"
            GrantAccessToTable rst.fields("JobScheduleID"), "FoodOrderItems", "New"
            rst.MoveNext
        Loop
    End If
    
    Exit Sub
    UpdPrintLayout "PACU", "PrintProp", "100%||0||0"
    
    sql = "select JobScheduleID from JobSchedule where JobScheduleID like 'm0%' "
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            GrantAccessToTable rst.fields("JobScheduleID"), "PatientChartPoint", "new"
            rst.MoveNext
        Loop
    End If
    rst.Close
    Set rst = Nothing
    CopyPrintOutAccess "pctChartTypeViewer", "PACU"
    
    Exit Sub
    
    DeleteTableKey "UserProcess", "PACU"
    Exit Sub
    CopyPrintOutAccess "pctChartTypeViewer", "PACU"
    Exit Sub
    CopyPrintOutAccess "Investigation1TH", "PrintLabResults"
    
    Exit Sub
    
    GrantAccessToTransition "DPT014", "DrugPurOrderPro", "T001", "T005"
    GrantAccessToTransition "DPT014", "DrugPurOrderPro", "T001", "T005"
    GrantAccessToTransition "S33", "DrugPurOrderPro", "T001", "T005"
    GrantAccessToTransition "S33", "DrugPurOrderPro", "T001", "T005"
    
    AddTransitionProcess "ItemPurOrderPro", "T005", "Cancel Purchase Order"
    UpdTransitionProcessName "ItemPurOrderPro", "T005", "Cancel Purchase Order"
    AddTransitionProcess "DrugPurOrderPro", "T005", "Cancel Purchase Order"
    UpdTransitionProcessName "DrugPurOrderPro", "T005", "Cancel Purchase Order"
    
    
    Exit Sub
    
    AddProcessCall "IncomingStock", "ValidateIncomingStock", "P024"
    AddProcessCall "IncomingStock", "ValidateIncomingStock", "P027"
    
    AddProcessCall "IncomingStock", "UpdateIncomingStock", "P025"
    AddProcessCall "IncomingStock", "UpdateIncomingStock", "P028"
    AddProcessCall "IncomingStock", "ReloadOpener", "P025"
    AddProcessCall "IncomingStock", "ReloadOpener", "P028"
    
    Exit Sub
    GrantAccessToTable "S20", "ItemPurOrderPro", "View"
    GrantAccessToTable "S21", "DrugPurOrderPro", "View"
    
    RepAnyPrintOutAlloc "ItemPurOrderRCP", "S20"
    RepAnyPrintOutAlloc "DrugPurOrderRCP", "S21"
     
    Exit Sub
    
    AddProcessCall "ItemPurOrder", "ReloadOpener", "P025"
    AddProcessCall "ItemPurOrder", "ReloadOpener", "P028"
    
    AddProcessCall "ItemPurOrder", "ReloadOpener", "P031"
    AddProcessCall "ItemPurOrder", "ReloadOpener", "P032"
    
    AddProcessCall "IncomingDrug", "ReloadOpener", "P031"
    AddProcessCall "IncomingDrug", "ReloadOpener", "P032"
    
    Exit Sub
    DeleteTableKey "UserProcess", "InitItemPurOderPro"
    
    CopyPrintOutAccess "ItemPurOrder", "ItemPurOrderRCP"
    GrantAccessToTable "S33", "ItemPurOrderPro", "view"
    GrantAccessToTable "S33", "ItemPurOrderPro", "new"
    
    AddTransitionProcess "ItemPurOrderPro", "T001", "New Purchase Order"
    UpdTransitionProcessName "ItemPurOrderPro", "T001", "New Purchase Order"
    
    AddTransitionProcess "ItemPurOrderPro", "T002", "Approved By Supply Chain Manager"
    UpdTransitionProcessName "ItemPurOrderPro", "T002", "Approved By Supply Chain Manager"
    
    AddTransitionProcess "ItemPurOrderPro", "T003", "Approved By Admin Manager"
    UpdTransitionProcessName "ItemPurOrderPro", "T003", "Approved By Admin Manager"
    
    AddTransitionProcess "ItemPurOrderPro", "T004", "Approved By CEO"
    UpdTransitionProcessName "ItemPurOrderPro", "T004", "Approved By CEO"
    
    GrantAccessToTransition "DPT014", "ItemPurOrderPro", "T001", "T002"
    GrantAccessToTransition "DPT014", "ItemPurOrderPro", "T002", "T003"
    GrantAccessToTransition "DPT014", "ItemPurOrderPro", "T003", "T004"
    GrantAccessToTransition "DPT014", "ItemPurOrderPro", "T002", "T004"
    
    AddProcessCall "DrugPurOrderPro", "InitDrugPurOrderPro", "P010"
    AddProcessCall "DrugPurOrderPro", "InitDrugPurOrderPro", "P011"
    
    AddProcessCall "DrugPurOrderPro", "ReloadOpener", "P010"
    AddProcessCall "DrugPurOrderPro", "ReloadOpener", "P011"
    
    AddProcessCall "ItemPurOrderPro", "InitItemPurOrderPro", "P010"
    AddProcessCall "ItemPurOrderPro", "InitItemPurOrderPro", "P011"
    
    AddProcessCall "ItemPurOrderPro", "ReloadOpener", "P010"
    AddProcessCall "ItemPurOrderPro", "ReloadOpener", "P011"
    
    Exit Sub
    
    GrantAccessToTable "S33", "UserFileUpload", "new"
    GrantAccessToTable "S33", "UserFileUpload", "view"
    
    Exit Sub
    
    GrantAccessToTransition "DPT014", "DrugPurOrderPro", "T003", "T004"

    Exit Sub
    
    AddProcessCall "IncomingDrug", "ReloadOpener", "P025"
    AddProcessCall "IncomingDrug", "ReloadOpener", "P028"
    
    Exit Sub
    
    AddProcessCall "DrugPurOrder", "ValidateDrugPurOrder", "P024"
    AddProcessCall "DrugPurOrder", "ValidateDrugPurOrder", "P027"
    
    Exit Sub
    AddProcessCall "DrugPurOrder", "ReloadOpener", "P025"
    AddProcessCall "DrugPurOrder", "ReloadOpener", "P028"
    
    Exit Sub
    AddProcessCall "DrugPurOrderPro", "InitCSSDRequestPro", "P010"
    AddProcessCall "DrugPurOrderPro", "InitCSSDRequestPro", "P011"
    
    AddProcessCall "DrugPurOrderPro", "ReloadOpener", "P010"
    AddProcessCall "DrugPurOrderPro", "ReloadOpener", "P011"
    
    Exit Sub
    AddTransitionProcess "DrugPurOrderPro", "T001", "New Purchase Order"
    UpdTransitionProcessName "DrugPurOrderPro", "T001", "New Purchase Order"
    
    AddTransitionProcess "DrugPurOrderPro", "T002", "Approved By Supply Chain Manager"
    UpdTransitionProcessName "DrugPurOrderPro", "T002", "Approved By Supply Chain Manager"
    
    AddTransitionProcess "DrugPurOrderPro", "T003", "Approved By Admin Manager"
    UpdTransitionProcessName "DrugPurOrderPro", "T003", "Approved By Admin Manager"
    
    AddTransitionProcess "DrugPurOrderPro", "T004", "Approved By CEO"
    UpdTransitionProcessName "DrugPurOrderPro", "T004", "Approved By CEO"
    
    GrantAccessToTransition "DPT014", "DrugPurOrderPro", "T001", "T002"
    GrantAccessToTransition "DPT014", "DrugPurOrderPro", "T002", "T003"
'    GrantAccessToTransition "DPT014", "DrugPurOrderPro", "T003", "T004"
    GrantAccessToTransition "DPT014", "DrugPurOrderPro", "T002", "T004"
    Exit Sub
    
    CopyPrintOutAccess "DrugPurOrder", "DrugPurOrderRCP"
    GrantAccessToTable "S33", "DrugPurOrderPro", "view"
    GrantAccessToTable "S33", "DrugPurOrderPro", "new"
    
    
    Exit Sub
    sql = "select distinct TableID from TableField where TableFieldID='TransProcessValID' and DisplayName='TransProcessVal'"
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            'UpdTableField rst.fields("TableID"), "TransProcessValID", "LabelName", "Unit Of Measure"
            UpdTableField rst.fields("TableID"), "TransProcessValID", "DisplayName", "Current Stage"
            rst.MoveNext
        Loop
    End If
    rst.Close
    
    Exit Sub
    'show transition stage on browseviews
    sql = "select distinct TableID from TableField where TableFieldID='TransProcessValID' "
    Set rst = CreateObject("ADODB.RecordSet")
    Set rst2 = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
        
            'sql = "select * from BrowseView where TableID='" & rst.fields("TableID") & "' and (( FieldList not like '%TransProcessValID%' ) or (FieldList like '%||TransProcessValID') )"
            sql = "select * from BrowseView where TableID='" & rst.fields("TableID") & "' and FieldList like '%||TransProcessStatID' "
            rst2.open sql, conn, 3, 4
            If rst2.RecordCount > 0 Then
                rst2.movefirst
                Do While Not rst2.EOF
                    rst2.fields("FieldList") = Replace(rst2.fields("FieldList"), "TransProcessStatID", "TransProcessValID")
                    rst2.UpdateBatch
                    rst2.MoveNext
                Loop
            End If
            rst2.Close
            rst.MoveNext
        Loop
    End If
    rst.Close
    Set rst = Nothing
    
    Exit Sub
    ConvertNavToMenu3 "M16"
    Exit Sub
    
    AddTransitionProcess "CSSDRequestPro", "T007", "Failed Sterilization"
    UpdTransitionProcessName "CSSDRequestPro", "T007", "Failed Sterilization"
    
    GrantAccessToTransition "M16", "CSSDRequestPro", "T003", "T007"
    GrantAccessToTransition "M16", "CSSDRequestPro", "T004", "T007"
    
    Exit Sub
    AddProcessCall "CSSDReceive", "ReloadOpener", "P030"
    AddProcessCall "CSSDReceive", "ReloadOpener", "P033"
    
    Exit Sub
    AddProcessCall "CSSDRequestPro", "InitCSSDRequestPro", "P010"
    AddProcessCall "CSSDRequestPro", "InitCSSDRequestPro", "P011"
    
    AddProcessCall "CSSDRequestPro", "ReloadOpener", "P010"
    AddProcessCall "CSSDRequestPro", "ReloadOpener", "P011"
    
    Exit Sub
    
    GrantAccessToTransition "M16", "CSSDRequestPro", "T002", "T003"
    GrantAccessToTransition "M16", "CSSDRequestPro", "T003", "T004"
    GrantAccessToTransition "M16", "CSSDRequestPro", "T004", "T005"
    GrantAccessToTransition "M16", "CSSDRequestPro", "T005", "T006"
    Exit Sub
    RepAnyPrintOutAlloc "CSSDRequestRCP", "M16"
    Exit Sub
    AddTransitionProcess "CSSDRequestPro", "T004", "Decontamination"
    UpdTransitionProcessName "CSSDRequestPro", "T004", "Decontamination"
    
    AddTransitionProcess "CSSDRequestPro", "T005", "Testing of Instruments"
    UpdTransitionProcessName "CSSDRequestPro", "T005", "Testing of Instruments"
    
    AddTransitionProcess "CSSDRequestPro", "T006", "Assembling and packaging"
    UpdTransitionProcessName "CSSDRequestPro", "T006", "Assembling and packaging"
    
    Exit Sub
    ConvertNavToMenu3 "M26"
    Exit Sub
    UpdPrintLayout "OPDDietCharges", "PrintInputFilter", "Date||Period"
    UpdPrintLayout "OPDDietCharges", "PrintProp", "100%||0||0"
    Exit Sub
    ConvertNavToMenu3 "m0318"
    Exit Sub
    ConvertNavToMenu3 "S01"
    Exit Sub
    ConvertNavToMenu3 "m0310"
    Exit Sub
    ConvertNavToMenu3 "s40"
    Exit Sub
    AddProcessCall "FoodMenu", "ReloadOpener", "P005"
    AddProcessCall "FoodMenu", "ReloadOpener", "P008"
    Exit Sub
    AddTransitionProcess "FoodOrder", "T001", "Pending Dieticians's Approval"
    AddTransitionProcess "FoodOrder", "T002", "Approved By Dieticians"
    Exit Sub
    ConvertNavToMenu3 "M0318"
    Exit Sub
    RepAccessRightAlloc "S13", "frmInvestigationURT001", "S22", "frmInvestigationURT001"
    RepAccessRightAlloc "S22", "frmInvestigationURT002", "S22", "frmInvestigationURT002"
    RepAccessRightAlloc "S13", "frmInvestigation2URT001", "S22", "frmInvestigation2URT001"
    RepAccessRightAlloc "S22", "frmInvestigation2URT002", "S22", "frmInvestigation2URT002"
    Exit Sub
    ConvertNavToMenu3 "S01"
    ConvertNavToMenu3 "S22"
    Exit Sub
    sql = "select distinct TableID from TableField where TableFieldID='UnitOfMEasureID'"
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            UpdTableField rst.fields("TableID"), "UnitOfMeasureID", "VisibleField", "Yes||Yes"
            UpdTableField rst.fields("TableID"), "UnitOfMeasureID", "LabelName", "Unit Of Measure"
            UpdTableField rst.fields("TableID"), "UnitOfMeasureID", "DisplayName", "Unit Of Measure"
            rst.MoveNext
        Loop
    End If
    
    
    Exit Sub
    ConvertNavToMenu3 "m0301"
    
    Exit Sub
    UpdPrintLayout "AppointmentDashboard", "PrintInputFilter", "Date||Period**Key||SpecialistID**Key||AppointmentCatID**Key||AppointmentCatTypeID**Key||SystemUserID"
    Exit Sub
    
    CopyPrintOutAccess "VisitationRCP", "AppointmentDashboard"
    Exit Sub
    AddProcessCall "PrescriptionAction", "ReloadOpener", "P005"
    AddProcessCall "PrescriptionAction", "ReloadOpener", "P008"
    
    Exit Sub
    sql = "Select * from Items where ItemCategoryID='CSSDSET' "
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            ReplaceTableKeyUpd "Items", rst.fields("ItemID"), Replace(rst.fields("ItemID"), "CSSDSET", "CSSDSET-"), ""
            rst.MoveNext
        Loop
    End If
    rst.Close
    
    
    Exit Sub
    ConvertNavToMenu3 "W001"
    Exit Sub
    UpdTableField "DrugSaleItems", "MainItemInfo1", "VisibleField", "NO||No"
    Exit Sub
    conn.execute " update bed set bedpos=1, bedNoID='001', wardID='W006', WardSectionID='W006' where bedid='W006-001' "
    conn.execute " update bed set bedpos=2, bedNoID='002', wardID='W006', WardSectionID='W006' where bedid='W006-002' "
    
    Exit Sub
    ReplaceTableKeyUpd "Bed", "W005-003", "W006-001", ""
    ReplaceTableKeyUpd "Bed", "W005-004", "W006-002", ""
    Exit Sub
    conn.execute "update LabTest set TestStatusID='TST002'; "
    Exit Sub
    UpdTableField "DrugSaleItems2", "DispenseInfo1", "VisibleField", "No||No"
    UpdTableField "DrugSaleItems2", "DispenseInfo1", "SubTableFieldSource", "TreeView"
    Exit Sub
    DeleteTableKey "Visitation", "V1220901001"
    Exit Sub
    conn.execute "update ProcessCall set ProcessPos=999 where UserProcessID='GeneratePatientInvoice'"
    Exit Sub
    CopyPrintOutAccess "VisitationRCP", "PatientReceiptSelector"
    AddProcesscallSibling "UpdatePatientBill", "GeneratePatientBill", "Receipt"
    conn.execute "update ProcessCall set ProcessPos=98 where TableID='Receipt' and UserProcessID='GeneratePatientBill'"
    
    UpdTableField "DrugSaleItems", "MainItemInfo1", "VisibleField", "Yes||Yes"
    UpdTableField "DrugSaleItems", "MainItemInfo1", "SubTableFieldSource", "UserText**16"
    UpdTableField "DrugSaleItems", "DrugGroupID", "VisibleField", "No||No"
    UpdTableField "DrugSaleItems", "DrugGroupID", "SubTableFieldSource", "TreeView"
    UpdTableField "DrugSaleItems", "DrugGroupID", "LabelName", "DrugGroupID"
    UpdTableField "DrugSaleItems", "DrugGroupID", "DisplayName", "DrugGroupID"
    
    
    conn.execute "update SystemSchedule set ScheduleInterval=21600 where SystemScheduleID='updateBillPeriod' "
    conn.execute "update BillPeriod set BillPeriodStatusID='B001' where BillMonthID='" & FormatWorkingMonth(Now) & "' "
    AddProcesscallSibling "GeneratePatientBill", "GeneratePatientInvoice", ""
    AddProcesscallSibling "GeneratePatientBill", "ReloadOpener", ""
'    Exit Sub
    
    conn.execute "DELETE FROM ""SystemVariables"" WHERE ""SystemVariableID""=N'AutoGenerateInvoiceForConsultBills';"
    conn.execute "INSERT INTO ""SystemVariables"" (""SystemVariableID"", ""SystemVariableName"", ""DataTypeID"", ""SystemVariableValue"", ""Description"", ""KeyPrefix"") VALUES ('AutoGenerateInvoiceForConsultBills', 'AutoGenerateInvoiceForConsultBills', 'DTY002', 'Yes', NULL, NULL);"
    conn.execute "DELETE FROM ""SystemVariables"" WHERE ""SystemVariableID""=N'AutoGenerateInvoiceForDrugBills';"
    conn.execute "INSERT INTO ""SystemVariables"" (""SystemVariableID"", ""SystemVariableName"", ""DataTypeID"", ""SystemVariableValue"", ""Description"", ""KeyPrefix"") VALUES ('AutoGenerateInvoiceForDrugBills', 'AutoGenerateInvoiceForDrugBills', 'DTY002', 'No', NULL, NULL);"
    conn.execute "DELETE FROM ""SystemVariables"" WHERE ""SystemVariableID""=N'AutoGenerateInvoiceForLabBills';"
    conn.execute "INSERT INTO ""SystemVariables"" (""SystemVariableID"", ""SystemVariableName"", ""DataTypeID"", ""SystemVariableValue"", ""Description"", ""KeyPrefix"") VALUES ('AutoGenerateInvoiceForLabBills', 'AutoGenerateInvoiceForLabBills', 'DTY002', 'No', NULL, NULL);"
    conn.execute "DELETE FROM ""SystemVariables"" WHERE ""SystemVariableID""=N'AutoGenerateInvoiceForStockBills';"
    conn.execute "INSERT INTO ""SystemVariables"" (""SystemVariableID"", ""SystemVariableName"", ""DataTypeID"", ""SystemVariableValue"", ""Description"", ""KeyPrefix"") VALUES ('AutoGenerateInvoiceForStockBills', 'AutoGenerateInvoiceForStockBills', 'DTY002', 'Yes', NULL, NULL);"
    conn.execute "DELETE FROM ""SystemVariables"" WHERE ""SystemVariableID""=N'AutoGenerateInvoiceForTreatBills';"
    conn.execute "INSERT INTO ""SystemVariables"" (""SystemVariableID"", ""SystemVariableName"", ""DataTypeID"", ""SystemVariableValue"", ""Description"", ""KeyPrefix"") VALUES ('AutoGenerateInvoiceForTreatBills', 'AutoGenerateInvoiceForTreatBills', 'DTY002', 'Yes', NULL, NULL);"

    
'    conn.execute "DELETE FROM ""ProcessCall"" WHERE ""ProcessCallID""=N'ConsultReview_GeneratePatientInvoice_P025';"
'    conn.execute "INSERT INTO ""ProcessCall"" (""ProcessCallID"", ""ProcessCallName"", ""TableID"", ""ProcessPointID"", ""ProcessPos"", ""UserProcessID"", ""UserAccessibleID"", ""Description"", ""KeyPrefix"") VALUES ('ConsultReview_GeneratePatientInvoice_P025', 'ConsultReview_GeneratePatientInvoice_P025', 'ConsultReview', 'P025', 9999, 'GeneratePatientInvoice', 'USA001', NULL, NULL);"
'    conn.execute "DELETE FROM ""ProcessCall"" WHERE ""ProcessCallID""=N'ConsultReview_GeneratePatientInvoice_P028';"
'    conn.execute "INSERT INTO ""ProcessCall"" (""ProcessCallID"", ""ProcessCallName"", ""TableID"", ""ProcessPointID"", ""ProcessPos"", ""UserProcessID"", ""UserAccessibleID"", ""Description"", ""KeyPrefix"") VALUES ('ConsultReview_GeneratePatientInvoice_P028', 'ConsultReview_GeneratePatientInvoice_P028', 'ConsultReview', 'P028', 9999, 'GeneratePatientInvoice', 'USA001', NULL, NULL);"
'    conn.execute "DELETE FROM ""ProcessCall"" WHERE ""ProcessCallID""=N'DrugSale_GeneratePatientInvoice_P025';"
'    conn.execute "INSERT INTO ""ProcessCall"" (""ProcessCallID"", ""ProcessCallName"", ""TableID"", ""ProcessPointID"", ""ProcessPos"", ""UserProcessID"", ""UserAccessibleID"", ""Description"", ""KeyPrefix"") VALUES ('DrugSale_GeneratePatientInvoice_P025', 'DrugSale_GeneratePatientInvoice_P025', 'DrugSale', 'P025', 9999, 'GeneratePatientInvoice', 'USA001', NULL, NULL);"
'    conn.execute "DELETE FROM ""ProcessCall"" WHERE ""ProcessCallID""=N'DrugSale_GeneratePatientInvoice_P028';"
'    conn.execute "INSERT INTO ""ProcessCall"" (""ProcessCallID"", ""ProcessCallName"", ""TableID"", ""ProcessPointID"", ""ProcessPos"", ""UserProcessID"", ""UserAccessibleID"", ""Description"", ""KeyPrefix"") VALUES ('DrugSale_GeneratePatientInvoice_P028', 'DrugSale_GeneratePatientInvoice_P028', 'DrugSale', 'P028', 9999, 'GeneratePatientInvoice', 'USA001', NULL, NULL);"
'    conn.execute "DELETE FROM ""ProcessCall"" WHERE ""ProcessCallID""=N'GeneratePatientInvoiceLabRequest1';"
'    conn.execute "INSERT INTO ""ProcessCall"" (""ProcessCallID"", ""ProcessCallName"", ""TableID"", ""ProcessPointID"", ""ProcessPos"", ""UserProcessID"", ""UserAccessibleID"", ""Description"", ""KeyPrefix"") VALUES ('GeneratePatientInvoiceLabRequest1', 'GeneratePatientInvoiceLabRequest1', 'LabRequest', 'P025', 999, 'GeneratePatientInvoice', 'USA001', NULL, NULL);"
'    conn.execute "DELETE FROM ""ProcessCall"" WHERE ""ProcessCallID""=N'GeneratePatientInvoiceLabRequest2';"
'    conn.execute "INSERT INTO ""ProcessCall"" (""ProcessCallID"", ""ProcessCallName"", ""TableID"", ""ProcessPointID"", ""ProcessPos"", ""UserProcessID"", ""UserAccessibleID"", ""Description"", ""KeyPrefix"") VALUES ('GeneratePatientInvoiceLabRequest2', 'GeneratePatientInvoiceLabRequest2', 'LabRequest', 'P028', 999, 'GeneratePatientInvoice', 'USA001', NULL, NULL);"
'    conn.execute "DELETE FROM ""ProcessCall"" WHERE ""ProcessCallID""=N'GeneratePatientInvoiceVisitation1';"
'    conn.execute "INSERT INTO ""ProcessCall"" (""ProcessCallID"", ""ProcessCallName"", ""TableID"", ""ProcessPointID"", ""ProcessPos"", ""UserProcessID"", ""UserAccessibleID"", ""Description"", ""KeyPrefix"") VALUES ('GeneratePatientInvoiceVisitation1', 'GeneratePatientInvoiceVisitation1', 'Visitation', 'P005', 999, 'GeneratePatientInvoice', 'USA001', NULL, NULL);"
'    conn.execute "DELETE FROM ""ProcessCall"" WHERE ""ProcessCallID""=N'GeneratePatientInvoiceVisitation2';"
'    conn.execute "INSERT INTO ""ProcessCall"" (""ProcessCallID"", ""ProcessCallName"", ""TableID"", ""ProcessPointID"", ""ProcessPos"", ""UserProcessID"", ""UserAccessibleID"", ""Description"", ""KeyPrefix"") VALUES ('GeneratePatientInvoiceVisitation2', 'GeneratePatientInvoiceVisitation1', 'Visitation', 'P008', 999, 'GeneratePatientInvoice', 'USA001', NULL, NULL);"

    RepAnyAccessRightAlloc "W001", "frmStockIssueURT001"
    RepAnyAccessRightAlloc "W001", "frmStockIssueURT002"
    RepAnyAccessRightAlloc "W001", "frmStockIssueURT003"
    RepAnyAccessRightAlloc "W001", "frmStockIssueItemsURT001"
    'RepAccessRightAlloc "W011", "frmStockIssueItemsURT001", "ClaimManager", "frmStockIssueItemsURT001"
    
    AddTransitionProcess "AdmissionPro", "T006", "Return Patient to ward"
    AddTransitionProcess "AdmissionPro", "T004", "Complete For Printing"
    AddTransitionProcess "AdmissionPro", "T005", "Discharge Patient"
    UpdTransitionProcessName "AdmissionPRo", "T004", "Complete For Printing"
    
    GrantAccessToTransition "M13", "AdmissionPro", "T003", "T006"
    
    GrantAccessToTransition "W001", "AdmissionPro", "T006", "T003"
    GrantAccessToTransition "W002", "AdmissionPro", "T006", "T003"
    GrantAccessToTransition "W003", "AdmissionPro", "T006", "T003"
    GrantAccessToTransition "W004", "AdmissionPro", "T006", "T003"
    GrantAccessToTransition "W005", "AdmissionPro", "T006", "T003"
    
    'Exit Sub
    
    conn.execute "DELETE FROM ConsultReview WHERE LEN(consultreviewid)=0"
    conn.execute "update PatientMode Set KeyPrefix='/' where PatientModeID='P001' "
    conn.execute "update SystemVariables Set SystemVariableValue='MR/' where SystemVariableID='KeyPrefixPatientPatientID' "
    conn.execute "delete from KeyAllocation " 'where TableID='Patient' "
    conn.execute "delete from KeyCurrCount " 'where TableID='Patient' "
    conn.execute "update insuredPatient set KeyPrefix='-' where len(KeyPrefix)=0"
    conn.execute "update LabTest set TestStatusID='TST002'; "
    conn.execute "truncate table ConsultCostMatrix;"
    conn.execute "truncate table LabTestCostMatrix;"
    conn.execute "truncate table DrugPriceMatrix;"
    conn.execute "truncate table ItemPriceMatrix;"
    
    Exit Sub
    CopyPrintOutAccess "VisitationRCP", "PatientReceiptHistory"
    Exit Sub
    CopyPrintOutAccess "TreatmentSheet", "TreatmentSheet2"
    Exit Sub
    RecompileMenus ""
    Exit Sub
    RepAnyPrintOutAlloc "PatientSurgeryInformation", "MedicalRecords"
    RepAnyPrintOutAlloc "PatientSurgeryInformation", "S01"
    
    Exit Sub
    RepAccessRightAlloc "S01", "frmPatientURT001", "MedicalRecords", "frmPatientURT001"
    RepAccessRightAlloc "S01", "frmPatientURT002", "MedicalRecords", "frmPatientURT001"
    RepAccessRightAlloc "S01", "frmPatientURT003", "MedicalRecords", "frmPatientURT001"
    Exit Sub
    'Dim dy
    dy = Day(Now())
    'If dy > 20 Then
      UpdateBillPeriod Now()
    'End If
    
    Exit Sub
    
    AddConsultations
    
    Exit Sub
    
    sql = "select PatientID from Patient"
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            ReplaceTableKeyUpd "Patient", rst.fields("PatientID"), Replace(rst.fields("PatientID"), "-", "/"), ""
            rst.MoveNext
        Loop
    End If
    rst.Close
    
    Exit Sub
    UpdPrintLayout "SelectPatientQuick", "PrintProp", "100%||0||0"
    Exit Sub
    CopyPrintOutAccess "SelectPatient", "SelectPatientQuick"
    
    Exit Sub
    CopyDrugPriceToIns "I102", "I109"
    CopyBedPriceToIns "I102", "I109"
    CopyInvPriceToIns "I102", "I109"
    CopyItemPriceToIns "I102", "I109"
    CopyTreatPriceToIns "I102", "I109"
    
    Exit Sub
    DeleteTableKey "SystemSchedule", "MigrateRMCData01"
    DeleteTableKey "SystemSchedule", "MigrateRMCData02"
    DeleteTableKey "SystemSchedule", "MigrateRMCData03"
    DeleteTableKey "SystemSchedule", "MigrateRMCData04"
    DeleteTableKey "SystemSchedule", "MigrateRMCData05"
    DeleteTableKey "SystemSchedule", "MigrateRMCData06"
    DeleteTableKey "SystemSchedule", "MigrateRMCData07"
    DeleteTableKey "SystemSchedule", "MigrateRMCData08"
    DeleteTableKey "SystemSchedule", "MigrateRMCData09"
    DeleteTableKey "SystemSchedule", "MigrateRMCData10"
    DeleteTableKey "SystemSchedule", "MigrateRMCData11"
    DeleteTableKey "SystemSchedule", "MigrateRMCData12"
    DeleteTableKey "SystemSchedule", "MigrateRMCData13"
    DeleteTableKey "SystemSchedule", "MigrateRMCData14"
    DeleteTableKey "SystemSchedule", "MigrateRMCData15"
    
    Exit Sub
    ReplaceTableKeyUpd "InsuranceType", "I002", "I109", "ALISA HOTEL"
    
    Exit Sub
    DeleteTableKey "ItemRequest2", "IR000037"
    Exit Sub
    AddAppointTimes
    Exit Sub
    
    GrantAccessToTransition "DPT014", "ReceiptPro", "T002", "T002" 'refund, refund
    
    GrantAccessToTransition "DPT014", "ReceiptPro", "T001", "T003" ' new to cancel
    GrantAccessToTransition "DPT014", "ReceiptPro", "T004", "T003" 'change owner to cancel
    GrantAccessToTransition "DPT014", "ReceiptPro", "T005", "T003" 'change info to cancel
    
    GrantAccessToTransition "DPT014", "ReceiptPro", "T001", "T004" 'new tp change owner
    GrantAccessToTransition "DPT014", "ReceiptPro", "T005", "T004" 'change info to change owner
    
    GrantAccessToTransition "DPT014", "ReceiptPro", "T001", "T005" 'new to change info
    GrantAccessToTransition "DPT014", "ReceiptPro", "T003", "T005" 'refund to change info
    GrantAccessToTransition "DPT014", "ReceiptPro", "T005", "T005" 'change info to change info
    GrantAccessToTransition "DPT014", "ReceiptPro", "T005", "T002" 'change info, refund
    GrantAccessToTransition "DPT014", "ReceiptPro", "T004", "T004" 'change owner, change owner
    GrantAccessToTransition "DPT014", "ReceiptPro", "T004", "T002" 'change owner, change owner
    
    GrantAccessToTransition "DPT014", "ReceiptPro", "T001", "T002"
    GrantAccessToTransition "DPT014", "VisitationPro", "T001", "T002"
    GrantAccessToTransition "DPT014", "VisitationPro", "T001", "T003"
    GrantAccessToTransition "DPT014", "VisitationPro", "T002", "T003"
    
    GrantAccessToTransition "DPT014", "VisitationPro", "T001", "T002"
    GrantAccessToTransition "DPT014", "VisitationPro", "T001", "T003"
    GrantAccessToTransition "DPT014", "VisitationPro", "T002", "T003"
    
    Exit Sub
    MigrateConsumables
    Exit Sub
    
    MergeJobSchedule "DPT014", ""
    RecompileMenus "DPT014"
    
    MergeJobSchedule "DPT014", "DPT014||S01||S13||S19||Cashier||ChiefCashier||CreditControl||ClaimManager||Transactions||S02||S22||S21A||DPT001||S17"
    RecompileMenus "DPT014"
    
    Exit Sub
    SetUpBillGroups
    
    Exit Sub
    FixVisitBillForSponsor "S048"
    
    Exit Sub
    
    MergeJobSchedule "DPT014", ""
    CompileMergedNav "DPT014"
    ConvertNavToMenu3 "DPT014"
    
    MergeJobSchedule "DPT014", "S01||S13||S19||Cashier|ChiefCashier||CreditControl||ClaimManager||Transactions||S02||S22||S21A||DPT001||S17"
    CompileMergedNav "DPT014"
    ConvertNavToMenu3 "DPT014"
    
    Exit Sub
    SetUpBillGroups
    Exit Sub
    RepAccessRightAlloc "S22", "frmDrugSaleItemsURT001", "ClaimManager", "frmDrugSaleItemsURT001"
    RepAccessRightAlloc "S22", "frmDrugSaleItems2URT001", "ClaimManager", "frmDrugSaleItems2URT001"
    RepAccessRightAlloc "W011", "frmStockIssueItemsURT001", "ClaimManager", "frmStockIssueItemsURT001"
    RepAccessRightAlloc "W011", "frmStockIssueURT001", "ClaimManager", "frmStockIssueURT001"
    'RepAccessRightAlloc "W011", "frmTreatChargesURT001", "ClaimManager", "frmStockIssueItemsURT001"


    Exit Sub
    SetUpBillGroups
    Exit Sub
    RepAnyPrintOutAlloc "VisitationRCP", "ClaimManager"
    Exit Sub
    SetUpBillGroups
    Exit Sub
    RecompileMenus "M13"
    RecompileMenus "S13"
    Exit Sub
    mapRXDrug2
    Exit Sub
    Call SetUpBillGroups
    
    Exit Sub
    
    UpdPrintLayout "SponsorBill2", "TableID", "Sponsor"
    UpdPrintLayout "SponsorBill2B", "TableID", "Sponsor"
    UpdPrintLayout "SponsorBill2B", "PrintInputFilter", GetComboNameFld("Printlayout", "SponsorBill2", "PrintInputFilter")
    UpdPrintLayout "SponsorBill2B", "PrintProp", GetComboNameFld("Printlayout", "SponsorBill2", "PrintProp")
    
    RepAnyPrintOutAlloc "SponsorBill2", "M13"
    RepAnyPrintOutAlloc "SponsorBill2B", "M13"
    
    Exit Sub
    fixVisitBillsNationwide
    Exit Sub
    FixInsuredPatient
    FixVisitBills
    Exit Sub
    fixVisitBillsNationwide
    Exit Sub
    FixVisitBills
    Exit Sub
    RepAnyPrintOutAlloc "ApproveClaim", "m13"

    Exit Sub
    GrantAccessToTransition "M13", "VisitationPro", "T001", "T002"
    GrantAccessToTransition "M13", "VisitationPro", "T001", "T003"
    GrantAccessToTransition "M13", "VisitationPro", "T002", "T003"
    
    GrantAccessToTransition "ClaimManager", "VisitationPro", "T001", "T002"
    GrantAccessToTransition "ClaimManager", "VisitationPro", "T001", "T003"
    GrantAccessToTransition "ClaimManager", "VisitationPro", "T002", "T003"
    
    'FixInsuredPatient
    'FixDrugMat
    Exit Sub
    
    RecompileMenus "Dpt014"
    RecompileMenus "Dpt014"
    RepBrowseView "DrugIssueByTransProcessStat", "DrugIssueByWorkingMonthByDrugStorePedingApprov"
    UpdBrowseView "DrugIssueByWorkingMonthByDrugStorePedingApprov", "BrowseViewName", "Medical Item Issue Month pending Approval "
    UpdBrowseView "DrugIssueByWorkingMonthByDrugStorePedingApprov", "UserAccessibleID", "USA001"
    UpdBrowseView "DrugIssueByWorkingMonthByDrugStorePedingApprov", "ReportGroupByID", "WorkingMonth"
    
    
    UpdTransitionProcessName "DrugIssuePro", "T001", "Initial Item Issue"
    UpdTransitionProcessName "DrugIssuePro", "T002", "Approved By Auditor"
    GrantAccessToTransition "DPT014", "DrugIssuePro", "T001", "T002"
    
    Exit Sub
    
    
    RepAnyAccessRightAlloc "m13", "frmInsuredPatientURT001"
    RepAnyAccessRightAlloc "m13", "frmInsuredPatientURT002"
    RepAnyAccessRightAlloc "m13", "frmInsuredPatientURT003"
    
    Exit Sub
    DeleteTableKey "ItemRequest2", "IR000030"
    
    Exit Sub
    RepAnyAccessRightAlloc "m13", "frmItemStockLevelURT001"
    Exit Sub
    GrantAccessToTransition "DPT014", "ItemIssuePro", "T001", "T002"
    Exit Sub
    AddProcessCall "ItemIssuePro", "UpdateItemIssuePro", "P005"
    
    Exit Sub
    RecompileMenus "DPT014"
    
    Exit Sub
 
    RepBrowseView "ItemIssueByTransProcessStat", "ItemIssueByWorkingMonthByItemStorePedingApprov"
    UpdBrowseView "ItemIssueByWorkingMonthByItemStorePedingApprov", "BrowseViewName", "Item Issue Month pending Approval "
    UpdBrowseView "ItemIssueByWorkingMonthByItemStorePedingApprov", "UserAccessibleID", "USA001"
    UpdBrowseView "ItemIssueByWorkingMonthByItemStorePedingApprov", "ReportGroupByID", "WorkingMonth"
'
'    RepBrowseView "ItemIssueByTransProcessStat", "ItemIssueByWorkingMonth_TransProcessVal"
'    UpdBrowseView "ItemIssueByWorkingMonth_TransProcessVal", "BrowseViewName", "Item Issue Month by Stages "
'    UpdBrowseView "ItemIssueByWorkingMonth_TransProcessVal", "UserAccessibleID", "USA001"
'    UpdBrowseView "ItemIssueByWorkingMonth_TransProcessVal", "ReportGroupByID", "WorkingMonth_TransProcessVal"

    
    
    RepBrowseView "ItemIssueByTransProcessStat", "ItemIssueByTransProcessVal"
    UpdBrowseView "ItemIssueByTransProcessVal", "BrowseViewName", "Item Issue Stages"
    UpdBrowseView "ItemIssueByTransProcessVal", "UserAccessibleID", "USA001"
    
    
    Exit Sub
    sql = "select DrugRequest2ID from DrugRequest2 where RequestDate<cast('01 Jul 2022 00:00:00' as date)"
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            DeleteTableKey "DrugRequest2", rst.fields("DrugRequest2ID")
            rst.MoveNext
        Loop
        
    End If
    rst.Close
    Set rst = Nothing
    
    
    Exit Sub
    
    
    Exit Sub
    sql = "select ItemRequest2ID from ItemRequest2 where RequestDate<cast('01 Jul 2022 00:00:00' as date)"
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            DeleteTableKey "ItemRequest2", rst.fields("ItemRequest2ID")
            rst.MoveNext
        Loop
        
    End If
    rst.Close
    Set rst = Nothing
    
    
    Exit Sub
    
    UpdTransitionProcessName "ItemIssuePro", "T001", "Initial Item Issue"
    UpdTransitionProcessName "ItemIssuePro", "T002", "Approved By Auditor"
    RepAnyAccessRightAlloc "m13", "frmItemAcceptItemsURT001"
    RepAnyAccessRightAlloc "m13", "frmItemIssueItemsURT001"

    Exit Sub
    MsgBox HasPostTransaction("ItemIssue")
    Exit Sub
    MsgBox GetComboName("DrugStore2", "Transactions**M13")
    Exit Sub
    RecompileMenus "M13"
    RecompileMenus "CreditControl"
    Exit Sub
    RecompileMenus "Cashier"
    RecompileMenus "ClaimManager"
    RecompileMenus "CreditControl"
    RecompileMenus "DPT014"
    RecompileMenus "ITSupport"
    RecompileMenus "M13"
    
    Exit Sub
    
    ResetUserPwd "Bluefaces"
    ResetUserPwd3 "Bluefaces"
    
    Exit Sub
    RecompileMenus "s01"
    RepAnyAccessRightAlloc "W011", "frmStockIssueItemsURT001"
    Exit Sub
    AddTransitionProcess "AdmissionPro", "T006", "Return Patient to ward"
    
    GrantAccessToTransition "S17", "AdmissionPro", "T003", "T006"
    GrantAccessToTransition "S17A", "AdmissionPro", "T003", "T006"
    GrantAccessToTransition "M13", "AdmissionPro", "T003", "T006"
    
    GrantAccessToTransition "W011", "AdmissionPro", "T006", "T003"
    
    Exit Sub

    GrantAccessToTransition "M13", "AdmissionPro", "T003", "T004"
    GrantAccessToTransition "M13", "AdmissionPro", "T004", "T005"
    GrantAccessToTransition "M13", "AdmissionPro", "T003", "T004"
    
    GrantAccessToTransition "M13", "AdmissionPro", "T003", "T004"
    GrantAccessToTransition "M13", "AdmissionPro", "T004", "T005"
    GrantAccessToTransition "M13", "AdmissionPro", "T003", "T004"
    
    RepAnyPrintOutAlloc "PatientReceiptHistory", "m13"
    
    RepAnyPrintOutAlloc "markAsRXSponsor", "m13"
    RepAnyPrintOutAlloc "RxClaimPushALT", "m13"
    RepAnyPrintOutAlloc "rxunmappedItems", "m13"
    
    RepAnyAccessRightAlloc "m13", "frmEMRVar99URT001"
    RepAnyAccessRightAlloc "m13", "frmEMRVar99URT002"
    RepAnyAccessRightAlloc "m13", "frmEMRVar99URT003"
    RepAnyAccessRightAlloc "m13", "frmEMRVar99URT004"
    
    RepAnyAccessRightAlloc "m13", "frmPerformVar11URT001"
    RepAnyAccessRightAlloc "m13", "frmPerformVar11URT002"
    RepAnyAccessRightAlloc "m13", "frmPerformVar11URT003"
    RepAnyAccessRightAlloc "m13", "frmPerformVar11URT004"
    
    Exit Sub
    MsgBox GetRecordKey("MaritalStatus", "MaritalStatusID", "NONE")
    MsgBox GetRecordKey("InsSchemeMode", "InsSchemeModeID", "NONE")
    Exit Sub
    RepAnyAccessRightAlloc "S01", "frmSponsorURT001"
    RepAnyAccessRightAlloc "S01", "frmSponsorURT002"
    RepAnyAccessRightAlloc "S01", "frmSponsorURT003"
    RepAnyAccessRightAlloc "S01", "frmSponsorURT004"
    
    RepAnyAccessRightAlloc "S01", "frmInsuranceSchemeURT001"
    RepAnyAccessRightAlloc "S01", "frmInsuranceSchemeURT002"
    RepAnyAccessRightAlloc "S01", "frmInsuranceSchemeURT003"
    RepAnyAccessRightAlloc "S01", "frmInsuranceSchemeURT004"
    
    RecompileMenus "S01"
    RecompileMenus "Cashier"
    
    Exit Sub
    RepAnyAccessRightAlloc "S01", "frmInsSchemeModeURT001"
    RepAnyAccessRightAlloc "S01", "frmInsSchemeModeURT002"
    RepAnyAccessRightAlloc "S01", "frmInsSchemeModeURT003"
    RepAnyAccessRightAlloc "S01", "frmInsSchemeModeURT004"
    
    RepAnyAccessRightAlloc "S01", "frmBenefitOptionURT001"
    RepAnyAccessRightAlloc "S01", "frmBenefitOptionURT002"
    RepAnyAccessRightAlloc "S01", "frmBenefitOptionURT003"
    RepAnyAccessRightAlloc "S01", "frmBenefitOptionURT004"
    
    RepAnyAccessRightAlloc "S01", "frmBenefitCompanyURT001"
    RepAnyAccessRightAlloc "S01", "frmBenefitCompanyURT002"
    RepAnyAccessRightAlloc "S01", "frmBenefitCompanyURT003"
    RepAnyAccessRightAlloc "S01", "frmBenefitCompanyURT004"
    
    
    Exit Sub
    
    RecompileMenus "M032"
    Exit Sub
'    MsgBox HasPostTransaction("LabRequest")
'    MsgBox HasPostTransaction("DrugSaleItems2")
'    MsgBox HasPostTransaction("StockIssueItems")
'    MsgBox HasPostTransaction("DrugAcceptItems")
'    MsgBox HasPostTransaction("ItemAcceptItems")
'    MsgBox HasPostTransaction("DrugAdjustItems")
'    MsgBox HasPostTransaction("StockAdjustItems")
'    MsgBox HasPostTransaction("DrugReturnItems")
'    MsgBox HasPostTransaction("DrugReturnItems2")
'    MsgBox HasPostTransaction("StockReturnItems")
'    MsgBox HasPostTransaction("DrugToSupplierItem")
'    MsgBox HasPostTransaction("StockToSupplierItm")
'    MsgBox HasPostTransaction("IncomingDrugItems")
'    MsgBox HasPostTransaction("IncomingDrugItems2")
'    MsgBox HasPostTransaction("IncomingStockItems")
    
'    Exit Sub

'    sql = "select * from DrugPurOrder where PurchaseOrderDate < '01 Jul 2022' "
'    Set rst = CreateObject("ADODB.RecordSet")
'    rst.open sql, conn, 3, 4
'    If rst.RecordCount > 0 Then
'        rst.MoveFirst
'        Do While Not rst.EOF
'            DeleteTableKey "DrugPurOrder", rst.fields("DrugPurOrderID")
'            rst.MoveNext
'        Loop
'    End If
'    Set rst = Nothing
'
'    Exit Sub
    
    UpdPrintLayout "markAsRXSponsor", "PrintINputFilter", ""
    RepAnyPrintOutAlloc "markAsRXSponsor", "ClaimManager"
    RepAnyPrintOutAlloc "RxClaimPushALT", "ClaimManager"
    
    UpdPrintLayout "rxclaimPushAlt", "PrintINputFilter", "Key||WorkingMonthID**Key||VisitationID**Text||Re-Push (YES/NO)**Text||Limit**Text||Start From"
    UpdPrintLayout "rxclaimsubmit", "PrintInputFilter", "Key||WorkingMonthID**Key||VisitationID**Text||Re-Push (YES/NO)**Text||Limit**Text"
    RepAnyPrintOutAlloc "rxunmappedItems", "Claimmanager"
    RepAnyPrintOutAlloc "rxclaimsubmit", "Claimmanager"
    UpdPrintLayout "rxunmappeditems", "PrintInputFilter", "Key||WorkingMonthID"
    
    RepAnyAccessRightAlloc "Claimmanager", "frmEMRVar99URT001"
    RepAnyAccessRightAlloc "Claimmanager", "frmEMRVar99URT002"
    RepAnyAccessRightAlloc "Claimmanager", "frmEMRVar99URT003"
    RepAnyAccessRightAlloc "Claimmanager", "frmEMRVar99URT004"
    
    RepAnyAccessRightAlloc "Claimmanager", "frmPerformVar11URT001"
    RepAnyAccessRightAlloc "Claimmanager", "frmPerformVar11URT002"
    RepAnyAccessRightAlloc "Claimmanager", "frmPerformVar11URT003"
    RepAnyAccessRightAlloc "Claimmanager", "frmPerformVar11URT004"
    
    AddProcessCall "EMRVar99", "initEMRVar99", "P010"
    AddProcessCall "EMRVar99", "initEMRVar99", "P011"
    
    AddProcessCall "PerformVar11", "initPerformVar11", "P010"
    AddProcessCall "PerformVar11", "initPerformVar11", "P011"
    
    
    Exit Sub
    RepAnyPrintOutAlloc "PatientRCP", "M13"
    RepAnyPrintOutAlloc "PatientRCP", "Claimmanager"
    RepAnyPrintOutAlloc "Admission1", "M13"
    RepAnyPrintOutAlloc "Admission1", "Claimmanager"
    
    Exit Sub
    RecompileMenus "Claimmanager"
    RecompileMenus "M13"
    Exit Sub
    
    AddTransitionProcess "AdmissionPro", "T004", "Complete For Printing"
    AddTransitionProcess "AdmissionPro", "T005", "Discharge Patient"
    UpdTransitionProcessName "AdmissionPRo", "T004", "Complete For Printing"
    
    ReplicateAccess "Admission", "S17A"
    ReplicateAccess "Admission", "S17"
    
    GrantAccessToTransition "S17", "AdmissionPro", "T003", "T004"
    GrantAccessToTransition "S17", "AdmissionPro", "T004", "T005"
    GrantAccessToTransition "S17", "AdmissionPro", "T003", "T004"
    
    GrantAccessToTransition "S17A", "AdmissionPro", "T003", "T004"
    GrantAccessToTransition "S17A", "AdmissionPro", "T004", "T005"
    GrantAccessToTransition "S17A", "AdmissionPro", "T003", "T004"
    
    Exit Sub
    ReplicateAccess "S17", "S17A"
    Exit Sub
    
    DeleteTableKey "Patient", "RMC06004/22"
    DeleteTableKey "Patient", "RMC06005/22"
    DeleteTableKey "Patient", "RMC06003/22"
    DeleteTableKey "Patient", "RMC06001/22"
    
    DeleteTableKey "Visitation", "V1220630007"
    DeleteTableKey "Visitation", "V1220630008"
    
    Exit Sub
    
    'RecompileMenus ""
    
    CompileMergedNav "DPT014"
    ConvertNavToMenu3 "DPT014"
    
    Exit Sub
    
    ReplicateAccess "W001", "W010"
    ReplicateAccess "W001", "W001"
    ReplicateAccess "W001", "W002"
    ReplicateAccess "W001", "W003"
    ReplicateAccess "W001", "W004"
    ReplicateAccess "W001", "W005"
    ReplicateAccess "W001", "W006"
    ReplicateAccess "W001", "W007"
    ReplicateAccess "W001", "W008"
    ReplicateAccess "W001", "W009"
    ReplicateAccess "W001", "W011"
    
    Exit Sub
    ReplicateProfile "M0304", "M0326", "UROLOGY [DOCTOR]"
    ReplicateProfile "M0304", "M0327", "OPHTAMOLOGY [DOCTOR]"
    
    'Exit Sub
    CopyPrintOutAccess "SponsorBillI", "SponsorBill2"
    UpdPrintLayout "SponsorBill2", "PrintInputFilter", "Key||WorkingMonthID"
    UpdPrintLayout "SponsorBill2", "PrintProp", "100%||0||0"
    
    'Exit Sub
    UpdTableField "Disease", "DiseaseGroupID", "VisibleInput", "1|1|1"
    'Exit Sub
    AddProcessCall2 "ConsultReview", "GeneratePatientInvoice", "P028", 9999
    AddProcessCall2 "ConsultReview", "GeneratePatientInvoice", "P025", 9999
    
    'Exit Sub
    
    UpdTableField "DrugSaleItems", "MainItemInfo1", "VisibleField", "Yes||Yes"
    UpdTableField "DrugSaleItems", "MainItemInfo1", "SubTableFieldSource", "UserText**50"
    UpdTableField "DrugSaleItems", "MainItemInfo1", "LabelName", "Instructions/Comments"
    UpdTableField "DrugSaleItems", "MainItemInfo1", "DisplayName", "Instructions/Comments"
    
    UpdTableField "DrugSaleItems2", "MainInfo1", "VisibleField", "Yes||Yes"
    UpdTableField "DrugSaleItems2", "MainInfo1", "SubTableFieldSource", "UserText**50"
    UpdTableField "DrugSaleItems2", "MainInfo1", "LabelName", "Instructions/Comments"
    UpdTableField "DrugSaleItems2", "MainInfo1", "DisplayName", "Instructions/Comments"
    
    UpdTableField "DrugSaleItems2", "DispenseInfo1", "SubTableFieldSource", "UserText**50**-"
    UpdTableField "DrugSaleItems2", "DispenseInfo1", "LabelName", "DispenseInfo1"
    UpdTableField "DrugSaleItems2", "DispenseInfo1", "DisplayName", "DispenseInfo1"
    UpdTableField "DrugSaleItems2", "DispenseInfo1", "VisibleField", "No||No"
    
    AddProcessCall "Visitation", "ReloadOpener", "P005"
    AddProcessCall "Visitation", "ReloadOpener", "P008"
    
'    Exit Sub
    CopyPrintOutAccess "VisitationRCP", "BlockPrompt"
    
    UpdTableField "SpecialistGroup", "KeyPrefix", "VisibleInput", "1|0|0"
    UpdTableField "SpecialistGroup", "KeyPrefix", "LabelName", "Key Flags"
    UpdTableField "SpecialistGroup", "KeyPrefix", "DisplayName", "Key Flags"
    
    UpdPullupData "ReceiptPro", "Receipt", "WithEdit", "Client"
    
'    //ReceiptPro-T002 refund
'    //ReceiptPro-T003 cancel
'    //ReceiptPro-T004 change owner
'    //ReceiptPro-T005 change payment info

    UpdTransitionProcessName "ReceiptPro", "T003", "Cancel"
    UpdTransitionProcessName "ReceiptPro", "T004", "Change Owner"
    AddTransitionProcess "ReceiptPro", "T005", "Change Payment Info"
    
    GrantAccessToTransition "ChiefCashier", "ReceiptPro", "T002", "T002" 'refund, refund
    
    GrantAccessToTransition "ChiefCashier", "ReceiptPro", "T001", "T003" ' new to cancel
    GrantAccessToTransition "ChiefCashier", "ReceiptPro", "T004", "T003" 'change owner to cancel
    GrantAccessToTransition "ChiefCashier", "ReceiptPro", "T005", "T003" 'change info to cancel
    
    GrantAccessToTransition "ChiefCashier", "ReceiptPro", "T001", "T004" 'new tp change owner
    GrantAccessToTransition "ChiefCashier", "ReceiptPro", "T005", "T004" 'change info to change owner
    
    GrantAccessToTransition "ChiefCashier", "ReceiptPro", "T001", "T005" 'new to change info
    GrantAccessToTransition "ChiefCashier", "ReceiptPro", "T003", "T005" 'refund to change info
    GrantAccessToTransition "ChiefCashier", "ReceiptPro", "T005", "T005" 'change info to change info
    GrantAccessToTransition "ChiefCashier", "ReceiptPro", "T005", "T002" 'change info, refund
    GrantAccessToTransition "ChiefCashier", "ReceiptPro", "T004", "T004" 'change owner, change owner
    GrantAccessToTransition "ChiefCashier", "ReceiptPro", "T004", "T002" 'change owner, change owner
    
    ''!!extreme cases, remember to remove access!!!
'    GrantAccessToTransition "ChiefCashier", "ReceiptPro", "T002", "T003" ' refund to cancel
'    GrantAccessToTransition "ChiefCashier", "ReceiptPro", "T003", "T004" ' cancel to change owner
'    GrantAccessToTransition "ChiefCashier", "ReceiptPro", "T002", "T004" ' refund to change owner
'    GrantAccessToTransition "ChiefCashier", "ReceiptPro", "T002", "T005" ' refund to change info
    
    UpdTableField "ReceiptPro", "PaymentModeID", "DefaultValue", ""
    UpdTableField "Receipt", "PaymentModeID", "DefaultValue", ""
    
    Exit Sub
    'ApplyInventoryTrans "DrugSaleItems", " where DrugSaleID='D122060954' " 'post drugsale
    'RebuildInventoryTransactions
    RepAnyPrintOutAlloc "KenUpdatePack", "SystemAdmin"
    
    AddProcessCall2 "DrugSale", "GeneratePatientInvoice", "P028", 9999
    AddProcessCall2 "DrugSale", "GeneratePatientInvoice", "P025", 9999
    
    AddProcessCall2 "DrugSale", "ReloadOpener", "P028", 9999
    AddProcessCall2 "DrugSale", "ReloadOpener", "P025", 9999
    
    
    
    Exit Sub
    tmp = Split("", ", ")
        
    If UBound(tmp) >= 0 Then
        phoneNo = tmp(0)
        idx = 0
        Do While Left(phoneNo, 1) = "0" And idx <= 10
            phoneNo = Mid(phoneNo, 2, Len(phoneNo))
            idx = idx + 1
        Loop
        MsgBox phoneNo
    End If
    
    
    Exit Sub
    '06 Jun 2022
    
    AddProcessCall "Patient", "MoveToNextFlow", "P008"
    UpdProcessCall "Patient_MoveToNextFlow_P008", "ProcessPos", 100000
    
    AddProcessCall "Patient", "MoveToNextFlow", "P005"
    UpdProcessCall "Patient_MoveToNextFlow_P005", "ProcessPos", 100000

'''  AddProcessCall "Visitation", "MoveToNextFlow", "P008"
    AddProcessCall "Visitation", "MoveToNextFlow", "P005"
    UpdProcessCall "Visitation_MoveToNextFlow_P005", "ProcessPos", 100000

    AddProcessCall "InsuredPatient", "MoveToNextFlow", "P008"
    UpdProcessCall "InsuredPatient_MoveToNextFlow_P008", "ProcessPos", 100000
    
    AddProcessCall "InsuredPatient", "MoveToNextFlow", "P005"
    UpdProcessCall "InsuredPatient_MoveToNextFlow_P005", "ProcessPos", 100000
    
    UpdPrintLayout "SelectPatientQuick", "PrintProp", "100%||0||0"
    
    ConvertNavToMenu3 "s01"
    
    CopyPrintOutAccess "SelectPAtient", "SelectPatientQuick"
    RepAnyPrintOutAlloc "PatientReceiptHistory", "S17A"
    ReplicateProfile "S01", "S01A", "OPD / Records"
    
    source = "M0303"
    target = "M0322"
    ReplicateProfile source, target, "ORTHOPEDICS [Doctor]"
    
    source = "M0303"
    target = "M0323"
    ReplicateProfile source, target, "CARDIOLOGY [Doctor]"
    
    source = "M0303"
    target = "M0324"
    ReplicateProfile source, target, "DERMATOLOGY [Doctor]"
    
    
    
    
    
    
    Exit Sub
    ConvertNavToMenu3 "DPT020"
    ConvertNavToMenu3 "M0322"
    ConvertNavToMenu3 "M0323"
    ConvertNavToMenu3 "M0324"
    ConvertNavToMenu3 "M0325"
    
    Exit Sub
    UpdPrintLayout "Nurses24HReport", "PrintProp", "80%||0||0"
    
    Exit Sub
    UpdJobSchedule "DPT020", "JobScheduleName", "NURSING HEAD"
    
    Exit Sub
    RepAnyPrintOutAlloc "Admission1", "W001"
    RepAnyPrintOutAlloc "Admission1", "S17A"
    Exit Sub
    source = "M0303"
    target = "M0322"
    ReplicateProfile source, target, "ORTHOPEDICS [Doctor]"
    
    source = "M0303"
    target = "M0323"
    ReplicateProfile source, target, "CARDIOLOGY [Doctor]"
    
    source = "M0303"
    target = "M0324"
    ReplicateProfile source, target, "DERMATOLOGY [Doctor]"
    
    
    Exit Sub
    RepAnyPrintOutAlloc "ViewTestTemplate", "SystemAdmin"
    RepAnyPrintOutAlloc "ViewTestTemplate", "S13"
    
    'Exit Sub
    sql = "select systemUserID from SystemUser"
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            ResetUserPwd rst.fields("SystemUserID")
            rst.MoveNext
        Loop
    End If
    
    Exit Sub
    ConvertNavToMenu3 "S01"
    
    RepAnyPrintOutAlloc "CovidCert", "s01"
   
    UpdPrintLayout "Nurses24HReport", "PrintInputFilter", "Date||Period"
    
    RepAnyPrintOutAlloc "ViewTestTemplate", "SystemAdmin"
    RepAnyPrintOutAlloc "ViewTestTemplate", "S13"
    
    ConvertNavToMenu3 "DPT020"
    UpdPrintLayout "Nurses24HReport", "PrintInputFilter", "Date||Period"
    
    ReplicateAccess "W001", "DPT020"
    ConvertNavToMenu3 "DPT020"
    
    RepAnyPrintOutAlloc "PatientReceiptSelector", "W001"
    RepAnyPrintOutAlloc "PatientReceiptSelector", "W002"
    
    RepAnyPrintOutAlloc "ReceiptSlip2", "S17A"
    
    RepAnyPrintOutAlloc "MonitorVisitationDrug", "S22"
    RepAnyPrintOutAlloc "MonitorVisitationDrugEmerg", "S22"
    
    RepAnyPrintOutAlloc "Admission1", "S17A"
    RepAnyPrintOutAlloc "VisitationBill", "S17A"
    RepAnyPrintOutAlloc "DynamicTableLoader", "S17A"
    RepAnyPrintOutAlloc "VisitationRCP", "S17A"
    ConvertNavToMenu3 "S17A"
    
    UpdJobSchedule "S17A", "JobscheduleName", "Cashier (Airport Clinic)"
    UpdJobSchedule "S17A", "DefaultUrl", "wpgPrtPrintlayoutAll.asp?PositionForTableName=WorkingDay&PrintLayoutName=DynamicTableLoader&LoadInterval=4000&ProcedureName=IncomingInvoicesJSON"
    
    RepUserRoleAlloc "S17", "S17", "S17A", "S17A"
    RepUserRoleAlloc "Cashier", "Cashier", "S17A", "S17A"
    UpdSystemUser "S17A", "JobScheduleID", "S17A"
    
    Exit Sub
    
    RepJobSchedule "S17", "S17A"
    RepSystemUser "S17", "S17A"
    ResetUserPwd "S17A"
    
    Exit Sub
    RepAnyPrintOutAlloc "PatientMedicalRecord", "W001"
    
    Exit Sub
    ConvertNavToMenu3 "S01"
    Exit Sub
    CopyPrintOutAccess "Investigation1TH", "PrintLabResults"
    
    Exit Sub
    ConvertNavToMenu3 "S21A"
    Exit Sub
    UpdJobSchedule "S20", "JobScheduleName", "GENERAL STORE"
    UpdJobSchedule "S21A", "JobScheduleName", "GENERAL & MEDICAL STORE"
    CompileMergedNav "S21A"
    ConvertNavToMenu3 "S21A"
    
    Exit Sub
    RepJobSchedule "S21", "S21A"
    MergeJobSchedule "S21A", "S21||S20"
    
    RepSystemUser "S21", "S21A"
    UpdSystemUser "S21A", "JobScheduleID", "S21A"
    ResetUserPwd "S21A"
    
    Exit Sub
    
    
End Sub
Sub GrantAccessToTransition(profile, tableName, proCodeStart, proCodeEnd)
    Dim sql, rst, rt
    
    Set rst = CreateObject("ADODB.RecordSet")
    sql = "select * from TransProcessor2 where TransProcessor2ID='" & profile & "'"
    rst.open sql, conn, 3, 4
    If rst.RecordCount = 0 Then
        rst.AddNew
        rst.fields("TransProcessor2ID") = profile
        rst.fields("TransProcessor2Name") = GetComboName("JobSchedule", profile)
        rst.fields("TransProcessorTypeID") = "T001"
        rst.fields("ProcessorLevel") = 1
        rst.fields("InitialScheduleID") = profile
        rst.UpdateBatch
    End If
    rst.Close
    
    rt = tableName & "-" & proCodeStart & "-" & proCodeEnd
    sql = "select * from TransProcessorAcc2 where TransProcessor2ID='" & profile & "' and TransProcessRightID='" & rt & "' "
    rst.open sql, conn, 3, 4
    If rst.RecordCount = 0 Then
        rst.AddNew
        rst.fields("TransProcessor2ID") = profile
        rst.fields("TransProcessRightID") = rt
        rst.fields("TableID") = tableName
        rst.fields("TransProcessStatID") = proCodeStart
        rst.fields("TransProcessStat2ID") = proCodeEnd
        rst.fields("TransProcessorTypeID") = "L001"
        rst.fields("InitialScheduleID") = profile
        rst.fields("AccessPos") = 1
        rst.fields("AccessInfo") = 1
        rst.UpdateBatch
    End If
    rst.Close
    
    Set rst = Nothing
End Sub
Sub AddTransitionProcess(tableName, proCode, proName)
    Dim sql, rst, trnsPro, idx
    
    trnsPro = tableName & "-" & proCode
    Set rst = CreateObject("ADODB.RecordSet")
    
    sql = "select * from TransProcessVal where TransProcessValID='" & trnsPro & "'"
    rst.open sql, conn, 3, 4
    If rst.RecordCount = 0 Then
        rst.AddNew
        rst.fields("TransProcessValID") = trnsPro
        rst.fields("TransProcessValName") = proName
        rst.UpdateBatch
    End If
    rst.Close
    
    sql = "select * from TransProcessVal2 where TransProcessVal2ID='" & trnsPro & "' "
    rst.open sql, conn, 3, 4
    If rst.RecordCount = 0 Then
        rst.AddNew
        rst.fields("TransProcessVal2ID") = trnsPro
        rst.fields("TransProcessVal2Name") = proName
        rst.fields("TransProcessTblID") = tableName
        rst.fields("TransProcessUseID") = "T001"
        rst.UpdateBatch
    End If
    rst.Close
    
    idx = Right(proCode, 2)
    If IsNumeric(idx) Then
        idx = Right(CStr(1000 + CInt(idx)), 2)
        sql = "select * from TransProcessStat where TransProcessStatID='T0" & idx & "'"
        rst.open sql, conn, 3, 4
        If rst.RecordCount = 0 Then
            rst.AddNew
            rst.fields("TransProcessStatID") = "T0" & idx
            rst.fields("TransProcessStatName") = "Stage " & CInt(idx)
            rst.UpdateBatch
        End If
        rst.Close
        
        sql = "select * from TransProcessStat2 where TransProcessStat2ID='T0" & idx & "'"
        rst.open sql, conn, 3, 4
        If rst.RecordCount = 0 Then
            rst.AddNew
            rst.fields("TransProcessStat2ID") = "T0" & idx
            rst.fields("TransProcessStat2Name") = "Stage " & CInt(idx)
            rst.UpdateBatch
        End If
        rst.Close
    End If
    
    ''''
    'add possible transitions
    sql = " insert into TransProcessRight(TransProcessRightID, TransProcessRightName, TableID, TransProcessStatID"
    sql = sql & " , TransProcessStat2ID, UserAccessibleID, RightDetail, RightPos, Description, KeyPrefix)"
    sql = sql & " select '" & tableName & "-' + TransProcessStat.TransProcessStatID + '-' + TransProcessStat2.TransProcessStat2ID "
    sql = sql & "   , '" & tableName & " > '+ TransProcessVal.TransProcessValName +' > ' + TransProcessVal2.TransProcessVal2Name "
    sql = sql & "   , '" & tableName & "', TransProcessStat.TransProcessStatID, TransProcessStat2.TransProcessStat2ID, 'USA001' "
    sql = sql & "   , '-', 1, null, null "
    sql = sql & " from TransProcessVal2 cross join TransProcessVal"
    sql = sql & " left join TransProcessStat ON TransProcessVal.TransProcessValid= '" & tableName & "-' +TransProcessStat.TransProcessStatID "
    sql = sql & " left join TransProcessStat2 ON TransProcessVal2.TransProcessVal2id= '" & tableName & "-' +TransProcessStat2.TransProcessStat2ID"
    sql = sql & " left join TransProcessRight on TransProcessRight.TransProcessRightID=('" & tableName & "-' + TransProcessStat.TransProcessStatID + '-' + TransProcessStat2.TransProcessStat2ID)"
    sql = sql & "   where TransProcessVal2.TransProcessTblID='" & tableName & "' and TransProcessVal.TransProcessValID like '" & tableName & "%'"
    sql = sql & "       and TransProcessRIght.TransProcessRightID is null;"
    
    conn.execute sql
    ''''
    ''''
    
    
    
    Set rst = Nothing
End Sub
Sub UpdTransitionProcessName(tableName, proCode, proName)
    Dim sql, rst, trnsPro, idx
    
    trnsPro = tableName & "-" & proCode
    Set rst = CreateObject("ADODB.RecordSet")
    
    sql = "select * from TransProcessVal where TransProcessValID='" & trnsPro & "'"
    rst.open sql, conn, 3, 4
    If rst.RecordCount = 0 Then
        rst.AddNew
        rst.fields("TransProcessValID") = trnsPro
        rst.fields("TransProcessValName") = proName
        rst.UpdateBatch
    Else
        rst.fields("TransProcessValName") = proName
        rst.UpdateBatch
    End If
    rst.Close
    
    sql = "select * from TransProcessVal2 where TransProcessVal2ID='" & trnsPro & "' "
    rst.open sql, conn, 3, 4
    If rst.RecordCount = 0 Then
        rst.AddNew
        rst.fields("TransProcessVal2ID") = trnsPro
        rst.fields("TransProcessVal2Name") = proName
        rst.fields("TransProcessTblID") = tableName
        rst.fields("TransProcessUseID") = "T001"
        rst.UpdateBatch
    Else
        rst.fields("TransProcessVal2Name") = proName
        rst.UpdateBatch
    End If
    rst.Close
    
    Set rst = Nothing
End Sub
Sub ReplicateProfile(source, target, targetName)
    RepJobSchedule source, target
    UpdJobSchedule target, "JobScheduleName", targetName
    RepSystemUser source, target
    UpdSystemUser target, "JobScheduleID", target
    ResetUserPwd target
    ReplicateAccess source, target
    ReplicateModuleManager source, target
    ConvertNavToMenu3 target
End Sub

Sub ReplicateModuleManager(source, target)
    Dim sql, rst, rst2
    
    Set rst = CreateObject("ADODB.RecordSet")
    Set rst2 = CreateObject("ADODB.RecordSet")
    sql = "select * from ModuleManagerAlloc where JobScheduleID='" & source & "' "
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
        
            sql = "select * from ModuleManagerAlloc where JobScheduleID='" & target & "' "
            rst2.open sql, conn, 3, 4
            If rst2.RecordCount = 0 Then
                rst2.AddNew
                For Each field In rst2.fields
                    If UCase(field.name) = UCase("JobScheduleID") Then
                        rst2.fields(field.name) = target
                    Else
                        rst2.fields(field.name) = rst.fields(field.name)
                    End If
                Next
                rst2.UpdateBatch
            End If
            
            rst2.Close
            rst.MoveNext
        Loop
        
        rst.Close
    End If
    
End Sub
Sub GrantTblAccessToPrintlayout(tbl, printlayoutName)
    Dim sql, rst
    
    sql = "select UserRoleID from AccessRightAlloc where TableID='" & tbl & "' "
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            RepAnyPrintOutAlloc printlayoutName, rst.fields("UserRoleID")
            rst.MoveNext
        Loop
        rst.Close
    End If
    Set rst = Nothing
End Sub
Sub CopyPrintOutAccess(sourcePrint, targetPrint)
    Dim sql, rst
    
    Set rst = CreateObject("ADODB.RecordSet")
    sql = "select * from PrintOutAlloc where PrintLayoutID='" & sourcePrint & "' "
    rst.open sql, conn, 3, 4
    
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            
            RepAnyPrintOutAlloc targetPrint, rst.fields("JobScheduleID")
            rst.MoveNext
        Loop
    End If
    rst.Close
End Sub


Sub AddProcessCall2(tbl, usrPro, proPnt, proPos)
    AddProcessCall tbl, usrPro, proPnt
    UpdProcessCall (tbl & "_" & usrPro & "_" & proPnt), "ProcessPos", proPos
End Sub



Sub ReplicateAccess(sourceJb, targetJb)
    Dim sql, rst
    'tables
    sql = " select * from AccessRightAlloc where UserRoleID='" & sourceJb & "' "
    Set rst = CreateObject("ADODB.Recordset")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            RepAccessRightAlloc rst.fields("UserRoleID"), rst.fields("AccessRightID"), targetJb, rst.fields("AccessRightID")
            rst.MoveNext
        Loop
        rst.Close
    End If
    
    'printouts
    sql = " select * from PrintOutAlloc where JobScheduleID='" & sourceJb & "' "
    Set rst = CreateObject("ADODB.Recordset")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            RepPrintOutAlloc rst.fields("PrintLayoutID"), rst.fields("JobScheduleID"), rst.fields("PrintLayoutID"), targetJb
            rst.MoveNext
        Loop
        rst.Close
    End If
    
    'userrole
    sql = " select * from UserRoleAlloc where JobScheduleID='" & sourceJb & "' "
    Set rst = CreateObject("ADODB.Recordset")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            'MsgBox rst.fields("JobScheduleID")
            RepUserRoleAlloc rst.fields("UserRoleID"), rst.fields("JobScheduleID"), rst.fields("UserRoleID"), targetJb
            
            'transitions
            sql = "select * from TransProcessorAcc2 where TransProcessor2ID='" & rst.fields("UserRoleID") & "'  "
            Set rst2 = CreateObject("ADODB.RecordSet")
            rst2.open sql, conn, 3, 4
            If rst2.RecordCount > 0 Then
                rst2.movefirst
                Do While Not rst2.EOF
                    GrantAccessToTransition targetJb, rst2.fields("TableID").value, rst2.fields("TransProcessStatID").value, rst2.fields("TransProcessStat2ID").value
                    rst2.MoveNext
                Loop
            End If
            rst2.Close
            
            rst.MoveNext
        Loop
        rst.Close
    End If
    
    
End Sub







'fixbatch
Sub fixbatch()
    Dim sql, rst, rst2, dt
    
    sql = "select distinct insuredPatient.FirstDayID From InsuredPatient "
    sql = sql & " left join FirstDay on FirstDay.Firstdayid=insuredPatient.Firstdayid "
    sql = sql & " where FirstDay.Firstdayid is null "
    sql = sql & " order by insuredPatient.Firstdayid desc"
    
    
    Set rst = CreateObject("ADODB.RecordSet")
    Set rst2 = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            dt = rst.fields("FirstDayID")
            dt = gExtractWorkingDate(dt)
            FormatWorkingDayAdd dt
            
            rst.MoveNext
        Loop
        rst.Close
    End If
    
End Sub
Function gExtractWorkingDate(wkDay)
    Dim str, ot
    ot = Null
    str = Trim(wkDay)
      If UCase(Left(str, 3)) = "YRS" Then
        If Len(str) = 7 Then
        ot = CDate("1 " & MonthName(1, 1) & " " & Mid(str, 4, 4))
        End If
      ElseIf UCase(Left(str, 3)) = "MTH" Then
        If Len(str) = 9 Then
        ot = CDate("1 " & MonthName(CInt(Mid(str, 8, 2)), 1) & " " & Mid(str, 4, 4))
        End If
      ElseIf UCase(Left(str, 3)) = "QTR" Then
        If Len(str) = 9 Then
        ot = CDate("1 " & MonthName(((CInt(Mid(str, 8, 2)) * 3) - 2), 1) & " " & Mid(str, 4, 4))
        End If
      ElseIf UCase(Left(str, 3)) = "DAY" Then
        If Len(str) = 11 Then
        ot = CDate(Mid(str, 10, 2) & " " & MonthName(CInt(Mid(str, 8, 2)), 1) & " " & Mid(str, 4, 4))
        End If
      End If
    gExtractWorkingDate = ot
End Function
Function FormatWorkingMonthAdd(dt)
    Dim ot, sql, wkMthName, kp

    ot = ""
    If IsDate(dt) Then
        ot = "MTH" & CStr(Year(CDate(dt))) & Right(CStr(Month(CDate(dt)) + 100), 2)
        wkMthName = MonthName(Month(dt)) & " " & Year(dt)
        kp = Month(dt) & Year(dt)
        kp = Right(CStr(Year(dt)), 2) & Right(CStr(100 + Month(dt)), 2)

        sql = " if not exists( select * from WorkingMonth where WorkingMonthID='" & ot & "' ) "
        sql = sql & " INSERT INTO WorkingMonth (WorkingMonthID, WorkingMonthName, WorkingYearID, WorkingQuarterID, WorkMonth, KeyPrefix, Description)"
        sql = sql & " VALUES ('" & ot & "', '" & wkMthName & "', '" & FormatWorkingYearAdd(dt) & "', '" & FormatWorkingQuarterAdd(dt) & "', '" & wkMthName & "', '" & kp & "', '/');"
        conn.execute qryPro.FltQry(sql)

        sql = " if not exists( select * from FirstMonth where FirstMonthID='" & ot & "' ) "
        sql = sql & " INSERT INTO FirstMonth (FirstMonthID, FirstMonthName, FirstYearID, FirstQuarterID, FirstMonth, KeyPrefix, Description)"
        sql = sql & " VALUES ('" & ot & "', '" & wkMthName & "', '" & FormatWorkingYearAdd(dt) & "', '" & FormatWorkingQuarterAdd(dt) & "', '" & wkMthName & "', '" & kp & "', '/');"
        conn.execute qryPro.FltQry(sql)

        sql = " if not exists( select * from BillMonth where BillMonthID='" & ot & "' ) "
        sql = sql & " INSERT INTO BillMonth (BillMonthID, BillMonthName, BillYearID, BillQuarterID, BillMonth, KeyPrefix, Description)"
        sql = sql & " VALUES ('" & ot & "', '" & wkMthName & "', '" & FormatWorkingYearAdd(dt) & "', '" & FormatWorkingQuarterAdd(dt) & "', '" & wkMthName & "', '" & kp & "', '/');"
        conn.execute qryPro.FltQry(sql)

        sql = " if not exists( select * from AppointMonth where AppointMonthID='" & ot & "' ) "
        sql = sql & " INSERT INTO AppointMonth (AppointMonthID, AppointMonthName, AppointYearID, AppointQuarterID, AppointMonth, KeyPrefix, Description)"
        sql = sql & " VALUES ('" & ot & "', '" & wkMthName & "', '" & FormatWorkingYearAdd(dt) & "', '" & FormatWorkingQuarterAdd(dt) & "', '" & wkMthName & "', '" & kp & "', '/');"
        conn.execute qryPro.FltQry(sql)
    End If
         
    FormatWorkingMonthAdd = ot
End Function
Function FormatWorkingDayAdd(dt)
    Dim ot, sql, wkMth, wkDyName, kp, wkYr

    ot = ""
    If IsDate(dt) Then
        dt = CDate(dt)

        ot = "DAY" & CStr(Year(CDate(dt))) & Right(CStr(Month(CDate(dt)) + 100), 2) & Right(CStr(Day(CDate(dt)) + 100), 2)
        wkMth = FormatWorkingMonthAdd(dt)
        wkYr = FormatWorkingYearAdd(dt)
        wkDyName = Day(dt) & " " & MonthName(Month(dt)) & " " & Year(dt) & " [" & WeekdayName(Weekday(dt), True) & "]"
        kp = Right(CStr(Year(dt)), 2) & Right(100 + Month(dt), 2) & Right(100 + Day(dt), 2)

        sql = " if not exists( select * from WorkingDay where WorkingDayID='" & ot & "' ) "
        sql = sql & "INSERT INTO WorkingDay (WorkingDayID, WorkingDayName, WorkingMonthID, WorkDate, KeyPrefix, Description) "
        sql = sql & " VALUES ('" & ot & "', '" & wkDyName & "', '" & wkMth & "', '" & dt & "', '" & kp & "', '/');"
        conn.execute qryPro.FltQry(sql)

        sql = " if not exists( select * from FirstDay where FirstDayID='" & ot & "' ) "
        sql = sql & "INSERT INTO FirstDay (FirstDayID, FirstDayName, FirstMonthID, FirstDate, KeyPrefix, Description) "
        sql = sql & " VALUES ('" & ot & "', '" & wkDyName & "', '" & wkMth & "', '" & dt & "', '" & kp & "', '/');"
        conn.execute qryPro.FltQry(sql)

        sql = " if not exists( select * from BillDay where BillDayID='" & ot & "' ) "
        sql = sql & "INSERT INTO BillDay (BillDayID, BillDayName, BillMonthID, BillDate, KeyPrefix, Description) "
        sql = sql & " VALUES ('" & ot & "', '" & wkDyName & "', '" & wkMth & "', '" & dt & "', '" & kp & "', '/');"
        conn.execute qryPro.FltQry(sql)

        sql = " if not exists( select * from AppointDay where AppointDayID='" & ot & "' ) "
        sql = sql & "INSERT INTO AppointDay (AppointDayID, AppointDayName, AppointMonthID, AppointDate, KeyPrefix, Description) "
        sql = sql & " VALUES ('" & ot & "', '" & wkDyName & "', '" & wkMth & "', '" & dt & "', '" & kp & "', '/');"
        conn.execute qryPro.FltQry(sql)

        sql = " if not exists(select * from BranchBatch where BranchBatchID='B001-" & ot & "' )"
        sql = sql & "insert into BranchBatch (BranchBatchID, BranchBatchName, BranchBatchTypeID, BranchBatchStatusID, BatchPos, BranchID, WorkingYearID, WorkingMonthID, WorkingDayID, BatchDate, BatchInfo, Description, KeyPrefix) "
        sql = sql & " values ('B001-" & ot & "', 'FOCOS [" & wkDyName & "]', 'B001', 'B003', 1, 'B001', '" & wkYr & "', '" & wkMth & "', '" & ot & "', '" & dt & "', NULL, NULL, NULL)"
        conn.execute qryPro.FltQry(sql)

        sql = "if not exists(select * from BranchSubBatch where BranchSubBatchID='B001-" & ot & "-Receipt') "
        sql = sql & " insert into BranchSubBatch (BranchSubBatchID, BranchSubBatchName, BranchSubBatchTypeID, BranchBatchID, BranchBatchTypeID, BranchSubBatchStatID, SubBatchPos, BranchID, WorkingYearID, WorkingMonthID, WorkingDayID, SubBatchDate, SubBatchInfo, Description, KeyPrefix)"
        sql = sql & " values('B001-" & ot & "-Receipt', 'FOCOS [" & wkDyName & "Receipt]', 'Receipt', 'B001-" & ot & "', 'B001', 'B003', 1, 'B001', '" & wkYr & "', '" & wkMth & "', '" & wkDyName & "', '" & dt & "', NULL, NULL, NULL) "
        conn.execute qryPro.FltQry(sql)

        sql = "if not exists(select * from BranchSubBatch where BranchSubBatchID='B001-" & ot & "-Visitation') "
        sql = sql & " insert into BranchSubBatch (BranchSubBatchID, BranchSubBatchName, BranchSubBatchTypeID, BranchBatchID, BranchBatchTypeID, BranchSubBatchStatID, SubBatchPos, BranchID, WorkingYearID, WorkingMonthID, WorkingDayID, SubBatchDate, SubBatchInfo, Description, KeyPrefix)"
        sql = sql & " values('B001-" & ot & "-Visitation', 'FOCOS [" & wkDyName & "Visitation]', 'Visitation', 'B001-" & ot & "', 'B001', 'B003', 1, 'B001', '" & wkYr & "', '" & wkMth & "', '" & wkDyName & "', '" & dt & "', NULL, NULL, NULL) "
        conn.execute qryPro.FltQry(sql)

    End If

    FormatWorkingDayAdd = ot
End Function
Function FormatWorkingQuarterAdd(dt)
    Dim mth, ot

    ot = ""
    If IsDate(dt) Then
        mth = Month(CDate(dt))
        ot = "QTR" & CStr(Year(CDate(dt))) & Right(CStr((Int((mth - 1) / 3) + 1) + 100), 2)
    End If

    FormatWorkingQuarterAdd = ot
End Function
Function FormatWorkingYearAdd(dt)
    Dim ot, sql

    ot = ""
    If IsDate(dt) Then
        ot = "YRS" & CStr(Year(CDate(dt)))
        
        sql = " if not exists( select * from WorkingYear where WorkingYearID='" & ot & "' ) "
        sql = sql & " INSERT INTO WorkingYear (WorkingYearID, WorkingYearName, WorkYear, KeyPrefix, Description) "
        sql = sql & " VALUES ('" & ot & "', '" & Year(dt) & "', '" & Year(dt) & "', '" & Right(CStr(Year(dt)), 2) & "', '/'); "
        conn.execute qryPro.FltQry(sql)

        sql = " if not exists( select * from FirstYear where FirstYearID='" & ot & "' ) "
        sql = sql & " INSERT INTO FirstYear (FirstYearID, FirstYearName, FirstYear, KeyPrefix, Description) "
        sql = sql & " VALUES ('" & ot & "', '" & Year(dt) & "', '" & Year(dt) & "', '" & Right(CStr(Year(dt)), 2) & "', '/'); "
        conn.execute qryPro.FltQry(sql)
        
        sql = " if not exists( select * from BillYear where BillYearID='" & ot & "' ) "
        sql = sql & " INSERT INTO BillYear (BillYearID, BillYearName, BillYear, KeyPrefix, Description) "
        sql = sql & " VALUES ('" & ot & "', '" & Year(dt) & "', '" & Year(dt) & "', '" & Right(CStr(Year(dt)), 2) & "', '/'); "
        conn.execute qryPro.FltQry(sql)

        sql = " if not exists( select * from AppointYear where AppointYearID='" & ot & "' ) "
        sql = sql & " INSERT INTO AppointYear (AppointYearID, AppointYearName, AppointYear, KeyPrefix, Description) "
        sql = sql & " VALUES ('" & ot & "', '" & Year(dt) & "', '" & Year(dt) & "', '" & Right(CStr(Year(dt)), 2) & "', '/'); "
        conn.execute qryPro.FltQry(sql)
    End If

    FormatWorkingYearAdd = ot
End Function

Sub RecompileMenus(pFx)
    Dim sql, rst
    
    sql = "select JobscheduleID from JobSchedule "
    If pFx <> "" Then
        sql = sql & " where JobScheduleID like '" & pFx & "%' "
    End If
    sql = sql & " order by JobScheduleID asc "
    
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            CompileMergedNav rst.fields("JobScheduleID")
            ConvertNavToMenu3 rst.fields("JobScheduleID")
            rst.MoveNext
        Loop
    End If
    
    rst.Close
    Set rst = Nothing
End Sub
Sub mapRXDrug()
    Dim sql, rst, rst2, rst3, xDconn
    
    Set xDconn = CreateObject("ADODB.Connection")
    xDconn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=admin;Password=drowssaprofas;Initial Catalog=RMCData;Data Source=192.168.1.81"
    xDconn.cursorlocation = 3
    xDconn.open
    
    sql = "select * from Drug where DrugStatusID='IST001' "
    Set rst = CreateObject("ADODB.RecordSet")
    Set rst2 = CreateObject("ADODB.RecordSet")
    Set rst3 = CreateObject("ADODB.RecordSet")
    
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
        
            sql = "select top 1 * from [service_items$] where item='" & rst.fields("DrugName") & "' "
            rst2.open sql, xDconn, 3, 4
            If rst2.RecordCount > 0 Then
                sql = "update [service_items$] set [map_status]='completed', [medical_item_code]='" & rst.fields("DrugID") & "'  where code='" & rst2.fields("code") & "'"
                xDconn.execute sql
                
                sql = "select * from EMRVar99 where EMRVar99ID='DRG||" & rst.fields("DrugID") & "' "
                rst3.open sql, conn, 3, 4
                If rst3.RecordCount = 0 Then
                    rst3.AddNew
                    rst3.fields("EMRVar99ID") = "DRG||" & rst.fields("DrugID")
                    
                End If
                rst3.fields("EMRVar99Name") = rst.fields("DrugName")
                rst3.fields("KeyPrefix") = rst2.fields("code")
                rst3.fields("Description") = rst2.fields("item")
                rst3.fields("VarPos") = 0
                rst3.UpdateBatch
                
                rst3.Close
            End If
            rst2.Close
            rst.MoveNext
        Loop
    End If
End Sub
Sub mapRXDrug2()
    Dim sql, rst, rst2, rst3, xDconn
    
    Set xDconn = CreateObject("ADODB.Connection")
    xDconn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=admin;Password=drowssaprofas;Initial Catalog=RMCData;Data Source=192.168.1.81"
    xDconn.cursorlocation = 3
    xDconn.open
    
    sql = "select * from Drug where DrugStatusID='IST001' "
    Set rst = CreateObject("ADODB.RecordSet")
    Set rst2 = CreateObject("ADODB.RecordSet")
    Set rst3 = CreateObject("ADODB.RecordSet")
    
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
        
            sql = "select top 1 * from [item_list_27_07_22$] where item='" & rst.fields("DrugName") & "' and ins_item_code is not null;"
            rst2.open sql, xDconn, 3, 4
            If rst2.RecordCount > 0 Then
                sql = "update [item_list_27_07_22$] set [map_status]='completed', [medical_item_code]='" & rst.fields("DrugID") & "'  where ins_item_code='" & rst2.fields("ins_item_code") & "'"
                xDconn.execute sql
                
                sql = "select * from EMRVar99 where EMRVar99ID='DRG||" & rst.fields("DrugID") & "' "
                rst3.open sql, conn, 3, 4
                If rst3.RecordCount = 0 Then
                    rst3.AddNew
                    rst3.fields("EMRVar99ID") = "DRG||" & rst.fields("DrugID")
                    
                End If
                rst3.fields("EMRVar99Name") = rst.fields("DrugName")
                rst3.fields("KeyPrefix") = rst2.fields("ins_item_code")
                rst3.fields("Description") = rst2.fields("item")
                rst3.fields("VarPos") = 0
                rst3.UpdateBatch
                
                rst3.Close
            End If
            rst2.Close
            rst.MoveNext
        Loop
    End If
End Sub
Sub updateNationwideDrugPrice()
    Dim sql, rst, rst2, rst3, xDconn
    
    Set xDconn = CreateObject("ADODB.Connection")
    xDconn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=admin;Password=drowssaprofas;Initial Catalog=RMCData;Data Source=192.168.1.81"
    xDconn.cursorlocation = 3
    xDconn.open
    
    sql = "select * from Sheet2$ where [RETAIL PRICE] is null;"
    Set rst = CreateObject("ADODB.RecordSet")
    Set rst2 = CreateObject("ADODB.RecordSet")
    
    rst.open sql, xDconn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            sql = "select * from Drug where DrugName='" & Replace(rst.fields("DRUG NAME"), "'", "''") & "' "
            rst2.open sql, conn, 3, 4
            If rst2.RecordCount > 0 Then
                sql = "update DrugPriceMatrix2 set ItemUnitCost=" & rst.fields("RETAIL PRICE") & " where DrugID='" & rst2.fields("DrugID") & "' and InsuranceTypeID='I108' "
                conn.execute sql
                rst.fields("map_status") = "completed"
                rst.UpdateBatch
            End If
            rst2.Close
            rst.MoveNext
        Loop
    End If
    rst.Close
    Set rst = Nothing
End Sub
Sub FixDrugMat()
    Dim sql, rst, matRst, insRst
    
    Set rst = CreateObject("ADODB.RecordSet")
    Set matRst = CreateObject("ADODB.RecordSet")
    Set insRst = CreateObject("ADODB.RecordSet")

    sql = "select *  from DrugPriceMatrix2 where InsuranceTypeID='I100' "
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
        
            sql = "select * from Insurancetype where InsuranceTypeID not in ('I108', 'I100', 'I101') "
            insRst.open sql, conn, 3, 4
            If insRst.RecordCount > 0 Then
                Do While Not insRst.EOF
                
                    sql = "select * from DrugPriceMatrix2 where InsuranceTypeID='" & insRst.fields("InsuranceTypeID") & "' "
                    sql = sql & " and DrugID='" & rst.fields("DrugID") & "' and DrugStoreTypeID='" & rst.fields("DrugStoreTypeID") & "'"
                    
                    matRst.open sql, conn, 3, 4
                    If matRst.RecordCount = 0 Then
                        matRst.AddNew
                        For Each field In rst.fields
                            matRst.fields(field.name) = field.value
                        Next
                        matRst.fields("InsuranceTypeID") = insRst.fields("InsuranceTypeID")
                    Else
                        matRst.fields("ItemUnitCost") = rst.fields("ItemUnitCost")
                    End If
                    
                    matRst.UpdateBatch
                    matRst.Close
                    
                    insRst.MoveNext
               Loop
                
            End If
            
            insRst.Close
            rst.MoveNext
        Loop
    End If
    
    rst.Close
End Sub

Sub FixInsuredPatient()
    Dim tmp, field, rst, insRst, tmp2, vstRst
    
    tmp = Split("InsuranceTypeID||ReceiptTypeID||InsuranceZoneID||InsuranceGroupID||SponsorID||SponsorTypeID||VettingGroupID", "||")
    tmp2 = Split("InsSchemeModeID||PatientID||BenefitTypeID||GenderID||InsuranceNo||InsuranceSchemeID||InsuranceTypeID||ReceiptTypeID||InsuranceZoneID||InitialSystemUserID||InitialScheduleID||InitialBranchID||PatientRankID||PatientRankLevelID||PatientUnitID||ServiceNo||InitialDependantID||InsuredModeID||InsuredPrincipalID", "||")
    
    sql = "select  * from InsuredPatient where Len(PatientID)>4 order by InsuranceSchemeID asc"
    
    Set rst = CreateObject("ADODB.RecordSet")
    Set insRst = CreateObject("ADODB.RecordSet")
    Set vstRst = CreateObject("ADODB.RecordSet")
    
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            sql = "select * from InsuranceScheme where InsuranceSchemeID='" & rst.fields("InsuranceSchemeID") & "'"
            insRst.open sql, conn, 3, 4
            If insRst.RecordCount > 0 Then
                For Each field In tmp
                    rst.fields(field) = insRst.fields(field)
                Next
                rst.UpdateBatch
            End If
            insRst.Close
            
            sql = "select * from Visitation where InsuredPatientID='" & rst.fields("InsuredPatientID") & "' "
            vstRst.open sql, conn, 3, 4
            If vstRst.RecordCount > 0 Then
                vstRst.movefirst
                Do While Not vstRst.EOF
                    For Each field In tmp2
                        vstRst.fields(field) = rst.fields(field)
                    Next
                    
                    vstRst.UpdateBatch
                    vstRst.MoveNext
                Loop
            End If
            vstRst.Close
            
            rst.MoveNext
        Loop
    End If
    
    rst.Close
    Set rst = Nothing
End Sub
Sub FixOneInsuredPatient(inspat)
    Dim tmp, field, rst, insRst, tmp2, vstRst
    
    tmp = Split("InsuranceTypeID||ReceiptTypeID||InsuranceZoneID||InsuranceGroupID||SponsorID||SponsorTypeID||VettingGroupID", "||")
    tmp2 = Split("InsSchemeModeID||PatientID||BenefitTypeID||GenderID||InsuranceNo||InsuranceSchemeID||InsuranceTypeID||ReceiptTypeID||InsuranceZoneID||InitialSystemUserID||InitialScheduleID||InitialBranchID||PatientRankID||PatientRankLevelID||PatientUnitID||ServiceNo||InitialDependantID||InsuredModeID||InsuredPrincipalID", "||")
    
    sql = "select  * from InsuredPatient where InsuredPatientID='" & inspat & "' "
    
    Set rst = CreateObject("ADODB.RecordSet")
    Set insRst = CreateObject("ADODB.RecordSet")
    Set vstRst = CreateObject("ADODB.RecordSet")
    
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            sql = "select * from InsuranceScheme where InsuranceSchemeID='" & rst.fields("InsuranceSchemeID") & "'"
            insRst.open sql, conn, 3, 4
            If insRst.RecordCount > 0 Then
                For Each field In tmp
                    rst.fields(field) = insRst.fields(field)
                Next
                rst.UpdateBatch
            End If
            insRst.Close
            
            sql = "select * from Visitation where InsuredPatientID='" & rst.fields("InsuredPatientID") & "' "
            vstRst.open sql, conn, 3, 4
            If vstRst.RecordCount > 0 Then
                vstRst.movefirst
                Do While Not vstRst.EOF
                    For Each field In tmp2
                        vstRst.fields(field) = rst.fields(field)
                    Next
                    
                    vstRst.UpdateBatch
                    vstRst.MoveNext
                Loop
            End If
            vstRst.Close
            
            rst.MoveNext
        Loop
    End If
    
    rst.Close
    Set rst = Nothing
End Sub
Sub FixVisitBillForSponsor(spn)
    Dim xmlhttp, url, sql, rst
    
    sql = "select top 1 * from Visitation where SponsorID ='" & spn & "' "
    sql = sql & " and VisitationID='V1220703047' "
    sql = sql & " order by VisitationID asc "
    
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    
    MsgBox "Working on " & rst.RecordCount
    
    If rst.RecordCount > 0 Then
        rst.movefirst
        
        Do While Not rst.EOF
            'FixOneInsuredPatient rst.fields("InsuredPatientID")
            'Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
            Set xmlhttp = CreateObject("MSXML2.XMLHTTP.3.0")
            'Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
            'Set xmlhttp = CreateObject("Microsoft.XMLHttp")
            
            url = "http://192.168.1.80/hms/wpgXMLHttp.asp?ProcedureName=GeneratePatientBill&inpPatientID=" & rst.fields("PatientID") & "&inpVisitationID=" & rst.fields("VisitationID") & "&TableName=Visitation&ApplyVisitCost=Yes"
            xmlhttp.open "GET", url, False
        
            xmlhttp.send ""
            
'            If xmlhttp.waitForResponse(100) Then 'timeout in 100secs'
'                MsgBox "here"
'                If Len(xmlhttp.responseText) > 0 Then
'                    MsgBox xmlhttp.responseText
'                Else
'                    sql = "update PerformVar30 set PerformVar30Name=PerformVar30Name + '||" & vst & "'  where PerformVar30ID='TESTVAR' "
'                    conn.execute sql
'                End If
'            Else 'wait timeout exceeded
'                sql = "update PerformVar30 set PerformVar30Name=PerformVar30Name + '||" & vst & "'  where PerformVar30ID='TESTVAR' "
'                conn.execute sql
'            End If

            MsgBox url
            MsgBox xmlhttp.readyState & "--" & xmlhttp.Status
            MsgBox xmlhttp.responseText
            rst.MoveNext
        Loop
    End If
    
    rst.Close
    Set rst = Nothing
End Sub
Sub FixVisitBills()
    Dim xmlhttp, url, sql, rst
    
    sql = "select * from Visitation where SponsorID in ('S001', 'S004', 'S005', 'S006', 'S007', 'S008', 'S009', 'S011', 'S012', 'S013'"
    sql = sql & " , 'S014', 'S015', 'S016', 'S017', 'S018', 'S019', 'S020', 'S021', 'S022', 'S023', 'S024', 'S025', 'S026', 'S027', 'S028'"
    sql = sql & " , 'S029', 'S030', 'S031', 'S032', 'S033', 'S034', 'S035', 'S036', 'S038', 'S039', 'S040', 'S041', 'S042', 'S043', 'S044'"
    sql = sql & " , 'S045', 'S046', 'S047', 'S048', 'S050', 'S052', 'S053', 'S054', 'S055', 'S056', 'S057', 'S059', 'S060', 'S061', 'S062'"
    sql = sql & " , 'S063', 'S064', 'S065', 'S066', 'S067', 'S068', 'S069', 'S070', 'S071', 'S072', 'S073', 'S074', 'S075', 'S076', 'S077'"
    sql = sql & " , 'S078', 'S079', 'S081', 'S082', 'S083', 'S084', 'S086', 'S087', 'S089', 'S090', 'S091', 'S092', 'S093', 'S094', 'S095'"
    sql = sql & " , 'S096', 'S097', 'S098', 'S099', 'S100', 'S101', 'S102', 'S103', 'S104', 'S106', 'S107', 'S108', 'S109', 'S110', 'S111'"
    sql = sql & " , 'S112', 'S113', 'S114', 'S115', 'S116', 'S117', 'S118', 'S119', 'S120', 'S121', 'S122', 'S123', 'S124', 'S125', 'S126'"
    sql = sql & " , 'S127', 'S128', 'S129', 'S130', 'S131', 'S132', 'S133', 'S134', 'S135', 'S136', 'S137', 'S138', 'S139', 'S140', 'S141'"
    sql = sql & " , 'S142', 'S143', 'S144', 'S145') "
    sql = sql & " "
    sql = sql & " order by VisitationID asc "
    
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    
    MsgBox "Working on " & rst.RecordCount
    
    If rst.RecordCount > 0 Then
        rst.movefirst
        
        Do While Not rst.EOF
            Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
            
            url = "http://192.168.1.80/hms/wpgXMLHttp.asp?ProcedureName=GeneratePatientBill&inpPatientID=" & rst.fields("PatientID") & "&inpVisitationID=" & rst.fields("PatientID") & "&TableName=Visitation&ApplyVisitCost=Yes"
            xmlhttp.open "GET", url ', True ', strRxUsername, strRxPassword 'async call'
            
            xmlhttp.send
            
            If xmlhttp.waitForResponse(100) Then 'timeout in 100secs'
                If Len(xmlhttp.responseText) > 0 Then
                Else
                    sql = "update PerformVar30 set PerformVar30Name=PerformVar30Name + '||" & vst & "'  where PerformVar30ID='TESTVAR' "
                    conn.execute sql
                End If
            Else 'wait timeout exceeded
                sql = "update PerformVar30 set PerformVar30Name=PerformVar30Name + '||" & vst & "'  where PerformVar30ID='TESTVAR' "
                conn.execute sql
            End If
            
            rst.MoveNext
        Loop
    End If
    
    rst.Close
    Set rst = Nothing
    
End Sub
Sub fixVisitBillsNationwide()
    Dim xmlhttp, url, sql, rst
    
    sql = "select * from Visitation where SponsorID in ('S088') and Visitation.VisitDate>='16 Jul 2022 00:00:00' order by VisitationID asc "
    
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    MsgBox "Working on " & rst.RecordCount
    If rst.RecordCount > 0 Then
        rst.movefirst
        
        Do While Not rst.EOF
            Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
            
            url = "http://192.168.1.80/hms/wpgXMLHttp.asp?ProcedureName=GeneratePatientBill&inpPatientID=" & rst.fields("PatientID") & "&inpVisitationID=" & rst.fields("PatientID") & "&TableName=Visitation&ApplyVisitCost=Yes"
            xmlhttp.open "GET", url ', True ', strRxUsername, strRxPassword 'async call'
            
            xmlhttp.send
            
            
            If xmlhttp.waitForResponse(100) Then 'timeout in 100secs'
                If Len(xmlhttp.responseText) > 0 Then
                Else
                    sql = "update PerformVar30 set PerformVar30Name=PerformVar30Name + '||" & vst & "'  where PerformVar30ID='TESTVAR' "
                    conn.execute sql
                End If
            Else 'wait timeout exceeded
                sql = "update PerformVar30 set PerformVar30Name=PerformVar30Name + '||" & vst & "'  where PerformVar30ID='TESTVAR' "
                conn.execute sql
            End If
            
            rst.MoveNext
        Loop
    End If
    
    rst.Close
    Set rst = Nothing
    
End Sub

Sub SetUpBillGroups()
    Dim sql, rst, bGrp, bGrpName
    
    UpdTableField "BillGroup", "BillGroupCatID", "VisibleInput", "1|1|1"
    UpdTableField "BillGroup", "BillGroupCatID", "VisibleField", "Yes||Yes"

    
    sql = "INSERT INTO PaymentType(PaymentTypeID, PaymentTypeName, PaymentCategoryID, BillGroupCatID, BillGroupID) "
    sql = sql & " SELECT BillGroup.billgroupid, BillGroup.BillGroupName, BillGroup.BillGroupID, BillGroup.BillGroupCatID, BillGroup.BillGroupID"
    sql = sql & " From BillGroup"
    sql = sql & " LEFT JOIN PaymentType ON PaymentType.paymenttypeid=BillGroup.BillGroupID"
    sql = sql & " WHERE PaymentType.PaymentTypeID IS NULL;"
    
    sql = sql & "  INSERT INTO PaymentCategory(paymentcategoryid, paymentcategoryname)"
    sql = sql & " SELECT BillGroup.BillGroupid, BillGroup.BillGroupname"
    sql = sql & " From BillGroup"
    sql = sql & " LEFT JOIN PaymentCategory ON BillGroup.billgroupid=PaymentCategory.paymentcategoryid"
    sql = sql & "  WHERE PaymentCategory.PaymentCategoryID IS NULL;"
    
    sql = sql & "  INSERT INTO PatientBillType(PatientBillTypeid, PatientBillTypename)"
    sql = sql & "  SELECT BillGroup.BillGroupid, BillGroup.BillGroupname"
    sql = sql & "  From BillGroup"
    sql = sql & "  LEFT JOIN PatientBillType ON BillGroup.billgroupid=PatientBillType.PatientBillTypeid"
    sql = sql & "  WHERE PatientBillType.PatientBillTypeid IS NULL;"
    
    conn.execute sql
    
    
    bGrp = "BG001"
    bGrpName = "GP Consult"
    AddBillGroup bGrp, bGrpName, "", ""

    sql = "update SpecialistType set BillGroupID='" & bGrp & "', BillGroupCatID='" & bGrp & "' where SpecialistGroupID='CD011' "
    conn.execute sql
    
    bGrp = "BG001.1"
    bGrpName = "Specialist Consult"
    AddBillGroup bGrp, bGrpName, "", ""

    sql = "update SpecialistType set BillGroupID='" & bGrp & "', BillGroupCatID='" & bGrp & "' where SpecialistGroupID<>'CD011' "
    conn.execute sql
    ApplySpecialistTypePullUp
    

    bGrp = "BG002"
    bGrpName = "Laboratory"
    AddBillGroup bGrp, bGrpName, "", ""
    sql = " ;update LabTest set TestContainerID='DPT005' where testcontainerid='T001' AND billgroupcatid='B13'"
    sql = sql & " ;update LabTest set BillGroupID='" & bGrp & "', BillGroupCatID='" & bGrp & "' where TestContainerID='DPT005' "
    conn.execute sql

    bGrp = "BG003"
    bGrpName = "Drugs"
    AddBillGroup bGrp, bGrpName, "", ""
    sql = " update Drug set BillGroupID='" & bGrp & "', billGroupCatID='" & bGrp & "' where BillGroupID='B15' "
    sql = sql & ";update Drug set BillGroupID='" & bGrp & "', billGroupCatID='" & bGrp & "' where BillGroupID='PHA' "
    sql = sql & ";update Drug set BillGroupID='" & bGrp & "', billGroupCatID='" & bGrp & "' where BillGroupID='REVALL' "
    sql = sql & " ;update BillGroup set BillGroupCatID='" & bGrp & "', BillGroupTypeID='" & bGrp & "' where BillGroupName like '%pha%'"
    sql = sql & " ;update BillGroup set BillGroupCatID='" & bGrp & "', BillGroupTypeID='" & bGrp & "' where BillGroupName like '%drug%'"
    'sql = sql & " ;update BillGroup set BillGroupCatID='" & bGrp & "', BillGroupTypeID='" & bGrp & "' where BillGroupName like '%all reve%'"

    conn.execute sql

    ApplyDrugPullup


    bGrp = "BG004"
    bGrpName = "Optical / Dental"
    AddBillGroup bGrp, bGrpName, "", ""

    sql = " update BillGroup set BillGroupCatID='" & bGrp & "', BillGroupTypeID='" & bGrp & "' where BillGroupName like '%dental%'"
    sql = sql & " ;update BillGroup set BillGroupCatID='" & bGrp & "', BillGroupTypeID='" & bGrp & "' where BillGroupName like '%opt%'"
    conn.execute sql


    bGrp = "BG005"
    bGrpName = "Procedures / Surgery"
    AddBillGroup bGrp, bGrpName, "", ""
    sql = " update BillGroup set BillGroupCatID='" & bGrp & "', BillGroupTypeID='" & bGrp & "' where BillGroupName like '%surg%'"
    sql = sql & " ;update BillGroup set BillGroupCatID='" & bGrp & "', BillGroupTypeID='" & bGrp & "' where BillGroupName like '%treatment%room%'"

    conn.execute sql


    bGrp = "BG006"
    bGrpName = "Admission / Ward"
    AddBillGroup bGrp, bGrpName, "", ""
    sql = " update BillGroup set BillGroupCatID='" & bGrp & "', BillGroupTypeID='" & bGrp & "' where BillGroupName like '%ward%'"
    sql = sql & "; update Items set BillGroupCatID='" & bGrp & "', BillGroupID='" & bGrp & "' where BillGroupID='B11'; "
    conn.execute sql


    bGrp = "BG007"
    bGrpName = "Radiology (X-Ray/Scan)"
    AddBillGroup bGrp, bGrpName, "", ""
    sql = sql & " ;update LabTest set TestContainerID='DPT011' where testcontainerid='T001' AND billgroupcatid<>'B13'"
    sql = sql & " ;update LabTest set BillGroupID='" & bGrp & "', BillGroupCatID='" & bGrp & "' where TestContainerID='DPT011' "
    conn.execute sql

    ApplyLabTestPullup
    
    
    bGrp = "BG999"
    bGrpName = "Admin Charges"
    AddBillGroup bGrp, bGrpName, "", ""
    sql = " update BillGroup set BillGroupCatID='" & bGrp & "', BillGroupTypeID='" & bGrp & "' where BillGroupName like '%admin%'"

    sql = sql & " ; update Treatment set BillGroupID=SpecialistType.BillGroupID, BillGroupCatID=SpecialistType.BillGroupCatID "
    sql = sql & " from Treatment inner join SpecialistType on SpecialistType.SpecialistTypeID=Treatment.TreatmentID"
    sql = sql & " ; update Treatment set BillGroupID='BG001.1', BillGroupCatID='BG001.1' where TreatmentName like '%consultation%'"
    sql = sql & " ; update Treatment set BillGroupID='BG001', BillGroupCatID='BG001' where TreatmentName like '%gp%consult%'"
    sql = sql & " ; update Treatment set BillGroupID='BG001', BillGroupCatID='BG001' where TreatmentName like '%fast%track%'"
    conn.execute sql
    
    ApplyTreatPullUp
    
    ApplyItemsPullUp
    
End Sub
Sub ApplyDrugPullup()
    Dim fld, sql, rst, bRst
    
    fld = Split("BillGroupCatID||BillGroupID", "||")
    
    Set rst = CreateObject("ADODB.RecordSet")
    Set bRst = CreateObject("ADODB.RecordSet")
    
    sql = "select   * from DrugSaleItems "
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            sql = "select * from Drug where DrugID='" & rst.fields("DrugID") & "' "
            bRst.open sql, conn, 3, 4
            If bRst.RecordCount > 0 Then
                bRst.movefirst
                For Each field In fld
                    rst.fields(field) = bRst.fields(field)
                Next
                rst.UpdateBatch
            End If
            bRst.Close
            rst.MoveNext
        Loop
        
    End If
    rst.Close
    
    sql = "select  * from DrugSaleItems2 "
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            sql = "select * from Drug where DrugID='" & rst.fields("DrugID") & "' "
            bRst.open sql, conn, 3, 4
            If bRst.RecordCount > 0 Then
                bRst.movefirst
                For Each field In fld
                    rst.fields(field) = bRst.fields(field)
                Next
                rst.UpdateBatch
            End If
            bRst.Close
            rst.MoveNext
        Loop
        
    End If
    rst.Close
    
    sql = "select   * from DrugReturnItems "
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            sql = "select * from Drug where DrugID='" & rst.fields("DrugID") & "' "
            bRst.open sql, conn, 3, 4
            If bRst.RecordCount > 0 Then
                bRst.movefirst
                For Each field In fld
                    rst.fields(field) = bRst.fields(field)
                Next
                rst.UpdateBatch
            End If
            bRst.Close
            rst.MoveNext
        Loop
        
    End If
    rst.Close
    
    sql = "select  * from DrugReturnItems2 "
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            sql = "select * from Drug where DrugID='" & rst.fields("DrugID") & "' "
            bRst.open sql, conn, 3, 4
            If bRst.RecordCount > 0 Then
                bRst.movefirst
                For Each field In fld
                    rst.fields(field) = bRst.fields(field)
                Next
                rst.UpdateBatch
            End If
            bRst.Close
            rst.MoveNext
        Loop
        
    End If
    rst.Close
    
End Sub
Sub ApplyLabTestPullup()
    Dim fld, sql, rst, bRst
    
    fld = Split("BillGroupCatID||BillGroupID||TestContainerID", "||")
    
    Set rst = CreateObject("ADODB.RecordSet")
    Set bRst = CreateObject("ADODB.RecordSet")
    
    sql = " select  * from Investigation "
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            
            sql = "select * from LabTest where LabTestID='" & rst.fields("LabTestID") & "' "
            bRst.open sql, conn, 3, 4
            If bRst.RecordCount > 0 Then
                bRst.movefirst
                For Each field In fld
                    rst.fields(field) = bRst.fields(field)
                Next
                rst.UpdateBatch
            End If
            bRst.Close
            rst.MoveNext
        Loop
    End If
    rst.Close
    
    sql = " select  * from Investigation2 "
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF

            sql = "select * from LabTest where LabTestID='" & rst.fields("LabTestID") & "' "
            bRst.open sql, conn, 3, 4
            If bRst.RecordCount > 0 Then
                
                bRst.movefirst
                For Each field In fld
                    rst.fields(field) = bRst.fields(field)
                Next
                rst.UpdateBatch
            End If
            bRst.Close
            rst.MoveNext
        Loop
    End If
    rst.Close
    
End Sub
Sub ApplySpecialistTypePullUp()
    Dim fld, sql, rst, bRst
    
    fld = Split("BillGroupID||BillGroupCatID||SpecialistClassID||SpecialistGroupID", "||")
    
    Set rst = CreateObject("ADODB.RecordSet")
    Set bRst = CreateObject("ADODB.RecordSet")
    
    sql = "select  * from Visitation "
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        'MsgBox "Visitation " & rst.RecordCount

        Do While Not rst.EOF
            sql = "select * from SpecialistType where SpecialistTypeID='" & rst.fields("SpecialistTypeID") & "' "
            bRst.open sql, conn, 3, 4
            If bRst.RecordCount > 0 Then
                bRst.movefirst
                For Each field In fld
                    rst.fields(field) = bRst.fields(field)
                Next
                rst.UpdateBatch
            End If
            bRst.Close
            
            rst.MoveNext
        Loop
        
    End If
    rst.Close
End Sub
Sub ApplyTreatPullUp()
    Dim fld, sql, rst, bRst, tRst
    
    fld = Split("BillGroupID||BillGroupCatID", "||")
    
    Set rst = CreateObject("ADODB.RecordSet")
    Set bRst = CreateObject("ADODB.RecordSet")
    Set tRst = CreateObject("ADODB.RecordSet")
    
    sql = "select  * from TreatCharges "
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            
            sql = "select * from Treatment where TreatmentID='" & rst.fields("TreatmentID") & "' "
            bRst.open sql, conn, 3, 4
            If bRst.RecordCount > 0 Then
                bRst.movefirst
                sql = "select * from BillGroup where BillGroupID='" & bRst.fields("BillGroupID") & "' "
                tRst.open sql, conn, 3, 4
                If tRst.RecordCount > 0 Then
                    tRst.movefirst
                    For Each field In fld
                        bRst.fields(field) = tRst.fields(field)
                    Next
                    bRst.UpdateBatch
                End If
                tRst.Close
                
                For Each field In fld
                    rst.fields(field) = bRst.fields(field)
                Next
                rst.UpdateBatch
            End If
            bRst.Close
            rst.MoveNext
        Loop
        
    End If
    rst.Close
End Sub
Sub AddBillGroup(bGrp, bGrpName, bGrpCat, bGrpCatName)
    Dim sql, rst
    
    Set rst = CreateObject("ADODB.RecordSet")

    If Trim(bGrpCat) = "" Then
        bGrpCat = bGrp
    End If
    
    If Trim(bGrpCatName) = "" Then
        bGrpCatName = bGrpName
    End If
    
    sql = "select  * from BillGroup where BillGroupID='" & bGrp & "' "
    rst.open sql, conn, 3, 4
    If rst.RecordCount = 0 Then
        rst.AddNew
    End If
    rst.fields("BillGroupID") = bGrp
    rst.fields("BillGroupName") = bGrpName
    rst.fields("BillGroupCatID") = bGrpCat
    rst.fields("BillGroupTypeID") = bGrpCat
    rst.UpdateBatch
    rst.Close
    
    
    sql = "select  * from BillGroupCat where BillGroupCatID='" & bGrpCat & "' "
    rst.open sql, conn, 3, 4
    If rst.RecordCount = 0 Then
        rst.AddNew
    End If
    rst.fields("BillGroupCatID") = bGrpCat
    rst.fields("BillGroupCatName") = bGrpCatName
    rst.UpdateBatch
    rst.Close
    
    sql = "select  * from BillGroupType where BillGroupTypeID='" & bGrpCat & "' "
    rst.open sql, conn, 3, 4
    If rst.RecordCount = 0 Then
        rst.AddNew
    End If
    rst.fields("BillGroupTypeID") = bGrpCat
    rst.fields("BillGroupTypeName") = bGrpCatName
    rst.UpdateBatch
    rst.Close
    
End Sub
Sub ApplyItemsPullUp()
    Dim fld, sql, rst, bRst
    
    fld = Split("BillGroupCatID||BillGroupID", "||")
    
    Set rst = CreateObject("ADODB.RecordSet")
    Set bRst = CreateObject("ADODB.RecordSet")
    
    sql = "select   * from StockIssueItems "
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            sql = "select * from Items where ItemID='" & rst.fields("ItemID") & "' "
            bRst.open sql, conn, 3, 4
            If bRst.RecordCount > 0 Then
                bRst.movefirst
                For Each field In fld
                    rst.fields(field) = bRst.fields(field)
                Next
                rst.UpdateBatch
            End If
            bRst.Close
            rst.MoveNext
        Loop
        
    End If
    rst.Close
    
    sql = "select * from StockReturnItems "
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            sql = "select * from Items where ItemID='" & rst.fields("ItemID") & "' "
            bRst.open sql, conn, 3, 4
            If bRst.RecordCount > 0 Then
                bRst.movefirst
                For Each field In fld
                    rst.fields(field) = bRst.fields(field)
                Next
                rst.UpdateBatch
            End If
            bRst.Close
            rst.MoveNext
        Loop
        
    End If
    rst.Close
    
End Sub
Sub UpdateBillGroups()

End Sub
Sub MigrateDrugStock(tbl, ky, dt)
    Dim sql, rst, stk, stkItm, itm, rst2
    
    
    xlsConn.execute " update " & tbl & " set rec_count=0, status=''"
    
    sql = "select * from " & tbl & " where qty is not null "
    Set rst = CreateObject("ADODB.RecordSet")
    Set rst2 = CreateObject("ADODB.RecordSet")
    
    rst.open sql, xlsConn, 3, 4

    If rst.RecordCount > 0 Then
        rst.movefirst
        
        Set stk = CreateObject("ADODB.RecordSet")
        Set stkItm = CreateObject("ADODB.RecordSet")
        Set itm = CreateObject("ADODB.RecordSet")

        sql = "select * from DrugAdjustment where DrugAdjustmentID='" & ky & "'"
        stk.open sql, conn, 3, 4
        If stk.RecordCount = 0 Then
            stk.AddNew
        End If
        If True Then
            stk.fields("AdjustDate") = dt
            stk.fields("BranchID") = "B001"
            stk.fields("EntryDate") = stk.fields("AdjustDate")
'            stk.fields("EntryInfo") = ""
            stk.fields("EntryValue") = 0
            stk.fields("DrugAdjustStatusID") = "D001"
            stk.fields("DrugStoreID") = "S22"
            stk.fields("JobScheduleID") = "S22"
'            stk.fields("KeyPrefix") = ""
            stk.fields("MainDate1") = stk.fields("AdjustDate")
            stk.fields("MainDate2") = stk.fields("AdjustDate")
'            stk.fields("MainInfo1") = ""
'            stk.fields("MainInfo2") = ""
            stk.fields("MainValue1") = 0
            stk.fields("MainValue2") = 0
            
            stk.fields("PostTransactionID") = "P002"
            stk.fields("DrugAdjustmentID") = ky
            stk.fields("DrugAdjustmentName") = GetComboName("DrugStore", stk.fields("DrugStoreID")) & "[" & FormatDateDetail(stk.fields("AdjustDate")) & "]"
            stk.fields("DrugAdjustPointID") = "NONE"
            stk.fields("DrugAdjustTypeID") = "D001"
            stk.fields("SystemUserID") = "S22"
            stk.fields("TransProcessStatID") = "T001"
            stk.fields("TransProcessValID") = "DrugAdjustmentPro-T001"
            stk.fields("WorkingDayID") = FormatWorkingDay(stk.fields("AdjustDate"))
            stk.fields("WorkingMonthID") = FormatWorkingMonth(stk.fields("AdjustDate"))
            stk.fields("WorkingYearID") = FormatWorkingYear(stk.fields("AdjustDate"))

            stk.UpdateBatch
        End If
        
        Do While Not rst.EOF
            
            If Len(rst.fields("ID")) > 0 Then
                sql = "select * from Drug where DrugID='" & rst.fields("ID") & "' "
            ElseIf Len(rst.fields("id_1")) > 0 Then
                Set rst2 = CreateObject("ADODB.RecordSet")
                sql = "select * from f_Itemmaster where ID='" & rst.fields("id_1") & "' "
                rst2.open sql, GetMySQLConnection, 3, 4
                If rst2.RecordCount > 0 Then
                    sql = "select * from Drug where DrugID='" & rst2.fields("ItemID") & "' "
                    rst.fields("ID") = rst2.fields("ItemID")
                End If
            Else
                sql = "select * from Drug where DrugID='' "
            End If
            
            itm.open sql, conn, 3, 4
            
            If itm.RecordCount > 0 Then
                AddDrugInventory itm
                sql = "select * from DrugAdjustItems where DrugAdjustmentID='" & ky & "' and DrugID='" & rst.fields("ID") & "' "
                
                stkItm.open sql, conn, 3, 4
                If stkItm.RecordCount = 0 Then
                    stkItm.AddNew
                End If
               
               'MsgBox "here"
                avqty = getAvQty("Drug", itm.fields("DrugID"), stk.fields("AdjustDate"))
                
                stkItm.fields("AdjustDate") = stk.fields("AdjustDate")
                stkItm.fields("AdjustDate1") = stk.fields("AdjustDate")
                stkItm.fields("AdjustDate2") = stk.fields("AdjustDate")
                stkItm.fields("AdjustmentDate1") = stk.fields("AdjustDate")
                stkItm.fields("AdjustmentDate2") = stk.fields("AdjustDate")
    '            stkItm.fields("AdjustmentInfo1") = ""
    '            stkItm.fields("AdjustmentInfo2") = ""
                stkItm.fields("AdjustmentValue1") = rst.fields("Qty")
    '            stkItm.fields("AdjustmentValue2") = 0
    '            stkItm.fields("AdjustmentValue3") = 0
    '            stkItm.fields("AdjustmentValue4") = 0
                stkItm.fields("AdjustValue1") = (-1 * avqty) + stkItm.fields("AdjustmentValue1") 'Formula**Expression**-1||*||AvailableQty||+||AdjustmentValue1
                stkItm.fields("AdjustValue2") = 0
                stkItm.fields("AdjustValue3") = 0
                
                If stkItm.fields("AdjustValue1") > 0 Then
                    stkItm.fields("AdjustValue2") = stkItm.fields("AdjustValue1")
                Else
                    stkItm.fields("AdjustValue3") = -1 * stkItm.fields("AdjustValue1")
                End If
                stkItm.fields("AdjustValue3") = 0
                stkItm.fields("AfterAcceptQty") = 0
                stkItm.fields("AvailableQty") = avqty
                stkItm.fields("BranchID") = stk.fields("BranchID")
                stkItm.fields("BulkUnitCost") = itm.fields("BulkUnitCost")
                stkItm.fields("EntryDate") = stk.fields("AdjustDate")
    '            stkItm.fields("EntryInfo") = ""
                stkItm.fields("EntryValue") = 0
    '            stkItm.fields("ExpiryDate") = ""
                stkItm.fields("FinalAmt") = itm.fields("BulkUnitCost") * rst.fields("QTY")
                stkItm.fields("DrugAdjustStatusID") = stk.fields("DrugAdjustStatusID")
                stkItm.fields("DrugCategoryID") = itm.fields("DrugCategoryID")
                stkItm.fields("DrugID") = itm.fields("DrugID")
                stkItm.fields("DrugStoreID") = stk.fields("DrugStoreID")
                stkItm.fields("DrugTypeID") = itm.fields("DrugTypeID")
                stkItm.fields("JobScheduleID") = stk.fields("JobScheduleID")
                stkItm.fields("MainDate1") = stk.fields("AdjustDate")
    '            stkItm.fields("MainDate2") = ""
    '            stkItm.fields("MainInfo1") = ""
    '            stkItm.fields("MainInfo2") = ""
                stkItm.fields("MainItemDate1") = stkItm.fields("MainDate1")
                
    '            stkItm.fields("MainItemInfo1") = ""
                stkItm.fields("MainItemValue1") = 0
                stkItm.fields("MainItemValue2") = 0
                stkItm.fields("MainItemValue3") = 0
                stkItm.fields("MainItemValue4") = 0
                stkItm.fields("MainValue1") = 0
                stkItm.fields("MainValue2") = 0
                stkItm.fields("PendingAcceptQty") = 0
                stkItm.fields("PostTransactionID") = stk.fields("PostTransactionID")
                stkItm.fields("Qty") = rst.fields("Qty")
                stkItm.fields("QtyBeforeReorder") = 0
                stkItm.fields("RetailUnitCost") = itm.fields("RetailUnitCost")
                stkItm.fields("DrugAdjustmentID") = stk.fields("DrugAdjustmentID")
                stkItm.fields("DrugAdjustPointID") = stk.fields("DrugAdjustPointID")
                stkItm.fields("DrugAdjustTypeID") = stk.fields("DrugAdjustTypeID")
    '            stkItm.fields("StockDate1") = ""
    '            stkItm.fields("StockDate2") = ""
    '            stkItm.fields("StockInfo1") = ""
    '            stkItm.fields("StockInfo2") = ""
    '            stkItm.fields("StockValue1") = ""
    '            stkItm.fields("StockValue2") = ""
                stkItm.fields("SystemUserID") = stk.fields("SystemUserID")
                stkItm.fields("TotalCost") = itm.fields("BulkUnitCost") * rst.fields("QTY")
                stkItm.fields("TransProcessStatID") = stk.fields("TransProcessStatID")
                stkItm.fields("TransProcessValID") = stk.fields("TransProcessValID")
                stkItm.fields("UnitOfMeasureID") = itm.fields("UnitOfMeasureID")
                stkItm.fields("WorkingDayID") = stk.fields("WorkingDayID")
                stkItm.fields("WorkingMonthID") = stk.fields("WorkingMonthID")
                stkItm.fields("WorkingYearID") = stk.fields("WorkingYearID")
    
                stkItm.UpdateBatch
                stkItm.Close
                
                rst.fields("status") = "DONE"
                
                rst.fields("rec_count") = rst.fields("rec_count") + 1
            End If
            itm.Close
            rst.UpdateBatch
            rst.MoveNext
        Loop
    End If
    rst.Close
End Sub
Function getAvQty(tbl, ky, dt)
    Dim sql, rst, acc
    Dim ot
    ot = 0
    acc = "DRG-" & ky
    sql = "select sum(balanceQty) as balanceQty from Inventtransentry where inventaccountid='" & acc & "' and EntryDate<'" & dt & "' group by inventaccountid"
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        
        If Not IsNull(rst.fields("balanceQty")) Then ot = rst.fields("balanceQty")
    End If
    'If ot <> 0 Then MsgBox ot
    getAvQty = ot
End Function
Sub MigrateItemStock()
    Dim sql, rst, stk, stkItm, ky, itm
    
    ky = "ITM-S00001"
    
    sql = "select * from consumables$ where qty is not null "
    Set rst = CreateObject("ADODB.RecordSet")
    
    
    rst.open sql, xlsConn, 3, 4

    If rst.RecordCount > 0 Then
        rst.movefirst
        
        Set stk = CreateObject("ADODB.RecordSet")
        Set stkItm = CreateObject("ADODB.RecordSet")
        Set itm = CreateObject("ADODB.RecordSet")

        sql = "select * from StockAdjustment where StockAdjustmentID='" & ky & "'"
        stk.open sql, conn, 3, 4
        If stk.RecordCount = 0 Then
            stk.AddNew

            stk.fields("AdjustDate") = "01 Oct 2022 00:00:00"
            stk.fields("BranchID") = "B001"
            stk.fields("EntryDate") = stk.fields("AdjustDate")
'            stk.fields("EntryInfo") = ""
            stk.fields("EntryValue") = 0
            stk.fields("ItemAdjustStatusID") = "I001"
            stk.fields("ItemStoreID") = "S22"
            stk.fields("JobScheduleID") = "S22"
'            stk.fields("KeyPrefix") = ""
            stk.fields("MainDate1") = stk.fields("AdjustDate")
            stk.fields("MainDate2") = stk.fields("AdjustDate")
'            stk.fields("MainInfo1") = ""
'            stk.fields("MainInfo2") = ""
            stk.fields("MainValue1") = 0
            stk.fields("MainValue2") = 0
            stk.fields("PostTransactionID") = "P002"
            stk.fields("StockAdjustmentID") = ky
            stk.fields("StockAdjustmentName") = GetComboName("ItemStore", stk.fields("ItemStoreID")) & "[" & FormatDateDetail(stk.fields("AdjustDate")) & "]"
            stk.fields("StockAdjustPointID") = "NONE"
            stk.fields("StockAdjustTypeID") = "S001"
            stk.fields("SystemUserID") = "S22"
            stk.fields("TransProcessStatID") = "T001"
            stk.fields("TransProcessValID") = "StockAdjustmentPro-T001"
            stk.fields("WorkingDayID") = FormatWorkingDay(stk.fields("AdjustDate"))
            stk.fields("WorkingMonthID") = FormatWorkingMonth(stk.fields("AdjustDate"))
            stk.fields("WorkingYearID") = FormatWorkingYear(stk.fields("AdjustDate"))

            stk.UpdateBatch
        End If
        
        Do While Not rst.EOF
        
            sql = "select * from Items where ItemID='" & rst.fields("ID") & "' "
            itm.open sql, conn, 3, 4
            
            If itm.RecordCount > 0 Then
                sql = "select * from StockAdjustItems where StockAdjustmentID='" & ky & "' and ItemID='" & rst.fields("ID") & "' "
                
                stkItm.open sql, conn, 3, 4
                If stkItm.RecordCount = 0 Then
                    stkItm.AddNew
                End If
    
                stkItm.fields("AdjustDate") = stk.fields("AdjustDate")
                stkItm.fields("AdjustDate1") = stk.fields("AdjustDate")
                stkItm.fields("AdjustDate2") = stk.fields("AdjustDate")
                stkItm.fields("AdjustmentDate1") = stk.fields("AdjustDate")
                stkItm.fields("AdjustmentDate2") = stk.fields("AdjustDate")
    '            stkItm.fields("AdjustmentInfo1") = ""
    '            stkItm.fields("AdjustmentInfo2") = ""
                stkItm.fields("AdjustmentValue1") = rst.fields("Qty")
    '            stkItm.fields("AdjustmentValue2") = 0
    '            stkItm.fields("AdjustmentValue3") = 0
    '            stkItm.fields("AdjustmentValue4") = 0
                stkItm.fields("AdjustValue1") = (-1 * 0) + stkItm.fields("AdjustmentValue1") 'Formula**Expression**-1||*||AvailableQty||+||AdjustmentValue1
                stkItm.fields("AdjustValue2") = 0
                stkItm.fields("AdjustValue3") = 0
                stkItm.fields("AfterAcceptQty") = 0
                stkItm.fields("AvailableQty") = rst.fields("QTY")
                stkItm.fields("BranchID") = stk.fields("BranchID")
                stkItm.fields("BulkUnitCost") = itm.fields("BulkUnitCost")
                stkItm.fields("EntryDate") = stk.fields("AdjustDate")
    '            stkItm.fields("EntryInfo") = ""
                stkItm.fields("EntryValue") = 0
    '            stkItm.fields("ExpiryDate") = ""
                stkItm.fields("FinalAmt") = itm.fields("BulkUnitCost") * rst.fields("QTY")
                stkItm.fields("ItemAdjustStatusID") = stk.fields("ItemAdjustStatusID")
                stkItm.fields("ItemCategoryID") = itm.fields("ItemCategoryID")
                stkItm.fields("ItemID") = itm.fields("ItemID")
                stkItm.fields("ItemStoreID") = stk.fields("ItemStoreID")
                stkItm.fields("ItemTypeID") = itm.fields("ItemTypeID")
                stkItm.fields("JobScheduleID") = "S21A"
                stkItm.fields("MainDate1") = stk.fields("AdjustDate")
    '            stkItm.fields("MainDate2") = ""
    '            stkItm.fields("MainInfo1") = ""
    '            stkItm.fields("MainInfo2") = ""
                stkItm.fields("MainItemDate1") = stkItm.fields("MainDate1")
                
    '            stkItm.fields("MainItemInfo1") = ""
                stkItm.fields("MainItemValue1") = 0
                stkItm.fields("MainItemValue2") = 0
                stkItm.fields("MainItemValue3") = 0
                stkItm.fields("MainItemValue4") = 0
                stkItm.fields("MainValue1") = 0
                stkItm.fields("MainValue2") = 0
                stkItm.fields("PendingAcceptQty") = 0
                stkItm.fields("PostTransactionID") = stk.fields("PostTransactionID")
                stkItm.fields("Qty") = rst.fields("Qty")
                stkItm.fields("QtyBeforeReorder") = 0
                stkItm.fields("RetailUnitCost") = itm.fields("RetailUnitCost")
                stkItm.fields("StockAdjustmentID") = stk.fields("StockAdjustmentID")
                stkItm.fields("StockAdjustPointID") = stk.fields("StockAdjustPointID")
                stkItm.fields("StockAdjustTypeID") = stk.fields("StockAdjustTypeID")
    '            stkItm.fields("StockDate1") = ""
    '            stkItm.fields("StockDate2") = ""
    '            stkItm.fields("StockInfo1") = ""
    '            stkItm.fields("StockInfo2") = ""
    '            stkItm.fields("StockValue1") = ""
    '            stkItm.fields("StockValue2") = ""
                stkItm.fields("SystemUserID") = stk.fields("SystemUserID")
                stkItm.fields("TotalCost") = itm.fields("BulkUnitCost") * rst.fields("QTY")
                stkItm.fields("TransProcessStatID") = stk.fields("TransProcessStatID")
                stkItm.fields("TransProcessValID") = stk.fields("TransProcessValID")
                stkItm.fields("UnitOfMeasureID") = itm.fields("UnitOfMeasureID")
                stkItm.fields("WorkingDayID") = stk.fields("WorkingDayID")
                stkItm.fields("WorkingMonthID") = stk.fields("WorkingMonthID")
                stkItm.fields("WorkingYearID") = stk.fields("WorkingYearID")
    
                stkItm.UpdateBatch
                stkItm.Close
            End If
            itm.Close
            
            rst.MoveNext
        Loop
    End If
    rst.Close
End Sub
Function xlsConn()
    Dim tmp
    Set xlsConn = CreateObject("ADODB.Connection")
    xlsConn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=admin;Password=drowssaprofas;Initial Catalog=FOHData;Data Source=172.2.2.32"
    xlsConn.cursorlocation = 3

    xlsConn.open

End Function
Sub AddAppointTimes()
    Dim sql, rst, idx, nm, hr, min
    
    Set rst = CreateObject("ADODB.RecordSet")
    
    i = 6
    For i = 6 To 21
        For j = 0 To 60 Step 5
            If j < 60 Then
                hr = Right((10000 + i), 2)
                min = Right((10000 + j), 2)
                
                idx = hr & min
                
                If hr <> 12 Then
                    nm = (hr Mod 12) & ":" & min
                Else
                    nm = hr & ":" & min
                End If
                If CInt(hr) >= 12 Then
                    nm = nm & " PM"
                Else
                    nm = nm & " AM"
                End If
                
                sql = "select * from AppointStartTime where AppointStartTimeID='" & idx & "'"
                rst.open sql, conn, 3, 4
                If rst.RecordCount = 0 Then
                    rst.AddNew
                End If
                If True Then
                    rst.fields("AppointStartTimeID") = idx
                    rst.fields("AppointStartTimeName") = nm
                    rst.fields("AppointStartTime") = CDate("01 Jan 2011 " & nm)
                    
                    rst.UpdateBatch
                End If
                rst.Close
                
                sql = "select * from AppointEndTime where AppointEndTimeID='" & idx & "'"
                rst.open sql, conn, 3, 4
                If rst.RecordCount = 0 Then
                    rst.AddNew
                End If
                If True Then
                    rst.fields("AppointEndTimeID") = idx
                    rst.fields("AppointEndTimeName") = nm
                    rst.fields("AppointEndTime") = CDate("01 Jan 2011 " & nm)
                    
                    rst.UpdateBatch
                End If
                rst.Close
            End If
        Next
    Next
        
End Sub
Sub CopyDrugPriceToIns(fromIns, toIns)
    Dim sql, frmRst, toRst, field
    
    
    sql = "select  * from DrugPriceMatrix2 where InsuranceTypeID='" & fromIns & "'"
    Set frmRst = CreateObject("ADODB.RecordSet")
    Set toRst = CreateObject("ADODB.RecordSet")
    frmRst.open sql, conn, 3, 4
    If frmRst.RecordCount > 0 Then
        frmRst.movefirst
        Do While Not frmRst.EOF
            sql = "select * from DrugPriceMatrix2 where InsuranceTypeID='" & toIns & "' and DrugID='" & frmRst.fields("DrugID") & "' and DrugStoreTypeID='" & frmRst.fields("DrugStoreTypeID") & "'"
            toRst.open sql, conn, 3, 4
            
            If toRst.RecordCount = 0 Then
                toRst.AddNew
            End If
            
            For Each field In frmRst.fields
                toRst.fields(field.name) = field.value
            Next
            
            toRst.fields("InsuranceTypeID") = toIns
            toRst.UpdateBatch
            toRst.Close
            frmRst.MoveNext
        Loop
    End If
    frmRst.Close
End Sub
Sub CopyInvPriceToIns(fromIns, toIns)
    Dim sql, frmRst, toRst, field
    
    sql = "select * from LabTestCostMatrix where InsuranceTypeID='" & fromIns & "'"
    Set frmRst = CreateObject("ADODB.RecordSet")
    Set toRst = CreateObject("ADODB.RecordSet")
    frmRst.open sql, conn, 3, 4
    If frmRst.RecordCount > 0 Then
        frmRst.movefirst
        Do While Not frmRst.EOF
            sql = "select * from LabTestCostMatrix where InsuranceTypeID='" & toIns & "' and AgeGroupID='" & frmRst.fields("AgeGroupID") & "' and LabTestID='" & frmRst.fields("LabTestID") & "'"
            toRst.open sql, conn, 3, 4
            If toRst.RecordCount = 0 Then
                toRst.AddNew
            End If
            For Each field In frmRst.fields
                toRst.fields(field.name) = field.value
            Next
            toRst.fields("InsuranceTypeID") = toIns
            toRst.UpdateBatch
            toRst.Close
            frmRst.MoveNext
        Loop
    End If
    frmRst.Close
End Sub
Sub CopyItemPriceToIns(fromIns, toIns)
    Dim sql, frmRst, toRst, field
    
    sql = "select * from ItemPriceMatrix2 where InsuranceTypeID='" & fromIns & "'"
    Set frmRst = CreateObject("ADODB.RecordSet")
    Set toRst = CreateObject("ADODB.RecordSet")
    frmRst.open sql, conn, 3, 4
    If frmRst.RecordCount > 0 Then
        frmRst.movefirst
        Do While Not frmRst.EOF
            sql = "select * from ItemPriceMatrix2 where InsuranceTypeID='" & toIns & "' and ItemID='" & frmRst.fields("ItemID") & "' and ItemStoreTypeID='" & frmRst.fields("ItemStoreTypeID") & "' and RecipientTypeID='" & frmRst.fields("RecipientTypeID") & "' "
            toRst.open sql, conn, 3, 4
            If toRst.RecordCount = 0 Then
                toRst.AddNew
            End If
            For Each field In frmRst.fields
                toRst.fields(field.name) = field.value
            Next
            toRst.fields("InsuranceTypeID") = toIns
            toRst.UpdateBatch
            toRst.Close
            frmRst.MoveNext
        Loop
    End If
    frmRst.Close
End Sub
Sub CopyBedPriceToIns(fromIns, toIns)
    Dim sql, frmRst, toRst, field
    
    sql = "select * from BedCostMatrix where InsuranceTypeID='" & fromIns & "'"
    Set frmRst = CreateObject("ADODB.RecordSet")
    Set toRst = CreateObject("ADODB.RecordSet")
    frmRst.open sql, conn, 3, 4
    If frmRst.RecordCount > 0 Then
        frmRst.movefirst
        Do While Not frmRst.EOF
            sql = "select * from BedCostMatrix where InsuranceTypeID='" & toIns & "' and BedID='" & frmRst.fields("BedID") & "' and AgeGroupID='" & frmRst.fields("AgeGroupID") & "' and AdmissionTypeID='" & frmRst.fields("AdmissionTypeID") & "'"
            toRst.open sql, conn, 3, 4
            If toRst.RecordCount = 0 Then
                toRst.AddNew
            End If
            For Each field In frmRst.fields
                toRst.fields(field.name) = field.value
            Next
            toRst.fields("InsuranceTypeID") = toIns
            toRst.UpdateBatch
            toRst.Close
            frmRst.MoveNext
        Loop
    End If
    frmRst.Close
End Sub
Sub CopyTreatPriceToIns(fromIns, toIns)
    Dim sql, frmRst, toRst, field
    
    sql = "select * from TreatCostMatrix where InsuranceTypeID='" & fromIns & "'"
    Set frmRst = CreateObject("ADODB.RecordSet")
    Set toRst = CreateObject("ADODB.RecordSet")
    frmRst.open sql, conn, 3, 4
    If frmRst.RecordCount > 0 Then
        frmRst.movefirst
        Do While Not frmRst.EOF
            sql = "select * from TreatCostMatrix where InsuranceTypeID='" & toIns & "' and MedicalServiceID='" & frmRst.fields("MedicalServiceID") & "' and AgeGroupID='" & frmRst.fields("AgeGroupID") & "' and TreatmentID='" & frmRst.fields("TreatmentID") & "'"
            toRst.open sql, conn, 3, 4
            If toRst.RecordCount = 0 Then
                toRst.AddNew
            End If
            For Each field In frmRst.fields
                toRst.fields(field.name) = field.value
            Next
            toRst.fields("InsuranceTypeID") = toIns
            toRst.UpdateBatch
            toRst.Close
            frmRst.MoveNext
        Loop
    End If
    frmRst.Close
End Sub




'Server.ScriptTimeOut = 60 * 60 * 60
'Call MigrateData

Function GetMySQLConnection()
    Dim mySqlConn, mySqlConnStr, svrIp, srcRst

    svrIp = "localhost"
    Set mySqlConn = CreateObject("ADODB.Connection")

    mySqlConnStr = "DSN=FOCOS;Uid=admin;Pwd=drowssaprofas;"
    'mySqlConnStr = "DSN=FOCOS_LIVE;Uid=admin;Pwd=drowssaprofas;"

    mySqlConn.cursorlocation = 3

    mySqlConn.open mySqlConnStr
    Set GetMySQLConnection = mySqlConn
End Function
Sub MigrateData()
    Dim sql
    ''add missing records
    'Call AddMissingRecords
'    Call MigrateCountry
'    Call MigrateWard
'    Call MigrateDisease
'
'    Call MigrateProcedures
'    Call MigrateInvestigations
'    Call MigrateMedicalConsumables
'    Call MigrateSurgery
'    Call MigrateItems
'    Call MigrateStaff
'    Call MigrateSponsor
'
'    sql = "update SystemSchedule set ScheduleStatusID='S001' where SystemScheduleID='Migrate00'"
'    conn.execute sql
    Call AddConsultations
End Sub
Function IIF(expression, trueVal, falseVal)
    If expression = True Then
        IIF = trueVal
    Else
        IIF = falseVal
    End If
End Function
Sub AddMissingRecords()
    Dim sql

    sql = "if not exists(select * from MaritalStatus where MaritalStatusID='M006')"
    sql = sql & " insert into MaritalStatus(MaritalStatusID, MaritalStatusName)"
    sql = sql & " values ('M006', 'Unmarried') "
    conn.execute qryPro.FltQry(sql)

End Sub
Sub AddConsultations()
    Dim sql, rst, tmpRst, matRst, vstat

    'consultation group
    sql = " if not exists(select * from SpecialistGroup where SpecialistGroupID='LSHHI1') "
    sql = sql & " insert into SpecialistGroup(SpecialistGroupID, SpecialistGroupName)"
    sql = sql & " values('LSHHI1', 'OPD CONSULTATIONS VISIT') "
    conn.execute qryPro.FltQry(sql)

    'consultation type
    sql = " select * from f_subcategorymaster where f_subcategorymaster.CategoryID='LSHHI1' "

    sql = " select * "
    sql = sql & " from (select f_subcategorymaster.SubCategoryID, f_subcategorymaster.Name AS sub_category_name "
    sql = sql & "     , f_ratelist.Rate  "
    sql = sql & "     , f_ratelist.EntryDate"
    sql = sql & "     , f_ratelist.Panel_ID "
    sql = sql & "     , ROW_NUMBER () over(PARTITION BY f_subcategorymaster.SubCategoryID,  f_ratelist.Panel_ID order by f_ratelist.EntryDate desc) AS row_num "
    sql = sql & " FROM f_itemmaster  "
    sql = sql & " left join f_subcategorymaster ON f_subcategorymaster.SubCategoryID=f_itemmaster.SubCategoryID "
    sql = sql & " left join f_ratelist ON f_ratelist.ItemID=f_itemmaster.ItemID "
    sql = sql & " left join f_ratelist AS cmp ON cmp.itemid=f_ratelist.itemid  and cmp.entrydate>f_ratelist.entrydate "
    sql = sql & " where 1=1 "
    sql = sql & "     and f_subcategorymaster.CategoryID='LSHHI1' "
    sql = sql & "  and cmp.ratelistid IS NULL and f_ratelist.rate IS NOT NULL "
    'sql = sql & "  and f_ratelist.entrydate >= '2021-01-01 00:00:00' "
    sql = sql & " ) AS tbl "
    sql = sql & " where row_num = 1"

    Set rst = CreateObject("ADODB.RecordSet")
    Set matRst = CreateObject("ADODB.RecordSet")
    Set tmpRst = CreateObject("ADODB.RecordSet")
    rst.open sql, GetMySQLConnection(), 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            sql = "select * from SpecialistType where SpecialistTypeID='" & rst.fields("SubCategoryID") & "' "
            tmpRst.open qryPro.FltQry(sql), conn, 3, 4
            If tmpRst.RecordCount = 0 Then
                tmpRst.AddNew
                tmpRst.fields("SpecialistTypeID") = rst.fields("SubCategoryID")
                tmpRst.fields("SpecialistTypeName") = rst.fields("sub_category_name")
                tmpRst.fields("SpecialistGroupID") = "LSHHI1"
                tmpRst.fields("SpecialistClassID") = "S001"
                tmpRst.fields("BillGroupCatID") = "B14"
                tmpRst.fields("BillGroupID") = "OPD"
                tmpRst.fields("LevelOfAccessID") = "All"
                tmpRst.fields("ConsultPercent") = 0
                tmpRst.fields("TreatPercent") = 0
                tmpRst.fields("DiagnosePercent") = 0
                tmpRst.fields("SpecialistTypeVal1") = 0
                tmpRst.fields("SpecialistTypeVal2") = 0
                tmpRst.fields("SpecialistTypeVal3") = 0
                tmpRst.fields("SpecialistTypeVal4") = 0
                tmpRst.fields("SpecialistTypeInfo1") = ""
                tmpRst.fields("SpecialistTypeInfo2") = ""
                'tmpRst.Fields("SpecialistTypeDate1") = ""
                'tmpRst.Fields("SpecialistTypeDate2") = ""
                tmpRst.fields("KeyPrefix") = ""
                tmpRst.fields("BillInput1") = 0
                tmpRst.fields("BillInput2") = 0
                tmpRst.fields("BillInput3") = 0
                tmpRst.fields("BillInout4") = 0
            
                tmpRst.UpdateBatch
            End If
            tmpRst.Close
            UpdateConsultMatrixFromTemplate rst.fields("SubCategoryID"), rst.fields("Panel_ID"), rst.fields("Rate")
            rst.MoveNext
        Loop
        rst.Close
    End If
    Set rst = Nothing

End Sub
Sub UpdateConsultMatrixFromTemplate(pk, insType, cst)
    Dim html, sql, rstAge, rstVstTyp, vstTyp, agegrp, rstIns, target, isValid
    Dim errorFields, matSQL, tmpSQL, insTypeVal, ageGrpVal, VstTypVal
    
    isValid = True
    
    Set rstVstTyp = CreateObject("ADODB.RecordSet")
    Set rstAge = CreateObject("ADODB.RecordSet")
    Set rstIns = CreateObject("ADODB.RecordSet")
    
    sql = "select AgeGroupID, AgeGroupName, ConsultPercent from AgeGroup"
    rstAge.open qryPro.FltQry(sql), conn, 3, 4
    
    sql = "select VisitTypeName, VisitTypeID, VisitCost from VisitType"
    rstVstTyp.open qryPro.FltQry(sql), conn, 3, 4
    
    sql = " select InsuranceTypeID, InsuranceTypeName, ConsultPercent from InsuranceType where InsuranceTypeName<>'-' "
    rstIns.open qryPro.FltQry(sql), conn, 3, 4
    
    If rstIns.RecordCount > 0 Then
        rstIns.movefirst
        Do While Not rstIns.EOF
            If rstAge.RecordCount > 0 Then
                rstAge.movefirst
                Do While Not rstAge.EOF
                    If rstVstTyp.RecordCount > 0 Then
                        rstVstTyp.movefirst
                        Do While Not rstVstTyp.EOF
                            insType = rstIns.fields("InsuranceTypeID")
                            agegrp = rstAge.fields("AgeGroupID")
                            vstTyp = rstVstTyp.fields("VisitTypeID")
                            
                            'target = pk & "||" & insType & "||" & ageGrp & "||" & vstTyp
                            'target = pk & "||" & insType & "||" & "-" & "||" & vstTyp
                            'If Request(target).Count <> 0 Then 'field exists
                                'cst = Trim(Request(target))
                                
                                If Not IsValidNumber(cst) Then
                                    isValid = isValid And False
                                End If
                                MsgBox isValid
                                If isValid Then
                                    cst = CDbl(cst)
                                    vstCst = GetComboNameFld("SpecialistType", pk, "ConsultPercent")
                                    ageGrpVal = rstAge.fields("ConsultPercent")
                                    VstTypVal = rstVstTyp.fields("VisitCost")
                                    insTypeVal = rstIns.fields("ConsultPercent")
                                    
                                    tmpSQL = vbCrLf & " if exists(select * from ConsultCostMatrix where SpecialistTypeID='" & pk & "' "
                                    tmpSQL = tmpSQL & "   and InsuranceTypeID='" & insType & "' and AgeGroupID='" & agegrp & "' and VisitTypeID='" & vstTyp & "' "
                                    tmpSQL = tmpSQL & " )"
                                    tmpSQL = tmpSQL & "   update ConsultCostMatrix set VisitCost=" & cst & " where SpecialistTypeID='" & pk & "' "
                                    tmpSQL = tmpSQL & "   and InsuranceTypeID='" & insType & "' and AgeGroupID='" & agegrp & "' and VisitTypeID='" & vstTyp & "' "
                                    tmpSQL = tmpSQL & " "
                                    tmpSQL = tmpSQL & vbCrLf & " else "
                                    tmpSQL = tmpSQL & "   insert into ConsultCostMatrix(SpecialistTypeID, InsuranceTypeID, AgeGroupID, VisitTypeID, PermuteUpdateID, PermuteStatusID, VisitCost, SpecialistTypeVal, AgeGroupVal, InsuranceTypeVal, VisitTypeVal)"
                                    tmpSQL = tmpSQL & "   values('" & pk & "', '" & insType & "',  '" & agegrp & "', '" & vstTyp & "', 'P002', 'P001', " & cst & ", " & vstCst & ", " & ageGrpVal & ", " & insTypeVal & ", " & VstTypVal & "); "
                                    
                                    matSQL = matSQL & tmpSQL
                                End If
                            'End If
                            rstVstTyp.MoveNext
                        Loop
                    End If
                    rstAge.MoveNext
                Loop
            End If
            rstIns.MoveNext
        Loop
    End If
    If isValid Then
        conn.execute qryPro.FltQry(matSQL)
    End If
End Sub
Sub MigrateStaff()
    Dim srcSql, srcRst, lastID, updDt

     updDt = GetSystemVariableValue("focosPayEmployeeLastUpdateDate")
    If Not IsDate(updDt) Then
       updDt = CDate("01 Jan 1990")
    End If

    lastID = GetSystemVariableValue("focosPayEmployeeLastID")
    If lastID = "" Then
        lastID = 0
    End If

    srcSql = "select Employee_Master.Employee_ID,  Employee_Master.PayRollEmployeeID "
    srcSql = srcSql & " , Employee_Master.Title, Employee_Master.Name, Employee_Master.IsActive, Pay_Employee_Master.Gender"
    srcSql = srcSql & " , Pay_Employee_Master.Dept_Name, Pay_Employee_Master.IsActive, Pay_Employee_Master.Phone, Pay_Employee_Master.Mobile "
    srcSql = srcSql & " , Pay_Employee_Master.Email, Pay_Employee_Master.House_No, Pay_Employee_Master.Street_Name, Pay_Employee_Master.Locality"
    srcSql = srcSql & " , Pay_Employee_Master.City, Employee_Master.Updatedate, Employee_Master.ID"
    srcSql = srcSql & " , Pay_Employee_Master.Desi_Name "
    srcSql = srcSql & " from Employee_Master "
    srcSql = srcSql & " left join  Pay_Employee_Master on Employee_Master.PayrollEmployeeID=Pay_Employee_Master.Employee_ID"
    srcSql = srcSql & " where Employee_Master.ID>" & lastID & " or Employee_Master.Updatedate>=cast('" & updDt & "' as datetime)"
    'srcSql = srcSql & " limit 10"
    srcSql = srcSql & " order by Employee_Master.ID asc"
    Set srcRst = CreateObject("ADODB.RecordSet")
    srcRst.open srcSql, GetMySQLConnection(), 3, 4
    If srcRst.RecordCount > 0 Then
        srcRst.movefirst
        Do While Not srcRst.EOF
            AddInitialStaff srcRst.fields
            AddStaff srcRst.fields
            AddInitialSystemUser srcRst.fields
            AddSystemUser srcRst.fields

            Call SetSystemVariableValue("focosPayEmployeeLastUpdateDate", srcRst.fields("Updatedate"))
            Call SetSystemVariableValue("focosPayEmployeeLastID", srcRst.fields("ID"))

            srcRst.MoveNext

        Loop
        srcRst.Close
    End If
    Set srcRst = Nothing

End Sub
Sub AddInitialSystemUser(usrFields)
    Dim sql, rst

    sql = "select * from InitialSystemUser where InitialSystemUserID='" & usrFields("Employee_ID") & "' "
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount = 0 Then
        rst.AddNew
        rst.fields("InitialSystemUserID") = usrFields("Employee_ID")
        rst.fields("InitialSystemUserName") = usrFields("Title") & " " & usrFields("Name")
        rst.UpdateBatch
    End If
    rst.Close
    Set rst = Nothing
End Sub
Sub AddInitialStaff(usrFields)
    Dim sql, rst

    sql = "select * from InitialStaff where InitialStaffID='" & usrFields("Employee_ID") & "' "
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount = 0 Then
        rst.AddNew
        rst.fields("InitialStaffID") = usrFields("Employee_ID")
        rst.fields("InitialStaffName") = usrFields("Title") & " " & usrFields("Name")
        rst.UpdateBatch
    End If
    rst.Close
    Set rst = Nothing
End Sub
Sub AddStaff(usrFields)
    Dim sql, rst, staffID

    If Len(usrFields("PayrollEmployeeID")) > 0 Then
        staffID = usrFields("PayrollEmployeeID")
        sql = " select * from Staff where StaffID='" & usrFields("PayrollEmployeeID") & "'  "
    Else
        sql = " select * from Staff where StaffID='" & usrFields("Employee_ID") & "'  "
        staffID = usrFields("Employee_ID")
    End If

    Set rst = CreateObject("ADODB.RecordSet")
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount = 0 Then
        rst.AddNew
    End If
    rst.fields("StaffID") = staffID
    rst.fields("StaffName") = usrFields("Title") & " " & usrFields("Name")
    rst.fields("InitialStaffID") = IIF(Len(usrFields("Employee_ID")) > 0, usrFields("Employee_ID"), "NONE")
    rst.fields("StaffTypeID") = GetStaffTypeID(usrFields("Dept_Name"))
    rst.fields("GenderID") = GetGender(usrFields("Gender"))
    rst.fields("StaffLevelID") = "S001"
    rst.fields("StaffStatusID") = IIF(usrFields("IsActive") = 1, "S001", "S002")
    rst.fields("ContactInfo") = usrFields("Phone") & "||" & usrFields("Mobile") & "||" & usrFields("Email")
    rst.fields("KeyPrefix") = ""
    rst.fields("Address") = usrFields("House_No") & "||" & usrFields("Street_Name") & "||" & usrFields("Locality")
    rst.fields("City") = usrFields("City")
    rst.fields("Location") = ""
    rst.UpdateBatch
    rst.Close
    Set rst = Nothing
End Sub
Sub AddSystemUser(userFields)
    Dim sql, rst, staffID, jsd, rst2

    If Len(userFields("PayrollEmployeeID")) > 0 Then
        staffID = userFields("PayrollEmployeeID")
    Else
        staffID = userFields("Employee_ID")
    End If

    sql = "select * from SystemUser where SystemUserID='" & staffID & "' "
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount = 0 Then
        rst.AddNew
        jsd = GetJobSchedule(userFields("Title"), userFields("Dept_Name"), userFields("Desi_Name"))
   
        rst.fields("SystemUserID") = staffID
        rst.fields("SystemUserName") = staffID
        rst.fields("InitialSystemUserID") = userFields("Employee_ID")
        rst.fields("UserPassword") = ""
        rst.fields("ConfirmPassword") = ""
        rst.fields("StaffID") = staffID
        If jsd = "Test" Or (userFields("IsActive") <> 1) Then
            rst.fields("UserStatusID") = "UST002"
        Else
            rst.fields("UserStatusID") = "UST001"
        End If
        rst.fields("BranchID") = "B001"
        rst.fields("DepartmentID") = GetComboNameFld("Jobschedule", jsd, "DepartmentID")
        rst.fields("UnitID") = GetComboNameFld("JobSchedule", jsd, "UnitID")
        rst.fields("JobScheduleID") = jsd
        rst.fields("AccountStartDate") = Now()
        rst.fields("KeyPrefix") = ""
        'rst.Fields("AccountEndDate") = ""
        'rst.Fields("PasswordChangeDate") = ""
        'rst.Fields("LastLoginDate") = ""
    End If

    If UCase(Left(jsd, 3)) = "M03" Then
        'add specialist
        sql = "select * from Specialist where SpecialistID='" & rst.fields("SystemUserID") & "' "
        Set rst2 = CreateObject("ADODB.Recordset")
        rst2.open qryPro.FltQry(sql), conn, 3, 4
        If rst2.RecordCount = 0 Then
            rst2.AddNew
        End If

        rst2.fields("SpecialistID") = rst.fields("SystemUserID")
        rst2.fields("SpecialistName") = userFields("Name")
        rst2.fields("SpecialistTypeID") = "000"
        rst2.fields("SpecialistStatusID") = GetComboNameFld("Jobschedule", jsd, "DepartmentID")
        rst2.fields("AppointInterval") = 0
        rst2.fields("NumberOfAppoint") = 0
        rst2.fields("ContactInfo") = ""
        'rst2.Fields("SpecialistDate") = ""
        rst2.fields("SpecialistVal") = 0
        rst2.fields("SpecialistInfo") = ""
        rst2.fields("KeyPrefix") = ""

        rst2.AddNew
        rst2.Close
        Set rst2 = Nothing
    End If

    rst.UpdateBatch
    rst.Close
    Set rst = Nothing
End Sub
Function GetJobSchedule(title, dept, desi)
    Dim ot

    ot = "Test"
    Select Case UCase(dept)
        Case "IT"
            ot = "DPT001"
        Case "LABORATORY"
            ot = "S13"
        Case "PHARMACY"
            ot = "S22"
        Case "PHYSIOTHERAPY"
            ot = "M0208"
        Case "RADIOLOGY"
            ot = "S19"
        Case "RESEARCH"
            ot = "Research"
        Case "NURSING"
            ot = "M0207"
        Case "MEDICAL SERVICES"
            ot = "M0307"
        Case "DIET & NUTRITION"
            ot = "M26"
        Case "FINANCE"
            ot = "CreditControl"
        Case "PROCUREMENT & MATERIALS MGT"
            ot = "S20"
    End Select

    Select Case UCase(desi)
        Case "SENIOR PHYSIOTHERAPIST", "SENIOR ORTHOTICS AIDE", UCase("Senior Physio Aide"), UCase("Snr Physiotherapist")
            ot = "M0308"
        Case UCase("Dietician"), "PRINCIPAL DIETICIAN"
            ot = "M0318"
        Case "FRONT DESK AIDE"
            ot = "S01"
        Case "SENIOR MEDICAL RECORDS OFFICER", UCase("Medical Records Officer")
            ot = "MedicalRecords"
        Case UCase("Patient Advocate/Social Work")
            ot = "PatientServices"
    End Select

    If ot = "Test" And title = "Dr." Then
        ot = "M0307"
    End If
    GetJobSchedule = ot
End Function
Function GetStaffTypeID(deptName)
    Dim sql, rst, cnt, tmpRst

    If Len(deptName) < 1 Or IsNull(deptName) Then
        GetStaffTypeID = "000"
    Else
        sql = " select * from StaffType where StaffTypeName='" & deptName & "' "
        Set rst = CreateObject("ADODB.RecordSet")
        rst.open qryPro.FltQry(sql), conn, 3, 4
        If rst.RecordCount = 0 Then
            Set tmpRst = CreateObject("ADODB.RecordSet")
            sql = " select count(*) as [idx] from StaffType "
            tmpRst.open qryPro.FltQry(sql), conn, 3, 4

            cnt = tmpRst.fields("idx")

            rst.AddNew
            rst.fields("StaffTypeName") = deptName
            rst.fields("StaffTypeID") = "S" & Right(CStr(1001 + cnt), 3)
            rst.UpdateBatch
        End If
        GetStaffTypeID = rst.fields("StaffTypeID")
    End If

End Function
Sub MigrateCountry()
    Dim srcRst, srcSql, mySqlConn, targSql, targRst, CountryID

    Set mySqlConn = GetMySQLConnection()
    Set srcRst = CreateObject("ADODB.RecordSet")

    srcSql = " select * from country_master where CountryID<>'C002'"
    srcRst.open srcSql, mySqlConn, 3, 4

    If srcRst.RecordCount > 0 Then
        Set targRst = CreateObject("ADODB.RecordSet")
        Do While Not srcRst.EOF
            CountryID = "C" & Right((1001 + srcRst.fields("CountryID")), 3)
            targSql = " select * from Country where CountryID='" & CountryID & "' "

            targRst.open qryPro.FltQry(targSql), conn, 3, 4
            If targRst.RecordCount = 0 Then
                targRst.AddNew

            End If
            targRst.fields("CountryID") = CountryID
            targRst.fields("CountryName") = srcRst.fields("Name")
            targRst.fields("CountryPhoneCode") = ""
            targRst.fields("CountryAbbreviation") = ""
            targRst.fields("Description") = "Currency=" & srcRst.fields("Currency") & "||Notation=" & srcRst.fields("Notation") '& "||Address=" & srcRst.Fields("Address")
            targRst.fields("KeyPrefix") = ""

            targRst.UpdateBatch

            targRst.Close
            srcRst.MoveNext
        Loop
        srcRst.Close
    End If

    Set srcRst = Nothing
    Set targRst = Nothing

End Sub
Function GetDoctorEmployeeID(docID)
    Dim sql, rst, ot
    ot = "M0307"
    sql = "select Employee_ID from doctor_employee where doctor_id='" & docID & "'"
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, GetMySQLConnection(), 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        ot = rst.fields("Employee_ID")
    End If
    GetDoctorEmployeeID = ot
End Function
Sub MigrateSponsor()
    Dim sql, rst, spn, spnPk, ins, insPk, tmpRst, insTyp

    sql = "select * from f_panel_master where Panel_ID>1"
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, GetMySQLConnection(), 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Set spn = CreateObject("ADODB.RecordSet")
        Set ins = CreateObject("ADODB.RecordSet")
        Set tmpRst = CreateObject("ADODB.RecordSet")

        insTyp = "INS001"
        Do While Not rst.EOF
            spnPk = "SPN-" & Right(CStr(1000 + rst.fields("Panel_ID")), 3)
            sql = "select * from Sponsor where SponsorID='" & spnPk & "' "

            spn.open qryPro.FltQry(sql), conn, 3, 4
            If spn.RecordCount = 0 Then
                spn.AddNew
            End If
            spn.fields("SponsorID") = spnPk
            spn.fields("SponsorName") = rst.fields("Company_Name")
            spn.fields("SponsorTypeID") = "S001"
            spn.fields("SponsorStatusID") = IIF(rst.fields("DateTo") > Now(), "S001", "S002")
            spn.fields("Address") = rst.fields("Add1") & vbCrLf & rst.fields("Add2")
            spn.fields("City") = ""
            spn.fields("Location") = rst.fields("EmailID")
            spn.fields("OfficePhone") = rst.fields("Contact_Person") & "/" & rst.fields("Phone") & "/" & rst.fields("Mobile")
            spn.fields("OfficeFax") = rst.fields("Fax_No")
            spn.fields("Description") = rst.fields("Agreement")
            spn.fields("KeyPrefix") = ""
            spn.UpdateBatch

            'insurance scheme
            insPk = spn.fields("SponsorID") & "-001"
            sql = "select * from InsuranceScheme where InsuranceSchemeID='" & insPk & "' "
            ins.open qryPro.FltQry(sql), conn, 3, 4
            If ins.RecordCount = 0 Then
                ins.AddNew
            End If
            ins.fields("InsuranceSchemeID") = insPk
            ins.fields("InsuranceSchemeName") = spn.fields("SponsorName")
            ins.fields("InsuranceTypeID") = insTyp
            ins.fields("InsuranceZoneID") = insTyp
            ins.fields("InsuranceGroupID") = "INS"
            ins.fields("SponsorID") = spn.fields("SponsorID")
            ins.fields("SponsorTypeID") = spn.fields("SponsorTypeID")
            ins.fields("InitialSchemeID") = "NONE"
            ins.fields("ReceiptTypeID") = IIF(UCase(rst.fields("PaymentMode")) = "CASH", "R001", "R002")
            ins.fields("InsSchemeModeID") = "I001"
            ins.fields("VettingGroupID") = "INS"
            ins.fields("InsSchemeStatusID") = "I001"
            ins.fields("CostPercent") = 100
            ins.fields("SchemeVal1") = 0
            ins.fields("SchemeVal2") = 0
            ins.fields("SchemeInfo1") = ""
            ins.fields("SchemeInfo2") = ""
            'ins.FIelds("SchemeDate1") = ""
            'ins.FIelds("SchemeDate2") = ""
            ins.fields("Address") = spn.fields("Address")
            ins.fields("City") = spn.fields("City")
            ins.fields("Location") = spn.fields("Location")
            ins.fields("OfficePhone") = spn.fields("OfficePhone")
            ins.fields("OfficeFax") = spn.fields("OfficeFax")
            ins.fields("Description") = ""
            ins.fields("KeyPrefix") = ""
            ins.UpdateBatch

            'insurance Type
            sql = "select * from InsuranceType where InsuranceTypeID='" & insTyp & "'"
            tmpRst.open qryPro.FltQry(sql), conn, 3, 4
            If tmpRst.RecordCount = 0 Then
                tmpRst.AddNew
            End If
            tmpRst.fields("InsuranceTypeID") = insTyp
            tmpRst.fields("InsuranceTypeName") = "INSURANCE"
            tmpRst.fields("ServicePercent") = 100
            tmpRst.fields("DrugPricePercent") = 100
            tmpRst.fields("ItemPricePercent") = 100
            tmpRst.fields("ConsultPercent") = 100
            tmpRst.fields("TreatPercent") = 100
            tmpRst.fields("DiagnosePercent") = 0
            tmpRst.fields("LabTestPercent") = 100
            tmpRst.fields("BedPercent") = 100
            tmpRst.fields("Address") = ""
            tmpRst.fields("City") = ""
            tmpRst.fields("Description") = ""
            tmpRst.fields("Location") = ""
            tmpRst.fields("OfficePhone") = ""
            tmpRst.fields("OfficeFax") = ""
            tmpRst.fields("KeyPrefix") = insTyp
            tmpRst.fields("FridgePercent") = 0
            tmpRst.fields("MortChargePercent") = 100
            tmpRst.UpdateBatch
            tmpRst.Close

            'insurance Zone
            sql = "select * from InsuranceZone where InsuranceZoneID='" & insTyp & "'"
            tmpRst.open qryPro.FltQry(sql), conn, 3, 4
            If tmpRst.RecordCount = 0 Then
                tmpRst.AddNew
            End If
            tmpRst.fields("InsuranceZoneID") = insTyp
            tmpRst.fields("InsuranceZoneName") = "INSURANCE"
            tmpRst.fields("InsuranceTypeID") = insTyp
            tmpRst.fields("InsuranceGroupID") = "INS"
            tmpRst.fields("Address") = ""
            tmpRst.fields("Description") = ""
            tmpRst.fields("City") = ""
            tmpRst.fields("KeyPrefix") = insTyp
            tmpRst.fields("Location") = ""
            tmpRst.fields("OfficePhone") = ""
            tmpRst.fields("OfficeFax") = ""
            tmpRst.UpdateBatch
            tmpRst.Close

            spn.Close
            ins.Close
            rst.MoveNext
        Loop
        rst.Close
    End If
    Set rst = Nothing
End Sub
Function GetSpecialistType(subcatName)
    Dim sql, rst, ot
    ot = "LSHHI1"
    sql = "select * from SpecialistType where SpecialistTypeName='" & subcatName & "'"
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount > 0 Then
        ot = rst.fields("SpecialistTypeID")
    End If

    GetSpecialistType = ot
End Function
Function GetSystemVariableValue(VarName)
    Dim ot
    ot = GetComboNameFld("SystemVariables", VarName, "SystemVariableValue")

    GetSystemVariableValue = ot
End Function
Function SetSystemVariableValue(VarName, val)
    Dim sql, dTyp

    If IsNumeric(val) Then
        dTyp = "DTY004"
    Else
        dTyp = "DTY002"
    End If


    sql = " begin"
    sql = sql & " if not exists(select * from SystemVariables where SystemVariableID='" & VarName & "') "
    sql = sql & "   insert into SystemVariables(SystemVariableID, SystemVariableName, DataTypeID, SystemVariableValue)"
    sql = sql & "   values('" & VarName & "', '" & VarName & "', '" & dTyp & "', '" & val & "')"
    sql = sql & " else"
    sql = sql & "   update SystemVariables set SystemVariableValue='" & val & "' where SystemVariableID='" & VarName & "' "
    sql = sql & " "
    sql = sql & " end"
    conn.execute qryPro.FltQry(sql)
End Function
Function GetMaritalStatus(maritalStatus)
    Dim ot

    ot = "M001"
    maritalStatus = Trim(maritalStatus)
    If maritalStatus = "Married" Then
        ot = "M003"
    ElseIf maritalStatus = "Single" Then
        ot = "M002"
    ElseIf maritalStatus = "Widow" Then
        ot = "M004"
    ElseIf maritalStatus = "Divorced" Then
        ot = "M005"
    ElseIf maritalStatus = "Unmarried" Then
        ot = "M006"
    End If

    GetMaritalStatus = ot
End Function
Function GetCountry(countryName)
    Dim rst, sql, ot

    sql = " select * from Country Where CountryName='" & countryName & "'"
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount > 0 Then
        ot = rst.fields("CountryID")
    End If

    GetCountry = ot
End Function
Function GetPhoneNo(rst)
    Dim phone, tmp

    phone = ""
    tmp = Trim(rst.fields("Phone"))
    If Len(tmp) > 0 Then
        phone = tmp
    End If

    tmp = Trim(rst.fields("EmergencyPhone"))
    If Len(tmp) > 0 Then
        If Len(phone) > 0 Then
            phone = phone & ", "
        End If
        phone = phone & tmp
    End If

    tmp = Trim(rst.fields("EmergencyPhoneNo"))
    If Len(tmp) > 0 Then
        If Len(phone) > 0 Then
            phone = phone & ", "
        End If
        phone = phone & tmp
    End If

    GetPhoneNo = phone
End Function
Function GetGender(gender)
    Dim ot

    gender = Trim(gender)
    If gender = "Male" Then
        ot = "GEN01"
    ElseIf gender = "Female" Then
        ot = "GEN02"
    Else
        ot = "G001"
    End If
    GetGender = ot
End Function
Function GetTitleID(ByVal rst)
    Dim ot, titleName

    ot = "0"
    titleName = UCase(rst.fields("Title"))
    If titleName = UCase("Mr.") Then
        ot = "1"
    ElseIf titleName = UCase("Miss.") Then
        ot = "3"
    ElseIf titleName = UCase("Mrs.") Then
        ot = "2"
    ElseIf titleName = UCase("Mst.") Then
        ot = "0"
    ElseIf titleName = UCase("Madam") Then
        ot = "4"
    ElseIf titleName = UCase("Master") Then
        ot = "9"
    ElseIf titleName = UCase("Dr.") Then
        ot = "5"
    ElseIf titleName = UCase("Prof.") Then
        ot = "7"
    ElseIf titleName = UCase("Nana") Then
        ot = "17"
    ElseIf titleName = UCase("Hajia") Then
        ot = "23"
    ElseIf titleName = UCase("Baby") Then
        If rst.fields("Gender") = "Male" Then
            ot = "12"
        Else
            'Female
            ot = "13"
        End If
    ElseIf titleName = UCase("Alhaji") Then
        ot = "19"
    Else

    End If

    GetTitleID = ot
End Function
Function FormatWorkingMonthAdd(dt)
    Dim ot, sql, wkMthName, kp

    ot = ""
    If IsDate(dt) Then
        ot = "MTH" & CStr(Year(CDate(dt))) & Right(CStr(Month(CDate(dt)) + 100), 2)
        wkMthName = MonthName(Month(dt)) & " " & Year(dt)
        kp = Month(dt) & Year(dt)
        kp = Right(CStr(Year(dt)), 2) & Right(CStr(100 + Month(dt)), 2)

        sql = " if not exists( select * from WorkingMonth where WorkingMonthID='" & ot & "' ) "
        sql = sql & " INSERT INTO WorkingMonth (WorkingMonthID, WorkingMonthName, WorkingYearID, WorkingQuarterID, WorkMonth, KeyPrefix, Description)"
        sql = sql & " VALUES ('" & ot & "', '" & wkMthName & "', '" & FormatWorkingYearAdd(dt) & "', '" & FormatWorkingQuarterAdd(dt) & "', '" & wkMthName & "', '" & kp & "', '/');"
        conn.execute qryPro.FltQry(sql)

        sql = " if not exists( select * from FirstMonth where FirstMonthID='" & ot & "' ) "
        sql = sql & " INSERT INTO FirstMonth (FirstMonthID, FirstMonthName, FirstYearID, FirstQuarterID, FirstMonth, KeyPrefix, Description)"
        sql = sql & " VALUES ('" & ot & "', '" & wkMthName & "', '" & FormatWorkingYearAdd(dt) & "', '" & FormatWorkingQuarterAdd(dt) & "', '" & wkMthName & "', '" & kp & "', '/');"
        conn.execute qryPro.FltQry(sql)

        sql = " if not exists( select * from BillMonth where BillMonthID='" & ot & "' ) "
        sql = sql & " INSERT INTO BillMonth (BillMonthID, BillMonthName, BillYearID, BillQuarterID, BillMonth, KeyPrefix, Description)"
        sql = sql & " VALUES ('" & ot & "', '" & wkMthName & "', '" & FormatWorkingYearAdd(dt) & "', '" & FormatWorkingQuarterAdd(dt) & "', '" & wkMthName & "', '" & kp & "', '/');"
        conn.execute qryPro.FltQry(sql)

        sql = " if not exists( select * from AppointMonth where AppointMonthID='" & ot & "' ) "
        sql = sql & " INSERT INTO AppointMonth (AppointMonthID, AppointMonthName, AppointYearID, AppointQuarterID, AppointMonth, KeyPrefix, Description)"
        sql = sql & " VALUES ('" & ot & "', '" & wkMthName & "', '" & FormatWorkingYearAdd(dt) & "', '" & FormatWorkingQuarterAdd(dt) & "', '" & wkMthName & "', '" & kp & "', '/');"
        conn.execute qryPro.FltQry(sql)
    End If
         
    FormatWorkingMonthAdd = ot
End Function
Function FormatWorkingDayAdd(dt)
    Dim ot, sql, wkMth, wkDyName, kp, wkYr

    ot = ""
    If IsDate(dt) Then
        dt = CDate(dt)

        ot = "DAY" & CStr(Year(CDate(dt))) & Right(CStr(Month(CDate(dt)) + 100), 2) & Right(CStr(Day(CDate(dt)) + 100), 2)
        wkMth = FormatWorkingMonthAdd(dt)
        wkYr = FormatWorkingYearAdd(dt)
        wkDyName = Day(dt) & " " & MonthName(Month(dt)) & " " & Year(dt) & " [" & WeekdayName(Weekday(dt), True) & "]"
        kp = Right(CStr(Year(dt)), 2) & Right(100 + Month(dt), 2) & Right(100 + Day(dt), 2)

        sql = " if not exists( select * from WorkingDay where WorkingDayID='" & ot & "' ) "
        sql = sql & "INSERT INTO WorkingDay (WorkingDayID, WorkingDayName, WorkingMonthID, WorkDate, KeyPrefix, Description) "
        sql = sql & " VALUES ('" & ot & "', '" & wkDyName & "', '" & wkMth & "', '" & dt & "', '" & kp & "', '/');"
        conn.execute qryPro.FltQry(sql)

        sql = " if not exists( select * from FirstDay where FirstDayID='" & ot & "' ) "
        sql = sql & "INSERT INTO FirstDay (FirstDayID, FirstDayName, FirstMonthID, FirstDate, KeyPrefix, Description) "
        sql = sql & " VALUES ('" & ot & "', '" & wkDyName & "', '" & wkMth & "', '" & dt & "', '" & kp & "', '/');"
        conn.execute qryPro.FltQry(sql)

        sql = " if not exists( select * from BillDay where BillDayID='" & ot & "' ) "
        sql = sql & "INSERT INTO BillDay (BillDayID, BillDayName, BillMonthID, BillDate, KeyPrefix, Description) "
        sql = sql & " VALUES ('" & ot & "', '" & wkDyName & "', '" & wkMth & "', '" & dt & "', '" & kp & "', '/');"
        conn.execute qryPro.FltQry(sql)

        sql = " if not exists( select * from AppointDay where AppointDayID='" & ot & "' ) "
        sql = sql & "INSERT INTO AppointDay (AppointDayID, AppointDayName, AppointMonthID, AppointDate, KeyPrefix, Description) "
        sql = sql & " VALUES ('" & ot & "', '" & wkDyName & "', '" & wkMth & "', '" & dt & "', '" & kp & "', '/');"
        conn.execute qryPro.FltQry(sql)

        sql = " if not exists(select * from BranchBatch where BranchBatchID='B001-" & ot & "' )"
        sql = sql & "insert into BranchBatch (BranchBatchID, BranchBatchName, BranchBatchTypeID, BranchBatchStatusID, BatchPos, BranchID, WorkingYearID, WorkingMonthID, WorkingDayID, BatchDate, BatchInfo, Description, KeyPrefix) "
        sql = sql & " values ('B001-" & ot & "', 'FOCOS [" & wkDyName & "]', 'B001', 'B003', 1, 'B001', '" & wkYr & "', '" & wkMth & "', '" & ot & "', '" & dt & "', NULL, NULL, NULL)"
        conn.execute qryPro.FltQry(sql)

        sql = "if not exists(select * from BranchSubBatch where BranchSubBatchID='B001-" & ot & "-Receipt') "
        sql = sql & " insert into BranchSubBatch (BranchSubBatchID, BranchSubBatchName, BranchSubBatchTypeID, BranchBatchID, BranchBatchTypeID, BranchSubBatchStatID, SubBatchPos, BranchID, WorkingYearID, WorkingMonthID, WorkingDayID, SubBatchDate, SubBatchInfo, Description, KeyPrefix)"
        sql = sql & " values('B001-" & ot & "-Receipt', 'FOCOS [" & wkDyName & "Receipt]', 'Receipt', 'B001-" & ot & "', 'B001', 'B003', 1, 'B001', '" & wkYr & "', '" & wkMth & "', '" & wkDyName & "', '" & dt & "', NULL, NULL, NULL) "
        conn.execute qryPro.FltQry(sql)

        sql = "if not exists(select * from BranchSubBatch where BranchSubBatchID='B001-" & ot & "-Visitation') "
        sql = sql & " insert into BranchSubBatch (BranchSubBatchID, BranchSubBatchName, BranchSubBatchTypeID, BranchBatchID, BranchBatchTypeID, BranchSubBatchStatID, SubBatchPos, BranchID, WorkingYearID, WorkingMonthID, WorkingDayID, SubBatchDate, SubBatchInfo, Description, KeyPrefix)"
        sql = sql & " values('B001-" & ot & "-Visitation', 'FOCOS [" & wkDyName & "Visitation]', 'Visitation', 'B001-" & ot & "', 'B001', 'B003', 1, 'B001', '" & wkYr & "', '" & wkMth & "', '" & wkDyName & "', '" & dt & "', NULL, NULL, NULL) "
        conn.execute qryPro.FltQry(sql)

    End If

    FormatWorkingDayAdd = ot
End Function
Function FormatWorkingQuarterAdd(dt)
    Dim mth, ot

    ot = ""
    If IsDate(dt) Then
        mth = Month(CDate(dt))
        ot = "QTR" & CStr(Year(CDate(dt))) & Right(CStr((Int((mth - 1) / 3) + 1) + 100), 2)

    End If

    FormatWorkingQuarterAdd = ot
End Function
Function FormatWorkingYearAdd(dt)
    Dim ot, sql

    ot = ""
    If IsDate(dt) Then
        ot = "YRS" & CStr(Year(CDate(dt)))
        
        sql = " if not exists( select * from WorkingYear where WorkingYearID='" & ot & "' ) "
        sql = sql & " INSERT INTO WorkingYear (WorkingYearID, WorkingYearName, WorkYear, KeyPrefix, Description) "
        sql = sql & " VALUES ('" & ot & "', '" & Year(dt) & "', '" & Year(dt) & "', '" & Right(CStr(Year(dt)), 2) & "', '/'); "
        conn.execute qryPro.FltQry(sql)

        sql = " if not exists( select * from FirstYear where FirstYearID='" & ot & "' ) "
        sql = sql & " INSERT INTO FirstYear (FirstYearID, FirstYearName, FirstYear, KeyPrefix, Description) "
        sql = sql & " VALUES ('" & ot & "', '" & Year(dt) & "', '" & Year(dt) & "', '" & Right(CStr(Year(dt)), 2) & "', '/'); "
        conn.execute qryPro.FltQry(sql)
        
        sql = " if not exists( select * from BillYear where BillYearID='" & ot & "' ) "
        sql = sql & " INSERT INTO BillYear (BillYearID, BillYearName, BillYear, KeyPrefix, Description) "
        sql = sql & " VALUES ('" & ot & "', '" & Year(dt) & "', '" & Year(dt) & "', '" & Right(CStr(Year(dt)), 2) & "', '/'); "
        conn.execute qryPro.FltQry(sql)

        sql = " if not exists( select * from AppointYear where AppointYearID='" & ot & "' ) "
        sql = sql & " INSERT INTO AppointYear (AppointYearID, AppointYearName, AppointYear, KeyPrefix, Description) "
        sql = sql & " VALUES ('" & ot & "', '" & Year(dt) & "', '" & Year(dt) & "', '" & Right(CStr(Year(dt)), 2) & "', '/'); "
        conn.execute qryPro.FltQry(sql)
    End If

    FormatWorkingYearAdd = ot
End Function
Function FormatAppointStartTimeAdd(appTime)
    Dim ot, tmp, tmName, sql

    ot = "0600"
    If IsDate(appTime) Then
        ot = Right(CStr(100 + Hour(appTime)), 2) & Right(CStr(100 + Minute(appTime)), 2)
        tmName = CStr(TimeSerial(Hour(appTime), BucketValue(Minute(appTime), 5), 0))
        tmName = Replace(tmName, ":00 ", " ")
        
        If Len(tmName) < 8 Then
            tmName = "0" & tmName
        End If
        
        sql = " if not exists(select * from AppointStartTime where AppointStartTimeID='" & ot & "' )"
        sql = sql & " insert into AppointStartTime(AppointStartTimeID, AppointStartTimeName, AppointStartTime)"
        sql = sql & " values ('" & ot & "', '" & tmName & "', '2022-02-08 " & tmName & "' )"
        conn.execute qryPro.FltQry(sql)

        sql = " if not exists(select * from AppointEndTime where AppointEndTimeID='" & ot & "' )"
        sql = sql & " insert into AppointEndTime(AppointEndTimeID, AppointEndTimeName, AppointEndTime)"
        sql = sql & " values ('" & ot & "', '" & tmName & "', '2022-02-08 " & tmName & "' )"
        conn.execute qryPro.FltQry(sql)
    End If

    FormatAppointStartTimeAdd = ot
End Function
Function BucketValue(number, giveOrTake)
    Dim ot
     
    ot = Round(number / giveOrTake) * giveOrTake
    
    BucketValue = ot
End Function
Function GetSystemUserID(employee_id)
    Dim srcSql, srcRst, ot

    Set srcRst = CreateObject("ADODB.RecordSet")
    srcSql = "select SystemUserID from SystemUser where InitialSystemUserID='" & employee_id & "' "
    srcRst.open qryPro.FltQry(srcSql), conn, 3, 4

    If srcRst.RecordCount > 0 Then
        srcRst.movefirst
        ot = srcRst.fields("SystemUserID")
        srcRst.Close
    Else
        ot = employee_id
    End If
    Set srcRst = Nothing

    GetSystemUserID = ot
End Function
Sub MigrateWard()
    Dim sql, rst, lastWard, ward, bedName, idx

    Set rst = CreateObject("ADODB.RecordSet")
    sql = " select * from room_master order by Name, Room_No, Bed_No asc "
    rst.open sql, GetMySQLConnection(), 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            If lastWard <> rst.fields("Name") Then
                idx = 0
                ward = GetWard(rst.fields("Name"))
            End If

            idx = idx + 1
            If Len(ward) > 0 Then
                bedName = "Room " & rst.fields("Room_No") & " Bed " & rst.fields("Bed_No")
                AddBed bedName, idx, ward
            End If
            lastWard = rst.fields("Name")
            rst.MoveNext
        Loop
        rst.Close
    End If
End Sub
Function GetWard(roomName)
    Dim wardID, wRst, tmpRst, sql

    Set wRst = CreateObject("ADODB.RecordSet")
    Set tmpRst = CreateObject("ADODB.RecordSet")

    wardID = ""
    Select Case UCase(roomName)
        Case UCase("2 BED")
            wardID = "W001"
        Case UCase("4 BED")
            wardID = "W002"
        Case UCase("CHILDREN WARD")
            wardID = "W003"
        Case UCase("RECOVERY ROOM")
            wardID = "W004"
        Case UCase("OR")
            wardID = "W005"
        
        'Case UCase("3 BED")
        'Case UCase("A001 (Boys)")
        'Case UCase("A002 (Girls)")
        'Case UCase("A003 (Boys)")
        'Case UCase("A004 (Boys)")
        'Case UCase("A005 (Girls)")
        'Case UCase("B001 (Boys)")
        'Case UCase("B002 (Girls)")
        'Case UCase("B003 (Boys)")
        'Case UCase("B004 (Boys)")
        'Case UCase("MALE ADM.WARD")
        'Case UCase("FEMALE ADM.")
        'Case UCase("PATIENT ROOM")
        'Case UCase("REHAB")
    End Select

    If Len(wardID) > 0 Then
        sql = "select * from Ward where WardID='" & wardID & "' "
        wRst.open qryPro.FltQry(sql), conn, 3, 4
        If wRst.RecordCount = 0 Then
            wRst.AddNew
        End If
        wRst.fields("WardID") = wardID
        wRst.fields("WardName") = roomName
        wRst.fields("BlockID") = "B001"
        wRst.fields("Description") = ""
        wRst.fields("KeyPrefix") = ""
        wRst.UpdateBatch

        conn.execute "update JobSchedule set JobScheduleName='" & roomName & "' where JobSCheduleID='" & wardID & "'"
        'waiting list
        AddBed "Waiting List", 0, wardID
    End If

    GetWard = wardID
End Function
Sub AddBed(bedName, bedPos, wardID)
    Dim sql, tmpRst, bedID, bedNo

    bedNo = Right(CStr(1000 + bedPos), 3)
    bedID = wardID & "-" & bedNo
    Set tmpRst = CreateObject("ADODB.RecordSet")
    sql = "select * from Bed where BedID='" & bedID & "' "
    tmpRst.open qryPro.FltQry(sql), conn, 3, 4
    If tmpRst.RecordCount = 0 Then
        tmpRst.AddNew
    End If
    If True Then
        tmpRst.fields("BedID") = bedID
        tmpRst.fields("BedName") = bedName
        tmpRst.fields("BedNoID") = bedNo
        tmpRst.fields("BedPos") = bedPos
        tmpRst.fields("BedTypeID") = "ECO"
        tmpRst.fields("BedStatusID") = "B001"
        tmpRst.fields("BedModeID") = "B001"
        tmpRst.fields("BlockID") = "B001"
        tmpRst.fields("WardID") = wardID
        tmpRst.fields("WardSectionID") = wardID
        tmpRst.fields("BedGroupID") = "B17"
        tmpRst.fields("BedClassID") = "B001"
        tmpRst.fields("BillGroupCatID") = "NONE"
        tmpRst.fields("BillGroupID") = "001"
        tmpRst.fields("BedCharge") = 0
        tmpRst.fields("BedVal1") = 0
        tmpRst.fields("BedVal2") = 0
        tmpRst.fields("BedVal3") = 0
        tmpRst.fields("BedVal4") = 0
        tmpRst.fields("BedInfo1") = ""
        tmpRst.fields("BedInfo2") = ""
        'tmpRst.Fields("BedDate1") = ""
        'tmpRst.Fields("BedDate2") = ""
        tmpRst.fields("Description") = ""
        tmpRst.fields("KeyPrefix") = ""
        tmpRst.fields("BillInput1") = 0
        tmpRst.fields("BillInput2") = 0
        tmpRst.fields("BillInput3") = 0
        tmpRst.fields("BillInout4") = 0
        tmpRst.UpdateBatch
    End If
    UpdateBedMatrixFromTemplate tmpRst.fields("BedID"), tmpRst.fields("BedCharge")
    tmpRst.Close
End Sub
Sub UpdateBedMatrixFromTemplate(pk, cst)
    Dim html, sql, rstAge, rstAdmTyp, insType, admTyp, agegrp, rstIns, target, isValid
    Dim errorFields, matSQL, tmpSQL, insTypeVal, ageGrpVal, admTypVal
    Dim treatVal
    isValid = True
    
    Set rstAdmTyp = CreateObject("ADODB.RecordSet")
    Set rstAge = CreateObject("ADODB.RecordSet")
    Set rstIns = CreateObject("ADODB.RecordSet")
    
    sql = "select AgeGroupID, AgeGroupName, BedPercent from AgeGroup"
    rstAge.open qryPro.FltQry(sql), conn, 3, 4
    
    sql = "select AdmissionTypeName, AdmissionTypeID, BedPercent from AdmissionType"
    rstAdmTyp.open qryPro.FltQry(sql), conn, 3, 4
    
    sql = " select InsuranceTypeID, InsuranceTypeName, BedPercent from InsuranceType where InsuranceTypeName<>'-' "
    rstIns.open qryPro.FltQry(sql), conn, 3, 4
    
    If rstIns.RecordCount > 0 Then
        rstIns.movefirst
        Do While Not rstIns.EOF
            If rstAge.RecordCount > 0 Then
                rstAge.movefirst
                Do While Not rstAge.EOF
                    If rstAdmTyp.RecordCount > 0 Then
                        rstAdmTyp.movefirst
                        Do While Not rstAdmTyp.EOF
                            insType = rstIns.fields("InsuranceTypeID")
                            agegrp = rstAge.fields("AgeGroupID")
                            admTyp = rstAdmTyp.fields("AdmissionTypeID")
                            
'                            If request(target).Count <> 0 Then 'field exists
'                                cst = Trim(request(target))
                                
                                If Not IsValidNumber(cst) Then
                                    isValid = isValid And False
                                End If
                                If isValid Then
                                    agegrp = rstAge.fields("AgeGroupID")
                                    
                                    treatVal = GetComboNameFld("Bed", pk, "BedCharge")
                                    ageGrpVal = rstAge.fields("BedPercent")
                                    admTypVal = rstAdmTyp.fields("BedPercent")
                                    insTypeVal = rstIns.fields("BedPercent")
                                    
                                    tmpSQL = vbCrLf & " if exists(select * from BedCostMatrix where BedID='" & pk & "' "
                                    tmpSQL = tmpSQL & "   and InsuranceTypeID='" & insType & "' and AgeGroupID='" & agegrp & "' and AdmissionTypeID='" & admTyp & "' "
                                    tmpSQL = tmpSQL & " )"
                                    tmpSQL = tmpSQL & "   update BedCostMatrix set BedCharge=" & cst & " where BedID='" & pk & "' "
                                    tmpSQL = tmpSQL & "   and InsuranceTypeID='" & insType & "' and AgeGroupID='" & agegrp & "' and AdmissionTypeID='" & admTyp & "' "
                                    tmpSQL = tmpSQL & " "
                                    tmpSQL = tmpSQL & vbCrLf & " else "
                                    tmpSQL = tmpSQL & "   insert into BedCostMatrix(BedID, InsuranceTypeID, AgeGroupID, AdmissionTypeID, PermuteUpdateID, PermuteStatusID, BedCharge, BedVal, AgeGroupVal, InsuranceTypeVal, AdmissionTypeVal)"
                                    tmpSQL = tmpSQL & "   values('" & pk & "', '" & insType & "',  '" & agegrp & "', '" & admTyp & "', 'P002', 'P001', " & cst & ", " & treatVal & ", " & ageGrpVal & ", " & insTypeVal & ", " & admTypVal & ") "
                                    
                                    matSQL = matSQL & tmpSQL
                                End If
'                            End If
                            rstAdmTyp.MoveNext
                        Loop
                    End If
                    rstAge.MoveNext
                Loop
            End If
            rstIns.MoveNext
        Loop
    End If
    If isValid Then
        conn.execute qryPro.FltQry(matSQL)
    End If
End Sub
Sub MigrateSurgery()
    Dim sql, rst, treat, upDt, field, sgID, VarName

    VarName = "focos_surgery_master"
    sgID = GetSystemVariableValue(VarName)

    sql = "select * from(select f_surgery_master.SurgeryCode, f_surgery_master.Department, f_surgery_master.SubDepartment, f_surgery_master.Name "
    sql = sql & " , f_surgery_rate_list.Rate "
    sql = sql & " , row_number() over(partition by f_surgery_master.Surgery_ID order by f_surgery_master.Surgery_ID, f_surgery_rate_list.DateFrom desc) as idx "
    sql = sql & " from f_surgery_master "
    sql = sql & " left join f_surgery_rate_list on f_surgery_rate_list.Surgery_ID=f_surgery_master.Surgery_ID "
    sql = sql & " where f_surgery_master.IsActive=1 and f_surgery_master.SurgeryCode>='" & sgID & "' "
    sql = sql & " ) AS surg WHERE surg.idx=1 "
    
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, GetMySQLConnection, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            Set treat = CreateObject("ADODB.RecordSet")
            sql = "select * from Treatment where TreatmentID='" & rst.fields("SurgeryCode") & "' "
            treat.open qryPro.FltQry(sql), conn, 3, 4
            If treat.RecordCount = 0 Then
                treat.AddNew
            End If
            If True Then
                treat.fields("TreatmentID") = rst.fields("SurgeryCode")
                treat.fields("TreatmentName") = rst.fields("Name")
                treat.fields("TreatTypeID") = GetTreatType(rst.fields("SubDepartment"), rst.fields("Department"))
                treat.fields("TreatCategoryID") = GetComboNameFld("TreatType", treat.fields("TreatTypeID"), "TreatCategoryID")
                treat.fields("TreatGroupID") = treat.fields("TreatCategoryID")
                treat.fields("TreatClassID") = "T001"
                treat.fields("TreatModeID") = treat.fields("TreatCategoryID")
                treat.fields("BillGroupCatID") = "NONE"
                treat.fields("BillGroupID") = "001"
                treat.fields("UnitCost") = IIF(IsNull(rst.fields("Rate")), 0, rst.fields("Rate"))
                treat.fields("TreatVal1") = 0
                treat.fields("TreatVal2") = 0
                treat.fields("TreatVal3") = 0
                treat.fields("TreatVal4") = 0
                treat.fields("TreatInfo1") = IIF(IsNull(rst.fields("Rate")), "YES", "NO")
                treat.fields("TreatInfo2") = ""
                'treat.Fields("TreatDate1") = ""
                'treat.Fields("TreatDate2") = ""
                treat.fields("Description") = ""
                treat.fields("KeyPrefix") = ""
                treat.fields("BillInput1") = 0
                treat.fields("BillInput2") = 0
                treat.fields("BillInput3") = 0
                treat.fields("BillInout4") = 0
            
                sql = "if Not exists(select * from TreatMode where TreatModeID='" & treat.fields("TreatModeID") & "')"
                sql = sql & " insert into TreatMode(TreatModeID, TreatModeName)"
                sql = sql & " values ('" & treat.fields("TreatModeID") & "', '" & GetComboName("TreatCategory", treat.fields("TreatCategoryID")) & "') "
                conn.execute qryPro.FltQry(sql)

                sql = "if not exists(select * from TreatGroup where TreatGroupID='" & treat.fields("TreatGroupID") & "')"
                sql = sql & " insert into TreatGroup(TreatGroupID, TreatGroupName)"
                sql = sql & " values ('" & treat.fields("TreatGroupID") & "', '" & GetComboName("TreatCategory", treat.fields("TreatCategoryID")) & "') "
                conn.execute qryPro.FltQry(sql)

                treat.UpdateBatch
                UpdateTreatMatrixFromTemplate treat.fields("TreatmentID"), treat.fields("UnitCost")
            End If

            SetSystemVariableValue VarName, treat.fields("TreatmentID")
            treat.Close
            rst.MoveNext
        Loop
        rst.Close
    End If
End Sub
Sub UpdateTreatMatrixFromTemplate(pk, cst)
    Dim html, sql, rstAge, rstMedSer, insType, medSer, agegrp, rstIns, target, isValid
    Dim errorFields, matSQL, tmpSQL, insTypeVal, ageGrpVal, medSerVal, treatVal
    
    isValid = True
    
    Set rstMedSer = CreateObject("ADODB.RecordSet")
    Set rstAge = CreateObject("ADODB.RecordSet")
    Set rstIns = CreateObject("ADODB.RecordSet")
    
    sql = "select AgeGroupID, AgeGroupName, TreatPercent from AgeGroup"
    rstAge.open qryPro.FltQry(sql), conn, 3, 4
    
    sql = "select MedicalServiceName, MedicalServiceID, TreatPercent from MedicalService"
    rstMedSer.open qryPro.FltQry(sql), conn, 3, 4
    
    sql = " select InsuranceTypeID, InsuranceTypeName, TreatPercent from InsuranceType where InsuranceTypeName<>'-' "
    rstIns.open qryPro.FltQry(sql), conn, 3, 4
    
    If rstIns.RecordCount > 0 Then
        rstIns.movefirst
        Do While Not rstIns.EOF
            If rstAge.RecordCount > 0 Then
                rstAge.movefirst
                Do While Not rstAge.EOF
                    If rstMedSer.RecordCount > 0 Then
                        rstMedSer.movefirst
                        Do While Not rstMedSer.EOF
                            insType = rstIns.fields("InsuranceTypeID")
                            agegrp = rstAge.fields("AgeGroupID")
                            medSer = rstMedSer.fields("MedicalServiceID")
                            
                            medSer = "-"
                            agegrp = "-"
                            If (Not IsValidNumber(cst)) Then
                                isValid = isValid And False
                            End If
                            If isValid Then
                                cst = CDbl(cst)
                                agegrp = rstAge.fields("AgeGroupID")
                                medSer = rstMedSer.fields("MedicalServiceID")
                                    
                                treatVal = GetComboNameFld("Treatment", pk, "UnitCost")
                                ageGrpVal = rstAge.fields("TreatPercent")
                                medSerVal = rstMedSer.fields("TreatPercent")
                                insTypeVal = rstIns.fields("TreatPercent")
                                    
                                tmpSQL = vbCrLf & " if exists(select * from TreatCostMatrix where TreatmentID='" & pk & "' "
                                tmpSQL = tmpSQL & "   and InsuranceTypeID='" & insType & "' and AgeGroupID='" & agegrp & "' and MedicalServiceID='" & medSer & "' "
                                tmpSQL = tmpSQL & " )"
                                tmpSQL = tmpSQL & "   update TreatCostMatrix set UnitCost=" & cst & " where TreatmentID='" & pk & "' "
                                tmpSQL = tmpSQL & "   and InsuranceTypeID='" & insType & "' and AgeGroupID='" & agegrp & "' and MedicalServiceID='" & medSer & "' "
                                tmpSQL = tmpSQL & " "
                                tmpSQL = tmpSQL & vbCrLf & " else "
                                tmpSQL = tmpSQL & "   insert into TreatCostMatrix(TreatmentID, InsuranceTypeID, AgeGroupID, MedicalServiceID, PermuteUpdateID, PermuteStatusID, UnitCost, TreatmentVal, AgeGroupVal, InsuranceTypeVal, MedicalServiceVal)"
                                tmpSQL = tmpSQL & "   values('" & pk & "', '" & insType & "',  '" & agegrp & "', '" & medSer & "', 'P002', 'P001', " & cst & ", " & treatVal & ", " & ageGrpVal & ", " & insTypeVal & ", " & medSerVal & ") "
                                    
                                matSQL = matSQL & tmpSQL
                            End If
                            rstMedSer.MoveNext
                        Loop
                    End If
                    rstAge.MoveNext
                Loop
            End If
            rstIns.MoveNext
        Loop
    End If
    
    If isValid Then
        conn.execute qryPro.FltQry(matSQL)
    End If
End Sub
Sub MigrateProcedures()
    Dim sql, rst, treat, upDt, field, lastID, nPfx, pk

   lastID = GetSystemVariableValue("focosProcedureLastID")
    If lastID = "" Then
        lastID = 0
    End If
    sql = "select * from (select f_itemmaster.*, f_subcategorymaster.Name as subcategory_name "
    sql = sql & " , f_categorymaster.Name as CategoryName, f_categorymaster.CategoryID "
    sql = sql & " , f_ratelist.Rate, row_number() over(partition by f_itemmaster.itemID order by f_ratelist.updatedate desc, f_ratelist.EntryDate desc) as idx "
    sql = sql & " from f_itemmaster "
    sql = sql & " left join f_subcategorymaster on f_subcategorymaster.SubCategoryID=f_itemmaster.SubCategoryID "
    sql = sql & " left join f_categorymaster on f_categorymaster.CategoryID=f_subcategorymaster.CategoryID"
    sql = sql & " left join f_ratelist on f_ratelist.ItemID=f_itemmaster.ItemID"
    sql = sql & " where f_subcategorymaster.CategoryID in ('LSHHI10','LSHHI11', 'LSHHI13', 'LSHHI14', 'LSHHI15', 'LSHHI16', 'LSHHI17', 'LSHHI169', 'LSHHI19') and f_itemmaster.IsActive=1"
    sql = sql & "   and f_itemmaster.ID>= " & lastID
    sql = sql & " ) as itm where idx=1"
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, GetMySQLConnection, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            Set treat = CreateObject("ADODB.RecordSet")
            pk = rst.fields("ItemID")
            If Len(rst.fields("ItemCode")) > 0 Then
                If UCase(rst.fields("ItemCode")) <> UCase("&nbsp;") Then
                    pk = rst.fields("ItemCode")
                End If
            End If
            sql = "select * from Treatment where TreatmentID='" & pk & "' "
            treat.open qryPro.FltQry(sql), conn, 3, 4
            If treat.RecordCount = 0 Then
                treat.AddNew
            End If
            If True Then
                treat.fields("TreatmentID") = pk
                If UCase(rst.fields("CategoryID")) = "LSHHI10" Then
                    nPfx = "(" & rst.fields("CategoryName") & "/" & rst.fields("subCategory_Name") & ") "
                Else
                    nPfx = ""
                End If
                treat.fields("TreatmentName") = nPfx & rst.fields("TypeName")
                treat.fields("TreatTypeID") = GetTreatTypeAdd(rst.fields("SubCategoryID"), rst.fields("subcategory_name"), rst.fields("CategoryID"), rst.fields("CategoryName"))
                treat.fields("TreatCategoryID") = GetComboNameFld("TreatType", treat.fields("TreatTypeID"), "TreatCategoryID")
                treat.fields("TreatGroupID") = treat.fields("TreatCategoryID")
                treat.fields("TreatClassID") = "T001"
                treat.fields("TreatModeID") = treat.fields("TreatCategoryID")
                treat.fields("BillGroupCatID") = "NONE"
                treat.fields("BillGroupID") = "001"
                treat.fields("UnitCost") = IIF(IsNull(rst.fields("Rate")), 0, rst.fields("Rate"))
                treat.fields("TreatVal1") = 0
                treat.fields("TreatVal2") = 0
                treat.fields("TreatVal3") = 0
                treat.fields("TreatVal4") = 0
                treat.fields("TreatInfo1") = IIF(IsNull(rst.fields("Rate")), "YES", "NO")
                treat.fields("TreatInfo2") = ""
                'treat.Fields("TreatDate1") = ""
                'treat.Fields("TreatDate2") = ""
                treat.fields("Description") = ""
                treat.fields("KeyPrefix") = ""
                treat.fields("BillInput1") = 0
                treat.fields("BillInput2") = 0
                treat.fields("BillInput3") = 0
                treat.fields("BillInout4") = 0
            
                sql = "if Not exists(select * from TreatMode where TreatModeID='" & treat.fields("TreatModeID") & "')"
                sql = sql & " insert into TreatMode(TreatModeID, TreatModeName)"
                sql = sql & " values ('" & treat.fields("TreatModeID") & "', '" & GetComboName("TreatCategory", treat.fields("TreatCategoryID")) & "') "
                conn.execute qryPro.FltQry(sql)

                sql = "if not exists(select * from TreatGroup where TreatGroupID='" & treat.fields("TreatGroupID") & "')"
                sql = sql & " insert into TreatGroup(TreatGroupID, TreatGroupName)"
                sql = sql & " values ('" & treat.fields("TreatGroupID") & "', '" & GetComboName("TreatCategory", treat.fields("TreatCategoryID")) & "') "
                conn.execute qryPro.FltQry(sql)

                treat.UpdateBatch
                UpdateTreatMatrixFromTemplate treat.fields("TreatmentID"), treat.fields("UnitCost")
                    
            End If

            Call SetSystemVariableValue("focosProcedureLastID", rst.fields("ID"))
            treat.Close
            rst.MoveNext
        Loop
        rst.Close
    End If
End Sub
Function IsValidNumber(num)
    Dim tmp
    tmp = False
    If IsNumeric(num) And num <> "" Then
        tmp = True
    End If
    IsValidNumber = tmp
End Function
Function GetTreatType(typename, department)
    Dim ot, sql, rst

    If Not (Len(typename) > 0) Then
       typename = "NONE"
    End If

    sql = " select TreatTypeID "
    sql = sql & " from TreatType "
    sql = sql & " left join TreatCategory on TreatType.TreatCategoryID=TreatCategory.TreatCategoryID"
    sql = sql & " where 1=1 "
    sql = sql & " and TreatTypeName = '" & typename & "' "
    sql = sql & "  and TreatCategoryName = '" & department & "' "

    Set rst = CreateObject("ADODB.RecordSet")
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount > 0 Then
        ot = rst.fields("TreatTypeID")
    End If

    GetTreatType = ot
End Function
Function GetTreatTypeAdd(TypeId, typename, deptID, department)
    Dim ot, sql, rst

    If Not (Len(typename) > 0) Then
       typename = "NONE"
    End If

    sql = " select * "
    sql = sql & " from TreatType "
    sql = sql & " where 1=1 "
    sql = sql & " and TreatTypeID = '" & TypeId & "' "
    sql = sql & "  and TreatCategoryID = '" & deptID & "' "

    Set rst = CreateObject("ADODB.RecordSet")
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount > 0 Then
        ot = rst.fields("TreatTypeID")
    Else
        ot = TypeId

        rst.AddNew
        rst.fields("TreatTypeID") = TypeId
        rst.fields("TreatTypeName") = typename
        rst.fields("TreatCategoryID") = deptID
        rst.UpdateBatch
        rst.Close
        rst.open qryPro.FltQry("select * from TreatCategory where TreatCategoryID='" & deptID & "'"), conn, 3, 4
        If rst.RecordCount = 0 Then
            rst.AddNew
            rst.fields("TreatCategoryID") = deptID
            rst.fields("TreatCategoryName") = department
            rst.UpdateBatch
        End If
        rst.Close
        
        rst.open qryPro.FltQry("select * from TreatGroup where TreatGroupID='" & deptID & "'"), conn, 3, 4
        If rst.RecordCount = 0 Then
            rst.AddNew
            rst.fields("TreatGroupID") = deptID
            rst.fields("TreatGroupName") = department
            rst.UpdateBatch
        End If
        rst.Close

    End If

    GetTreatTypeAdd = ot
End Function
Function MigrateMedicalConsumables()
    Dim sql, rst, lastID, dRst, stkRst, sale, pk

    lastID = GetSystemVariableValue("focosMedicalConsumablesLastID")
    If lastID = "" Then
        lastID = 0
    End If
    sql = "select f_itemmaster.*, f_subcategorymaster.Name as subcategory_name, f_subcategorymaster.SubCategoryID "
    sql = sql & " from f_itemmaster "
    sql = sql & " left join f_subcategorymaster on f_subcategorymaster.SubCategoryID=f_itemmaster.SubCategoryID "
    sql = sql & " where f_subcategorymaster.CategoryID='LSHHI5' and length(f_itemmaster.ItemCode) > 0"
    sql = sql & "   and f_itemmaster.SubCategoryID not in ('LSHHI100', 'LSHHI101', 'LSHHI174', 'LSHHI176', 'LSHHI34' "
    sql = sql & "           , 'LSHHI38', 'LSHHI40', 'LSHHI42', 'LSHHI51', 'LSHHI52', 'LSHHI55', 'LSHHI59', 'LSHHI61' "
    sql = sql & "           , 'LSHHI64', 'LSHHI65', 'LSHHI85', 'LSHHI84', 'LSHHI86', 'LSHHI87', 'LSHHI88', 'LSHHI89' "
    sql = sql & "           , 'LSHHI92', 'LSHHI93', 'LSHHI95', 'LSHHI96', 'LSHHI97', 'LSHHI98' ) "
    sql = sql & "   And f_itemmaster.ID>= " & lastID
    sql = sql & " order by f_itemmaster.ID asc"

    Set rst = CreateObject("ADODB.RecordSet")
    Set stkRst = CreateObject("ADODB.RecordSet")
    Set sale = CreateObject("ADODB.RecordSet")
    rst.open sql, GetMySQLConnection, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Set dRst = CreateObject("ADODB.RecordSet")
        Do While Not rst.EOF
            pk = Trim(rst.fields("ItemCode"))
            If Not (Len(pk) > 0) Or IsNull(pk) Then
                pk = rst.fields("ItemID")
            End If
            sql = "select * from Drug where DrugID='" & pk & "' "
            dRst.open qryPro.FltQry(sql), conn, 3, 4
            If dRst.RecordCount = 0 Then
                dRst.AddNew
            End If
            If True Then
                sql = "select * from f_stock where ItemID='" & rst.fields("ItemID") & "' order by PostDate desc limit 1"
                stkRst.open sql, GetMySQLConnection, 3, 4
                If stkRst.RecordCount > 0 Then
                    stkRst.movefirst
                End If
                sql = "select * from f_salesdetails where ItemID='" & rst.fields("ItemID") & "' order by f_salesdetails.Date desc, f_salesdetails.Time desc limit 1"
                sale.open sql, GetMySQLConnection, 3, 4
                If sale.RecordCount > 0 Then
                    sale.movefirst
                End If

                dRst.fields("DrugID") = pk
                dRst.fields("DrugName") = rst.fields("TypeName")
                dRst.fields("DrugCategoryID") = "DRC001"
                dRst.fields("DrugTypeID") = GetDrugType(rst.fields("subcategory_name"), rst.fields("SubCategoryID"))
                dRst.fields("UnitOfMeasureID") = rst.fields("UnitType")
                dRst.fields("DrugStatusID") = IIF(rst.fields("IsActive") = 1, "IST001", "IST002")
                dRst.fields("DrugLocationID") = "ILC001"
                dRst.fields("BillGroupCatID") = "B000007"
                dRst.fields("BillGroupID") = "B15"
                dRst.fields("BulkUnitCost") = IIF(sale.RecordCount > 0, sale.fields("PerUnitBuyPrice"), 0)
                dRst.fields("RetailUnitCost") = IIF(sale.RecordCount > 0, sale.fields("PerUnitSellingPrice"), 0)
                dRst.fields("MaxStockQty") = IIF(IsNull(rst.fields("MaxLevel")), 0, rst.fields("MaxLevel"))
                dRst.fields("ReOrderModeID") = "MOD01"
                dRst.fields("ReOrderLevel") = IIF(IsNull(rst.fields("ReOrderLevel")), 0, rst.fields("ReOrderLevel"))
                dRst.fields("ReOrderLevelQty") = IIF(IsNull(rst.fields("ReOrderQty")), 0, rst.fields("ReOrderQty"))
                dRst.fields("QtyBeforeReorder") = IIF(IsNull(rst.fields("ReOrderQty")), 0, rst.fields("ReOrderQty"))
                dRst.fields("ReOrderQty") = IIF(IsNull(rst.fields("ReOrderQty")), 0, rst.fields("ReOrderQty"))
                dRst.fields("AvailableQty") = IIF(IsNull(rst.fields("QtyInHand")), 0, rst.fields("QtyInHand"))
                dRst.fields("Remark") = "0"
                dRst.fields("KeyPrefix") = ""
                dRst.fields("DrugGroupID") = "B15"
                dRst.fields("DrugBrandID") = "B001"
                dRst.fields("DrugModelID") = "D001"
                dRst.fields("DrugModeID") = "B15"
                dRst.fields("Rate1") = 0
                dRst.fields("Rate2") = IIF(IsNull(rst.fields("MinLevel")), 0, rst.fields("MinLevel"))
                dRst.fields("Rate3") = 0
                dRst.fields("Rate4") = 0
                'dRst.Fields("ProdDate1") = ""
                'dRst.Fields("ProdDate2") = ""
                dRst.fields("ProdInfo1") = ""
                dRst.fields("ProdInfo2") = ""

                dRst.UpdateBatch
                    
                sql = "if not exists(select * from UnitOfMeasure where UnitOfMeasureID='" & rst.fields("UnitType") & "')"
                sql = sql & " insert into UnitOfMeasure(UnitOfMeasureID, UnitOfMeasureName)"
                sql = sql & " values('" & rst.fields("UnitType") & "', '" & rst.fields("UnitType") & "') "
                conn.execute qryPro.FltQry(sql)

                
                stkRst.Close
                sale.Close
            End If

            UpdateDrugMatrixFromTemplate dRst.fields("DrugID"), dRst.fields("RetailUnitCost")
            'AddDrugInventory dRst
            Call SetSystemVariableValue("focosMedicalConsumablesLastID", rst.fields("ID"))

            dRst.Close
            rst.MoveNext
        Loop
        rst.Close
    End If
    Set rst = Nothing
End Function
Sub AddDrugInventory(ByVal drug)
    Dim sql, dt, rst, sl

    dt = Now
    sql = "select * from DrugStore"
    Set rst = CreateObject("ADODB.RecordSet")
    Set sl = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            sql = "select * from DrugStockLevel where DrugID='" & drug.fields("DrugID") & "' and DrugStoreID='" & rst.fields("DrugStoreID") & "' "
            sl.open sql, conn, 3, 4
            If sl.RecordCount = 0 Then
                sl.AddNew

                sl.fields("DrugStoreTypeID") = rst.fields("DrugStoreTypeID")
                sl.fields("DrugStoreID") = rst.fields("DrugStoreID")
                sl.fields("DrugID") = drug.fields("DrugID")
                sl.fields("DrugStoreLocationID") = rst.fields("DrugStoreLocationID")
                sl.fields("DrugCategoryID") = drug.fields("DrugCategoryID")
                sl.fields("DrugTypeID") = drug.fields("DrugTypeID")
                sl.fields("UnitOfMeasureID") = drug.fields("UnitOfMeasureID")
                sl.fields("DrugStatusID") = drug.fields("DrugStatusID")
                sl.fields("DrugLocationID") = drug.fields("DrugLocationID")
                sl.fields("BulkUnitCost") = drug.fields("BulkUnitCost")
                sl.fields("RetailUnitCost") = drug.fields("RetailUnitCost")
                sl.fields("TotalCost") = 0
                sl.fields("MaxStockQty") = drug.fields("MaxStockQty")
                sl.fields("ReOrderModeID") = drug.fields("ReOrderModeID")
                sl.fields("ReOrderLevel") = drug.fields("ReOrderLevel")
                sl.fields("ReOrderLevelQty") = drug.fields("ReOrderLevelQty")
                sl.fields("ReOrderQty") = drug.fields("ReOrderQty")
                sl.fields("AvailableQty") = 0 'drug.Fields("AvailableQty")
                sl.fields("QtyBeforeReorder") = drug.fields("QtyBeforeReorder")
                sl.fields("Remark") = drug.fields("Remark")
                sl.fields("PendingAcceptQty") = 0
                sl.fields("AfterAcceptQty") = 0
                sl.fields("StockValue1") = 0
                sl.fields("StockValue2") = 0
                sl.fields("ExpiryDate") = dt
                sl.fields("StockDate1") = dt
                sl.fields("StockDate2") = dt
                sl.fields("StockInfo1") = "-"
                sl.fields("DrugStockStatusID") = "D001"
                sl.fields("StockInfo2") = "-"
                sl.fields("DrugAdjustStatusID") = "D001"
                sl.fields("AdjustValue1") = 0
                sl.fields("AdjustValue2") = 0
                sl.fields("AdjustValue3") = 0
                sl.fields("AdjustDate1") = dt
                sl.fields("AdjustDate2") = dt

                sl.UpdateBatch
            End If
            sl.Close

            rst.MoveNext
        Loop
        rst.Close
    End If

    AddInventoryAccount "Drug", drug.fields("DrugID"), drug.fields("DrugName")
    AddInventoryLookup "Drug", drug.fields("DrugID"), drug.fields("DrugName")

    Set rst = Nothing
End Sub
Sub AddItemInventory(ByVal item)
    Dim sql, dt, rst, sl

    dt = Now
    sql = "select * from ItemStore"
    Set rst = CreateObject("ADODB.RecordSet")
    Set sl = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            sql = "select * from ItemStockLevel where ItemID='" & item.fields("ItemID") & "' and ItemStoreID='" & rst.fields("ItemStoreID") & "' "
            sl.open sql, conn, 3, 4
            If sl.RecordCount = 0 Then
                sl.AddNew

                sl.fields("ItemStoreTypeID") = rst.fields("ItemStoreTypeID")
                sl.fields("ItemStoreID") = rst.fields("ItemStoreID")
                sl.fields("ItemID") = item.fields("ItemID")
                sl.fields("ItemStoreLocationID") = rst.fields("ItemStoreLocationID")
                sl.fields("ItemCategoryID") = item.fields("ItemCategoryID")
                sl.fields("ItemTypeID") = item.fields("ItemTypeID")
                sl.fields("UnitOfMeasureID") = item.fields("UnitOfMeasureID")
                sl.fields("ItemStatusID") = item.fields("ItemStatusID")
                sl.fields("ItemLocationID") = item.fields("ItemLocationID")
                sl.fields("BulkUnitCost") = item.fields("BulkUnitCost")
                sl.fields("RetailUnitCost") = item.fields("RetailUnitCost")
                sl.fields("TotalCost") = 0
                sl.fields("MaxStockQty") = item.fields("MaxStockQty")
                sl.fields("ReOrderModeID") = item.fields("ReOrderModeID")
                sl.fields("ReOrderLevel") = item.fields("ReOrderLevel")
                sl.fields("ReOrderLevelQty") = item.fields("ReOrderLevelQty")
                sl.fields("ReOrderQty") = item.fields("ReOrderQty")
                sl.fields("AvailableQty") = 0 'Item.Fields("AvailableQty")
                sl.fields("QtyBeforeReorder") = item.fields("QtyBeforeReorder")
                sl.fields("Remark") = item.fields("Remark")
                sl.fields("PendingAcceptQty") = 0
                sl.fields("AfterAcceptQty") = 0
                sl.fields("StockValue1") = 0
                sl.fields("StockValue2") = 0
                sl.fields("ExpiryDate") = dt
                sl.fields("StockDate1") = dt
                sl.fields("StockDate2") = dt
                sl.fields("StockInfo1") = "-"
                sl.fields("ItemStockStatusID") = "D001"
                sl.fields("StockInfo2") = "-"
                sl.fields("ItemAdjustStatusID") = "D001"
                sl.fields("AdjustValue1") = 0
                sl.fields("AdjustValue2") = 0
                sl.fields("AdjustValue3") = 0
                sl.fields("AdjustDate1") = dt
                sl.fields("AdjustDate2") = dt

                sl.UpdateBatch
            End If
            sl.Close

            rst.MoveNext
        Loop
        rst.Close
    End If

    AddInventoryAccount "Item", item.fields("ItemID"), item.fields("ItemName")
    AddInventoryLookup "Item", item.fields("ItemID"), item.fields("ItemName")

    Set rst = Nothing
End Sub
Sub AddInventoryLookup(TableID, pk, name)
    Dim sql, lkpId

    If TableID = "DRUG" Then
        lkpId = "DRG-"
    ElseIf TableID = "ITEMS" Then
        lkpId = "ITM-"
    End If

    If Len(lkpId) > 0 Then
        lkpId = lkpId & pk

        sql = "if not exists(select * from InventLookup where InventLookupID='" & lkpId & "' )"
        sql = sql & " insert into InventLookup(InventLookupID, InventLookupName, TableID, KeyValue, InventAccountID, Description, KeyPrefix) "
        sql = sql & " values ('" & lkpId & "', '" & (TableID & " - [" & name & "]") & "', '" & TableID & "', '" & pk & "', '" & lkpId & "', 'System Generated', null) "

        conn.execute sql
    End If

End Sub
Sub AddInventoryAccount(TableID, pk, name)
    Dim sql, lkpId

    If TableID = "DRUG" Then
        lkpId = "DRG-"
    ElseIf TableID = "ITEMS" Then
        lkpId = "ITM-"
    End If

    If Len(lkpId) > 0 Then
        lkpId = lkpId & pk

        sql = "if not exists(select * from InventAccount where InventLookupID='" & lkpId & "' )"
        sql = sql & " insert into InventAccount(InventAccountID, InventAccountName, OpeningBalance, Description, KeyPrefix) "
        sql = sql & " values ('" & lkpId & "', '" & (TableID & " - [" & name & "]") & "', 0, 'System Generated', null) "

        conn.execute sql
    End If
End Sub
Function MigrateItems()
    Dim sql, rst, lastID, iRst, stkRst, sale, pk

    lastID = GetSystemVariableValue("focosItemsLastID")
    If lastID = "" Then
        lastID = 0
    End If
    sql = "select f_itemmaster.*, f_subcategorymaster.Name as subcategory_name, f_subcategorymaster.SubCategoryID "
    sql = sql & " , f_categorymaster.Name as CategoryName, f_categorymaster.CategoryID "
    sql = sql & " from f_itemmaster "
    sql = sql & " left join f_subcategorymaster on f_subcategorymaster.SubCategoryID=f_itemmaster.SubCategoryID "
    sql = sql & " left join f_categorymaster on f_categorymaster.CategoryID=f_subcategorymaster.CategoryID"
    sql = sql & " where f_subcategorymaster.CategoryID='LSHHI8' " 'and length(f_itemmaster.ItemCode) >0"
    sql = sql & "   or f_itemmaster.SubCategoryID in ('LSHHI100', 'LSHHI101', 'LSHHI174', 'LSHHI176', 'LSHHI34' "
    sql = sql & "           , 'LSHHI38', 'LSHHI40', 'LSHHI42', 'LSHHI51', 'LSHHI52', 'LSHHI55', 'LSHHI59', 'LSHHI61' "
    sql = sql & "           , 'LSHHI64', 'LSHHI65', 'LSHHI85', 'LSHHI84', 'LSHHI86', 'LSHHI87', 'LSHHI88', 'LSHHI89' "
    sql = sql & "           , 'LSHHI92', 'LSHHI93', 'LSHHI95', 'LSHHI96', 'LSHHI97', 'LSHHI98' ) "
    sql = sql & "   and f_itemmaster.ID>= " & lastID
    sql = sql & " order by f_itemmaster.ID asc"

    Set rst = CreateObject("ADODB.RecordSet")
    Set stkRst = CreateObject("ADODB.RecordSet")
    Set sale = CreateObject("ADODB.RecordSet")

    rst.open sql, GetMySQLConnection, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Set iRst = CreateObject("ADODB.RecordSet")
        Do While Not rst.EOF
            pk = Trim(rst.fields("ItemCode"))
            If Not (Len(pk) > 0) Or IsNull(pk) Then
                pk = rst.fields("ItemID")
            End If

            sql = "select * from Items where ItemID='" & pk & "' "
            iRst.open qryPro.FltQry(sql), conn, 3, 4
            If iRst.RecordCount = 0 Then
                iRst.AddNew
            End If
            If True Then
                sql = "select * from f_stock where ItemID='" & rst.fields("ItemID") & "' order by PostDate desc limit 1"
                stkRst.open sql, GetMySQLConnection, 3, 4
                If stkRst.RecordCount > 0 Then
                    stkRst.movefirst
                End If
                sql = "select * from f_salesdetails where ItemID='" & rst.fields("ItemID") & "' order by f_salesdetails.Date desc, f_salesdetails.Time desc limit 1"
                sale.open sql, GetMySQLConnection, 3, 4
                If sale.RecordCount > 0 Then
                    sale.movefirst
                End If
                
                iRst.fields("ItemID") = pk
                iRst.fields("ItemName") = rst.fields("TypeName")
                iRst.fields("ItemCategoryID") = GetItemCategory(rst.fields("CategoryName"), rst.fields("CategoryID"))
                iRst.fields("ItemTypeID") = GetItemType(rst.fields("subcategory_name"), rst.fields("subcategoryID"), iRst.fields("ItemCategoryID"))
                iRst.fields("UnitOfMeasureID") = rst.fields("UnitType")
                iRst.fields("ItemStatusID") = "IST001"
                iRst.fields("ItemLocationID") = "ILC002"
                iRst.fields("BillGroupCatID") = "NONE"
                iRst.fields("BillGroupID") = "B11"
                iRst.fields("BulkUnitCost") = IIF(sale.RecordCount > 0, sale.fields("PerUnitBuyPrice"), 0)
                iRst.fields("RetailUnitCost") = IIF(sale.RecordCount > 0, sale.fields("PerUnitSellingPrice"), 0)
                iRst.fields("MaxStockQty") = rst.fields("MaxLevel")
                iRst.fields("ReOrderModeID") = "MOD01"
                iRst.fields("ReOrderLevel") = rst.fields("ReorderLevel")
                iRst.fields("ReOrderLevelQty") = rst.fields("ReorderQty")
                iRst.fields("QtyBeforeReorder") = rst.fields("MinReorderQty")
                iRst.fields("ReOrderQty") = rst.fields("ReorderQty")
                iRst.fields("AvailableQty") = IIF(IsNull(rst.fields("QtyInHand")), 0, rst.fields("QtyInHand"))
                iRst.fields("Rate1") = 0
                iRst.fields("ItemGroupID") = "ADM"
                iRst.fields("Remark") = ""
                iRst.fields("ItemBrandID") = "B001"
                iRst.fields("Rate2") = 0
                iRst.fields("ItemModelID") = "I001"
                iRst.fields("KeyPrefix") = ""
                iRst.fields("ItemModeID") = "I001"
                iRst.fields("Rate3") = 0
                iRst.fields("Rate4") = 0
                'dRst.Fields("ProdDate1") = ""
                'dRst.Fields("ProdDate2") = ""
                'dRst.Fields("ProdInfo1") = ""
                'dRst.Fields("ProdInfo2") = ""

                iRst.UpdateBatch
                sql = "if not exists(select * from UnitOfMeasure where UnitOfMeasureID='" & rst.fields("UnitType") & "')"
                sql = sql & " insert into UnitOfMeasure(UnitOfMeasureID, UnitOfMeasureName)"
                sql = sql & " values('" & rst.fields("UnitType") & "', '" & rst.fields("UnitType") & "') "
                conn.execute qryPro.FltQry(sql)

                
                stkRst.Close
                sale.Close
            End If
            
            Call SetSystemVariableValue("focosItemsLastID", rst.fields("ID"))

            UpdateItemMatrixFromTemplate iRst.fields("ItemID"), iRst.fields("RetailUnitCost")
            'AddItemInventory iRst
            iRst.Close
            rst.MoveNext
        Loop
        rst.Close
    End If
    Set rst = Nothing
End Function
Sub UpdateItemMatrixFromTemplate(pk, cst)
    Dim html, sql, rstItemStTyp, insType, itmStTyp, rstIns, target, isValid, ItemVal, rstRecTyp
    Dim errorFields, fields, matSQL, tmpSQL, ItmStTypVal, insTypeVal

    isValid = True

    Set rstItemStTyp = CreateObject("ADODB.RecordSet")
    Set rstIns = CreateObject("ADODB.RecordSet")
    Set rstRecTyp = CreateObject("ADODB.RecordSet")
    
    sql = "select ItemStoreTypeName, ItemStoreTypeID, ItemPricePercent2 from ItemStoreType;"
    rstItemStTyp.open qryPro.FltQry(sql), conn, 3, 4

    sql = "select * from RecipientType;"
    rstRecTyp.open qryPro.FltQry(sql), conn, 3, 4
    
    'sql = " select InsuranceTypeID, InsuranceTypeName, ItemPricePercent from InsuranceType where InsuranceTypeName<>'-' "
    sql = " select distinct InsuranceType.InsuranceTypeID, InsuranceType.InsuranceTypeName, InsuranceType.ItemPricePercent "
    sql = sql & " from InsuranceType "
    sql = sql & " left join InsuranceScheme on InsuranceScheme.InsuranceTypeID=InsuranceType.InsuranceTypeID "
    sql = sql & " where InsuranceTypeName<>'-' and InsuranceScheme.InsuranceTypeID is not null "
    
    rstIns.open qryPro.FltQry(sql), conn, 3, 4

    If rstIns.RecordCount > 0 Then
        rstIns.movefirst
        Do While Not rstIns.EOF
            If rstItemStTyp.RecordCount > 0 Then
                rstItemStTyp.movefirst
                Do While Not rstItemStTyp.EOF
                    insType = rstIns.fields("InsuranceTypeID")
                    itmStTyp = rstItemStTyp.fields("ItemStoreTypeID")

                    If Not IsValidNumber(cst) Then
                        isValid = isValid And False
                    End If
                    If isValid Then
                        cst = CDbl(cst)
                        ItemVal = GetComboNameFld("Items", pk, "RetailUnitCost")
                        ItmStTypVal = rstItemStTyp.fields("ItemPricePercent2")
                        insTypeVal = rstIns.fields("ItemPricePercent")
                        
                        If rstRecTyp.RecordCount > 0 Then
                            rstRecTyp.movefirst
                            Do While Not rstRecTyp.EOF
                                tmpSQL = vbCrLf & " if exists(select * from ItemPriceMatrix2 where ItemID='" & pk & "' and RecipientTypeID='" & rstRecTyp.fields("RecipientTypeID") & "'"
                                tmpSQL = tmpSQL & "   and InsuranceTypeID='" & insType & "'  and ItemStoreTypeID='" & itmStTyp & "' "
                                tmpSQL = tmpSQL & " )"
                                tmpSQL = tmpSQL & "   update ItemPriceMatrix2 set ItemUnitCost=" & cst & " where ItemID='" & pk & "'  and RecipientTypeID='" & rstRecTyp.fields("RecipientTypeID") & "'"
                                tmpSQL = tmpSQL & "   and InsuranceTypeID='" & insType & "'  and ItemStoreTypeID='" & itmStTyp & "' "
                                tmpSQL = tmpSQL & " "
                                tmpSQL = tmpSQL & vbCrLf & " else "
                                tmpSQL = tmpSQL & " insert into ItemPriceMatrix2 (Column2, Column3, Column4, Column5, RecipientTypeID, Column6, Column1, ItemID, RecipientTypeVal, ItemStoreTypeID, InsuranceTypeID, ItemsVal, ItemStoreTypeVal, ItemUnitCost, PermuteUpdateID, PermuteStatusID, InsuranceTypeVal, KeyPrefix)"
                                tmpSQL = tmpSQL & " values ('-', '-', '-', '-', '" & rstRecTyp.fields("RecipientTypeID") & "', '-', '-', '" & pk & "', " & rstRecTyp.fields("PriceMatrixValue") & ", '" & itmStTyp & "', '" & insType & "', " & ItemVal & ", " & ItmStTypVal & ", " & cst & ", 'P001', 'P001', " & insTypeVal & ", NULL);"
        
                                matSQL = matSQL & tmpSQL
                                rstRecTyp.MoveNext
                            Loop
                        End If
                    End If
                    rstItemStTyp.MoveNext
                Loop
            End If
            rstIns.MoveNext
        Loop
    End If
    If isValid Then
        'SetPageMessages matSQL
        conn.execute qryPro.FltQry(matSQL)
    End If
End Sub
Function GetItemType(typename, TypeId, catID)
    Dim ot, rst, sql

    sql = "select * from ItemType where ItemTypeName='" & typename & "' and ItemCategoryID='" & catID & "' "
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        ot = rst.fields("ItemTypeID")
    Else
        ot = TypeId
        rst.AddNew
        rst.fields("ItemTypeID") = TypeId
        rst.fields("ItemTypeName") = typename
        rst.fields("ItemCategoryID") = catID
        rst.UpdateBatch
    End If

    rst.Close
    Set rst = Nothing

    GetItemType = ot
End Function
Function GetItemCategory(catName, catID)
    Dim sql, rst, ot
    ot = "G11"
    sql = " select * from ItemCategory where ItemCategoryName='" & catName & "' "
    sql = sql & " "
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        ot = rst.fields("ItemCategoryID")
    Else
        ot = catID
        rst.AddNew
        rst.fields("ItemCategoryID") = catID
        rst.fields("ItemCategoryName") = catName
        rst.UpdateBatch
    End If
    rst.Close
    Set rst = Nothing
    GetItemCategory = ot
End Function
Sub UpdateDrugMatrixFromTemplate(pk, cst)
    Dim html, sql, rstDrgStTyp, insType, drgStTyp, rstIns, target, isValid, drgVal, drgStTypVal, insTypeVal
    Dim errorFields, fields, matSQL, tmpSQL

    isValid = True

    Set rstDrgStTyp = CreateObject("ADODB.RecordSet")
    Set rstIns = CreateObject("ADODB.RecordSet")

    sql = "select DrugStoreTypeName, DrugStoreTypeID, DrugPricePercent2 from DrugStoreType"
    rstDrgStTyp.open qryPro.FltQry(sql), conn, 3, 4

    sql = " select distinct InsuranceType.InsuranceTypeID, InsuranceType.InsuranceTypeName, InsuranceType.DrugPricePercent "
    sql = sql & " from InsuranceType "
    sql = sql & " left join InsuranceScheme on InsuranceScheme.InsuranceTypeID=InsuranceType.InsuranceTypeID "
    sql = sql & " where InsuranceTypeName<>'-' and InsuranceScheme.InsuranceTypeID is not null "
    
    rstIns.open qryPro.FltQry(sql), conn, 3, 4

    If rstIns.RecordCount > 0 Then
        rstIns.movefirst
        Do While Not rstIns.EOF
            If rstDrgStTyp.RecordCount > 0 Then
                rstDrgStTyp.movefirst
                Do While Not rstDrgStTyp.EOF
                    insType = rstIns.fields("InsuranceTypeID")
                    drgStTyp = rstDrgStTyp.fields("DrugStoreTypeID")


                    If Not IsValidNumber(cst) Then
                        isValid = isValid And False
                    End If
                    If isValid Then
                        cst = CDbl(cst)
                        drgVal = GetComboNameFld("Drug", pk, "RetailUnitCost")
                        drgStTypVal = rstDrgStTyp.fields("DrugPricePercent2")
                        insTypeVal = rstIns.fields("DrugPricePercent")

                        tmpSQL = vbCrLf & " if exists(select * from DrugPriceMatrix2 where DrugID='" & pk & "' "
                        tmpSQL = tmpSQL & "   and InsuranceTypeID='" & insType & "'  and DrugStoreTypeID='" & drgStTyp & "' "
                        tmpSQL = tmpSQL & " )"
                        tmpSQL = tmpSQL & "   update DrugPriceMatrix2 set ItemUnitCost=" & cst & " where DrugID='" & pk & "' "
                        tmpSQL = tmpSQL & "   and InsuranceTypeID='" & insType & "'  and DrugStoreTypeID='" & drgStTyp & "' "
                        tmpSQL = tmpSQL & " "
                        tmpSQL = tmpSQL & vbCrLf & " else "
                        tmpSQL = tmpSQL & "   insert into DrugPriceMatrix2(DrugID, InsuranceTypeID, DrugStoreTypeID, PermuteUpdateID, PermuteStatusID, ItemUnitCost, DrugVal, InsuranceTypeVal, DrugStoreTypeVal)"
                        tmpSQL = tmpSQL & "   values('" & pk & "', '" & insType & "', '" & drgStTyp & "', 'P002', 'P001', " & cst & ", " & drgVal & ", " & insTypeVal & ", " & drgStTypVal & "); "

                        matSQL = matSQL & tmpSQL
                    End If
                    rstDrgStTyp.MoveNext
                Loop
            End If
            rstIns.MoveNext
        Loop
    End If
    If isValid Then
        conn.execute qryPro.FltQry(matSQL)
    End If
End Sub
Function GetDrugType(typename, TypeId)
    Dim sql, rst, ot

    sql = "select * from DrugType where DrugTypeName = '" & typename & "'"
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount > 0 Then
        ot = rst.fields("DrugTypeID")
    Else
        rst.AddNew
        ot = TypeId

        rst.fields("DrugTypeID") = ot
        rst.fields("DrugTypeName") = typename
        rst.fields("DrugCategoryID") = "DRC001"
        rst.fields("Description") = ""
        rst.fields("KeyPrefix") = "/"
        rst.UpdateBatch
        rst.Close
    End If

    GetDrugType = ot
End Function
Sub MigrateInvestigations()
    Dim sql, rst, lt, upDt, field, lastID, nPfx, tmpRst
 
   lastID = GetSystemVariableValue("focosInvestigationLastID")
    If lastID = "" Then
        lastID = 0
    End If
    sql = "select * from (select f_itemmaster.*, f_subcategorymaster.Name as subcategory_name "
    sql = sql & " , f_categorymaster.Name as CategoryName, f_categorymaster.CategoryID "
    sql = sql & " , f_ratelist.Rate, row_number() over(partition by f_itemmaster.itemID order by f_ratelist.updatedate desc, f_ratelist.EntryDate desc) as idx "
    sql = sql & " from f_itemmaster "
    sql = sql & " left join f_subcategorymaster on f_subcategorymaster.SubCategoryID=f_itemmaster.SubCategoryID "
    sql = sql & " left join f_categorymaster on f_categorymaster.CategoryID=f_subcategorymaster.CategoryID"
    sql = sql & " left join f_ratelist on f_ratelist.ItemID=f_itemmaster.ItemID"
    sql = sql & " where f_subcategorymaster.CategoryID in ('LSHHI3', 'LSHHI7', 'LSHHI18') and f_itemmaster.IsActive=1 and length(f_itemmaster.ItemCode)>0"
    sql = sql & "   and f_itemmaster.ID>= " & lastID
    sql = sql & " ) as itm where idx=1 order by itm.ID asc"

    Set rst = CreateObject("ADODB.RecordSet")
    Set tmpRst = CreateObject("ADODB.RecordSet")
    rst.open sql, GetMySQLConnection, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            Set lt = CreateObject("ADODB.RecordSet")
            sql = "select * from LabTest where LabTestID='" & rst.fields("ItemCode") & "' "
            lt.open qryPro.FltQry(sql), conn, 3, 4
            If lt.RecordCount = 0 Then
                lt.AddNew
            End If
            If True Then
                lt.fields("LabTestID") = rst.fields("ItemCode")
                If UCase(rst.fields("CategoryID")) = "LSHHI7" Then
                    lt.fields("TestCategoryID") = "B19" 'RAD
                    lt.fields("TestGroupID") = "B19"
                    lt.fields("BillGroupCatID") = "B19"
                    lt.fields("TestContainerID") = "DPT011"
                    lt.fields("BillGroupID") = "B19"
                Else
                    lt.fields("TestCategoryID") = "B13" 'LAB
                    lt.fields("TestGroupID") = "B13"
                    lt.fields("BillGroupCatID") = "B13"
                    lt.fields("TestContainerID") = "DPT005"
                    lt.fields("BillGroupID") = "B13"
                End If

                lt.fields("LabTestName") = rst.fields("TypeName")
                lt.fields("TestTypeID") = rst.fields("SubcategoryID")
                lt.fields("TestClassID") = "T001"
                lt.fields("LevelOfAccessID") = "All"
                lt.fields("UnitCost") = IIF(IsNull(rst.fields("Rate")), 0, rst.fields("Rate"))
                lt.fields("TestStatusID") = "TST001"
                lt.fields("ResultTypeID") = "RTY001"
                lt.fields("TestSampleTypeID") = "T001"
                lt.fields("TestDuration") = 0
                lt.fields("TestAmt1") = 0
                lt.fields("TestAmt2") = 0
                lt.fields("LabTestVal1") = 0
                lt.fields("LabTestVal2") = 0
                lt.fields("LabTestVal3") = 0
                lt.fields("LabTestVal4") = 0
                lt.fields("LabTestInfo1") = ""
                lt.fields("LabTestInfo2") = ""
                'lt.Fields("LabTestDate1") = ""
                'lt.Fields("LabTestDate2") = ""
                lt.fields("Description") = ""
                lt.fields("KeyPrefix") = ""
                lt.fields("BillInput1") = 0
                lt.fields("BillInput2") = 0
                lt.fields("BillInput3") = 0
                lt.fields("BillInout4") = 0
            
                lt.UpdateBatch

                sql = "select * from TestType where TestTypeID='" & lt.fields("TestTypeID") & "'"
                tmpRst.open qryPro.FltQry(sql), conn, 3, 4
                If tmpRst.RecordCount = 0 Then
                    tmpRst.AddNew
                        
                    tmpRst.fields("TestTypeID") = lt.fields("TestTypeID")
                    tmpRst.fields("TestTypeName") = rst.fields("subcategory_name")
                    tmpRst.fields("TestCategoryID") = lt.fields("TestCategoryID")
                    tmpRst.fields("Description") = ""
                    tmpRst.fields("KeyPrefix") = ""

                    tmpRst.UpdateBatch
                End If
                tmpRst.Close

                UpdateLabTestMatrixFromTemplate lt.fields("LabTestID"), lt.fields("UnitCost")
                    
            End If

            Call SetSystemVariableValue("focosInvestigationLastID", rst.fields("ID"))
            lt.Close
            rst.MoveNext
        Loop
        rst.Close
    End If
End Sub
Sub UpdateLabTestMatrixFromTemplate(pk, cst)
    Dim html, sql, rstAge, rstVstTyp, insType, vstTyp, agegrp, rstIns, target, isValid
    Dim errorFields, matSQL, tmpSQL, insTypeVal, ageGrpVal, VstTypVal, ltCst
    
    isValid = True
    
    Set rstAge = CreateObject("ADODB.RecordSet")
    Set rstIns = CreateObject("ADODB.RecordSet")
    
    sql = "select AgeGroupID, AgeGroupName, LabTestPercent from AgeGroup"
    rstAge.open qryPro.FltQry(sql), conn, 3, 4
    
    sql = " select InsuranceTypeID, InsuranceTypeName, LabTestPercent from InsuranceType where InsuranceTypeName<>'-' "
    rstIns.open qryPro.FltQry(sql), conn, 3, 4
    
    If rstIns.RecordCount > 0 Then
        rstIns.movefirst
        Do While Not rstIns.EOF
            If rstAge.RecordCount > 0 Then
                rstAge.movefirst
                Do While Not rstAge.EOF
                    insType = rstIns.fields("InsuranceTypeID")
                    agegrp = rstAge.fields("AgeGroupID")
                    
                    If True Then 'field exists
                        agegrp = rstAge.fields("AgeGroupID")
                        
                        isValid = True
                        If isValid Then
                            cst = CDbl(cst)
                            ltCst = GetComboNameFld("LabTest", pk, "UnitCost")
                            ageGrpVal = rstAge.fields("LabTestPercent")
                            insTypeVal = rstIns.fields("LabTestPercent")
                            
                            tmpSQL = vbCrLf & " if exists(select * from LabTestCostMatrix where LabTestID='" & pk & "' "
                            tmpSQL = tmpSQL & "   and InsuranceTypeID='" & insType & "' and AgeGroupID='" & agegrp & "' "
                            tmpSQL = tmpSQL & " )"
                            tmpSQL = tmpSQL & "   update LabTestCostMatrix set UnitCost=" & cst & " where LabTestID='" & pk & "' "
                            tmpSQL = tmpSQL & "   and InsuranceTypeID='" & insType & "' and AgeGroupID='" & agegrp & "' "
                            tmpSQL = tmpSQL & " "
                            tmpSQL = tmpSQL & vbCrLf & " else "
                            tmpSQL = tmpSQL & "   insert into LabTestCostMatrix(LabTestID, InsuranceTypeID, AgeGroupID, PermuteUpdateID, PermuteStatusID, UnitCost, LabTestVal, AgeGroupVal, InsuranceTypeVal)"
                            tmpSQL = tmpSQL & "   values('" & pk & "', '" & insType & "',  '" & agegrp & "', 'P002', 'P001', " & cst & ", " & ltCst & ", " & ageGrpVal & ", " & insTypeVal & "); "
                            
                            matSQL = matSQL & tmpSQL
                        End If
                    End If
                    rstAge.MoveNext
                Loop
            End If
            rstIns.MoveNext
        Loop
    End If
    If isValid Then
        conn.execute qryPro.FltQry(matSQL)
    End If
End Sub
Sub MigrateDisease()
    Dim sql, rst, lastID, iRst, dCat, dGroup, dTyp, sale, pk, dRst

    lastID = GetSystemVariableValue("focosDiseaseLastID")
    If lastID = "" Then
        lastID = 0
    End If
    sql = "select icd_10_new.* "
    sql = sql & " from icd_10_new "
    sql = sql & "   where icd_10_new.ICD10_Code > '" & lastID & "' "
    sql = sql & " order by icd_10_new.ICD10_Code asc"

    Set rst = CreateObject("ADODB.RecordSet")
    Set dCat = CreateObject("ADODB.RecordSet")
    Set dGroup = CreateObject("ADODB.RecordSet")
    Set dTyp = CreateObject("ADODB.RecordSet")
    Set dRst = CreateObject("ADODB.RecordSet")
    rst.open sql, GetMySQLConnection, 3, 4

    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            sql = "select * from DiseaseCategory where DiseaseCategoryID='" & Clean(rst.fields("Chapter_No")) & "' "
            dCat.open sql, conn, 3, 4
            If dCat.RecordCount = 0 Then
                dCat.AddNew
                dCat.fields("DiseaseCategoryID") = Clean(rst.fields("Chapter_No"))
                dCat.fields("DiseaseCategoryName") = rst.fields("Chapter_Desc")
                dCat.UpdateBatch
            End If

            sql = "select * from DiseaseGroup where DiseaseGroupID='" & Clean(rst.fields("Group_Code")) & "' "
            dGroup.open sql, conn, 3, 4
            If dGroup.RecordCount = 0 Then
                dGroup.AddNew
                dGroup.fields("DiseaseGroupID") = Clean(rst.fields("Group_Code"))
                dGroup.fields("DiseaseGroupName") = rst.fields("Group_Desc")
                dGroup.UpdateBatch
            End If

            sql = "select * from DiseaseType where DiseaseTypeID='" & Clean(rst.fields("icd10_3_code")) & "'"
            dTyp.open sql, conn, 3, 4
            If dTyp.RecordCount = 0 Then
                dTyp.AddNew
                dTyp.fields("DiseaseTypeID") = Clean(rst.fields("icd10_3_code"))
                dTyp.fields("DiseaseTypeName") = rst.fields("icd10_3_code_desc")
                dTyp.fields("DiseaseCategoryID") = Clean(rst.fields("Chapter_No"))
                'dTyp.Fields("DiseaseGroupID") = rst.Fields("Group_Code")
                dTyp.fields("BillGroupID") = "B14"
                dTyp.UpdateBatch
            End If

            If UCase(rst.fields("valid_icd10_clinicaluse")) = "Y" Then
                sql = "select * from Disease where DiseaseID='" & Clean(rst.fields("icd10_code")) & "' "
                dRst.open sql, conn, 3, 4
                If dRst.RecordCount = 0 Then
                    dRst.AddNew
                    dRst.fields("DiseaseID") = Clean(rst.fields("icd10_code"))
                    dRst.fields("DiseaseName") = rst.fields("who_full_desc")
                    dRst.fields("DiseaseTypeID") = Clean(rst.fields("icd10_3_code"))
                    dRst.fields("DiseaseCategoryID") = Clean(rst.fields("Chapter_No"))
                    dRst.fields("DiseaseGroupID") = Clean(rst.fields("Group_Code"))
                    dRst.fields("BillGroupID") = "B14"
                    dRst.fields("BillGroupCatID") = "NONE"
                    dRst.fields("DiseaseClassID") = "D001"

                    dRst.UpdateBatch
                End If
                dRst.Close

            End If

            dCat.Close
            dGroup.Close
            dTyp.Close

            Call SetSystemVariableValue("focosDiseaseLastID", rst.fields("icd10_code"))
            rst.MoveNext
        Loop
        rst.Close
    End If
    Set rst = Nothing
End Sub
Function Clean(str)
    Clean = Replace(Replace(Replace(str, ")", "_"), "(", "_"), " ", "_")
End Function
Sub AddProcesscallSibling(oldUsrProc, newUsrProc, TableID)
    Dim sql, rst, rst2
    
    sql = "select * from ProcessCall where UserProcessID='" & oldUsrProc & "'"
    If Len(TableID) > 0 Then
        sql = sql & " and TableID='" & TableID & "'"
    End If
    
    Set rst = CreateObject("ADODB.RecordSet")
    Set rst2 = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            AddProcessCall rst.fields("TableID"), newUsrProc, rst.fields("ProcessPointID")
            rst.MoveNext
        Loop
    End If
    Set rst = Nothing
    Set rst2 = Nothing
End Sub
Sub GrantAccessToTable(profile, tableName, uRight)
    Dim sql, rst, accessRight, userRight
    accessRight = ""
    
    Select Case UCase(uRight)
        Case "VIEW"
            userRight = "URT001"
            accessRight = "frm" & tableName & userRight
        Case "NEW"
            userRight = "URT002"
            accessRight = "frm" & tableName & userRight
        Case "EDIT"
            userRight = "URT003"
            accessRight = "frm" & tableName & userRight
        Case "SAVE"
            userRight = "URT004"
            accessRight = "frm" & tableName & userRight
        Case "DELETE"
            userRight = "URT005"
            accessRight = "frm" & tableName & userRight
        Case "SEARCH"
            userRight = "URT006"
            accessRight = "frm" & tableName & userRight
    End Select
    
    If Len(accessRight) > 0 Then
        sql = "select * from AccessRightAlloc where UserRoleID='" & profile & "' and AccessRightID='" & accessRight & "' "
        Set rst = CreateObject("ADODB.RecordSet")
        rst.open sql, conn, 3, 4
        If rst.RecordCount = 0 Then
            rst.AddNew
            rst.fields("UserRoleID") = profile
            rst.fields("AccessRightID") = accessRight
            rst.fields("UserRightID") = userRight
            rst.fields("TableID") = tableName
            rst.fields("AccessDetail") = "YES"
            rst.UpdateBatch
        End If
        rst.Close
    End If
    
End Sub


    





Sub DeleteOldData()
    Dim sql, rst, dt
    
    dt = "27 Sep 2022 23:59:59"
    
    sql = "select * from DrugAdjustment  where  AdjustDate<'" & dt & "'"
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        Do While Not rst.EOF
            DeleteTableKey "DrugAdjustment", rst.fields("DrugAdjustmentID")
            rst.MoveNext
        Loop
    End If
    
    sql = "select * from StockAdjustment  where  AdjustDate<'" & dt & "'"
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        Do While Not rst.EOF
            DeleteTableKey "StockAdjustment", rst.fields("StockAdjustmentID")
            rst.MoveNext
        Loop
    End If
    
    sql = "select * from DrugSale  where  DispenseDate<'" & dt & "'"
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        Do While Not rst.EOF
            DeleteTableKey "DrugSale", rst.fields("DrugSaleID")
            rst.MoveNext
        Loop
    End If
    
    sql = "select * from StockIssue  where  IssueDate<'" & dt & "'"
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        Do While Not rst.EOF
            DeleteTableKey "StockIssue", rst.fields("StockIssueID")
            rst.MoveNext
        Loop
    End If
    
    sql = "select * from LabRequest  where  RequestDate<'" & dt & "'"
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        Do While Not rst.EOF
            DeleteTableKey "LabRequest", rst.fields("LabRequestID")
            rst.MoveNext
        Loop
    End If
    
    sql = "select * from ConsultReview  where  ConsultReviewDate<'" & dt & "'"
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        Do While Not rst.EOF
            DeleteTableKey "ConsultReview", rst.fields("ConsultReviewID")
            rst.MoveNext
        Loop
    End If
    
    sql = "select * from DrugPurOrder  where  PurchaseOrderDate<'" & dt & "'"
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        Do While Not rst.EOF
            DeleteTableKey "DrugPurOrder", rst.fields("DrugPurOrderID")
            rst.MoveNext
        Loop
    End If
    
    sql = "select * from ItemPurOrder  where  PurchaseOrderDate<'" & dt & "'"
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        Do While Not rst.EOF
            DeleteTableKey "ItemPurOrder", rst.fields("ItemPurOrderID")
            rst.MoveNext
        Loop
    End If
    
    sql = "select * from DrugRequest2  where  RequestDate<'" & dt & "'"
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        Do While Not rst.EOF
            DeleteTableKey "DrugRequest2", rst.fields("DrugRequest2ID")
            rst.MoveNext
        Loop
    End If
    
    sql = "select * from ItemRequest2  where  RequestDate<'" & dt & "'"
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        Do While Not rst.EOF
            DeleteTableKey "ItemRequest2", rst.fields("ItemRequest2ID")
            rst.MoveNext
        Loop
    End If
End Sub













Sub AddDrugInventory(ByVal drug)
    Dim sql, dt, rst, sl

    dt = Now
    sql = "select * from DrugStore where DrugStoreID='S22' "
    Set rst = CreateObject("ADODB.RecordSet")
    Set sl = CreateObject("ADODB.RecordSet")
    rst.open sql, conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            sql = "select * from DrugStockLevel where DrugID='" & drug.fields("DrugID") & "' and DrugStoreID='" & rst.fields("DrugStoreID") & "' "
            sl.open sql, conn, 3, 4
            If sl.RecordCount = 0 Then
                sl.AddNew

                sl.fields("DrugStoreTypeID") = rst.fields("DrugStoreTypeID")
                sl.fields("DrugStoreID") = rst.fields("DrugStoreID")
                sl.fields("DrugID") = drug.fields("DrugID")
                sl.fields("DrugStoreLocationID") = rst.fields("DrugStoreLocationID")
                sl.fields("DrugCategoryID") = drug.fields("DrugCategoryID")
                sl.fields("DrugTypeID") = drug.fields("DrugTypeID")
                sl.fields("UnitOfMeasureID") = drug.fields("UnitOfMeasureID")
                sl.fields("DrugStatusID") = drug.fields("DrugStatusID")
                sl.fields("DrugLocationID") = drug.fields("DrugLocationID")
                sl.fields("BulkUnitCost") = drug.fields("BulkUnitCost")
                sl.fields("RetailUnitCost") = drug.fields("RetailUnitCost")
                sl.fields("TotalCost") = 0
                sl.fields("MaxStockQty") = drug.fields("MaxStockQty")
                sl.fields("ReOrderModeID") = drug.fields("ReOrderModeID")
                sl.fields("ReOrderLevel") = drug.fields("ReOrderLevel")
                sl.fields("ReOrderLevelQty") = drug.fields("ReOrderLevelQty")
                sl.fields("ReOrderQty") = drug.fields("ReOrderQty")
                sl.fields("AvailableQty") = 0 'drug.Fields("AvailableQty")
                sl.fields("QtyBeforeReorder") = drug.fields("QtyBeforeReorder")
                sl.fields("Remark") = drug.fields("Remark")
                sl.fields("PendingAcceptQty") = 0
                sl.fields("AfterAcceptQty") = 0
                sl.fields("StockValue1") = 0
                sl.fields("StockValue2") = 0
                sl.fields("ExpiryDate") = dt
                sl.fields("StockDate1") = dt
                sl.fields("StockDate2") = dt
                sl.fields("StockInfo1") = "-"
                sl.fields("DrugStockStatusID") = "D001"
                sl.fields("StockInfo2") = "-"
                sl.fields("DrugAdjustStatusID") = "D001"
                sl.fields("AdjustValue1") = 0
                sl.fields("AdjustValue2") = 0
                sl.fields("AdjustValue3") = 0
                sl.fields("AdjustDate1") = dt
                sl.fields("AdjustDate2") = dt

                sl.UpdateBatch
            End If
            sl.Close

            rst.MoveNext
        Loop
        rst.Close
    End If

    'AddInventoryAccount "Drug", drug.fields("DrugID"), drug.fields("DrugName")
    'AddInventoryLookup "Drug", drug.fields("DrugID"), drug.fields("DrugName")

    Set rst = Nothing
End Sub