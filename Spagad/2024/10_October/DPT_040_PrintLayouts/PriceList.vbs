'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>


ShowPriceList

Sub ShowPriceList()
    Dim rptGen, sql, args

    Set rptGen = New PRTGLO_RptGen

    GetConsultConsultCostMatrix sql, args
    rptGen.AddReport sql, args

    GetDrugPriceMatrix sql, args
    rptGen.AddReport sql, args

    GetItemPriceMatrix sql, args
    rptGen.AddReport sql, args

    GetBedCostMatrix sql, args
    rptGen.AddReport sql, args

    GetTreatCostMatrix sql, args
    rptGen.AddReport sql, args

    GetLabTestPriceMatrix sql, args
    rptGen.AddReport sql, args
        
    GetRadTestPriceMatrix sql, args
    rptGen.AddReport sql, args

    rptGen.ShowReport
End Sub
Function GetActiveInsuranceTypes()
    Dim ot, rst, tmp, sql, field

    Set ot = CreateObject("Scripting.Dictionary")

    sql = "select distinct InsuranceType.InsuranceTypeID, InsuranceType.InsuranceTypeName "
    sql = sql & " from InsuranceType "
    sql = sql & " inner join InsuranceScheme on InsuranceScheme.InsuranceTypeID=InsuranceType.InsuranceTypeID"
    sql = sql & " and InsuranceScheme.InsSchemeStatusID='I001' "
    sql = sql & " order by InsuranceType.InsuranceTypeName"
    
    Set rst = CreateObject("ADODB.RecordSet")
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            Set tmp = CreateObject("Scripting.Dictionary")
            
            For Each field In rst.fields
                tmp.add field.name, field.value
            Next

            ot.add rst.fields("InsuranceTypeID").value, tmp
            rst.MoveNext
        Loop
    End If
    rst.Close

    Set GetActiveInsuranceTypes = ot
End Function
Sub GetRadTestPriceMatrix(ByRef sql, ByRef args)
    Dim ins, key, insType

    Set insType = GetActiveInsuranceTypes()
    sql = "select "
    sql = sql & " LabTestCostMatrix.LabTestID"
    sql = sql & " , LabTestCostMatrix.LabTestName"
    For Each key In insType.keys()
        Set ins = insType(key)
        sql = sql & " , max(case when LabTestCostMatrix.InsuranceTypeID='" & ins("InsuranceTypeID") & "' "
        sql = sql & "       then LabTestCostMatrix.UnitCost "
        sql = sql & "       else 0 "
        sql = sql & "   end) as ""[" & key & "] " & ins("InsuranceTypeName") & """ "
    Next
    sql = sql & " from (" & GetTableSql("LabTestCostMatrix") & ") as LabTestCostMatrix"
    sql = sql & "  inner join LabTest on LabTest.LabTestID=LabTestCostMatrix.LabTestID and LabTest.TestStatusID='TST001'"
    sql = sql & " where 1=1 and LabTest.TestContainerID='DPT011' and LabTest.TestStatusID='TST001' "
    sql = sql & "   and LabTestCostMatrix.InsuranceTypeID in ('" & Join(insType.keys(), "', '") & "') "
    sql = sql & " group by "
    sql = sql & " LabTestCostMatrix.LabTestID"
    sql = sql & " , LabTestCostMatrix.LabTestName"
    sql = sql & " order by LabTestCostMatrix.LabTestName, LabTestCostMatrix.LabTestID "

    args = "Title=Radiology Investigations"
    args = args & ";PageRecordCount=9999999"

End Sub
Sub GetLabTestPriceMatrix(ByRef sql, ByRef args)
    Dim ins, key, insType

    Set insType = GetActiveInsuranceTypes()
    sql = "select "
    sql = sql & " LabTestCostMatrix.LabTestID"
    sql = sql & " , LabTestCostMatrix.LabTestName"
    For Each key In insType.keys()
        Set ins = insType(key)
        sql = sql & " , max(case when LabTestCostMatrix.InsuranceTypeID='" & ins("InsuranceTypeID") & "' "
        sql = sql & "       then LabTestCostMatrix.UnitCost "
        sql = sql & "       else 0 "
        sql = sql & "   end) as ""[" & key & "] " & ins("InsuranceTypeName") & """ "
    Next
    sql = sql & " from (" & GetTableSql("LabTestCostMatrix") & ") as LabTestCostMatrix"
    sql = sql & "  inner join LabTest on LabTest.LabTestID=LabTestCostMatrix.LabTestID and LabTest.TestStatusID='TST001'"
    sql = sql & " where 1=1 and LabTest.TestContainerID='DPT005' and LabTest.TestStatusID='TST001' "
    sql = sql & "   and LabTestCostMatrix.InsuranceTypeID in ('" & Join(insType.keys(), "', '") & "') "
    sql = sql & " group by "
    sql = sql & " LabTestCostMatrix.LabTestID"
    sql = sql & " , LabTestCostMatrix.LabTestName"
    sql = sql & " order by LabTestCostMatrix.LabTestName, LabTestCostMatrix.LabTestID "

    args = "Title=Laboratory Investigations"
    args = args & ";PageRecordCount=9999999"

End Sub
Sub GetDrugPriceMatrix(ByRef sql, ByRef args)
    Dim ins, key, insType

    Set insType = GetActiveInsuranceTypes()
    sql = "select "
    sql = sql & " DrugPriceMatrix2.DrugID"
    sql = sql & " , DrugPriceMatrix2.DrugName"
    For Each key In insType.keys()
        Set ins = insType(key)
        sql = sql & " , max(case when DrugPriceMatrix2.InsuranceTypeID='" & ins("InsuranceTypeID") & "' "
        sql = sql & "       then DrugPriceMatrix2.ItemUnitCost "
        sql = sql & "       else 0 "
        sql = sql & "   end) as ""[" & key & "] " & ins("InsuranceTypeName") & """ "
    Next
    sql = sql & " from (" & GetTableSql("DrugPriceMatrix2") & ") as DrugPriceMatrix2"
    sql = sql & "  inner join Drug on Drug.DrugID=DrugPriceMatrix2.DrugID and Drug.DrugStatusID='IST001'"
    sql = sql & " where 1=1 and Drug.DrugStatusID='IST001' "
    sql = sql & "   and DrugPriceMatrix2.InsuranceTypeID in ('" & Join(insType.keys(), "', '") & "') "
    sql = sql & " group by "
    sql = sql & " DrugPriceMatrix2.DrugID"
    sql = sql & " , DrugPriceMatrix2.DrugName"
    sql = sql & " order by DrugPriceMatrix2.DrugName, DrugPriceMatrix2.DrugID "

    args = "Title=Medical Items"
    args = args & ";PageRecordCount=9999999"

End Sub
Sub GetItemPriceMatrix(ByRef sql, ByRef args)
    Dim ins, key, insType

    Set insType = GetActiveInsuranceTypes()
    sql = "select "
    sql = sql & " ItemPriceMatrix2.ItemID"
    sql = sql & " , ItemPriceMatrix2.ItemName"
    For Each key In insType.keys()
        Set ins = insType(key)
        sql = sql & " , max(case when ItemPriceMatrix2.InsuranceTypeID='" & ins("InsuranceTypeID") & "' "
        sql = sql & "       then ItemPriceMatrix2.ItemUnitCost "
        sql = sql & "       else 0 "
        sql = sql & "   end) as ""[" & key & "] " & ins("InsuranceTypeName") & """ "
    Next
    sql = sql & " from (" & GetTableSql("ItemPriceMatrix2") & ") as ItemPriceMatrix2"
    sql = sql & "  inner join Items on Items.ItemID=ItemPriceMatrix2.ItemID and Items.ItemStatusID='IST001'"
    sql = sql & " where 1=1 and Items.ItemStatusID='IST001' "
    sql = sql & "   and ItemPriceMatrix2.InsuranceTypeID in ('" & Join(insType.keys(), "', '") & "') "
    sql = sql & " group by "
    sql = sql & " ItemPriceMatrix2.ItemID"
    sql = sql & " , ItemPriceMatrix2.ItemName"
    sql = sql & " order by ItemPriceMatrix2.ItemName, ItemPriceMatrix2.ItemID "

    args = "Title=General Items"
    args = args & ";PageRecordCount=9999999"

End Sub
Sub GetTreatCostMatrix(ByRef sql, ByRef args)
    Dim ins, key, insType

    Set insType = GetActiveInsuranceTypes()
    sql = "select "
    sql = sql & " TreatCostMatrix.TreatmentID"
    sql = sql & " , TreatCostMatrix.TreatmentName"
    For Each key In insType.keys()
        Set ins = insType(key)
        sql = sql & " , max(case when TreatCostMatrix.InsuranceTypeID='" & ins("InsuranceTypeID") & "' "
        sql = sql & "       then TreatCostMatrix.UnitCost "
        sql = sql & "       else 0 "
        sql = sql & "   end) as ""[" & key & "] " & ins("InsuranceTypeName") & """ "
    Next
    sql = sql & " from (" & GetTableSql("TreatCostMatrix") & ") as TreatCostMatrix"
    sql = sql & " inner join Treatment on Treatment.TreatmentID=TreatCostMatrix.TreatmentID"
    sql = sql & "   and (Treatment.TreatInfo1 is null or Treatment.TreatInfo1<>'YES') "
    sql = sql & " where TreatCostMatrix.InsuranceTypeID in ('" & Join(insType.keys(), "', '") & "') "
    sql = sql & " group by "
    sql = sql & " TreatCostMatrix.TreatmentID"
    sql = sql & " , TreatCostMatrix.TreatmentName"
    sql = sql & " order by TreatCostMatrix.TreatmentName, TreatCostMatrix.TreatmentID "


    args = "Title=Services / Treatments"
    args = args & ";PageRecordCount=9999999"

End Sub
Sub GetConsultConsultCostMatrix(ByRef sql, ByRef args)
    Dim ins, key, insType

    Set insType = GetActiveInsuranceTypes()
    sql = "select "
    sql = sql & " ConsultCostMatrix.SpecialistTypeID"
    sql = sql & " , ConsultCostMatrix.SpecialistTypeName"
    sql = sql & " , '[' + ConsultCostMatrix.VisitTypeID + '] ' + ConsultCostMatrix.VisitTypeName as [Visit Type]"
    For Each key In insType.keys()
        Set ins = insType(key)
        sql = sql & " , max(case when ConsultCostMatrix.InsuranceTypeID='" & ins("InsuranceTypeID") & "' "
        sql = sql & "       then ConsultCostMatrix.VisitCost "
        sql = sql & "       else 0 "
        sql = sql & "   end) as ""[" & key & "] " & ins("InsuranceTypeName") & """ "
    Next
    sql = sql & " from (" & GetTableSql("ConsultCostMatrix") & ") as ConsultCostMatrix"
    sql = sql & " where ConsultCostMatrix.InsuranceTypeID in ('" & Join(insType.keys(), "', '") & "') "
    sql = sql & " group by "
    sql = sql & " ConsultCostMatrix.SpecialistTypeID"
    sql = sql & " , ConsultCostMatrix.SpecialistTypeName"
    sql = sql & " , ConsultCostMatrix.VisitTypeID"
    sql = sql & " , ConsultCostMatrix.VisitTypeName"
    sql = sql & " order by ConsultCostMatrix.SpecialistTypeName, ConsultCostMatrix.SpecialistTypeID, ConsultCostMatrix.VisitTypeID "


    args = "Title=Consultations"
    args = args & ";PageRecordCount=9999999"

End Sub
Sub GetBedCostMatrix(ByRef sql, ByRef args)
    Dim ins, key, insType

    Set insType = GetActiveInsuranceTypes()
    sql = "select "
    sql = sql & " Ward.WardName"
    sql = sql & " , BedCostMatrix.BedID"
    sql = sql & " , BedCostMatrix.BedName"
    For Each key In insType.keys()
        Set ins = insType(key)
        sql = sql & " , max(case when BedCostMatrix.InsuranceTypeID='" & ins("InsuranceTypeID") & "' "
        sql = sql & "       then BedCostMatrix.BedCharge "
        sql = sql & "       else 0 "
        sql = sql & "   end) as ""[" & key & "] " & ins("InsuranceTypeName") & """ "
    Next
    sql = sql & " from (" & GetTableSql("BedCostMatrix") & ") as BedCostMatrix"
    sql = sql & " inner join Bed on BedCostMatrix.BedID=Bed.BedID"
    sql = sql & " inner join Ward on Bed.WardID=Ward.WardID"
    sql = sql & " where BedCostMatrix.InsuranceTypeID in ('" & Join(insType.keys(), "', '") & "') "
    sql = sql & " group by "
    sql = sql & " BedCostMatrix.BedID"
    sql = sql & " , BedCostMatrix.BedName, Bed.WardID, Ward.WardName"
    sql = sql & " order by Bed.WardID, Ward.WardName, BedCostMatrix.BedName, BedCostMatrix.BedID "

    args = "Title=Bed Charges"
    args = args & ";PageRecordCount=9999999"

End Sub

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
