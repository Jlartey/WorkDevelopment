BEGIN 
  IF NOT EXISTS (
    SELECT * FROM EMRData WHERE ( EMRDataID='MED001') 
  ) 
  BEGIN 
    INSERT INTO EMRData ( EMRDataID, EMRDataName, EMRCategoryID, EMRTypeID, EMRGroupID, EMRClassID, BillGroupCatID, BillGroupID, LevelOfAccessID, UnitCost, EMRStatusID, ResultTypeID, EMRContainerID, EMRSampleTypeID, EMRDuration, EMRDataAmt1, EMRDataAmt2, EMRDataVal1, EMRDataVal2, EMRDataVal3, EMRDataVal4, EMRDataInfo1, EMRDataInfo2, EMRDataDate1, EMRDataDate2, Description, KeyPrefix, BillInput1, BillInput2, BillInput3, BillInout4) 
    VALUES ( 'MED001', 'MEDICATION AUDIT TOOL', 'E001', 'E001', 'E001', 'E001', 'COMPUCARE', 'B16', 'All', '0', 'E001', 'RTY003', 'E001', 'E001', '0', '0', '0', '0', '0', '0', '0', '', '', null, null, '', '', '0', '0', '0', '0') 
  END 
END 

