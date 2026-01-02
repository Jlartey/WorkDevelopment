BEGIN 
  IF NOT EXISTS (
    SELECT * FROM EMRComponent WHERE ( EMRComponentID='MED001.7') 
  ) 
  BEGIN 
    INSERT INTO EMRComponent ( EMRComponentID, EMRComponentName, CompPos, FlagPos1, Column1, ResultPos1, Column2, RefPos1, Column3, FlagPos2, Column4, ResultPos2, Column5, RefPos2, Column6, Description, KeyPrefix, EMRCompCatID, EMRCompTypeID, EMRCompStatusID, EMRCompModeID) 
    VALUES ( 'MED001.7', '&nbsp;', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'NONE', 'E001', 'E001', 'E001') 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM EMRComponent WHERE ( EMRComponentID='MED001.1') 
  ) 
  BEGIN 
    INSERT INTO EMRComponent ( EMRComponentID, EMRComponentName, CompPos, FlagPos1, Column1, ResultPos1, Column2, RefPos1, Column3, FlagPos2, Column4, ResultPos2, Column5, RefPos2, Column6, Description, KeyPrefix, EMRCompCatID, EMRCompTypeID, EMRCompStatusID, EMRCompModeID) 
    VALUES ( 'MED001.1', '&nbsp;', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'NONE', 'E001', 'E001', 'E001') 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM EMRComponent WHERE ( EMRComponentID='MED001.3') 
  ) 
  BEGIN 
    INSERT INTO EMRComponent ( EMRComponentID, EMRComponentName, CompPos, FlagPos1, Column1, ResultPos1, Column2, RefPos1, Column3, FlagPos2, Column4, ResultPos2, Column5, RefPos2, Column6, Description, KeyPrefix, EMRCompCatID, EMRCompTypeID, EMRCompStatusID, EMRCompModeID) 
    VALUES ( 'MED001.3', '&nbsp;', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'NONE', 'E001', 'E001', 'E001') 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM EMRComponent WHERE ( EMRComponentID='MED001.4') 
  ) 
  BEGIN 
    INSERT INTO EMRComponent ( EMRComponentID, EMRComponentName, CompPos, FlagPos1, Column1, ResultPos1, Column2, RefPos1, Column3, FlagPos2, Column4, ResultPos2, Column5, RefPos2, Column6, Description, KeyPrefix, EMRCompCatID, EMRCompTypeID, EMRCompStatusID, EMRCompModeID) 
    VALUES ( 'MED001.4', '&nbsp;', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'NONE', 'E001', 'E001', 'E001') 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM EMRComponent WHERE ( EMRComponentID='MED001.5') 
  ) 
  BEGIN 
    INSERT INTO EMRComponent ( EMRComponentID, EMRComponentName, CompPos, FlagPos1, Column1, ResultPos1, Column2, RefPos1, Column3, FlagPos2, Column4, ResultPos2, Column5, RefPos2, Column6, Description, KeyPrefix, EMRCompCatID, EMRCompTypeID, EMRCompStatusID, EMRCompModeID) 
    VALUES ( 'MED001.5', '&nbsp;', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'NONE', 'E001', 'E001', 'E001') 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM EMRComponent WHERE ( EMRComponentID='MED001.6') 
  ) 
  BEGIN 
    INSERT INTO EMRComponent ( EMRComponentID, EMRComponentName, CompPos, FlagPos1, Column1, ResultPos1, Column2, RefPos1, Column3, FlagPos2, Column4, ResultPos2, Column5, RefPos2, Column6, Description, KeyPrefix, EMRCompCatID, EMRCompTypeID, EMRCompStatusID, EMRCompModeID) 
    VALUES ( 'MED001.6', '&nbsp;', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'NONE', 'E001', 'E001', 'E001') 
  END 
END 

