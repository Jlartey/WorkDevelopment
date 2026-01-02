BEGIN 
  IF NOT EXISTS (
    SELECT * FROM EMRTemplate WHERE ( EMRComponentID='MED001.7'  and EMRDataID='MED001'  and EMRCompTabID='MED001') 
  ) 
  BEGIN 
    INSERT INTO EMRTemplate ( EMRComponentID, EMRDataID, CompPos, Column1, Column2, Column3, Column4, Column5, Column6, Description, TabDisplayName, TabInfo, TabPos, EMRCompTabID) 
    VALUES ( 'MED001.7', 'MED001', '07', '', '', '', '', '', '', '', '', '', '1', 'MED001') 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM EMRTemplate WHERE ( EMRComponentID='MED001.1'  and EMRDataID='MED001'  and EMRCompTabID='MED001') 
  ) 
  BEGIN 
    INSERT INTO EMRTemplate ( EMRComponentID, EMRDataID, CompPos, Column1, Column2, Column3, Column4, Column5, Column6, Description, TabDisplayName, TabInfo, TabPos, EMRCompTabID) 
    VALUES ( 'MED001.1', 'MED001', '01', '', '', '', '', '', '', '', '', '', '1', 'MED001') 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM EMRTemplate WHERE ( EMRComponentID='MED001.3'  and EMRDataID='MED001'  and EMRCompTabID='MED001') 
  ) 
  BEGIN 
    INSERT INTO EMRTemplate ( EMRComponentID, EMRDataID, CompPos, Column1, Column2, Column3, Column4, Column5, Column6, Description, TabDisplayName, TabInfo, TabPos, EMRCompTabID) 
    VALUES ( 'MED001.3', 'MED001', '03', '', '', '', '', '', '', '', '', '', '1', 'MED001') 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM EMRTemplate WHERE ( EMRComponentID='MED001.4'  and EMRDataID='MED001'  and EMRCompTabID='MED001') 
  ) 
  BEGIN 
    INSERT INTO EMRTemplate ( EMRComponentID, EMRDataID, CompPos, Column1, Column2, Column3, Column4, Column5, Column6, Description, TabDisplayName, TabInfo, TabPos, EMRCompTabID) 
    VALUES ( 'MED001.4', 'MED001', '04', '', '', '', '', '', '', '', '', '', '1', 'MED001') 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM EMRTemplate WHERE ( EMRComponentID='MED001.5'  and EMRDataID='MED001'  and EMRCompTabID='MED001') 
  ) 
  BEGIN 
    INSERT INTO EMRTemplate ( EMRComponentID, EMRDataID, CompPos, Column1, Column2, Column3, Column4, Column5, Column6, Description, TabDisplayName, TabInfo, TabPos, EMRCompTabID) 
    VALUES ( 'MED001.5', 'MED001', '05', '', '', '', '', '', '', '', '', '', '1', 'MED001') 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM EMRTemplate WHERE ( EMRComponentID='MED001.6'  and EMRDataID='MED001'  and EMRCompTabID='MED001') 
  ) 
  BEGIN 
    INSERT INTO EMRTemplate ( EMRComponentID, EMRDataID, CompPos, Column1, Column2, Column3, Column4, Column5, Column6, Description, TabDisplayName, TabInfo, TabPos, EMRCompTabID) 
    VALUES ( 'MED001.6', 'MED001', '06', '', '', '', '', '', '', '', '', '', '1', 'MED001') 
  END 
END 

