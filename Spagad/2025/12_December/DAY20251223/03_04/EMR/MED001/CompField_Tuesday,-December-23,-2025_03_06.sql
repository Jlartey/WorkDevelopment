BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.7'  and CompFieldID='Column1') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.7', 'Column1', '&nbsp;', 'TDProp**1**Left**Top%%Default**<b>Interventions and Outcome</b>', '', '10', 'DTY002', '', '', 'FALSE', '01', null, null, '01', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.7'  and CompFieldID='Column2') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.7', 'Column2', '&nbsp;', 'TDProp**5**Left**Top%%UserTextArea**3**60', '', '10', 'DTY002', '', '', 'FALSE', '02', null, null, '02', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.7'  and CompFieldID='Column3') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.7', 'Column3', '&nbsp;', 'TREEVIEW', '', '10', 'DTY002', '', '', 'FALSE', '03', null, null, '03', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.7'  and CompFieldID='Column4') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.7', 'Column4', '&nbsp;', 'TREEVIEW', '', '10', 'DTY002', '', '', 'FALSE', '04', null, null, '04', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.7'  and CompFieldID='Column5') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.7', 'Column5', '&nbsp;', 'TREEVIEW', '', '10', 'DTY002', '', '', 'FALSE', '05', null, null, '05', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.7'  and CompFieldID='Column6') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.7', 'Column6', '&nbsp;', 'TREEVIEW', '', '10', 'DTY002', '', '', 'FALSE', '06', null, null, '06', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.1'  and CompFieldID='Column1') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.1', 'Column1', '&nbsp;', 'TDProp**1**Left**Top%%Default**<b>Date</b>', '', '10', 'DTY002', '', '', 'FALSE', '01', null, null, '01', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.1'  and CompFieldID='Column2') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.1', 'Column2', '&nbsp;', 'TDProp**5**Left**Top%%UserDate', '', '10', 'DTY002', '', '', 'FALSE', '02', null, null, '02', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.1'  and CompFieldID='Column3') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.1', 'Column3', '&nbsp;', 'TREEVIEW', '', '10', 'DTY002', '', '', 'FALSE', '03', null, null, '03', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.1'  and CompFieldID='Column4') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.1', 'Column4', '&nbsp;', 'TREEVIEW', '', '10', 'DTY002', '', '', 'FALSE', '04', null, null, '04', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.1'  and CompFieldID='Column5') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.1', 'Column5', '&nbsp;', 'TREEVIEW', '', '10', 'DTY002', '', '', 'FALSE', '05', null, null, '05', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.1'  and CompFieldID='Column6') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.1', 'Column6', '&nbsp;', 'TREEVIEW', '', '10', 'DTY002', '', '', 'FALSE', '06', null, null, '06', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.3'  and CompFieldID='Column1') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.3', 'Column1', '&nbsp;', 'TDProp**1**Left**Top%%Default**<b>Appropriateness of Drug</b>', '', '10', 'DTY002', '', '', 'FALSE', '01', null, null, '01', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.3'  and CompFieldID='Column2') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.3', 'Column2', '&nbsp;', 'TDProp**2**left**TOP****%%UserCombo**EMRVAR3B****EMRVAR3B.EMRVAR3AID = ''E059''', '', '10', 'DTY002', '', '', 'FALSE', '02', null, null, '02', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.3'  and CompFieldID='Column3') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.3', 'Column3', '&nbsp;', 'TREEVIEW', '', '10', 'DTY002', '', '', 'FALSE', '03', null, null, '03', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.3'  and CompFieldID='Column4') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.3', 'Column4', '&nbsp;', 'TDProp**1**Left**Top%%Default**<b>Appropriateness of Dose</b>', '', '10', 'DTY002', '', '', 'FALSE', '04', null, null, '04', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.3'  and CompFieldID='Column5') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.3', 'Column5', '&nbsp;', 'TDProp**2**left**TOP****%%UserCombo**EMRVAR3B****EMRVAR3B.EMRVAR3AID = ''E059''', '', '10', 'DTY002', '', '', 'FALSE', '05', null, null, '05', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.3'  and CompFieldID='Column6') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.3', 'Column6', '&nbsp;', 'TREEVIEW', '', '10', 'DTY002', '', '', 'FALSE', '06', null, null, '06', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.4'  and CompFieldID='Column1') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.4', 'Column1', '&nbsp;', 'TDProp**1**Left**Top%%Default**<b>Appropriateness of Frequency</b>', '', '10', 'DTY002', '', '', 'FALSE', '01', null, null, '01', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.4'  and CompFieldID='Column2') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.4', 'Column2', '&nbsp;', 'TDProp**2**left**TOP****%%UserCombo**EMRVAR3B****EMRVAR3B.EMRVAR3AID = ''E059''', '', '10', 'DTY002', '', '', 'FALSE', '02', null, null, '02', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.4'  and CompFieldID='Column3') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.4', 'Column3', '&nbsp;', 'TREEVIEW', '', '10', 'DTY002', '', '', 'FALSE', '03', null, null, '03', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.4'  and CompFieldID='Column4') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.4', 'Column4', '&nbsp;', 'TDProp**1**Left**Top%%Default**<b>Appropriateness of Route</b>', '', '10', 'DTY002', '', '', 'FALSE', '04', null, null, '04', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.4'  and CompFieldID='Column5') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.4', 'Column5', '&nbsp;', 'TDProp**2**Left**TOP****%%UserCombo**EMRVAR3B****EMRVAR3B.EMRVAR3AID = ''E059''', '', '10', 'DTY002', '', '', 'FALSE', '05', null, null, '05', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.4'  and CompFieldID='Column6') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.4', 'Column6', '&nbsp;', 'TREEVIEW', '', '10', 'DTY002', '', '', 'FALSE', '06', null, null, '06', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.5'  and CompFieldID='Column1') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.5', 'Column1', '&nbsp;', 'TDProp**1**Left**Top%%Default**<b>Therapeutic Duplication</b>', '', '10', 'DTY002', '', '', 'FALSE', '01', null, null, '01', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.5'  and CompFieldID='Column2') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.5', 'Column2', '&nbsp;', 'TDProp**5**Left**TOP****%%UserCombo**EMRVAR3B****EMRVAR3B.EMRVAR3AID = ''E060''', '', '10', 'DTY002', '', '', 'FALSE', '02', null, null, '02', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.5'  and CompFieldID='Column3') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.5', 'Column3', '&nbsp;', 'TREEVIEW', '', '10', 'DTY002', '', '', 'FALSE', '03', null, null, '03', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.5'  and CompFieldID='Column4') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.5', 'Column4', '&nbsp;', 'TREEVIEW', '', '10', 'DTY002', '', '', 'FALSE', '04', null, null, '04', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.5'  and CompFieldID='Column5') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.5', 'Column5', '&nbsp;', 'TREEVIEW', '', '10', 'DTY002', '', '', 'FALSE', '05', null, null, '05', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.5'  and CompFieldID='Column6') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.5', 'Column6', '&nbsp;', 'TREEVIEW', '', '10', 'DTY002', '', '', 'FALSE', '06', null, null, '06', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.6'  and CompFieldID='Column1') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.6', 'Column1', '&nbsp;', 'TDProp**1**Left**Top%%Default**<b>Comments</b>', '', '10', 'DTY002', '', '', 'FALSE', '01', null, null, '01', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.6'  and CompFieldID='Column2') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.6', 'Column2', '&nbsp;', 'TDProp**5**Left**Top%%UserTextArea**3**60', '', '10', 'DTY002', '', '', 'FALSE', '02', null, null, '02', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.6'  and CompFieldID='Column3') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.6', 'Column3', '&nbsp;', 'TREEVIEW', '', '10', 'DTY002', '', '', 'FALSE', '03', null, null, '03', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.6'  and CompFieldID='Column4') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.6', 'Column4', '&nbsp;', 'TREEVIEW', '', '10', 'DTY002', '', '', 'FALSE', '04', null, null, '04', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.6'  and CompFieldID='Column5') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.6', 'Column5', '&nbsp;', 'TREEVIEW', '', '10', 'DTY002', '', '', 'FALSE', '05', null, null, '05', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

BEGIN 
  IF NOT EXISTS (
    SELECT * FROM CompField WHERE ( RecordKey='MED001.6'  and CompFieldID='Column6') 
  ) 
  BEGIN 
    INSERT INTO CompField ( CompTableKeyID, RecordKey, CompFieldID, CompFieldName, SubTableFieldSource, DisplayName, FieldLen, DataTypeID, DefaultValue, DataFormat, Required, DatagridPos, RecordFilterField, ForeignKeyField, TreeviewDatagridPos, DisabledField, VisibleField, PrimaryKey, TransientField, DetailField, DetailField2, SummaryField, GraphField, HelpDescription, KeyPrefix) 
    VALUES ( 'EMRComponentID', 'MED001.6', 'Column6', '&nbsp;', 'TREEVIEW', '', '10', 'DTY002', '', '', 'FALSE', '06', null, null, '06', 'FALSE', 'YES||YES', 'NO', null, null, null, null, null, null, null) 
  END 
END 

