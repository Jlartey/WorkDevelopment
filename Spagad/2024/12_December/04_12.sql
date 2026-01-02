INSERT INTO "TEstComponent" ("TestComponentID", "TestComponentName", "CompPos", "FlagPos1", "Column1", "ResultPos1", "Column2", "RefPos1", "Column3", "FlagPos2", "Column4", "ResultPos2", "Column5", "RefPos2", "Column6", "Description", "KeyPrefix", "TestCompCatID", "TestCompTypeID", "TestCompStatusID", "TestCompModeID") VALUES ('L0836', '&nbsp', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'T001', 'T001', 'T001', 'T001');
INSERT INTO "TEstComponent" ("TestComponentID", "TestComponentName", "CompPos", "FlagPos1", "Column1", "ResultPos1", "Column2", "RefPos1", "Column3", "FlagPos2", "Column4", "ResultPos2", "Column5", "RefPos2", "Column6", "Description", "KeyPrefix", "TestCompCatID", "TestCompTypeID", "TestCompStatusID", "TestCompModeID") VALUES ('L0837', '&nbsp', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'T001', 'T001', 'T001', 'T001');

--TestVar3A
INSERT INTO "TestVar3A" ("TestVar3AID", "TestVar3AName", "VarPos", "Description", "KeyPrefix") VALUES ('T0013A', 'Result', 1, '', '');

INSERT INTO "TestVar3B" ("TestVar3BID", "TestVar3BName", "TestVar3AID", "VarPos", "Description", "KeyPrefix") VALUES ('T0143B', 'REACTIVE', 'T0013A', 1, '', '');
INSERT INTO "TestVar3B" ("TestVar3BID", "TestVar3BName", "TestVar3AID", "VarPos", "Description", "KeyPrefix") VALUES ('T0143B2', 'NON-REACTIVE', 'T0013A', 2, '', '');

