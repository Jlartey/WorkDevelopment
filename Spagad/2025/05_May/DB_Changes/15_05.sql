SELECT tc.CONSTRAINT_NAME, cc.CHECK_CLAUSE FROM "INFORMATION_SCHEMA"."CHECK_CONSTRAINTS" AS cc, "INFORMATION_SCHEMA"."TABLE_CONSTRAINTS" AS tc WHERE tc.CONSTRAINT_SCHEMA='dbo' AND tc.TABLE_NAME='Bed' AND tc.CONSTRAINT_TYPE='CHECK' AND tc.CONSTRAINT_SCHEMA=cc.CONSTRAINT_SCHEMA AND tc.CONSTRAINT_NAME=cc.CONSTRAINT_NAME;
SELECT * FROM "YBMHMS01"."dbo"."Bed" ORDER BY 1 OFFSET 0 ROWS FETCH NEXT 1000 ROWS ONLY;
SELECT * FROM "YBMHMS01"."dbo"."Bed" WHERE BEDNAME = 'Waiting List'
 ORDER BY 1 OFFSET 0 ROWS FETCH NEXT 1000 ROWS ONLY;
UPDATE "YBMHMS01"."dbo"."Bed" SET "BedCharge"=0 WHERE  "BedID"=N'W002-000';
SELECT "BedID", "BedName", "BedNoID", "BedPos", "BedTypeID", "BedStatusID", "BedModeID", "BlockID", "WardID", "WardSectionID", "BedGroupID", "BedClassID", "BillGroupCatID", "BillGroupID", "BedCharge", "BedVal1", "BedVal2", "BedVal3", "BedVal4", "BedInfo1", "BedInfo2", "BedDate1", "BedDate2", "Description", "KeyPrefix", "BillInput1", "BillInput2", "BillInput3", "BillInout4" FROM "YBMHMS01"."dbo"."Bed" WHERE  "BedID"=N'W002-000';
UPDATE "YBMHMS01"."dbo"."Bed" SET "BedCharge"=0 WHERE  "BedID"=N'W004-000';
SELECT "BedID", "BedName", "BedNoID", "BedPos", "BedTypeID", "BedStatusID", "BedModeID", "BlockID", "WardID", "WardSectionID", "BedGroupID", "BedClassID", "BillGroupCatID", "BillGr