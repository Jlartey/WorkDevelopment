UPDATE "FOHMS02"."dbo"."EMRVar3B" SET "VarPos"=3 WHERE  "EMRVar3BID"=N'RES023016004';
SELECT "EMRVar3BID", "EMRVar3BName", "EMRVar3AID", "VarPos", "Description", "KeyPrefix" FROM "FOHMS02"."dbo"."EMRVar3B" WHERE  "EMRVar3BID"=N'RES023016004';
SELECT SUM("rows") FROM "sys"."partitions" WHERE "index_id" IN (0, 1) AND "object_id" = object_id('FOHMS02.dbo.EMRVar3B');
UPDATE "FOHMS02"."dbo"."EMRVar3B" SET "VarPos"=4 WHERE  "EMRVar3BID"=N'RES023016005';
SELECT "EMRVar3BID", "EMRVar3BName", "EMRVar3AID", "VarPos", "Description", "KeyPrefix" FROM "FOHMS02"."dbo"."EMRVar3B" WHERE  "EMRVar3BID"=N'RES023016005';
SELECT SUM("rows") FROM "sys"."partitions" WHERE "index_id" IN (0, 1) AND "object_id" = object_id('FOHMS02.dbo.EMRVar3B');
UPDATE "FOHMS02"."dbo"."EMRVar3B" SET "VarPos"=5 WHERE  "EMRVar3BID"=N'RES023016006';
SELECT "EMRVar3BID", "EMRVar3BName", "EMRVar3AID", "VarPos", "Description", "KeyPrefix" FROM "FOHMS02"."dbo"."EMRVar3B" WHERE  "EMRVar3BID"=N'RES023016006';
SELECT SUM("rows") FROM "sys"."partitions" WHERE "index_id" IN (0, 1) AND "object_id" = object_id('FOHMS02.dbo.EMRVar3B');