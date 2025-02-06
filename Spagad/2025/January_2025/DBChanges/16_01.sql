Patient is deemed high fall-risk per protocol 
(e.g seizure precautions <br>Low Fall Rick - implement Low Fall Risk interventions per protocol
Complete paralysis or completely immobilized

SELECT SUM("rows") FROM "sys"."partitions" WHERE "index_id" IN (0, 1) AND "object_id" = object_id('FOHMS02.dbo.EMRVar3B');
SELECT "EMRVar3BID", "EMRVar3BName", "EMRVar3AID", "VarPos", "Description", "KeyPrefix" FROM "FOHMS02"."dbo"."EMRVar3B" WHERE  "EMRVar3BID"=N'E045';
UPDATE "FOHMS02"."dbo"."EMRVar3B" SET "EMRVar3BName"=N'Patient is deemed high fall-risk per protocol ' WHERE  "EMRVar3BID"=N'E045';
SELECT "EMRVar3BID", "EMRVar3BName", "EMRVar3AID", "VarPos", "Description", "KeyPrefix" FROM "FOHMS02"."dbo"."EMRVar3B" WHERE  "EMRVar3BID"=N'E045';
SELECT SUM("rows") FROM "sys"."partitions" WHERE "index_id" IN (0, 1) AND "object_id" = object_id('FOHMS02.dbo.EMRVar3B');

SELECT SUM("rows") FROM "sys"."partitions" WHERE "index_id" IN (0, 1) AND "object_id" = object_id('FOHMS02.dbo.EMRVar3B');
DELETE FROM "FOHMS02"."dbo"."EMRVar3B" WHERE  "EMRVar3BID"=N'E046';
SELECT SUM("rows") FROM "sys"."partitions" WHERE "index_id" IN (0, 1) AND "object_id" = object_id('FOHMS02.dbo.EMRVar3B');