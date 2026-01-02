RES006^^RES006034Column5

DELETE FROM "FOHMS02"."dbo"."EMRVar3B" WHERE  "EMRVar3BID"=N'E235';
SELECT SUM("rows") FROM "sys"."partitions" WHERE "index_id" IN (0, 1) AND "object_id" = object_id('FOHMS02.dbo.EMRVar3B');

EMRVAR3AID = 'RES024004 - to be used to correct