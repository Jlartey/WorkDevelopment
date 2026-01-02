DELETE FROM "IMAHMS01"."dbo"."EMRData" WHERE  "EMRDataID"=N'E000048';
SELECT SUM("rows") FROM "sys"."partitions" WHERE "index_id" IN (0, 1) AND "object_id" = object_id('IMAHMS01.dbo.EMRData');