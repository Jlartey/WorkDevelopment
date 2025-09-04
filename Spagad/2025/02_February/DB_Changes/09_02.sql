DELETE FROM "FOHMS02"."dbo"."AssetPurchase" WHERE  "AssetPurchaseID"=N'A162';
SELECT SUM("rows") FROM "sys"."partitions" WHERE "index_id" IN (0, 1) AND "object_id" = object_id('FOHMS02.dbo.AssetPurchase');

FXTBR1800347TJ