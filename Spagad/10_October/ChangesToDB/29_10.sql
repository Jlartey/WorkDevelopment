/* Affected rows: 0  Found rows: 6  Warnings: 0  Duration for 1 query: 0.015 sec. */
UPDATE "FOHMS02"."dbo"."EMRVar3B" SET "EMRVar3BName"=N'My social life is normal and does not increase  my pain.' WHERE  "EMRVar3BID"=N'RES013H01';
SELECT "EMRVar3BID", "EMRVar3BName", "EMRVar3AID", "VarPos", "Description", "KeyPrefix" FROM "FOHMS02"."dbo"."EMRVar3B" WHERE  "EMRVar3BID"=N'RES013H01';

/* Affected rows: 0  Found rows: 6  Warnings: 0  Duration for 1 query: 0.015 sec. */
UPDATE "FOHMS02"."dbo"."EMRVar3B" SET "EMRVar3BName"=N'My social life is normal and does not increase  my pain.' WHERE  "EMRVar3BID"=N'RES013H01';
SELECT "EMRVar3BID", "EMRVar3BName", "EMRVar3AID", "VarPos", "Description", "KeyPrefix" FROM "FOHMS02"."dbo"."EMRVar3B" WHERE  "EMRVar3BID"=N'RES013H01';
UPDATE "FOHMS02"."dbo"."EMRVar3B" SET "EMRVar3BName"=N'I have hardly any social life because of my  pain.' WHERE  "EMRVar3BID"=N'RES013H06';
SELECT "EMRVar3BID", "EMRVar3BName", "EMRVar3AID", "VarPos", "Description", "KeyPrefix" FROM "FOHMS02"."dbo"."EMRVar3B" WHERE  "EMRVar3BID"=N'RES013H06';
SELECT TOP 20 * FROM EMRVar3B
WHERE EMRVAR3AID = 'RES013I';
/*  */
/* Affected rows: 0  Found rows: 6  Warnings: 0  Duration for 1 query: 0.000 sec. */
UPDATE "FOHMS02"."dbo"."EMRVar3B" SET "EMRVar3BName"=N'My pain restricts my travel to short necessary journeys under 1/2 hour.' WHERE  "EMRVar3BID"=N'RES013I05';
SELECT "EMRVar3BID", "EMRVar3BName", "EMRVar3AID", "VarPos", "Description", "KeyPrefix" FROM "FOHMS02"."dbo"."EMRVar3B" WHERE  "EMRVar3BID"=N'RES013I05';
UPDATE "FOHMS02"."dbo"."EMRVar3B" SET "EMRVar3BName"=N'My pain prevents all travel except for visits to the physician/therapist or hospital.' WHERE  "EMRVar3BID"=N'RES013I06';
SELECT "EMRVar3BID", "EMRVar3BName", "EMRVar3AID", "VarPos", "Description", "KeyPrefix" FROM "FOHMS02"."dbo"."EMRVar3B" WHERE  "EMRVar3BID"=N'RES013I06';

UPDATE "FOHMS02"."dbo"."EMRVar3B" SET "EMRVar3BName"=N'My normal homemaking/job activities do not  cause pain.' WHERE  "EMRVar3BID"=N'RES013J01';
SELECT "EMRVar3BID", "EMRVar3BName", "EMRVar3AID", "VarPos", "Description", "KeyPrefix" FROM "FOHMS02"."dbo"."EMRVar3B" WHERE  "EMRVar3BID"=N'RES013J01';