SELECT * FROM "FOHMS02"."dbo"."ManufacturerType" ORDER BY 1 OFFSET 0 ROWS FETCH NEXT 1000 ROWS ONLY;
DELETE FROM "FOHMS02"."dbo"."ManufacturerType" WHERE  "ManufacturerTypeID"=N'M006';
DELETE FROM "FOHMS02"."dbo"."ManufacturerType" WHERE  "ManufacturerTypeID"=N'M007';

SELECT  "ManufacturerID",  "ManufacturerName",  "ManufacturerTypeID",  "CountryID", LEFT(CAST("Address" AS NVARCHAR(256)), 256),  "City",  "Location",  "OfficePhone",  "OfficeFax",  "KeyPrefix" FROM "FOHMS02"."dbo"."Manufacturer" ORDER BY 1 OFFSET 0 ROWS FETCH NEXT 1000 ROWS ONLY;
SELECT  "ManufacturerID",  "ManufacturerName",  "ManufacturerTypeID",  "CountryID", LEFT(CAST("Address" AS NVARCHAR(256)), 256),  "City",  "Location",  "OfficePhone",  "OfficeFax",  "KeyPrefix" FROM "FOHMS02"."dbo"."Manufacturer" WHERE MANUFACTURERNAME = 'STAR X-RAY'
 ORDER BY 1 OFFSET 0 ROWS FETCH NEXT 1000 ROWS ONLY;
 --DELETED STAR X-RAY
DELETE FROM "FOHMS02"."dbo"."Manufacturer" WHERE  "ManufacturerID"=N'M026'; 
DELETE FROM "FOHMS02"."dbo"."Manufacturer" WHERE  "ManufacturerID"=N'M027';
DELETE FROM "FOHMS02"."dbo"."Manufacturer" WHERE  "ManufacturerID"=N'M028';
DELETE FROM "FOHMS02"."dbo"."Manufacturer" WHERE  "ManufacturerID"=N'M029';