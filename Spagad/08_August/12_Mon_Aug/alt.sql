SELECT DISTINCT drug.drugname, 
                drug.RetailUnitCost, 
                drugstocklevel.drugid, 
                CASE 
                    WHEN (SELECT TOP 1 IncomingDrugItems.ExpiryDate 
                          FROM IncomingDrugItems 
                          WHERE IncomingDrugItems.drugid = drug.drugid) IS NULL 
                    THEN 'N/A' 
                    ELSE CONVERT(VARCHAR(20), 
                                (SELECT TOP 1 IncomingDrugItems.ExpiryDate 
                                 FROM IncomingDrugItems 
                                 WHERE IncomingDrugItems.drugid = drug.drugid), 103) 
                END AS ExpiryDate
FROM DrugStockLevel, drug
WHERE drugstocklevel.drugid = drug.drugid
AND drug.DrugStatusID <> 'IST002'
AND drugstocklevel.drugstockstatusid = 'D001'
AND drugstocklevel.drugstoreid = 's21a'
ORDER BY drug.drugName;
