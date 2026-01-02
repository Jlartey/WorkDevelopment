SELECT DISTINCT DrugStoreID FROM DrugStockLevel

SELECT 
    d.DrugName,
    COALESCE(SUM(CASE WHEN ds.DrugStoreID = 'S21A' THEN ds.AvailableQty ELSE 0 END), 0) AS Store1,
    COALESCE(SUM(CASE WHEN ds.DrugStoreID = 'S22' THEN ds.AvailableQty ELSE 0 END), 0) AS Store2,
    COALESCE(SUM(CASE WHEN ds.DrugStoreID = 'S21C' THEN ds.AvailableQty ELSE 0 END), 0) AS Store3,
    COALESCE(SUM(CASE WHEN ds.DrugStoreID = 'S23' THEN ds.AvailableQty ELSE 0 END), 0) AS Store4,  
    SUM(ds.AvailableQty) AS [Total Drug Quantity]
FROM Drug d
INNER JOIN DrugStockLevel ds ON d.DrugID = ds.DrugID
GROUP BY d.DrugName
ORDER BY [Total Drug Quantity] DESC;