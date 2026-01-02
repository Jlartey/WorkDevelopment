sql = " WITH DoctorXray AS ("
sql = sql & " SELECT SystemUserID, COUNT(LabRequestID) Tests"
sql = sql & " From Investigation2"
sql = sql & " Join LabTest"
sql = sql & " ON Investigation2.LabtestID = LabTest.LabTestID"
sql = sql & " WHERE Investigation2.TestCategoryID = 'B19'"
sql = sql & " AND LabtestName LIKE '%ray%'
sql = sql & " AND LabTest.TestStatusID = 'TST001'
sql = sql & " AND RequestDate BETWEEN '2025-01-01' AND '2025-12-17'
sql = sql & " GROUP BY SystemUserID
sql = sql & " ),
sql = sql & " DoctorScans AS (
sql = sql & " SELECT SystemUserID, COUNT(LabRequestID) Tests
sql = sql & " From Investigation2
sql = sql & " Join LabTest
sql = sql & "     ON Investigation2.LabtestID = LabTest.LabTestID
sql = sql & " WHERE Investigation2.TestCategoryID = 'B19'
sql = sql & " AND LabtestName NOT LIKE '%ray%'
sql = sql & " AND LabTest.TestStatusID = 'TST001'
sql = sql & " AND RequestDate BETWEEN '2025-01-01' AND '2025-12-17'
sql = sql & " GROUP BY SystemUserID
sql = sql & " )
sql = sql & " SELECT stf.Staffname, dx.Tests Xrays, dc.Tests Scans
sql = sql & " FROM DoctorXray dx
sql = sql & " JOIN DoctorScans dc
sql = sql & " ON dx.SystemuserID = dc.SystemUserID
sql = sql & " JOIN SystemUser sys
sql = sql & " ON dc.SystemUserID = sys.SystemUserID
sql = sql & " JOIN Staff stf
sql = sql & " ON sys.StaffID = stf.StaffID
sql = sql & " ORDER BY Staffname