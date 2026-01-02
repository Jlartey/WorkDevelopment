WITH DoctorConsults AS (
 SELECT DISTINCT SystemUserID, VisitationID FROM 
    EMRRequestItems 
    WHERE EMRDataID IN ('TH060', 'IM051') 
    AND VisitationID = 'V1190807038' 
    ) 
    SELECT Staff.StaffName 
    FROM DoctorConsults dc 
    Join SystemUser 
    ON SystemUser.SystemUserID = dc.SystemUserID 
    Join Staff 
    ON Staff.StaffID = SystemUser.StaffID 