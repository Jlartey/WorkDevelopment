'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

Dim dateRange, args, sql, rptGen
Set rptGen = New PRTGLO_RptGen

dateRange = Split(Trim(Request.QueryString("PrintFilter0")), "||")

If UBound(dateRange) > 0 Then
    sql = GetMorgueOccupancySQL(dateRange)
    args = "title=Morgue Occupancy from " & dateRange(0) & " to " & dateRange(1)
    args = args & ";hiddenFields=scenario;"
    args = args & ";ShowColumnTotal=No"
    rptGen.AddReport sql, args
    
    rptGen.ShowReport
End If

Function GetMorgueOccupancySQL(dateRange)
   Dim sql

   sql = " SELECT MortuaryID AS [ID],  VisitationId AS [Visit ID], patientid AS [Patient ID], MortFridgeID AS [Fridge Name], MortuaryName AS [Occupant Name]"
   sql = sql & "   , [Gender] "
   sql = sql & "   , CAST(NextOfKinBirthDate AS VARCHAR) AS [Date of Birth] , CAST(MortuaryDate2 AS VARCHAR) AS [Date of Death] "
   sql = sql & "   , [Age], [Cause of Death] "
   sql = sql & "   , CAST(DepositDate AS VARCHAR) AS [Deposit Date], CAST(ReleaseDate AS VARCHAR) AS [Release Date] "
   sql = sql & "   , NoOfDays AS [Days spent in period] "
   sql = sql & "   , DATEDIFF(DAY, DepositDate, ISNULL(ReleaseDate, '" & dateRange(1) & "')) AS [Projected Days Spent (from deposit to release)] " 'from actual fridge days -> projected days spent
   sql = sql & "   , (CASE WHEN (CAST(ReleaseDate AS DATE) <= GETDATE()) OR (CAST('" & dateRange(1) & "' AS DATE) <= GETDATE()) THEN "
   sql = sql & "                DATEDIFF(DAY, DepositDate, '" & dateRange(1) & "') "
   sql = sql & "            ELSE '-' "
   sql = sql & "      END) AS [Actual Days Spent (from deposit to " & dateRange(1) & ")]"
   sql = sql & "   , [Scenario] "
   sql = sql & "   FROM "
   sql = sql & " ( "
   sql = sql & " SELECT MortuaryID, VisitationId, patientid, MortFridgeID, MortuaryName, DepositDate, ReleaseDate "
   sql = sql & "   , DATEDIFF(day, DepositDate, ReleaseDate) AS NoOfDays, 1 AS [Scenario] "
   sql = sql & "    , DATEDIFF(YEAR, NextOfKinBirthDate, MortuaryDate2) AS [Age] "
   sql = sql & "    , MortuaryInfo12 AS [Cause of Death] "
   sql = sql & "     , (SELECT GenderName FROM Gender WHERE Gender.GenderID=Mortuary.GenderID ) AS [Gender] "
   sql = sql & "    , NextOfKinBirthDate, MortuaryDate2 "
   sql = sql & "   FROM Mortuary "
   sql = sql & "   WHERE 1=1 "
   sql = sql & "       AND DepositDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "'  "
   sql = sql & "       AND ReleaseDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "'  " 'deposited and released within chosen period
   sql = sql & " UNION ALL"
   sql = sql & " SELECT MortuaryID, VisitationId, patientid, MortFridgeID, MortuaryName, DepositDate, ReleaseDate "
   sql = sql & "    , DATEDIFF(day, DepositDate, '" & dateRange(1) & "') AS NoOfDays, 2 AS [Scenario] "
   sql = sql & "    , DATEDIFF(YEAR, NextOfKinBirthDate, MortuaryDate2) AS [Age] "
   sql = sql & "    , MortuaryInfo12 AS [Cause of Death] "
   sql = sql & "    , (SELECT GenderName FROM Gender WHERE Gender.GenderID=Mortuary.GenderID ) AS [Gender] "
   sql = sql & "    , NextOfKinBirthDate, MortuaryDate2 "
   sql = sql & "   FROM Mortuary "
   sql = sql & "   WHERE 1=1 "
   sql = sql & "       AND DepositDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "' "
   sql = sql & "       AND (LEN(ReleaseDate)=0 OR ReleaseDate > '" & dateRange(1) & "'        )" 'deposited within, unreleased
   sql = sql & " UNION ALL"
   sql = sql & " SELECT MortuaryID, VisitationId, patientid, MortFridgeID, MortuaryName, DepositDate, ReleaseDate "
   sql = sql & "    , DATEDIFF(day, '" & dateRange(0) & "', ReleaseDate ) AS NoOfDays, 3 AS [Scenario] "
   sql = sql & "    , DATEDIFF(YEAR, NextOfKinBirthDate, MortuaryDate2) AS [Age] "
   sql = sql & "    , MortuaryInfo12 AS [Cause of Death] "
   sql = sql & "     , (SELECT GenderName FROM Gender WHERE Gender.GenderID=Mortuary.GenderID ) AS [Gender] "
   sql = sql & "    , NextOfKinBirthDate, MortuaryDate2 "
   sql = sql & "    FROM Mortuary "
   sql = sql & "   WHERE 1=1 "
   sql = sql & "       AND DepositDate < CAST('" & dateRange(0) & "' AS DATE) "
   sql = sql & "       AND ReleaseDate BETWEEN '" & dateRange(0) & "' AND '" & dateRange(1) & "'  " 'deposited before, released within
   sql = sql & " UNION ALL"
   sql = sql & " SELECT MortuaryID, VisitationId, patientid, MortFridgeID, MortuaryName, DepositDate, ReleaseDate "
   sql = sql & "    , DATEDIFF(day, '" & dateRange(0) & "', '" & dateRange(1) & "') AS NoOfDays, 4 AS [Scenario] "
   sql = sql & "    , DATEDIFF(YEAR, NextOfKinBirthDate, MortuaryDate2) AS [Age] "
   sql = sql & "    , MortuaryInfo12 AS [Cause of Death] "
   sql = sql & "     , (SELECT GenderName FROM Gender WHERE Gender.GenderID=Mortuary.GenderID ) AS [Gender] "
   sql = sql & "    , NextOfKinBirthDate, MortuaryDate2 "
   sql = sql & "   FROM Mortuary "
   sql = sql & "   WHERE 1=1 "
   sql = sql & "       AND DepositDate < CAST('" & dateRange(0) & "' AS DATE) "
   sql = sql & "       AND ( LEN(ReleaseDate)=0 OR ReleaseDate > '" & dateRange(1) & "' )       " 'deposited before, still there
   sql = sql & " )  "
   sql = sql & " AS MortDays "
   sql = sql & " ORDER BY MortuaryID ASC "

   GetMorgueOccupancySQL = sql
End Function

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
