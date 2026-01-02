--Physical function (9)
select SUM((50 * varpos) - 50) as score, count(varpos) as total from EMRResults e, emrvar3b e2 
where cast(e.Column3 as varchar) = e2.EMRVar3BID 
and EMRRequestID =  'E100085851' 
AND e.EmrComponentID IN ('RES017003', 'RES017004', 'RES017005', 'RES017006', 'RES017007')
UNION ALL
select sum((50 * varpos) - 50) as score, count(varpos) as total from EMRResults e, emrvar3b e2 
where cast(e.Column6 as varchar) = e2.EMRVar3BID and EMRRequestID =  'E100085851'
AND e.EmrComponentID IN ('RES017003', 'RES017004', 'RES017005', 'RES017006', 'RES017007') 

-- Role limitations due to physical health (4)
-- Role limitations due to emotional problems. (7)
select SUM((25 * varpos) - 25) as score, count(varpos) as total from EMRResults e, emrvar3b e2 
where cast(e.Column3 as varchar) = e2.EMRVar3BID 
and EMRRequestID =  'E100085851' 
AND e.EmrComponentID IN ('RES017009', 'RES017010', 'RES017013', 'RES017014')
UNION ALL
select SUM((25 * varpos) - 25) as score, count(varpos) as total from EMRResults e, emrvar3b e2 
where cast(e.Column6 as varchar) = e2.EMRVar3BID and EMRRequestID =  'E100085851'
AND e.EmrComponentID IN ('RES017009', 'RES017010', 'RES017013', 'RES017014')

SELECT DISTINCT EMRRequestID, PatientID,visitationID FROM EMRRequestItems 
WHERE EMRDataID = 'RES017' 

 select sum(score) as userScore, sum(total) as totalScore from ( 
    select sum(ABS(varpos-5)+1) as score, count(varpos) as total from EMRResults e, emrvar3b e2 
    where cast(e.Column3 as varchar) = e2.EMRVar3BID and EMRRequestID =  'E100116704' 
    union all 
    select sum(ABS(varpos-5)+1) as score, count(varpos) as total from EMRResults e, emrvar3b e2 
    where cast(e.Column6 as varchar) = e2.EMRVar3BID and EMRRequestID =  'E100116704' 
    ) as results 


SELECT DISTINCT EMRRequestID, PatientID,visitationID FROM EMRRequestItems 
WHERE EMRDataID = 'RES017' 

 select sum(score) as userScore, sum(total) as totalScore from ( 
    select sum(ABS(varpos-5)+1) as score, count(varpos) as total from EMRResults e, emrvar3b e2 
    where cast(e.Column3 as varchar) = e2.EMRVar3BID and EMRRequestID =  'E100116704' 
    union all 
    select sum(ABS(varpos-5)+1) as score, count(varpos) as total from EMRResults e, emrvar3b e2 
    where cast(e.Column6 as varchar) = e2.EMRVar3BID and EMRRequestID =  'E100116704' 
    ) as results 
 
--  07/11/2024
-- physical function  
select sum(score) as userScore, sum(total) as totalScore from ( 
select SUM((50 * varpos) - 50) as score, count(varpos) as total from EMRResults e, emrvar3b e2 
where cast(e.Column3 as varchar) = e2.EMRVar3BID 
and EMRRequestID =  'E100085851' 
AND e.EmrComponentID IN ('RES017003', 'RES017004', 'RES017005', 'RES017006', 'RES017007')
UNION ALL

select sum((50 * varpos) - 50) as score, count(varpos) as total from EMRResults e, emrvar3b e2 
where cast(e.Column6 as varchar) = e2.EMRVar3BID and EMRRequestID =  'E100085851'
AND e.EmrComponentID IN ('RES017003', 'RES017004', 'RES017005', 'RES017006', 'RES017007') 
UNION ALL

--role limitation due to physical limitation & emotional problems
select SUM((25 * varpos) - 25) as score, count(varpos) as total from EMRResults e, emrvar3b e2 
where cast(e.Column3 as varchar) = e2.EMRVar3BID 
and EMRRequestID =  'E100085851' 
AND e.EmrComponentID IN ('RES017009', 'RES017010', 'RES017013', 'RES017014')
UNION ALL

select SUM((25 * varpos) - 25) as score, count(varpos) as total from EMRResults e, emrvar3b e2 
where cast(e.Column6 as varchar) = e2.EMRVar3BID and EMRRequestID =  'E100085851'
AND e.EmrComponentID IN ('RES017009', 'RES017010', 'RES017013', 'RES017014')
UNION ALL

--Energy/Fatigue / (E1 & E2)
select SUM(125 - (25 * varpos)) as score, count(varpos) as total from EMRResults e, emrvar3b e2 
where cast(e.Column3 as varchar) = e2.EMRVar3BID 
and EMRRequestID =  'E100085851' 
AND e.EmrComponentID = 'RES017016'
UNION ALL

select SUM(125 - (25 * varpos)) as score, count(varpos) as total from EMRResults e, emrvar3b e2 
where cast(e.Column6 as varchar) = e2.EMRVar3BID 
and EMRRequestID =  'E100085851' 
AND e.EmrComponentID = 'RES017016'
UNION ALL

--Energy/Fatigue / (E3 & E4)
select SUM((25 * varpos) - 25) as score, count(varpos) as total from EMRResults e, emrvar3b e2 
where cast(e.Column3 as varchar) = e2.EMRVar3BID 
and EMRRequestID =  'E100085851' 
AND e.EmrComponentID = 'RES017017'
UNION ALL

select SUM((25 * varpos) - 25) as score, count(varpos) as total from EMRResults e, emrvar3b e2 
where cast(e.Column6 as varchar) = e2.EMRVar3BID 
and EMRRequestID =  'E100085851' 
AND e.EmrComponentID = 'RES017017'
UNION ALL

--Emotional WellBeing E1 and E2
select SUM((25 * varpos) - 25) as score, count(varpos) as total from EMRResults e, emrvar3b e2 
where cast(e.Column3 as varchar) = e2.EMRVar3BID 
and EMRRequestID =  'E100085851' 
AND e.EmrComponentID = 'RES017020'
UNION ALL

select SUM((25 * varpos) - 25) as score, count(varpos) as total from EMRResults e, emrvar3b e2 
where cast(e.Column6 as varchar) = e2.EMRVar3BID 
and EMRRequestID =  'E100085851' 
AND e.EmrComponentID = 'RES017020'
UNION ALL

--Emotional WellBeing (E3)
select SUM(125 - (25 * varpos)) as score, count(varpos) as total from EMRResults e, emrvar3b e2 
where cast(e.Column3 as varchar) = e2.EMRVar3BID 
and EMRRequestID =  'E100085851' 
AND e.EmrComponentID = 'RES017021'
UNION ALL

--Emotional WellBeing E4
select SUM((25 * varpos) - 25) as score, count(varpos) as total from EMRResults e, emrvar3b e2 
where cast(e.Column6 as varchar) = e2.EMRVar3BID 
and EMRRequestID =  'E100085851' 
AND e.EmrComponentID = 'RES017021'
UNION ALL

--Emotional WellBeing E5
select SUM(125 - (25 * varpos)) as score, count(varpos) as total from EMRResults e, emrvar3b e2 
where cast(e.Column3 as varchar) = e2.EMRVar3BID 
and EMRRequestID =  'E100085851' 
AND e.EmrComponentID = 'RES017022'
UNION ALL

--Social Functioning
select SUM((25 * varpos) - 25) as score, count(varpos) as total from EMRResults e, emrvar3b e2 
where cast(e.Column3 as varchar) = e2.EMRVar3BID 
and EMRRequestID =  'E100085851' 
AND e.EmrComponentID = 'RES017024'
UNION ALL

select SUM((25 * varpos) - 25) as score, count(varpos) as total from EMRResults e, emrvar3b e2 
where cast(e.Column6 as varchar) = e2.EMRVar3BID 
and EMRRequestID =  'E100085851' 
AND e.EmrComponentID = 'RES017024'
UNION ALL

--Pain
select SUM((25 * varpos) - 25) as score, count(varpos) as total from EMRResults e, emrvar3b e2 
where cast(e.Column3 as varchar) = e2.EMRVar3BID 
and EMRRequestID =  'E100085851' 
AND e.EmrComponentID = 'RES017027'
UNION ALL

select SUM((25 * varpos) - 25) as score, count(varpos) as total from EMRResults e, emrvar3b e2 
where cast(e.Column6 as varchar) = e2.EMRVar3BID 
and EMRRequestID =  'E100085851' 
AND e.EmrComponentID = 'RES017027'
UNION ALL

-- General Health (G1)
select SUM((25 * varpos) - 25) as score, count(varpos) as total from EMRResults e, emrvar3b e2 
where cast(e.Column3 as varchar) = e2.EMRVar3BID 
and EMRRequestID =  'E100085851' 
AND e.EmrComponentID = 'RES017030'
UNION ALL

-- General Health (G2)
select SUM((25 * varpos) - 25) as score, count(varpos) as total from EMRResults e, emrvar3b e2 
where cast(e.Column6 as varchar) = e2.EMRVar3BID 
and EMRRequestID =  'E100085851' 
AND e.EmrComponentID = 'RES017030'
UNION ALL

-- General Health (G3)
select SUM(125 - (25 * varpos)) as score, count(varpos) as total from EMRResults e, emrvar3b e2 
where cast(e.Column3 as varchar) = e2.EMRVar3BID 
and EMRRequestID =  'E100085851' 
AND e.EmrComponentID = 'RES017031'
UNOIN ALL

--General Health (G4)
select SUM((25 * varpos) - 25) as score, count(varpos) as total from EMRResults e, emrvar3b e2 
where cast(e.Column6 as varchar) = e2.EMRVar3BID 
and EMRRequestID =  'E100085851' 
AND e.EmrComponentID = 'RES017031'
UNION ALL

--General Health (G5)
select SUM(125 - (25 * varpos)) as score, count(varpos) as total from EMRResults e, emrvar3b e2 
where cast(e.Column3 as varchar) = e2.EMRVar3BID 
and EMRRequestID =  'E100085851' 
AND e.EmrComponentID = 'RES017032'
) as results