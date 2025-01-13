E100152136
E100069142
E100022551

Pain Assessment tool
1. E100022475
2. E100022461
3. E100015689

Pairs
E100022475/V1221231010
E100022461/

select sum(score) as userScore, sum(total) as totalScore from ( 
select sum(varpos) as score, count(varpos) as total from EMRResults e, emrvar3b e2 
where cast(e.Column4 as varchar) = e2.EMRVar3BID and EMRRequestID =  'E100062812' 
union all 
select sum(varpos) as score, count(varpos) as total from EMRResults e, emrvar3b e2 
where cast(e.Column1 as varchar) = e2.EMRVar3BID and EMRRequestID =  'E100062812' 
) as results 
