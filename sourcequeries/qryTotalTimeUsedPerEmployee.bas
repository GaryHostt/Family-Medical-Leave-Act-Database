SELECT DISTINCTROW qryTotalTimeUsedPerRequest.Employee, Sum(qryTotalTimeUsedPerRequest.SumOfTime_used) AS SumOfSumOfTime_used, Last(qryTotalTimeUsedPerRequest.Start) AS LastOfStart, Sum(qryTotalTimeUsedPerRequest.HoursPerDay) AS SumOfHoursPerDay
FROM tblRequestTimeLog, qryTotalTimeUsedPerRequest
GROUP BY qryTotalTimeUsedPerRequest.Employee;

