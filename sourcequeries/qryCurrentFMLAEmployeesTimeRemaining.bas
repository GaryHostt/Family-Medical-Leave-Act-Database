SELECT DISTINCT tblEmployees.[First Name], tblEmployees.[Last Name], qryTotalTimeUsedPerEmployee.Employee AS Employee, qryTotalTimeUsedPerEmployee.SumOfSumOfTime_used, 480-[SumOfSumOfTime_used] AS Remaining
FROM tblRequests, tblEmployees, tblRequestTimeLog, qryTotalTimeUsedPerEmployee
WHERE (((tblEmployees.[Employee ])=[qryTotalTimeUsedPerEmployee].[Employee]));

