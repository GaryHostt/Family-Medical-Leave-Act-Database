SELECT tblRequestTimeLog.Employee, tblRequestTimeLog.RequestNumber, Sum(tblRequestTimeLog.Time_used) AS SumOfTime_used, tblRequests.Start, [Time_used]/((DateDiff("d",[Start],Date()))+0.00000001) AS HoursPerDay, 365-DateDiff("d",[Start],Date()) AS DaysUntilExpiration
FROM tblRequestTimeLog, tblRequests
GROUP BY tblRequestTimeLog.Employee, tblRequestTimeLog.RequestNumber, tblRequests.Start, [Time_used]/((DateDiff("d",[Start],Date()))+0.00000001), tblRequestTimeLog.Date_Time_Used, DateDiff("d",[Date_Time_Used],Date()), tblRequests.RequestNumber, 365-DateDiff("d",[Start],Date())
HAVING (((DateDiff("d",[Date_Time_Used],Date()))<365) And ((tblRequests.RequestNumber)=tblRequestTimeLog.RequestNumber));

