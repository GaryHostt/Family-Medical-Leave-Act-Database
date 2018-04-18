SELECT DISTINCTROW tblRequests.Employee, tblEmployees.[First Name], tblEmployees.[Last Name], (DateDiff("d",[EndDate],Date()))*(-1) AS Days_until_over, tblRequests.RequestNumber
FROM tblRequests, tblEmployees, tblRequestTimeLog
WHERE (((tblRequests.Employee)=tblEmployees.Employee) And ((tblRequests.[Approved?])=Yes) And ((DateDiff("d",[EndDate],Date()))<0));

