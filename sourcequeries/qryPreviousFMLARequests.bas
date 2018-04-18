SELECT DISTINCT tblRequests.Employee, tblRequests.RequestNumber, tblEmployees.[First Name], tblEmployees.[Last Name], (DateDiff("d",[EndDate],Date())) AS Days_since_expired, tblRequests.[Approved?]
FROM tblRequests, tblEmployees, tblRequestTimeLog
WHERE (((tblRequests.Employee)=tblEmployees.Employee) And (((DateDiff("d",[EndDate],Date())))>0));

