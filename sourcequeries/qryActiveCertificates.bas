SELECT tblRequests.Employee, tblRequests.Date_given_application, tblRequests.Date_HR_received_application, tblRequests.Self, tblRequests.Child, tblRequests.Spouse, tblRequests.Parent, tblRequests.[Approved?], tblRequests.Reason, tblRequests.Frequency, tblRequests.Frequency_2, tblRequests.Frequency_3, tblRequests.Frequency_4, tblRequests.Type, tblRequests.[Condition (illness)], tblRequests.Start, tblRequests.EndDate, tblRequests.RequestNumber, tblRequests.Total_Allocated_Time, DateDiff("d",[Start],Date()) AS DaySince
FROM tblRequests
WHERE (((tblRequests.[Approved?])=Yes) AND ((DateDiff("d",[Start],Date()))<91));

