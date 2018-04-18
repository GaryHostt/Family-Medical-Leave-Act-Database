SELECT tblEmployees.[Employee ], tblEmployees.[First Name], tblEmployees.[Last Name]
FROM tblEmployees
WHERE (((DateDiff("d",[Date_hired],Date()))>364));

