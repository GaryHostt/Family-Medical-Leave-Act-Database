SELECT tblEmployees.[Employee ], tblEmployees.[First Name], tblEmployees.[Last Name], 365-DateDiff("d",[Date_hired],Date()) AS [Days until eligible]
FROM tblEmployees
WHERE (((DateDiff("d",[Date_hired],Date()))<364));

