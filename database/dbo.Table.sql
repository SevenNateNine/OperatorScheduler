CREATE TABLE [dbo].[Table]
(
	[Id] INT NOT NULL PRIMARY KEY, 
    [EmployeeID] NUMERIC(10) NULL, 
    [FirstName] NCHAR(50) NOT NULL, 
    [LastName] NCHAR(50) NOT NULL, 
    [Seniority] NCHAR(10) NOT NULL DEFAULT 0
)
