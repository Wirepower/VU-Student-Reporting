-- Run once on ElectrotechnologyReports (adjust schema if needed).
-- Reminder-only table: does not change AgreementsDetails or email behaviour.

IF OBJECT_ID(N'dbo.StudentEmploymentStatus', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.StudentEmploymentStatus (
        StudentID            NVARCHAR(50)  NOT NULL,
        EmploymentStatus     NVARCHAR(50)  NULL,
        DateofUnemployment   DATE          NULL,
        CONSTRAINT PK_StudentEmploymentStatus PRIMARY KEY CLUSTERED (StudentID)
    );
END
GO
