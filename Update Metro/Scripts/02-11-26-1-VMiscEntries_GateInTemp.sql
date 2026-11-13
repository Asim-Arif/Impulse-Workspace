USE [Impulse_AWM]
GO

/****** Object:  View [dbo].[VMiscEntries_GateInTemp]    Script Date: 11/2/2026 10:51:27 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO


CREATE   VIEW [dbo].[VMiscEntries_GateInTemp]
AS
SELECT
    m.EntryID,
    m.GatePassNo,
    m.DT,
    m.PersonName,
    m.VehicleNo,
    m.Description,
    m.UserName,
    m.MachineName,
    m.DTEntry,
    CAST(CONVERT(date, m.DT) AS datetime) AS OnlyDT,
    m.EmpID,
    m.GatePassType,
    m.DeptID,
    m.VendID,
    ISNULL(e.Name, m.EmpNameManual) AS Name,
    m.EmpNameManual,
    m.VendNameManual
FROM dbo.MiscEntries_GateInTemp m
LEFT JOIN dbo.Employees e ON e.empid COLLATE DATABASE_DEFAULT = m.EmpID COLLATE DATABASE_DEFAULT;
GO


